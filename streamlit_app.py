"""Moo Housing Investor Reporting: AppFolio + Google Sheets + reserves + reports."""
import io,re,sqlite3,json
from datetime import datetime,timezone
from pathlib import Path
import pandas as pd
import streamlit as st
from engine import number,pid,period,read_appfolio,parse_income,parse_rent_roll,calculate
from reporting import make_docx,make_pdf

st.set_page_config(page_title='Moo Housing | Investor Report Generator',layout='wide')
st.title('MOO HOUSING · Investor Report Generator')
st.caption('AppFolio Excel → NOI / Occupancy → Cash & Reserve → Save → Word / PDF')
HERE=Path(__file__).parent
SHEETS=['Properties','Monthly Statements','Investors','Distribution History']
DB=HERE/'moo_local_records.sqlite3'

def local_db():
    con=sqlite3.connect(DB)
    con.execute('CREATE TABLE IF NOT EXISTS statements (property_id TEXT, month TEXT, payload TEXT, saved_at TEXT, PRIMARY KEY(property_id,month))')
    return con

def excel_load(blob):
    x=pd.ExcelFile(io.BytesIO(blob));return {name:pd.read_excel(x,sheet_name=name,dtype=object).fillna('') for name in x.sheet_names if name in SHEETS}

GS_SPREADSHEET_NAME = "Reports"

def get_gc_client():
    """沿用现有支票系统的 Streamlit Secrets / Service Account 授权。"""
    import gspread
    from google.oauth2.service_account import Credentials
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_dict = dict(st.secrets["GOOGLE_APPLICATION_CREDENTIALS"])
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    credentials = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(credentials)

def book_open():
    return get_gc_client().open(GS_SPREADSHEET_NAME)

def google_load():
    book = book_open()
    data = {}
    for title in SHEETS:
        try:
            ws = book.worksheet(title)
        except Exception as exc:
            if title in ('Properties', 'Monthly Statements'):
                raise ValueError(f"Reports 缺少必要工作表 {title}: {exc}") from exc
            continue
        data[title] = pd.DataFrame(ws.get_all_records(default_blank=''))
    return data

def records(df):return df.fillna('').to_dict('records') if df is not None and not df.empty else []

def store_google(rec):
    book=book_open();ws=book.worksheet('Monthly Statements')
    heads=ws.row_values(1)
    if not heads:raise ValueError('Monthly Statements 缺少表头')
    missing=[k for k in rec if k not in heads]
    if missing:
        heads+=missing
        ws.update(range_name='A1',values=[heads])
    key=pid(rec['Property ID']);month=rec['Month (YYYY-MM)']
    vals=ws.get_all_values();existing=None
    for i,r in enumerate(vals[1:],start=2):
        if len(r)>=2 and pid(r[0])==key and period(r[1])==month:existing=i;break
    new=[rec.get(h,'') for h in heads]
    if existing:
        old=vals[existing-1];oldstatus=old[heads.index('Status')] if 'Status' in heads and len(old)>heads.index('Status') else ''
        if str(oldstatus).lower()=='confirmed':raise ValueError('该月份已经 Confirmed，不能直接覆盖。请先在 Google Sheets 中审核并改为 Draft。')
        ws.update(range_name=f'A{existing}',values=[new])
    else:ws.append_row(new,value_input_option='USER_ENTERED')

def store_local(rec):
    with local_db() as c:
        old=c.execute('SELECT payload FROM statements WHERE property_id=? AND month=?',(pid(rec['Property ID']),rec['Month (YYYY-MM)'])).fetchone()
        if old and json.loads(old[0]).get('Status')=='Confirmed':raise ValueError('本地记录已 Confirmed；不能直接覆盖。')
        c.execute('INSERT OR REPLACE INTO statements VALUES (?,?,?,?)',(pid(rec['Property ID']),rec['Month (YYYY-MM)'],json.dumps(rec,ensure_ascii=False),datetime.now(timezone.utc).isoformat()))

def local_records():
    with local_db() as c:return [json.loads(r[0]) for r in c.execute('SELECT payload FROM statements')]

with st.sidebar:
    st.header('数据来源')
    source=st.radio('Master 数据', ['Google Sheets（Reports）','上传 Excel（本机测试）'])
    data=None
    if source.startswith('上传'):
        file=st.file_uploader('上传 Moo_Housing_Master.xlsx',type=['xlsx','xlsm'],key='master')
        if file:
            try:data=excel_load(file.getvalue())
            except Exception as e:st.error(str(e))
    else:
        st.caption('自动连接 Google Sheets：Reports')
        if st.button('读取 / 刷新 Reports') or 'google_data' not in st.session_state:
            try:st.session_state['google_data']=google_load();st.success('已连接 Reports')
            except Exception as e:st.error(f'读取失败：{e}')
        data=st.session_state.get('google_data')
    st.caption('固定读取名为 Reports 的 Google Sheets；请将其共享给 Service Account（Editor）。')
if not data or 'Properties' not in data or 'Monthly Statements' not in data:
    st.info('请先上传压缩包中的 Master Excel，或连接 Google Sheets。');st.stop()
props=records(data['Properties']);sheet_monthly=records(data['Monthly Statements'])
if not props:st.error('Properties 没有物业信息');st.stop()
pmap={pid(p.get('Property ID')):p for p in props if pid(p.get('Property ID'))}
if not pmap:st.stop()
selected=st.selectbox('Property',list(pmap),format_func=lambda k:f"{pmap[k].get('Property Name / LLC',k)} ({k})")
p=pmap[selected]
st.caption(f"{p.get('Property Address','')}  |  Report Language: {p.get('Report Language','English')}")
month=st.text_input('Reporting Month (YYYY-MM)',value=datetime.now().strftime('%Y-%m'))
if not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])',month):st.error('月份格式必须是 YYYY-MM');st.stop()

st.subheader('1 · 上传 AppFolio 报表')
a,b=st.columns(2)
with a:income_file=st.file_uploader('AppFolio Income Statement Excel',type=['xlsx','xls'],key='inc')
with b:rent_file=st.file_uploader('AppFolio Rent Roll Excel',type=['xlsx','xls'],key='roll')
financial=None;occupancy=None
if income_file:
    try:financial=parse_income(read_appfolio(income_file))
    except Exception as e:st.error(f'Income Statement: {e}')
if rent_file:
    try:occupancy=parse_rent_roll(read_appfolio(rent_file))
    except Exception as e:st.error(f'Rent Roll: {e}')
if financial:
    st.write('AppFolio 识别结果',{'Income':financial.get('income'),'Expenses':financial.get('expenses'),'NOI':financial.get('noi')})
    if all(financial.get(x) is not None for x in ['income','expenses','noi']) and abs(financial['income']-financial['expenses']-financial['noi'])>.05:
        st.warning('Income - Expenses 与 NOI 不一致。请确认报表符号和费用分类。')
if occupancy:
    st.metric('Occupancy',f"{occupancy['occupancy']:.1%}",f"{occupancy['occupied']} / {occupancy['rooms']} rooms")
    for w in occupancy['warnings']:st.warning(w)

previous=next((r for r in sheet_monthly if pid(r.get('Property ID'))==selected and period(r.get('Month (YYYY-MM)'))==month),{})
if source.startswith('上传'):
    previous=next((r for r in local_records() if pid(r.get('Property ID'))==selected and r.get('Month (YYYY-MM)')==month),previous)

st.subheader('2 · 月度现金与储备')
st.caption('上传 AppFolio 后自动带入财务数据；如需要，可以人工修正。')
with st.form('monthly'):
    c1,c2,c3=st.columns(3)
    with c1:
        inc=st.number_input('Total Income',value=float(number(financial.get('income') if financial else previous.get('Income'))),step=100.,format='%.2f')
        exp=st.number_input('Operating Expenses',value=float(number(financial.get('expenses') if financial else previous.get('Expenses'))),step=100.,format='%.2f')
        noi=st.number_input('NOI',value=float(number(financial.get('noi') if financial else previous.get('NOI'))),step=100.,format='%.2f')
    with c2:
        tax=st.number_input('Property Tax Paid from Reserve',value=float(number(previous.get('Property Tax Paid'))),min_value=0.,format='%.2f')
        ins=st.number_input('Insurance Paid from Reserve',value=float(number(previous.get('Insurance Paid'))),min_value=0.,format='%.2f')
        tax_in_noi=st.checkbox('Tax / Insurance 已经包含在 NOI 费用里',value=str(previous.get('Tax/Insurance Included in NOI?','Yes')).lower()!='no')
    with c3:
        other=st.number_input('Other Cash Expenses NOT in NOI',value=float(number(previous.get('Other Cash Expense'))),min_value=0.,format='%.2f')
        adjustments=st.number_input('Other Cash Adjustments (+/-)',value=float(number(previous.get('Other Cash Adjustments'))),format='%.2f')
        dist=st.number_input('Distribution Paid (本月已支付)',value=float(number(previous.get('Distribution Paid'))),min_value=0.,format='%.2f')
    st.caption('银行余额请从银行核对；不填写则仅提供根据 NOI 推算的现金余额，不能用于正式分红。')
    bank_known=st.checkbox('我已核对本月期末银行余额',value=str(previous.get('Verified Ending Bank Balance','')).strip()!='')
    bank=st.number_input('Verified Ending Bank Balance',value=float(number(previous.get('Verified Ending Bank Balance'))),format='%.2f',disabled=not bank_known)
    notes=st.text_area('Notes / Operating Commentary',value=str(previous.get('Notes','')))
    status=st.selectbox('保存状态',['Draft','Confirmed'],index=0)
    submit=st.form_submit_button('计算 / Preview')

rec={'Property ID':p.get('Property ID'),'Month (YYYY-MM)':month,'Income':inc,'Expenses':exp,'NOI':noi,
     'Property Tax Paid':tax,'Insurance Paid':ins,'Tax/Insurance Included in NOI?':'Yes' if tax_in_noi else 'No',
     'Other Cash Expense':other,'Other Cash Adjustments':adjustments,'Distribution Paid':dist,
     'Verified Ending Bank Balance':bank if bank_known else '', 'Notes':notes,'Status':status,
     'Occupancy':occupancy['occupancy'] if occupancy else previous.get('Occupancy',''),
     'Occupied Rooms':occupancy['occupied'] if occupancy else previous.get('Occupied Rooms',''),
     'Total Rooms':occupancy['rooms'] if occupancy else previous.get('Total Rooms','')}
combined=[r for r in sheet_monthly if not(pid(r.get('Property ID'))==selected and period(r.get('Month (YYYY-MM)'))==month)]
if source.startswith('上传'):
    local=local_records();combined=[r for r in combined if not any(pid(r.get('Property ID'))==pid(x.get('Property ID')) and period(r.get('Month (YYYY-MM)'))==period(x.get('Month (YYYY-MM)')) for x in local)]+[x for x in local if not(pid(x.get('Property ID'))==selected and x.get('Month (YYYY-MM)')==month)]
combined.append(rec)
rows,issues=calculate(props,combined)
for issue in issues:st.warning(issue)
match=[r for r in rows if r['Property ID']==selected and r['Month (YYYY-MM)']==month]
if not match:st.stop()
r=match[0]
st.subheader('3 · 自动计算结果')
x,y,z,w=st.columns(4)
x.metric('NOI',f"${r['NOI']:,.2f}");y.metric('Ending Cash',f"${r['Ending Cash']:,.2f}");z.metric('Required Reserve',f"${r['Required Reserve']:,.2f}");w.metric('Shortfall' if r['Reserve Shortfall']>0 else 'Available Distribution',f"${max(r['Reserve Shortfall'],r['Available to Distribute']):,.2f}")
st.dataframe(pd.DataFrame([r]),hide_index=True,use_container_width=True)
if not bank_known:st.warning('银行余额尚未核对。Ending Cash 是根据 NOI 推算的金额，不能视为已确认可分红现金。')
if r['Reserve Shortfall']>0:st.error('存在 Reserve Shortfall，建议暂停分红。')
if status=='Confirmed' and (not bank_known or not income_file or not rent_file):st.warning('确认状态要求上传两份 AppFolio 报表并核对银行余额。')
if st.button('💾 Save to Google Sheets' if source.startswith('Google') else '💾 Save locally (Excel 模式)'):
    if status=='Confirmed' and (not bank_known or not income_file or not rent_file):st.error('Confirmed 必须提供两份 AppFolio 报表和核对银行余额。')
    else:
        try:
            if source.startswith('Google'):store_google(rec);st.success('已保存到 Google Sheets。请点击侧栏刷新读取。')
            else:store_local(rec);st.success('已保存到本机 SQLite 数据库。Excel 文件本身未被修改；切换电脑不会自动同步。')
        except Exception as e:st.error(f'保存失败：{e}')

st.subheader('4 · Investor Report 下载')
if not income_file or not rent_file:st.info('未同时上传两份 AppFolio 报表，以下报告为待核对草稿。')
lang=str(p.get('Report Language','English')).strip().lower()
languages=['中文','English'] if lang in ('both','bilingual','中英文') else (['中文'] if lang in ('中文','chinese','cn') else ['English'])
for language in languages:
    cn=language=='中文';stem=f"Moo_{selected}_{month}_{'CN' if cn else 'EN'}"
    try:
        d=make_docx(r,p,cn);pdf=make_pdf(r,p,cn)
        left,right=st.columns(2)
        left.download_button(f'{language} · Download Word',d,stem+'.docx',mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',key='doc'+language)
        right.download_button(f'{language} · Download PDF',pdf,stem+'.pdf',mime='application/pdf',key='pdf'+language)
    except Exception as e:st.error(f'{language} 报告生成失败：{e}')
st.subheader('5 · 历史月度记录')
if rows:st.dataframe(pd.DataFrame([x for x in rows if x['Property ID']==selected]).sort_values('Month (YYYY-MM)'),hide_index=True,use_container_width=True)
st.caption('内部管理工具。系统不自动付款、不自动发送邮件；确认分红前请核对银行余额、应付款项和实际分红流水。')
