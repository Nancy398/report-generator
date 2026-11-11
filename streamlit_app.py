"""Moo Housing Investor Reporting: AppFolio + Google Sheets + reserves + reports."""
import io,re,sqlite3,json
from datetime import datetime,timezone
from pathlib import Path
import pandas as pd
import streamlit as st
from engine import number,pid,period,read_appfolio,parse_income,parse_rent_roll,calculate
from reporting import make_docx,make_pdf,append_pdf_documents
from chinese_appendix import translate_income, translate_rent_roll
from pdf_appfolio import parse_income_pdf,parse_rent_roll_pdf
from investor_communication import show_communications

st.set_page_config(page_title='Moo Housing | Investor Report Generator',layout='wide')
st.title('MOO HOUSING · Investor Report Generator')
st.caption('AppFolio PDF → NOI / Occupancy → Cash & Reserve → Word / PDF with original PDF appendices')
HERE=Path(__file__).parent
SHEETS=['Properties','Monthly Statements','Investors','Distribution History','Summaries']
DB=Path('/tmp')/'moo_local_records.sqlite3'

def local_db():
    con=sqlite3.connect(DB)
    con.execute('CREATE TABLE IF NOT EXISTS statements (property_id TEXT, month TEXT, payload TEXT, saved_at TEXT, PRIMARY KEY(property_id,month))')
    return con

def excel_load(blob):
    x=pd.ExcelFile(io.BytesIO(blob));return {name:pd.read_excel(x,sheet_name=name,dtype=object).fillna('') for name in x.sheet_names if name in SHEETS}

GS_SPREADSHEET_NAME = "Reports"
GS_SPREADSHEET_ID = "1vESBpLsjZEO67UCDnTuZgU0GlOes4jSZLzi9QGIhNLY"

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
    return get_gc_client().open_by_key(GS_SPREADSHEET_ID)

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
        if st.button('读取 / 刷新 Reports'):
            try:st.session_state['google_data']=google_load();st.success('已连接 Reports')
            except Exception as e:st.error(f'读取失败：{e}')
        data=st.session_state.get('google_data')
        if data is None:
            st.info('点击「读取 / 刷新 Reports」后再连接 Google Sheets。')
    st.caption('固定读取名为 Reports 的 Google Sheets；请将其共享给 Service Account（Editor）。')
if not data or 'Properties' not in data or 'Monthly Statements' not in data:
    st.info('请先上传压缩包中的 Master Excel，或连接 Google Sheets。');st.stop()
props=records(data['Properties']);sheet_monthly=records(data['Monthly Statements']);summaries=records(data.get('Summaries'))
if not props:st.error('Properties 没有物业信息');st.stop()
pmap={pid(p.get('Property ID')):p for p in props if pid(p.get('Property ID'))}
if not pmap:st.stop()
report_type=st.radio('Report Type',['Single Property Report','Combined LLC Report'],horizontal=True)
if report_type=='Combined LLC Report':
    from combined_llc import build_model,make_combined_docx,make_combined_pdf,generated_summary
    from reporting import append_pdf_documents
    groups={}
    for prop in props:
        group=str(prop.get('Property Name / LLC','') or '').strip()
        if group and pid(prop.get('Property ID')):
            key=re.sub(r'[^a-z0-9]','',group.lower())
            if key not in groups:groups[key]={'name':group,'members':[]}
            groups[key]['members'].append(prop)
    options=[key for key,entry in groups.items() if len({pid(x.get('Property ID')) for x in entry['members']})>=2]
    if not options:
        st.error('Properties 中没有发现两个或以上 Property ID 对应相同 Property Name / LLC 的公司。')
        st.stop()
    group_key=st.selectbox('LLC',options,format_func=lambda key:groups[key]['name'])
    group=groups[group_key]['name']
    month_combined=st.text_input('Reporting Month (YYYY-MM)',value=datetime.now().strftime('%Y-%m'),key='combined_month')
    st.caption('Included properties: '+', '.join(str(p.get('Property Address') or p.get('Property ID')) for p in groups[group_key]['members']))
    st.info('先在 Single Property Report 分别保存各物业当月数据。合并报告会相加各物业 Ending Cash、Required Reserve、Distribution Paid；请确保各物业现金记录不重复包含同一银行账户余额。')
    try:
        model=build_model(group,month_combined,props,sheet_monthly,summaries)
    except Exception as e:
        st.error(f'无法生成合并报告：{e}')
        st.stop()
    st.subheader('Combined Executive Summary')
    preview_lang=st.radio('Summary Preview Language',['中文','English'],horizontal=True,key='llc_summary_preview_lang')
    preview_cn=preview_lang=='中文'
    operating_text,cash_text=generated_summary(model,preview_cn)
    st.markdown('**经营表现**' if preview_cn else '**OPERATING PERFORMANCE**')
    st.text(operating_text)
    st.markdown('**LLC 现金及分红**' if preview_cn else '**LLC CASH & DISTRIBUTIONS**')
    st.write(cash_text)
    manual_note=model['summary_cn' if preview_cn else 'summary_en'].strip()
    if manual_note:
        st.markdown('**其他经营事项**' if preview_cn else '**OTHER OPERATING NOTES**')
        st.write(manual_note)

    st.dataframe(pd.DataFrame(model['properties'])[['name','occupancy','prior_occupancy','noi','rent','income','expenses']],hide_index=True,use_container_width=True)
    c1,c2,c3=st.columns(3)
    c1.metric('LLC Available Cash',f"${model['available']:,.2f}")
    c2.metric('LLC Distribution Paid',f"${model['distribution']:,.2f}")
    c3.metric('LLC Reserve Shortfall',f"${model['shortfall']:,.2f}")
    st.caption('可分配余额 = max(0, LLC Ending Cash − LLC Required Reserve)；季度余额为期末累计余额，不叠加各月余额。')
    if model['shortfall']>0:st.warning('LLC 存在 Reserve Shortfall，请先审核分红。')
    st.subheader('Optional · AppFolio Original PDF Appendices')
    uploads=st.file_uploader('上传各物业原始 Income Statement / Rent Roll PDF（可多选）',type=['pdf'],accept_multiple_files=True,key='llc_appendices')
    for language,cn in [('中文',True),('English',False)]:
        try:
            doc=make_combined_docx(model,cn)
            stem=f"Moo_{re.sub(r'[^A-Za-z0-9]+','_',group)}_{month_combined}_{'CN' if cn else 'EN'}_Combined"
            col1,col2=st.columns(2)
            col1.download_button(f'{language} · Download Combined Word',doc,stem+'.docx',mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',key='llc_doc'+language)
            try:
                pdf=make_combined_pdf(doc)
                if uploads:pdf=append_pdf_documents(pdf,[(f.name,f.getvalue()) for f in uploads])
                col2.download_button(f'{language} · Download Combined PDF',pdf,stem+'.pdf',mime='application/pdf',key='llc_pdf'+language)
            except Exception as e:st.info(f'{language} PDF 暂不可用：{e}')
        except Exception as e:st.error(f'{language} Word 生成失败：{e}')
    op_cn, cash_cn = generated_summary(model, True)
    op_en, cash_en = generated_summary(model, False)
    show_communications(group, month_combined, op_cn, cash_cn, model.get('summary_cn',''),
                        op_en, cash_en, model.get('summary_en',''), key='llc_'+group_key)
    st.stop()
selected=st.selectbox('Property',list(pmap),format_func=lambda k:f"{pmap[k].get('Property Name / LLC',k)} ({k})")
p=pmap[selected]
st.caption(f"{p.get('Property Address','')}  |  Report Language: {p.get('Report Language','English')}")
month=st.text_input('Reporting Month (YYYY-MM)',value=datetime.now().strftime('%Y-%m'))
if not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])',month):st.error('月份格式必须是 YYYY-MM');st.stop()

st.subheader('1 · 上传 AppFolio PDF 报表')
a,b=st.columns(2)
with a:income_file=st.file_uploader('AppFolio Income Statement PDF',type=['pdf'],key='inc')
with b:rent_file=st.file_uploader('AppFolio Rent Roll PDF',type=['pdf'],key='roll')
st.caption('两份 PDF 同时用于识别数据，并自动按 Income Statement → Rent Roll 顺序附在最终 PDF 后面。无需单独上传 Appendix。')
appendix_files=[f for f in (income_file,rent_file) if f is not None]
financial=None;occupancy=None
if income_file:
    try:financial=parse_income_pdf(income_file.getvalue())
    except Exception as e:st.error(f'Income Statement PDF 识别失败：{e}')
if rent_file:
    try:occupancy=parse_rent_roll_pdf(rent_file.getvalue())
    except Exception as e:st.error(f'Rent Roll PDF 识别失败：{e}')
for label,record in [('Income Statement',financial),('Rent Roll',occupancy)]:
    if record and record.get('period') and record['period']!=month:
        st.error(f'{label} 的报表月份为 {record["period"]}，当前选择 {month}。请修改 Reporting Month，否则不能保存或生成正式报告。')
period_mismatch=any(x and x.get('period') and x['period']!=month for x in (financial,occupancy))
if financial:
    for field, title in [('rent_income','Rent Income'),('utility_income','Utility Income')]:
        if field not in financial:
            st.warning(f'{title} 未从 PDF 识别到；请核对并在下方手动填写。')
    st.write('AppFolio 识别结果',{'Rent Income':financial.get('rent_income'),'Utility Income (Total FEES)':financial.get('utility_income'),'Other Income':financial.get('other_income'),'Total Income':financial.get('income'),'Expenses':financial.get('expenses'),'NOI':financial.get('noi')})
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
        rent_income=st.number_input('Rent Income',value=float(number(financial.get('rent_income') if financial else previous.get('Rent Income'))),format='%.2f')
        utility_income=st.number_input('Utility Income (AppFolio Total FEES)',value=float(number(financial.get('utility_income') if financial else previous.get('Utility Income'))),format='%.2f')
        other_income=st.number_input('Other Income',value=float(number(financial.get('other_income') if financial else previous.get('Other Income'))),format='%.2f')
    with c2:
        tax=st.number_input('Property Tax Paid from Reserve',value=float(number(previous.get('Property Tax Paid'))),min_value=0.,format='%.2f')
        ins=st.number_input('Insurance Paid from Reserve',value=float(number(previous.get('Insurance Paid'))),min_value=0.,format='%.2f')
        st.caption('Tax / Insurance Paid 按独立现金支付扣除；请确保 NOI 未重复扣除相同支出。')
    with c3:
        other=st.number_input('Other Cash Expenses NOT in NOI',value=float(number(previous.get('Other Cash Expense'))),min_value=0.,format='%.2f')
        adjustments=st.number_input('Other Cash Adjustments (+/-)',value=float(number(previous.get('Other Cash Adjustments'))),format='%.2f')
        dist=st.number_input('Distribution Paid (本月已支付)',value=float(number(previous.get('Distribution Paid'))),min_value=0.,format='%.2f')
    st.caption('Ending Cash 自动按上月余额 + NOI − Tax − Insurance − 其他支出 − 分红 + 调整计算，不需要核对银行余额。')
    notes=st.text_area('Executive Summary (manual)',value=str(previous.get('Notes','')))
    status=st.selectbox('保存状态',['Draft','Confirmed'],index=0)
    submit=st.form_submit_button('计算 / Preview')

rec={'Property ID':p.get('Property ID'),'Month (YYYY-MM)':month,'Rent Income':rent_income,'Utility Income':utility_income,'Other Income':other_income,'Income':inc,'Expenses':exp,'NOI':noi,
     'Property Tax Paid':tax,'Insurance Paid':ins,'Calculation Version':'modeled-v2',
     'Other Cash Expense':other,'Other Cash Adjustments':adjustments,'Distribution Paid':dist,
     'Notes':notes,'Status':status,
     'Occupancy':occupancy['occupancy'] if occupancy else previous.get('Occupancy',''),
     'Occupied Rooms':occupancy['occupied'] if occupancy else previous.get('Occupied Rooms',''),
     'Total Rooms':occupancy['rooms'] if occupancy else previous.get('Total Rooms','')}
if abs((rent_income+utility_income+other_income)-inc)>.05:
    st.warning('Income Breakdown 与 Total Income 不一致，请核对 Rent / FEES / Other Income；Total Income 和 NOI 不会自动改变。')
combined=[r for r in sheet_monthly if not(pid(r.get('Property ID'))==selected and period(r.get('Month (YYYY-MM)'))==month)]
if source.startswith('上传'):
    local=local_records();combined=[r for r in combined if not any(pid(r.get('Property ID'))==pid(x.get('Property ID')) and period(r.get('Month (YYYY-MM)'))==period(x.get('Month (YYYY-MM)')) for x in local)]+[x for x in local if not(pid(x.get('Property ID'))==selected and x.get('Month (YYYY-MM)')==month)]
combined.append(rec)
rows,issues=calculate(props,combined)
for issue in issues:st.warning(issue)
match=[r for r in rows if r['Property ID']==selected and r['Month (YYYY-MM)']==month]
if not match:st.stop()
r=match[0]
# Preserve the current unsaved manual summary in both Word and PDF output.
r = dict(r)
r['Notes'] = notes
r['Summary'] = notes
# Match the separate Summaries tab by Property ID, property name, or address + month.
def _match_property(value):
    def norm(x):return re.sub(r'[^a-z0-9]','',str(x or '').lower())
    aliases=[selected,p.get('Property ID'),p.get('Property Name / LLC'),p.get('Property Address'),p.get('Property Name')]
    return bool(norm(value)) and any(norm(value)==norm(x) for x in aliases)
summary_record=next((s for s in summaries if _match_property(s.get('Property')) and period(s.get('Month'))==month),{})
r['Summary CN']=str(summary_record.get('Summary CN','') or '').strip()
r['Summary EN']=str(summary_record.get('Summary EN','') or '').strip()
r['Distribution Frequency']=str(p.get('Distribution Frequency','Monthly') or 'Monthly').strip()

# Use the current form's income breakdown for report generation, even when
# a historical calculated row or cached Google Sheets row has missing fields.
for _field in ('Rent Income', 'Utility Income', 'Other Income'):
    r[_field] = rec[_field]
from datetime import date
report_year,report_month=map(int,month.split('-'))
prior_month=f'{report_year-1:04d}-12' if report_month==1 else f'{report_year:04d}-{report_month-1:02d}'
prior_row=next((x for x in rows if pid(x.get('Property ID'))==selected and period(x.get('Month (YYYY-MM)'))==prior_month),None)
# If no calculated historical row exists, use saved source data (never fabricate).
if prior_row is None:
    prior_row=next((x for x in combined if pid(x.get('Property ID'))==selected and period(x.get('Month (YYYY-MM)'))==prior_month),None)
# Single Property Executive Summary preview uses the same financial inputs as the Word report.
def _summary_number(value):
    try:
        if value is None or str(value).strip() == '':
            return None
        return float(str(value).replace(',', '').replace('$', '').replace('%', ''))
    except (TypeError, ValueError):
        return None

def _single_summary_preview(row, previous_row, chinese):
    fmt=lambda n: f'${n:,.2f}'
    current=_summary_number(row.get('Occupancy'))
    previous_occ=_summary_number((previous_row or {}).get('Occupancy'))
    pct=lambda v: v*100 if v is not None and v<=1 else v
    if current is None:
        occ='本月入住率数据尚未提供。' if chinese else 'Current-month occupancy data is unavailable.'
    elif previous_occ is None:
        occ=(f'本月入住率为 {pct(current):.1f}%；暂无上月可比数据。' if chinese else f'Occupancy was {pct(current):.1f}% this month; prior-month data is unavailable.')
    else:
        delta=pct(current)-pct(previous_occ)
        if chinese:
            occ=f'本月入住率为 {pct(current):.1f}%，上月为 {pct(previous_occ):.1f}%，' + (f'环比上升 {delta:.1f} 个百分点。' if delta>0.005 else f'环比下降 {abs(delta):.1f} 个百分点。' if delta< -0.005 else '环比持平。')
        else:
            occ=f'Occupancy was {pct(current):.1f}% versus {pct(previous_occ):.1f}% last month; ' + (f'it increased by {delta:.1f} percentage points.' if delta>0.005 else f'it decreased by {abs(delta):.1f} percentage points.' if delta< -0.005 else 'it remained unchanged.')
    noi_value=_summary_number(row.get('NOI'))
    noi_text=(f'本月净营业收入（NOI）为 {fmt(noi_value)}。' if chinese else f'Net Operating Income (NOI) for the month was {fmt(noi_value)}.') if noi_value is not None else ('本月 NOI 数据尚未提供。' if chinese else 'NOI data for this month is unavailable.')
    available=_summary_number(row.get('Available to Distribute'))
    paid=_summary_number(row.get('Distribution Paid'))
    quarterly=str(row.get('Distribution Frequency','Monthly')).strip().lower() in ('quarterly','quarter','季度','季度分红')
    if available is None:
        cash='可分配现金余额尚未提供。' if chinese else 'Available cash for distribution is unavailable.'
    elif quarterly:
        cash=(f'截至本月末，累计可分配现金为 {fmt(available)}。' if chinese else f'Cumulative available cash for distribution at month-end was {fmt(available)}.')
        cash+=('该余额将结转至下一季度分红评估。' if chinese else ' This balance will carry forward to the next quarterly distribution review.')
    else:
        cash=(f'本月期末可分配现金余额为 {fmt(available)}。' if chinese else f'Available cash for distribution at month-end was {fmt(available)}.')
    note=str(row.get('Summary CN' if chinese else 'Summary EN','') or '').strip() or str(row.get('Summary','') or '').strip()
    return occ+' '+noi_text,cash,note

st.subheader('Executive Summary · 自动经营摘要')
preview_language=st.radio('Summary Preview Language',['中文','English'],horizontal=True,key='single_summary_preview_language')
preview_chinese=preview_language=='中文'
operation_summary,distribution_summary,manual_summary=_single_summary_preview(r,prior_row,preview_chinese)
st.markdown('**经营表现**' if preview_chinese else '**OPERATING PERFORMANCE**')
st.write(operation_summary)
st.markdown('**现金及分红情况**' if preview_chinese else '**CASH & DISTRIBUTIONS**')
st.write(distribution_summary)
if manual_summary:
    st.markdown('**其他经营事项**' if preview_chinese else '**OTHER OPERATING NOTES**')
    st.write(manual_summary)

st.subheader('3 · 自动计算结果')
x,y,z,w=st.columns(4)
x.metric('NOI',f"${r['NOI']:,.2f}");y.metric('Modeled Ending Cash',f"${r['Ending Cash']:,.2f}");z.metric('Cumulative Required Reserve',f"${r['Required Reserve']:,.2f}");w.metric('Shortfall' if r['Reserve Shortfall']>0 else 'Available Distribution',f"${max(r['Reserve Shortfall'],r['Available to Distribute']):,.2f}")
st.dataframe(pd.DataFrame([r]),hide_index=True,use_container_width=True)
st.caption(f"Beginning Cash ${r['Beginning Cash']:,.2f} + NOI ${r['NOI']:,.2f} − Tax ${tax:,.2f} − Insurance ${ins:,.2f} − Other ${other:,.2f} − Distribution ${dist:,.2f} + Adjustments ${adjustments:,.2f} = Modeled Ending Cash ${r['Modeled Ending Cash']:,.2f}")
st.caption('Ending Cash 为系统推算余额；Required Reserve 为上月累计预留 + 本月固定预留 − 本月实际支付的 Property Tax 和 Insurance。')
if r['Reserve Shortfall']>0:st.error('存在 Reserve Shortfall，建议暂停分红。')
if status=='Confirmed' and (not income_file or not rent_file):st.warning('Confirmed 状态要求上传两份 AppFolio 报表。')
# Save both source inputs and calculated columns, using the existing Reports sheet headings.
save_rec = dict(rec)
save_rec.update({
    'Beginning Cash': r['Beginning Cash'],
    'Modeled Ending Cash': r['Modeled Ending Cash'],
    'Ending Cash': r['Ending Cash'],
    'Tax Reserve Required': r['Tax Reserve Required'],
    'Insurance Reserve Required': r['Insurance Reserve Required'],
    'Total Required Reserve': r['Required Reserve'],
    'Required Reserve': r['Required Reserve'],
    'Reserve Shortfall': r['Reserve Shortfall'],
    'Available to Distribute': r['Available to Distribute'],
    'Monthly Reserve Addition': r['Monthly Reserve Addition'],
})
if st.button('💾 Save to Google Sheets' if source.startswith('Google') else '💾 Save locally (Excel 模式)',disabled=period_mismatch):
    if status=='Confirmed' and (not income_file or not rent_file):st.error('Confirmed 必须提供两份 AppFolio 报表。')
    else:
        try:
            if source.startswith('Google'):
                store_google(save_rec)
                st.session_state['google_data']=google_load()
                st.success('已保存计算结果到 Google Sheets，并刷新数据。')
            else:store_local(save_rec);st.success('已保存计算结果到本机 SQLite 数据库。Excel 文件本身未被修改；切换电脑不会自动同步。')
        except Exception as e:st.error(f'保存失败：{e}')

st.subheader('4 · Investor Report 下载')
st.caption('自动经营摘要：入住率环比 + 可分配现金/季度累计；额外备注读取 Summaries（Property、Month、Summary CN、Summary EN）。Properties 可新增 Distribution Frequency 列，填写 Monthly 或 Quarterly。')
st.caption('报告收入拆分：Rent Income ${:,.2f} · Utility Income ${:,.2f} · Other Income ${:,.2f}'.format(r['Rent Income'],r['Utility Income'],r['Other Income']))
if not income_file or not rent_file:st.info('未同时上传两份 AppFolio PDF 报表，以下报告为待核对草稿。')
if period_mismatch:st.error('报表月份不一致：请先修改 Reporting Month，下载功能已停用。')
lang=str(p.get('Report Language','English')).strip().lower()
languages=['中文','English'] if lang in ('both','bilingual','中英文') else (['中文'] if lang in ('中文','chinese','cn') else ['English'])
for language in languages:
    cn=language=='中文';stem=f"Moo_{selected}_{month}_{'CN' if cn else 'EN'}"
    try:
        if period_mismatch:continue
        appendix_documents=[]
        if income_file: appendix_documents.append(('AppFolio Income Statement', month))
        if rent_file: appendix_documents.append(('AppFolio Rent Roll', 'As of '+(month+' month-end')))
        d=make_docx(r,p,cn,previous=prior_row,appendix_documents=appendix_documents)
        try:
            pdf=make_pdf(r,p,cn,previous=prior_row,appendix_documents=appendix_documents)
            if appendix_files:
                if cn:
                    translated=[]
                    property_name=str(p.get('Property Name / LLC',p.get('Property Name','')))
                    if income_file:
                        translated.append(('物业经营利润表.pdf',translate_income(income_file.getvalue(),property_name,month)))
                    if rent_file:
                        translated.append(('物业出租及入住情况表.pdf',translate_rent_roll(rent_file.getvalue(),property_name,month)))
                    pdf=append_pdf_documents(pdf,translated)
                else:
                    pdf=append_pdf_documents(pdf,[(f.name,f.getvalue()) for f in appendix_files])
        except Exception as pdf_error:
            pdf=None
            st.info(f'Word report ready. PDF unavailable: {pdf_error}')
        left,right=st.columns(2)
        left.download_button(f'{language} · Download Word',d,stem+'.docx',mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',key='doc'+language)
        if pdf is not None:
            right.download_button(f"{language} · Download PDF + Appendix",pdf,stem+".pdf",mime="application/pdf",key="pdf"+language)
    except Exception as e:st.error(f'{language} 报告生成失败：{e}')
op_cn, cash_cn, note_cn = _single_summary_preview(r, prior_row, True)
op_en, cash_en, note_en = _single_summary_preview(r, prior_row, False)
show_communications(str(p.get('Property Name / LLC') or selected), month,
                    op_cn, cash_cn, note_cn, op_en, cash_en, note_en,
                    key='single_'+re.sub(r'[^a-zA-Z0-9]', '_', selected))

st.subheader('5 · 历史月度记录')
if rows:st.dataframe(pd.DataFrame([x for x in rows if x['Property ID']==selected]).sort_values('Month (YYYY-MM)'),hide_index=True,use_container_width=True)
st.caption('内部管理工具。系统不自动付款、不自动发送邮件；现金余额为推算值，确认分红前仍需审核应付款项和实际分红流水。')
