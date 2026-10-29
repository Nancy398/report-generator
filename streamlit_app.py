import io, json, os, re, sqlite3
from datetime import datetime
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Moo Housing | Investor Reporting V3', layout='wide')
SHEETS=['Properties','Investors','Report Periods','Monthly Performance','Distributions','Reserve Setup','Reserve Ledger','Reserve Transactions','Cash Transactions','Cash Summary']
DB=Path('moo_portal_confirmations.sqlite3')

def num(v, default=0.0):
    try:
        if pd.isna(v) or str(v).strip()=='':return default
        return float(str(v).replace(',','').replace('$','').replace('%',''))
    except (ValueError,TypeError):return default

def norm_id(v):
    s=str(v).strip().upper()
    if s in ('','NAN','NONE'):return ''
    s=re.sub(r'[^A-Z0-9]','',s)
    if s.startswith('MH'):s=s[2:]
    return s

def month(v):
    if pd.isna(v):return ''
    s=str(v).strip()
    if re.fullmatch(r'\d{4}-\d{2}',s):return s
    try:return pd.to_datetime(v).strftime('%Y-%m')
    except Exception:return ''

def rows(df):
    if df is None:return []
    return df.fillna('').to_dict('records')

def load_excel(blob):
    x=pd.ExcelFile(io.BytesIO(blob))
    return {name:pd.read_excel(x,sheet_name=name,header=3 if name in ('Reserve Setup','Reserve Ledger') else 2 if name in ('Cash Transactions','Cash Summary') else 0,dtype=object).dropna(how='all') for name in x.sheet_names if name in SHEETS}

def load_google(url, credentials):
    import gspread
    from google.oauth2.service_account import Credentials
    scopes=['https://www.googleapis.com/auth/spreadsheets.readonly','https://www.googleapis.com/auth/drive.readonly']
    client=gspread.authorize(Credentials.from_service_account_info(credentials,scopes=scopes))
    book=client.open_by_url(url)
    output={}
    for ws in book.worksheets():
        if ws.title not in SHEETS:continue
        vals=ws.get_all_values()
        h=3 if ws.title in ('Reserve Setup','Reserve Ledger') else 2 if ws.title in ('Cash Transactions','Cash Summary') else 0
        if len(vals)<=h:continue
        headers=vals[h];data=[(r+['']*len(headers))[:len(headers)] for r in vals[h+1:]]
        df=pd.DataFrame(data,columns=headers)
        output[ws.title]=df.loc[df.apply(lambda row:any(str(v).strip() for v in row),axis=1)].copy()
    return output

def pick(df,pid):return [r for r in rows(df) if norm_id(r.get('Property ID',r.get('Property_ID','')))==pid]

def get_setup(data,pid):
    r=pick(data.get('Reserve Setup'),pid)
    return r[0] if r else {}

def get_property(data,pid):
    r=pick(data.get('Properties'),pid)
    return r[0] if r else {}

def reserve_tx(data,pid,m):
    amounts={'Property Tax':0.,'Insurance':0.,'Other':0.}
    for r in pick(data.get('Reserve Transactions'),pid):
        pm=month(r.get('Reporting Month (Auto)')) or month(r.get('Payment Date (YYYY-MM-DD)'))
        if pm!=m:continue
        typ=str(r.get('Type','')).lower();a=max(0,num(r.get('Amount ($)')))
        if 'tax' in typ:amounts['Property Tax']+=a
        elif 'insurance' in typ:amounts['Insurance']+=a
        else:amounts['Other']+=a
    return amounts

def calc_ledger(data,pid):
    setup=get_setup(data,pid)
    opening_balance=num(setup.get('Opening Reserve Balance ($)'))
    opening_shortfall=num(setup.get('Opening Reserve Shortfall ($)'))
    tax=num(setup.get('Tax Reserve / Period ($)'))
    ins=num(setup.get('Insurance Reserve / Period ($)'))
    # Fallback when Google Sheets formula cells have not yet been recalculated.
    if not tax and num(setup.get('Annual Property Tax ($)')):
        tax=num(setup.get('Annual Property Tax ($)'))/max(1,num(setup.get('Tax Periods'),6))
    if not ins and num(setup.get('Annual Insurance ($)')):
        ins=num(setup.get('Annual Insurance ($)'))/max(1,num(setup.get('Insurance Periods'),12))
    entries={}
    for r in pick(data.get('Reserve Ledger'),pid):
        m=month(r.get('Period (YYYY-MM)'))
        if m:
            if m in entries:raise ValueError(f'{pid} / {m} 在 Reserve Ledger 中重复。请每个物业每月只保留一行。')
            entries[m]=r
    balance=opening_balance;shortfall=opening_shortfall;out=[]
    for m in sorted(entries):
        r=entries[m]
        noi=num(r.get('NOI ($)'))
        prior=max(0,num(r.get('Prior Unrestricted Cash Available ($)')))
        due=tax+ins+shortfall
        from_noi=min(max(0,noi),due)
        from_prior=min(prior,max(0,due-from_noi))
        funded=from_noi+from_prior
        tx=reserve_tx(data,pid,m)
        # Preserve manually entered reserve usage when transaction log is empty; never add both.
        if not sum(tx.values()):
            tx={'Property Tax':num(r.get('Tax Paid from Reserve ($)')),'Insurance':num(r.get('Insurance Paid from Reserve ($)')),'Other':num(r.get('Other Approved Reserve Used ($)'))}
        used=sum(tx.values())
        begin=balance
        balance=begin+funded-used
        new_short=max(0,due-funded)
        issues=[]
        if used>begin+funded+0.005:issues.append('Reserve 使用超过可用储备')
        if noi<0:issues.append('NOI 为负；需核实现金情况')
        confirmed=str(r.get('Cash / Funding Confirmed?','')).strip().lower()=='yes'
        out.append({'Property ID':pid,'Month':m,'NOI':noi,'Prior Cash Input':prior,'Tax Due':tax,'Insurance Due':ins,'Previous Shortfall':shortfall,'Total Due':due,'From NOI':from_noi,'From Prior Cash':from_prior,'Total Funded':funded,'Reserve Used':used,'Beginning Reserve':begin,'Ending Reserve':balance,'Ending Shortfall':new_short,'NOI After Allocation':noi-from_noi,'Sheet Confirmed':confirmed,'Issues':'; '.join(issues)})
        shortfall=new_short
    return pd.DataFrame(out)

def cash_check(data,pid,m,ledger):
    summary=[r for r in pick(data.get('Cash Summary'),pid) if month(r.get('Period (YYYY-MM)'))==m]
    r=summary[0] if summary else {}
    latest=ledger.loc[ledger['Month']==m].iloc[0]
    bank=num(r.get('Ending Bank Balance (verified)'),None)
    minimum=num(r.get('Minimum Operating Cash')) or num(get_setup(data,pid).get('Minimum Operating Cash ($)'))
    restricted=num(r.get('Other Restricted Cash (excluding reserve)'))
    obligations=num(r.get('Unpaid Obligations (not in reserve)'))
    reserve=max(0,float(latest['Ending Reserve']))
    shortfall=float(latest['Ending Shortfall'])
    capacity=None if bank is None else max(0,bank-reserve-minimum-restricted-obligations-shortfall)
    proposed=num(r.get('Proposed Distribution (Report Periods)'))
    if not proposed:
        for x in pick(data.get('Report Periods'),pid):
            if month(x.get('Period End'))==m:proposed+=num(x.get('Proposed Distribution'))
    reconciled=str(r.get('Bank Reconciled?','')).strip().lower()=='yes'
    reserve_confirmed=str(r.get('Reserve Confirmed?','')).strip().lower()=='yes'
    return {'bank':bank,'minimum':minimum,'restricted':restricted,'obligations':obligations,'capacity':capacity,'proposed':proposed,'reconciled':reconciled,'reserve_confirmed':reserve_confirmed,'shortfall':shortfall,'reserve':reserve}

def connect_db():
    db=sqlite3.connect(DB)
    db.execute('CREATE TABLE IF NOT EXISTS approvals (property_id TEXT, period TEXT, confirmed_at TEXT, snapshot TEXT, PRIMARY KEY(property_id,period))')
    return db

st.title('MOO HOUSING | Investor Reporting & Reserve V3')
st.caption('Google Sheets / Excel → Reserve 结转 → 现金与分红检查 → 人工确认。不会自动转账或修改 Google Sheets。')
with st.sidebar:
    st.header('数据来源')
    source=st.radio('选择', ['上传 Excel（先测试）','连接私有 Google Sheets'],index=0)
    data=None
    if source.startswith('上传'):
        upload=st.file_uploader('上传 Master Excel (.xlsx)',type=['xlsx','xlsm'])
        if upload:
            try:data=load_excel(upload.getvalue())
            except Exception as e:st.error(f'读取失败：{e}')
    else:
        url=st.text_input('Google Sheets URL',value=st.secrets.get('google_sheet_url',''))
        try:creds=dict(st.secrets['gcp_service_account'])
        except Exception:creds=None
        if not creds:st.warning('请先在 .streamlit/secrets.toml 配置服务账号。')
        if st.button('读取 Google Sheets') and url and creds:
            try:st.session_state['sheet_data']=load_google(url,creds);st.success('读取成功')
            except Exception as e:st.error(f'读取失败：{e}')
        data=st.session_state.get('sheet_data')
if not data:
    st.info('先上传 V3 Master Excel 测试；连接 Google Sheets 的步骤见 README。')
    st.stop()

props={norm_id(r.get('Property_ID')):r for r in rows(data.get('Properties')) if norm_id(r.get('Property_ID'))}
setups={norm_id(r.get('Property ID')) for r in rows(data.get('Reserve Setup')) if norm_id(r.get('Property ID'))}
all_ids=sorted(set(props)|setups)
if not all_ids:st.error('没有找到 Property ID。');st.stop()
with st.expander('数据检查',expanded=False):
    st.write('已读取工作表：',', '.join(data))
    missing=[(name,sorted({norm_id(r.get('Property ID',r.get('Property_ID',''))) for r in rows(df) if norm_id(r.get('Property ID',r.get('Property_ID','')))}-set(all_ids))) for name,df in data.items()]
    for name,ids in missing:
        if ids:st.warning(f'{name} 存在未在 Properties/Reserve Setup 找到的 ID：{ids}')
    if 'Reserve Setup' in data:
        st.dataframe(data['Reserve Setup'].head(15),hide_index=True)

pid=st.selectbox('Property',all_ids,format_func=lambda x:f'{x} — {get_property(data,x).get("Property / LLC Name",get_setup(data,x).get("Property / LLC",""))}')
try:ledger=calc_ledger(data,pid)
except Exception as e:st.error(str(e));st.stop()
if ledger.empty:st.warning('该项目在 Reserve Ledger 没有月份记录。请先添加 Property ID、Period、NOI。');st.stop()
m=st.selectbox('Reporting Month',ledger['Month'].tolist(),index=len(ledger)-1)
row=ledger.loc[ledger['Month']==m].iloc[0]
st.subheader('Reserve Summary')
a,b,c,d=st.columns(4)
a.metric('Current NOI',f'${row["NOI"]:,.2f}')
b.metric('Funded This Period',f'${row["Total Funded"]:,.2f}')
c.metric('Ending Reserve',f'${row["Ending Reserve"]:,.2f}')
d.metric('Ending Shortfall',f'${row["Ending Shortfall"]:,.2f}')
st.dataframe(pd.DataFrame([row.to_dict()]).T.rename(columns={0:'Amount / Status'}),use_container_width=True)
if row['Issues']:st.error(row['Issues'])
if not row['Sheet Confirmed']:st.warning('Reserve Ledger 中 Cash / Funding Confirmed? 仍为 No：以下结果为计算草稿。')
st.subheader('Cash & Distribution Check')
check=cash_check(data,pid,m,ledger)
x,y,z=st.columns(3)
x.metric('Verified Bank Balance','未填写' if check['bank'] is None else f'${check["bank"]:,.2f}')
y.metric('Available for Distribution','未核对' if check['capacity'] is None else f'${check["capacity"]:,.2f}')
z.metric('Proposed Distribution',f'${check["proposed"]:,.2f}')
checks=[]
if check['bank'] is None:checks.append('未填写银行余额')
if not check['reconciled']:checks.append('银行尚未对账确认')
if not check['reserve_confirmed'] or not row['Sheet Confirmed']:checks.append('Reserve 尚未确认')
if check['shortfall']>0.005:checks.append('存在 Reserve Shortfall')
if check['capacity'] is not None and check['proposed']>check['capacity']+0.005:checks.append('拟分红超过可分配现金')
if row['Issues']:checks.append(row['Issues'])
if checks:st.error('HOLD — '+ '；'.join(checks))
else:st.success('通过表内现金检查；仍须管理层审批，且不代表银行资金已经划拨。')
with st.expander('历史月份 Reserve 结转'):
    st.dataframe(ledger,hide_index=True,use_container_width=True)
    st.download_button('下载 Reserve 计算结果 CSV',ledger.to_csv(index=False).encode('utf-8-sig'),file_name=f'{pid}_reserve_ledger.csv',mime='text/csv')

st.subheader('管理层确认（本机留痕）')
st.caption('确认仅保存在本机 SQLite，不写回 Google Sheets。更新源数据后需重新审核；正式部署请加登录权限和持久化存储。')
with connect_db() as db:
    old=db.execute('SELECT confirmed_at,snapshot FROM approvals WHERE property_id=? AND period=?',(pid,m)).fetchone()
    if old:st.info(f'本机历史确认：{old[0]}。源数据变更后需要重新确认。')
    attested=st.checkbox('我已核对银行余额、Reserve 来源、支付流水及分红金额，并确认本次记录')
    disabled=bool(checks) or not attested
    if st.button('Confirm & Save Snapshot',disabled=disabled):
        payload={'reserve':row.to_dict(),'cash':check}
        db.execute('INSERT OR REPLACE INTO approvals VALUES (?,?,?,?)',(pid,m,datetime.now().isoformat(timespec='seconds'),json.dumps(payload,default=str)))
        db.commit();st.success('已保存本机确认快照。')
