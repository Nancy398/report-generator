import io, re, sqlite3, os, smtplib, ssl
from email.message import EmailMessage
from datetime import datetime
from pathlib import Path
import pandas as pd
import streamlit as st
from docx import Document
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

st.set_page_config(page_title='Moo Housing | Admin Portal',page_icon='🏢',layout='wide')
DB=Path(os.environ.get('MOO_DB_PATH',str(Path(__file__).parent/'moo_portal.sqlite')))
def db():
 c=sqlite3.connect(DB);c.execute('PRAGMA foreign_keys=ON');return c
def init():
 with db() as c:
  c.execute('CREATE TABLE IF NOT EXISTS properties (id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL, tax REAL NOT NULL DEFAULT 0, insurance REAL NOT NULL DEFAULT 0, min_cash REAL NOT NULL DEFAULT 0, opening_tax REAL NOT NULL DEFAULT 0, opening_insurance REAL NOT NULL DEFAULT 0)')
  c.execute('CREATE TABLE IF NOT EXISTS reserve_months (property_id INTEGER NOT NULL, period TEXT NOT NULL, bank_cash REAL NOT NULL, committed REAL NOT NULL, tax_paid REAL NOT NULL, insurance_paid REAL NOT NULL, tax_funded REAL NOT NULL, insurance_funded REAL NOT NULL, tax_shortfall REAL NOT NULL, insurance_shortfall REAL NOT NULL, tax_balance REAL NOT NULL, insurance_balance REAL NOT NULL, notes TEXT, PRIMARY KEY(property_id,period))')
  c.execute('CREATE TABLE IF NOT EXISTS investors (id INTEGER PRIMARY KEY, name TEXT NOT NULL, email TEXT)')
  c.execute('CREATE TABLE IF NOT EXISTS investments (id INTEGER PRIMARY KEY, investor_id INTEGER NOT NULL, property_id INTEGER NOT NULL, invested REAL NOT NULL DEFAULT 0, ownership REAL NOT NULL DEFAULT 0, distribution_pct REAL NOT NULL DEFAULT 0, UNIQUE(investor_id,property_id))')
  c.execute("CREATE TABLE IF NOT EXISTS report_preferences (property_id INTEGER PRIMARY KEY, language TEXT NOT NULL DEFAULT 'English', report_type TEXT NOT NULL DEFAULT 'Monthly', recipient_emails TEXT NOT NULL DEFAULT '', wechat_group TEXT NOT NULL DEFAULT '')")
  c.execute('CREATE TABLE IF NOT EXISTS distributions (id INTEGER PRIMARY KEY, investor_id INTEGER NOT NULL, property_id INTEGER NOT NULL, paid_on TEXT NOT NULL, amount REAL NOT NULL, notes TEXT)')
init()
def amount(v):
 try:return float(v) if pd.notna(v) else 0.
 except:return 0.
def dollars(v):return f'${v:,.2f}'
def read_appfolio(upload):
 if upload is None:return None
 try:
  df=pd.read_excel(upload,header=None,dtype=object)
  return df.fillna('')
 except Exception as e:st.error(f'Cannot read Excel: {e}');return None
def parse_income(df):
 if df is None:return None
 data={};meta=[]
 for _,row in df.iterrows():
  label=str(row.iloc[0]).strip() if len(row)>0 else ''
  if not label:continue
  if any(s in label for s in ('Owned By:','Period Range:')):meta.append(label)
  val=amount(row.iloc[1]) if len(row)>1 else 0.
  data[label]=val
 def find(keys):
  for key in keys:
   for k,v in data.items():
    if k.lower()==key.lower():return v
  return None
 income=find(['Total Operating Income']);expenses=find(['Total Operating Expense','Total Operating Expenses']);noi=find(['Net Operating Income','NOI','Net Income'])
 if noi is None and income is not None and expenses is not None:noi=income-expenses
 return {'income':income,'expenses':expenses,'noi':noi,'metadata':meta,'accounts':data,'rent_income':find(['Rent Income']),'concessions':find(['Concessions'])}
def parse_roll(df):
 if df is None:return None
 header=None
 for i,row in df.iterrows():
  if str(row.iloc[0]).strip().lower()=='unit' and len(row)>2 and str(row.iloc[1]).strip().lower()=='tenant':header=i;break
 if header is None:raise ValueError('Rent Roll header Unit / Tenant not found')
 groups={};unmatched=[];current=None
 for i in range(header+1,len(df)):
  row=df.iloc[i];unit=str(row.iloc[0]).strip();tenant=str(row.iloc[1]).strip() if len(row)>1 else '';status=str(row.iloc[2]).strip() if len(row)>2 else '';rent=amount(row.iloc[3]) if len(row)>3 else 0
  if not unit or unit.lower().startswith(('total','units:')):continue
  if ' - ' not in unit:continue
  prefix,suffix=unit.split(' - ',1);prefix=prefix.strip();suffix=suffix.strip()
  if re.match(r'(?i)^rm\s*\d+',suffix):
   if prefix not in groups:groups[prefix]={'master':None,'rooms':[]}
   groups[prefix]['rooms'].append({'unit':unit,'tenant':tenant,'status':status,'rent':rent})
  elif prefix.lower()==suffix.lower():
   if prefix not in groups:groups[prefix]={'master':None,'rooms':[]}
   groups[prefix]['master']={'unit':unit,'tenant':tenant,'status':status,'rent':rent}
  else:unmatched.append(unit)
 occupied=0;rooms=0;monthly_rent=0;detail=[];warnings=[]
 for prefix,g in groups.items():
  master=g['master'];room_list=g['rooms'];master_active=bool(master and master['tenant'] and master['status'].lower() not in ('vacant','vacant-unrented'))
  if not room_list:
   warnings.append(f'{prefix}: master unit without identifiable rooms; excluded from room occupancy denominator')
   continue
  room_occupied=sum(bool(r['tenant'] and r['status'].lower() not in ('vacant','vacant-unrented')) for r in room_list)
  if master_active:
   n=len(room_list);rent=master['rent']
   if room_occupied: warnings.append(f'{prefix}: master and {room_occupied} room(s) both have tenants; review overlap')
   if any(r['rent'] for r in room_list):warnings.append(f'{prefix}: room rent exists under occupied master; excluded to avoid double counting')
  else:
   n=room_occupied;rent=sum(r['rent'] for r in room_list if r['tenant'] and r['status'].lower() not in ('vacant','vacant-unrented'))
  occupied+=n;rooms+=len(room_list);monthly_rent+=rent
  detail.append({'Building':prefix,'Total Rooms':len(room_list),'Occupied Rooms':n,'Occupancy':f'{n/len(room_list):.1%}','Master Leased':master_active,'Scheduled Monthly Rent':rent})
 if unmatched:warnings.append(f'{len(unmatched)} unrecognized unit rows excluded; verify naming format')
 return {'rooms':rooms,'occupied':occupied,'rate':occupied/rooms if rooms else None,'scheduled_rent':monthly_rent,'detail':detail,'warnings':warnings}
def get_properties():
 with db() as c:return c.execute('SELECT id,name,tax,insurance,min_cash,opening_tax,opening_insurance FROM properties ORDER BY name').fetchall()
def history(pid):
 with db() as c:return c.execute('SELECT period,bank_cash,committed,tax_paid,insurance_paid,tax_funded,insurance_funded,tax_shortfall,insurance_shortfall,tax_balance,insurance_balance,notes FROM reserve_months WHERE property_id=? ORDER BY period',(pid,)).fetchall()
def build_docx(name,period,financial,roll,reserve):
 d=Document();d.add_heading('MOO HOUSING INC.',0);d.add_heading('Property Investor Report',1);d.add_paragraph(f'{name} | {period}')
 d.add_heading('Financial Performance',2)
 for label,value in [('Total Operating Income',financial.get('income')),('Total Operating Expenses',financial.get('expenses')),('Net Operating Income',financial.get('noi'))]:d.add_paragraph(f'{label}: {dollars(value) if value is not None else "Not available"}')
 d.add_heading('Occupancy',2)
 d.add_paragraph(f"Occupied Rooms: {roll['occupied']} / {roll['rooms']} ({roll['rate']:.1%})" if roll and roll['rate'] is not None else 'Rent Roll not uploaded')
 if roll:d.add_paragraph(f"Scheduled Monthly Rent (Rent Roll): {dollars(roll['scheduled_rent'])}")
 d.add_heading('Reserve Management',2)
 for label,value in reserve.items():d.add_paragraph(f'{label}: {dollars(value)}')
 d.add_paragraph('Draft for internal review. Reserve allocations are cash designations, not additional P&L expenses. Scheduled rent is not recognized rental income.')
 bio=io.BytesIO();d.save(bio);return bio.getvalue()
def build_pdf(name,period,financial,roll,reserve):
 bio=io.BytesIO();styles=getSampleStyleSheet();story=[Paragraph('MOO HOUSING INC. | INVESTOR REPORT',styles['Title']),Paragraph(f'{name} | {period}',styles['Heading2']),Spacer(1,14)]
 items=[('Total Operating Income',financial.get('income')),('Total Operating Expenses',financial.get('expenses')),('Net Operating Income',financial.get('noi'))]
 if roll and roll['rate'] is not None:items += [('Room Occupancy',f"{roll['occupied']}/{roll['rooms']} ({roll['rate']:.1%})"),('Scheduled Monthly Rent',roll['scheduled_rent'])]
 items += list(reserve.items())
 cells=[['Metric','Value']]+[[k,dollars(v) if isinstance(v,(int,float)) else ('N/A' if v is None else str(v))] for k,v in items]
 table=Table(cells,colWidths=[300,170]);table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#17324D')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F1F5F9')]),('BOTTOMPADDING',(0,0),(-1,-1),9),('TOPPADDING',(0,0),(-1,-1),9)]));story.append(table)
 story.append(Spacer(1,15));story.append(Paragraph('Internal draft: confirm cash balances, reserves, and distributions before publication. Scheduled rent is not recognized income.',styles['Normal']))
 SimpleDocTemplate(bio).build(story);return bio.getvalue()

st.title('🏢 Moo Housing | Admin Portal')
st.caption('Phase 1 • AppFolio Excel imports • Master lease occupancy • Reserve management • Draft investor reports')
tabs=st.tabs(['Property Setup','AppFolio Import','Reserve Management','Investor Records','Generate Report'])
with tabs[0]:
 st.subheader('Property / LLC Settings')
 with st.form('property'):
  pname=st.text_input('Property / LLC Name',value='1153 37th Dr Development LLC')
  a,b,c=st.columns(3)
  tax=a.number_input('Monthly Property Tax Reserve ($)',min_value=0.,value=0.,step=100.)
  ins=b.number_input('Monthly Insurance Reserve ($)',min_value=0.,value=0.,step=100.)
  minimum=c.number_input('Minimum Operating Cash ($)',min_value=0.,value=0.,step=1000.)
  d,e=st.columns(2)
  opening_tax=d.number_input('Opening Tax Reserve ($)',min_value=0.,value=0.)
  opening_ins=e.number_input('Opening Insurance Reserve ($)',min_value=0.,value=0.)
  save=st.form_submit_button('Save / Update Property')
 if save and pname.strip():
  with db() as conn:
   conn.execute('INSERT INTO properties(name,tax,insurance,min_cash,opening_tax,opening_insurance) VALUES(?,?,?,?,?,?) ON CONFLICT(name) DO UPDATE SET tax=excluded.tax,insurance=excluded.insurance,min_cash=excluded.min_cash,opening_tax=excluded.opening_tax,opening_insurance=excluded.opening_insurance',(pname.strip(),tax,ins,minimum,opening_tax,opening_ins))
  st.success('Property settings saved.')
 props=get_properties()
 if props:st.dataframe(pd.DataFrame(props,columns=['ID','LLC','Monthly Tax','Monthly Insurance','Min Cash','Opening Tax','Opening Insurance']),hide_index=True,use_container_width=True)
with tabs[1]:
 st.subheader('Upload AppFolio Exports')
 f1=st.file_uploader('Income Statement (.xlsx)',type=['xlsx'],key='income')
 f2=st.file_uploader('Rent Roll (.xlsx)',type=['xlsx'],key='roll')
 financial=parse_income(read_appfolio(f1)) if f1 else None
 roll=None
 if f2:
  try:roll=parse_roll(read_appfolio(f2))
  except Exception as ex:st.error(f'Rent Roll parsing failed: {ex}')
 if financial:
  st.write('**Income Statement**')
  for k in financial['metadata']:st.caption(k)
  x,y,z=st.columns(3)
  x.metric('Operating Income',dollars(financial['income']) if financial['income'] is not None else 'Missing')
  y.metric('Operating Expenses',dollars(financial['expenses']) if financial['expenses'] is not None else 'Missing')
  z.metric('NOI',dollars(financial['noi']) if financial['noi'] is not None else 'Missing')
  if financial['income'] is not None and financial['expenses'] is not None and financial['noi'] is not None and abs(financial['income']-financial['expenses']-financial['noi'])>0.02:st.error('NOI reconciliation failed: check report before publishing')
 if roll:
  st.write('**Room-level Occupancy (Master Lease Aware)**')
  a,b,c=st.columns(3);a.metric('Occupied Rooms',f"{roll['occupied']} / {roll['rooms']}");b.metric('Occupancy',f"{roll['rate']:.1%}" if roll['rate'] is not None else 'N/A');c.metric('Scheduled Monthly Rent',dollars(roll['scheduled_rent']))
  st.dataframe(pd.DataFrame(roll['detail']),hide_index=True,use_container_width=True)
  for warning in roll['warnings']:st.warning(warning)
 st.info('Uploads are processed in this session only. To generate a report, keep files uploaded and switch to the Generate Report tab.')
with tabs[2]:
 st.subheader('Monthly Reserve Close')
 props=get_properties()
 if not props:st.info('Create a property in Property Setup first.')
 else:
  chosen=st.selectbox('Property',props,format_func=lambda p:p[1],key='reserve_property');pid,pname,tax,ins,minimum,ot,oi=chosen
  hist=history(pid)
  st.caption(f'Fixed monthly reserve: Tax {dollars(tax)} | Insurance {dollars(ins)} | Minimum operating cash {dollars(minimum)}')
  period=st.text_input('Closing period (YYYY-MM)',value=datetime.today().strftime('%Y-%m'))
  previous=[r for r in hist if r[0]<period]
  latest=previous[-1] if previous else None
  prev_tax_balance=latest[9] if latest else ot;prev_ins_balance=latest[10] if latest else oi
  prev_tax_gap=latest[7] if latest else 0.;prev_ins_gap=latest[8] if latest else 0.
  a,b=st.columns(2)
  bank=a.number_input('Unrestricted available bank cash at closing ($)',min_value=0.,value=0.,step=1000.)
  committed=b.number_input('Unpaid committed obligations ($)',min_value=0.,value=0.,step=500.)
  c,d=st.columns(2)
  tax_paid=c.number_input('Tax paid from reserved cash this month ($)',min_value=0.,value=0.)
  ins_paid=d.number_input('Insurance paid from reserved cash this month ($)',min_value=0.,value=0.)
  notes=st.text_input('Close notes / adjustments')
  due_tax=tax+prev_tax_gap;due_ins=ins+prev_ins_gap
  # Existing reserves are ring-fenced: do not count them as free cash.
  available=max(0.,bank-minimum-committed-prev_tax_balance-prev_ins_balance)
  funded_tax=min(available,due_tax);funded_ins=min(max(0.,available-funded_tax),due_ins)
  # Reserve spending occurs after beginning balances; cannot spend more than held + funded.
  balance_tax=prev_tax_balance+funded_tax-tax_paid;balance_ins=prev_ins_balance+funded_ins-ins_paid
  gap_tax=due_tax-funded_tax;gap_ins=due_ins-funded_ins
  if balance_tax<0 or balance_ins<0:st.error('Reserve payments exceed available designated reserve balance; correct payment or record funding first.')
  st.write('**Calculated close (tax funded first; insurance second)**')
  st.dataframe(pd.DataFrame([{'Category':'Property Tax','Prior Balance':prev_tax_balance,'Monthly Required':tax,'Prior Shortfall':prev_tax_gap,'New Funding':funded_tax,'Paid':tax_paid,'Ending Balance':balance_tax,'Shortfall':gap_tax},{'Category':'Insurance','Prior Balance':prev_ins_balance,'Monthly Required':ins,'Prior Shortfall':prev_ins_gap,'New Funding':funded_ins,'Paid':ins_paid,'Ending Balance':balance_ins,'Shortfall':gap_ins}]),hide_index=True,use_container_width=True)
  st.metric('Preliminary unallocated cash after funding',dollars(max(0.,available-funded_tax-funded_ins)))
  if gap_tax+gap_ins>0:st.warning(f'Reserve shortfall: {dollars(gap_tax+gap_ins)}. No automatic distribution recommendation.')
  if any(r[0]==period for r in hist):st.warning('This period is already closed. Saving will replace its entry; later periods must be recalculated manually.')
  if st.button('Save Monthly Reserve Close',disabled=balance_tax<0 or balance_ins<0):
   if not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])',period):st.error('Period must be YYYY-MM')
   elif any(r[0]>period for r in hist):st.error('Cannot modify an earlier month after later months have been closed. Reopen/recalculate chronologically.')
   else:
    with db() as conn:conn.execute('INSERT OR REPLACE INTO reserve_months VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)',(pid,period,bank,committed,tax_paid,ins_paid,funded_tax,funded_ins,gap_tax,gap_ins,balance_tax,balance_ins,notes))
    st.success('Reserve close saved.');st.rerun()
  if hist:st.dataframe(pd.DataFrame(hist,columns=['Period','Bank Cash','Committed','Tax Paid','Insurance Paid','Tax Funded','Insurance Funded','Tax Shortfall','Insurance Shortfall','Tax Balance','Insurance Balance','Notes']),hide_index=True,use_container_width=True)
  st.caption('Cash model assumes bank cash is unrestricted, as of the same closing date, and includes funds already earmarked for reserves. Exclude escrow/restricted accounts and avoid double counting pending payments. Funding is an internal designation, not a P&L expense.')
with tabs[3]:
 st.subheader('Investor & Investment Records (Internal Only)')
 props=get_properties()
 with st.form('investor'):
  iname=st.text_input('Investor name');email=st.text_input('Investor email')
  if props:invest_property=st.selectbox('Investment LLC',props,format_func=lambda p:p[1])
  a,b,c=st.columns(3)
  invested=a.number_input('Invested capital ($)',min_value=0.,value=0.);ownership=b.number_input('Ownership (%)',min_value=0.,max_value=100.,value=0.);dist=c.number_input('Distribution (%)',min_value=0.,max_value=100.,value=0.)
  submit=st.form_submit_button('Save Investor / Investment')
 if submit:
  if not iname.strip() or not props:st.error('Name and LLC required')
  else:
   with db() as conn:
    row=conn.execute('SELECT id FROM investors WHERE name=?',(iname.strip(),)).fetchone()
    if row:inv_id=row[0];conn.execute('UPDATE investors SET email=? WHERE id=?',(email,inv_id))
    else:inv_id=conn.execute('INSERT INTO investors(name,email) VALUES(?,?)',(iname.strip(),email)).lastrowid
    conn.execute('INSERT INTO investments(investor_id,property_id,invested,ownership,distribution_pct) VALUES(?,?,?,?,?) ON CONFLICT(investor_id,property_id) DO UPDATE SET invested=excluded.invested,ownership=excluded.ownership,distribution_pct=excluded.distribution_pct',(inv_id,invest_property[0],invested,ownership,dist))
   st.success('Investor saved')
 with db() as conn:
  data=conn.execute('SELECT i.name,i.email,p.name,v.invested,v.ownership,v.distribution_pct FROM investments v JOIN investors i ON i.id=v.investor_id JOIN properties p ON p.id=v.property_id ORDER BY i.name').fetchall()
 if data:st.dataframe(pd.DataFrame(data,columns=['Investor','Email','LLC','Invested','Ownership %','Distribution %']),hide_index=True,use_container_width=True)
 st.warning('This is an internal prototype, not a secured investor website. Do not deploy publicly with real investor personal information until authentication, role-based authorization, encryption, backups, and audit logs are added.')
with tabs[4]:
 st.subheader('Draft Investor Report')
 props=get_properties()
 if not props:st.info('Add a property first.')
 else:
  selected=st.selectbox('Report LLC',props,format_func=lambda p:p[1],key='report_property')
  report_period=st.text_input('Reporting period',value='September 2026')
  if not financial:st.warning('Upload an AppFolio Income Statement in the AppFolio Import tab first.')
  else:
   h=history(selected[0]);last=h[-1] if h else None
   reserve={'Monthly Tax Reserve Target':selected[2],'Monthly Insurance Reserve Target':selected[3],'Minimum Operating Cash':selected[4],'Tax Reserve Balance':last[9] if last else selected[5],'Insurance Reserve Balance':last[10] if last else selected[6],'Tax Reserve Shortfall':last[7] if last else 0.,'Insurance Reserve Shortfall':last[8] if last else 0.}
   doc=build_docx(selected[1],report_period,financial,roll,reserve)
   pdf=build_pdf(selected[1],report_period,financial,roll,reserve)
   a,b=st.columns(2)
   a.download_button('Download Word Draft',doc,file_name='Moo_Investor_Report_Draft.docx',mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
   b.download_button('Download PDF Draft',pdf,file_name='Moo_Investor_Report_Draft.pdf',mime='application/pdf')
   st.info('Draft reports require administrator verification and approval. They are not published to investors.')


# --- Report Studio: bilingual project preferences, controlled delivery ---
def get_preference(pid):
 with db() as c:
  row=c.execute('SELECT language,report_type,recipient_emails,wechat_group FROM report_preferences WHERE property_id=?',(pid,)).fetchone()
 return row or ('English','Monthly','','')

def safe_num(v):
 return dollars(v) if v is not None else 'Not available / 未提供'

def report_text(property_name,period,financial,roll,reserve,lang,notes):
 income,expense,noi=financial.get('income'),financial.get('expenses'),financial.get('noi')
 occ=f"{roll['occupied']}/{roll['rooms']} ({roll['rate']:.1%})" if roll and roll.get('rate') is not None else None
 if lang=='中文':
  lines=[f'【{property_name}｜{period} 经营报告】','', '一、财务表现',f'营业收入：{safe_num(income)}',f'运营费用：{safe_num(expense)}',f'净运营收入（NOI）：{safe_num(noi)}','', '二、出租情况',f'房间入住率：{occ or "未提供 Rent Roll"}','', '三、税费及保险储备',f'每月房产税预留：{dollars(reserve["Monthly Tax Reserve Target"])}',f'每月保险预留：{dollars(reserve["Monthly Insurance Reserve Target"])}',f'房产税储备余额：{dollars(reserve["Tax Reserve Balance"])}',f'保险储备余额：{dollars(reserve["Insurance Reserve Balance"])}',f'尚未补足的储备缺口：{dollars(reserve["Tax Reserve Shortfall"]+reserve["Insurance Reserve Shortfall"])}']
  if notes.strip():lines+=['','四、运营说明',notes.strip()]
  lines+=['','注：储备金为资金安排，不是新增的利润表费用。报告须经 Moo Housing 审核后发送。']
 else:
  lines=[f'{property_name} | {period} Investor Report','', 'Financial Performance',f'Total Operating Income: {safe_num(income)}',f'Total Operating Expenses: {safe_num(expense)}',f'Net Operating Income (NOI): {safe_num(noi)}','', 'Occupancy',f'Room Occupancy: {occ or "Rent Roll not provided"}','', 'Property Tax & Insurance Reserves',f'Monthly Property Tax Reserve: {dollars(reserve["Monthly Tax Reserve Target"])}',f'Monthly Insurance Reserve: {dollars(reserve["Monthly Insurance Reserve Target"])}',f'Tax Reserve Balance: {dollars(reserve["Tax Reserve Balance"])}',f'Insurance Reserve Balance: {dollars(reserve["Insurance Reserve Balance"])}',f'Outstanding Reserve Shortfall: {dollars(reserve["Tax Reserve Shortfall"]+reserve["Insurance Reserve Shortfall"])}']
  if notes.strip():lines+=['','Operations Update',notes.strip()]
  lines+=['','Note: Reserves are designated cash, not additional P&L expenses. Subject to Moo Housing review before distribution.']
 return '\n'.join(lines)

def send_smtp(host,port,username,password,sender,recipients,subject,body,attachment_name=None,attachment=None):
 msg=EmailMessage();msg['From']=sender;msg['To']=', '.join(recipients);msg['Subject']=subject;msg.set_content(body)
 if attachment:
  msg.add_attachment(attachment,maintype='application',subtype='vnd.openxmlformats-officedocument.wordprocessingml.document',filename=attachment_name or 'Investor_Report.docx')
 with smtplib.SMTP_SSL(host,port,timeout=20,context=ssl.create_default_context()) as server:
  server.login(username,password);server.send_message(msg)

st.divider()
st.header('✉️ Report Studio | 中英文报告与发送')
st.caption('按项目保存语言和收件人设置；先审核报告，再选择邮件发送或复制微信文案。微信不支持直接自动发送（除非另外接入获授权的企业微信接口）。')
props=get_properties()
if not props:
 st.info('请先在 Property Setup 新增项目。')
else:
 selected=st.selectbox('Project / LLC',props,format_func=lambda p:p[1],key='studio_project')
 pid=selected[0]; pref=get_preference(pid)
 with st.form('studio_settings'):
  language=st.selectbox('Default language / 默认语言',['English','中文','Both'],index=['English','中文','Both'].index(pref[0]) if pref[0] in ('English','中文','Both') else 0)
  report_type=st.selectbox('Report frequency',['Monthly','Quarterly'],index=0 if pref[1]=='Monthly' else 1)
  recipients=st.text_area('Default email recipients (comma separated)',value=pref[2])
  wechat_group=st.text_input('WeChat group / 联系人备注（仅用于内部记录）',value=pref[3])
  if st.form_submit_button('Save project preferences / 保存设置'):
   with db() as c:c.execute('INSERT INTO report_preferences(property_id,language,report_type,recipient_emails,wechat_group) VALUES(?,?,?,?,?) ON CONFLICT(property_id) DO UPDATE SET language=excluded.language,report_type=excluded.report_type,recipient_emails=excluded.recipient_emails,wechat_group=excluded.wechat_group',(pid,language,report_type,recipients,wechat_group))
   st.success('Saved.');st.rerun()
 st.caption('上传文件请使用上方 AppFolio Import 标签。')
 if not financial:
  st.warning('请先上传 Income Statement Excel；没有财务数据不会生成可发送的报告。')
 else:
  h=history(pid);last=h[-1] if h else None
  reserve={'Monthly Tax Reserve Target':selected[2],'Monthly Insurance Reserve Target':selected[3],'Minimum Operating Cash':selected[4],'Tax Reserve Balance':last[9] if last else selected[5],'Insurance Reserve Balance':last[10] if last else selected[6],'Tax Reserve Shortfall':last[7] if last else 0.,'Insurance Reserve Shortfall':last[8] if last else 0.}
  period=st.text_input('Report period / 报告期间',value=datetime.today().strftime('%B %Y'),key='studio_period')
  notes_en=st.text_area('Operations notes (English)',key='studio_en',placeholder='Leasing, maintenance, collection updates; do not invent details.')
  notes_cn=st.text_area('运营备注（中文）',key='studio_cn',placeholder='填写真实运营情况；未提供的信息不会自动编造。')
  language,report_type,recipients,_=get_preference(pid)
  eng=report_text(selected[1],period,financial,roll,reserve,'English',notes_en)
  chi=report_text(selected[1],period,financial,roll,reserve,'中文',notes_cn)
  content=(eng if language=='English' else chi if language=='中文' else eng+'\n\n'+'='*25+'\n\n'+chi)
  st.subheader('Preview / 报告预览')
  st.text_area('Investor report draft',value=content,height=440,key='studio_preview')
  # Export actual bilingual copy to editable Word, avoiding font issues in PDFs.
  d=Document();d.add_heading('MOO HOUSING INC.',0)
  for para in content.split('\n'):
   if para.strip():d.add_paragraph(para)
  out=io.BytesIO();d.save(out);doc_bytes=out.getvalue()
  st.download_button('⬇️ Download bilingual Word / 下载 Word',doc_bytes,file_name=f'Investor_Report_{pid}_{period.replace(" ","_")}.docx',mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
  st.download_button('⬇️ Download email / WeChat text (.txt)',content.encode('utf-8-sig'),file_name=f'Investor_Message_{pid}.txt',mime='text/plain')
  st.markdown('**Email / 邮件发送**')
  st.caption('邮件仅在你主动确认并点击发送后发出。SMTP 凭据只在本次会话中使用，不写入数据库。')
  to=st.text_input('To (comma-separated)',value=recipients,key='studio_to')
  subject=st.text_input('Subject',value=f'{selected[1]} — {period} Investor Report',key='studio_subject')
  with st.expander('SMTP configuration (company mail server / app password)'):
   smtp_host=st.text_input('SMTP host',key='smtp_host')
   smtp_port=st.number_input('SMTP SSL port',min_value=1,max_value=65535,value=465,key='smtp_port')
   smtp_user=st.text_input('SMTP username',key='smtp_user')
   smtp_password=st.text_input('SMTP password / app password',type='password',key='smtp_pass')
   sender=st.text_input('From address',key='smtp_sender')
  reconcile=(financial['income'] is not None and financial['expenses'] is not None and financial['noi'] is not None and abs(financial['income']-financial['expenses']-financial['noi'])<=0.02)
  if not reconcile:st.error('Financial totals missing or inconsistent. Email sending disabled.')
  if roll and roll['warnings']:st.warning('Rent Roll has review warnings. Confirm before sending.')
  reviewed=st.checkbox('I verified financial data, occupancy, reserve balances and recipients / 已人工核对数据及收件人',key='studio_review')
  if st.button('📧 Send approved email / 发送邮件',disabled=not(reconcile and reviewed),key='studio_send'):
   addresses=[x.strip() for x in re.split('[,;\n]',to) if x.strip()]
   if not addresses or any(not re.fullmatch(r'[^@\s]+@[^@\s]+\.[^@\s]+',a) for a in addresses):st.error('Invalid recipient email addresses.')
   elif not all((smtp_host,smtp_user,smtp_password,sender)):st.error('Please complete SMTP settings.')
   else:
    try:
     send_smtp(smtp_host,int(smtp_port),smtp_user,smtp_password,sender,addresses,subject,content,'Investor_Report.docx',doc_bytes)
     st.success(f'Email sent to {len(addresses)} recipient(s).')
    except Exception as e:st.error(f'Sending failed. Check SMTP settings and access. Error type: {type(e).__name__}')
  st.markdown('**WeChat / 微信**')
  st.text_area('复制以下内容到微信（无法直接向个人微信自动发送）',value=chi,height=300,key='wechat_copy')
