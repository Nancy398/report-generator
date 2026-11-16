"""Direct-upload Combined LLC workflow; no Monthly Statements pre-save required."""
import io, re
from datetime import datetime
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Pt
from pdf_appfolio import parse_income_pdf, parse_rent_roll_pdf
from reporting import make_docx, make_pdf, append_pdf_documents
from engine import number, pid, period, calculate
from investor_communication import show_communications
from combined_chapters import make_chapter_combined_docx


def _n(v):
    if v is None or str(v).strip() == '': return None
    try: return float(str(v).replace('$','').replace(',',''))
    except (ValueError,TypeError): return None


def _money(v): return 'Not provided' if v is None else f'${v:,.2f}'


def _summary_line(item, cn):
    r=item['row']; occ=item['occ']['occupancy']*100
    name=item['name']
    return (f'{name}：入住率 {occ:.1f}%，NOI {_money(r["NOI"])}。' if cn else
            f'{name}: occupancy {occ:.1f}%; NOI {_money(r["NOI"])}.')

def _parse_uploads(members,month):
    entries=[];missing=[];errors=[]
    for p in members:
        key=p['id'];inc=st.session_state.get(f'combined_income_{key}');roll=st.session_state.get(f'combined_roll_{key}')
        if not inc or not roll:
            missing.append(p['name']);continue
        try:
            f=parse_income_pdf(inc.getvalue());o=parse_rent_roll_pdf(roll.getvalue())
            for label,d in [('Income Statement',f),('Rent Roll',o)]:
                if d.get('period') and d['period']!=month:
                    raise ValueError(f'{label} month {d["period"]} differs from {month}')
            if any(f.get(k) is None for k in ('income','expenses','noi')):
                raise ValueError('Missing Total Income, Expenses or NOI')
            entries.append(dict(id=key,name=p['name'],financial=f,occ=o,
                                income_pdf=inc.getvalue(),rent_pdf=roll.getvalue()))
        except Exception as exc: errors.append(f'{p["name"]}: {exc}')
    return entries,missing,errors


def show_combined_upload_workflow(google_load,records,store_google,save_reports):
    st.subheader('Combined LLC Report · Upload PDFs')
    st.caption('Select LLC → Upload AppFolio PDFs → Review & Edit → Cash / Reserve → Preview → Word / PDF + Appendices → Investor Communication')
    @st.cache_data(ttl=300,show_spinner=False)
    def directory():
        data=google_load();result=[];seen=set()
        for r in records(data.get('Properties')):
            group=str(r.get('Property Name / LLC') or '').strip()
            name=str(r.get('Property Address') or '').strip()
            k=pid(r.get('Property ID'))
            if group and name and k and k not in seen:
                seen.add(k);result.append(dict(group=group,name=name,id=k,settings=r))
        return result
    try: all_props=directory()
    except Exception as exc:
        st.error(f'无法读取 Reports → Properties：{exc}');return
    groups=sorted({x['group'] for x in all_props},key=str.casefold)
    llc=st.selectbox('Select LLC',groups,index=None,placeholder='Choose an LLC',key='upload_combined_llc')
    month=st.text_input('Reporting Month (YYYY-MM)',datetime.now().strftime('%Y-%m'),key='upload_combined_month')
    if not llc:return
    if not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])',month):
        st.error('Reporting Month must be YYYY-MM');return
    members=[p for p in all_props if p['group']==llc]
    # Read saved history from Reports → Monthly Statements. The previous month is
    # calculated with the SAME roll-forward engine used by Single Property Report.
    history_rows=[];history_calculated=[];history_issues=[]
    try:
        master=google_load()
        history_rows=records(master.get('Monthly Statements'))
        history_calculated,history_issues=calculate(records(master.get('Properties')),history_rows)
    except Exception as exc:
        st.warning(f'历史 Monthly Statements 暂时无法读取：{exc}')
    year_num,month_num=map(int,month.split('-'))
    previous_month=f'{year_num-1:04d}-12' if month_num==1 else f'{year_num:04d}-{month_num-1:02d}'
    def history_for(prop_id):
        matches=[r for r in history_calculated if pid(r.get('Property ID'))==prop_id and period(r.get('Month (YYYY-MM)'))==previous_month]
        if not matches:return None
        row=dict(matches[-1])
        raw=next((r for r in history_rows if pid(r.get('Property ID'))==prop_id and period(r.get('Month (YYYY-MM)'))==previous_month),{})
        for field in ('Rent Income','Outstanding Rent Receivable','Occupancy','Occupied Rooms','Total Rooms'):
            if row.get(field) is None or str(row.get(field))=='':
                row[field]=raw.get(field)
        return row
    st.caption(f'{len(members)} properties found under {llc}.')
    if len(members)<2:
        st.warning('Combined Report requires at least two properties.');return
    for i,p in enumerate(members,1):
        st.markdown(f'**Property {i}: {p["name"]}**')
        a,b=st.columns(2)
        with a:st.file_uploader('Income Statement PDF',type=['pdf'],key=f'combined_income_{p["id"]}')
        with b:st.file_uploader('Rent Roll PDF',type=['pdf'],key=f'combined_roll_{p["id"]}')
    entries,missing,errors=_parse_uploads(members,month)
    for msg in errors:st.error(msg)
    if missing:
        st.info('请上传全部物业的两份 PDF 后继续。')
        return
    if errors:return
    st.subheader('1 · Review AppFolio Data / 核对识别结果')
    st.caption('以下数值可以像单物业报告一样修改。Income − Expenses 应等于 NOI；Rent + Utility + Other 应等于 Total Income。')
    edited=[]
    for item in entries:
        f=item['financial'];o=item['occ'];k=item['id']
        with st.expander(item['name'],expanded=True):
            a,b,c=st.columns(3)
            with a:
                rent=st.number_input('Rent Income',value=float(f.get('rent_income') or 0),format='%.2f',key=f'cu_rent_{k}')
                utility=st.number_input('Utility Income',value=float(f.get('utility_income') or 0),format='%.2f',key=f'cu_utility_{k}')
            with b:
                income=st.number_input('Total Income',value=float(f['income']),format='%.2f',key=f'cu_income_{k}')
                expenses=st.number_input('Operating Expenses',value=float(f['expenses']),format='%.2f',key=f'cu_exp_{k}')
            with c:
                noi=st.number_input('NOI',value=float(f['noi']),format='%.2f',key=f'cu_noi_{k}')
                st.metric('Occupancy',f'{o["occupancy"]:.1%}',f'{o["occupied"]}/{o["rooms"]} rooms')
            other=st.number_input('Other Income',value=float(f.get('other_income') if f.get('other_income') is not None else f['income']-(f.get('rent_income') or 0)-(f.get('utility_income') or 0)),format='%.2f',key=f'cu_other_{k}')
            if abs(income-expenses-noi)>.05:st.error('Total Income − Expenses ≠ NOI')
            if abs(rent+utility+other-income)>.05:st.warning('Rent + Utility + Other ≠ Total Income; please review the breakdown.')
            for warning in o.get('warnings',[]):st.warning(warning)
            item=dict(item,row={'Rent Income':rent,'Utility Income':utility,'Other Income':other,
                                'Income':income,'Expenses':expenses,'NOI':noi})
            edited.append(item)
    st.subheader('2 · Property Operating Results / 各物业经营数据')
    st.caption('各物业独立展示，不计算 LLC 合计；每个章节都会分别列出各物业。')
    st.dataframe(pd.DataFrame([{'Property':e['name'],**e['row'],
       'Occupancy':f'{e["occ"]["occupancy"]:.1%}'} for e in edited]),
       hide_index=True,use_container_width=True)
    st.subheader('3 · Monthly Cash & Reserve / 各物业独立计算')
    st.caption('每个 Property 使用自己的 Beginning Cash、Tax/Insurance Reserve 和 Distribution。与单物业计算公式一致，不混合两个物业。')
    for item in edited:
        key=item['id']; settings=next(m['settings'] for m in members if m['id']==key)
        prior=history_for(key)
        item['previous']=prior or {}
        with st.expander(f'{item["name"]} · Cash & Reserve',expanded=True):
            a,b,c=st.columns(3)
            with a:
                beginning=st.number_input('Beginning Cash',value=float(number(prior.get('Ending Cash') if prior else settings.get('Opening Cash'))),format='%.2f',key=f'cu_begin_{key}_{month}')
                reserve_open=st.number_input('Beginning Required Reserve',value=float(number(prior.get('Required Reserve') if prior else settings.get('Opening Required Reserve'))),format='%.2f',key=f'cu_reserve_open_{key}_{month}')
                ar=st.number_input('Outstanding Rent Receivable',min_value=0.,format='%.2f',key=f'cu_ar_{key}')
                if prior:
                    st.success(f'已读取 {previous_month} 历史数据：Ending Cash {_money(number(prior.get("Ending Cash")))} · Reserve {_money(number(prior.get("Required Reserve")))}')
                else:
                    st.warning(f'找不到 {previous_month} 的历史记录，暂以 Properties 的 Opening Cash / Reserve 作为默认值；请核对。')
            with b:
                tax_month=st.number_input('Monthly Tax Reserve',value=float(number(settings.get('Monthly Tax Reserve'))),min_value=0.,format='%.2f',key=f'cu_tax_month_{key}')
                ins_month=st.number_input('Monthly Insurance Reserve',value=float(number(settings.get('Monthly Insurance Reserve'))),min_value=0.,format='%.2f',key=f'cu_ins_month_{key}')
                tax_paid=st.number_input('Property Tax Paid from Reserve',min_value=0.,format='%.2f',key=f'cu_tax_paid_{key}')
                ins_paid=st.number_input('Insurance Paid from Reserve',min_value=0.,format='%.2f',key=f'cu_ins_paid_{key}')
            with c:
                other_cash=st.number_input('Other Cash Expenses NOT in NOI',min_value=0.,format='%.2f',key=f'cu_other_cash_{key}')
                adjustment=st.number_input('Other Cash Adjustments (+/-)',format='%.2f',key=f'cu_adjust_{key}')
                paid=st.number_input('Distribution Paid',min_value=0.,format='%.2f',key=f'cu_paid_{key}')
            ending=beginning+item['row']['NOI']-tax_paid-ins_paid-other_cash-paid+adjustment
            reserve=reserve_open+tax_month+ins_month-tax_paid-ins_paid
            item['row'].update({'Beginning Cash':beginning,'Ending Cash':ending,'Required Reserve':reserve,
                'Reserve Shortfall':max(0.,reserve-ending),'Available to Distribute':max(0.,ending-reserve),
                'Distribution Paid':paid,'Monthly Reserve Addition':tax_month+ins_month,
                'Property Tax Paid':tax_paid,'Insurance Paid':ins_paid,'Other Cash Expense':other_cash,
                'Other Cash Adjustments':adjustment,'Outstanding Rent Receivable':ar,
                'Occupancy':item['occ']['occupancy'],'Occupied Rooms':item['occ']['occupied'],
                'Total Rooms':item['occ']['rooms'],
                'Distribution Frequency':str(settings.get('Distribution Frequency') or 'Monthly')})
            x,y,z=st.columns(3)
            x.metric('Ending Cash',_money(ending));y.metric('Required Reserve',_money(reserve))
            z.metric('Reserve Shortfall' if reserve>ending else 'Available Distribution',_money(abs(ending-reserve)))
            if reserve>ending:st.warning(f'{item["name"]}: Reserve Shortfall; review distributions.')
            st.caption('Ending Cash = Beginning Cash + Property NOI − Tax − Insurance − Other Cash Expenses − Distribution + Adjustments')
            st.caption('Reserve = Beginning Reserve + Monthly Tax Reserve + Monthly Insurance Reserve − Tax Paid − Insurance Paid')
    st.caption('Beginning Cash / Reserve 优先使用 Reports → Monthly Statements 上个月的滚动计算结果；没有历史数据时才回退到 Properties Opening 值。')
    st.subheader('4 · Executive Summary / 各物业独立说明')
    for item in edited:
        key=item['id']
        with st.expander(item['name']+' · Executive Summary',expanded=True):
            item['notes_cn']=st.text_area('Additional Operating Notes · 中文',key=f'cu_notes_cn_{key}')
            item['notes_en']=st.text_area('Additional Operating Notes · English',key=f'cu_notes_en_{key}')
            item['row']['Summary CN']=item['notes_cn']
            item['row']['Summary EN']=item['notes_en']
            item['row']['Notes']=item['notes_en']
            st.write(_summary_line(item,False))
    op_cn='\n'.join(_summary_line(item,True) for item in edited)
    op_en='\n'.join(_summary_line(item,False) for item in edited)
    cash_cn='\n'.join(f'{e["name"]}：期末现金 {_money(e["row"]["Ending Cash"])}，要求预留 {_money(e["row"]["Required Reserve"])}，可分配 {_money(e["row"]["Available to Distribute"])}。' for e in edited)
    cash_en='\n'.join(f'{e["name"]}: ending cash {_money(e["row"]["Ending Cash"])}; required reserve {_money(e["row"]["Required Reserve"])}; available {_money(e["row"]["Available to Distribute"])}.' for e in edited)
    notes_cn='\n'.join(f'{e["name"]}: {e["notes_cn"]}' for e in edited if e['notes_cn'])
    notes_en='\n'.join(f'{e["name"]}: {e["notes_en"]}' for e in edited if e['notes_en'])
    st.subheader('5 · Investor Report 下载')
    ready=all(abs(e['row']['Income']-e['row']['Expenses']-e['row']['NOI'])<=.05 for e in edited)
    if not ready:st.error('财务数据未配平，请先修正各物业 NOI。')
    clicked=st.button('Generate Combined Report',type='primary',disabled=not ready,key='cu_generate')
    import hashlib
    source_signature=tuple((e['id'],hashlib.sha256(e['income_pdf']).hexdigest(),
        hashlib.sha256(e['rent_pdf']).hexdigest(),str(e['row'])) for e in edited)
    signature=(llc,month,notes_cn,notes_en,source_signature)
    if not clicked and st.session_state.get('cu_artifacts_signature')!=signature:
        return
    if st.session_state.get('cu_artifacts_signature') != signature or clicked:
        artifacts=[]
        for cn,language in [(True,'CN'),(False,'EN')]:
            try:
                report_entries=[]
                for e in edited:
                    settings=next(m['settings'] for m in members if m['id']==e['id'])
                    report_entries.append({'name':e['name'],'row':e['row'],
                        'property':settings,'previous':e.get('previous') or {}})
                doc=make_chapter_combined_docx(llc,month,report_entries,cn)
                stem=f'Moo_{re.sub(r"[^A-Za-z0-9]+","_",llc)}_{month}_{language}_Combined'
                artifacts.append((stem+'.docx',doc,'application/vnd.openxmlformats-officedocument.wordprocessingml.document'))
                try:
                    selected_attachments=[]
                    for e in edited:
                        if cn:
                            try:
                                from chinese_appendix import translate_income,translate_rent_roll
                                income_payload=translate_income(e['income_pdf'],e['name'],month)
                                roll_payload=translate_rent_roll(e['rent_pdf'],e['name'],month)
                            except Exception as exc:
                                st.warning(f'{e["name"]} Chinese appendix translation unavailable: {exc}')
                                income_payload=e['income_pdf'];roll_payload=e['rent_pdf']
                        else:
                            income_payload=e['income_pdf']
                            try:
                                from english_appendix import corrected_rent_roll
                                roll_payload=corrected_rent_roll(e['rent_pdf'],e['name'],month)
                            except Exception as exc:
                                st.warning(f'{e["name"]} corrected English rent roll unavailable: {exc}')
                                roll_payload=e['rent_pdf']
                        selected_attachments.extend([(f'{e["name"]} Income Statement',income_payload),
                                                     (f'{e["name"]} Rent Roll',roll_payload)])
                    import shutil,subprocess,tempfile
                    from pathlib import Path
                    with tempfile.TemporaryDirectory() as tmp:
                        src=Path(tmp)/'combined.docx';src.write_bytes(doc)
                        exe=shutil.which('libreoffice') or shutil.which('soffice')
                        if not exe:raise RuntimeError('LibreOffice is not installed')
                        result=subprocess.run([exe,'-env:UserInstallation=file://'+tmp+'/lo','--headless',
                            '--convert-to','pdf','--outdir',tmp,str(src)],capture_output=True,text=True,timeout=90)
                        if result.returncode or not (Path(tmp)/'combined.pdf').exists():raise RuntimeError(result.stderr)
                        pdf=append_pdf_documents((Path(tmp)/'combined.pdf').read_bytes(),selected_attachments)
                    artifacts.append((stem+'.pdf',pdf,'application/pdf'))
                except Exception as exc:st.warning(f'{language} PDF unavailable: {exc}')
            except Exception as exc:st.error(f'{language} report generation failed: {exc}')
        st.session_state['cu_artifacts']=artifacts
        st.session_state['cu_artifacts_signature']=signature
    artifacts=st.session_state.get('cu_artifacts',[])
    for name,blob,mime in artifacts:
        st.download_button('Download '+name,blob,file_name=name,mime=mime,key='cu_dl_'+name)
    if artifacts and st.button('☁️ Save Combined Reports to Google Drive',key='cu_save_drive'):
        try:
            saved=save_reports(st.secrets,llc,month[:4],artifacts)
            st.success(f'Saved {len(saved["files"])} files to Google Drive')
            st.link_button('Open LLC Reports Folder',saved['folder_url'])
        except Exception as exc:st.error(f'Drive save failed: {exc}')
    show_communications(llc,month,op_cn,cash_cn,notes_cn,op_en,cash_en,notes_en,
                        key='cu_'+re.sub(r'[^a-zA-Z0-9]','_',llc))
