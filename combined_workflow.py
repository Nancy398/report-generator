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
    # Use the exact Single Property Report engine and monthly source records.
    history_rows=[];history_calculated=[];history_issues=[];all_property_settings=[]
    try:
        master=google_load()
        history_rows=records(master.get('Monthly Statements'))
        all_property_settings=records(master.get('Properties'))
        history_calculated,history_issues=calculate(all_property_settings,history_rows)
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
    save_slot=st.empty()  # Fixed location: visible before PDF uploads / report downloads.
    with save_slot.container():
        st.subheader('Save to Google Sheets · Reports → Monthly Statements')
        st.info('请上传全部物业 PDF 并完成 Cash & Reserve 计算，随后这里会出现保存按钮。')
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
                ar=st.number_input('Outstanding Rent Receivable',min_value=0.,
                    value=float(number(item['previous'].get('Outstanding Rent Receivable'))),
                    format='%.2f',key=f'cu_ar_{key}_{month}')
                if prior:
                    st.info(f'{previous_month} previous month: Ending Cash {_money(number(prior.get("Ending Cash")))}; Required Reserve {_money(number(prior.get("Required Reserve")))}')
                else:
                    st.warning(f'没有 {previous_month} 历史记录；计算将使用 Single Property 相同的 Opening Balance / 现有历史记录。')
            with b:
                tax_paid=st.number_input('Property Tax Paid from Reserve',min_value=0.,format='%.2f',key=f'cu_tax_paid_{key}_{month}')
                ins_paid=st.number_input('Insurance Paid from Reserve',min_value=0.,format='%.2f',key=f'cu_ins_paid_{key}_{month}')
            with c:
                other_cash=st.number_input('Other Cash Expenses NOT in NOI',min_value=0.,format='%.2f',key=f'cu_other_cash_{key}_{month}')
                adjustment=st.number_input('Other Cash Adjustments (+/-)',format='%.2f',key=f'cu_adjust_{key}_{month}')
                paid=st.number_input('Distribution Paid (actual)',min_value=0.,format='%.2f',key=f'cu_paid_{key}_{month}')
            # Match the Single Property form: submit only INPUT fields to engine.calculate().
            # No manual beginning balance or duplicate combined-only cash formula.
            input_row={
                'Property ID':key,'Month (YYYY-MM)':month,
                'Rent Income':item['row']['Rent Income'],
                'Utility Income':item['row']['Utility Income'],
                'Other Income':item['row']['Other Income'],
                'Income':item['row']['Income'],'Expenses':item['row']['Expenses'],
                'NOI':item['row']['NOI'],
                'Property Tax Paid':tax_paid,'Insurance Paid':ins_paid,
                'Other Cash Expense':other_cash,'Other Cash Adjustments':adjustment,
                'Distribution Paid':paid,'Outstanding Rent Receivable':ar,
                'Occupancy':item['occ']['occupancy'],
                'Occupied Rooms':item['occ']['occupied'],
                'Total Rooms':item['occ']['rooms'],
                'Calculation Version':'modeled-v2','Status':'Draft',
            }
            # Replace an existing draft of this month (same behavior as Single Property).
            current_history=[h for h in history_rows if not (
                pid(h.get('Property ID'))==key and period(h.get('Month (YYYY-MM)'))==month)]
            current_history.append(input_row)
            calculated,calc_issues=calculate(all_property_settings or [m['settings'] for m in members],current_history)
            current=next((r for r in calculated if pid(r.get('Property ID'))==key
                and period(r.get('Month (YYYY-MM)'))==month),None)
            if current is None:
                st.error(f'{item["name"]}: Single Property calculation failed. Check Properties and Monthly Statements.')
                item['calculation_error']=True
                continue
            for issue in calc_issues:
                if key in issue: st.warning(issue)
            item['save_input']=input_row
            item['row'].update(current)
            item['row'].update({'Outstanding Rent Receivable':ar,
                'Occupancy':item['occ']['occupancy'],
                'Occupied Rooms':item['occ']['occupied'],
                'Total Rooms':item['occ']['rooms'],
                'Distribution Frequency':str(settings.get('Distribution Frequency') or 'Monthly')})
            x,y,z=st.columns(3)
            x.metric('Beginning Cash',_money(current['Beginning Cash']))
            y.metric('Ending Cash',_money(current['Ending Cash']))
            z.metric('Required Reserve',_money(current['Required Reserve']))
            st.caption(f'Available to Distribute: {_money(current["Available to Distribute"])} · Reserve Shortfall: {_money(current["Reserve Shortfall"])}')
            if current['Reserve Shortfall']>0:st.warning('Reserve Shortfall — review distribution.')
            st.caption('Calculated with the same engine.calculate() and monthly history as Single Property Report.')
    st.caption('Beginning Cash 和 Reserve 由 Single Property 的 engine.calculate() 从历史记录自动滚动；不在 Combined 中重复计算。')
    if any(e.get('calculation_error') for e in edited):
        with save_slot.container():
            st.subheader('Save to Google Sheets · 暂不可保存')
            st.error('物业计算未完成；请先修正上方错误。')
        return
    llc_available=sum(e['row']['Available to Distribute'] for e in edited)
    llc_paid=sum(e['row']['Distribution Paid'] for e in edited)
    st.subheader('4 · LLC Distribution Summary / LLC 分红汇总')
    st.dataframe(pd.DataFrame([{'Property':e['name'],'Available to Distribute':e['row']['Available to Distribute'],'Distribution Paid':e['row']['Distribution Paid']} for e in edited] + [{'Property':'LLC TOTAL','Available to Distribute':llc_available,'Distribution Paid':llc_paid}]),hide_index=True,use_container_width=True)
    st.caption('先分别计算各物业 Available to Distribute，再加总 LLC Total；实际已分红金额单独合计，不等同于 Available。')
    st.subheader('5 · Executive Summary / 各物业独立说明')
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
    # Persist EACH property separately in the same Monthly Statements sheet used
    # by Single Property Report. No LLC aggregate row is written.
    with save_slot.container():
        st.subheader('💾 Save to Google Sheets · Reports → Monthly Statements')
        st.caption('保存当前月份的每个 Property（独立行），和 Single Property Report 使用相同的 Google Sheets 保存函数。')
        valid_for_save=all(abs(e['row']['Income']-e['row']['Expenses']-e['row']['NOI'])<=.05 for e in edited)
        if not valid_for_save:
            st.error('Income − Expenses 与 NOI 不一致，请先核对，保存按钮暂不可用。')
        if st.button('💾 Save Both Properties to Google Sheets',
                     disabled=not valid_for_save,key='cu_save_monthly',type='primary'):
            saved_names=[];failures=[]
            for e in edited:
                r=e['row']
                rec=dict(e['save_input'])
                for field in ('Beginning Cash','Modeled Ending Cash','Ending Cash',
                              'Tax Reserve Required','Insurance Reserve Required',
                              'Total Required Reserve','Required Reserve',
                              'Reserve Shortfall','Available to Distribute',
                              'Monthly Reserve Addition'):
                    if field in r:
                        rec[field]=r[field]
                rec['Notes']=e.get('notes_en','')
                rec['Summary CN']=e.get('notes_cn','')
                rec['Summary EN']=e.get('notes_en','')
                try:
                    store_google(rec)
                    saved_names.append(e['name'])
                except Exception as exc:
                    failures.append(f'{e["name"]}: {exc}')
            if saved_names:
                st.success('已写入 Reports → Monthly Statements：'+', '.join(saved_names))
                directory.clear()
            for failure in failures:
                st.error('保存失败：'+failure)
            if failures:
                st.warning('仅部分物业保存成功；请处理错误后重试。')
        st.caption('Google Drive 文件保存是另一个按钮，不会写入 Monthly Statements。')
    st.subheader('7 · Investor Report 下载')
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
