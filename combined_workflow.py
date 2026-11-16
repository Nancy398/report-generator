"""Direct-upload Combined LLC workflow; no Monthly Statements pre-save required."""
import io, re
from datetime import datetime
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Pt
from pdf_appfolio import parse_income_pdf, parse_rent_roll_pdf
from reporting import make_docx, make_pdf, append_pdf_documents
from engine import number, pid, period
from investor_communication import show_communications
from combined_llc import generated_summary


def _n(v):
    if v is None or str(v).strip() == '': return None
    try: return float(str(v).replace('$','').replace(',',''))
    except (ValueError,TypeError): return None


def _money(v): return 'Not provided' if v is None else f'${v:,.2f}'


def _build_model(llc, month, items, cash, notes_cn, notes_en, frequency):
    props=[]
    for item in items:
        row=item['row']; occ=item['occ']
        props.append(dict(id=item['id'],name=item['name'],rent=row['Rent Income'],
            utility=row['Utility Income'],other=row['Other Income'],income=row['Income'],
            expenses=row['Expenses'],noi=row['NOI'],occupancy=occ['occupancy']*100,
            prior_occupancy=None,rooms=occ['rooms'],occupied=occ['occupied']))
    return dict(group=llc,month=month,properties=props,cash=cash.get('Ending Cash'),
        reserve=cash.get('Required Reserve'),shortfall=cash.get('Reserve Shortfall'),
        available=cash.get('Available to Distribute'),distribution=cash.get('Distribution Paid'),
        frequency=frequency,summary_cn=notes_cn,summary_en=notes_en,
        total_noi=sum(p['noi'] for p in props))


def _docx(model, row, cn, appendix):
    """Reuse the exact single-property Word master and append per-property breakdown."""
    prop={'Property Name / LLC':model['group'],
          'Property Address':f'{len(model["properties"])} properties · Combined LLC Report',
          'Distribution Frequency':model['frequency']}
    source=make_docx(row,prop,cn,appendix_documents=appendix)
    doc=Document(io.BytesIO(source))
    doc.add_page_break()
    doc.add_heading('各物业经营明细' if cn else 'PROPERTY OPERATING DETAILS',level=1)
    for item in model['properties']:
        doc.add_heading(item['name'],level=2)
        pairs=[('Occupancy',f'{item["occupancy"]:.1f}%'),
               ('Occupied / Total Rooms',f'{item["occupied"]} / {item["rooms"]}'),
               ('Rent Income',_money(item['rent'])),('Utility Income',_money(item['utility'])),
               ('Other Income',_money(item['other'])),('Total Income',_money(item['income'])),
               ('Operating Expenses',_money(item['expenses'])),('NOI',_money(item['noi']))]
        t=doc.add_table(rows=0,cols=2)
        t.style='Table Grid'
        for label,value in pairs:
            cells=t.add_row().cells;cells[0].text=label;cells[1].text=value
        doc.add_paragraph()
    doc.add_heading('LLC TOTAL' if not cn else 'LLC 合计',level=2)
    t=doc.add_table(rows=0,cols=2);t.style='Table Grid'
    for label,value in [('Rent Income',sum(x['rent'] for x in model['properties'])),
                        ('Total Income',sum(x['income'] for x in model['properties'])),
                        ('Operating Expenses',sum(x['expenses'] for x in model['properties'])),
                        ('NOI',model['total_noi'])]:
        cells=t.add_row().cells;cells[0].text=label;cells[1].text=_money(value)
    out=io.BytesIO();doc.save(out);return out.getvalue()


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
    totals={field:sum(e['row'][field] for e in edited) for field in ('Rent Income','Utility Income','Other Income','Income','Expenses','NOI')}
    total_rooms=sum(e['occ']['rooms'] for e in edited);occupied=sum(e['occ']['occupied'] for e in edited)
    st.subheader('2 · Combined Operating Results')
    a,b,c,d=st.columns(4)
    a.metric('Combined Income',_money(totals['Income']))
    b.metric('Combined Expenses',_money(totals['Expenses']))
    c.metric('Combined NOI',_money(totals['NOI']))
    d.metric('Combined Occupancy',f'{occupied/total_rooms:.1%}',f'{occupied}/{total_rooms} rooms')
    st.dataframe(pd.DataFrame([{'Property':e['name'],**e['row'],
       'Occupancy':f'{e["occ"]["occupancy"]:.1%}'} for e in edited]),hide_index=True,use_container_width=True)
    st.subheader('3 · Monthly Cash & Reserve / LLC 统一计算')
    st.caption('与 Single Property 一样采用现金滚动及固定 Tax / Insurance Reserve。只在 LLC 层面计算一次，避免重复计算共用银行账户。')
    include_cash=st.checkbox('Calculate LLC Cash / Reserve / Distribution',value=True,key='cu_include_cash')
    cash={}
    freq_values={str(x['settings'].get('Distribution Frequency') or 'Monthly').strip().title() for x in members}
    frequency=next(iter(freq_values)) if len(freq_values)==1 else 'Monthly'
    if len(freq_values)>1:st.warning('LLC 下物业的 Distribution Frequency 不一致，请核对。')
    if include_cash:
        a,b,c=st.columns(3)
        with a:
            beginning=st.number_input('LLC Beginning Cash',format='%.2f',key='cu_beginning')
            opening_reserve=st.number_input('LLC Beginning Required Reserve',min_value=0.,format='%.2f',key='cu_open_reserve')
            ar=st.number_input('Outstanding Rent Receivable',min_value=0.,format='%.2f',key='cu_ar')
        with b:
            monthly_tax=sum(number(x['settings'].get('Monthly Tax Reserve')) for x in members)
            monthly_ins=sum(number(x['settings'].get('Monthly Insurance Reserve')) for x in members)
            tax_res=st.number_input('Monthly Tax Reserve',value=monthly_tax,min_value=0.,format='%.2f',key='cu_monthly_tax')
            ins_res=st.number_input('Monthly Insurance Reserve',value=monthly_ins,min_value=0.,format='%.2f',key='cu_monthly_ins')
            tax_paid=st.number_input('Property Tax Paid from Reserve',min_value=0.,format='%.2f',key='cu_tax_paid')
            ins_paid=st.number_input('Insurance Paid from Reserve',min_value=0.,format='%.2f',key='cu_ins_paid')
        with c:
            other_cash=st.number_input('Other Cash Expenses NOT in NOI',min_value=0.,format='%.2f',key='cu_other_cash')
            adjustments=st.number_input('Other Cash Adjustments (+/-)',format='%.2f',key='cu_adjustments')
            distribution=st.number_input('Distribution Paid',min_value=0.,format='%.2f',key='cu_distribution')
        ending=beginning+totals['NOI']-tax_paid-ins_paid-other_cash-distribution+adjustments
        reserve=opening_reserve+tax_res+ins_res-tax_paid-ins_paid
        cash={'Beginning Cash':beginning,'Ending Cash':ending,'Required Reserve':reserve,
              'Reserve Shortfall':max(0,reserve-ending),'Available to Distribute':max(0,ending-reserve),
              'Distribution Paid':distribution,'Monthly Reserve Addition':tax_res+ins_res,
              'Property Tax Paid':tax_paid,'Insurance Paid':ins_paid,
              'Other Cash Expense':other_cash,'Other Cash Adjustments':adjustments,
              'Outstanding Rent Receivable':ar}
        x,y,z=st.columns(3)
        x.metric('Modeled Ending Cash',_money(ending));y.metric('Required Reserve',_money(reserve))
        z.metric('Reserve Shortfall' if reserve>ending else 'Available Distribution',_money(abs(ending-reserve)))
        st.caption('Ending Cash = Beginning Cash + Combined NOI − Tax Paid − Insurance Paid − Other Cash Expenses − Distribution + Adjustments')
        st.caption('Required Reserve = Beginning Required Reserve + Monthly Tax Reserve + Monthly Insurance Reserve − Tax Paid − Insurance Paid')
        if reserve>ending:st.warning('Reserve Shortfall：当前没有可供分红审核的现金。')
        st.warning('请核对 LLC Beginning Cash 和 Reserve：本次直接上传模式不会自动读取历史现金滚动记录。')
    st.subheader('4 · Executive Summary / 报告说明')
    notes_cn=st.text_area('Additional Operating Notes · 中文',key='cu_notes_cn')
    notes_en=st.text_area('Additional Operating Notes · English',key='cu_notes_en')
    model=_build_model(llc,month,edited,cash,notes_cn,notes_en,frequency)
    op_cn,cash_cn=generated_summary(model,True) if cash else ('\n'.join(f'{p["name"]}：NOI {_money(p["noi"])}，入住率 {p["occupancy"]:.1f}%。' for p in model['properties']),'现金及分红数据未提供。')
    op_en,cash_en=generated_summary(model,False) if cash else ('\n'.join(f'{p["name"]}: NOI {_money(p["noi"])}; occupancy {p["occupancy"]:.1f}%.' for p in model['properties']),'Cash and distribution information was not supplied.')
    with st.expander('Report Preview / 预览',expanded=True):
        st.write(op_en);st.write(cash_en)
        if notes_en:st.write(notes_en)
    st.subheader('5 · Investor Report 下载')
    ready=all(abs(e['row']['Income']-e['row']['Expenses']-e['row']['NOI'])<=.05 for e in edited)
    if not ready:st.error('财务数据未配平，请先修正各物业 NOI。')
    clicked=st.button('Generate Combined Report',type='primary',disabled=not ready,key='cu_generate')
    import hashlib
    source_signature=tuple((e['id'],hashlib.sha256(e['income_pdf']).hexdigest(),hashlib.sha256(e['rent_pdf']).hexdigest()) for e in edited)
    signature=(llc,month,str(totals),str(cash),notes_cn,notes_en,source_signature)
    if not clicked and st.session_state.get('cu_artifacts_signature')!=signature:
        return
    report_row={'Month (YYYY-MM)':month,**totals,'Occupancy':occupied/total_rooms,
                'Occupied Rooms':occupied,'Total Rooms':total_rooms,
                'Summary CN':notes_cn,'Summary EN':notes_en,'Distribution Frequency':frequency,**cash}
    if st.session_state.get('cu_artifacts_signature') != signature or clicked:
        appendix=[(f'{e["name"]} — {kind}',month) for e in edited for kind in ('Income Statement','Rent Roll')]
        attachments=[]
        for e in edited:
            attachments.extend([(f'{e["name"]} Income Statement',e['income_pdf']),
                                (f'{e["name"]} Rent Roll',e['rent_pdf'])])
        artifacts=[]
        for cn,language in [(True,'CN'),(False,'EN')]:
            try:
                doc=_docx(model,report_row,cn,appendix)
                stem=f'Moo_{re.sub(r"[^A-Za-z0-9]+","_",llc)}_{month}_{language}_Combined'
                artifacts.append((stem+'.docx',doc,'application/vnd.openxmlformats-officedocument.wordprocessingml.document'))
                try:
                    # Match Single Property: Chinese translated appendices and corrected
                    # English rent rolls, preserving the original PDFs as fallback.
                    selected_attachments=[]
                    for e in edited:
                        if cn:
                            try:
                                from chinese_appendix import translate_income,translate_rent_roll
                                income_payload=translate_income(e['income_pdf'],e['name'],month)
                                roll_payload=translate_rent_roll(e['rent_pdf'],e['name'],month)
                            except Exception as exc:
                                st.warning(f'{e["name"]} Chinese appendix conversion failed; original PDF attached: {exc}')
                                income_payload=e['income_pdf'];roll_payload=e['rent_pdf']
                        else:
                            income_payload=e['income_pdf']
                            try:
                                from english_appendix import corrected_rent_roll
                                roll_payload=corrected_rent_roll(e['rent_pdf'],e['name'],month)
                            except Exception as exc:
                                st.warning(f'{e["name"]} corrected English rent roll unavailable; original PDF attached: {exc}')
                                roll_payload=e['rent_pdf']
                        selected_attachments.extend([(f'{e["name"]} Income Statement',income_payload),
                                                     (f'{e["name"]} Rent Roll',roll_payload)])
                    # Reuse single-property PDF rendering and append source appendices.
                    import shutil,subprocess,tempfile
                    from pathlib import Path
                    with tempfile.TemporaryDirectory() as tmp:
                        src=Path(tmp)/'combined.docx';src.write_bytes(doc)
                        exe=shutil.which('libreoffice') or shutil.which('soffice')
                        if not exe:raise RuntimeError('LibreOffice is not installed')
                        result=subprocess.run([exe,'-env:UserInstallation=file://'+tmp+'/lo','--headless','--convert-to','pdf','--outdir',tmp,str(src)],capture_output=True,text=True,timeout=90)
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
                        key='cu_'+re.sub(r'[^a-zA-Z0-9]','_',llc),report_row=report_row)
