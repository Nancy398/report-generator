"""Investor reports built from the actual 1438 Word master, not a visual recreation."""
import io
import os
import shutil
import subprocess
import tempfile
from copy import deepcopy
from pathlib import Path
from datetime import datetime
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn

MASTER = Path(__file__).with_name('1438_template.docx')


def _get(row, *names, default='Not provided'):
    for name in names:
        val = row.get(name)
        if val is not None and str(val).strip() != '':
            return val
    return default


def _cash(v):
    try: return '${:,.2f}'.format(float(str(v).replace(',','').replace('$','')))
    except (TypeError, ValueError): return 'Not provided'


def _pct(v):
    try:
        n=float(v)
        return f'{n*100 if n<=1 else n:.1f}%'
    except (TypeError, ValueError): return 'Not provided'


def _replace_para(p, new_text):
    # Retain paragraph formatting and first run's character formatting.
    if p.runs:
        p.runs[0].text = str(new_text)
        for run in p.runs[1:]: run.text = ''
    else: p.add_run(str(new_text))


def _clone_paragraph(doc, source_element, text=None):
    el=deepcopy(source_element)
    doc.element.body.insert(len(doc.element.body)-1, el)
    p=Paragraph(el,doc)
    if text is not None: _replace_para(p,text)
    return p


def _clone_table(doc, source_element, headings, data, total_last=False):
    el=deepcopy(source_element)
    doc.element.body.insert(len(doc.element.body)-1,el)
    table=Table(el,doc)
    while len(table.rows)>1:
        table._tbl.remove(table.rows[-1]._tr)
    template_row=deepcopy(source_element.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr')[1])
    # Extend the original 1438 two-column table to four columns when needed.
    if len(headings) > len(table._tbl.tblGrid.gridCol_lst):
        extra=len(headings)-len(table._tbl.tblGrid.gridCol_lst)
        grid=table._tbl.tblGrid
        for _ in range(extra):
            grid.append(deepcopy(grid[-1]))
        for tr in (table.rows[0]._tr, template_row):
            cells=tr.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc')
            for _ in range(extra):tr.append(deepcopy(cells[-1]))
        # Original table is two columns; split width evenly across four.
        for gridcol in grid:
            width=gridcol.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w')
            if width: gridcol.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w',str(max(1,int(width)//2)))
        for tr in (table.rows[0]._tr,template_row):
            for tc in tr.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'):
                tcw=tc.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcW')
                if tcw is not None and tcw.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w'):
                    tcw.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w','2300')
    for item in data:
        table._tbl.append(deepcopy(template_row))
    for ri, vals in enumerate([headings]+data):
        for ci, val in enumerate(vals):
            if ci>=len(table.rows[ri].cells): continue
            cell=table.rows[ri].cells[ci]
            _replace_para(cell.paragraphs[0],val)
            for extra in cell.paragraphs[1:]: _replace_para(extra,'')
    if total_last and data:
        # Copy the master table's original total-row shading and bold emphasis.
        original_rows=source_element.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr')
        if len(original_rows)>2:
            last=deepcopy(original_rows[-1])
            table._tbl.replace(table.rows[-1]._tr,last)
            for ci,val in enumerate(data[-1]):
                _replace_para(table.rows[-1].cells[ci].paragraphs[0],val)
    if len(headings)==4:
        grid=table._tbl.tblGrid
        while len(grid.gridCol_lst)<4:
            grid.append(deepcopy(grid.gridCol_lst[-1]))
        for gc in grid.gridCol_lst:
            gc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w','2300')
    return table


def _number(value):
    try:return float(str(value).replace(',','').replace('$','').replace('%','').strip())
    except (ValueError,TypeError):return None

def _change(current,prior,kind):
    a,b=_number(current),_number(prior)
    if a is None or b is None:return 'N/A'
    if kind=='percent':
        a=a*100 if a<=1 else a
        b=b*100 if b<=1 else b
        return f'{a-b:+.1f} pp'
    if kind=='money':return f'{a-b:+,.2f}' if a-b<0 else f'+${a-b:,.2f}' if a-b>0 else '$0.00'
    return f'{a-b:+,.0f}'

def _set_font(run, western='Cambria', east_asia='Microsoft YaHei', size=None, bold=None, color=None):
    run.font.name = western
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.rFonts
    if fonts is None:
        fonts = OxmlElement('w:rFonts')
        rpr.insert(0, fonts)
    fonts.set(qn('w:eastAsia'), east_asia)
    if size is not None: run.font.size = Pt(size)
    if bold is not None: run.bold = bold
    if color is not None: run.font.color.rgb = RGBColor(*color)


def _format_banner(paragraph, company, address, cn):
    """Compact, left-aligned two-line running banner; keep template's right logo."""
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    logo_runs = [run for run in paragraph.runs if run._element.xpath('.//w:drawing')]
    for run in logo_runs:
        paragraph._p.remove(run._element)
    for run in paragraph.runs:
        if run.text:
            run.text = ''
    title = '投资人财务报告' if cn else 'INVESTOR FINANCIAL REPORT'
    for index, (value, size, bold) in enumerate(((title, 9, True), (address, 8, True))):
        run = paragraph.add_run()
        if index:
            run.add_break()
        run.add_text(value)
        _set_font(run, western='Arial', east_asia='Microsoft YaHei', size=size,
                  bold=bold, color=(35, 69, 92))
    if logo_runs:
        paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(6.25), WD_TAB_ALIGNMENT.RIGHT)
        paragraph.add_run("\t")
        for logo in logo_runs:
            paragraph._p.append(logo._element)
    paragraph.paragraph_format.space_after = Pt(10)
    paragraph.paragraph_format.line_spacing = 1.0


def _polish_chinese(doc):
    # Explicit East Asian fonts are needed for reliable Word/LibreOffice PDF output.
    for paragraph in doc.paragraphs:
        if not paragraph.text.strip(): continue
        for run in paragraph.runs:
            if not run.text: continue
            if any('\u3400' <= c <= '\u9fff' for c in run.text):
                _set_font(run, western='Cambria', east_asia='Microsoft YaHei')
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if any('\u3400' <= c <= '\u9fff' for c in run.text):
                            _set_font(run, western='Cambria', east_asia='Microsoft YaHei', size=10)


def _apply_sans_typography(doc):
    """Apply a consistent sans-serif family without disturbing template geometry/logo."""
    for paragraph in list(doc.paragraphs) + [p for sec in doc.sections for p in sec.header.paragraphs]:
        for run in paragraph.runs:
            if not run.text: continue
            _set_font(run, western='Arial', east_asia='Noto Sans CJK SC')
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if run.text:
                            _set_font(run, western='Arial', east_asia='Noto Sans CJK SC')
    for style in doc.styles:
        if style.type == 1:
            style.font.name = 'Arial'
            rpr = style.element.get_or_add_rPr()
            fonts = rpr.rFonts
            if fonts is None:
                fonts = OxmlElement('w:rFonts')
                rpr.insert(0, fonts)
            fonts.set(qn('w:eastAsia'), 'Noto Sans CJK SC')


def make_docx(row,p,cn=False,previous=None,appendix_documents=None):
    if not MASTER.exists(): raise FileNotFoundError('1438_template.docx must be beside reporting.py')
    doc=Document(MASTER)
    originals=[deepcopy(e) for e in doc.element.body]
    # Keep the original cover (including logo, fonts and spacing), replace content only.
    prop=str(_get(p,'Property Name / LLC','Property Name',default='PROPERTY NAME'))
    addr=str(_get(p,'Property Address',default='Los Angeles, California'))
    period=str(_get(row,'Month (YYYY-MM)',default=datetime.now().strftime('%Y-%m')))
    try: report_date=datetime.strptime(period[:7],'%Y-%m').strftime('%B %Y')
    except ValueError: report_date=period
    cover={3:'Project Overview, Operations & Distribution Summary',4:prop.upper(),5:addr,7:report_date}
    for i,txt in cover.items(): _replace_para(Paragraph(doc.element.body[i],doc),txt)
    for e in list(doc.element.body)[12:-1]: doc.element.body.remove(e)

    occupancy=_pct(_get(row,'Occupancy'))
    occupied=_get(row,'Occupied Rooms')
    total=_get(row,'Total Rooms')
    try: vacant=str(int(float(total))-int(float(occupied)))
    except (ValueError,TypeError): vacant='Not provided'
    summary=str(_get(row,'Summary CN' if cn else 'Summary EN',default='') or '').strip()
    if not summary:
        summary=str(_get(row,'Summary','Executive Summary','Notes',default='') or '').strip()
    frequency=str(_get(row,'Distribution Frequency',default=_get(p,'Distribution Frequency',default='Monthly'))).strip().lower()
    quarterly=frequency in ('quarterly','quarter','季度','季度分红')
    previous=previous or {}
    cur_occ=_number(_get(row,'Occupancy',default=''))
    prev_occ=_number(_get(previous,'Occupancy',default=''))
    def occ_percent(v): return v*100 if v is not None and v<=1 else v
    if cur_occ is None:
        occ_line=('本月入住率数据尚未提供。' if cn else 'Current-month occupancy data is unavailable.')
    elif prev_occ is None:
        occ_line=(f'本月入住率为 {occ_percent(cur_occ):.1f}%；暂无上月可比数据。' if cn else f'Occupancy was {occ_percent(cur_occ):.1f}% this month; prior-month data is unavailable.')
    else:
        delta=occ_percent(cur_occ)-occ_percent(prev_occ)
        if cn:
            direction='上升' if delta>0.005 else '下降' if delta< -0.005 else '持平'
            occ_line=f'本月入住率为 {occ_percent(cur_occ):.1f}%，上月为 {occ_percent(prev_occ):.1f}%，' + (f'环比{direction} {abs(delta):.1f} 个百分点。' if direction!='持平' else '环比持平。')
        else:
            direction='increased' if delta>0.005 else 'decreased' if delta< -0.005 else 'remained unchanged'
            occ_line=f'Occupancy was {occ_percent(cur_occ):.1f}% versus {occ_percent(prev_occ):.1f}% last month; ' + (f'it {direction} by {abs(delta):.1f} percentage points.' if direction!='remained unchanged' else 'it remained unchanged.')
    available_num=_number(_get(row,'Available to Distribute',default=''))
    distribution_num=_number(_get(row,'Distribution Paid',default=''))
    money=lambda x: '${:,.2f}'.format(x)
    if available_num is None:
        cash_line=('可分配现金余额尚未提供。' if cn else 'Available cash for distribution is unavailable.')
    elif quarterly:
        cash_line=(f'截至本月末，累计可分配现金为 {money(available_num)}。' if cn else f'Cumulative available cash for distribution at month-end was {money(available_num)}.')
        cash_line+=('该余额将结转至下一季度分红评估。' if cn else ' This balance will carry forward to the next quarterly distribution review.')
    else:
        cash_line=(f'本月期末可分配现金余额为 {money(available_num)}。' if cn else f'Available cash for distribution at month-end was {money(available_num)}.')

    noi_num=_number(_get(row,'NOI',default=''))
    if noi_num is None:
        noi_line=('本月 NOI 数据尚未提供。' if cn else 'NOI data for this month is unavailable.')
    else:
        noi_line=(f'本月净营业收入（NOI）为 {money(noi_num)}。' if cn else f'Net Operating Income (NOI) for the month was {money(noi_num)}.')
    extra_notes=summary

    # Income Statement rent is accrual/accounting revenue, not verified cash collected.
    rent=_cash(_get(row,'Rent Income'))
    ar=_cash(_get(row,'Outstanding Rent Receivable','Outstanding AR','Accounts Receivable','AR'))
    income=_cash(_get(row,'Income','Total Income'))
    rent_income=_cash(_get(row,'Rent Income'))
    utility_income=_cash(_get(row,'Utility Income'))
    other_income=_get(row,'Other Income',default=None)
    if rent_income == 'Not provided' or utility_income == 'Not provided':
        raise ValueError('Income Breakdown missing in report row. Re-upload the Income Statement and check Rent Income / Utility Income before generating.')
    income_breakdown=[('Rent Income',rent_income),('Utility Income',utility_income)]
    if other_income is not None and abs(float(other_income or 0))>0.005:
        income_breakdown.append(('Other Income',_cash(other_income)))
    income_breakdown.append(('Total Income',income))
    expenses=_cash(_get(row,'Expenses','Operating Expenses'))
    noi=_cash(_get(row,'NOI'))
    cash=_cash(_get(row,'Ending Cash'))
    reserve=_cash(_get(row,'Required Reserve'))
    shortfall=_cash(_get(row,'Reserve Shortfall'))
    available=_cash(_get(row,'Available to Distribute'))
    distribution=_cash(_get(row,'Distribution Paid'))
    previous=previous or {}
    prior_occupied=_get(previous,'Occupied Rooms',default='')
    prior_total=_get(previous,'Total Rooms',default='')
    try:prior_vacant=str(int(float(prior_total))-int(float(prior_occupied)))
    except (ValueError,TypeError):prior_vacant='N/A'
    def comp(label,key,kind,cur_display=None,prior_display=None):
        now=_get(row,key,default='')
        old=_get(previous,key,default='')
        formatter=_pct if kind=='percent' else _cash if kind=='money' else lambda x:str(x) if str(x).strip() else 'N/A'
        return (label,prior_display or (formatter(old) if str(old).strip() else 'N/A'),cur_display or (formatter(now) if str(now).strip() else 'N/A'),_change(now,old,kind))
    comparison=[comp('Occupancy Rate','Occupancy','percent'),
                comp('Occupied Rooms','Occupied Rooms','count'),
                ('Vacant Rooms',prior_vacant,vacant,_change(vacant,prior_vacant,'count')),
                comp('Total Rooms','Total Rooms','count'),
                comp('Rent Income','Rent Income','money'),
                comp('Outstanding Rent Receivable','Outstanding Rent Receivable','money')]
    appendix_documents = appendix_documents or []
    appendix_items = [(f'Appendix {chr(65+i)}', label.split(' — ')[0].strip() + (f' — {date[:7]}' if date else '')) for i,(label,date) in enumerate(appendix_documents)]
    appendix_intro = ('The following original AppFolio reports are attached after this report for reference. ' + 'They provide supporting details for the financial and occupancy information presented above.') if appendix_items else 'No supporting PDF documents were uploaded for this report.'
    sections=[
      ('01  EXECUTIVE SUMMARY','',[('AT A GLANCE',['KEY METRIC','MONTH-END VALUE'],[
          ('Occupancy Rate',occupancy),('Net Operating Income (NOI)',_cash(_get(row,'NOI'))),('Cumulative Available Cash' if quarterly else 'Available Cash',available)])]),
      ('02  OCCUPANCY & RENT COLLECTION','Occupancy and collection results for the reporting month.',[
          ('OCCUPANCY & COLLECTION',['METRIC','PREVIOUS MONTH','CURRENT MONTH','CHANGE'],comparison)]),
      ('03  FINANCIAL SUMMARY','Operating results shown below are based on the supplied financial records.',[
          ('OPERATING RESULTS',['ITEM','AMOUNT'],income_breakdown+[('Operating Expenses',expenses),('Net Operating Income (NOI)',noi)]),
          ('CASH & RESERVES',['ITEM','AMOUNT'],[('Ending Cash Balance',cash),('Required Reserve',reserve),('Reserve Shortfall',shortfall)])]),
      ('04  DISTRIBUTION DETAILS','',[
          ('DISTRIBUTION OVERVIEW',['CASH POSITION','USD'],[('Ending Cash Balance',cash),('Less: Required Reserve',reserve),('Available for Distribution Review',available)])]),
      ('05  APPENDIX',appendix_intro,([('SUPPORTING DOCUMENTS',['DOCUMENT','DESCRIPTION'],appendix_items)] if appendix_items else []) + [
          ('SOURCE & RECONCILIATION',['ITEM','DETAILS'],[('Property',prop),('Reporting Period',period),('Property Tax Paid',_cash(_get(row,'Property Tax Paid'))),('Insurance Paid',_cash(_get(row,'Insurance Paid'))),('Other Cash Adjustments',_cash(_get(row,'Other Cash Adjustments')))])])
    ]
    for idx,(heading,body,blocks) in enumerate(sections):
        if idx: doc.add_page_break()
        # The original page-top banner is a body paragraph with embedded logo; clone exactly.
        banner=_clone_paragraph(doc, originals[12])
        _format_banner(banner, prop, addr, cn)
        title=_clone_paragraph(doc, originals[13])
        if len(title.runs)>=2:
            parts=heading.split('  ',1)
            title.runs[0].text=(parts[0]+'  ') if len(parts)==2 else ''
            title.runs[1].text=parts[1] if len(parts)==2 else heading
            for run in title.runs[2:]: run.text=''
        else: _replace_para(title,heading)
        if body:
            _clone_paragraph(doc, originals[14],body)
        for j,(subtitle,cols,items) in enumerate(blocks):
            _clone_paragraph(doc, originals[16],subtitle)
            tbl = _clone_table(doc,originals[17],cols,items,total_last=(idx in (0,2,3) and j==len(blocks)-1))
            if subtitle == 'SUPPORTING DOCUMENTS':
                for row_cells in tbl.rows[1:]:
                    row_cells.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
            _clone_paragraph(doc, originals[18], '')
        if idx == 0:
            for sub, paragraph_text in [
                ('经营表现' if cn else 'OPERATING PERFORMANCE', occ_line + ' ' + noi_line),
                ('现金及分红情况' if cn else 'CASH & DISTRIBUTIONS', cash_line),
            ]:
                _clone_paragraph(doc, originals[16], sub)
                _clone_paragraph(doc, originals[14], paragraph_text)
            if extra_notes:
                _clone_paragraph(doc, originals[16], '其他经营事项' if cn else 'OTHER OPERATING NOTES')
                _clone_paragraph(doc, originals[14], extra_notes)
    if cn:
        translations = {
            'Project Overview, Operations & Distribution Summary':'项目概况、经营情况及分红摘要',
            'INVESTOR FINANCIAL REPORT':'投资人财务报告',
            'EXECUTIVE SUMMARY':'经营摘要','OCCUPANCY & RENT COLLECTION':'出租率及租金收取',
            'FINANCIAL SUMMARY':'财务汇总','DISTRIBUTION DETAILS':'分红明细','APPENDIX':'附录',
            'AT A GLANCE':'关键指标','KEY METRIC':'指标','MONTH-END VALUE':'期末数值','Cumulative Available Cash':'累计可分配现金','Available Cash':'可分配现金','KEY FINANCIAL HIGHLIGHTS':'主要财务指标','FINANCIAL METRIC':'财务指标',
            'AMOUNT':'金额','METRIC':'指标','PREVIOUS MONTH':'上个月','CURRENT MONTH':'本月','CHANGE':'变化',
            'OCCUPANCY & COLLECTION':'出租及收款情况','OPERATING RESULTS':'经营业绩',
            'CASH & RESERVES':'现金及预留资金','DISTRIBUTION OVERVIEW':'分红概览',
            'CASH POSITION':'现金情况','SOURCE & RECONCILIATION':'数据来源及核对',
            'SUPPORTING DOCUMENTS':'附属报表','DOCUMENT':'文件','DESCRIPTION':'说明',
            'Occupancy Rate':'入住率','Occupied Rooms':'已入住房间','Vacant Rooms':'空置房间',
            'Total Rooms':'总房间数','Rent Collected':'实收租金','Outstanding AR':'未收应收款','Outstanding Rent Receivable':'未收应收款',
            'Rent Income':'租金收入','Utility Income':'水电及其他服务收入','Other Income':'其他收入','Total Income':'总收入','Operating Expenses':'营业费用','Net Operating Income (NOI)':'净营业收入（NOI）',
            'Available for Distribution Review':'本月可分配现金','Ending Cash Balance':'期末现金余额',
            'Required Reserve':'要求保留资金','Reserve Shortfall':'预留资金缺口',
            'Less: Required Reserve':'减：要求保留资金','Distribution Paid This Period':'本期已分红',
            'Property':'物业','Reporting Period':'报告月份',
            'Property Tax Paid':'已付房产税','Insurance Paid':'已付保险费','Other Cash Adjustments':'其他现金调整',
            'AppFolio Income Statement':'AppFolio 物业经营利润表',
            'AppFolio Rent Roll':'AppFolio 物业出租及入住情况表',
            'Occupancy and collection results for the reporting month.':'本月入住情况及租金收取情况。',
            'Operating results shown below are based on the supplied financial records.':'以下经营业绩依据已提供的财务报表整理。',
            'Distributions are subject to management approval, cash verification and unpaid obligations.':'分红金额须经管理层批准，并核实现金余额及未付义务。',
            'The following original AppFolio reports are attached after this report for reference. They provide supporting details for the financial and occupancy information presented above.':'以下附上根据 AppFolio 原始报表生成的中文财务及出租明细。',
            'No supporting PDF documents were uploaded for this report.':'本报告未上传附属报表。',
            'As of':'截至','month-end':'月末','Appendix A':'附录 A','Appendix B':'附录 B',
            'PREPARED BY':'编制单位','CONFIDENTIAL':'保密文件',
        }
        import re
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                old=paragraph.text
                new=old
                for a,b in sorted(translations.items(),key=lambda item:-len(item[0])):
                    new=new.replace(a,b)
                if new!=old: _replace_para(paragraph,new)
                if re.search(r'[\u3400-\u9fff]',paragraph.text):
                    for run in paragraph.runs:
                        run.font.name='Microsoft YaHei'
                        run._element.get_or_add_rPr().rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia','Microsoft YaHei')
        for table in doc.tables:
            for tr in table.rows:
                for cell in tr.cells:
                    for paragraph in cell.paragraphs:
                        old=paragraph.text;new=old
                        for a,b in sorted(translations.items(),key=lambda item:-len(item[0])):
                            new=new.replace(a,b)
                        if new!=old:_replace_para(paragraph,new)
                        if re.search(r'[\u3400-\u9fff]',paragraph.text):
                            for run in paragraph.runs:
                                run.font.name='Microsoft YaHei'
                                run._element.get_or_add_rPr().rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia','Microsoft YaHei')
    if cn:
        _polish_chinese(doc)
        # Cover title and property lines: consistent typography across projects.
        for index in (2, 3, 4, 5):
            paragraph = doc.paragraphs[index]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                if run.text:
                    _set_font(run, east_asia='Microsoft YaHei', size=17 if index == 2 else (12 if index == 4 else 10.5), bold=(index in (2, 4)))
    _apply_sans_typography(doc)
    output=io.BytesIO();doc.save(output);return output.getvalue()


def make_pdf(row,p,cn=False,previous=None,appendix_documents=None):
    # Use Word-to-PDF conversion, preserving the exact template. This requires LibreOffice.
    binary=shutil.which('libreoffice') or shutil.which('soffice')
    if not binary:
        raise RuntimeError('PDF conversion requires LibreOffice. Download the Word report or deploy with LibreOffice support.')
    with tempfile.TemporaryDirectory() as tmp:
        input_path=Path(tmp)/'investor_report.docx'
        input_path.write_bytes(make_docx(row,p,cn,previous=previous,appendix_documents=appendix_documents))
        proc=subprocess.run([binary,'-env:UserInstallation=file://'+tmp+'/lo_profile','--headless','--convert-to','pdf','--outdir',tmp,str(input_path)],capture_output=True,text=True,timeout=90)
        pdf=Path(tmp)/'investor_report.pdf'
        if proc.returncode or not pdf.exists(): raise RuntimeError(f'Word-to-PDF conversion failed: {proc.stderr}')
        return pdf.read_bytes()


def append_pdf_documents(report_pdf, attachments):
    """Append original PDF pages in upload order, without rasterizing or changing them."""
    from pypdf import PdfReader, PdfWriter
    writer = PdfWriter()
    writer.append(io.BytesIO(report_pdf))
    for name, payload in attachments:
        try:
            reader = PdfReader(io.BytesIO(payload), strict=False)
            if reader.is_encrypted:
                raise ValueError('PDF is password-protected')
            if len(reader.pages) == 0:
                raise ValueError('PDF has no pages')
            writer.append(reader)
        except Exception as exc:
            raise ValueError(f'Appendix PDF {name}: {exc}') from exc
    result = io.BytesIO()
    writer.write(result)
    writer.close()
    return result.getvalue()
