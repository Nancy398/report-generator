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
from docx.shared import Inches

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

def make_docx(row, p, cn=False, previous=None, appendix_documents=None):
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
    summary=str(_get(row,'Summary','Executive Summary','Notes',default=''))
    if not summary.strip(): summary='[Enter your Executive Summary in the Streamlit Summary field.]'
    rent=_cash(_get(row,'Rent Collected','Rent Collection','Collected Rent'))
    ar=_cash(_get(row,'Outstanding AR','Accounts Receivable','AR'))
    income=_cash(_get(row,'Income','Total Income'))
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
                comp('Rent Collected','Rent Collected','money'),
                comp('Outstanding AR','Outstanding AR','money')]
    sections=[
      ('01  EXECUTIVE SUMMARY',summary,[('KEY FINANCIAL HIGHLIGHTS',['FINANCIAL METRIC','AMOUNT'],[
          ('Occupancy Rate',occupancy),('Total Income',income),('Net Operating Income (NOI)',noi),('Available for Distribution Review',available)])]),
      ('02  OCCUPANCY & RENT COLLECTION','Occupancy and collection results for the reporting month.',[
          ('OCCUPANCY & COLLECTION',['METRIC','PREVIOUS MONTH','CURRENT MONTH','CHANGE'],comparison)]),
      ('03  FINANCIAL SUMMARY','Operating results shown below are based on the supplied financial records.',[
          ('OPERATING RESULTS',['ITEM','AMOUNT'],[('Total Income',income),('Operating Expenses',expenses),('Net Operating Income (NOI)',noi)]),
          ('CASH & RESERVES',['ITEM','AMOUNT'],[('Ending Cash Balance',cash),('Required Reserve',reserve),('Reserve Shortfall',shortfall)])]),
      ('04  DISTRIBUTION DETAILS','Distributions are subject to management approval, cash verification and unpaid obligations.',[
          ('DISTRIBUTION OVERVIEW',['CASH POSITION','USD'],[('Ending Cash Balance',cash),('Less: Required Reserve',reserve),('Available for Distribution Review',available),('Distribution Paid This Period',distribution)])]),
      ('APPENDIX','Supporting data and financial statements may be attached separately.',[
          ('SOURCE & RECONCILIATION',['ITEM','DETAILS'],[('Property',prop),('Reporting Period',period),('Verified Ending Bank Balance',_cash(_get(row,'Verified Ending Bank Balance'))),('Property Tax Paid',_cash(_get(row,'Property Tax Paid'))),('Insurance Paid',_cash(_get(row,'Insurance Paid'))),('Other Cash Adjustments',_cash(_get(row,'Other Cash Adjustments')))])])
    ]
    for idx,(heading,body,blocks) in enumerate(sections):
        if idx: doc.add_page_break()
        # The original page-top banner is a body paragraph with embedded logo; clone exactly.
        banner=_clone_paragraph(doc, originals[12])
        for run in banner.runs:
            if '1438 West 37th Drive' in run.text:
                run.text=run.text.replace('1438 West 37th Drive · Los Angeles, CA',f'{prop} · {addr}')
        title=_clone_paragraph(doc, originals[13])
        if len(title.runs)>=2:
            parts=heading.split('  ',1)
            title.runs[0].text=(parts[0]+'  ') if len(parts)==2 else ''
            title.runs[1].text=parts[1] if len(parts)==2 else heading
            for run in title.runs[2:]: run.text=''
        else: _replace_para(title,heading)
        _clone_paragraph(doc, originals[14],body)
        for j,(subtitle,cols,items) in enumerate(blocks):
            _clone_paragraph(doc, originals[16],subtitle)
            _clone_table(doc,originals[17],cols,items,total_last=(idx in (0,2,3) and j==len(blocks)-1))
            _clone_paragraph(doc, originals[18], '')
    output=io.BytesIO();doc.save(output);return output.getvalue()


def make_pdf(row,p,cn=False,previous=None):
    # Use Word-to-PDF conversion, preserving the exact template. This requires LibreOffice.
    binary=shutil.which('libreoffice') or shutil.which('soffice')
    if not binary:
        raise RuntimeError('PDF conversion requires LibreOffice. Download the Word report or deploy with LibreOffice support.')
    with tempfile.TemporaryDirectory() as tmp:
        input_path=Path(tmp)/'investor_report.docx'
        input_path.write_bytes(make_docx(row,p,cn,previous=previous))
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
