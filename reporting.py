import io
from docx import Document
from docx.shared import Inches,Pt,RGBColor
from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer,Table,TableStyle,PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pathlib import Path

def narrative(row,cn):
    occ=row.get('Occupancy');vac=row.get('Total Rooms')
    if occ is None:return '本期未提供可核对的 Rent Roll，出租率尚未确认。' if cn else 'Rent Roll was not provided; occupancy is pending verification.'
    rate=f'{occ:.1%}'
    if cn:return f'本期房间出租率为 {rate}。'+('截至本期存在现金储备缺口，建议暂缓分红，优先补足储备。' if row['Reserve Shortfall']>0 else '当前现金超过累计所需储备，超出部分可供分红审核。')
    return f'Room occupancy for the reporting period was {rate}. '+('Cash is below the required reserve; distributions should be deferred pending reserve replenishment.' if row['Reserve Shortfall']>0 else 'Cash exceeds the required reserve; excess cash may be reviewed for distribution.')

def sections(row,p,cn):
    money=lambda v:f'${float(v or 0):,.2f}'
    fin=[('营业收入' if cn else 'Total Income',money(row['Income'])),('运营费用' if cn else 'Operating Expenses',money(row['Expenses'])),('净运营收入（NOI）' if cn else 'Net Operating Income (NOI)',money(row['NOI']))]
    ops=[('房间出租率' if cn else 'Room Occupancy',f"{row['Occupancy']:.1%}" if row.get('Occupancy') is not None else 'N/A'),('已出租 / 总房间' if cn else 'Occupied / Total Rooms',f"{int(row['Occupied Rooms'])} / {int(row['Total Rooms'])}" if row.get('Total Rooms') is not None else 'N/A')]
    cash=[('期末现金余额' if cn else 'Ending Cash Balance',money(row['Ending Cash'])),('累计应保留储备' if cn else 'Required Reserve Balance',money(row['Required Reserve']))]
    if row['Reserve Shortfall']>0:cash.append(('储备缺口' if cn else 'Reserve Shortfall',money(row['Reserve Shortfall'])))
    cash += [('可供分红审核现金' if cn else 'Cash Available for Distribution Review',money(row['Available to Distribute'])),('本期已支付分红' if cn else 'Distribution Paid This Period',money(row['Distribution Paid']))]
    return [('01  财务表现' if cn else '01  Financial Performance',fin),('02  运营表现' if cn else '02  Operating Performance',ops),('03  现金及储备' if cn else '03  Cash & Reserve Summary',cash)]

def make_docx(row,p,cn):
    d=Document();sec=d.sections[0];sec.top_margin=Inches(.8);sec.bottom_margin=Inches(.7)
    style=d.styles['Normal'];style.font.name='Calibri';style.font.size=Pt(10)
    d.add_paragraph('MOO HOUSING INC.','Subtitle')
    d.add_heading('投资人财务报告' if cn else 'INVESTOR FINANCIAL REPORT',0)
    d.add_paragraph(str(p.get('Property Name / LLC','')))
    d.add_paragraph(str(p.get('Property Address','')))
    d.add_paragraph(('报告期间：' if cn else 'Reporting Period: ')+row['Month (YYYY-MM)'])
    d.add_paragraph('CONFIDENTIAL / 仅供投资人审阅')
    d.add_page_break()
    for heading,items in sections(row,p,cn):
        d.add_heading(heading,level=1)
        t=d.add_table(rows=0,cols=2);t.style='Light Shading Accent 1'
        for k,v in items:
            a,b=t.add_row().cells;a.text=k;b.text=v
        d.add_paragraph('')
    d.add_heading('04  经营总结' if cn else '04  Operating Commentary',level=1)
    d.add_paragraph(narrative(row,cn))
    d.add_paragraph('注：NOI 不一定等于现金流。储备不是额外的 P&L 费用；分红前应核对银行余额和未付款项。' if cn else 'Note: NOI is not necessarily cash flow. Reserve designations are not additional P&L expenses. Verify bank cash and outstanding obligations before distribution.')
    out=io.BytesIO();d.save(out);return out.getvalue()

def make_pdf(row,p,cn):
    font='Helvetica'
    for candidate in ['/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc','/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc','/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf']:
        if Path(candidate).exists():
            try:pdfmetrics.registerFont(TTFont('MooUnicode',candidate));font='MooUnicode';break
            except Exception:pass
    if cn and font=='Helvetica':
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'));font='STSong-Light'
    styles=getSampleStyleSheet();styles.add(ParagraphStyle(name='MooTitle',parent=styles['Title'],fontName=font,fontSize=20,leading=28,textColor=colors.HexColor('#183451')))
    styles.add(ParagraphStyle(name='MooBody',parent=styles['Normal'],fontName=font,fontSize=10,leading=16))
    styles.add(ParagraphStyle(name='MooHead',parent=styles['Heading2'],fontName=font,textColor=colors.HexColor('#183451')))
    buf=io.BytesIO();doc=SimpleDocTemplate(buf,pagesize=(595,842),leftMargin=48,rightMargin=48,topMargin=48,bottomMargin=48)
    story=[Paragraph('MOO HOUSING INC.',styles['MooHead']),Spacer(1,70),Paragraph('投资人财务报告' if cn else 'INVESTOR FINANCIAL REPORT',styles['MooTitle']),Spacer(1,30),Paragraph(str(p.get('Property Name / LLC','')),styles['MooBody']),Paragraph(str(p.get('Property Address','')),styles['MooBody']),Paragraph(row['Month (YYYY-MM)'],styles['MooBody']),Spacer(1,60),Paragraph('CONFIDENTIAL',styles['MooBody']),PageBreak()]
    for heading,items in sections(row,p,cn):
        story.append(Paragraph(heading,styles['MooHead']))
        table=Table([[Paragraph(str(k),styles['MooBody']),Paragraph(str(v),styles['MooBody'])] for k,v in items],colWidths=[330,160],hAlign='LEFT')
        table.setStyle(TableStyle([('ROWBACKGROUNDS',(0,0),(-1,-1),[colors.white,colors.HexColor('#EDF3F7')]),('BOTTOMPADDING',(0,0),(-1,-1),9),('TOPPADDING',(0,0),(-1,-1),9)]))
        story.extend([table,Spacer(1,16)])
    story.extend([Paragraph('04  经营总结' if cn else '04  Operating Commentary',styles['MooHead']),Paragraph(narrative(row,cn),styles['MooBody']),Spacer(1,20),Paragraph('NOI 不等于现金流；正式分红前须核对银行及未付款项。' if cn else 'NOI is not necessarily cash flow. Verify bank cash and obligations before distribution.',styles['MooBody'])]);doc.build(story)
    return buf.getvalue()
