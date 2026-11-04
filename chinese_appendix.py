"""Generate Chinese supporting PDFs from text-based AppFolio PDFs (no external AI)."""
import io,re
from pathlib import Path
from reportlab.lib.utils import ImageReader
import fitz
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, KeepTogether
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.enums import TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import landscape,A4
from xml.sax.saxutils import escape
pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
FONT='STSong-Light'
MAP={
'Income Statement - 12 Month':'物业经营利润表','Rent Roll for Owner':'物业出租及入住情况表',
'Operating Income & Expense':'营业收入与费用','Total Operating Income':'营业收入合计','Total Operating Expense':'营业费用合计','NOI - Net Operating Income':'净营业收入（NOI）','Total Income':'总收入','Total Expense':'总费用','Net Income':'净利润',
'Income':'收入','Expense':'费用','RENTS':'租金','Rent Income':'租金收入','Total RENTS':'租金合计','Parking Income':'停车收入','Concessions':'租金优惠及减免','Advertising':'广告费','CLEANING AND MAINTENANCE':'清洁与维护','General Maintenance Labor':'日常维护人工费','Landscaping & Trash Pick Up':'园艺及垃圾清运','Cleaning and Maintenance -Other':'其他清洁维护费','Site Visit Fee':'现场检查费','Maintenance Coordination Fee':'维修协调费','LEGAL AND OTHER PROFESSIONAL FEES':'法律及其他专业服务费','Accounting':'会计费','MANAGEMENT FEES':'物业管理费','Management Fees':'管理费','Commissions/Placement Fees':'佣金／招租服务费','MORTGAGE':'房屋贷款','Mortgage Interest':'房贷利息','REPAIRS':'维修','Repairs - Other':'其他维修费','Supplies':'物料用品费','UTILITIES':'公用事业费','LADWP':'LADWP（水电费）','Internet':'网络费','CAPITAL EXPENSES':'资本性支出','Appliances':'家用电器',
'Unit':'房间编号','Tenant':'租客','Status':'入住状态','Rent':'月租金','Deposit':'押金','Lease From':'租约开始','Lease To':'租约结束','Move-in':'入住日期',
'Current':'在住','Notice-Unrented':'已通知退租（未重新出租）','Notice-Rented':'已通知退租（已重新出租）','Vacant-Rented':'空置（已出租）','Vacant-Unrented':'空置（未出租）','Vacant':'空置',
'Account Name':'科目名称','Total':'合计','Properties':'物业','Period Range':'报表期间','As of':'截至日期','Units':'房间','Include Advertised Rent':'包含挂牌租金','Base Report':'基础报表','Fund Type':'资金类型','Level of Detail':'明细级别'
}
def translate(s):
    s=str(s).strip()
    if s in MAP:return MAP[s]
    if s.startswith('Total '):return MAP.get(s[6:],s[6:])+'合计' if s[6:] in MAP else s
    return s

def _lines(pdf):
    doc=fitz.open(stream=pdf,filetype='pdf')
    if doc.needs_pass:raise ValueError('加密 PDF 无法翻译')
    result=[]
    for pg in doc:
        if len(pg.get_text().strip())<30:raise ValueError('PDF 无可提取文字，请使用 AppFolio 原生导出的 PDF')
        groups={}
        for w in pg.get_text('words'):
            y=round(w[1]/3)*3
            groups.setdefault(y,[]).append((w[0],w[4]))
        for y,spans in sorted(groups.items()):
            result.append((pg.number,y,sorted(spans)))
    return result

def _styles():
    return (ParagraphStyle('zh',fontName=FONT,fontSize=8,leading=13,textColor=colors.HexColor('#263747')),
            ParagraphStyle('zhs',fontName=FONT,fontSize=7,leading=11,textColor=colors.HexColor('#263747')))

def _p(s,style):return Paragraph(escape(str(s)),style)

def _base(title,property_name,period,land=False):
    buf=io.BytesIO();page=landscape(A4) if land else A4
    doc=SimpleDocTemplate(buf,pagesize=page,leftMargin=30,rightMargin=30,topMargin=94,bottomMargin=30)
    normal,small=_styles()
    story=[]
    # Page header is drawn on the canvas so the left title and right logo
    # stay aligned, including on continuation pages.
    def draw_header(canvas,document):
        canvas.saveState()
        w,h=page
        canvas.setFillColor(colors.HexColor('#17334f'))
        canvas.setFont(FONT,13)
        canvas.drawString(30,h-40,title)
        canvas.setFont(FONT,8)
        label=str(property_name or 'Moo Housing Inc.')
        canvas.drawString(30,h-56,label[:90])
        canvas.setFont(FONT,7.5)
        canvas.drawString(30,h-70,(str(period)+'  ·  USD').strip())
        logo=Path(__file__).with_name('moo_logo.png')
        if logo.exists():
            try:
                image=ImageReader(str(logo));iw,ih=image.getSize()
                maxw,maxh=76,52
                scale=min(maxw/iw,maxh/ih)
                ww,hh=iw*scale,ih*scale
                canvas.drawImage(image,w-30-ww,h-22-hh,width=ww,height=hh,mask='auto')
            except Exception:
                pass
        canvas.setStrokeColor(colors.HexColor('#d5e0e9'))
        canvas.line(30,h-82,w-30,h-82)
        canvas.restoreState()
    doc._moo_header=draw_header
    return buf,doc,story,normal,small

def _table(story,rows,widths,small,headers=True):
    if not rows:return
    header_style=ParagraphStyle('table_header',parent=small,textColor=colors.white,fontName=FONT)
    data=[[_p(c,header_style if headers and i==0 else small) for c in row] for i,row in enumerate(rows)]
    table=Table(data,colWidths=widths,repeatRows=1 if headers else 0,hAlign='LEFT')
    cmds=[('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),6),('RIGHTPADDING',(0,0),(-1,-1),6),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),('LINEBELOW',(0,0),(-1,0),.7,colors.HexColor('#d5e0e9'))]
    if headers:cmds.extend([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#17334f')),('TEXTCOLOR',(0,0),(-1,0),colors.white)])
    for i in range(1,len(rows)):
        if i%2==0:cmds.append(('BACKGROUND',(0,i),(-1,i),colors.HexColor('#edf4f8')))
    table.setStyle(TableStyle(cmds));story.append(table)

def translate_income(pdf,property_name='',period=''):
    lines=_lines(pdf);normal,small=_styles()
    rows=[['科目名称',period or '本期','合计']]
    summary={}
    money=re.compile(r'^-?\$?[\d,]+\.\d{2}$')
    pending_label=''
    for pg,y,spans in lines:
        if not spans:continue
        text=' '.join(t for x,t in spans)
        if any(x in text for x in ('Created on','Page ','Properties:','Period Range:','Account Name','Fund Type:','Level of Detail:','Include Zero Balance','Income Statement - 12 Month')):continue
        nums=[t for x,t in spans if money.fullmatch(t)]
        if not nums:
            if text.strip() in ('Total Operating Expense','Net Income','Total Operating Income'):
                pending_label=text.strip()
            continue
        label=' '.join(t for x,t in spans if not money.fullmatch(t)).strip()
        if not label:
            nearby=[ ' '.join(t for x,t in ss) for pp,yy,ss in lines if pp==pg and 0<abs(yy-y)<=5 and not any(money.fullmatch(t) for x,t in ss)]
            label=next((v for v in nearby if v in ('Total Operating Expense','Net Income','Total Operating Income')), '')
        if not label and pending_label:label=pending_label
        pending_label=''
        if not label:continue
        # Some PDFs split a label into separate spans; preserve all labels and numeric values.
        zh=translate(label)
        rows.append([zh,nums[0],nums[1] if len(nums)>1 else nums[0]])
        if label in ('Total Operating Income','Total Operating Expense','NOI - Net Operating Income'):
            summary[label]=nums[0]
    if len(rows)<4:raise ValueError('无法识别 Income Statement 的财务科目')
    buf,doc,story,normal,small=_base('物业经营利润表',property_name,period)
    highlights=[['营业收入','营业费用','净营业收入（NOI）'],[summary.get('Total Operating Income','—'),summary.get('Total Operating Expense','—'),summary.get('NOI - Net Operating Income','—')]]
    _table(story,highlights,[177,177,177],normal);story.append(Spacer(1,16))
    _table(story,rows,[300,115,116],small)
    story.extend([Spacer(1,12),_p('财务管理报表  |  数据来源：原始 AppFolio Income Statement',small)])
    doc.build(story,onFirstPage=doc._moo_header,onLaterPages=doc._moo_header);return buf.getvalue()

def translate_rent_roll(pdf,property_name='',period=''):
    lines=_lines(pdf);room_re=re.compile(r'\b(?:Rm|Room|Bed)\s*[-#]?\s*\w+',re.I)
    status_re=re.compile(r'\b(?:Notice-Unrented|Notice-Rented|Vacant-Unrented|Vacant-Rented|Current|Vacant)\b',re.I)
    date_re=re.compile(r'\d{2}/\d{2}/\d{4}')
    amt_re=re.compile(r'^-?[\d,]+\.\d{2}$')
    rows=[['房间编号','租客','入住状态','月租金','押金','租约开始','租约结束','入住日期']]
    seen=set()
    for pg,y,spans in lines:
        txt=' '.join(t for x,t in spans)
        if not room_re.search(txt) or not status_re.search(txt):continue
        match=status_re.search(txt)
        left=txt[:match.start()].strip();right=txt[match.end():].strip()
        # Unit usually appears before tenant; split at first date/status column via x positions when possible.
        left_spans=[(x,t) for x,t in spans if x<340]
        unit=' '.join(t for x,t in left_spans if x<220).strip()
        tenant=' '.join(t for x,t in left_spans if 220<=x<340).strip()
        if not room_re.search(unit):
            unit=left.split('  ')[0] if '  ' in left else left
            tenant=left[len(unit):].strip()
        if not room_re.search(unit) or unit in seen:continue
        seen.add(unit)
        amounts=[t for t in right.split() if amt_re.fullmatch(t)]
        dates=date_re.findall(right)
        rows.append([unit,tenant,translate(match.group()),*(amounts+['—','—'])[:2],*(dates+['—','—','—'])[:3]])
    if len(rows)<2:raise ValueError('无法识别 Rent Roll 房间记录')
    buf,doc,story,normal,small=_base('物业出租及入住情况表',property_name,period,land=True)
    occupied=sum(1 for r in rows[1:] if not r[2].startswith('空置'))
    story.extend([_p(f'实际房间：{len(rows)-1}  |  已入住：{occupied}  |  当前出租率：{occupied/(len(rows)-1):.1%}',normal),Spacer(1,12)])
    _table(story,rows,[158,86,104,63,63,76,76,76],small)
    story.extend([Spacer(1,10),_p('注：仅统计实际 Rm 房间行，不计入楼栋及楼层主行。数据来源：AppFolio Rent Roll。',small)])
    doc.build(story,onFirstPage=doc._moo_header,onLaterPages=doc._moo_header);return buf.getvalue()
