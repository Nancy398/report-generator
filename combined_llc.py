"""Combined LLC reporting: property operating results stay separate; LLC cash is entered once.
No property-level cash figures are summed or reused for LLC distributions.
"""
import io, re, shutil, subprocess, tempfile
from pathlib import Path
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from engine import number, pid, period

NAVY=RGBColor(36,77,107)

def normalize(v):return re.sub(r'[^a-z0-9]','',str(v or '').lower())
def val(row,key):
    v=row.get(key,'')
    return None if v is None or str(v).strip()=='' else number(v)
def money(v):return '—' if v is None else f'${v:,.2f}'
def occupancy(v):
    n=val(v,'Occupancy')
    return None if n is None else (n*100 if n<=1 else n)
def previous_month(month):
    y,m=map(int,month.split('-'))
    return f'{y-1}-12' if m==1 else f'{y}-{m-1:02d}'

def build_model(group,month,properties,statements,llc_records,summaries,llcs):
    members=[p for p in properties if normalize(p.get('Report Group'))==normalize(group)]
    if len(members)<2:raise ValueError('Combined Report 至少需要两个 Property，且 Properties 中的 Report Group 必须一致。')
    month=period(month)
    if not month:raise ValueError('Month 必须是 YYYY-MM')
    ids=[pid(p.get('Property ID')) for p in members]
    data=[];prior=previous_month(month)
    for p,k in zip(members,ids):
        row=next((x for x in statements if pid(x.get('Property ID'))==k and period(x.get('Month (YYYY-MM)'))==month),None)
        if row is None:raise ValueError(f'{k} 缺少 {month} 的 Monthly Statements 记录。请先分别保存两个物业的运营数据。')
        last=next((x for x in statements if pid(x.get('Property ID'))==k and period(x.get('Month (YYYY-MM)'))==prior),{})
        # A missing field must remain missing rather than silently becoming $0.
        required=['Rent Income','Income','Expenses','NOI']
        missing=[x for x in required if val(row,x) is None]
        if missing:raise ValueError(f'{k} {month} 缺少 {", ".join(missing)}；请先完成该物业月报。')
        data.append({'id':k,'name':str(p.get('Property Address') or p.get('Property Name / LLC') or k),
            'rent':val(row,'Rent Income'),'utility':val(row,'Utility Income'),
            'other':val(row,'Other Income'),'income':val(row,'Income'),
            'expenses':val(row,'Expenses'),'noi':val(row,'NOI'),
            'occupancy':occupancy(row),'prior_occupancy':occupancy(last),
            'rooms':val(row,'Total Rooms'),'occupied':val(row,'Occupied Rooms')})
    lr=next((r for r in llc_records if normalize(r.get('Report Group'))==normalize(group) and period(r.get('Month'))==month),None)
    if lr is None:raise ValueError(f'LLC Monthly 缺少 {group} / {month} 的记录。合并报告的现金及分红必须由 LLC 单独提供，不能从物业记录相加。')
    required=['Ending Cash','Required Reserve','Distribution Paid']
    missing=[x for x in required if val(lr,x) is None]
    if missing:raise ValueError('LLC Monthly 缺少 '+', '.join(missing)+'；请填写 LLC 统一余额及实际分红（无分红填 0）。')
    cash=val(lr,'Ending Cash');reserve=val(lr,'Required Reserve');dist=val(lr,'Distribution Paid')
    if reserve<0 or dist<0:raise ValueError('Reserve 和 Distribution Paid 不能为负数。')
    available=max(0,cash-reserve);shortfall=max(0,reserve-cash)
    config=next((x for x in llcs if normalize(x.get('Report Group'))==normalize(group)),{})
    freq=str(config.get('Distribution Frequency','') or '').strip()
    if freq.lower() not in ('monthly','quarterly'):raise ValueError('请在 LLCs 工作表设置 Distribution Frequency = Monthly 或 Quarterly。')
    summary=next((s for s in summaries if normalize(s.get('Property'))==normalize(group) and period(s.get('Month'))==month),{})
    return {'group':group,'month':month,'properties':data,'cash':cash,'reserve':reserve,'shortfall':shortfall,
        'available':available,'distribution':dist,'frequency':freq,'summary_cn':str(summary.get('Summary CN','') or ''),
        'summary_en':str(summary.get('Summary EN','') or ''), 'total_noi':sum(x['noi'] for x in data)}

def _font(run,cn=False,bold=False,size=10,color=None):
    run.font.name='Arial';run.font.size=Pt(size);run.bold=bold
    if color:run.font.color.rgb=color
    pr=run._element.get_or_add_rPr();f=pr.rFonts
    if f is None:f=OxmlElement('w:rFonts');pr.insert(0,f)
    f.set(qn('w:eastAsia'),'Noto Sans CJK SC' if cn else 'Arial')

def add_text(doc,text,cn=False,size=10,bold=False,color=None,after=5):
    p=doc.add_paragraph();p.paragraph_format.space_after=Pt(after)
    _font(p.add_run(text),cn,bold,size,color)
    return p

def heading(doc,num,label,cn):
    p=doc.add_paragraph();p.paragraph_format.space_before=Pt(17);p.paragraph_format.space_after=Pt(8)
    _font(p.add_run(num+'   '),cn,False,17,RGBColor(154,179,195))
    _font(p.add_run(label),cn,True,15,NAVY)

def table(doc,heads,rows,cn=False):
    t=doc.add_table(rows=1,cols=len(heads));t.style='Table Grid'
    for i,h in enumerate(heads):
        cell=t.rows[0].cells[i];cell.text='';cell._tc.get_or_add_tcPr().append(OxmlElement('w:shd'))
        cell._tc.tcPr[-1].set(qn('w:fill'),'244D6B')
        _font(cell.paragraphs[0].add_run(str(h)),cn,True,9,RGBColor(255,255,255))
    for idx,record in enumerate(rows):
        cells=t.add_row().cells
        for j,v in enumerate(record):
            cells[j].text='';p=cells[j].paragraphs[0]
            if j:p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
            _font(p.add_run(str(v)),cn,False,9)
            if idx%2==1:
                sh=OxmlElement('w:shd');sh.set(qn('w:fill'),'F2F6F9');cells[j]._tc.get_or_add_tcPr().append(sh)
    doc.add_paragraph().paragraph_format.space_after=Pt(1)
    return t

def _occ_sentence(item,cn):
    now=item['occupancy'];old=item['prior_occupancy'];name=item['name']
    if now is None:return (f'{name}：本月入住率暂无数据。' if cn else f'{name}: occupancy data was not provided.')
    if old is None:return (f'{name}：本月入住率 {now:.1f}%，暂无上月可比数据。' if cn else f'{name}: occupancy was {now:.1f}%; prior-month data was unavailable.')
    d=now-old
    if abs(d)<.005:change=('环比持平' if cn else 'unchanged month over month')
    elif cn:change=f'环比{"上升" if d>0 else "下降"} {abs(d):.1f} 个百分点'
    else:change=f'{"up" if d>0 else "down"} {abs(d):.1f} percentage points month over month'
    return (f'{name}：入住率 {now:.1f}%（{change}），NOI 为 {money(item["noi"])}。' if cn else f'{name}: occupancy {now:.1f}% ({change}); NOI {money(item["noi"])}.')

def make_combined_docx(model,cn=False):
    doc=Document();sec=doc.sections[0];sec.top_margin=Inches(.72);sec.bottom_margin=Inches(.68)
    sec.left_margin=Inches(.8);sec.right_margin=Inches(.8)
    title='投资人财务报告' if cn else 'INVESTOR FINANCIAL REPORT'
    add_text(doc,'MOO HOUSING  /  '+title,cn,9,True,NAVY,7)
    add_text(doc,model['group'],cn,20,True,NAVY,2)
    add_text(doc,model['month']+'  |  '+ '合并物业报告' if cn else model['month']+'  |  COMBINED LLC REPORT',cn,10,False,None,15)
    heading(doc,'01','综合经营摘要' if cn else 'EXECUTIVE SUMMARY',cn)
    overview=[]
    for p in model['properties']:
        overview.append((p['name'], '—' if p['occupancy'] is None else f'{p["occupancy"]:.1f}%',money(p['noi'])))
    table(doc,['物业','入住率','本月 NOI'] if cn else ['PROPERTY','OCCUPANCY','MONTHLY NOI'],overview,cn)
    add_text(doc,'经营表现' if cn else 'OPERATING PERFORMANCE',cn,10,True,NAVY,4)
    for p in model['properties']:add_text(doc,_occ_sentence(p,cn),cn,9.5,False,None,5)
    add_text(doc,'LLC 现金及分红' if cn else 'LLC CASH & DISTRIBUTIONS',cn,10,True,NAVY,4)
    if model['frequency'].lower()=='quarterly':
        text=(f'截至本月末，LLC 累计可分配现金为 {money(model["available"])}。' if cn else f'Cumulative LLC cash available for distribution at month-end was {money(model["available"])}.')
        if model['distribution']>0:text+=(f'本月实际分红 {money(model["distribution"])}，上述余额为分红后金额。' if cn else f' Distributions paid this month were {money(model["distribution"])}; the balance shown is after distributions.')
        else:text+=('该余额将结转至下一次季度分红评估。' if cn else ' The balance will carry forward to the next quarterly distribution review.')
    else:text=(f'本月 LLC 期末可分配现金为 {money(model["available"])}，本月实际分红总额为 {money(model["distribution"])}。' if cn else f'LLC available cash at month-end was {money(model["available"])}; actual distributions this month totaled {money(model["distribution"])}.')
    add_text(doc,text,cn,9.5)
    notes=model['summary_cn' if cn else 'summary_en'].strip()
    if notes:
        add_text(doc,'其他经营事项' if cn else 'OTHER OPERATING NOTES',cn,10,True,NAVY,4)
        add_text(doc,notes,cn,9.5)
    for idx,p in enumerate(model['properties'],start=2):
        heading(doc,f'{idx:02d}',f'{p["name"]} · '+('物业经营' if cn else 'PROPERTY OPERATIONS'),cn)
        data=[('Occupancy', '—' if p['occupancy'] is None else f'{p["occupancy"]:.1f}%'),
              ('Rent Income',money(p['rent'])),('Utility Income',money(p['utility'])),('Other Income',money(p['other'])),
              ('Total Income',money(p['income'])),('Operating Expenses',money(p['expenses'])),('NOI',money(p['noi']))]
        table(doc,['指标','金额'] if cn else ['METRIC','AMOUNT'],data,cn)
    heading(doc,f'{len(model["properties"])+2:02d}','LLC 统一现金及分红' if cn else 'LLC CASH & DISTRIBUTION',cn)
    table(doc,['项目','金额'] if cn else ['ITEM','AMOUNT'],[
        ('Combined NOI',money(model['total_noi'])),('LLC Ending Cash',money(model['cash'])),
        ('LLC Required Reserve',money(model['reserve'])),('LLC Reserve Shortfall',money(model['shortfall'])),
        ('LLC Available to Distribute',money(model['available'])),('LLC Distribution Paid',money(model['distribution']))],cn)
    add_text(doc,'现金及 Reserve 使用 LLC Monthly 工作表中的统一记录；不对两个物业的现金余额求和。' if cn else 'LLC cash and reserves come from the single LLC Monthly record; property cash balances are not summed.',cn,8,False,None,3)
    out=io.BytesIO();doc.save(out);return out.getvalue()

def make_combined_pdf(docx_bytes):
    binary=shutil.which('libreoffice') or shutil.which('soffice')
    if not binary:raise RuntimeError('PDF conversion requires LibreOffice; Word download is still available.')
    with tempfile.TemporaryDirectory() as d:
        source=Path(d)/'combined.docx';source.write_bytes(docx_bytes)
        proc=subprocess.run([binary,'-env:UserInstallation=file://'+d+'/lo_profile','--headless','--convert-to','pdf','--outdir',d,str(source)],capture_output=True,text=True,timeout=90)
        dest=Path(d)/'combined.pdf'
        if proc.returncode or not dest.exists():raise RuntimeError(proc.stderr)
        return dest.read_bytes()
