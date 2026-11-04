"""Extract values from text-based AppFolio PDF exports; never OCR or guess missing values."""
import io,re
from collections import defaultdict
import fitz

AMOUNT=re.compile(r'^\(?-?\$?[\d,]+\.\d{2}\)?$')
ROOM=re.compile(r'\s+-\s+(?:Rm|Room|Bed)\s*[-#]?\s*\w+',re.I)
STATUSES=('Current','Notice-Unrented','Notice-Rented','Vacant-Rented','Vacant-Unrented','Vacant')

def _lines(data):
    doc=fitz.open(stream=data,filetype='pdf')
    if doc.needs_pass:raise ValueError('Password-protected PDF is not supported')
    lines=[]
    for page in doc:
        if len(page.get_text().strip())<60:raise ValueError('PDF has no extractable text; export a text-based PDF from AppFolio.')
        groups=defaultdict(list)
        for w in page.get_text('words'):
            groups[round(w[1]/2)*2].append(w)
        for y,ws in sorted(groups.items()):
            ws.sort(key=lambda w:w[0]);lines.append((page.number,y,ws))
    return lines

def _tokens(ws,xmin=0):
    return ' '.join(w[4] for w in ws if w[0]>=xmin)

def _amounts(ws):
    return [float(w[4].replace(',','').replace('$','').replace('(','-').replace(')','')) for w in ws if w[0]>420 and AMOUNT.fullmatch(w[4])]

def parse_income_pdf(data):
    lines=_lines(data); found={};period=None
    for page,y,ws in lines:
        text=_tokens(ws);label=_tokens([w for w in ws if w[0]<430]).strip().lower()
        if 'period range:' in text.lower():
            m=re.search(r'Period Range:\s*([A-Za-z]{3}\s+\d{4})\s+to\s+([A-Za-z]{3}\s+\d{4})',text,re.I)
            if m:
                from datetime import datetime
                start=datetime.strptime(m.group(1),'%b %Y').strftime('%Y-%m')
                end=datetime.strptime(m.group(2),'%b %Y').strftime('%Y-%m')
                if start!=end:raise ValueError('Income Statement must be a single reporting month, not multiple months.')
                period=start
        if label in ('total operating income','total operating expense','total operating expenses','noi - net operating income'):
            nums=_amounts(ws)
            if not nums:
                nums=[a for pg,yy,ww in lines if pg==page and abs(yy-y)<=4 for a in _amounts(ww)]
            if not nums:continue
            key='income' if label=='total operating income' else 'expenses' if 'expense' in label else 'noi'
            found[key]=nums[0]
    if len(found)<3:raise ValueError('Missing Total Operating Income, Total Operating Expense or NOI in PDF.')
    if abs(found['income']-found['expenses']-found['noi'])>.05:raise ValueError('Income / Expenses / NOI reconciliation failed.')
    found['period']=period
    return found

def parse_rent_roll_pdf(data):
    lines=_lines(data);rooms=[];asof=None;parent_occupied={};seen=set()
    for page,y,ws in lines:
        text=_tokens(ws)
        if text.lower().startswith('as of:'):
            m=re.search(r'(\d{2}/\d{2}/\d{4})',text)
            if m:
                from datetime import datetime
                asof=datetime.strptime(m.group(1),'%m/%d/%Y').strftime('%Y-%m')
        unit=_tokens([w for w in ws if w[0]<330]).strip()
        status=next((s for s in STATUSES if re.search(r'(?<!\w)'+re.escape(s)+r'(?!\w)',_tokens([w for w in ws if 325<=w[0]<600]),re.I)),None)
        if not status:continue
        if ROOM.search(unit):
            if unit in seen:raise ValueError('Duplicate room record: '+unit)
            seen.add(unit);rooms.append((unit,status))
        elif ' - ' in unit and status.lower() not in ('vacant','vacant-unrented','vacant-rented'):
            parent_occupied[unit.split(' - ')[0].lower()]=True
    if not rooms:raise ValueError('No room records detected in Rent Roll PDF.')
    occupied=0
    for unit,status in rooms:
        base=unit.split(' - ')[0].lower()
        occupied+= int(not status.lower().startswith('vacant') or parent_occupied.get(base,False))
    total=len(rooms)
    return {'rooms':total,'occupied':occupied,'vacant':total-occupied,'occupancy':occupied/total,'period':asof,'warnings':[]}
