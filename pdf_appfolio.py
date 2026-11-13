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
    return [float(w[4].replace(',','').replace('$','').replace('(','-').replace(')','')) for w in ws if w[0]>350 and AMOUNT.fullmatch(w[4])]

def parse_income_pdf(data):
    lines=_lines(data); found={};period=None
    for page,y,ws in lines:
        text=_tokens(ws);label=_tokens([w for w in ws if w[0]<350]).strip().lower()
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
    # Management reporting convention: all FEES are grouped as Utility Income.
    # Keep the original AppFolio total income / NOI unchanged.
    fees = None
    rents = None
    for page, y, ws in lines:
        label = _tokens([w for w in ws if w[0] < 350]).strip().lower()
        if label in ('total fees', 'total rents'):
            nums = _amounts(ws)
            if not nums:
                nums = [a for pg, yy, ww in lines if pg == page and abs(yy-y) <= 4 for a in _amounts(ww)]
            if nums:
                if label == 'total fees': fees = nums[0]
                else: rents = nums[0]
    # Report convention agreed with management: Rent Income is the remaining
    # operating income after grouping all FEES as Utility Income. This may
    # include concessions and other non-FEES rent adjustments.
    if fees is not None:
        found['utility_income'] = round(fees, 2)
        found['rent_income'] = round(found['income'] - fees, 2)
    elif rents is not None:
        # Without a FEES subtotal, do not guess the missing split.
        found['rent_income'] = round(rents, 2)
        found['other_income'] = round(found['income'] - rents, 2)
    found['period']=period
    return found

def _normalize_unit(value):
    """Normalize spacing/case without merging distinct fractional street addresses."""
    return re.sub(r'\s+', ' ', str(value or '').strip()).casefold()


def _lease_scope(unit):
    """Recognize whole-address and floor-specific parent leases without mixing addresses."""
    if ' - ' not in unit:
        return None
    left, right = (_normalize_unit(part) for part in unit.split(' - ', 1))
    left_address, left_floor = _room_scope(left)
    right_address, right_floor = _room_scope(right)
    if left_address != right_address:
        return None
    # A repeated floor-specific name is a floor lease, not an address-wide lease.
    if left == right:
        return (left_address, left_floor)
    # An address-wide left side can explicitly name a floor on the right.
    if left_floor is None and right_floor is not None:
        return (left_address, right_floor)
    return None


def _room_scope(base):
    base = _normalize_unit(base)
    match = re.fullmatch(r'(.+?)\s+ffl\s*(\w+)', base, re.I)
    if match:
        return match.group(1), match.group(2).casefold()
    return base, None


def rent_roll_details(data):
    """Count physical rooms, inheriting occupied whole-address/floor leases.

    Exact floor parents take precedence over address-wide parents. A vacant
    floor parent does not cancel an occupied address-wide lease, but an
    occupied floor parent never covers another floor. Parent rent is never
    copied into child-room rent.
    """
    lines = _lines(data)
    parents = {}
    rooms = []
    seen = set()
    asof = None
    status_pattern = re.compile(r'(?<!\w)(?:' + '|'.join(map(re.escape, STATUSES)) + r')(?!\w)', re.I)
    for page, y, ws in lines:
        text = _tokens(ws)
        if text.lower().startswith('as of:'):
            match = re.search(r'(\d{2}/\d{2}/\d{4})', text)
            if match:
                from datetime import datetime
                asof = datetime.strptime(match.group(1), '%m/%d/%Y').strftime('%Y-%m')
        unit = _tokens([w for w in ws if w[0] < 260]).strip()
        if ' - ' not in unit:
            continue
        status_text = _tokens([w for w in ws if 345 <= w[0] < 450])
        status_match = status_pattern.search(status_text)
        if not status_match:
            continue
        status = status_match.group(0)
        tenant = _tokens([w for w in ws if 260 <= w[0] < 345]).strip()
        def column_text(lo, hi):
            return _tokens([w for w in ws if lo <= w[0] < hi]).strip()
        record = {
            'unit': unit, 'status': status, 'tenant': tenant,
            'rent': column_text(450, 530), 'deposit': column_text(530, 580),
            'lease_from': column_text(580, 650),
            'lease_to': column_text(650, 705), 'move_in': column_text(705, 800),
        }
        base, suffix = unit.rsplit(' - ', 1)
        if re.fullmatch(r'(?:Rm|Room|Bed)\s*[-#]?\s*\w+', suffix.strip(), re.I):
            key = _normalize_unit(unit)
            if key in seen:
                raise ValueError('Duplicate room record: ' + unit)
            seen.add(key)
            record['base'] = _normalize_unit(base)
            rooms.append(record)
        else:
            scope = _lease_scope(unit)
            if scope:
                if scope in parents and parents[scope]['status'].casefold() != status.casefold():
                    raise ValueError('Conflicting whole-unit lease statuses for: ' + unit)
                parents[scope] = record
    if not rooms:
        raise ValueError('No room records detected in Rent Roll PDF.')
    for room in rooms:
        address, floor = _room_scope(room['base'])
        candidates = []
        if floor is not None:
            candidates.append(parents.get((address, floor)))
        candidates.append(parents.get((address, None)))
        parent = next((p for p in candidates if p and p['status'].casefold() in ('current', 'notice-unrented', 'notice-rented')), None)
        inherited = parent is not None
        direct = room['status'].casefold() in ('current', 'notice-unrented', 'notice-rented')
        room['occupied'] = direct or inherited
        room['whole_unit_lease'] = inherited
        if inherited and not direct:
            room['display_status'] = '整租已出租'
            room['display_tenant'] = parent['tenant'] or '整租租客'
            for field in ('lease_from', 'lease_to', 'move_in'):
                if not room[field]:
                    room[field] = parent[field]
        else:
            room['display_status'] = room['status']
            room['display_tenant'] = room['tenant']
    return {'rooms': rooms, 'period': asof}


def parse_rent_roll_pdf(data):
    details = rent_roll_details(data)
    rooms = details['rooms']
    occupied = sum(bool(r['occupied']) for r in rooms)
    total = len(rooms)
    return {
        'rooms': total, 'occupied': occupied, 'vacant': total - occupied,
        'occupancy': occupied / total, 'period': details['period'], 'warnings': [],
    }
