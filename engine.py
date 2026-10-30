"""Pure calculation and AppFolio parsing routines. No Streamlit dependency."""
import io,re,math
import pandas as pd

def number(x, default=0.):
    try:
        if x is None or pd.isna(x) or str(x).strip()=='': return default
        s=str(x).replace(',','').replace('$','').replace('(','-').replace(')','').strip()
        return float(s)
    except (ValueError,TypeError):return default

def pid(x):
    s=re.sub(r'[^A-Z0-9]','',str(x).upper())
    return s[2:] if s.startswith('MH') else s

def period(x):
    if x is None or str(x).strip()=='':return ''
    s=str(x).strip()
    if re.fullmatch(r'\d{4}-\d{2}',s):return s
    try:return pd.to_datetime(x).strftime('%Y-%m')
    except:return ''

def read_appfolio(file):
    raw=file.getvalue() if hasattr(file,'getvalue') else file
    book=pd.ExcelFile(io.BytesIO(raw))
    return [pd.read_excel(book,sheet_name=s,header=None,dtype=object).fillna('') for s in book.sheet_names]

def parse_income(sheets):
    results=[]
    for df in sheets:
        found={};meta=[]
        for _,row in df.iterrows():
            vals=list(row)
            for j,v in enumerate(vals):
                label=str(v).strip().lower()
                if 'period range:' in label or 'owned by:' in label:meta.append(str(v))
                aliases={'income':['total operating income','total income'], 'expenses':['total operating expenses','total operating expense','total expenses'], 'noi':['net operating income','noi']}
                for key, names in aliases.items():
                    if label in names:
                        nums=[number(x,None) for x in vals[j+1:]]
                        nums=[n for n in nums if n is not None]
                        if nums:found[key]=nums[-1]
        if 'noi' not in found and 'income' in found and 'expenses' in found:found['noi']=found['income']-found['expenses']
        if found:results.append((len(found),found,meta))
    if not results:raise ValueError('Income Statement 中未找到 Total Operating Income / Expenses / NOI；请检查 AppFolio 导出格式。')
    _,d,meta=max(results,key=lambda x:x[0]);d['metadata']=meta
    return d

def parse_rent_roll(sheets):
    """Count actual room records; ignore AppFolio property/building/floor headings.

    A whole-unit lease counts its underlying rooms as occupied only when the
    corresponding parent row actually has a tenant and a non-vacant status.
    """
    best = None
    for df in sheets:
        for i in range(min(len(df), 45)):
            cells = [str(x).strip().lower() for x in df.iloc[i]]
            if 'unit' in cells and 'tenant' in cells:
                best = (df, i, cells)
                break
        if best:
            break
    if not best:
        raise ValueError('Rent Roll 未找到 Unit / Tenant 表头')
    df, hi, head = best
    ix, tx = head.index('unit'), head.index('tenant')
    sx = next((j for j, h in enumerate(head) if 'status' in h), None)
    rows = []
    for _, row in df.iloc[hi + 1:].iterrows():
        vals = list(row)
        unit = str(vals[ix]).strip()
        if not unit or unit.lower() in ('nan', 'none') or unit.lower().startswith(('total', 'units:', 'property:')):
            continue
        tenant = str(vals[tx]).strip()
        status = str(vals[sx]).strip().lower() if sx is not None else ''
        occupied = (tenant.lower() not in ('', 'nan', 'none', '-', 'vacant')
                    and not status.startswith('vacant'))
        rows.append((unit, occupied, status))
    # AppFolio exports e.g. '1186 W 37th Pl FFL 2 - Rm 1'.
    # Only a suffix that is an actual room/bed identifier counts as a room.
    room_pattern = re.compile(r'^(.*?)\s+-\s+(?:rm|room|bed)\s*[-#]?\s*\w+', re.I)
    room_rows = []
    parent_occupied = {}
    for unit, occupied, status in rows:
        m = room_pattern.match(unit)
        if m:
            room_rows.append((m.group(1).strip().lower(), occupied))
        else:
            # Parent rows often say Vacant-Unrented even when their rooms are leased.
            # Such parent rows are NOT additional rentable rooms.
            if occupied:
                parent_occupied[unit.split(' - ')[0].strip().lower()] = True
    if not room_rows:
        raise ValueError('未找到可统计的房间 (Rm / Room / Bed)')
    total = len(room_rows)
    occupied_count = sum(occupied or parent_occupied.get(base, False)
                         for base, occupied in room_rows)
    warnings = []
    if any(is_occupied for unit, is_occupied, status in rows if not room_pattern.match(unit)):
        warnings.append('检测到已出租整租主行，按其下属房间计算，不重复计数。')
    return {'rooms': total, 'occupied': occupied_count,
            'vacant': total - occupied_count,
            'occupancy': occupied_count / total, 'warnings': warnings}

def calculate(properties, statements):
    """One property-month row, chronological. Tax/insurance payments reduce required reserve.
    Cash movement: NOI + cash adjustments - distribution - cash expenses not included in NOI.
    Verified bank balance overrides modeled ending cash and carries forward.
    """
    pmap={pid(r.get('Property ID')):r for r in properties if pid(r.get('Property ID'))}
    grouped={};issues=[];out=[]
    for r in statements:
        k=pid(r.get('Property ID'));m=period(r.get('Month (YYYY-MM)'))
        if not k or not m:continue
        if k not in pmap:issues.append(f'{k} {m}: Property ID 不在 Properties');continue
        if (k,m) in grouped:issues.append(f'{k} {m}: 重复月份');continue
        grouped[(k,m)]=r
    for k in pmap:
        p=pmap[k];cash=number(p.get('Opening Cash'));reserve=number(p.get('Opening Required Reserve'))
        months=sorted(m for (x,m) in grouped if x==k)
        last=None
        for m in months:
            r=grouped[(k,m)]
            if last:
                months_gap=(int(m[:4])-int(last[:4]))*12+int(m[5:])-int(last[5:])
                if months_gap>1:issues.append(f'{k}: {last} → {m} 缺少 {months_gap-1} 个月；累计 Reserve 可能偏低')
            noi=number(r.get('NOI'))
            tax=number(r.get('Property Tax Paid'));ins=number(r.get('Insurance Paid'))
            other=number(r.get('Other Cash Expense'));dist=number(r.get('Distribution Paid'))
            adjustments=number(r.get('Other Cash Adjustments'))
            # Default YES: Tax/Insurance already recognized in NOI. Still release reserve.
            already=str(r.get('Tax/Insurance Included in NOI?','Yes')).strip().lower() not in ('no','false','0')
            extra=tax+ins if not already else 0
            modeled=cash+noi+adjustments-other-dist-extra
            verified=number(r.get('Verified Ending Bank Balance'),None)
            cash=verified if verified is not None else modeled
            reserve=max(0,reserve+number(p.get('Monthly Tax Reserve'))+number(p.get('Monthly Insurance Reserve'))-tax-ins)
            if cash<0:issues.append(f'{k} {m}: Ending Cash 为负数')
            if verified is not None and abs(verified-modeled)>0.01:issues.append(f'{k} {m}: 银行余额与推算差额 ${verified-modeled:,.2f}，请核对现金调整项')
            available=max(0,cash-reserve)
            shortfall=max(0,reserve-cash)
            out.append({'Property ID':k,'Month (YYYY-MM)':m,'Income':number(r.get('Income')),'Expenses':number(r.get('Expenses')),'NOI':noi,
                        'Occupancy':number(r.get('Occupancy'),None),'Occupied Rooms':number(r.get('Occupied Rooms'),None),'Total Rooms':number(r.get('Total Rooms'),None),
                        'Beginning Cash':round(modeled-noi-adjustments+other+dist+extra,2),'Modeled Ending Cash':round(modeled,2),'Ending Cash':round(cash,2),
                        'Required Reserve':round(reserve,2),'Reserve Shortfall':round(shortfall,2),'Available to Distribute':round(available,2),
                        'Distribution Paid':dist,'Property Tax Paid':tax,'Insurance Paid':ins,'Bank Reconciled':verified is not None,
                        'Status':str(r.get('Status','Draft') or 'Draft')})
            last=m
    return out,issues
