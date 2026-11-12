"""English investor-facing rent roll with corrected whole-unit lease status."""
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Spacer
from chinese_appendix import _base, _table, _p
from pdf_appfolio import rent_roll_details

STATUS = {
    'Current': 'Occupied',
    'Notice-Unrented': 'Occupied - Notice Given',
    'Notice-Rented': 'Occupied - Notice Given / Re-Leased',
    'Vacant-Rented': 'Vacant - Leased for Future Move-in',
    'Vacant-Unrented': 'Vacant - Unleased',
    'Vacant': 'Vacant',
}

def corrected_rent_roll(pdf, property_name='', period=''):
    rooms = rent_roll_details(pdf)['rooms']
    occupied = sum(bool(room['occupied']) for room in rooms)
    rows = [['Room / Unit', 'Tenant / Whole-Unit Tenant', 'Lease Status', 'Monthly Rent', 'Deposit', 'Lease Start', 'Lease End', 'Move-in']]
    for room in rooms:
        if room.get('whole_unit_lease') and room.get('display_status') == '整租已出租':
            status = 'Leased under Whole-Unit Agreement'
        else:
            status = STATUS.get(room['status'], room['status'])
        rows.append([room['unit'], room.get('display_tenant') or '—', status,
                     room.get('rent') or '—', room.get('deposit') or '—',
                     room.get('lease_from') or '—', room.get('lease_to') or '—', room.get('move_in') or '—'])
    buf, doc, story, normal, small = _base('Corrected Rent Roll / Occupancy', property_name, period, land=True)
    story.extend([_p(f'Physical rooms: {len(rooms)}  |  Occupied rooms: {occupied}  |  Occupancy: {occupied / len(rooms):.1%}', normal), Spacer(1, 12)])
    _table(story, rows, [145, 100, 110, 60, 57, 74, 74, 74], small)
    story.extend([Spacer(1, 10), _p('Whole-unit leases override vacant child-room statuses for occupancy purposes. Parent lease rent is not duplicated across rooms. Source: AppFolio Rent Roll; status presentation adjusted for investor reporting.', small)])
    doc.build(story)
    return buf.getvalue()
