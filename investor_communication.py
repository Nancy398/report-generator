"""Investor communication drafts; no emails or WeChat messages are sent."""
import re
import streamlit as st


def money(value):
    try:
        return f"${float(str(value).replace(',', '').replace('$', '')):,.2f}"
    except (ValueError, TypeError):
        return 'N/A'


def _month(month, cn):
    match = re.fullmatch(r'(\d{4})-(\d{2})', str(month))
    if not match:
        return str(month)
    year, m = match.groups()
    return f'{year}年{int(m)}月' if cn else f'{year}-{m}'


def drafts(name, month, operation, cash, notes='', cn=True):
    label = _month(month, cn)
    operation = str(operation or '').strip()
    cash = str(cash or '').strip()
    notes = str(notes or '').strip()
    if cn:
        subject = f'{name}｜{label}投资人财务报告'
        paragraphs = [f'您好，', f'附件为 {name} {label}的投资人财务报告。',
                      f'经营情况：{operation}', f'现金及分红：{cash}']
        if notes:
            paragraphs.append(f'其他经营事项：{notes}')
        paragraphs.extend(['请查阅附件。如有任何问题，欢迎随时与我们联系。', 'Moo Housing'])
        email = '\n\n'.join(paragraphs)
        wechat_parts = [f'【{name}｜{label}投资人报告】', operation, cash]
        if notes:
            wechat_parts.append(notes)
        wechat_parts.append('完整财务报告请参阅随附文件。')
        wechat = '\n'.join(wechat_parts)
    else:
        subject = f'{name} | Investor Report | {label}'
        paragraphs = ['Hello,', f'Please find attached the investor report for {name} ({label}).',
                      f'Operating performance: {operation}', f'Cash and distributions: {cash}']
        if notes:
            paragraphs.append(f'Additional notes: {notes}')
        paragraphs.extend(['Please review the attached report and let us know if you have any questions.', 'Moo Housing'])
        email = '\n\n'.join(paragraphs)
        wechat_parts = [f'【{name} | {label} Investor Report】', operation, cash]
        if notes:
            wechat_parts.append(notes)
        wechat_parts.append('Please see the attached report for full details.')
        wechat = '\n'.join(wechat_parts)
    return subject, email, wechat



def _numeric(row, *keys):
    for key in keys:
        value = (row or {}).get(key)
        if value is None or str(value).strip() == '':
            continue
        try:
            return float(str(value).replace(',', '').replace('$', '').replace('%', ''))
        except (ValueError, TypeError):
            pass
    return None


def _formatted_month(month):
    from datetime import datetime
    try:
        return datetime.strptime(str(month), '%Y-%m').strftime('%B %Y')
    except ValueError:
        return str(month)


def _percent(value):
    return value * 100 if value is not None and abs(value) <= 1 else value


def _money_or_na(value):
    return money(value) if value is not None else 'Not available'


def english_email(name, month, row=None, previous=None, notes='', model=None):
    """Data-grounded investor email; no claims about leasing, payments or renewals without notes."""
    label = _formatted_month(month)
    notes = str(notes or '').strip()
    if model is not None:
        properties = model.get('properties') or []
        op_lines = []
        for prop in properties:
            prop_name = str(prop.get('name') or prop.get('property') or 'Property')
            occ = _percent(_numeric(prop, 'occupancy'))
            prior = _percent(_numeric(prop, 'prior_occupancy'))
            description = f'{prop_name}: occupancy was {occ:.1f}%' if occ is not None else f'{prop_name}: occupancy data is unavailable'
            if prior is not None and occ is not None:
                description += f' versus {prior:.1f}% in the previous month'
            op_lines.append(description + '.')
        operations = '\n'.join(op_lines) if op_lines else 'Operating information is provided in the attached report.'
        rental_lines = []
        for prop in properties:
            rent = _numeric(prop, 'rent', 'Rent Income')
            if rent is not None:
                rental_lines.append(f"{prop.get('name') or 'Property'}: {_money_or_na(rent)} in rental income.")
        rental = '\n'.join(rental_lines) if rental_lines else 'Please see the attached report for rental income details.'
        total_income = sum(_numeric(p, 'income') or 0 for p in properties) if properties else None
        expenses = sum(_numeric(p, 'expenses') or 0 for p in properties) if properties else None
        noi = sum(_numeric(p, 'noi') or 0 for p in properties) if properties else None
        utility = sum(_numeric(p, 'utility_income', 'Utility Income') or 0 for p in properties) if properties and all(_numeric(p, 'utility_income', 'Utility Income') is not None for p in properties) else None
        beginning = _numeric(model, 'beginning', 'beginning_cash', 'Beginning Cash')
        ending = _numeric(model, 'ending', 'ending_cash', 'Ending Cash')
        reserve = _numeric(model, 'reserve', 'required_reserve', 'Required Reserve')
        available = _numeric(model, 'available', 'Available to Distribute')
        shortfall = _numeric(model, 'shortfall', 'Reserve Shortfall')
        reserve_addition = _numeric(model, 'reserve_addition', 'Monthly Reserve Addition')
    else:
        row = row or {}
        previous = previous or {}
        occ = _percent(_numeric(row, 'Occupancy'))
        prior_occ = _percent(_numeric(previous, 'Occupancy'))
        if occ is None:
            operations = 'Occupancy information is not available for this month.'
        elif prior_occ is None:
            operations = f'The property reported {occ:.1f}% occupancy in {label}.'
        else:
            change = occ - prior_occ
            direction = 'remained consistent with' if abs(change) < 0.05 else ('increased from' if change > 0 else 'decreased from')
            operations = (f'The property maintained {occ:.1f}% occupancy in {label}, consistent with the prior month.'
                          if abs(change) < 0.05 else
                          f'The property\'s occupancy {direction} {prior_occ:.1f}% to {occ:.1f}% in {label}.')
        total_rooms = _numeric(row, 'Total Rooms')
        occupied = _numeric(row, 'Occupied Rooms')
        if total_rooms is not None and occupied is not None:
            vacant = max(0, int(round(total_rooms - occupied)))
            operations += f' {vacant} room' + (' was' if vacant == 1 else 's were') + ' vacant at month-end.'
        rent = _numeric(row, 'Rent Income')
        prior_rent = _numeric(previous, 'Rent Income')
        if rent is None:
            rental = 'Rental income information is not available for this month.'
        elif prior_rent is None:
            rental = f'Rental income for {label} was {money(rent)}.'
        else:
            rental = f'Rental income was {money(rent)} in {label}, compared with {money(prior_rent)} in the previous month.'
        ar = _numeric(row, 'Outstanding Rent Receivable')
        if ar is not None:
            rental += (' There was no outstanding rent receivable at month-end.' if abs(ar) < 0.005
                       else f' Outstanding rent receivable at month-end was {money(ar)}.')
        total_income = _numeric(row, 'Income', 'Total Income')
        utility = _numeric(row, 'Utility Income')
        expenses = _numeric(row, 'Expenses', 'Operating Expenses')
        noi = _numeric(row, 'NOI')
        reserve_addition = _numeric(row, 'Monthly Reserve Addition')
        beginning = _numeric(row, 'Beginning Cash')
        ending = _numeric(row, 'Ending Cash')
        reserve = _numeric(row, 'Required Reserve')
        available = _numeric(row, 'Available to Distribute')
        shortfall = _numeric(row, 'Reserve Shortfall')
    financial = '\n'.join([
        f'- **Total Income:** {_money_or_na(total_income)}',
        f'- **Utilities Income:** {_money_or_na(utility)}',
        f'- **Operating Expenses:** {_money_or_na(expenses)}',
        f'- **Net Operating Income (NOI):** **{_money_or_na(noi)}**',
        f'- **Reserve Contribution:** {_money_or_na(reserve_addition)}',
    ])
    cash_lines = [
        f'- **Total Balance Last Month:** {_money_or_na(beginning)}',
        f'- **Total Balance This Month:** {_money_or_na(ending)}',
        f'- **Total Required Reserve:** {_money_or_na(reserve)}',
        f'- **Available Cash for Distribution Review:** {_money_or_na(available)}',
    ]
    if shortfall is not None and shortfall > 0.005:
        cash_lines.append(f'- **Reserve Shortfall:** {money(shortfall)}')
        conclusion = ('The current cash balance is below the required reserve. '
                      'No cash is available for investor distribution review this month; '
                      'funds will remain in the property account to support reserves and ongoing operations.')
    elif available is not None and available > 0.005:
        conclusion = (f'Available cash for distribution review at month-end was {money(available)}. '
                      'Any distribution remains subject to management approval and verification of outstanding obligations.')
    elif available is not None:
        conclusion = 'No cash is available for investor distribution review this month.'
    else:
        conclusion = 'Distribution availability has not been determined from the supplied records.'
    parts = [
        'Dear Investors,',
        f'Please find below the key highlights for **{label}**:',
        '**Operations**', operations,
        '**Rental Income**', rental,
        '**Financial Performance**', financial,
        '**Cash Position & Distribution Update**', '\n'.join(cash_lines), conclusion,
    ]
    if notes:
        parts.extend(['**Additional Operating Updates**', notes])
    parts.extend([
        'We will continue to monitor rental income, operating expenses, and the property’s cash position.',
        'Please refer to the attached monthly report for detailed financial statements, reserve information, and the updated rent roll.',
        'Please feel free to reach out with any questions.',
        'Moo Housing',
    ])
    return f'{name} | Investor Report | {label}', '\n\n'.join(parts)

def show_communications(name, month, operation_cn, cash_cn, notes_cn='',
                        operation_en='', cash_en='', notes_en='', key='communication',
                        report_row=None, previous_row=None, combined_model=None):
    st.subheader('Investor Communication · 邮件草稿 / 微信通知')
    st.caption('仅生成可复制的草稿，不会自动发送。请先下载报告，并在发送前核对收件人、金额及附件。')
    language = st.radio('通知语言', ['中文', 'English'], horizontal=True, key=key+'_language')
    cn = language == '中文'
    subject, email, wechat = drafts(name, month,
                                   operation_cn if cn else operation_en,
                                   cash_cn if cn else cash_en,
                                   notes_cn if cn else notes_en, cn=cn)
    if not cn and (report_row is not None or combined_model is not None):
        subject, email = english_email(name, month, row=report_row, previous=previous_row,
                                       notes=notes_en, model=combined_model)
    email_tab, wechat_tab = st.tabs(['Email Draft · 邮件草稿', 'WeChat · 微信通知'])
    with email_tab:
        st.markdown('**Subject / 邮件标题**（右上角复制图标）')
        st.code(subject, language=None)
        st.markdown('**Email Body / 邮件正文**（右上角复制图标）')
        st.code(email, language=None)
        st.download_button('下载邮件草稿 .txt', f'Subject: {subject}\n\n{email}'.encode('utf-8'),
                           file_name=f'Moo_Email_{key}.txt', mime='text/plain', key=key+'_email_txt')
        st.info('邮件附件需手动添加本页生成的 Word 或 PDF。此功能不会直接发送邮件。')
    with wechat_tab:
        st.markdown('**微信消息**（点击右上角复制图标，粘贴到微信）')
        st.code(wechat, language=None)
        st.download_button('下载微信文案 .txt', wechat.encode('utf-8'),
                           file_name=f'Moo_WeChat_{key}.txt', mime='text/plain', key=key+'_wechat_txt')
        st.caption('微信中的报告文件需单独发送；不会自动发送到微信联系人。')
