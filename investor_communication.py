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


def show_communications(name, month, operation_cn, cash_cn, notes_cn='',
                        operation_en='', cash_en='', notes_en='', key='communication'):
    st.subheader('Investor Communication · 邮件草稿 / 微信通知')
    st.caption('仅生成可复制的草稿，不会自动发送。请先下载报告，并在发送前核对收件人、金额及附件。')
    language = st.radio('通知语言', ['中文', 'English'], horizontal=True, key=key+'_language')
    cn = language == '中文'
    subject, email, wechat = drafts(name, month,
                                   operation_cn if cn else operation_en,
                                   cash_cn if cn else cash_en,
                                   notes_cn if cn else notes_en, cn=cn)
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
