"""Investor communication drafts; no emails or WeChat messages are sent."""
import re
import html
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


def _email_html(body):
    """Convert the report's limited Markdown to safe, selectable HTML."""
    def inline(value):
        value = html.escape(value)
        return re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", value)

    blocks = []
    in_list = False
    for line in str(body).splitlines():
        line = line.strip()
        if not line:
            if in_list:
                blocks.append('</ul>')
                in_list = False
            continue
        if line.startswith('- '):
            if not in_list:
                blocks.append('<ul>')
                in_list = True
            blocks.append('<li>' + inline(line[2:]) + '</li>')
        else:
            if in_list:
                blocks.append('</ul>')
                in_list = False
            blocks.append('<p>' + inline(line) + '</p>')
    if in_list:
        blocks.append('</ul>')
    return '<div class="email-body">' + '\n'.join(blocks) + '</div>'


def _html_document(body):
    content = _email_html(body)
    return ("""<!doctype html><html><head><meta charset="utf-8">
<style>body{font-family:Arial,Helvetica,sans-serif;color:#202124;font-size:14px;line-height:1.6;max-width:850px;margin:30px auto;padding:0 16px}
p{margin:0 0 16px}ul{margin:0 0 16px;padding-left:25px}li{margin:3px 0}strong{font-weight:700}</style>
</head><body>""" + content + '</body></html>')


def show_communications(name, month, operation_cn, cash_cn, notes_cn='',
                        operation_en='', cash_en='', notes_en='', key='communication',
                        email_override_en=None, email_override_cn=None):
    st.subheader('Investor Communication · 邮件草稿 / 微信通知')
    st.caption('仅生成可复制的草稿，不会自动发送。发送前请核对金额及附件。')
    language = st.radio('通知语言', ['中文', 'English'], horizontal=True, key=key+'_language')
    cn = language == '中文'
    subject, email, wechat = drafts(name, month,
                                   operation_cn if cn else operation_en,
                                   cash_cn if cn else cash_en,
                                   notes_cn if cn else notes_en, cn=cn)
    if cn and email_override_cn is not None:
        email = email_override_cn
    if not cn and email_override_en is not None:
        email = email_override_en
    email_tab, wechat_tab = st.tabs(['Email Draft · 邮件草稿', 'WeChat · 微信通知'])
    with email_tab:
        st.markdown('**Subject / 邮件标题**')
        st.code(subject, language=None)
        st.markdown('**Formatted Email Preview · 富文本邮件预览**')
        st.caption('预览里的加粗是真正的 HTML 格式。推荐下载 HTML，用浏览器打开后 Ctrl+A / Ctrl+C，再粘贴到 Gmail。')
        import streamlit.components.v1 as components
        preview = _html_document(email)
        # HTML rendered in an iframe: use the selection button inside the iframe,
        # or open the downloaded HTML file in a browser for reliable rich-text copying.
        select_button = """<div style="margin-bottom:14px"><button onclick="var s=window.getSelection();s.removeAllRanges();var r=document.createRange();r.selectNodeContents(document.querySelector('.email-body'));s.addRange(r);" style="padding:9px 14px;cursor:pointer">Select Email Content · 选中邮件正文</button><span style="font:12px Arial;color:#666;margin-left:10px">Then press Ctrl+C (Mac: Cmd+C)</span></div>"""
        preview = preview.replace('<body>', '<body>' + select_button)
        components.html(preview, height=650, scrolling=True)
        st.download_button('⬇️ Download Formatted Email (.html)',
                           _html_document(email).encode('utf-8'),
                           file_name=f'Moo_Email_{key}.html', mime='text/html',
                           key=key+'_email_html')
        with st.expander('Copy Plain Text · 纯文本（无星号）'):
            plain = re.sub(r'\*\*(.*?)\*\*', r'\1', email)
            st.code(plain, language=None)
            st.download_button('下载纯文本邮件 .txt',
                               f'Subject: {subject}\n\n{plain}'.encode('utf-8'),
                               file_name=f'Moo_Email_{key}.txt', mime='text/plain', key=key+'_email_txt')
        st.info('HTML 文件用 Chrome/Edge 打开后，Ctrl+A → Ctrl+C → Gmail 粘贴，即可保留加粗。邮件附件需手动添加。')
    with wechat_tab:
        st.markdown('**微信消息**')
        clean_wechat = re.sub(r'\*\*(.*?)\*\*', r'\1', wechat)
        st.code(clean_wechat, language=None)
        st.download_button('下载微信文案 .txt', clean_wechat.encode('utf-8'),
                           file_name=f'Moo_WeChat_{key}.txt', mime='text/plain', key=key+'_wechat_txt')
        st.caption('微信中的报告文件需单独发送；不会自动发送到微信联系人。')
