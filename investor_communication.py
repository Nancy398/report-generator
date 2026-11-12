"""Investor communication drafts; no emails or WeChat messages are sent."""
import re
import html
import json
import streamlit as st
import streamlit.components.v1 as components


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


def plain_text(value):
    """Remove Markdown emphasis markers from copyable email/WeChat drafts."""
    value = str(value or '')
    value = value.replace('\\*', '*').replace('\\-', '-')
    value = re.sub(r'(?<!\w)\*{1,3}(?=\S)|(?<=\S)\*{1,3}(?!\w)', '', value)
    return value


def formatted_email_html(value):
    """Safe, Gmail-friendly HTML from the report's plain/Markdown-like draft."""
    source = str(value or '').replace('\\*', '*').replace('\\-', '-')
    blocks = []
    headings = {'Operations', 'Rental Income', 'Financial Performance',
                'Cash Position & Distribution Update', 'Additional Operating Updates'}
    for block in re.split(r'\n\s*\n', source.strip()):
        lines = [line.strip() for line in block.splitlines() if line.strip()]
        if not lines:
            continue
        clean = lambda x: re.sub(r'\*{1,3}', '', x).strip()
        if len(lines) == 1 and clean(lines[0]) in headings:
            blocks.append(f'<p style="margin:18px 0 8px"><strong>{html.escape(clean(lines[0]))}</strong></p>')
        elif all(re.match(r'^[-*]\s+', line) for line in lines):
            items = []
            for line in lines:
                text = clean(re.sub(r'^[-*]\s+', '', line))
                if ':' in text:
                    title, rest = text.split(':', 1)
                    content = f'<strong>{html.escape(title)}:</strong>{html.escape(rest)}'
                else:
                    content = html.escape(text)
                items.append(f'<li style="margin:3px 0">{content}</li>')
            blocks.append('<ul style="padding-left:24px;margin:8px 0 16px">'+''.join(items)+'</ul>')
        else:
            content = '<br>'.join(html.escape(clean(line)) for line in lines)
            blocks.append(f'<p style="margin:0 0 13px">{content}</p>')
    return '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.55;color:#222">'+''.join(blocks)+'</div>'


def formatted_copy_widget(body, plain):
    """Rich clipboard on user click; selectable rich-text fallback for blocked clipboard."""
    rich = formatted_email_html(body)
    # Escape closing script in JSON, even if future user notes contain markup.
    rich_js = json.dumps(rich).replace('<', '\\u003c')
    plain_js = json.dumps(plain).replace('<', '\\u003c')
    page = f"""<!doctype html><html><head><meta charset='utf-8'></head><body style='margin:0;font-family:Arial,sans-serif'>
    <button id='copy' style='padding:9px 15px;background:#165dba;color:white;border:0;border-radius:6px;cursor:pointer'>Copy Formatted Email</button>
    <span id='status' role='status' style='margin-left:10px;font-size:12px'></span>
    <div id='mail' contenteditable='true' aria-label='Formatted email; select all to copy' style='margin-top:12px;padding:15px;border:1px solid #ddd;border-radius:6px;min-height:140px'></div>
    <script>
    const rich={rich_js};const plain={plain_js};
    const mail=document.getElementById('mail');mail.innerHTML=rich;
    document.getElementById('copy').addEventListener('click',async()=>{{
      const status=document.getElementById('status');
      try {{
        if (!navigator.clipboard || !window.ClipboardItem) throw Error('Clipboard API unavailable');
        await navigator.clipboard.write([new ClipboardItem({{
          'text/html':new Blob([rich],{{type:'text/html'}}),
          'text/plain':new Blob([plain],{{type:'text/plain'}})
        }})]);
        status.textContent='Copied with formatting. Paste into Gmail.';
      }} catch(err) {{
        const range=document.createRange();range.selectNodeContents(mail);
        const selection=window.getSelection();selection.removeAllRanges();selection.addRange(range);
        status.textContent='Clipboard blocked. Press Ctrl+C (Mac: Cmd+C), then paste into Gmail.';
      }}
    }});
    </script></body></html>"""
    components.html(page, height=430, scrolling=True)


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
    email = plain_text(email)
    wechat = plain_text(wechat)
    email_tab, wechat_tab = st.tabs(['Email Draft · 邮件草稿', 'WeChat · 微信通知'])
    with email_tab:
        st.markdown('**Subject / 邮件标题**（右上角复制图标）')
        st.code(subject, language=None)
        st.markdown('**Email Body / 邮件正文**（右上角复制图标）')
        formatted_copy_widget(email, email)
        st.markdown('**Copy Plain Text**（右上角复制图标）')
        st.code(email, language=None)
        st.download_button('下载带格式邮件 .html', ('<!doctype html><html><meta charset=\"utf-8\"><body>'+formatted_email_html(email)+'</body></html>').encode('utf-8'), file_name=f'Moo_Email_{key}.html', mime='text/html', key=key+'_email_html')
        st.download_button('下载邮件草稿 .txt', f'Subject: {subject}\n\n{email}'.encode('utf-8'),
                           file_name=f'Moo_Email_{key}.txt', mime='text/plain', key=key+'_email_txt')
        st.info('邮件附件需手动添加本页生成的 Word 或 PDF。此功能不会直接发送邮件。')
    with wechat_tab:
        st.markdown('**微信消息**（点击右上角复制图标，粘贴到微信）')
        st.code(wechat, language=None)
        st.download_button('下载微信文案 .txt', wechat.encode('utf-8'),
                           file_name=f'Moo_WeChat_{key}.txt', mime='text/plain', key=key+'_wechat_txt')
        st.caption('微信中的报告文件需单独发送；不会自动发送到微信联系人。')
