"""Store approved investor report artifacts under a user-shared Drive folder."""
import io
import re
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.errors import HttpError

ROOT_FOLDER_ID = '178_BBqTmyIp-jFi8PljmhS2CQ_YeAp90'
SCOPES = ['https://www.googleapis.com/auth/drive']


def _service(secrets):
    info = dict(secrets['GOOGLE_APPLICATION_CREDENTIALS'])
    if 'private_key' in info:
        info['private_key'] = info['private_key'].replace('\\n', '\n')
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds, cache_discovery=False)


def _safe_name(name):
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]', '_', str(name or '').strip())
    return name[:150] or 'Unnamed LLC'


def _folder(service, parent, name):
    safe = name.replace('\\', '\\\\').replace("'", "\\'")
    query = "mimeType = 'application/vnd.google-apps.folder' and trashed = false and name = '%s' and '%s' in parents" % (safe, parent)
    result = service.files().list(q=query, fields='nextPageToken,files(id,name,webViewLink)', page_size=100, supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    if result.get('files'):
        return result['files'][0]['id']
    created = service.files().create(body={'name': name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent]}, fields='id', supportsAllDrives=True).execute()
    return created['id']


def save_reports(secrets, llc_name, year, artifacts):
    """artifacts: iterable of (filename, bytes, MIME type); update same-name files."""
    service = _service(secrets)
    try:
        root = service.files().get(fileId=ROOT_FOLDER_ID, fields='id,name,mimeType,webViewLink,capabilities(canAddChildren)', supportsAllDrives=True).execute()
    except HttpError as exc:
        if exc.resp.status == 404:
            email = dict(secrets['GOOGLE_APPLICATION_CREDENTIALS']).get('client_email', '(unknown)')
            raise PermissionError(
                f'Google Drive root folder is not visible to service account {email}. '
                f'Open https://drive.google.com/drive/folders/{ROOT_FOLDER_ID}, '
                'share this folder with that service-account email as Editor, '
                'or ask the Shared Drive administrator to grant access. '
                'A Drive API 404 can mean missing permissions, not a missing folder.'
            ) from exc
        raise
    if root.get('mimeType') != 'application/vnd.google-apps.folder':
        raise ValueError('Configured Drive root is not a folder.')
    if not root.get('capabilities', {}).get('canAddChildren', False):
        raise PermissionError('Service Account cannot create files in this Drive folder. Share it with the Service Account as Editor or use a writable Shared Drive.')
    llc_id = _folder(service, ROOT_FOLDER_ID, _safe_name(llc_name))
    year_id = _folder(service, llc_id, str(year))
    saved = []
    for filename, content, mime in artifacts:
        filename = _safe_name(filename)
        escaped = filename.replace('\\', '\\\\').replace("'", "\\'")
        existing = service.files().list(q=f"name = '{escaped}' and '{year_id}' in parents and trashed = false", fields='files(id,name)', supportsAllDrives=True, includeItemsFromAllDrives=True).execute().get('files', [])
        media = MediaIoBaseUpload(io.BytesIO(content), mimetype=mime, resumable=False)
        if existing:
            saved_file = service.files().update(fileId=existing[0]['id'], media_body=media, fields='id,name,webViewLink', supportsAllDrives=True).execute()
        else:
            saved_file = service.files().create(body={'name': filename, 'parents': [year_id]}, media_body=media, fields='id,name,webViewLink', supportsAllDrives=True).execute()
        saved.append(saved_file)
    return {'folder_url': f'https://drive.google.com/drive/folders/{year_id}', 'files': saved}
