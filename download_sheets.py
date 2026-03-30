#!/usr/bin/env python3
"""Download Google Sheets using Service Account authentication."""

import io
import json
import os
import sys

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly',
          'https://www.googleapis.com/auth/drive.readonly']

SHEETS = {
    'source': {
        'id': '1w7-6hlip3-sCT9SXqdzYm1jzpay3S185QLoqoDQUvZo',
        'output': 'data/source.xlsx',
        'name': '分析用HR_CA管理シート',
    },
    'main_sheet': {
        'id': '1YcyDzQLtjurl8iUyEXywsjiPkbxbNlLofErWCH7rMZQ',
        'output': 'data/main_sheet.xlsx',
        'name': 'HR_CA管理シート',
    },
    'targets': {
        'id': '1rJy2SvHLCN_8bAqyD5DOVrdLkOiuq9OI60rp77Uev-Q',
        'output': 'data/targets.xlsx',
        'name': '辰巳_振り返り（目標）',
    },
    'referral': {
        'id': '1dIPIOhqfOms9ab8QFEnmfEdzs9Am3IQLTTlMLN-UqgA',
        'output': 'data/referral.xlsx',
        'name': 'リファラル・アプローチ',
    },
}


def get_credentials():
    """Get credentials from service account key."""
    # Check for GOOGLE_SERVICE_ACCOUNT_KEY env var (JSON string, used in CI)
    key_json = os.environ.get('GOOGLE_SERVICE_ACCOUNT_KEY')
    if key_json:
        info = json.loads(key_json)
        return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)

    # Check for key file path
    key_file = os.environ.get('GOOGLE_KEY_FILE')
    if key_file and os.path.exists(key_file):
        return service_account.Credentials.from_service_account_file(key_file, scopes=SCOPES)

    # Default: look in common locations
    for path in [
        'service_account_key.json',
        os.path.expanduser('~/Downloads/dashboard-updater-491803-b2523c8cadc8.json'),
    ]:
        if os.path.exists(path):
            return service_account.Credentials.from_service_account_file(path, scopes=SCOPES)

    print("ERROR: No Google service account key found.")
    print("Set GOOGLE_SERVICE_ACCOUNT_KEY env var or GOOGLE_KEY_FILE path.")
    sys.exit(1)


def download_sheet(drive_service, sheet_id, output_path, name):
    """Download a Google Sheet as xlsx."""
    print(f"  Downloading {name}...")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    try:
        # Try export API first (works for most sheets)
        request = drive_service.files().export_media(
            fileId=sheet_id,
            mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        fh = io.FileIO(output_path, 'wb')
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
    except Exception as e:
        if 'exportSizeLimitExceeded' in str(e) or 'too large' in str(e):
            print(f"  Sheet too large for export API, using URL download...")
            # Fall back to public export URL with auth
            import google.auth.transport.requests
            creds = drive_service._http.credentials
            auth_req = google.auth.transport.requests.Request()
            creds.refresh(auth_req)
            headers = {'Authorization': f'Bearer {creds.token}'}
            url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx'
            import requests as req
            resp = req.get(url, headers=headers, stream=True)
            resp.raise_for_status()
            with open(output_path, 'wb') as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    f.write(chunk)
        else:
            raise

    size = os.path.getsize(output_path)
    print(f"  → {output_path} ({size:,} bytes)")


def main():
    print("Authenticating with Google...")
    creds = get_credentials()
    drive_service = build('drive', 'v3', credentials=creds)

    print("Downloading sheets...")
    for key, sheet in SHEETS.items():
        try:
            download_sheet(drive_service, sheet['id'], sheet['output'], sheet['name'])
        except Exception as e:
            print(f"  ERROR downloading {sheet['name']}: {e}")
            if 'source' in key:
                print("  This is a critical sheet. Aborting.")
                sys.exit(1)

    print("All sheets downloaded successfully.")


if __name__ == '__main__':
    main()
