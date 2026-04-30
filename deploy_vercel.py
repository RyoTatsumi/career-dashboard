#!/usr/bin/env python3
"""Deploy static dashboard to Vercel using REST API."""

import json
import os
import hashlib
import requests

# Read Vercel auth token
auth_path = os.path.expanduser("~/Library/Application Support/com.vercel.cli/auth.json")
with open(auth_path) as f:
    token = json.load(f)["token"]

HEADERS = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
TEAM_ID = "team_uIt1kEEUEPM0Z4ZFBFGYiNtZ"
PROJECT_NAME = "hr-ca-dashboard"
BASE = "https://api.vercel.com"

# Files to deploy
# NOTE: dashboard_data.json は GitHub Actions の毎日cronが最新を生成→Vercelにdeploy するので、
# ローカルからは含めない（含めると古いローカルデータで本番を上書きしてしまう）。
# データもどうしても更新したい場合は --with-data フラグを付けて実行する。
import sys
include_data = '--with-data' in sys.argv
deploy_files = ["index.html"]
if include_data:
    deploy_files.append("data/dashboard_data.json")
    print("⚠️  --with-data 指定: ローカルの dashboard_data.json も含めます")
    # 最新のproduction dataを取得して、ローカルが古ければ警告
    try:
        import urllib.request, json as _json
        with urllib.request.urlopen("https://hr-ca-dashboard-ryo-tatsumis-projects.vercel.app/data/dashboard_data.json") as r:
            prod_data = _json.loads(r.read())
        with open("data/dashboard_data.json") as f:
            local_data = _json.load(f)
        prod_gen = prod_data.get('generated_at', '')
        local_gen = local_data.get('generated_at', '')
        if prod_gen > local_gen:
            print(f"⚠️  本番データ ({prod_gen}) のほうが新しい！")
            print(f"   ローカルデータ ({local_gen}) で上書きしますか？ Ctrl+C で中断")
            input("Enterで続行...")
    except Exception:
        pass
else:
    print("ℹ️  index.html のみデプロイします（データ変更したい場合は --with-data 付けて実行）")

# Step 1: Upload files
print("Uploading files...")
file_entries = []
for filepath in deploy_files:
    with open(filepath, "rb") as f:
        content = f.read()

    sha = hashlib.sha1(content).hexdigest()
    size = len(content)

    # Upload file
    resp = requests.post(
        f"{BASE}/v2/files",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream",
            "x-vercel-digest": sha,
            "x-vercel-size": str(size),
        },
        params={"teamId": TEAM_ID},
        data=content,
    )

    if resp.status_code in [200, 409]:  # 409 means already uploaded
        print(f"  Uploaded: {filepath} ({size} bytes)")
    else:
        print(f"  Error uploading {filepath}: {resp.status_code} {resp.text}")
        continue

    file_entries.append({
        "file": filepath,
        "sha": sha,
        "size": size,
    })

# Step 2: Create deployment
print("\nCreating deployment...")
deploy_payload = {
    "name": PROJECT_NAME,
    "files": file_entries,
    "projectSettings": {
        "framework": None,
    },
    "target": "production",
}

resp = requests.post(
    f"{BASE}/v13/deployments",
    headers=HEADERS,
    params={"teamId": TEAM_ID},
    json=deploy_payload,
)

if resp.status_code in [200, 201]:
    result = resp.json()
    url = result.get("url", "")
    print(f"\nDeployment successful!")
    print(f"URL: https://{url}")
    ready_state = result.get("readyState", "unknown")
    print(f"Status: {ready_state}")

    # Also print alias URLs
    aliases = result.get("alias", [])
    if aliases:
        print(f"Production URL: https://{aliases[0]}")
else:
    print(f"Deployment failed: {resp.status_code}")
    print(resp.text[:500])
