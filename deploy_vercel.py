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
deploy_files = [
    "index.html",
    "data/dashboard_data.json",
]

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
