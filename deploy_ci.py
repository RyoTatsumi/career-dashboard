"""Deploy to Vercel from CI (GitHub Actions)."""
import json, os, hashlib, requests

token = os.environ['VERCEL_TOKEN']
team_id = os.environ.get('VERCEL_TEAM_ID', 'team_uIt1kEEUEPM0Z4ZFBFGYiNtZ')
HEADERS = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

files_to_deploy = ["index.html", "data/dashboard_data.json"]
file_entries = []

for filepath in files_to_deploy:
    with open(filepath, "rb") as f:
        content = f.read()
    sha = hashlib.sha1(content).hexdigest()
    resp = requests.post(
        "https://api.vercel.com/v2/files",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/octet-stream",
                 "x-vercel-digest": sha, "x-vercel-size": str(len(content))},
        params={"teamId": team_id},
        data=content,
    )
    print(f"Upload {filepath}: {resp.status_code}")
    file_entries.append({"file": filepath, "sha": sha, "size": len(content)})

resp = requests.post(
    "https://api.vercel.com/v13/deployments",
    headers=HEADERS,
    params={"teamId": team_id},
    json={"name": "hr-ca-dashboard", "files": file_entries,
          "projectSettings": {"framework": None}, "target": "production"},
)
if resp.ok:
    result = resp.json()
    print(f"Deployed: https://{result.get('url', '')}")
    aliases = result.get("alias", [])
    if aliases: print(f"Production: https://{aliases[0]}")
else:
    print(f"Deploy failed: {resp.status_code} {resp.text[:300]}")
    exit(1)
