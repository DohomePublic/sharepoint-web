#!/usr/bin/env python3
"""
build_dashboard.py — KYC Dashboard builder (fixed version)
"""

import os
import json
import datetime
import requests

# ── CONFIG ───────────────────────────────────────────────────────────────────
TENANT_ID     = os.environ.get("AZURE_TENANT_ID", "").strip()
CLIENT_ID     = os.environ.get("AZURE_CLIENT_ID", "").strip()
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "").strip()

SITE_HOST  = "dohomegroup.sharepoint.com"
SITE_PATH  = "/sites/KYC"
LIST_NAME  = "KYCData1"
SP_DISP    = f"https://{SITE_HOST}{SITE_PATH}/Lists/{LIST_NAME}/DispForm.aspx?ID="


# ── VALIDATE SECRETS ──────────────────────────────────────────────────────────
def validate_env():
    missing = [k for k, v in {
        "AZURE_TENANT_ID":     TENANT_ID,
        "AZURE_CLIENT_ID":     CLIENT_ID,
        "AZURE_CLIENT_SECRET": CLIENT_SECRET,
    }.items() if not v]
    if missing:
        print(f"❌ Missing required environment variables: {', '.join(missing)}")
        print("   → Go to: Repo → Settings → Secrets and variables → Actions")
        raise SystemExit(1)
    print(f"✅ Secrets loaded — Tenant: {TENANT_ID[:8]}...")


# ── AUTH ──────────────────────────────────────────────────────────────────────
def get_token():
    print("🔑 Authenticating to Microsoft Graph...")
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    r = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default",
    }, timeout=30)
    if not r.ok:
        print(f"❌ Auth failed: HTTP {r.status_code}")
        print(f"   Response: {r.text[:500]}")
        r.raise_for_status()
    print("✅ Token acquired")
    return r.json()["access_token"]


# ── FETCH ─────────────────────────────────────────────────────────────────────
def fetch_all_items(token):
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    # ✅ แก้: ตัด $select ออก เพื่อหลีกเลี่ยง Graph bug กับ $expand=fields
    url = (
        f"https://graph.microsoft.com/v1.0/sites/"
        f"{SITE_HOST}:{SITE_PATH}:/lists/{LIST_NAME}/items"
        f"?$expand=fields&$top=999"
    )
    items = []
    page = 0
    while url:
        page += 1
        print(f"  📄 Fetching page {page}...", flush=True)
        r = requests.get(url, headers=headers, timeout=60)
        if not r.ok:
            print(f"❌ Fetch failed: HTTP {r.status_code}")
            print(f"   URL: {url}")
            print(f"   Response: {r.text[:500]}")
            r.raise_for_status()
        data = r.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        print(f"  ✅ Total so far: {len(items)} items", flush=True)
    return items


# ── TRANSFORM ─────────────────────────────────────────────────────────────────
def clean_biz(v):
    if not v:
        return ""
    return v.replace(";#", "").strip(";").strip()


def to_row(item):
    f = item.get("fields", {})
    ts = f.get("Created", "")
    try:
        dt = datetime.datetime.fromisoformat(ts.replace("Z", "+00:00"))
        created = dt.strftime("%d/%m/%Y")
        ym = dt.strftime("%Y-%m")
    except Exception:
        created = ts[:10] if ts else ""
        ym = ts[:7] if ts else ""

    # ✅ แก้: safe int conversion
    try:
        item_id = int(item["id"])
    except (ValueError, KeyError):
        item_id = 0

    return {
        "id":         item_id,
        "title":      f.get("Title") or f"Item {item['id']}",
        "registered": f.get("Registered_Name") or "",
        "status":     f.get("Status") or "",
        "type_req":   f.get("Type_Request") or "",
        "biz":        clean_biz(f.get("business_type") or ""),
        "province":   f.get("province") or "",
        "created":    created,
        "ym":         ym,
        "contact":    f.get("contact_name") or "",
        "tel":        f.get("telephone") or "",
        "owner":      f.get("Owner") or "",
        "reg_num":    f.get("registration_number") or "",
    }


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    validate_env()          # ← fail-fast ถ้า secrets ขาด
    token = get_token()

    print("📋 Fetching SharePoint list items...")
    raw = fetch_all_items(token)
    print(f"✅ Total fetched: {len(raw)} items")

    # ... (ส่วนที่เหลือเหมือนเดิม: compute_stats, build_html, write file)
