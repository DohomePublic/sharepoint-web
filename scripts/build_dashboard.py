#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_dashboard.py — ดึงข้อมูลจาก SharePoint Online แล้ว generate index.html (Daily Dashboard)

แหล่งข้อมูล 2 ลิสต์บนไซต์ AC-Accounting
  1) DemoApp        -> ข้อมูลคำขอสินเชื่อ/KYC ที่นำมาทำ Dashboard
  2) Admin_KycNew   -> ทะเบียนอีเมลผู้มีสิทธิ์ (ACL) คอลัมน์ Title = อีเมล

โหมดการทำงาน (เลือกอัตโนมัติ)
  GRAPH  : มี TENANT_ID/CLIENT_ID/CLIENT_SECRET  -> Microsoft Graph (app-only) ** ใช้ใน GitHub Actions **
  CSV    : ไม่มี credential แต่มีไฟล์ใน data/   -> อ่านจาก CSV (dev / offline)

การใช้งาน
  python scripts/build_dashboard.py                 # เขียนทับ index.html
  python scripts/build_dashboard.py --out out.html  # กำหนดไฟล์ผลลัพธ์
  python scripts/build_dashboard.py --source csv    # บังคับใช้ CSV
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import pathlib
import re
import sys
import time
import urllib.parse
import urllib.request
from collections import Counter

ROOT = pathlib.Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "scripts" / "template.html"
DEFAULT_OUT = ROOT / "index.html"

# ---------------------------------------------------------------- config ----
SITE_HOSTNAME = os.getenv("SITE_HOSTNAME", "dohomegroup.sharepoint.com")
SITE_PATH = os.getenv("SITE_PATH", "/sites/AC-Accounting")
LIST_DATA = os.getenv("LIST_DEMOAPP", "DemoApp")
LIST_ACL = os.getenv("LIST_ACL", "Admin_KycNew")
TENANT_ID = os.getenv("TENANT_ID", "")
CLIENT_ID = os.getenv("CLIENT_ID", "")
CLIENT_SECRET = os.getenv("CLIENT_SECRET", "")
GRAPH = "https://graph.microsoft.com/v1.0"

# ชื่อคอลัมน์ใน SharePoint -> คีย์ที่ Dashboard ใช้
FIELD_MAP = {
    "Title": "Title", "Customer_id": "Customer_id", "Customer_x0020_Name": "CustomerName",
    "branch": "branch", "Request_x0020_TimeStamp": "RequestTimeStamp", "Status": "Status",
    "Status_1": "Status_1", "Type_Request": "TypeRequest", "Type1": "Type1",
    "type_teams": "type_teams", "Typr_Distribution": "Distribution", "Owner": "Owner",
    "limit": "limit", "limit_other": "limit_other", "registration_number": "registration_number",
    "Registered_Name": "Registered_Name", "business_type": "business_type",
    "contact_name": "contact_name", "position": "position", "contact_number": "contact_number",
    "telephone": "telephone", "county": "county", "district": "district", "province": "province",
    "credit_semester1": "credit_semester1", "credit_semester2": "credit_semester2",
    "land": "land", "other_property": "other_property", "Data": "Data",
    "Estimated_annual_income": "Estimated_annual_income",
}
# ชื่อคอลัมน์ใน CSV export -> คีย์ที่ Dashboard ใช้
CSV_MAP = {
    "_ID": "ID", "Title": "Title", "Customer_id": "Customer_id", "Customer Name": "CustomerName",
    "branch": "branch", "Request TimeStamp": "RequestTimeStamp", "Status": "Status",
    "Status_1": "Status_1", "Type_Request": "TypeRequest", "Type1": "Type1",
    "type_teams": "type_teams", "Typr_Distribution": "Distribution", "Owner": "Owner",
    "limit": "limit", "limit_other": "limit_other", "registration_number": "registration_number",
    "Registered_Name": "Registered_Name", "business_type": "business_type",
    "contact_name": "contact_name", "position": "position", "contact_number": "contact_number",
    "telephone": "telephone", "county": "county", "district": "district", "province": "province",
    "credit_semester1": "credit_semester1", "credit_semester2": "credit_semester2",
    "land": "land", "other_property": "other_property", "Data": "Data",
    "Estimated_annual_income": "Estimated_annual_income",
}


def log(msg: str) -> None:
    print(f"[build] {msg}", flush=True)


# ------------------------------------------------------------- http util ----
def http_json(url: str, token: str = "", data: bytes | None = None,
              headers: dict | None = None, retries: int = 4) -> dict:
    """เรียก REST + retry แบบ exponential backoff (กัน 429/503 ของ Graph)"""
    hdr = {"Accept": "application/json"}
    if token:
        hdr["Authorization"] = "Bearer " + token
    if headers:
        hdr.update(headers)
    last = None
    for attempt in range(retries):
        req = urllib.request.Request(url, data=data, headers=hdr,
                                     method="POST" if data else "GET")
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "ignore")[:400]
            last = RuntimeError(f"HTTP {e.code} {url.split('?')[0]} :: {body}")
            if e.code in (429, 500, 502, 503, 504) and attempt < retries - 1:
                wait = int(e.headers.get("Retry-After") or 2 ** attempt)
                log(f"  ↻ HTTP {e.code} รอ {wait}s แล้วลองใหม่ ({attempt + 1}/{retries})")
                time.sleep(wait)
                continue
            raise last
        except Exception as e:                                    # network error
            last = e
            if attempt < retries - 1:
                time.sleep(2 ** attempt)
                continue
            raise
    raise last  # pragma: no cover


def get_token() -> str:
    """OAuth2 client-credentials (app-only) — ไม่ต้องมีผู้ใช้ล็อกอิน"""
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    body = urllib.parse.urlencode({
        "client_id": CLIENT_ID, "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default", "grant_type": "client_credentials",
    }).encode()
    tok = http_json(url, data=body,
                    headers={"Content-Type": "application/x-www-form-urlencoded"})
    if "access_token" not in tok:
        raise RuntimeError("ขอ access token ไม่สำเร็จ: " + json.dumps(tok)[:300])
    return tok["access_token"]


def graph_all(url: str, token: str) -> list:
    """ไล่ paging ตาม @odata.nextLink จนครบทุกรายการ"""
    out, guard = [], 0
    while url and guard < 200:
        guard += 1
        j = http_json(url, token)
        out.extend(j.get("value", []))
        url = j.get("@odata.nextLink")
    return out


# ------------------------------------------------------------ graph mode ----
def fetch_graph() -> tuple[list, list]:
    log(f"เชื่อมต่อ Microsoft Graph · site {SITE_HOSTNAME}{SITE_PATH}")
    token = get_token()
    site = http_json(f"{GRAPH}/sites/{SITE_HOSTNAME}:{SITE_PATH}", token)
    site_id = site["id"]
    log(f"  site id = {site_id}")

    def items(list_name: str) -> list:
        url = (f"{GRAPH}/sites/{site_id}/lists/{urllib.parse.quote(list_name)}"
               f"/items?expand=fields&$top=999")
        rows = graph_all(url, token)
        log(f"  {list_name}: {len(rows)} รายการ")
        return rows

    return items(LIST_DATA), items(LIST_ACL)


def map_graph_row(row: dict) -> dict:
    f = row.get("fields", {}) or {}
    o = {"ID": int(f.get("id") or row.get("id") or 0)}
    for sp, key in FIELD_MAP.items():
        v = f.get(sp)
        if v in (None, ""):
            # SharePoint บางไซต์ไม่เข้ารหัส _x0020_ -> ลองชื่อแบบ underscore
            v = f.get(sp.replace("_x0020_", "_")) or f.get(sp.replace("_x0020_", ""))
        o[key] = "" if v is None else (str(v).strip() if not isinstance(v, (int, float)) else v)
    if not o.get("RequestTimeStamp"):
        o["RequestTimeStamp"] = f.get("Created") or row.get("createdDateTime") or ""
    return o


# -------------------------------------------------------------- csv mode ----
def fetch_csv(data_csv: str, acl_csv: str) -> tuple[list, list]:
    import csv
    log(f"อ่านจาก CSV · {data_csv} / {acl_csv}")

    def read(p):
        with open(p, encoding="utf-8-sig", newline="") as fh:
            return list(csv.DictReader(fh))      # อ่านเป็น str ทั้งหมด กันเลข 0 นำหน้าหาย

    return read(data_csv), read(acl_csv)


def map_csv_row(row: dict) -> dict:
    o = {}
    for src, key in CSV_MAP.items():
        v = row.get(src, "")
        o[key] = ("" if v is None else str(v).strip())
    o["ID"] = int(float(o["ID"])) if str(o.get("ID", "")).strip() else 0
    return o


# ------------------------------------------------------------------ acl -----
def classify_acl(item_id: int, email: str) -> dict:
    """กติกาเดียวกับฝั่ง JavaScript ใน template (classifyAcl)"""
    e = (email or "").strip().lower()
    local = e.split("@")[0]
    m = re.match(r"^bi-v?operation([a-z0-9]+)_gm$", local)
    m2 = re.match(r"^dohometogogm-([a-z0-9]+)$", local)
    m3 = re.match(r"^gm-([a-z0-9]+)$", local)
    if m:
        code = m.group(1).upper()
        kind, role, name = "BI_OPERATION", "BIOPS", "BI Operation " + code
    elif m2:
        code = m2.group(1).upper()
        kind, role, name = "TOGO_GM", "GM", "Dohome To Go GM " + code
    elif m3:
        code = m3.group(1).upper()
        trainee = code == "TRAINEE"
        kind = "BRANCH_GM"
        role = "VIEWER" if trainee else "GM"
        name = "GM Trainee" if trainee else "GM สาขา " + code
    else:
        code, kind, role, name = "", "NAMED_USER", "CREDIT", local
    branches = [code + "OO"] if (code and kind != "NAMED_USER" and code != "TRAINEE") else []
    return {"id": item_id, "email": e, "kind": kind, "code": code,
            "role": role, "name": name, "branches": branches}


# ---------------------------------------------------------------- render ----
# ------------------------------------------------------------ mask PII ------
# ใช้เมื่อ deploy ขึ้นที่สาธารณะ (GitHub Pages แบบ public) — ปิดบังข้อมูลส่วนบุคคล
# ตั้งแต่ต้นทางก่อนเขียนลงไฟล์ ไม่ใช่ปิดบังแค่ตอนแสดงผล
MASK_FIELDS = ["CustomerName", "contact_name", "contact_number", "telephone",
               "registration_number", "Registered_Name", "Customer_id"]


def mask_value(v: str) -> str:
    s = str(v or "")
    if len(s) <= 2:
        return "•" * len(s)
    if len(s) <= 6:
        return s[0] + "•" * (len(s) - 1)
    return s[:3] + "•" * (len(s) - 5) + s[-2:]


def mask_records(records: list) -> list:
    for r in records:
        for f in MASK_FIELDS:
            if r.get(f):
                r[f] = mask_value(r[f])
    log(f"  ปิดบัง PII แล้ว {len(MASK_FIELDS)} คอลัมน์: {MASK_FIELDS}")
    return records


def render(records: list, acl: list, mode: str, out_path: pathlib.Path) -> None:
    tpl = TEMPLATE.read_text(encoding="utf-8")
    for ph in ("__DATA__", "__ACL__", "__GENAT__"):
        if ph not in tpl:
            raise RuntimeError(f"template.html ไม่มี placeholder {ph}")
    gen = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    html = (tpl.replace("__DATA__", json.dumps(records, ensure_ascii=False, indent=1))
               .replace("__ACL__", json.dumps(acl, ensure_ascii=False, indent=1))
               .replace("__GENAT__", f"{gen} · {mode}"))
    out_path.write_text(html, encoding="utf-8")
    log(f"เขียน {out_path} ({len(html):,} bytes)")


def summarize(records: list, acl: list) -> None:
    log(f"สรุปข้อมูล: {len(records)} รายการ · {len(acl)} บัญชีในทะเบียน")
    log(f"  ประเภทบัญชี: {dict(Counter(a['kind'] for a in acl))}")
    branches = sorted({(r.get('branch') or '').strip() for r in records if (r.get('branch') or '').strip()})
    log(f"  สาขาที่มีข้อมูล ({len(branches)}): {branches}")
    log(f"  ไม่ระบุสาขา: {sum(1 for r in records if not (r.get('branch') or '').strip())} รายการ")
    log(f"  ผู้ดูแล: {len({r.get('Owner') for r in records if r.get('Owner')})} คน")
    acl_branches = {b for a in acl for b in a["branches"]}
    missing = [b for b in branches if b not in acl_branches]
    log(f"  สาขาที่ยังไม่มีบัญชีดูแลใน {LIST_ACL}: {missing or 'ไม่มี (ครบทุกสาขา)'}")


def main() -> int:
    ap = argparse.ArgumentParser(description="Build DemoApp Daily Dashboard")
    ap.add_argument("--out", default=str(DEFAULT_OUT))
    ap.add_argument("--source", choices=["auto", "graph", "csv"], default="auto")
    ap.add_argument("--data-csv", default=str(ROOT / "data" / "DemoApp.csv"))
    ap.add_argument("--acl-csv", default=str(ROOT / "data" / "Admin_KycNew.csv"))
    ap.add_argument("--min-records", type=int, default=1,
                    help="ถ้าดึงได้น้อยกว่านี้ให้ถือว่าล้มเหลว (กันเขียนทับด้วยข้อมูลว่าง)")
    ap.add_argument("--mask-pii", action="store_true",
                    default=os.getenv("MASK_PII", "").lower() in ("1", "true", "yes"),
                    help="ปิดบังข้อมูลส่วนบุคคลก่อนเขียนไฟล์ (ใช้เมื่อ repo/Pages เป็น public)")
    a = ap.parse_args()

    use_graph = a.source == "graph" or (
        a.source == "auto" and all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]))
    try:
        if use_graph:
            raw, acl_raw = fetch_graph()
            records = [map_graph_row(r) for r in raw]
            acl = [classify_acl(int((r.get("fields", {}) or {}).get("id") or r.get("id") or 0),
                                (r.get("fields", {}) or {}).get("Title", ""))
                   for r in acl_raw]
            mode = "LIVE via Microsoft Graph"
        else:
            if a.source == "auto":
                log("ไม่พบ TENANT_ID/CLIENT_ID/CLIENT_SECRET -> ใช้โหมด CSV")
            raw, acl_raw = fetch_csv(a.data_csv, a.acl_csv)
            records = [map_csv_row(r) for r in raw]
            acl = [classify_acl(int(float(r.get("_ID") or 0)), r.get("Title", "")) for r in acl_raw]
            mode = "CSV snapshot"
    except Exception as e:                                   # noqa: BLE001
        log(f"!! ดึงข้อมูลไม่สำเร็จ: {e}")
        return 1

    acl = [x for x in acl if x["email"]]
    records = [r for r in records if r.get("ID")]
    records.sort(key=lambda x: (str(x.get("RequestTimeStamp") or ""), x["ID"]), reverse=True)

    if len(records) < a.min_records:
        log(f"!! ได้ข้อมูลเพียง {len(records)} รายการ (< {a.min_records}) — ยกเลิกเพื่อไม่ให้ทับ index.html เดิม")
        return 2

    summarize(records, acl)
    if a.mask_pii:
        records = mask_records(records)
        mode += " · PII masked"
    render(records, acl, mode, pathlib.Path(a.out))
    log("เสร็จสมบูรณ์")
    return 0


if __name__ == "__main__":
    sys.exit(main())
