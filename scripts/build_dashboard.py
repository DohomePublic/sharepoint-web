"""
build_dashboard.py

ดึงข้อมูล KYCData1 จาก SharePoint ผ่าน Microsoft Graph API
แล้วสร้าง index.html สำหรับ GitHub Pages

SharePoint:
https://dohomegroup.sharepoint.com/sites/KYC/Lists/KYCData1/AllItems.aspx

Authentication:
Microsoft Entra ID - Client Credentials
"""

import os
import json
import requests

from collections import Counter
from datetime import datetime, timezone, timedelta


# ============================================================
# CONFIG
# ============================================================

CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]

SITE_HOST = "dohomegroup.sharepoint.com"
SITE_PATH = "/sites/KYC"

# SharePoint List
LIST_NAME = "KYCData1"

# Fields จาก SharePoint
SELECT_FIELDS = ",".join([
    "Title",
    "registration_number",
    "Status",
    "Type_Request",
    "business_type",
    "province",
    "district",
    "Estimated_annual_income",
    "limit",
    "contact_name",
    "telephone",
    "contact_number",
    "product_group",
    "Created"
])

# Power Apps
POWERAPPS_URL = (
    "https://apps.powerapps.com/play/e/"
    "default-7f8918d9-718a-495b-ac9a-17cba381c4a0/"
    "a/25b58c8a-551c-494d-acbb-ea4ed16fe8cd"
    "?tenantId=7f8918d9-718a-495b-ac9a-17cba381c4a0"
)


# ============================================================
# HTTP HELPER
# ============================================================

def graph_get(url, token, params=None):
    """
    GET Microsoft Graph พร้อมตรวจ Error
    """

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    response = requests.get(
        url,
        headers=headers,
        params=params,
        timeout=60
    )

    if not response.ok:
        print("")
        print("========== GRAPH API ERROR ==========")
        print(f"HTTP Status : {response.status_code}")
        print(f"URL         : {response.url}")
        print(f"Response    : {response.text[:5000]}")
        print("=====================================")
        print("")

        response.raise_for_status()

    return response.json()


# ============================================================
# AUTH
# ============================================================

def get_token():

    print("Getting Microsoft Graph access token...")

    url = (
        f"https://login.microsoftonline.com/"
        f"{TENANT_ID}/oauth2/v2.0/token"
    )

    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }

    response = requests.post(
        url,
        data=data,
        timeout=60
    )

    if not response.ok:

        print("")
        print("========== TOKEN ERROR ==========")
        print(f"HTTP Status : {response.status_code}")
        print(f"Response    : {response.text[:5000]}")
        print("=================================")
        print("")

        response.raise_for_status()

    token = response.json().get("access_token")

    if not token:
        raise RuntimeError(
            "ไม่พบ access_token จาก Microsoft Entra ID"
        )

    print("Graph token: OK")

    return token


# ============================================================
# GET SHAREPOINT SITE
# ============================================================

def get_site(token):

    print("")
    print("Getting SharePoint Site...")

    url = (
        f"https://graph.microsoft.com/v1.0/"
        f"sites/{SITE_HOST}:{SITE_PATH}"
    )

    data = graph_get(url, token)

    site_id = data.get("id")

    if not site_id:
        raise RuntimeError(
            "ไม่พบ Site ID จาก Microsoft Graph"
        )

    print(f"Site ID: {site_id}")

    return site_id


# ============================================================
# GET SHAREPOINT LIST
# ============================================================

def get_list(token, site_id):

    print("")
    print("Getting SharePoint Lists...")

    url = (
        f"https://graph.microsoft.com/v1.0/"
        f"sites/{site_id}/lists"
    )

    params = {
        "$select": "id,name,displayName,webUrl"
    }

    data = graph_get(
        url,
        token,
        params=params
    )

    lists = data.get("value", [])

    print(f"Found {len(lists)} SharePoint lists")

    # แสดง List ทั้งหมดใน Log
    for lst in lists:

        print(
            "LIST:"
            f" displayName={lst.get('displayName')}"
            f" | name={lst.get('name')}"
            f" | id={lst.get('id')}"
        )

    # ========================================================
    # หา KYCData1
    # ========================================================

    target = None

    for lst in lists:

        display_name = str(
            lst.get("displayName", "")
        ).strip().lower()

        internal_name = str(
            lst.get("name", "")
        ).strip().lower()

        if (
            display_name == LIST_NAME.lower()
            or internal_name == LIST_NAME.lower()
        ):
            target = lst
            break

    if not target:

        print("")
        print("==========================================")
        print(f"ERROR: ไม่พบ SharePoint List '{LIST_NAME}'")
        print("")
        print("Lists ที่พบ:")

        for lst in lists:
            print(
                f" - {lst.get('displayName')}"
                f" | {lst.get('name')}"
            )

        print("==========================================")
        print("")

        raise RuntimeError(
            f"ไม่พบ SharePoint List: {LIST_NAME}"
        )

    list_id = target.get("id")

    print("")
    print("========== TARGET LIST ==========")
    print(f"Display Name : {target.get('displayName')}")
    print(f"Internal Name: {target.get('name')}")
    print(f"List ID      : {list_id}")
    print(f"Web URL      : {target.get('webUrl')}")
    print("=================================")
    print("")

    return list_id


# ============================================================
# FETCH ALL LIST ITEMS
# ============================================================

def fetch_all_items(token):

    # --------------------------------------------------------
    # 1. Get Site ID
    # --------------------------------------------------------

    site_id = get_site(token)

    # --------------------------------------------------------
    # 2. Get List ID
    # --------------------------------------------------------

    list_id = get_list(
        token,
        site_id
    )

    # --------------------------------------------------------
    # 3. Get Items
    # --------------------------------------------------------

    print("Fetching SharePoint items...")

    url = (
        f"https://graph.microsoft.com/v1.0/"
        f"sites/{site_id}"
        f"/lists/{list_id}"
        f"/items"
    )

    params = {
        "$expand": f"fields($select={SELECT_FIELDS})",
        "$top": "500"
    }

    items = []
    page = 0

    while url:

        page += 1

        data = graph_get(
            url,
            token,
            params=params
        )

        batch = data.get("value", [])

        items.extend(batch)

        print(
            f"Page {page}: "
            f"{len(batch):,} items "
            f"(total {len(items):,})"
        )

        # nextLink จะมี query string มาให้แล้ว
        url = data.get("@odata.nextLink")

        # หลังหน้าแรก nextLink มี parameters ครบแล้ว
        params = None

    print("")
    print(
        f"Total fetched: {len(items):,} items"
    )

    return items


# ============================================================
# CLEAN DATA
# ============================================================

def clean_biz(value):

    if not value:
        return ""

    return (
        str(value)
        .replace(";#", " ")
        .strip()
    )


def clean_income(value):

    if not value:
        return "ไม่ระบุ"

    result = (
        str(value)
        .replace(";#", "")
        .strip()
    )

    return result or "ไม่ระบุ"


# ============================================================
# TRANSFORM
# ============================================================

def transform(items):

    rows = []

    for item in items:

        fields = item.get("fields", {})

        created = fields.get(
            "Created",
            ""
        ) or ""

        date_str = (
            created[:10]
            if len(created) >= 10
            else ""
        )

        year_ce = 0
        month_ce = 0

        if date_str:

            try:

                dt = datetime.strptime(
                    date_str,
                    "%Y-%m-%d"
                )

                year_ce = dt.year
                month_ce = dt.month

            except Exception:
                pass

        year_be = (
            year_ce + 543
            if year_ce > 0
            else 0
        )

        rows.append([
            str(item.get("id", "")),

            fields.get("Title", "") or "",

            fields.get(
                "registration_number",
                ""
            ) or "",

            fields.get(
                "Status",
                ""
            ) or "ไม่ระบุ",

            fields.get(
                "Type_Request",
                ""
            ) or "ไม่ระบุ",

            clean_biz(
                fields.get(
                    "business_type",
                    ""
                )
            ),

            fields.get(
                "province",
                ""
            ) or "ไม่ระบุ",

            fields.get(
                "district",
                ""
            ) or "",

            clean_income(
                fields.get(
                    "Estimated_annual_income",
                    ""
                )
            ),

            fields.get(
                "limit",
                ""
            ) or "",

            date_str,

            year_be,

            month_ce,

            fields.get(
                "contact_name",
                ""
            ) or "",

            fields.get(
                "telephone",
                ""
            ) or "",

            fields.get(
                "contact_number",
                ""
            ) or "",

            fields.get(
                "product_group",
                ""
            ) or ""
        ])

    return rows


# ============================================================
# HEADERS
# ============================================================

HEADERS_LIST = [
    "_ID",
    "Title",
    "registration_number",
    "Status",
    "Type_Request",
    "biz_clean",
    "province",
    "district",
    "Estimated_annual_income",
    "limit",
    "date_str",
    "year_BE",
    "month_CE",
    "contact_name",
    "telephone",
    "contact_number",
    "product_group"
]


# ============================================================
# AGGREGATE
# ============================================================

def aggregate(rows):

    idx = {
        header: i
        for i, header in enumerate(
            HEADERS_LIST
        )
    }

    status_cnt = Counter()
    type_cnt = Counter()
    prov_cnt = Counter()
    inc_cnt = Counter()
    biz_cnt = Counter()
    monthly = Counter()

    for row in rows:

        status = row[idx["Status"]]

        type_request = row[
            idx["Type_Request"]
        ]

        province = row[
            idx["province"]
        ]

        income = row[
            idx["Estimated_annual_income"]
        ]

        business = row[
            idx["biz_clean"]
        ]

        year_be = row[
            idx["year_BE"]
        ]

        month_ce = row[
            idx["month_CE"]
        ]

        status_cnt[status] += 1

        type_cnt[type_request] += 1

        if (
            province
            and province != "ไม่ระบุ"
        ):
            prov_cnt[province] += 1

        if (
            income
            and income != "ไม่ระบุ"
        ):
            inc_cnt[income] += 1

        # แยกประเภทธุรกิจ
        for part in [
            p.strip()
            for p in business.split()
            if p.strip()
        ]:

            biz_cnt[part] += 1

        if (
            year_be > 0
            and month_ce > 0
        ):

            monthly[
                (year_be, month_ce)
            ] += 1

    years = sorted(
        {
            row[idx["year_BE"]]
            for row in rows
            if row[idx["year_BE"]] > 543
        },
        reverse=True
    )

    return {
        "status": dict(status_cnt),

        "type_req": dict(type_cnt),

        "biz": dict(
            biz_cnt.most_common(10)
        ),

        "province": dict(
            prov_cnt.most_common(15)
        ),

        "income": dict(inc_cnt),

        "monthly": [
            {
                "year_BE": key[0],
                "month_CE": key[1],
                "cnt": value
            }
            for key, value
            in monthly.items()
        ],

        "years": years,

        "total": len(rows)
    }


# ============================================================
# BUILD HTML
# ============================================================

def build_html(
    rows,
    charts,
    update_date
):

    data_json = json.dumps(
        {
            "headers": HEADERS_LIST,
            "rows": rows
        },
        ensure_ascii=False
    )

    charts_json = json.dumps(
        charts,
        ensure_ascii=False
    )

    # ========================================================
    # HTML
    # ========================================================

    html = f"""<!DOCTYPE html>
<html lang="th">

<head>

<meta charset="UTF-8">

<meta
    name="viewport"
    content="width=device-width,initial-scale=1"
>

<title>KYC Dashboard | DoHome</title>

<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>

<style>

* {{
    box-sizing:border-box;
    margin:0;
    padding:0;
    font-family:'Segoe UI',Tahoma,sans-serif
}}

body {{
    background:#eef2f7;
    color:#1a202c
}}

:root {{
    --blue:#1565c0;
    --blue-l:#e3f0ff;
    --green:#2e7d32;
    --red:#c62828;
    --orange:#e65100;
    --gray:#546e7a;
    --card:#fff;
    --border:#dde3ea
}}

nav {{
    background:var(--blue);
    color:#fff;
    padding:0 24px;
    display:flex;
    align-items:center;
    justify-content:space-between;
    height:56px;
    box-shadow:0 2px 8px rgba(0,0,0,.2)
}}

.logo {{
    font-size:1.15rem;
    font-weight:800;
    letter-spacing:1px
}}

.nav-right {{
    display:flex;
    align-items:center;
    gap:12px
}}

.btn-home {{
    background:#fff;
    color:var(--blue);
    border:none;
    padding:6px 16px;
    border-radius:20px;
    font-weight:700;
    cursor:pointer;
    font-size:.85rem;
    text-decoration:none
}}

.upd {{
    font-size:.78rem;
    opacity:.85
}}

.container {{
    max-width:1440px;
    margin:0 auto;
    padding:16px 20px
}}

.stat-grid {{
    display:grid;
    grid-template-columns:
        repeat(auto-fit,minmax(160px,1fr));
    gap:14px;
    margin-bottom:18px
}}

.sc {{
    background:var(--card);
    border-radius:14px;
    padding:18px 16px;
    box-shadow:0 2px 8px rgba(0,0,0,.07);
    border-left:5px solid var(--blue);
    display:flex;
    flex-direction:column;
    gap:4px
}}

.sc.g {{ border-color:var(--green) }}
.sc.r {{ border-color:var(--red) }}
.sc.o {{ border-color:var(--orange) }}

.sl {{
    font-size:.73rem;
    color:var(--gray);
    font-weight:700
}}

.sv {{
    font-size:2rem;
    font-weight:800;
    line-height:1.1
}}

.sp {{
    font-size:.78rem;
    color:var(--gray)
}}

.filter-bar {{
    background:var(--card);
    border-radius:14px;
    padding:14px 18px;
    box-shadow:0 2px 8px rgba(0,0,0,.07);
    margin-bottom:18px;
    display:flex;
    flex-wrap:wrap;
    gap:12px;
    align-items:flex-end
}}

.fg {{
    display:flex;
    flex-direction:column;
    gap:4px;
    min-width:155px
}}

.fg label {{
    font-size:.72rem;
    font-weight:700;
    color:var(--gray)
}}

.fg select,
.fg input {{
    padding:8px 10px;
    border:1.5px solid var(--border);
    border-radius:8px;
    font-size:.88rem;
    background:#f7faff;
    outline:none
}}

.btn-s {{
    background:var(--blue);
    color:#fff;
    border:none;
    padding:9px 22px;
    border-radius:8px;
    font-weight:700;
    cursor:pointer
}}

.btn-c {{
    background:#f0f4f8;
    color:var(--gray);
    border:1.5px solid var(--border);
    padding:9px 18px;
    border-radius:8px;
    font-weight:600;
    cursor:pointer
}}

.year-tabs {{
    display:flex;
    gap:8px;
    flex-wrap:wrap;
    margin-bottom:16px
}}

.yt {{
    background:#fff;
    border:2px solid var(--border);
    padding:7px 18px;
    border-radius:20px;
    font-weight:700;
    cursor:pointer;
    color:var(--gray)
}}

.yt.active {{
    background:var(--blue);
    color:#fff;
    border-color:var(--blue)
}}

.chart-grid {{
    display:grid;
    grid-template-columns:
        repeat(auto-fit,minmax(420px,1fr));
    gap:16px;
    margin-bottom:18px
}}

.cc {{
    background:var(--card);
    border-radius:14px;
    padding:18px;
    box-shadow:0 2px 8px rgba(0,0,0,.07)
}}

.ct {{
    font-weight:700;
    font-size:.95rem;
    margin-bottom:12px
}}

.cw {{
    position:relative;
    height:240px
}}

.table-card {{
    background:var(--card);
    border-radius:14px;
    padding:18px;
    box-shadow:0 2px 8px rgba(0,0,0,.07);
    margin-bottom:20px
}}

.th-row {{
    display:flex;
    justify-content:space-between;
    align-items:center;
    margin-bottom:12px;
    flex-wrap:wrap;
    gap:10px
}}

.tt {{
    font-weight:700;
    font-size:.95rem
}}

.cb {{
    background:var(--blue-l);
    color:var(--blue);
    padding:4px 12px;
    border-radius:20px;
    font-weight:700;
    font-size:.82rem
}}

.si {{
    padding:8px 14px;
    border:1.5px solid var(--border);
    border-radius:8px;
    font-size:.88rem;
    width:260px;
    outline:none
}}

.tw {{
    overflow-x:auto;
    max-height:500px;
    overflow-y:auto
}}

table {{
    width:100%;
    border-collapse:collapse;
    font-size:.82rem
}}

thead th {{
    background:#f0f4f8;
    padding:10px 12px;
    text-align:left;
    position:sticky;
    top:0;
    font-weight:700;
    color:#455a64;
    border-bottom:2px solid var(--border);
    white-space:nowrap;
    z-index:2
}}

tbody tr {{
    border-bottom:1px solid #f0f4f8
}}

tbody tr:hover {{
    background:#f0f7ff
}}

tbody td {{
    padding:9px 12px;
    color:#374151;
    max-width:200px;
    overflow:hidden;
    text-overflow:ellipsis;
    white-space:nowrap
}}

.badge {{
    display:inline-block;
    padding:2px 10px;
    border-radius:12px;
    font-size:.75rem;
    font-weight:700
}}

.ba {{
    background:#e8f5e9;
    color:#2e7d32
}}

.br {{
    background:#ffebee;
    color:#c62828
}}

.bw {{
    background:#fff8e1;
    color:#f57f17
}}

.bo {{
    background:#f3e5f5;
    color:#6a1b9a
}}

.pg {{
    display:flex;
    gap:6px;
    align-items:center;
    justify-content:flex-end;
    margin-top:14px;
    flex-wrap:wrap
}}

.pb {{
    background:#fff;
    border:1.5px solid var(--border);
    padding:6px 12px;
    border-radius:8px;
    font-size:.82rem;
    cursor:pointer;
    font-weight:600
}}

.pb:hover,
.pb.active {{
    background:var(--blue);
    color:#fff;
    border-color:var(--blue)
}}

.pb:disabled {{
    opacity:.4;
    cursor:default
}}

.pi {{
    font-size:.82rem;
    color:var(--gray)
}}

.mo {{
    position:fixed;
    inset:0;
    background:rgba(0,0,0,.45);
    z-index:1000;
    display:none;
    align-items:center;
    justify-content:center;
    padding:20px
}}

.mo.open {{
    display:flex
}}

.md {{
    background:#fff;
    border-radius:16px;
    max-width:640px;
    width:100%;
    max-height:88vh;
    overflow-y:auto;
    padding:24px
}}

.mt {{
    font-weight:800;
    font-size:1.05rem;
    margin-bottom:16px;
    color:var(--blue);
    display:flex;
    justify-content:space-between
}}

.mc {{
    cursor:pointer;
    font-size:1.4rem;
    color:var(--gray)
}}

.dg {{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:12px
}}

.di {{
    display:flex;
    flex-direction:column;
    gap:3px
}}

.dl {{
    font-size:.72rem;
    color:var(--gray);
    font-weight:700
}}

.dv {{
    font-size:.9rem;
    color:#1a202c;
    font-weight:500;
    word-break:break-word
}}

@media(max-width:640px) {{
    .chart-grid {{
        grid-template-columns:1fr
    }}

    .dg {{
        grid-template-columns:1fr
    }}
}}

</style>

</head>

<body>

<nav>

<div class="logo">
📊 KYC Dashboard | DoHome
</div>

<div class="nav-right">

<span class="upd">
🗓 ข้อมูล ณ {update_date}
</span>

<a
    href="{POWERAPPS_URL}"
    target="_blank"
    class="btn-home"
>
🏠 กลับ KYC App
</a>

</div>

</nav>


<div class="container">

<div
    class="stat-grid"
    id="statGrid"
></div>


<div
    class="year-tabs"
    id="yearTabs"
></div>


<div class="filter-bar">

<div class="fg">

<label>
🔍 ชื่อกิจการ / เลขทะเบียน
</label>

<input
    type="text"
    id="fText"
    placeholder="พิมพ์ค้นหา..."
>

</div>


<div class="fg">

<label>
📋 สถานะ
</label>

<select id="fStatus">
<option value="">ทั้งหมด</option>
</select>

</div>


<div class="fg">

<label>
📝 ประเภทคำขอ
</label>

<select id="fType">
<option value="">ทั้งหมด</option>
</select>

</div>


<div class="fg">

<label>
🗺 จังหวัด
</label>

<select id="fProv">
<option value="">ทั้งหมด</option>
</select>

</div>


<div class="fg">

<label>
📅 วันที่เริ่ม
</label>

<input
    type="date"
    id="fDateFrom"
>

</div>


<div class="fg">

<label>
📅 วันที่สิ้นสุด
</label>

<input
    type="date"
    id="fDateTo"
>

</div>


<button
    class="btn-s"
    onclick="applyFilter()"
>
🔍 ค้นหา
</button>


<button
    class="btn-c"
    onclick="clearFilter()"
>
✕ ล้าง
</button>

</div>


<div class="chart-grid">

<div class="cc">
<div class="ct">
📊 สถานะคำขอ
</div>

<div class="cw">
<canvas id="cStatus"></canvas>
</div>
</div>


<div class="cc">
<div class="ct">
📝 ประเภทคำขอ
</div>

<div class="cw">
<canvas id="cType"></canvas>
</div>
</div>


<div class="cc">
<div class="ct">
🏢 ประเภทธุรกิจ Top 10
</div>

<div class="cw">
<canvas id="cBiz"></canvas>
</div>
</div>


<div class="cc">
<div class="ct">
🗺 จังหวัด Top 15
</div>

<div class="cw">
<canvas id="cProv"></canvas>
</div>
</div>


<div class="cc">
<div class="ct">
💰 รายได้ประจำปีโดยประมาณ
</div>

<div class="cw">
<canvas id="cInc"></canvas>
</div>
</div>


<div class="cc">
<div class="ct">
📈 แนวโน้มรายเดือน
</div>

<div class="cw">
<canvas id="cMonth"></canvas>
</div>
</div>

</div>


<div class="table-card">

<div class="th-row">

<div class="tt">
📋 รายการ KYC
</div>

<div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">

<span
    class="cb"
    id="rowCount"
>
0 รายการ
</span>

<input
    class="si"
    id="tSearch"
    placeholder="🔍 ค้นหาในตาราง..."
    oninput="renderTable()"
>

</div>

</div>


<div class="tw">

<table>

<thead>

<tr>

<th>#</th>
<th>ชื่อกิจการ</th>
<th>เลขทะเบียน</th>
<th>สถานะ</th>
<th>ประเภทคำขอ</th>
<th>จังหวัด</th>
<th>รายได้/ปี</th>
<th>วงเงิน</th>
<th>วันที่สร้าง</th>
<th></th>

</tr>

</thead>

<tbody id="tBody"></tbody>

</table>

</div>


<div
    class="pg"
    id="pg"
></div>

</div>

</div>


<div
    class="mo"
    id="modal"
    onclick="closeModal(event)"
>

<div class="md">

<div class="mt">

<span id="mTitle">
รายละเอียด
</span>

<span
    class="mc"
    onclick="closeModal()"
>
✕
</span>

</div>

<div
    class="dg"
    id="mBody"
></div>

</div>

</div>


<script>

const CD = {charts_json};
const RD = {data_json};

const H = RD.headers;
const RA = RD.rows;

const I = {{}};

H.forEach(
    (h,i) => I[h] = i
);

let FR = [...RA];

let AY = 0;

let CP = 1;

const PS = 50;

let CI = {{}};


window.onload = function() {{

    initYears();

    populateFilters();

    applyFilter();

}};


function initYears(){{

    const el =
        document.getElementById(
            "yearTabs"
        );

    let html =
        '<button class="yt active" '
        + 'data-y="0" '
        + 'onclick="setYear(0,this)">'
        + '📅 ทุกปี'
        + '</button>';

    CD.years.forEach(y => {{

        html +=
            '<button class="yt" '
            + 'data-y="' + y + '" '
            + 'onclick="setYear(' + y + ',this)">'
            + y
            + '</button>';

    }});

    el.innerHTML = html;
}}


function setYear(y,el){{

    AY = y;

    document
        .querySelectorAll(".yt")
        .forEach(
            t => t.classList.toggle(
                "active",
                t.dataset.y === String(y)
            )
        );

    applyFilter();
}}


function populateFilters(){{

    const statuses =
        [...new Set(
            RA.map(
                r => r[I.Status]
            )
        )]
        .filter(Boolean)
        .sort();

    const types =
        [...new Set(
            RA.map(
                r => r[I.Type_Request]
            )
        )]
        .filter(Boolean)
        .sort();

    const provinces =
        [...new Set(
            RA.map(
                r => r[I.province]
            )
        )]
        .filter(
            p => p && p !== "ไม่ระบุ"
        )
        .sort();

    const statusEl =
        document.getElementById(
            "fStatus"
        );

    statuses.forEach(s => {{

        statusEl.innerHTML +=
            '<option value="' +
            escapeHtml(s) +
            '">' +
            escapeHtml(s) +
            '</option>';

    }});


    const typeEl =
        document.getElementById(
            "fType"
        );

    types.forEach(t => {{

        typeEl.innerHTML +=
            '<option value="' +
            escapeHtml(t) +
            '">' +
            escapeHtml(t) +
            '</option>';

    }});


    const provinceEl =
        document.getElementById(
            "fProv"
        );

    provinces.forEach(p => {{

        provinceEl.innerHTML +=
            '<option value="' +
            escapeHtml(p) +
            '">' +
            escapeHtml(p) +
            '</option>';

    }});
}}


function escapeHtml(value){{

    return String(value ?? "")
        .replace(/&/g,"&amp;")
        .replace(/</g,"&lt;")
        .replace(/>/g,"&gt;")
        .replace(/"/g,"&quot;")
        .replace(/'/g,"&#039;");
}}


function applyFilter(){{

    const txt =
        document.getElementById(
            "fText"
        ).value
        .toLowerCase();

    const st =
        document.getElementById(
            "fStatus"
        ).value;

    const tp =
        document.getElementById(
            "fType"
        ).value;

    const pv =
        document.getElementById(
            "fProv"
        ).value;

    const df =
        document.getElementById(
            "fDateFrom"
        ).value;

    const dt =
        document.getElementById(
            "fDateTo"
        ).value;


    FR = RA.filter(r => {{

        if (
            AY > 0 &&
            r[I.year_BE] !== AY
        )
            return false;


        if (
            st &&
            r[I.Status] !== st
        )
            return false;


        if (
            tp &&
            r[I.Type_Request] !== tp
        )
            return false;


        if (
            pv &&
            r[I.province] !== pv
        )
            return false;


        if (
            txt &&
            !(
                String(
                    r[I.Title] || ""
                )
                .toLowerCase()
                .includes(txt)

                ||

                String(
                    r[I.registration_number] || ""
                )
                .includes(txt)

                ||

                String(
                    r[I.contact_name] || ""
                )
                .toLowerCase()
                .includes(txt)
            )
        )
            return false;


        if (
            df &&
            r[I.date_str] < df
        )
            return false;


        if (
            dt &&
            r[I.date_str] > dt
        )
            return false;


        return true;

    }});


    CP = 1;

    renderStats();

    renderCharts();

    renderTable();
}}


function clearFilter(){{

    [
        "fText",
        "fStatus",
        "fType",
        "fProv",
        "fDateFrom",
        "fDateTo"
    ]
    .forEach(
        id =>
            document.getElementById(
                id
            ).value = ""
    );

    AY = 0;

    document
        .querySelectorAll(".yt")
        .forEach(
            t =>
                t.classList.toggle(
                    "active",
                    t.dataset.y === "0"
                )
        );

    applyFilter();
}}


function renderStats(){{

    const total = FR.length;

    const approve =
        FR.filter(
            r =>
                r[I.Status] === "Approve"
        ).length;

    const reject =
        FR.filter(
            r =>
                String(
                    r[I.Status] || ""
                ).includes("Reject")
                ||
                String(
                    r[I.Status] || ""
                ).includes("ไม่อนุมัติ")
        ).length;

    const waiting =
        FR.filter(
            r =>
                String(
                    r[I.Status] || ""
                ).includes("รอ")
                ||
                String(
                    r[I.Status] || ""
                ).includes("ผ่านการพิจารณา")
        ).length;

    const newCredit =
        FR.filter(
            r =>
                r[I.Type_Request] ===
                "คำขอเปิดวงเงินลูกค้าใหม่"
        ).length;


    const pct = (n,t) =>
        t
        ? ((n/t)*100).toFixed(1) + "%"
        : "0%";


    document.getElementById(
        "statGrid"
    ).innerHTML = `

    <div class="sc">

        <div class="sl">
            📋 ทั้งหมด
        </div>

        <div class="sv">
            ${{total.toLocaleString()}}
        </div>

        <div class="sp">
            รายการ
        </div>

    </div>


    <div class="sc g">

        <div class="sl">
            ✅ Approve
        </div>

        <div
            class="sv"
            style="color:#2e7d32"
        >
            ${{approve.toLocaleString()}}
        </div>

        <div class="sp">
            ${{pct(approve,total)}}
        </div>

    </div>


    <div class="sc r">

        <div class="sl">
            ❌ Reject
        </div>

        <div
            class="sv"
            style="color:#c62828"
        >
            ${{reject.toLocaleString()}}
        </div>

        <div class="sp">
            ${{pct(reject,total)}}
        </div>

    </div>


    <div class="sc o">

        <div class="sl">
            ⏳ รอพิจารณา
        </div>

        <div
            class="sv"
            style="color:#e65100"
        >
            ${{waiting.toLocaleString()}}
        </div>

        <div class="sp">
            ${{pct(waiting,total)}}
        </div>

    </div>


    <div class="sc">

        <div class="sl">
            🆕 เปิดวงเงินใหม่
        </div>

        <div class="sv">
            ${{newCredit.toLocaleString()}}
        </div>

        <div class="sp">
            จาก ${{total.toLocaleString()}} รายการ
        </div>

    </div>

    `;
}}


const PAL = [
    "#1976d2",
    "#43a047",
    "#e53935",
    "#fb8c00",
    "#8e24aa",
    "#00acc1",
    "#f4511e",
    "#6d4c41",
    "#546e7a",
    "#39796b",
    "#c0ca33",
    "#fdd835",
    "#ff7043",
    "#78909c",
    "#66bb6a"
];


function mkChart(
    id,
    type,
    labels,
    data
){{

    if (CI[id])
        CI[id].destroy();


    const canvas =
        document.getElementById(id);

    if (!canvas)
        return;


    CI[id] = new Chart(
        canvas,
        {{

            type: type,

            data: {{

                labels: labels,

                datasets: [{{

                    data: data,

                    backgroundColor:
                        PAL.slice(
                            0,
                            labels.length
                        ),

                    borderWidth:
                        type === "line"
                        ? 2
                        : 0,

                    fill:
                        type === "line",

                    tension:.4,

                    pointRadius:4

                }}]

            }},

            options: {{

                responsive:true,

                maintainAspectRatio:false,

                plugins: {{

                    legend: {{

                        display:
                            type === "doughnut"
                            ||
                            type === "pie",

                        position:"right"

                    }}

                }},

                scales:
                    type === "bar"
                    ||
                    type === "line"
                    ? {{

                        y: {{
                            beginAtZero:true
                        }}

                    }}
                    : {{}}

            }}

        }}
    );
}}


function renderCharts(){{

    const statusCount = {{}};
    const typeCount = {{}};
    const provinceCount = {{}};
    const incomeCount = {{}};
    const monthCount = {{}};


    FR.forEach(r => {{

        const status =
            r[I.Status];

        const type =
            r[I.Type_Request];

        const province =
            r[I.province];

        const income =
            r[I.Estimated_annual_income];


        statusCount[status] =
            (statusCount[status] || 0) + 1;


        typeCount[type] =
            (typeCount[type] || 0) + 1;


        if (
            province &&
            province !== "ไม่ระบุ"
        ) {{

            provinceCount[province] =
                (provinceCount[province] || 0) + 1;

        }}


        if (
            income &&
            income !== "ไม่ระบุ"
        ) {{

            incomeCount[income] =
                (incomeCount[income] || 0) + 1;

        }}


        const month =
            r[I.month_CE];

        const year =
            r[I.year_BE];


        if (month && year) {{

            const key =
                year + "-" + month;

            monthCount[key] =
                (monthCount[key] || 0) + 1;

        }}

    }});


    mkChart(
        "cStatus",
        "doughnut",
        Object.keys(statusCount),
        Object.values(statusCount)
    );


    mkChart(
        "cType",
        "doughnut",
        Object.keys(typeCount),
        Object.values(typeCount)
    );


    const provinces =
        Object.entries(
            provinceCount
        )
        .sort(
            (a,b) => b[1] - a[1]
        )
        .slice(0,15);


    mkChart(
        "cProv",
        "bar",
        provinces.map(x => x[0]),
        provinces.map(x => x[1])
    );


    const incomes = [
        "น้อยกว่า 5 ล้านบาท/ปี",
        "5 ล้านแต่ไม่เกิน 10 ล้านบาท/ปี",
        "10 ล้านแต่ไม่เกิน 50 ล้านบาท/ปี",
        "50 ล้านแต่ไม่เกิน 100 ล้านบาท/ปี",
        "100 ล้านแต่ไม่เกิน 500 ล้านบาท/ปี",
        "500 ล้านบาท/ปี ขึ้นไป"
    ];


    const incomeLabels =
        incomes.filter(
            x => incomeCount[x]
        );


    mkChart(
        "cInc",
        "bar",
        incomeLabels,
        incomeLabels.map(
            x => incomeCount[x] || 0
        )
    );


    const months = [
        "ม.ค.",
        "ก.พ.",
        "มี.ค.",
        "เม.ย.",
        "พ.ค.",
        "มิ.ย.",
        "ก.ค.",
        "ส.ค.",
        "ก.ย.",
        "ต.ค.",
        "พ.ย.",
        "ธ.ค."
    ];


    const monthTotals = {{}};


    Object.entries(
        monthCount
    ).forEach(
        ([key,value]) => {{

            const parts =
                key.split("-");

            const month =
                Number(parts[1]);

            monthTotals[month] =
                (monthTotals[month] || 0)
                + value;

        }}
    );


    mkChart(
        "cMonth",
        "line",
        months,
        Array.from(
            {{length:12}},
            (_,i) =>
                monthTotals[i+1] || 0
        )
    );


    const businessCount = CD.biz || {{}};

    mkChart(
        "cBiz",
        "bar",
        Object.keys(businessCount),
        Object.values(businessCount)
    );
}}


function renderTable(){{

    const query =
        (
            document.getElementById(
                "tSearch"
            ).value || ""
        )
        .toLowerCase();


    let rows = FR;


    if (query) {{

        rows =
            rows.filter(
                r =>
                    String(
                        r[I.Title] || ""
                    )
                    .toLowerCase()
                    .includes(query)

                    ||

                    String(
                        r[I.registration_number] || ""
                    )
                    .includes(query)

                    ||

                    String(
                        r[I.province] || ""
                    )
                    .toLowerCase()
                    .includes(query)

                    ||

                    String(
                        r[I.contact_name] || ""
                    )
                    .toLowerCase()
                    .includes(query)
            );

    }}


    document.getElementById(
        "rowCount"
    ).textContent =
        rows.length.toLocaleString()
        + " รายการ";


    const totalPages =
        Math.ceil(
            rows.length / PS
        ) || 1;


    if (CP > totalPages)
        CP = 1;


    const start =
        (CP - 1) * PS;


    const pageRows =
        rows.slice(
            start,
            start + PS
        );


    document.getElementById(
        "tBody"
    ).innerHTML =
        pageRows
        .map(
            (r,i) => {{

                const status =
                    r[I.Status] || "";

                let cls = "bo";


                if (
                    status === "Approve"
                )
                    cls = "ba";

                else if (
                    status.includes("Reject")
                    ||
                    status.includes("ไม่อนุมัติ")
                )
                    cls = "br";

                else if (
                    status.includes("รอ")
                    ||
                    status.includes("ผ่าน")
                )
                    cls = "bw";


                return `

<tr>

<td>
${{start+i+1}}
</td>

<td
    title="${{escapeHtml(r[I.Title])}}"
    style="font-weight:600"
>
${{escapeHtml(r[I.Title] || "-")}}
</td>

<td>
${{escapeHtml(r[I.registration_number] || "-")}}
</td>

<td>

<span
    class="badge ${{cls}}"
>
${{escapeHtml(r[I.Status] || "-")}}
</span>

</td>

<td>
${{escapeHtml(r[I.Type_Request] || "-")}}
</td>

<td>
${{escapeHtml(r[I.province] || "-")}}
</td>

<td>
${{escapeHtml(r[I.Estimated_annual_income] || "-")}}
</td>

<td
    style="font-weight:700;color:#1565c0"
>
${{escapeHtml(r[I.limit] || "-")}}
</td>

<td>
${{escapeHtml(r[I.date_str] || "-")}}
</td>

<td>

<button
    class="pb"
    onclick='showDetail(${{JSON.stringify(r)}})'
>
ดู
</button>

</td>

</tr>

`;

            }
        )
        .join("");


    renderPg(
        totalPages,
        rows
    );
}}


function renderPg(
    total,
    rows
){{

    const el =
        document.getElementById(
            "pg"
        );


    if (total <= 1) {{

        el.innerHTML = "";

        return;

    }}


    let html =
        `<span class="pi">
        หน้า ${{CP}}/${{total}}
        (${{rows.length.toLocaleString()}} รายการ)
        </span>`;


    html +=
        `<button
            class="pb"
            onclick="goPg(1)"
            ${{CP === 1 ? "disabled" : ""}}
        >«</button>`;


    html +=
        `<button
            class="pb"
            onclick="goPg(${{CP-1}})"
            ${{CP === 1 ? "disabled" : ""}}
        >‹</button>`;


    const start =
        Math.max(
            1,
            CP - 2
        );


    const end =
        Math.min(
            total,
            CP + 2
        );


    for (
        let i = start;
        i <= end;
        i++
    ) {{

        html +=
            `<button
                class="pb ${{i === CP ? "active" : ""}}"
                onclick="goPg(${{i}})"
            >${{i}}</button>`;

    }}


    html +=
        `<button
            class="pb"
            onclick="goPg(${{CP+1}})"
            ${{CP === total ? "disabled" : ""}}
        >›</button>`;


    html +=
        `<button
            class="pb"
            onclick="goPg(${{total}})"
            ${{CP === total ? "disabled" : ""}}
        >»</button>`;


    el.innerHTML = html;
}}


function goPg(page){{

    CP = page;

    renderTable();

    document
        .querySelector(".table-card")
        .scrollIntoView({{
            behavior:"smooth"
        }});
}}


function showDetail(row){{

    const fields = [

        ["Title","ชื่อกิจการ"],

        [
            "registration_number",
            "เลขทะเบียน"
        ],

        ["Status","สถานะ"],

        [
            "Type_Request",
            "ประเภทคำขอ"
        ],

        [
            "biz_clean",
            "ประเภทธุรกิจ"
        ],

        [
            "province",
            "จังหวัด"
        ],

        [
            "district",
            "อำเภอ"
        ],

        [
            "contact_name",
            "ชื่อผู้ติดต่อ"
        ],

        [
            "telephone",
            "โทรศัพท์"
        ],

        [
            "contact_number",
            "เบอร์ผู้ติดต่อ"
        ],

        [
            "Estimated_annual_income",
            "รายได้ประจำปี"
        ],

        [
            "limit",
            "วงเงิน"
        ],

        [
            "product_group",
            "กลุ่มสินค้า"
        ],

        [
            "date_str",
            "วันที่สร้าง"
        ]

    ];


    document.getElementById(
        "mTitle"
    ).textContent =
        row[I.Title]
        || "รายละเอียด";


    document.getElementById(
        "mBody"
    ).innerHTML =
        fields
        .map(
            ([key,label]) => `

<div class="di">

<div class="dl">
${{label}}
</div>

<div class="dv">
${{escapeHtml(row[I[key]] || "-")}}
</div>

</div>

`
        )
        .join("");


    document.getElementById(
        "modal"
    ).classList.add("open");
}}


function closeModal(event){{

    const modal =
        document.getElementById(
            "modal"
        );


    if (
        !event
        ||
        event.target === modal
    )
        modal.classList.remove(
            "open"
        );
}}

</script>

</body>

</html>
"""

    return html


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":

    # --------------------------------------------------------
    # Thailand timezone
    # --------------------------------------------------------

    tz_thai = timezone(
        timedelta(hours=7)
    )

    now_thai = datetime.now(
        tz_thai
    )

    thai_months = [
        "ม.ค.",
        "ก.พ.",
        "มี.ค.",
        "เม.ย.",
        "พ.ค.",
        "มิ.ย.",
        "ก.ค.",
        "ส.ค.",
        "ก.ย.",
        "ต.ค.",
        "พ.ย.",
        "ธ.ค."
    ]

    update_date = (
        f"{now_thai.day} "
        f"{thai_months[now_thai.month - 1]} "
        f"{now_thai.year + 543}"
    )


    print("")
    print("========================================")
    print("=== KYC Dashboard Builder ===")
    print("========================================")

    print(
        f"Update date: {update_date}"
    )

    print(
        f"SharePoint: "
        f"https://{SITE_HOST}{SITE_PATH}"
    )

    print(
        f"List: {LIST_NAME}"
    )

    print("")


    # --------------------------------------------------------
    # Authentication
    # --------------------------------------------------------

    token = get_token()


    # --------------------------------------------------------
    # Fetch SharePoint
    # --------------------------------------------------------

    items = fetch_all_items(
        token
    )


    # --------------------------------------------------------
    # Transform
    # --------------------------------------------------------

    rows = transform(
        items
    )


    print(
        f"Transformed: {len(rows):,} rows"
    )


    # --------------------------------------------------------
    # Aggregate
    # --------------------------------------------------------

    charts = aggregate(
        rows
    )


    # --------------------------------------------------------
    # Build HTML
    # --------------------------------------------------------

    html = build_html(
        rows,
        charts,
        update_date
    )


    # --------------------------------------------------------
    # Write index.html
    # --------------------------------------------------------

    output_path = "index.html"

    with open(
        output_path,
        "w",
        encoding="utf-8"
    ) as file:

        file.write(html)


    print("")
    print("========================================")
    print("BUILD SUCCESS")
    print("========================================")

    print(
        f"Written: {output_path}"
    )

    print(
        f"HTML size: {len(html):,} chars"
    )

    print(
        f"Rows: {len(rows):,}"
    )

    print("========================================")
