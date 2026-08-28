"""
build_dashboard.py
ดึงข้อมูล KYCData1 จาก SharePoint ผ่าน Microsoft Graph API
แล้วสร้าง index.html ฝังข้อมูลทั้งหมด
"""
import os, json, requests
from collections import Counter
from datetime import datetime, timezone, timedelta

# ─── Config ───────────────────────────────────────────────
CLIENT_ID     = os.environ['AZURE_CLIENT_ID']
TENANT_ID     = os.environ['AZURE_TENANT_ID']
CLIENT_SECRET = os.environ['AZURE_CLIENT_SECRET']
SITE_HOST     = 'dohomegroup.sharepoint.com'
SITE_PATH     = '/sites/KYC'
LIST_NAME     = 'KYCData1'
SELECT_FIELDS = ','.join([
    'id','Title','registration_number','Status','Type_Request',
    'business_type','province','district','Estimated_annual_income',
    'limit','contact_name','telephone','contact_number',
    'product_group','Created'
])
POWERAPPS_URL = (
    'https://apps.powerapps.com/play/e/'
    'default-7f8918d9-718a-495b-ac9a-17cba381c4a0/'
    'a/25b58c8a-551c-494d-acbb-ea4ed16fe8cd'
    '?tenantId=7f8918d9-718a-495b-ac9a-17cba381c4a0'
)

# ─── Auth: Client Credentials Flow ────────────────────────
def get_token():
    url = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
    r = requests.post(url, data={
        'grant_type': 'client_credentials',
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'scope': 'https://graph.microsoft.com/.default'
    })
    r.raise_for_status()
    return r.json()['access_token']

# ─── Fetch all list items (paginate) ──────────────────────
def fetch_all_items(token):
    headers = {'Authorization': f'Bearer {token}', 'Accept': 'application/json'}

    # 1. Get site id
    r = requests.get(
        f'https://graph.microsoft.com/v1.0/sites/{SITE_HOST}:{SITE_PATH}',
        headers=headers
    )
    r.raise_for_status()
    site_id = r.json()['id']
    print(f'Site ID: {site_id}')

    # 2. Get list id
    r = requests.get(
        f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{LIST_NAME}',
        headers=headers
    )
    r.raise_for_status()
    list_id = r.json()['id']
    print(f'List ID: {list_id}')

    # 3. Paginate items
    url = (
        f'https://graph.microsoft.com/v1.0/sites/{site_id}'
        f'/lists/{list_id}/items'
        f'?expand=fields($select={SELECT_FIELDS})&$top=500'
    )
    items = []
    page = 0
    while url:
        page += 1
        r = requests.get(url, headers=headers)
        r.raise_for_status()
        data = r.json()
        batch = data.get('value', [])
        items.extend(batch)
        url = data.get('@odata.nextLink')
        print(f'  Page {page}: {len(batch)} items (total {len(items)})')

    print(f'Total fetched: {len(items)} items')
    return items

# ─── Transform ────────────────────────────────────────────
def clean_biz(v):
    if not v: return ''
    return str(v).replace(';#', ' ').strip()

def clean_income(v):
    if not v: return 'ไม่ระบุ'
    return str(v).replace(';#', '').strip() or 'ไม่ระบุ'

def transform(items):
    rows = []
    for item in items:
        f = item.get('fields', {})
        created = f.get('Created', '') or ''
        date_str = created[:10] if len(created) >= 10 else ''
        year_ce = 0
        month_ce = 0
        if date_str:
            try:
                dt = datetime.strptime(date_str, '%Y-%m-%d')
                year_ce = dt.year
                month_ce = dt.month
            except Exception:
                pass
        year_be = year_ce + 543 if year_ce > 0 else 0

        rows.append([
            str(item.get('id', '')),
            f.get('Title', '') or '',
            f.get('registration_number', '') or '',
            f.get('Status', '') or 'ไม่ระบุ',
            f.get('Type_Request', '') or 'ไม่ระบุ',
            clean_biz(f.get('business_type', '')),
            f.get('province', '') or 'ไม่ระบุ',
            f.get('district', '') or '',
            clean_income(f.get('Estimated_annual_income', '')),
            f.get('limit', '') or '',
            date_str,
            year_be,
            month_ce,
            f.get('contact_name', '') or '',
            f.get('telephone', '') or '',
            f.get('contact_number', '') or '',
            f.get('product_group', '') or '',
        ])
    return rows

HEADERS_LIST = [
    '_ID','Title','registration_number','Status','Type_Request','biz_clean',
    'province','district','Estimated_annual_income','limit','date_str',
    'year_BE','month_CE','contact_name','telephone','contact_number','product_group'
]

# ─── Aggregate ────────────────────────────────────────────
def aggregate(rows):
    idx = {h: i for i, h in enumerate(HEADERS_LIST)}
    status_cnt, type_cnt, prov_cnt, inc_cnt = Counter(), Counter(), Counter(), Counter()
    biz_cnt = Counter()
    monthly = Counter()

    for r in rows:
        st = r[idx['Status']]
        tp = r[idx['Type_Request']]
        pv = r[idx['province']]
        inc = r[idx['Estimated_annual_income']]
        biz = r[idx['biz_clean']]
        yb = r[idx['year_BE']]
        m  = r[idx['month_CE']]

        status_cnt[st] += 1
        type_cnt[tp] += 1
        if pv and pv != 'ไม่ระบุ': prov_cnt[pv] += 1
        if inc and inc != 'ไม่ระบุ': inc_cnt[inc] += 1
        for part in [p.strip() for p in biz.split() if p.strip()]:
            biz_cnt[part] += 1
        if yb > 0 and m > 0:
            monthly[(yb, m)] += 1

    years = sorted({r[idx['year_BE']] for r in rows if r[idx['year_BE']] > 543}, reverse=True)

    return {
        'status':  dict(status_cnt),
        'type_req': dict(type_cnt),
        'biz':     dict(biz_cnt.most_common(10)),
        'province': dict(prov_cnt.most_common(15)),
        'income':  dict(inc_cnt),
        'monthly': [{'year_BE': k[0], 'month_CE': k[1], 'cnt': v} for k, v in monthly.items()],
        'years':   years,
        'total':   len(rows)
    }

# ─── Build HTML ───────────────────────────────────────────
def build_html(rows, charts, update_date):
    data_json   = json.dumps({'headers': HEADERS_LIST, 'rows': rows}, ensure_ascii=False)
    charts_json = json.dumps(charts, ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>KYC Dashboard | DoHome</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0;font-family:'Segoe UI',Tahoma,sans-serif}}
body{{background:#eef2f7;color:#1a202c}}
:root{{--blue:#1565c0;--blue-l:#e3f0ff;--green:#2e7d32;--red:#c62828;--orange:#e65100;--gray:#546e7a;--card:#fff;--border:#dde3ea}}
nav{{background:var(--blue);color:#fff;padding:0 24px;display:flex;align-items:center;justify-content:space-between;height:56px;box-shadow:0 2px 8px rgba(0,0,0,.2)}}
.logo{{font-size:1.15rem;font-weight:800;letter-spacing:1px}}
.nav-right{{display:flex;align-items:center;gap:12px}}
.btn-home{{background:#fff;color:var(--blue);border:none;padding:6px 16px;border-radius:20px;font-weight:700;cursor:pointer;font-size:.85rem;text-decoration:none;transition:.2s}}
.btn-home:hover{{background:#e3f0ff}}
.upd{{font-size:.78rem;opacity:.85}}
.container{{max-width:1440px;margin:0 auto;padding:16px 20px}}
.stat-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:14px;margin-bottom:18px}}
.sc{{background:var(--card);border-radius:14px;padding:18px 16px;box-shadow:0 2px 8px rgba(0,0,0,.07);border-left:5px solid var(--blue);display:flex;flex-direction:column;gap:4px}}
.sc.g{{border-color:var(--green)}}.sc.r{{border-color:var(--red)}}.sc.o{{border-color:var(--orange)}}
.sl{{font-size:.73rem;color:var(--gray);font-weight:700;text-transform:uppercase;letter-spacing:.5px}}
.sv{{font-size:2rem;font-weight:800;line-height:1.1}}
.sp{{font-size:.78rem;color:var(--gray)}}
.filter-bar{{background:var(--card);border-radius:14px;padding:14px 18px;box-shadow:0 2px 8px rgba(0,0,0,.07);margin-bottom:18px;display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end}}
.fg{{display:flex;flex-direction:column;gap:4px;min-width:155px}}
.fg label{{font-size:.72rem;font-weight:700;color:var(--gray);text-transform:uppercase}}
.fg select,.fg input{{padding:8px 10px;border:1.5px solid var(--border);border-radius:8px;font-size:.88rem;background:#f7faff;outline:none;transition:.15s}}
.fg select:focus,.fg input:focus{{border-color:var(--blue);background:#fff}}
.btn-s{{background:var(--blue);color:#fff;border:none;padding:9px 22px;border-radius:8px;font-weight:700;font-size:.9rem;cursor:pointer;align-self:flex-end;transition:.2s}}
.btn-s:hover{{background:#0d47a1}}
.btn-c{{background:#f0f4f8;color:var(--gray);border:1.5px solid var(--border);padding:9px 18px;border-radius:8px;font-weight:600;font-size:.88rem;cursor:pointer;align-self:flex-end;transition:.2s}}
.btn-c:hover{{background:#e0e7ef}}
.year-tabs{{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:16px}}
.yt{{background:#fff;border:2px solid var(--border);padding:7px 18px;border-radius:20px;font-weight:700;font-size:.88rem;cursor:pointer;transition:.2s;color:var(--gray)}}
.yt.active{{background:var(--blue);color:#fff;border-color:var(--blue)}}
.yt:hover:not(.active){{border-color:var(--blue);color:var(--blue)}}
.chart-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:16px;margin-bottom:18px}}
.cc{{background:var(--card);border-radius:14px;padding:18px;box-shadow:0 2px 8px rgba(0,0,0,.07)}}
.ct{{font-weight:700;font-size:.95rem;margin-bottom:12px}}
.cw{{position:relative;height:240px}}
.table-card{{background:var(--card);border-radius:14px;padding:18px;box-shadow:0 2px 8px rgba(0,0,0,.07);margin-bottom:20px}}
.th-row{{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;flex-wrap:wrap;gap:10px}}
.tt{{font-weight:700;font-size:.95rem}}
.cb{{background:var(--blue-l);color:var(--blue);padding:4px 12px;border-radius:20px;font-weight:700;font-size:.82rem}}
.si{{padding:8px 14px;border:1.5px solid var(--border);border-radius:8px;font-size:.88rem;width:260px;outline:none}}
.si:focus{{border-color:var(--blue)}}
.tw{{overflow-x:auto;max-height:500px;overflow-y:auto}}
table{{width:100%;border-collapse:collapse;font-size:.82rem}}
thead th{{background:#f0f4f8;padding:10px 12px;text-align:left;position:sticky;top:0;font-weight:700;color:#455a64;border-bottom:2px solid var(--border);white-space:nowrap;z-index:2}}
tbody tr{{border-bottom:1px solid #f0f4f8;transition:.1s}}
tbody tr:hover{{background:#f0f7ff}}
tbody td{{padding:9px 12px;color:#374151;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.badge{{display:inline-block;padding:2px 10px;border-radius:12px;font-size:.75rem;font-weight:700;white-space:nowrap}}
.ba{{background:#e8f5e9;color:#2e7d32}}.br{{background:#ffebee;color:#c62828}}
.bw{{background:#fff8e1;color:#f57f17}}.bo{{background:#f3e5f5;color:#6a1b9a}}
.pg{{display:flex;gap:6px;align-items:center;justify-content:flex-end;margin-top:14px;flex-wrap:wrap}}
.pb{{background:#fff;border:1.5px solid var(--border);padding:6px 12px;border-radius:8px;font-size:.82rem;cursor:pointer;transition:.15s;font-weight:600}}
.pb:hover,.pb.active{{background:var(--blue);color:#fff;border-color:var(--blue)}}
.pb:disabled{{opacity:.4;cursor:default}}
.pi{{font-size:.82rem;color:var(--gray)}}
.mo{{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:1000;display:none;align-items:center;justify-content:center;padding:20px}}
.mo.open{{display:flex}}
.md{{background:#fff;border-radius:16px;max-width:640px;width:100%;max-height:88vh;overflow-y:auto;padding:24px;box-shadow:0 8px 40px rgba(0,0,0,.2)}}
.mt{{font-weight:800;font-size:1.05rem;margin-bottom:16px;color:var(--blue);display:flex;justify-content:space-between;align-items:center}}
.mc{{cursor:pointer;font-size:1.4rem;color:var(--gray);line-height:1}}
.dg{{display:grid;grid-template-columns:1fr 1fr;gap:12px}}
.di{{display:flex;flex-direction:column;gap:3px}}
.dl{{font-size:.72rem;color:var(--gray);font-weight:700;text-transform:uppercase}}
.dv{{font-size:.9rem;color:#1a202c;font-weight:500;word-break:break-word}}
@media(max-width:640px){{.chart-grid{{grid-template-columns:1fr}}.dg{{grid-template-columns:1fr}}}}
</style>
</head>
<body>
<nav>
  <div class="logo">📊 KYC Dashboard | DoHome</div>
  <div class="nav-right">
    <span class="upd">🗓 ข้อมูล ณ {update_date}</span>
    <a href="{POWERAPPS_URL}" target="_blank" class="btn-home">🏠 กลับ KYC App</a>
  </div>
</nav>
<div class="container">
  <div class="stat-grid" id="statGrid"></div>
  <div class="year-tabs" id="yearTabs"></div>
  <div class="filter-bar">
    <div class="fg"><label>🔍 ชื่อกิจการ / เลขทะเบียน</label><input type="text" id="fText" placeholder="พิมพ์ค้นหา..."></div>
    <div class="fg"><label>📋 สถานะ</label><select id="fStatus"><option value="">ทั้งหมด</option></select></div>
    <div class="fg"><label>📝 ประเภทคำขอ</label><select id="fType"><option value="">ทั้งหมด</option></select></div>
    <div class="fg"><label>🗺 จังหวัด</label><select id="fProv"><option value="">ทั้งหมด</option></select></div>
    <div class="fg"><label>📅 วันที่เริ่ม</label><input type="date" id="fDateFrom"></div>
    <div class="fg"><label>📅 วันที่สิ้นสุด</label><input type="date" id="fDateTo"></div>
    <button class="btn-s" onclick="applyFilter()">🔍 ค้นหา</button>
    <button class="btn-c" onclick="clearFilter()">✕ ล้าง</button>
  </div>
  <div class="chart-grid">
    <div class="cc"><div class="ct">📊 สถานะคำขอ <small style="color:#888;font-weight:400">(คลิกเลือก)</small></div><div class="cw"><canvas id="cStatus"></canvas></div></div>
    <div class="cc"><div class="ct">📝 ประเภทคำขอ <small style="color:#888;font-weight:400">(คลิกเลือก)</small></div><div class="cw"><canvas id="cType"></canvas></div></div>
    <div class="cc"><div class="ct">🏢 ประเภทธุรกิจ Top 10 <small style="color:#888;font-weight:400">(คลิกเลือก)</small></div><div class="cw"><canvas id="cBiz"></canvas></div></div>
    <div class="cc"><div class="ct">🗺 จังหวัด Top 15 <small style="color:#888;font-weight:400">(คลิกเลือก)</small></div><div class="cw"><canvas id="cProv"></canvas></div></div>
    <div class="cc"><div class="ct">💰 รายได้ประจำปีโดยประมาณ</div><div class="cw"><canvas id="cInc"></canvas></div></div>
    <div class="cc"><div class="ct">📈 แนวโน้มรายเดือน</div><div class="cw"><canvas id="cMonth"></canvas></div></div>
  </div>
  <div class="table-card">
    <div class="th-row">
      <div class="tt">📋 รายการ KYC</div>
      <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
        <span class="cb" id="rowCount">0 รายการ</span>
        <input class="si" id="tSearch" placeholder="🔍 ค้นหาในตาราง..." oninput="renderTable()">
      </div>
    </div>
    <div class="tw">
      <table>
        <thead><tr><th>#</th><th>ชื่อกิจการ</th><th>เลขทะเบียน</th><th>สถานะ</th><th>ประเภทคำขอ</th><th>จังหวัด</th><th>รายได้/ปี</th><th>วงเงิน</th><th>วันที่สร้าง</th><th></th></tr></thead>
        <tbody id="tBody"></tbody>
      </table>
    </div>
    <div class="pg" id="pg"></div>
  </div>
</div>
<div class="mo" id="modal" onclick="closeModal(event)">
  <div class="md">
    <div class="mt"><span id="mTitle">รายละเอียด</span><span class="mc" onclick="closeModal()">✕</span></div>
    <div class="dg" id="mBody"></div>
  </div>
</div>
<script>
const CD={charts_json};
const RD={data_json};
const H=RD.headers,RA=RD.rows;
const I={{}};H.forEach((h,i)=>I[h]=i);
let FR=[...RA],AY=0,CP=1;const PS=50;let CI={{}};
window.onload=function(){{initYears();populateFilters();applyFilter();}};
function initYears(){{
  const el=document.getElementById('yearTabs');
  let h='<button class="yt active" data-y="0" onclick="setYear(0,this)">📅 ทุกปี</button>';
  CD.years.forEach(y=>{{h+=`<button class="yt" data-y="${{y}}" onclick="setYear(${{y}},this)">${{y}}</button>"`;}});
  el.innerHTML=h;
}}
function setYear(y,el){{AY=y;document.querySelectorAll('.yt').forEach(t=>t.classList.toggle('active',t.dataset.y==String(y)));applyFilter();}}
function populateFilters(){{
  const ss=[...new Set(RA.map(r=>r[I.Status]))].filter(Boolean).sort();
  const ts=[...new Set(RA.map(r=>r[I.Type_Request]))].filter(Boolean).sort();
  const ps=[...new Set(RA.map(r=>r[I.province]))].filter(p=>p&&p!='ไม่ระบุ').sort();
  const fS=document.getElementById('fStatus');ss.forEach(s=>{{fS.innerHTML+=`<option value="${{s}}">${{s}}</option>`;}});
  const fT=document.getElementById('fType');ts.forEach(t=>{{fT.innerHTML+=`<option value="${{t}}">${{t}}</option>`;}});
  const fP=document.getElementById('fProv');ps.forEach(p=>{{fP.innerHTML+=`<option value="${{p}}">${{p}}</option>`;}});
}}
function applyFilter(){{
  const txt=document.getElementById('fText').value.toLowerCase();
  const st=document.getElementById('fStatus').value;
  const tp=document.getElementById('fType').value;
  const pv=document.getElementById('fProv').value;
  const df=document.getElementById('fDateFrom').value;
  const dt=document.getElementById('fDateTo').value;
  FR=RA.filter(r=>{{
    if(AY>0&&r[I.year_BE]!==AY)return false;
    if(st&&r[I.Status]!==st)return false;
    if(tp&&r[I.Type_Request]!==tp)return false;
    if(pv&&r[I.province]!==pv)return false;
    if(txt&&!((r[I.Title]||'').toLowerCase().includes(txt)||(r[I.registration_number]||'').includes(txt)||(r[I.contact_name]||'').toLowerCase().includes(txt)))return false;
    if(df&&r[I.date_str]<df)return false;
    if(dt&&r[I.date_str]>dt)return false;
    return true;
  }});
  CP=1;renderStats();renderCharts();renderTable();
}}
function clearFilter(){{
  ['fText','fStatus','fType','fProv','fDateFrom','fDateTo'].forEach(id=>{{document.getElementById(id).value='';}});
  AY=0;document.querySelectorAll('.yt').forEach(t=>t.classList.toggle('active',t.dataset.y==='0'));
  applyFilter();
}}
function renderStats(){{
  const total=FR.length;
  const ap=FR.filter(r=>r[I.Status]==='Approve').length;
  const rj=FR.filter(r=>r[I.Status].includes('Reject')||r[I.Status].includes('ไม่อนุมัติ')).length;
  const wt=FR.filter(r=>r[I.Status].includes('รอ')||r[I.Status].includes('ผ่านการพิจาณา')).length;
  const nw=FR.filter(r=>r[I.Type_Request]==='คำขอเปิดวงเงินลูกค้าใหม่').length;
  const pct=(n,t)=>t?(n/t*100).toFixed(1)+'%':'0%';
  document.getElementById('statGrid').innerHTML=`
    <div class="sc"><div class="sl">📋 ทั้งหมด</div><div class="sv">${{total.toLocaleString()}}</div><div class="sp">รายการ</div></div>
    <div class="sc g"><div class="sl">✅ Approve</div><div class="sv" style="color:#2e7d32">${{ap.toLocaleString()}}</div><div class="sp">${{pct(ap,total)}}</div></div>
    <div class="sc r"><div class="sl">❌ Reject</div><div class="sv" style="color:#c62828">${{rj.toLocaleString()}}</div><div class="sp">${{pct(rj,total)}}</div></div>
    <div class="sc o"><div class="sl">⏳ รอพิจารณา</div><div class="sv" style="color:#e65100">${{wt.toLocaleString()}}</div><div class="sp">${{pct(wt,total)}}</div></div>
    <div class="sc"><div class="sl">🆕 เปิดวงเงินใหม่</div><div class="sv">${{nw.toLocaleString()}}</div><div class="sp">จาก ${{total.toLocaleString()}} รายการ</div></div>`;
}}
const PAL=['#1976d2','#43a047','#e53935','#fb8c00','#8e24aa','#00acc1','#f4511e','#6d4c41','#546e7a','#39796b','#c0ca33','#fdd835','#ff7043','#78909c','#66bb6a'];
function mkChart(id,type,labels,data){{
  if(CI[id])CI[id].destroy();
  const ctx=document.getElementById(id);if(!ctx)return;
  CI[id]=new Chart(ctx,{{type,data:{{labels,datasets:[{{data,backgroundColor:PAL.slice(0,labels.length),
    borderColor:type==='line'?'#1976d2':undefined,borderWidth:type==='line'?2:0,
    fill:type==='line',tension:.4,pointRadius:4}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:type==='doughnut'||type==='pie',position:'right',labels:{{font:{{size:11}},boxWidth:14}}}},
        tooltip:{{callbacks:{{label:function(c){{const tot=c.dataset.data.reduce((a,b)=>a+b,0);const pct=tot?(c.raw/tot*100).toFixed(1):0;return ` ${{c.label}}: ${{Number(c.raw).toLocaleString()}} (${{pct}}%)`;}}}}}} }},
      scales:(type==='bar'||type==='line')?{{y:{{ticks:{{font:{{size:10}}}},grid:{{color:'#f0f4f8'}}}},x:{{ticks:{{font:{{size:10}},maxRotation:35}}}}}}:{{}},
      onClick:(evt,els)=>{{if(!els.length)return;handleClick(id,labels[els[0].index]);}}
    }}
  }});
}}
function renderCharts(){{
  const stC={{}},tpC={{}},pvC={{}},incC={{}},mC={{}};
  FR.forEach(r=>{{
    stC[r[I.Status]]=(stC[r[I.Status]]||0)+1;
    tpC[r[I.Type_Request]]=(tpC[r[I.Type_Request]]||0)+1;
    if(r[I.province]&&r[I.province]!='ไม่ระบุ')pvC[r[I.province]]=(pvC[r[I.province]]||0)+1;
    if(r[I.Estimated_annual_income]&&r[I.Estimated_annual_income]!='ไม่ระบุ')incC[r[I.Estimated_annual_income]]=(incC[r[I.Estimated_annual_income]]||0)+1;
    const m=r[I.month_CE],y=r[I.year_BE];if(m&&y)mC[`${{y}}-${{m}}`]=(mC[`${{y}}-${{m}}`]||0)+1;
  }});
  mkChart('cStatus','doughnut',Object.keys(stC),Object.values(stC));
  mkChart('cType','doughnut',Object.keys(tpC),Object.values(tpC));
  mkChart('cBiz','bar',Object.keys(CD.biz),Object.values(CD.biz));
  const pvTop=Object.entries(pvC).sort((a,b)=>b[1]-a[1]).slice(0,15);
  mkChart('cProv','bar',pvTop.map(x=>x[0]),pvTop.map(x=>x[1]));
  const INC=['น้อยกว่า 5 ล้านบาท/ปี','5 ล้านแต่ไม่เกิน 10 ล้านบาท/ปี','10 ล้านแต่ไม่เกิน 50 ล้านบาท/ปี','50 ล้านแต่ไม่เกิน 100 ล้านบาท/ปี','100 ล้านแต่ไม่เกิน 500 ล้านบาท/ปี','500 ล้านบาท/ปี ขึ้นไป'];
  const iL=INC.filter(k=>incC[k]);mkChart('cInc','bar',iL,iL.map(k=>incC[k]||0));
  const TH=['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  const byM={{}};Object.entries(mC).forEach(([k,v])=>{{const[,m]=k.split('-');byM[+m]=(byM[+m]||0)+v;}});
  mkChart('cMonth','line',TH,Array.from({{length:12}},(_,i)=>byM[i+1]||0));
}}
function handleClick(id,label){{
  if(id==='cStatus')document.getElementById('fStatus').value=label;
  else if(id==='cType')document.getElementById('fType').value=label;
  else if(id==='cProv')document.getElementById('fProv').value=label;
  applyFilter();
}}
function renderTable(){{
  const q=(document.getElementById('tSearch').value||'').toLowerCase();
  let rows=FR;
  if(q)rows=rows.filter(r=>(r[I.Title]||'').toLowerCase().includes(q)||(r[I.registration_number]||'').includes(q)||(r[I.province]||'').toLowerCase().includes(q)||(r[I.contact_name]||'').toLowerCase().includes(q));
  document.getElementById('rowCount').textContent=rows.length.toLocaleString()+' รายการ';
  const total=Math.ceil(rows.length/PS)||1;if(CP>total)CP=1;
  const s=(CP-1)*PS,pg=rows.slice(s,s+PS);
  document.getElementById('tBody').innerHTML=pg.map((r,i)=>{{
    let cls='bo';const st=r[I.Status]||'';
    if(st==='Approve')cls='ba';else if(st.includes('Reject')||st.includes('ไม่อนุมัติ'))cls='br';else if(st.includes('รอ')||st.includes('ผ่าน'))cls='bw';
    return `<tr><td style="color:#9e9e9e;font-size:.73rem">${{s+i+1}}</td><td title="${{r[I.Title]}}" style="font-weight:600">${{r[I.Title]||'-'}}</td><td style="font-family:monospace;font-size:.77rem">${{r[I.registration_number]||'-'}}</td><td><span class="badge ${{cls}}">${{r[I.Status]||'-'}}</span></td><td style="font-size:.77rem">${{r[I.Type_Request]||'-'}}</td><td>${{r[I.province]||'-'}}</td><td style="font-size:.74rem">${{r[I.Estimated_annual_income]||'-'}}</td><td style="font-weight:700;color:#1565c0">${{r[I.limit]||'-'}}</td><td style="font-size:.77rem;color:#78909c">${{r[I.date_str]||'-'}}</td><td><button class="pb" style="font-size:.75rem;padding:4px 10px" onclick="showDetail(${{JSON.stringify(r)}})">ดู</button></td></tr>`;
  }}).join('');
  renderPg(total,rows);
}}
function renderPg(total,rows){{
  const el=document.getElementById('pg');if(total<=1){{el.innerHTML='';return;}}
  let h=`<span class="pi">หน้า ${{CP}}/${{total}} (${{rows.length.toLocaleString()}} รายการ)</span>`;
  h+=`<button class="pb" onclick="goPg(1)" ${{CP===1?'disabled':''}}>«</button>`;
  h+=`<button class="pb" onclick="goPg(${{CP-1}})" ${{CP===1?'disabled':''}}>‹</button>`;
  const sp=Math.max(1,CP-2),ep=Math.min(total,CP+2);
  for(let i=sp;i<=ep;i++)h+=`<button class="pb ${{i===CP?'active':''}}" onclick="goPg(${{i}})">${{i}}</button>`;
  h+=`<button class="pb" onclick="goPg(${{CP+1}})" ${{CP===total?'disabled':''}}>›</button>`;
  h+=`<button class="pb" onclick="goPg(${{total}})" ${{CP===total?'disabled':''}}>»</button>`;
  el.innerHTML=h;
}}
function goPg(p){{CP=p;renderTable();document.querySelector('.table-card').scrollIntoView({{behavior:'smooth'}});}}
function showDetail(r){{
  const fields=[['Title','ชื่อกิจการ'],['registration_number','เลขทะเบียน'],['Status','สถานะ'],['Type_Request','ประเภทคำขอ'],['biz_clean','ประเภทธุรกิจ'],['province','จังหวัด'],['district','อำเภอ'],['contact_name','ชื่อผู้ติดต่อ'],['position','ตำแหน่ง'],['contact_number','เบอร์ผู้ติดต่อ'],['telephone','โทรศัพท์'],['Estimated_annual_income','รายได้ประจำปี'],['limit','วงเงิน'],['product_group','กลุ่มสินค้า'],['date_str','วันที่สร้าง']];
  document.getElementById('mTitle').textContent=r[I.Title]||'รายละเอียด';
  document.getElementById('mBody').innerHTML=fields.map(([k,label])=>`<div class="di"><div class="dl">${{label}}</div><div class="dv">${{r[I[k]]||'-'}}</div></div>`).join('');
  document.getElementById('modal').classList.add('open');
}}
function closeModal(evt){{if(!evt||evt.target===document.getElementById('modal'))document.getElementById('modal').classList.remove('open');}}
</script>
</body>
</html>"""

# ─── Main ─────────────────────────────────────────────────
if __name__ == '__main__':
    # Thai timezone
    tz_thai = timezone(timedelta(hours=7))
    now_thai = datetime.now(tz_thai)
    thai_months = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.',
                   'ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']
    update_date = f"{now_thai.day} {thai_months[now_thai.month-1]} {now_thai.year + 543}"

    print(f"=== KYC Dashboard Builder ===")
    print(f"Update date: {update_date}")

    token = get_token()
    items = fetch_all_items(token)
    rows  = transform(items)
    charts = aggregate(rows)

    html = build_html(rows, charts, update_date)

    out_path = 'index.html'
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Written: {out_path} ({len(html):,} chars, {len(rows):,} rows)")
