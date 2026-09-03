#!/usr/bin/env python3
"""
build_dashboard.py
------------------
1. Authenticate to Microsoft Graph via client-credentials flow
2. Fetch ALL items from SharePoint KYCData1 (handles pagination)
3. Generate index.html — dark KYC Dashboard, NO credit-limit fields shown
"""

import os
import sys
import json
import datetime
import requests

# ── CONFIG ───────────────────────────────────────────────────────────────────
TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]

SITE_HOST  = "dohomegroup.sharepoint.com"
SITE_PATH  = "/sites/KYC"
LIST_NAME  = "KYCData1"
SP_DISP    = f"https://{SITE_HOST}{SITE_PATH}/Lists/{LIST_NAME}/DispForm.aspx?ID="

STATUS_COLOR = {
    "Approve":                                                    "#22c55e",
    "\u0e2a\u0e34\u0e19\u0e40\u0e0a\u0e37\u0e48\u0e2d Reject":  "#ef4444",
    "Reject : \u0e44\u0e21\u0e48\u0e1c\u0e48\u0e32\u0e19\u0e40\u0e01\u0e13\u0e11\u0e4c": "#f97316",
    "Reject \u0e44\u0e21\u0e48\u0e2d\u0e19\u0e38\u0e21\u0e31\u0e15\u0e34":                "#dc2626",
    "D3 Reject":                                                  "#f87171",
    "GM Reject":                                                  "#fb923c",
    "\u0e1c\u0e48\u0e32\u0e19\u0e01\u0e32\u0e23\u0e1e\u0e34\u0e08\u0e32\u0e13\u0e32\u0e40\u0e1a\u0e37\u0e49\u0e2d\u0e07\u0e15\u0e49\u0e19": "#84cc16",
    "\u0e23\u0e2d\u0e1c\u0e39\u0e49\u0e08\u0e31\u0e14\u0e01\u0e32\u0e23 D3 \u0e2d\u0e19\u0e38\u0e21\u0e31\u0e15\u0e34":                       "#facc15",
    "\u0e23\u0e2d\u0e1c\u0e39\u0e49\u0e08\u0e31\u0e14\u0e01\u0e32\u0e23 D3 /\u0e1c\u0e39\u0e49\u0e08\u0e31\u0e14\u0e01\u0e32\u0e23 GM \u0e2d\u0e19\u0e38\u0e21\u0e31\u0e15\u0e34": "#eab308",
    "\u0e23\u0e2d GM \u0e2d\u0e19\u0e38\u0e21\u0e31\u0e15\u0e34":                          "#f59e0b",
    "\u0e23\u0e2d GM /\u0e1c\u0e08\u0e01.\u0e1c\u0e19.(\u0e2d\u0e32\u0e27\u0e38\u0e42\u0e2a)\u0e2a\u0e19\u0e31\u0e1a\u0e2a\u0e19\u0e38\u0e19\u0e07\u0e32\u0e19\u0e02\u0e32\u0e22 POS\u00a0\u0e2d\u0e19\u0e38\u0e21\u0e31\u0e15\u0e34": "#fbbf24",
    "\u0e23\u0e2d\u0e2a\u0e34\u0e19\u0e40\u0e0a\u0e37\u0e48\u0e2d\u0e2d\u0e19\u0e38\u0e21\u0e31\u0e15\u0e34": "#a78bfa",
    "Reject : \u0e44\u0e21\u0e48\u0e2a\u0e19\u0e43\u0e08 (\u0e44\u0e21\u0e48\u0e1c\u0e48\u0e32\u0e19\u0e40\u0e01\u0e13\u0e11\u0e4c)": "#fb7185",
}

# ── AUTH ──────────────────────────────────────────────────────────────────────
def get_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    r = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default",
    }, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]


# ── FETCH ─────────────────────────────────────────────────────────────────────
def fetch_all_items(token):
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    url = (
        f"https://graph.microsoft.com/v1.0/sites/"
        f"{SITE_HOST}:{SITE_PATH}:/lists/{LIST_NAME}/items"
        f"?$expand=fields&$top=999&$select=id,fields"
    )
    items = []
    while url:
        r = requests.get(url, headers=headers, timeout=60)
        r.raise_for_status()
        data = r.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        print(f"  fetched {len(items)} items...", flush=True)
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
    return {
        "id":         int(item["id"]),
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


# ── STATS ─────────────────────────────────────────────────────────────────────
def compute_stats(rows):
    from collections import Counter
    sc = Counter(r["status"]   for r in rows)
    tc = Counter(r["type_req"] for r in rows)
    mc = Counter(r["ym"]       for r in rows if r["ym"])
    last30 = sorted(mc.items())[-30:]
    return dict(sc), dict(tc), [{"ym": k, "cnt": v} for k, v in last30]


# ── CSS ───────────────────────────────────────────────────────────────────────
CSS = (
    ":root{--bg:#0f172a;--s:#1e293b;--s2:#263348;--bd:#334155;"
    "--t:#f1f5f9;--t2:#94a3b8;--ac:#38bdf8;--ra:12px;}\n"
    "*{box-sizing:border-box;margin:0;padding:0;}\n"
    "body{background:var(--bg);color:var(--t);"
    "font-family:'Sarabun','Noto Sans Thai',sans-serif;"
    "font-size:15px;min-height:100vh;}\n"
    "header{background:linear-gradient(135deg,#1e3a5f,#0f172a);"
    "padding:16px 24px;display:flex;align-items:center;gap:14px;"
    "border-bottom:1px solid var(--bd);position:sticky;top:0;z-index:100;}\n"
    "header h1{font-size:20px;font-weight:700;color:var(--ac);}\n"
    ".sub{font-size:12px;color:var(--t2);}\n"
    ".logo{width:40px;height:40px;background:var(--ac);border-radius:9px;"
    "display:flex;align-items:center;justify-content:center;font-size:20px;}\n"
    ".upd{font-size:11px;color:var(--t2);margin-left:auto;}\n"
    ".main{padding:20px 24px;max-width:1600px;margin:0 auto;}\n"
    ".kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));"
    "gap:14px;margin-bottom:20px;}\n"
    ".kpi{background:var(--s);border:1px solid var(--bd);border-radius:var(--ra);"
    "padding:18px;position:relative;overflow:hidden;transition:.2s;}\n"
    ".kpi:hover{border-color:var(--ac);transform:translateY(-2px);}\n"
    ".kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;"
    "background:var(--kc,var(--ac));}\n"
    ".kpi .lbl{font-size:11px;color:var(--t2);text-transform:uppercase;"
    "letter-spacing:.05em;margin-bottom:6px;}\n"
    ".kpi .val{font-size:28px;font-weight:800;color:var(--kc,var(--t));}\n"
    ".kpi .nt{font-size:11px;color:var(--t2);margin-top:3px;}\n"
    ".charts-row{display:grid;grid-template-columns:1fr 1fr 1.6fr;gap:14px;margin-bottom:20px;}\n"
    "@media(max-width:900px){.charts-row{grid-template-columns:1fr 1fr;}}\n"
    "@media(max-width:580px){.charts-row{grid-template-columns:1fr;}}\n"
    ".card{background:var(--s);border:1px solid var(--bd);border-radius:var(--ra);padding:18px;}\n"
    ".card h3{font-size:12px;color:var(--t2);margin-bottom:14px;"
    "text-transform:uppercase;letter-spacing:.05em;}\n"
    ".cw{position:relative;height:210px;}\n"
    ".frow{display:flex;gap:9px;flex-wrap:wrap;margin-bottom:12px;align-items:center;}\n"
    ".frow input,.frow select{background:var(--s2);border:1px solid var(--bd);"
    "color:var(--t);padding:8px 12px;border-radius:8px;font-size:13px;"
    "font-family:inherit;outline:none;transition:.2s;}\n"
    ".frow input{flex:1;min-width:200px;}\n"
    ".frow input:focus,.frow select:focus{border-color:var(--ac);}\n"
    ".badge-cnt{background:var(--ac);color:#0f172a;font-size:12px;font-weight:700;"
    "padding:4px 12px;border-radius:20px;}\n"
    ".btn-r{background:var(--s2);border:1px solid var(--bd);color:var(--t2);"
    "padding:8px 12px;border-radius:8px;cursor:pointer;font-size:13px;"
    "font-family:inherit;transition:.2s;}\n"
    ".btn-r:hover{border-color:var(--ac);color:var(--ac);}\n"
    ".tw{overflow-x:auto;border-radius:var(--ra);border:1px solid var(--bd);}\n"
    "table{width:100%;border-collapse:collapse;font-size:13px;}\n"
    "thead{background:var(--s2);position:sticky;top:0;z-index:10;}\n"
    "th{padding:11px 13px;text-align:left;font-size:11px;font-weight:600;"
    "color:var(--t2);text-transform:uppercase;letter-spacing:.04em;"
    "border-bottom:1px solid var(--bd);cursor:pointer;user-select:none;white-space:nowrap;}\n"
    "th:hover{color:var(--ac);}\n"
    "td{padding:10px 13px;border-bottom:1px solid var(--bd);vertical-align:middle;}\n"
    "tr:last-child td{border-bottom:none;}\n"
    "tbody tr{cursor:pointer;transition:.1s;}\n"
    "tbody tr:hover{background:var(--s2);}\n"
    "tbody tr:hover td:first-child{border-left:3px solid var(--ac);padding-left:10px;}\n"
    ".sb{display:inline-block;padding:2px 9px;border-radius:20px;"
    "font-size:11px;font-weight:600;}\n"
    ".pg{display:flex;align-items:center;gap:5px;margin-top:12px;flex-wrap:wrap;}\n"
    ".pg button{background:var(--s2);border:1px solid var(--bd);color:var(--t);"
    "padding:5px 11px;border-radius:7px;cursor:pointer;font-size:12px;"
    "font-family:inherit;transition:.2s;}\n"
    ".pg button:hover,.pg button.act{background:var(--ac);color:#0f172a;"
    "border-color:var(--ac);font-weight:700;}\n"
    ".pg button:disabled{opacity:.3;cursor:not-allowed;}\n"
    ".pgi{color:var(--t2);font-size:11px;margin-left:auto;}\n"
    ".ov{display:none;position:fixed;inset:0;background:rgba(0,0,0,.72);"
    "z-index:1000;backdrop-filter:blur(5px);}\n"
    ".ov.open{display:flex;align-items:center;justify-content:center;padding:16px;}\n"
    ".modal{background:var(--s);border:1px solid var(--bd);border-radius:14px;"
    "width:100%;max-width:720px;max-height:90vh;overflow-y:auto;animation:su .2s;}\n"
    "@keyframes su{from{transform:translateY(24px);opacity:0}"
    "to{transform:none;opacity:1}}\n"
    ".mh{padding:18px 22px 14px;border-bottom:1px solid var(--bd);"
    "display:flex;justify-content:space-between;align-items:flex-start;gap:12px;}\n"
    ".mh h2{font-size:17px;font-weight:700;color:var(--ac);line-height:1.4;}\n"
    ".mcl{background:var(--s2);border:1px solid var(--bd);color:var(--t2);"
    "width:34px;height:34px;border-radius:8px;cursor:pointer;font-size:17px;"
    "display:flex;align-items:center;justify-content:center;flex-shrink:0;transition:.2s;}\n"
    ".mcl:hover{border-color:#ef4444;color:#ef4444;}\n"
    ".mb{padding:18px 22px;}\n"
    ".dg{display:grid;grid-template-columns:1fr 1fr;gap:10px;}\n"
    "@media(max-width:480px){.dg{grid-template-columns:1fr;}}\n"
    ".di{background:var(--s2);border:1px solid var(--bd);border-radius:9px;padding:12px;}\n"
    ".di.full{grid-column:1/-1;}\n"
    ".dl{font-size:10px;color:var(--t2);text-transform:uppercase;"
    "letter-spacing:.05em;margin-bottom:4px;}\n"
    ".dv{font-size:14px;font-weight:600;color:var(--t);word-break:break-word;}\n"
    ".mf{padding:14px 22px;border-top:1px solid var(--bd);"
    "display:flex;justify-content:flex-end;gap:9px;}\n"
    ".btn-p{background:var(--ac);color:#0f172a;border:none;padding:9px 20px;"
    "border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;"
    "font-family:inherit;transition:.2s;}\n"
    ".btn-p:hover{background:#7dd3fc;}\n"
    ".btn-s{background:var(--s2);border:1px solid var(--bd);color:var(--t2);"
    "padding:9px 20px;border-radius:8px;font-size:14px;cursor:pointer;"
    "font-family:inherit;transition:.2s;}\n"
    ".btn-s:hover{border-color:var(--ac);color:var(--ac);}\n"
    ".wbox{margin-top:14px;padding:11px 14px;"
    "background:rgba(250,204,21,.08);border:1px solid rgba(250,204,21,.3);"
    "border-radius:9px;font-size:12px;color:var(--t2);}\n"
    "::-webkit-scrollbar{width:5px;height:5px;}\n"
    "::-webkit-scrollbar-track{background:var(--bg);}\n"
    "::-webkit-scrollbar-thumb{background:var(--bd);border-radius:3px;}\n"
)

# ── JavaScript ────────────────────────────────────────────────────────────────
JS = r"""
let filtered=[...ALL],curPage=1,sortCol="id",sortDir=1,curId=null;
const PG=25;
(function initCharts(){
  var sl=Object.keys(SS), sv=Object.values(SS);
  new Chart(document.getElementById("cSt"),{
    type:"doughnut",
    data:{labels:sl,datasets:[{data:sv,
      backgroundColor:sl.map(function(k){return SC[k]||"#64748b";}),
      borderColor:"#1e293b",borderWidth:2}]},
    options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){
      return " "+c.label+": "+c.raw.toLocaleString()+" ("+(c.raw/ALL.length*100).toFixed(1)+"%)";
    }}}},maintainAspectRatio:false}
  });
  new Chart(document.getElementById("cTy"),{
    type:"bar",
    data:{labels:["\u0e40\u0e1b\u0e34\u0e14\u0e27\u0e07\u0e40\u0e07\u0e34\u0e19\u0e43\u0e2b\u0e21\u0e48","\u0e40\u0e1e\u0e34\u0e48\u0e21\u0e27\u0e07\u0e40\u0e07\u0e34\u0e19"],
      datasets:[{data:[
        TC["\u0e04\u0e33\u0e02\u0e2d\u0e40\u0e1b\u0e34\u0e14\u0e27\u0e07\u0e40\u0e07\u0e34\u0e19\u0e25\u0e39\u0e01\u0e04\u0e49\u0e32\u0e43\u0e2b\u0e21\u0e48"]||0,
        TC["\u0e04\u0e33\u0e02\u0e2d\u0e40\u0e1e\u0e34\u0e48\u0e21\u0e27\u0e07\u0e40\u0e07\u0e34\u0e19"]||0
      ],backgroundColor:["#38bdf8","#fb923c"],borderRadius:7,borderSkipped:false}]},
    options:{indexAxis:"y",plugins:{legend:{display:false},tooltip:{callbacks:{
      label:function(c){return " "+c.raw.toLocaleString()+" \u0e23\u0e32\u0e22\u0e01\u0e32\u0e23";}
    }}},scales:{
      x:{grid:{color:"#334155"},ticks:{color:"#94a3b8"}},
      y:{grid:{display:false},ticks:{color:"#94a3b8"}}
    },maintainAspectRatio:false}
  });
  new Chart(document.getElementById("cMo"),{
    type:"line",
    data:{
      labels:MC.map(function(m){return m.ym;}),
      datasets:[{data:MC.map(function(m){return m.cnt;}),
        borderColor:"#38bdf8",backgroundColor:"rgba(56,189,248,.1)",
        fill:true,pointBackgroundColor:"#38bdf8",pointRadius:3,tension:.4}]
    },
    options:{plugins:{legend:{display:false}},scales:{
      x:{grid:{color:"#334155"},ticks:{color:"#94a3b8",maxRotation:45,font:{size:9}}},
      y:{grid:{color:"#334155"},ticks:{color:"#94a3b8"}}
    },maintainAspectRatio:false}
  });
})();

function badge(s){
  var c=SC[s]||"#64748b";
  return "<span class='sb' style='background:"+c+"22;color:"+c+";border:1px solid "+c+"44'>"+s+"</span>";
}
function applyF(){
  var q=document.getElementById("srch").value.trim().toLowerCase();
  var fs=document.getElementById("fSt").value;
  var ft=document.getElementById("fTy").value;
  filtered=ALL.filter(function(r){
    var mq=!q||[r.title,r.reg_num,r.province,r.contact,r.tel,r.owner,r.biz]
      .some(function(v){return v&&v.toLowerCase().indexOf(q)>=0;});
    return mq&&(!fs||r.status===fs)&&(!ft||r.type_req===ft);
  });
  doSort(sortCol,true); curPage=1; render();
}
function resetF(){
  document.getElementById("srch").value="";
  document.getElementById("fSt").value="";
  document.getElementById("fTy").value="";
  applyF();
}
function sortBy(col){
  if(sortCol===col) sortDir*=-1; else{sortCol=col;sortDir=1;}
  doSort(col,false); render();
}
function doSort(col,keep){
  if(!keep&&sortCol!==col) sortDir=1;
  filtered.sort(function(a,b){
    var av=a[col]||"", bv=b[col]||"";
    if(col==="id") return (av-bv)*sortDir;
    return av.toString().localeCompare(bv.toString(),"th")*sortDir;
  });
  document.querySelectorAll("[id^='si']").forEach(function(e){e.textContent="\u2195";});
  var si=document.getElementById("si"+col);
  if(si) si.textContent=sortDir===1?"\u2191":"\u2193";
}
function render(){
  var tot=filtered.length;
  var tp=Math.max(1,Math.ceil(tot/PG));
  curPage=Math.min(curPage,tp);
  var s=(curPage-1)*PG, sl=filtered.slice(s,s+PG);
  document.getElementById("rCnt").textContent=tot.toLocaleString()+" \u0e23\u0e32\u0e22\u0e01\u0e32\u0e23";
  var tb=document.getElementById("tBody");
  if(!sl.length){
    tb.innerHTML="<tr><td colspan='7' style='text-align:center;padding:40px;color:var(--t2)'>\ud83d\udd0d \u0e44\u0e21\u0e48\u0e1e\u0e1a\u0e23\u0e32\u0e22\u0e01\u0e32\u0e23</td></tr>";
  } else {
    tb.innerHTML=sl.map(function(r){
      return "<tr onclick='showD("+r.id+")'>"
        +"<td style='color:var(--t2);font-size:11px'>"+r.id+"</td>"
        +"<td style='font-weight:600;max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap'>"+r.title+"</td>"
        +"<td style='font-family:monospace;color:var(--t2);font-size:12px'>"+(r.reg_num||"\u2013")+"</td>"
        +"<td>"+badge(r.status)+"</td>"
        +"<td style='font-size:12px;color:var(--t2)'>"+(r.type_req||"\u2013")+"</td>"
        +"<td style='font-size:12px'>"+(r.province||"\u2013")+"</td>"
        +"<td style='font-size:12px;color:var(--t2)'>"+(r.created||"\u2013")+"</td>"
        +"</tr>";
    }).join("");
  }
  document.getElementById("bP").disabled=curPage<=1;
  document.getElementById("bN").disabled=curPage>=tp;
  document.getElementById("pI").textContent=
    "\u0e2b\u0e19\u0e49\u0e32 "+curPage+"/"+tp+" \u00b7 "+(s+1)+"\u2013"+Math.min(s+PG,tot)
    +" \u0e08\u0e32\u0e01 "+tot.toLocaleString();
  var btns=[],pv=0,dl=2;
  for(var i=1;i<=tp;i++){
    if(i===1||i===tp||Math.abs(i-curPage)<=dl){
      if(pv&&i-pv>1) btns.push("...");
      btns.push(i); pv=i;
    }
  }
  document.getElementById("pgN").innerHTML=btns.map(function(b){
    if(b==="...") return "<span style='padding:5px 3px;color:var(--t2)'>...</span>";
    return "<button onclick='goP("+b+")' class='"+(b===curPage?"act":"")+"'>"+b+"</button>";
  }).join("");
}
function goP(p){
  curPage=Math.max(1,Math.min(p,Math.ceil(filtered.length/PG)));
  render();
  document.querySelector(".tw").scrollIntoView({behavior:"smooth",block:"nearest"});
}
function showD(id){
  var r=ALL.find(function(x){return x.id===id;}); if(!r) return;
  curId=id;
  document.getElementById("mTi").textContent=r.title;
  var c=SC[r.status]||"#64748b";
  var sb="<span class='sb' style='background:"+c+"22;color:"+c+";border:1px solid "+c+"44'>"+r.status+"</span>";
  var flds=[
    ["","\u0e40\u0e25\u0e02\u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19\u0e19\u0e34\u0e15\u0e34\u0e1a\u0e38\u0e04\u0e04\u0e25",r.reg_num||"\u2013"],
    ["","\u0e0a\u0e37\u0e48\u0e2d\u0e08\u0e14\u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19",r.registered||r.title],
    ["","\u0e2a\u0e16\u0e32\u0e19\u0e30",sb],
    ["","\u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17\u0e04\u0e33\u0e02\u0e2d",r.type_req||"\u2013"],
    ["full","\u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17\u0e18\u0e38\u0e23\u0e01\u0e34\u0e08",r.biz||"\u2013"],
    ["","\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14",r.province||"\u2013"],
    ["","\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48\u0e2a\u0e23\u0e49\u0e32\u0e07\u0e23\u0e32\u0e22\u0e01\u0e32\u0e23",r.created||"\u2013"],
    ["","\u0e1c\u0e39\u0e49\u0e15\u0e34\u0e14\u0e15\u0e48\u0e2d",r.contact||"\u2013"],
    ["","\u0e40\u0e1a\u0e2d\u0e23\u0e4c\u0e42\u0e17\u0e23\u0e28\u0e31\u0e1e\u0e17\u0e4c",r.tel||"\u2013"],
    ["full","Owner / \u0e2a\u0e32\u0e02\u0e32",r.owner||"\u2013"]
  ];
  document.getElementById("mBo").innerHTML=
    "<div class='dg'>"+flds.map(function(f){
      return "<div class='di "+f[0]+"'><div class='dl'>"+f[1]+"</div><div class='dv'>"+f[2]+"</div></div>";
    }).join("")+"</div>"
    +"<div class='wbox'>\u26a0\ufe0f \u0e02\u0e49\u0e2d\u0e21\u0e39\u0e25\u0e27\u0e07\u0e40\u0e07\u0e34\u0e19 "
    +"\u0e40\u0e04\u0e23\u0e14\u0e34\u0e15\u0e40\u0e17\u0e2d\u0e23\u0e4c\u0e21 "
    +"\u0e41\u0e25\u0e30\u0e2b\u0e25\u0e31\u0e01\u0e1b\u0e23\u0e30\u0e01\u0e31\u0e19 "
    +"\u0e44\u0e21\u0e48\u0e41\u0e2a\u0e14\u0e07\u0e43\u0e19\u0e41\u0e14\u0e0a\u0e1a\u0e2d\u0e23\u0e4c\u0e14\u0e19\u0e35\u0e49\u0e15\u0e32\u0e21\u0e19\u0e42\u0e22\u0e1a\u0e32\u0e22</div>";
  document.getElementById("ov").classList.add("open");
}
function closeM(){document.getElementById("ov").classList.remove("open");}
function closeOut(e){if(e.target===document.getElementById("ov")) closeM();}
function openSP(){if(curId) window.open(BASE+curId,"_blank");}
document.addEventListener("keydown",function(e){if(e.key==="Escape") closeM();});
document.getElementById("srch").addEventListener("input",applyF);
document.getElementById("fSt").addEventListener("change",applyF);
document.getElementById("fTy").addEventListener("change",applyF);
doSort("id",false); render();
"""


# ── HTML BUILDER ──────────────────────────────────────────────────────────────
def build_html(rows, status_cnt, type_cnt, monthly, updated_at):
    total        = len(rows)
    approve      = status_cnt.get("Approve", 0)
    reject_total = sum(v for k, v in status_cnt.items() if "Reject" in k)
    pending      = total - approve - reject_total
    new_cr = type_cnt.get("\u0e04\u0e33\u0e02\u0e2d\u0e40\u0e1b\u0e34\u0e14\u0e27\u0e07\u0e40\u0e07\u0e34\u0e19\u0e25\u0e39\u0e01\u0e04\u0e49\u0e32\u0e43\u0e2b\u0e21\u0e48", 0)
    add_cr = type_cnt.get("\u0e04\u0e33\u0e02\u0e2d\u0e40\u0e1e\u0e34\u0e48\u0e21\u0e27\u0e07\u0e40\u0e07\u0e34\u0e19", 0)
    ap_pct = f"{approve/total*100:.1f}" if total else "0.0"
    rj_pct = f"{reject_total/total*100:.1f}" if total else "0.0"

    data_js = (
        "const ALL=" + json.dumps(rows, ensure_ascii=False) + ";\n"
        "const SC="  + json.dumps(STATUS_COLOR, ensure_ascii=False) + ";\n"
        "const MC="  + json.dumps(monthly, ensure_ascii=False) + ";\n"
        "const SS="  + json.dumps(status_cnt, ensure_ascii=False) + ";\n"
        "const TC="  + json.dumps(type_cnt, ensure_ascii=False) + ";\n"
        'const BASE="' + SP_DISP + '";\n'
    )

    status_opts = "\n".join(
        f"      <option>{s}</option>" for s in status_cnt
    )

    kpi_html = (
        f'<div class="kpi" style="--kc:#38bdf8">'
        f'<div class="lbl">\u0e23\u0e32\u0e22\u0e01\u0e32\u0e23\u0e17\u0e31\u0e49\u0e07\u0e2b\u0e21\u0e14</div>'
        f'<div class="val">{total:,}</div>'
        f'<div class="nt">KYC-Data</div></div>\n'

        f'<div class="kpi" style="--kc:#22c55e">'
        f'<div class="lbl">Approve</div>'
        f'<div class="val">{approve:,}</div>'
        f'<div class="nt">{ap_pct}%</div></div>\n'

        f'<div class="kpi" style="--kc:#ef4444">'
        f'<div class="lbl">Reject \u0e23\u0e27\u0e21</div>'
        f'<div class="val">{reject_total:,}</div>'
        f'<div class="nt">{rj_pct}%</div></div>\n'

        f'<div class="kpi" style="--kc:#facc15">'
        f'<div class="lbl">\u0e23\u0e2d\u0e14\u0e33\u0e40\u0e19\u0e34\u0e19\u0e01\u0e32\u0e23</div>'
        f'<div class="val">{pending:,}</div>'
        f'<div class="nt">\u0e23\u0e30\u0e2b\u0e27\u0e48\u0e32\u0e07\u0e1e\u0e34\u0e08\u0e32\u0e23\u0e13\u0e32</div></div>\n'

        f'<div class="kpi" style="--kc:#a78bfa">'
        f'<div class="lbl">\u0e40\u0e1b\u0e34\u0e14\u0e27\u0e07\u0e40\u0e07\u0e34\u0e19\u0e43\u0e2b\u0e21\u0e48</div>'
        f'<div class="val">{new_cr:,}</div>'
        f'<div class="nt">\u0e23\u0e32\u0e22\u0e01\u0e32\u0e23</div></div>\n'

        f'<div class="kpi" style="--kc:#fb923c">'
        f'<div class="lbl">\u0e40\u0e1e\u0e34\u0e48\u0e21\u0e27\u0e07\u0e40\u0e07\u0e34\u0e19</div>'
        f'<div class="val">{add_cr:,}</div>'
        f'<div class="nt">\u0e23\u0e32\u0e22\u0e01\u0e32\u0e23</div></div>\n'
    )

    return (
        '<!DOCTYPE html>\n<html lang="th">\n<head>\n'
        '<meta charset="UTF-8"/>\n'
        '<meta name="viewport" content="width=device-width,initial-scale=1.0"/>\n'
        '<title>KYC Dashboard \u2013 DoHome</title>\n'
        '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>\n'
        '<style>' + CSS + '</style>\n'
        '</head>\n<body>\n'
        '<header>'
        '<div class="logo">\U0001f3e0</div>'
        '<div><h1>KYC Credit Dashboard</h1>'
        '<div class="sub">DoHome Group \u00b7 KYC-Data \u00b7 \u0e44\u0e21\u0e48\u0e41\u0e2a\u0e14\u0e07\u0e02\u0e49\u0e2d\u0e21\u0e39\u0e25\u0e27\u0e07\u0e40\u0e07\u0e34\u0e19</div></div>'
        f'<span class="upd">\U0001f504 \u0e2d\u0e31\u0e1b\u0e40\u0e14\u0e15: {updated_at}</span>'
        '</header>\n'
        '<div class="main">\n'
        '<div class="kpi-grid">\n' + kpi_html + '</div>\n'
        '<div class="charts-row">\n'
        '  <div class="card"><h3>\u0e2a\u0e16\u0e32\u0e19\u0e30\u0e04\u0e33\u0e02\u0e2d</h3>'
        '<div class="cw"><canvas id="cSt"></canvas></div></div>\n'
        '  <div class="card"><h3>\u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17\u0e04\u0e33\u0e02\u0e2d</h3>'
        '<div class="cw"><canvas id="cTy"></canvas></div></div>\n'
        '  <div class="card"><h3>\u0e41\u0e19\u0e27\u0e42\u0e19\u0e49\u0e21\u0e23\u0e32\u0e22\u0e40\u0e14\u0e37\u0e2d\u0e19</h3>'
        '<div class="cw"><canvas id="cMo"></canvas></div></div>\n'
        '</div>\n'
        '<div class="card">\n'
        '  <h3 style="margin-bottom:12px;">\u0e23\u0e32\u0e22\u0e01\u0e32\u0e23 KYC \u0e17\u0e31\u0e49\u0e07\u0e2b\u0e21\u0e14'
        '    <small style="font-weight:400;font-size:11px;color:var(--t2)">'
        '      &nbsp;\u00b7 \u0e44\u0e21\u0e48\u0e41\u0e2a\u0e14\u0e07\u0e27\u0e07\u0e40\u0e07\u0e34\u0e19'
        '      \u00b7 \u0e04\u0e25\u0e34\u0e01\u0e41\u0e16\u0e27\u0e14\u0e39\u0e23\u0e32\u0e22\u0e25\u0e30\u0e40\u0e2d\u0e35\u0e22\u0e14</small></h3>\n'
        '  <div class="frow">\n'
        '    <input type="text" id="srch" placeholder="\ud83d\udd0d  \u0e04\u0e49\u0e19\u0e2b\u0e32\u0e0a\u0e37\u0e48\u0e2d\u0e1a\u0e23\u0e34\u0e29\u0e31\u0e17, \u0e40\u0e25\u0e02\u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19, \u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14\u2026"/>\n'
        '    <select id="fSt">\n'
        '      <option value="">\u2014 \u0e2a\u0e16\u0e32\u0e19\u0e30\u0e17\u0e31\u0e49\u0e07\u0e2b\u0e21\u0e14 \u2014</option>\n'
        + status_opts + '\n'
        '    </select>\n'
        '    <select id="fTy">\n'
        '      <option value="">\u2014 \u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17\u0e17\u0e31\u0e49\u0e07\u0e2b\u0e21\u0e14 \u2014</option>\n'
        '      <option>\u0e04\u0e33\u0e02\u0e2d\u0e40\u0e1b\u0e34\u0e14\u0e27\u0e07\u0e40\u0e07\u0e34\u0e19\u0e25\u0e39\u0e01\u0e04\u0e49\u0e32\u0e43\u0e2b\u0e21\u0e48</option>\n'
        '      <option>\u0e04\u0e33\u0e02\u0e2d\u0e40\u0e1e\u0e34\u0e48\u0e21\u0e27\u0e07\u0e40\u0e07\u0e34\u0e19</option>\n'
        '    </select>\n'
        '    <button class="btn-r" onclick="resetF()">\u21ba \u0e23\u0e35\u0e40\u0e0b\u0e47\u0e15</button>\n'
        '    <span class="badge-cnt" id="rCnt">0 \u0e23\u0e32\u0e22\u0e01\u0e32\u0e23</span>\n'
        '  </div>\n'
        '  <div class="tw"><table>\n'
        '    <thead><tr>\n'
        '      <th onclick="sortBy(\'id\')"># <span id="siid">\u2195</span></th>\n'
        '      <th onclick="sortBy(\'title\')">\u0e0a\u0e37\u0e48\u0e2d\u0e19\u0e34\u0e15\u0e34\u0e1a\u0e38\u0e04\u0e04\u0e25 <span id="sititle">\u2195</span></th>\n'
        '      <th onclick="sortBy(\'reg_num\')">\u0e40\u0e25\u0e02\u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19 <span id="sireg_num">\u2195</span></th>\n'
        '      <th onclick="sortBy(\'status\')">\u0e2a\u0e16\u0e32\u0e19\u0e30 <span id="sistatus">\u2195</span></th>\n'
        '      <th onclick="sortBy(\'type_req\')">\u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17 <span id="sitype_req">\u2195</span></th>\n'
        '      <th onclick="sortBy(\'province\')">\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14 <span id="siprovince">\u2195</span></th>\n'
        '      <th onclick="sortBy(\'created\')">\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48\u0e2a\u0e23\u0e49\u0e32\u0e07 <span id="sicreated">\u2195</span></th>\n'
        '    </tr></thead>\n'
        '    <tbody id="tBody"></tbody>\n'
        '  </table></div>\n'
        '  <div class="pg">\n'
        '    <button onclick="goP(curPage-1)" id="bP">\u2039 \u0e01\u0e48\u0e2d\u0e19\u0e2b\u0e19\u0e49\u0e32</button>\n'
        '    <div id="pgN" style="display:flex;gap:4px;flex-wrap:wrap;"></div>\n'
        '    <button onclick="goP(curPage+1)" id="bN">\u0e16\u0e31\u0e14\u0e44\u0e1b \u203a</button>\n'
        '    <span class="pgi" id="pI"></span>\n'
        '  </div>\n'
        '</div>\n</div>\n'
        '<div class="ov" id="ov" onclick="closeOut(event)">\n'
        '  <div class="modal">\n'
        '    <div class="mh"><h2 id="mTi"></h2>'
        '<button class="mcl" onclick="closeM()">\u2715</button></div>\n'
        '    <div class="mb" id="mBo"></div>\n'
        '    <div class="mf">\n'
        '      <button class="btn-s" onclick="closeM()">\u0e1b\u0e34\u0e14</button>\n'
        '      <button class="btn-p" onclick="openSP()">'
        '\U0001f517 \u0e14\u0e39\u0e43\u0e19 SharePoint</button>\n'
        '    </div>\n'
        '  </div>\n'
        '</div>\n'
        '<script>\n' + data_js + JS + '\n</script>\n'
        '</body>\n</html>\n'
    )


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    print("Authenticating to Microsoft Graph...")
    token = get_token()

    print("Fetching SharePoint list items...")
    raw = fetch_all_items(token)
    print(f"Total fetched: {len(raw)}")

    rows = [to_row(i) for i in raw]
    status_cnt, type_cnt, monthly = compute_stats(rows)

    tz_thai = datetime.timezone(datetime.timedelta(hours=7))
    updated_at = datetime.datetime.now(tz_thai).strftime("%d/%m/%Y %H:%M \u0e19. (\u0e44\u0e17\u0e22)")

    print("Building HTML...")
    html = build_html(rows, status_cnt, type_cnt, monthly, updated_at)

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    kb = len(html.encode("utf-8")) // 1024
    print(f"Done — {kb} KB written to index.html")


if __name__ == "__main__":
    main()
