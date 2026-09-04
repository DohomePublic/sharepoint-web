# DemoApp · Daily Dashboard (SharePoint → GitHub Pages)

Dashboard รายวันของ SharePoint List **DemoApp** (ไซต์ `AC-Accounting`)
ดึงข้อมูลอัตโนมัติด้วย GitHub Actions แล้ว publish เป็นหน้าเว็บไฟล์เดียว

```
sharepoint-web/
├── .github/
│   └── workflows/
│       └── update-dashboard.yml   ← GitHub Actions workflow (ดึงข้อมูล + deploy)
├── scripts/
│   ├── build_dashboard.py         ← ดึงข้อมูล SharePoint + สร้าง HTML
│   └── template.html              ← เทมเพลต (มี placeholder __DATA__/__ACL__/__GENAT__)
├── data/                          ← CSV สำรองสำหรับ dev/offline (ไม่ใช้ใน CI)
│   ├── DemoApp.csv
│   └── Admin_KycNew.csv
├── index.html                     ← Dashboard (auto-generated ห้ามแก้มือ)
└── README.md
```

---

## 1. Data Flow

```
SharePoint List DemoApp ─┐
                         ├─► Microsoft Graph (app-only) ─► build_dashboard.py ─► index.html ─► GitHub Pages
SharePoint List Admin_KycNew ─┘                                    │
                                                                   └─► ACL 65 บัญชี (คอลัมน์ Title = อีเมล)
```

| ลิสต์ | บทบาท | คอลัมน์ที่ใช้ |
|---|---|---|
| `DemoApp` | ข้อมูลคำขอ KYC/สินเชื่อ | 31 คอลัมน์ (ดู `FIELD_MAP` ใน `build_dashboard.py`) |
| `Admin_KycNew` | ทะเบียนผู้มีสิทธิ์ (ACL) | `Title` = อีเมล → แปลงเป็น role อัตโนมัติ |

**กติกาแปลงอีเมล → สิทธิ์** (ตรงกันทั้งฝั่ง Python และ JavaScript)

| รูปแบบอีเมล | ประเภท | Role | ขอบเขต |
|---|---|---|---|
| `BI-Operation<XX>_GM@` / `BI-VOperation<XX>_GM@` | BI_OPERATION | `BIOPS` | สาขา `<XX>OO` |
| `GM-<XX>@` | BRANCH_GM | `GM` | สาขา `<XX>OO` |
| `GM-trainee@` | BRANCH_GM | `VIEWER` | ดูอย่างเดียว ไม่เห็น PII |
| `Dohometogogm-<XX>@` | TOGO_GM | `GM` | สาขา `<XX>OO` |
| อีเมลรายบุคคล (HQ) | NAMED_USER | `CREDIT` | ทุกสาขา |

---

## 2. ตั้งค่าครั้งแรก

### 2.1 สร้าง Azure AD App (app-only, ไม่ต้องมีผู้ใช้ล็อกอิน)

1. Azure Portal → **Microsoft Entra ID → App registrations → New registration**
   ตั้งชื่อเช่น `sp-demoapp-dashboard` → Register
2. **Certificates & secrets → New client secret** → คัดลอกค่า **Value** (เห็นครั้งเดียว)
3. **API permissions → Add a permission → Microsoft Graph → Application permissions**
   - `Sites.Selected` *(แนะนำ — ให้สิทธิ์เฉพาะไซต์)* หรือ `Sites.Read.All`
   - กด **Grant admin consent**
4. ถ้าใช้ `Sites.Selected` ต้องให้สิทธิ์เฉพาะไซต์เพิ่ม (รันโดย SharePoint Admin):

```powershell
# PowerShell + PnP.PowerShell
Connect-PnPOnline -Url "https://dohomegroup.sharepoint.com/sites/AC-Accounting" -Interactive
Grant-PnPAzureADAppSitePermission -AppId "<CLIENT_ID>" -DisplayName "sp-demoapp-dashboard" -Permissions Read
```

### 2.2 ใส่ Secrets ใน GitHub

**Settings → Secrets and variables → Actions → New repository secret**

| ชื่อ | ค่า |
|---|---|
| `TENANT_ID` | Directory (tenant) ID |
| `CLIENT_ID` | Application (client) ID |
| `CLIENT_SECRET` | ค่า secret จากข้อ 2.1.2 |

**Variables** (ไม่บังคับ — มีค่า default ในสคริปต์อยู่แล้ว)

| ชื่อ | ค่า default |
|---|---|
| `SITE_HOSTNAME` | `dohomegroup.sharepoint.com` |
| `SITE_PATH` | `/sites/AC-Accounting` |
| `LIST_DEMOAPP` | `DemoApp` |
| `LIST_ACL` | `Admin_KycNew` |

### 2.3 เปิด GitHub Pages

**Settings → Pages → Source = GitHub Actions**

---

## 3. ตารางเวลาอัปเดต

| Trigger | เวลา |
|---|---|
| `schedule` | ทุก **30 นาที** ระหว่าง 07:00–20:00 น. (ไทย) จันทร์–ศุกร์ |
| `schedule` | รอบสรุป 21:00 น. (ไทย) |
| `workflow_dispatch` | กดรันเองได้ทันทีจากแท็บ **Actions** |
| `push` | เมื่อแก้ `scripts/**` หรือไฟล์ workflow |

> ⚠️ cron ของ GitHub Actions อาจดีเลย์ 5–15 นาทีในช่วงที่ระบบมีงานหนาแน่น เป็นพฤติกรรมปกติของ GitHub

---

## 4. รันในเครื่อง (dev)

```bash
# โหมด CSV — ใช้ไฟล์ใน data/ ไม่ต้องมี credential
python scripts/build_dashboard.py --source csv

# โหมด Graph — ดึงสดจาก SharePoint
export TENANT_ID=... CLIENT_ID=... CLIENT_SECRET=...
python scripts/build_dashboard.py --source graph

# ปิดบัง PII ก่อนเขียนไฟล์ (เมื่อ deploy ที่สาธารณะ)
python scripts/build_dashboard.py --mask-pii
```

**Options ทั้งหมด**

| Option | ค่า default | ความหมาย |
|---|---|---|
| `--out` | `index.html` | ไฟล์ผลลัพธ์ |
| `--source` | `auto` | `auto` / `graph` / `csv` |
| `--data-csv` | `data/DemoApp.csv` | CSV ของ DemoApp |
| `--acl-csv` | `data/Admin_KycNew.csv` | CSV ของ Admin_KycNew |
| `--min-records` | `1` | ถ้าดึงได้น้อยกว่านี้ **ไม่เขียนทับ** `index.html` (exit 2) |
| `--mask-pii` | ปิด | ปิดบัง 7 คอลัมน์ PII ก่อนเขียนไฟล์ |

`build_dashboard.py` ใช้เฉพาะ **standard library** ของ Python 3.10+ (ไม่ต้อง `pip install`)

---

## 5. ความปลอดภัย (อ่านก่อน deploy)

| หัวข้อ | รายละเอียด |
|---|---|
| 🔴 **Repo ต้องเป็น Private** | `index.html` ฝังข้อมูลลูกค้าจริง (ชื่อ, เลขทะเบียน, เบอร์โทร, วงเงิน) — ถ้า repo เป็น public ข้อมูลจะเปิดสู่สาธารณะทันที |
| 🔴 **GitHub Pages ของ repo private ยังเป็น public** ในแพ็กเกจ Free/Pro — ต้องใช้ **GitHub Enterprise Cloud** จึงจะจำกัดผู้เข้าถึงได้ ถ้าไม่มี ให้ deploy ด้วย `--mask-pii` |
| 🟡 **Login ฝั่ง client เป็น UI-layer** | ผู้ที่ View Source ยังเห็น `RAW_DATA` ได้ — ต้องกันการเข้าถึงที่ระดับ repo/Pages ควบคู่เสมอ |
| 🟢 **ไม่มีอีเมลใน UI ล็อกอิน** | หน้าล็อกอินไม่แสดง/ไม่ autocomplete รายชื่ออีเมล ผู้ใช้ต้องพิมพ์เอง และข้อความ error ไม่บอกว่าอีเมลมีอยู่จริงหรือไม่ (กัน account enumeration) |
| 🟢 **Secret ไม่อยู่ในโค้ด** | `CLIENT_SECRET` อยู่ใน GitHub Secrets เท่านั้น ไม่ถูกพิมพ์ลง log |
| 🟡 **หมุน Client Secret** | ตั้งอายุ 6–12 เดือน และจดวันหมดอายุไว้ ถ้าหมดอายุ workflow จะ fail ที่ขั้นขอ token |

---

## 6. Troubleshooting

| อาการ | สาเหตุ / วิธีแก้ |
|---|---|
| `HTTP 401` ตอนขอ token | `TENANT_ID`/`CLIENT_ID`/`CLIENT_SECRET` ผิด หรือ secret หมดอายุ → สร้างใหม่ |
| `HTTP 403` ตอนเรียก `/sites/...` | ยังไม่ได้ **Grant admin consent** หรือใช้ `Sites.Selected` แต่ยังไม่ได้ `Grant-PnPAzureADAppSitePermission` |
| `HTTP 404` ที่ `/lists/DemoApp` | ชื่อลิสต์ไม่ตรง (ใช้ **ชื่อที่แสดง**) → ตั้ง variable `LIST_DEMOAPP` ให้ถูก |
| ได้ข้อมูลไม่ครบ | สคริปต์ไล่ `@odata.nextLink` อยู่แล้ว ถ้ายังขาด ให้ตรวจ **View Threshold** และสิทธิ์ระดับ item |
| คอลัมน์บางช่องว่างเปล่า | ชื่อ internal name เปลี่ยน → แก้ `FIELD_MAP` (สคริปต์ลอง fallback แบบไม่มี `_x0020_` ให้แล้ว) |
| `exit code 2` | ดึงได้น้อยกว่า `--min-records` → ระบบ **ไม่เขียนทับ** `index.html` เดิมโดยตั้งใจ |
| Pages ไม่อัปเดต | ตรวจ Settings → Pages → Source ต้องเป็น **GitHub Actions**; และดู job `deploy` ใน Actions |
| หน้าเว็บขึ้น badge **OFFLINE** | ปกติเมื่อเปิดนอก SharePoint — ข้อมูลมาจาก snapshot ที่ workflow สร้าง (ดูเวลาที่ `srcDetail`) |

---

## 7. หมายเหตุการแสดงผล

`index.html` รองรับ 2 โหมด

| ที่ตั้ง | โหมด | ที่มาข้อมูล |
|---|---|---|
| GitHub Pages | `SNAPSHOT` | ข้อมูลที่ workflow ดึงมาตอน build (มี timestamp บนหน้าจอ) |
| ฝังใน SharePoint (Embed WebPart ไซต์เดียวกัน) | `LIVE` | ยิง REST ตรงไปที่ลิสต์ `DemoApp` แบบ no-cache + ปุ่มรีเฟรช + auto-refresh 5 นาที |

ปรับค่าได้ที่หัวสคริปต์ใน `scripts/template.html` ส่วน `[1] DATA BINDING`:
`LIVE_MODE` (`"auto"`/`true`/`false`) · `SITE_URL` · `LIST_TITLE` · `PAGE_SIZE_API` · `AUTO_REFRESH_MIN`
