# KYC Dashboard — Auto Update via GitHub Actions

Dashboard ข้อมูล KYC ของ DoHome  
อัปเดตอัตโนมัติทุกวัน 07:00 น. เวลาไทย ผ่าน GitHub Actions

## 🌐 URL
https://dohomepublic.github.io/sharepoint-web/

## 🔄 วิธีทำงาน
1. GitHub Actions รันทุกวัน 07:00 น. (cron: `0 0 * * *`)
2. ดึงข้อมูลจาก SharePoint KYCData1 ผ่าน Microsoft Graph API
3. สร้าง `index.html` ใหม่พร้อมข้อมูลล่าสุด
4. Commit + Push → GitHub Pages อัปเดตอัตโนมัติ

## ⚙️ การตั้งค่า GitHub Secrets

ไปที่ `Settings > Secrets and variables > Actions` แล้วเพิ่ม:

| Secret | Value |
|--------|-------|
| `AZURE_CLIENT_ID` | `012ac5e6-9487-4436-9e0e-246c19ab2a67` |
| `AZURE_TENANT_ID` | `7f8918d9-718a-495b-ac9a-17cba381c4a0` |
| `AZURE_CLIENT_SECRET` | (ค่า Client Secret จาก Azure AD) |

## 🔑 การตั้งค่า Azure AD (IT Admin ทำ 1 ครั้ง)

1. เปิด [Azure Portal](https://portal.azure.com)
2. ไปที่ **Azure Active Directory > App registrations**
3. เปิด App ID: `012ac5e6-9487-4436-9e0e-246c19ab2a67`
4. **Certificates & secrets** → New client secret → Copy value
5. **API permissions** → Add permission → Microsoft Graph → Application permissions
   - เพิ่ม `Sites.Read.All`
   - กด **Grant admin consent**

## 📁 โครงสร้างไฟล์

```
sharepoint-web/
├── .github/
│   └── workflows/
│       └── update-dashboard.yml   ← GitHub Actions workflow
├── scripts/
│   └── build_dashboard.py         ← Python script ดึงข้อมูล + สร้าง HTML
├── index.html                     ← Dashboard (auto-generated)
└── README.md
```

## 🖱️ รันด้วยตนเอง
ไปที่ **Actions** tab → เลือก "Update KYC Dashboard" → กด **Run workflow**
