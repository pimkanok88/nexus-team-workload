# Team Workload & OT — Netlify Frontend + Apps Script API

เวอร์ชันนี้ **ไม่ใช้ Google Cloud Project / Service Account**

โครงสร้าง:

```
Browser
  ↓
Netlify (public/index.html)
  ↓ /api
Netlify Function (ซ่อน API secret)
  ↓
Google Apps Script Web App
  ↓
Google Sheet เดิม
```

## ไฟล์ใน ZIP

```
team_workload_netlify_apps_script_v1/
├─ public/
│  └─ index.html
├─ netlify/
│  └─ functions/
│     └─ api.js
├─ apps-script/
│  ├─ Code.gs
│  └─ appsscript.json
├─ scripts/
│  └─ deploy-netlify.ps1
├─ netlify.toml
├─ package.json
├─ .env.example
└─ README_TH.md
```

---

# A. ตั้ง Apps Script Backend

## กรณีที่ใช้ Apps Script project เดิม (แนะนำ)

1. เปิด Apps Script project เดิมของ Team Workload
2. สำรอง `Code.gs` เดิมไว้ก่อน
3. แทน `Code.gs` ด้วย `apps-script/Code.gs` ใน ZIP
4. Save
5. **ไม่ต้องรัน setupNewSystem()**
6. รันฟังก์ชัน:

```
generateNetlifyApiSecret
```

7. กด Run
8. Authorize ถ้าระบบถาม
9. เปิด **Execution log** ด้านล่าง แล้ว Copy ค่าหลัง `APPS_SCRIPT_API_SECRET=`

ตัวอย่าง:

```
aabbcc...ยาวประมาณ 64 ตัวอักษร
```

> Secret นี้ห้ามใส่ใน index.html

## ถ้าคุณสร้าง Apps Script project ใหม่

หลังวาง `Code.gs` แล้ว ให้รัน:

```
connectExistingSpreadsheet
```

แต่ Apps Script editor ไม่สามารถกรอก argument ตอนกด Run ได้โดยตรง
ดังนั้นสร้าง temporary function แบบนี้:

```javascript
function connectMySheet() {
  return connectExistingSpreadsheet('SPREADSHEET_ID_ของคุณ');
}
```

รัน `connectMySheet()` หนึ่งครั้ง แล้วลบ temporary function ได้

จากนั้นรัน `generateNetlifyApiSecret()` ต่อ

---

# B. Deploy Apps Script เป็น API

Apps Script:

1. Deploy
2. Manage deployments
3. ถ้ามี deployment เดิม ให้ Edit
4. Select type: Web app
5. Execute as: **Me**
6. Who has access: **Anyone**
7. New version
8. Deploy

Copy URL ที่ลงท้าย:

```
/exec
```

ตัวอย่าง:

```
https://script.google.com/macros/s/XXXXXXXXXXXX/exec
```

เปิด URL นี้ใน browser แล้วควรเห็น JSON ประมาณ:

```json
{
  "ok": true,
  "service": "Team Workload API"
}
```

ถ้าเห็นหน้า login แปลว่ายังไม่ได้เปิด access แบบ Anyone

> หากบัญชีองค์กรไม่อนุญาต "Anyone" วิธีนี้จะใช้ไม่ได้ ต้องให้ผู้ดูแลระบบเปิดสิทธิ์ หรือใช้ backend แบบ OAuth/Service Account แทน

---

# C. ตั้ง Netlify

สร้าง/ใช้ Netlify project เดิมก็ได้

ไปที่:

```
Project configuration
→ Environment variables
```

เพิ่ม 2 ตัว:

## 1. APPS_SCRIPT_API_URL

Value = URL `/exec` ของ Apps Script

```
https://script.google.com/macros/s/XXXXXXXXXXXX/exec
```

## 2. APPS_SCRIPT_API_SECRET

Value = secret ที่ได้จาก:

```
generateNetlifyApiSecret()
```

อย่าใส่เครื่องหมาย quote

---

# D. Deploy Netlify

เวอร์ชันนี้มี Netlify Function ดังนั้น **อย่าใช้การลากแค่ public/index.html เข้า Netlify Drop**

วิธีง่ายบน Windows:

1. แตก ZIP
2. เปิด PowerShell ใน folder หลัก
3. รัน:

```powershell
.\scripts\deploy-netlify.ps1
```

หรือ:

```powershell
npx netlify-cli login
npx netlify-cli deploy --prod --dir public --functions netlify/functions
```

ครั้งแรก Netlify จะถามว่าจะ link project ไหน
เลือก project ที่ต้องการ

หากเพิ่งตั้ง Environment Variables หลัง deploy รอบแรก ให้ deploy ซ้ำอีกครั้ง

---

# E. ทดสอบ

เปิด Netlify URL

ควรเห็น Dashboard แล้วทำตามนี้:

1. เลือกชื่อ "ผู้บันทึกข้อมูล" ที่ Sidebar
2. กด Refresh
3. ตรวจ Dashboard
4. บันทึก Member Status
5. บันทึก Consult
6. Refresh → Consult ต้องไม่เพิ่มเอง
7. เปิด งานทีม
8. เพิ่ม/แก้ Task/Subtask
9. เปิด OT
10. เปิด Audit Log

---

# Logic สำคัญที่คงไว้

## Consult

```
1 Consult = Weight 1
```

Dashboard ใช้เฉพาะ Consult ของ **สัปดาห์ปัจจุบัน จันทร์–ศุกร์**

```
Workload
= Weight งาน Active
+ Consult สัปดาห์นี้
```

Refresh เป็น read-only สำหรับ Consult
ไม่มี auto migration จาก Tasks แล้ว

## Task Weight

นับงานที่:

```
category != Consult
status != Complete
status != Cancelled
```

## Member Status

รองรับ 0 ถึง 1 เช่น:

```
1   = เต็มกำลัง
0.5 = รับงาน 50%
0   = ไม่รับงานใหม่
```

## OT

ใช้ Calendar-first UI และโหลดเมื่อเปิดเมนู OT

---

# Security

Secret อยู่ใน:

```
Netlify Environment Variables
```

Netlify Function เป็นคนส่ง secret ไป Apps Script

Browser ไม่เห็น secret

ดังนั้นห้ามแก้ frontend ให้เรียก Apps Script URL โดยตรงแล้วฝัง secret ลง JavaScript

---

# ถ้าขึ้น Error

## "Apps Script ส่งหน้า HTML กลับมาแทน API"

ตรวจ:

```
Execute as: Me
Who has access: Anyone
```

และใช้ URL `/exec` ไม่ใช่ `/dev`

## "Unauthorized API request"

ค่า:

```
APPS_SCRIPT_API_SECRET
```

ใน Netlify ไม่ตรงกับค่าที่ Apps Script เก็บไว้

ให้รัน:

```
generateNetlifyApiSecret()
```

ใหม่ แล้วเอาค่าใหม่ไปแทนใน Netlify Environment Variables จากนั้น deploy ใหม่

## "ยังไม่ได้เชื่อมฐานข้อมูล"

ถ้าใช้ Apps Script project ใหม่ ต้องรัน `connectExistingSpreadsheet(...)`

ถ้าใช้ project เดิม ปกติไม่ต้องทำ เพราะ `DB_SPREADSHEET_ID` เดิมยังอยู่ใน Script Properties

---

# Backup เดิม

`weeklyBackup()` ยังอยู่ใน Code.gs และใช้ต่อได้

ถ้าใช้ Apps Script project เดิม trigger เดิมยังอยู่

Netlify ทำหน้าที่เฉพาะ frontend/proxy
Google Sheet เดิมยังเป็นฐานข้อมูลจริง
