# Bot Export Scada & JMS
<img width="1424" height="856" alt="image" src="https://github.com/user-attachments/assets/72501792-87b0-4f41-ad33-db5bbeacccab" />

Python desktop application สำหรับ export และ download report จากระบบ DWS Scada และ JMS ผ่าน API อัตโนมัติ พร้อมระบบ Scheduler, Multi-Task Export, GUI Monitoring และ File Download Automation

โปรแกรมรองรับการทำงานแบบ manual และ auto schedule สำหรับดึงไฟล์ report จากหลายระบบพร้อมกัน และสามารถ build เป็น `.exe` สำหรับใช้งานจริงบน Windows ได้

---

# Features

- Auto scheduler รายชั่วโมง
- Manual Run
- Multi-system export
- Download report ผ่าน API
- รองรับ DWS และ JMS
- Auto polling download status
- GUI monitoring
- Real-time log viewer
- Config persistence (`config.ini`)
- Multi-threaded worker
- Stop running task ได้
- Auto next-run calculation
- File auto save
- รองรับ build เป็น Windows EXE

---

# Application Overview

โปรแกรมนี้ถูกออกแบบมาเพื่อ automate การ export report จาก:

- DWS Scada System
- JMS Operating Platform

โดยระบบจะ:

1. Request report generation
2. Poll download queue
3. ตรวจสอบว่าไฟล์พร้อม download หรือไม่
4. Download file อัตโนมัติ
5. Save ลง output folder

เหมาะสำหรับ:

- Warehouse Automation
- Logistics Monitoring
- DWS Report Collection
- JMS Report Export
- Operation Dashboard Integration

---

# Supported Systems

## DWS9-11

ใช้ API สำหรับ export report จาก DWS Scada

Export file:

```text
DWS9-11.xlsx
```

---

## JMS Auto PDA

ใช้ API สำหรับ export:

```text
建包扫描
```

Export file:

```text
DWSXAUTOPDA.xlsx
```

---

## JMS DWS PDA

ใช้ API สำหรับ export:

```text
卸车扫描
```

Export file:

```text
DWSPDA.xlsx
```

---

# Tech Stack

- Python
- PySide6
- Requests
- Threading
- ConfigParser

---

# Project Structure

```text
project/
│
├── Main.py
├── config.ini
│
├── output/
│   ├── DWS9-11.xlsx
│   ├── DWSXAUTOPDA.xlsx
│   └── DWSPDA.xlsx
│
└── logs/
```

---

# GUI Overview

โปรแกรมแบ่งออกเป็น 2 Tabs

## Home Tab

ใช้สำหรับ:

- Start / Stop Scheduler
- Run Manual Export
- ตั้ง Time Range
- ดู Log
- ดู Status

---

## Setting Tab

ใช้สำหรับ:

- ตั้ง API URL
- ตั้ง Token
- ตั้ง Output Path
- ตั้งชื่อไฟล์ export
- ตั้ง Run Minute

---

# Scheduler Logic

ระบบ scheduler จะทำงานทุกชั่วโมงตาม minute ที่กำหนด

ตัวอย่าง:

```text
Run minute = 5
```

ระบบจะทำงานเวลา:

```text
10:05
11:05
12:05
13:05
```

---

# Export Workflow

## Full Workflow

```text
Start Scheduler
    ↓
Wait Until Run Minute
    ↓
Run DWS Export
    ↓
Wait 10 Seconds
    ↓
Run JMS Auto PDA Export
    ↓
Wait 10 Seconds
    ↓
Run JMS DWS PDA Export
    ↓
Download Complete
```

---

# JMS Download Logic

ระบบ JMS ใช้วิธี:

1. Request export task
2. รอ server generate file
3. Poll download center
4. ตรวจสอบว่า file ready หรือไม่
5. Request signed URL
6. Download file

---

# Smart File Detection

ระบบตรวจสอบ:

- finishOrNot == 1
- มี download URL
- downTime ตรงกับเวลาที่ request
- file size มากกว่า 100 KB

เพื่อป้องกัน download ไฟล์ผิดหรือไฟล์ยัง generate ไม่เสร็จ

---

# Time Range System

รองรับการกำหนด:

- Start Date
- Start Hour
- End Date
- End Hour

ระบบจะ auto swap เวลาให้หาก:

```text
Start > End
```

---

# Threading Design

โปรแกรมใช้:

```python
threading.Thread()
```

เพื่อแยก:

- GUI Thread
- Scheduler Thread
- Export Worker Thread

ข้อดี:

- GUI ไม่ค้าง
- Stop task ได้
- Scheduler ทำงาน background ได้
- รองรับ long-running export

---

# Config System

ระบบจะ save config อัตโนมัติลง:

```text
config.ini
```

---

# Example Config

```ini
[API]
dws_url=http://example.com/api
dws_token=YOUR_DWS_TOKEN
jms_token=YOUR_JMS_TOKEN

[PATH]
path=C:/EXPORT

[FILE]
name_dws=DWS9-11.xlsx
name_auto=DWSXAUTOPDA.xlsx
name_dwspda=DWSPDA.xlsx

[TIME]
run_minute=5
```

---

# Logging System

โปรแกรมมี real-time log viewer ภายใน GUI

ตัวอย่าง:

```text
10:05:00 [SYS] Auto scheduler started
10:05:01 [DWS9-11] Start DWS request...
10:05:15 [JMS] Generate DWSXAUTOPDA.xlsx
10:06:20 [JMS] Downloaded DWSPDA.xlsx
```

---

# Error Handling

ระบบรองรับ:

- API timeout
- Invalid token
- File generation timeout
- Download failure
- Invalid time range
- User stop request

เมื่อเกิด error:

- แสดงใน log
- update status
- stop current task

---

# UI State Management

ระหว่าง export:

- disable input fields
- lock config editing
- ป้องกัน run ซ้ำ
- อนุญาตให้ stop task ได้

---

# Installation

## 1. Clone Repository

```bash
git clone https://github.com/yourname/bot-export-scada-jms.git
```

---

## 2. Install Dependencies

```bash
pip install requests tkcalendar
```

---

## 3. Run Application

```bash
python Main.py
```

---

# Build EXE

ใช้ PyInstaller:

```bash
pyinstaller --onefile --windowed Main.py
```

หรือ:

```bash
pyinstaller --onefile --noconsole Main.py
```

---

# API Headers

## DWS API

ใช้:

```text
Content-Type: application/json;charset=UTF-8
token: YOUR_TOKEN
```

---

## JMS API

ใช้:

```text
Content-Type: application/json;charset=UTF-8
authtoken: YOUR_TOKEN
```

---

# Output Files

ระบบจะ save file ลง folder ที่กำหนด

ตัวอย่าง:

```text
C:/EXPORT/
├── DWS9-11.xlsx
├── DWSXAUTOPDA.xlsx
└── DWSPDA.xlsx
```

---

# Future Improvements

- Auto retry system
- Multi profile config
- Tray icon support
- Email notification
- Download history
- API health checker
- Parallel export worker
- Excel validation system
- Report merge system

---

# License

MIT License

---

# Author

Developed for warehouse operation automation and report export workflow integration.


