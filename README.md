# ⚡ MCO Maker

> **Automated Electricity Meter Change Order (MCO) Generator & PITC Bill Scraper**  
> Built for Sub Divisions and Regional Offices (RO) of Pakistan Power Distribution Companies (PESCO, LESCO, GEPCO, FESCO, IESCO, MEPCO, SEPCO, HESCO, QESCO, TESCO).

---

## 🌟 Overview

**MCO Maker** streamlines and automates the creation of official **Meter Change Orders (MCOs)**. When power utilities replace traditional meters with modern smart/AMR meters, administrative staff traditionally spent hours manually finding consumer bills, copying billing data, and typing MCO forms.

With **MCO Maker**:
1. You provide an Excel workbook containing consumer reference numbers and new meter details.
2. The application automatically queries the official PITC online billing system and extracts:
   - **Consumer Name & Address**
   - **Feeder Code & Feeder Name**
   - **Sub Division Name**
   - **Removed Meter Number (Old MSN)**
   - **Present Meter Reading (Final Reading)**
   - **Consumer Tariff Code**
3. Caches scraped data locally in `%LOCALAPPDATA%\MCO Maker\Temp\scraped_records.json`.
4. Instantly generates official **Sub Division MCOs** and **RO Office MCOs (CP-FORM-138)** in **Microsoft Word (.docx)** format!

---

## 📸 Key Features

- **VS Code Dark+ Palette & Windows 10/11 Immersive Dark Titlebar**: Clean, distraction-free interface matching developer-grade tools.
- **Smart Column Detection**: Automatically detects column names or falls back to cell values:
  - **Reference Number (Account No)**: Column A (e.g. `01269310099060`)
  - **MSN (Installed Meter No)**: Header matching `MSN` or `Meter No`
  - **SIM Number**: Detects Pakistani mobile/SIM numbers (`03142664719`, `3142664719`, `92320...`)
  - **SIM ID**: Detects ICCID numbers > 11 figures (e.g. `899204560526267198`)
  - **Meter Installation Date**: Any date format (`2026-09-25`, `25-09-2026`, `25-Sep-2026`)
  - **Telco**: Network operator (`Zong`, `Jazz`, `Telenor`, `Ufone`, `SCOM`)
  - **SCO Number / Book Number**: Order/book number (e.g. `33/CPM`, `01TH 06`)
- **Minimalist Clean Logging**: During scraping, the log window displays only clean progress:
  ```text
  1/20 Information from Bill fetched
  2/20 Information from Bill fetched
  ...
  ```
- **Colored Status Banner**: Displays a bright green confirmation box `✔ Information Fetched` upon completion.
- **Two Distinct MCO Formats**:
  1. **For Sub Division**: Generates Sub Division Meter Change Order with installed and removed meter grids, SDO signature lines, and SIM credentials.
  2. **For RO Office**: Generates Revenue Office CP-FORM-138 with Batch #, Sub Division, Account Number, Transaction Type (1- REPLACEMENT), Feeder Code, and audit signatures.
- **Word Export**: Generates clean, formatted **Microsoft Word (.docx)** documents ready for printing.
- **Hardware Binding & Licensing**: BIOS UUID & Windows MachineGuid hardware locking with instant startup (<50ms).
- **Multi-OS Support**: Native builds for both **Windows 10/11 (64-bit)** and **Windows 7 / Legacy (32-bit with bundled VC++ runtime)**.

---

## 📁 Project Structure

```
MCOMaker/
├── .github/
│   └── workflows/
│       └── build.yml               # Multi-OS CI/CD Pipeline (Win 10/11 & Win 7)
├── assets/
│   ├── app.ico                     # Multi-resolution application icon
│   ├── app.png                     # Window header icon
│   └── mcomaker.ico                # Installer icon
├── license_server/                 # Vercel Serverless Microservice
│   ├── api/index.js                # Vercel entrypoint
│   ├── server.js                   # Express API (License check, updates, admin)
│   ├── schema.sql                  # Supabase PostgreSQL schema with mco_ prefix
│   ├── package.json
│   ├── vercel.json
│   └── .env.example
├── scripts/
│   └── compile_nuitka.py           # Native C++ binary compilation (.pyd)
├── templates/
│   ├── sub_division_mco_template.docx
│   └── ro_mco_template.docx
├── installer.iss                   # Inno Setup professional installer script
├── MCOMaker.spec                   # PyInstaller specification
├── launcher.py                     # Bootstrap entry point for compiled binary
├── main.py                         # Complete Desktop GUI, Licensing client & Core logic
├── generate_key.py                 # Admin Serial Key Generator tool (Tkinter GUI & CLI)
├── generate_key.spec               # PyInstaller spec for Admin tool
├── release.py                      # One-command automated release publisher
├── requirements.txt                # Python dependencies
├── SUPABASE_QUERY.sql              # Supabase table creation & live user monitoring queries
└── VERCEL_PROCEDURE.md             # Step-by-step Vercel deployment guide
```

---

## 🚀 Quick Start (Development)

### 1. Prerequisites
- Python 3.8 to 3.12 (or 3.14)
- Git

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

### 3. Run Application
```bash
python main.py
```

---

## 🔑 Supabase Database Setup

MCO Maker uses Supabase Free Tier with table prefixing (`mco_`), allowing it to run alongside other tools (e.g. Bill Scrapper) inside a single Supabase project for $0 forever!

1. Open **Supabase Dashboard** -> **SQL Editor**.
2. Run the script in `SUPABASE_QUERY.sql` (Tab 1: Initial Setup Script).
   This creates:
   - `mco_licenses` (serial keys & subscriptions)
   - `mco_license_devices` (registered computer HWIDs)
   - `mco_app_updates` (version updates & download links)
3. Save the live query from `SUPABASE_QUERY.sql` in a second tab (`Users MCOMaker`) to monitor live activations!

---

## 🌐 Deploy License Server to Vercel

Follow the step-by-step guide in [`VERCEL_PROCEDURE.md`](VERCEL_PROCEDURE.md):
1. Push this repository to GitHub.
2. In Vercel, import the repo and select root directory `license_server`.
3. Add Environment Variables:
   - `SUPABASE_URL`
   - `SUPABASE_SERVICE_ROLE_KEY`
   - `ADMIN_SECRET`
   - `TABLE_LICENSES=mco_licenses`
   - `TABLE_DEVICES=mco_license_devices`
   - `TABLE_UPDATES=mco_app_updates`
   - `KEY_PREFIX=MCO`
4. Deploy!

---

## 🔑 Admin Serial Key Generator

Run the owner admin tool to generate subscription keys:
```bash
python generate_key.py
```
- Select plan: **6 Months (Rs. 2,500)** or **1 Year (Rs. 5,000)**
- Click **🔑 Generate Serial Key & Copy**
- The key (e.g. `MCO-9H3K-8W2Y-X7RM`) is automatically saved in Supabase and copied to your clipboard!

---

## 📦 Building & Releases

### Local Compilation (Nuitka C++ Binary Obfuscation)
```bash
python scripts/compile_nuitka.py
pyinstaller MCOMaker.spec --clean
```

### One-Command Release Publisher
```bash
python release.py 1.0.1 -m "Added support for high-resolution PDF generation"
```
This automatically:
1. Validates code syntax
2. Bumps `APP_VERSION` in `main.py` and `installer.iss`
3. Commits and pushes tag `v1.0.1` to GitHub
4. Triggers GitHub Actions CI/CD to build both `MCOMaker-Setup.exe` (Win 10/11) and `MCOMaker-Win7-Setup.exe` (Win 7 x86) and publish GitHub Releases!

---

## 📄 License & Attribution
Developed by **TN Developers Co.**  
For support and inquiries:
- **WhatsApp:** +92 333 9962474
- **Email:** tahir.naqash4@gmail.com
