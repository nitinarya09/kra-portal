# 🏛️ KRA Quarterly Reporting & Consolidation Portal
**Office of the Accountant General (A&E), Madhya Pradesh, Gwalior**

![Version](https://img.shields.io/badge/Version-2.2.0-blue.svg)
![Status](https://img.shields.io/badge/Deployment-Live-success.svg)
![Platform](https://img.shields.io/badge/Platform-React%20%7C%20Google%20Apps%20Script%20%7C%20Python-orange.svg)

---

## 📌 Project Overview

The **Key Result Area (KRA) Quarterly Reporting & Consolidation Portal** is an end-to-end digital system designed for the **Office of the Accountant General (A&E), MP, Gwalior**. It streamlines, automates, and secures the quarterly KRA data collection, multi-level supervisory review, ITA certification, and master report generation across all **17 operational sections**.

### 🌐 Live System URLs
- **Web Portal (GitHub Pages):** [https://nitinarya09.github.io/kra-portal/](https://nitinarya09.github.io/kra-portal/)
- **Render Cloud Compiler Service:** `https://kra-compiler-service.onrender.com`
- **GitHub Source Repository:** [https://github.com/nitinarya09/kra-portal](https://github.com/nitinarya09/kra-portal)

---

## ✨ Core Features & Technical Highlights

1. **🔐 Multi-Level Role-Based Access Control (RBAC):**
   - 59 pre-configured accounts across 7 permission levels (`OPERATOR`, `REVIEWER`, `SECTION_HEAD`, `TM_ADMIN`, `ITA_OFFICER`, `ADMIN`, `DEVELOPER`, `VIEWER`).
   - Strict section-level data boundary enforcement (Operators only see and edit their assigned section forms and templates).

2. **⚡ Dual Data Entry Modes:**
   - **Interactive Web Forms:** Auto-calculating fields, validation checks, auto-drafting.
   - **Excel Upload:** Section-wise standardized `.xlsx` templates uploaded directly to Google Drive.

3. **🔄 6-Stage Supervisory Approval Workflow:**
   - Accountant (Submission) ➔ AAO (Level 1 Review/Return) ➔ Sr. AO (Section Approval) ➔ TM Section (Office Consolidation) ➔ ITA Section (Audit Certificate) ➔ AG/Dy. AG (Final Sign-off).

4. **📜 Auto-Generated ITA Audit Certificate:**
   - Official Hindi print-ready certificate rendered automatically with custom ITA Audit Remarks & Objections.

5. **📈 Interactive Trend Analytics:**
   - 10 chart types powered by Chart.js tracking 29 quarterly performance metrics across historical KRA reports.

6. **📄 One-Click Master Document Compilation:**
   - Python-powered cloud document compiler merges webform/Excel inputs into the official **Consolidated KRA Report (.docx)**.

---

## 📁 Repository Structure

```
KRA Automation/
├── frontend/                     # React + Vite Web Portal Source
│   ├── src/
│   │   ├── components/          # React Components (Dashboard, FormWizard, AdminPanel, UserManagement, etc.)
│   │   ├── config/              # Roles, Sections, and API Endpoints Config
│   │   └── utils/               # Auth, Chart Helpers, Draft Hooks
│   ├── dist/                    # Compiled Production Assets (Pushed to GitHub Pages)
│   └── vite.config.js           # Vite Configuration with Cache-Busting Hashing
├── apps-script/                 # Google Apps Script Backend (Database API)
│   ├── Code.gs                  # Main Router (doPost API Endpoints)
│   ├── AuthService.gs           # Authentication & Session Management
│   ├── AdminService.gs          # Quarter Lock, ITA Certs, Audit Logs
│   ├── SheetService.gs          # Submission & Sheet CRUD
│   └── DbInitializer.gs         # Database Schema & Seed Data (59 Accounts)
├── consolidation/               # Master Word Report Generation Engine
│   ├── consolidate.py           # Python Data Aggregator
│   ├── populate_template.py     # python-docx Document Compiler
│   └── cloud_server.py          # Flask HTTP Compiler API (Render Cloud)
└── docs/                        # Official Operational & User Documentation
    ├── USER_DIRECTORY_AND_CREDENTIALS.md
    └── SYSTEM_WORKING_NOTE_AND_WORKFLOW.md
```

---

## 🛠️ Quickstart for Developers & Administrators

### 1. Running Frontend Locally
```bash
cd frontend
npm install
npm run dev
```
Open `http://localhost:5173/` in your browser.

### 2. Building & Deploying Frontend to GitHub Pages
```bash
cd frontend
npm run build
powershell -Command "Copy-Item -Path 'dist/index.html' -Destination 'dist/404.html' -Force"
cd dist
git add .
git commit -m "Deploy latest build"
git push origin main --force
```

---

## 📑 Official Documentation
- 📖 [User Directory & Credentials Reference](docs/USER_DIRECTORY_AND_CREDENTIALS.md)
- ⚙️ [System Working Note & Workflow Guide](docs/SYSTEM_WORKING_NOTE_AND_WORKFLOW.md)

---
*Developed for Digital Transformation Initiative | Office of the Accountant General (A&E), MP, Gwalior*
