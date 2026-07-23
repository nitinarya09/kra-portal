# KRA Consolidation Engine — Setup Guide

## Prerequisites

- **Python 3.10+** installed on your system
- **Microsoft Word** (to review the generated .docx)
- Access to the [KRA_Database Google Spreadsheet](https://docs.google.com/spreadsheets/d/12YfeR994vYcuQTdvA1y-8l-dXJzX64AlP9GG8EtSEjM)

---

## Step 1: Install Python Dependencies

Open a command prompt and navigate to the consolidation folder:

```powershell
cd "d:\BUILDING and TESTING\KRA Automation\consolidation"
pip install -r requirements.txt
```

---

## Step 2: Create Google Cloud Service Account

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Click **Select a project** > **New Project**
   - Name: `KRA-Automation`
   - Click **Create**
3. Navigate to **APIs & Services** > **Library**
4. Search and enable:
   - **Google Sheets API** — click Enable
   - **Google Drive API** — click Enable
5. Navigate to **APIs & Services** > **Credentials**
6. Click **+ Create Credentials** > **Service Account**
   - Name: `kra-reader`
   - Click **Create and Continue**
   - Role: **Viewer** (or skip role assignment)
   - Click **Done**
7. Click on the created service account email
8. Go to **Keys** tab > **Add Key** > **Create new key**
   - Key type: **JSON**
   - Click **Create**
   - A file like `kra-automation-xxxx.json` will download

---

## Step 3: Place credentials.json

1. Rename the downloaded file to `credentials.json`
2. Copy it to:
   ```
   d:\BUILDING and TESTING\KRA Automation\consolidation\credentials.json
   ```

---

## Step 4: Share the Spreadsheet

1. Open `credentials.json` in a text editor
2. Find the `"client_email"` field — it looks like:
   ```
   "client_email": "kra-reader@kra-automation-12345.iam.gserviceaccount.com"
   ```
3. Copy that email address
4. Open the [KRA_Database spreadsheet](https://docs.google.com/spreadsheets/d/12YfeR994vYcuQTdvA1y-8l-dXJzX64AlP9GG8EtSEjM)
5. Click **Share** > paste the service account email > give **Viewer** access > click **Send**

---

## Step 5: Test the Connection

```powershell
cd "d:\BUILDING and TESTING\KRA Automation\consolidation"
py fetch_data.py --fy 2026-27 --quarter Q1 --dry-run
```

Expected output:
```
Using credentials: D:\...\consolidation\credentials.json
Connected to spreadsheet: KRA_Database
DRY RUN: Connection successful. Spreadsheet has 40 worksheets.
Available worksheets: Submissions, Accounts_MCA, ...
```

---

## Step 6: Run the Consolidation

```powershell
py consolidate.py --fy 2026-27 --quarter Q1
```

The compiled report will be saved to:
```
d:\BUILDING and TESTING\KRA Automation\consolidation\output\Consolidated_KRA_Q1_2026-27.docx
```

---

## Command Reference

| Command | Description |
|---|---|
| `py consolidate.py --fy 2026-27 --quarter Q1` | Full compilation |
| `py consolidate.py --fy 2026-27 --quarter Q1 --dry-run` | Test connection only |
| `py consolidate.py --fy 2026-27 --quarter Q1 --skip-note` | Skip appreciation note |
| `py consolidate.py --fy 2026-27 --quarter Q1 --template "path.docx"` | Use custom template |
| `py fetch_data.py --fy 2026-27 --quarter Q1` | Download data only |

---

## Troubleshooting

### "credentials.json not found"
Place the Google service account JSON key file in the `consolidation/` directory with filename `credentials.json`.

### "Could not open spreadsheet"
Make sure you've shared the spreadsheet with the service account email (Step 4).

### "No data found for the specified FY and Quarter"
Check that the section staff have submitted data through the web portal for the correct FY and Quarter.

### Import errors (gspread, python-docx)
Run `pip install -r requirements.txt` to install all dependencies.
