# ⚙️ System Working Note & Operational Workflow
**Office of the Accountant General (A&E), Madhya Pradesh, Gwalior**

---

## 📑 1. Operational Overview & Purpose

The **KRA Automation Tool** automates the end-to-end quarterly Key Result Area (KRA) reporting cycle for the **Office of the Accountant General (A&E), MP, Gwalior**.

It replaces manual paper circulars and offline Word collection with:
- **Centralized Data Capture:** 17 office sections submit webforms or section `.xlsx` files.
- **Role-Based Supervision:** Multi-level supervisory review and approval chain.
- **Audit Certification:** Automated Hindi Internal Audit (ITA) Certificate generation.
- **Master Report Compilation:** One-click cloud generation of the official **Consolidated KRA Report (.docx)**.

---

## 🔄 2. Complete 6-Stage Approval & Review Lifecycle

```
[Stage 1: Accountant / DEO] ──► Fills Webform or Uploads Excel ──► Status: SUBMITTED
                                                                          │
[Stage 2: Section AAO] ────────◄────── Returns with Remarks ──────────────┤
         │                                                                │
         └────────────────────────── Approves ──────────────────────────► Status: APPROVED_L1
                                                                          │
[Stage 3: Section Sr. AO] ──────────────── Approves ──────────────────► Status: APPROVED_L2
                                                                          │
[Stage 4: TM Section] ──────────────────── Consolidates 17 Sections ────► Status: TM_ACCEPTED
                                                                          │
[Stage 5: ITA Section] ─────────────────── Issues Audit Certificate ────► Status: ITA_REVIEWED
                                                                          │
[Stage 6: Dy. AG / AG] ────────────────── Final Sign-off ───────────────► Status: FINALIZED & .DOCX COMPILED
```

### Stage Details

1. **Stage 1 (Section Data Entry - Accountant/DEO):**
   - Accountant logs in (e.g. `book1`), opens **Submit KRA Data**.
   - The dropdown automatically locks to `Book-1 (BOOK1)`.
   - Accountant fills data or uploads `BOOK1_Template.xlsx` and submits.
   - Status set to **`SUBMITTED`**.

2. **Stage 2 (Section Review - Asst. Accounts Officer / AAO):**
   - AAO logs in (e.g. `aaobook1`), opens **Review & Approve**.
   - **If Valid:** Clicks **Approve** (Status updates to **`APPROVED_L1`**).
   - **If Errors:** Clicks **Return**, types remarks. Status changes to **`RETURNED`**, enabling the Accountant to re-edit and re-submit.

3. **Stage 3 (Section Sign-off - Sr. Accounts Officer / Sr. AO):**
   - Sr. AO logs in (e.g. `sraobook1`) and approves. Status updates to **`APPROVED_L2`**.

4. **Stage 4 (Office Consolidation - Treasury Misc. Section):**
   - TM Section tracks all 17 sections on the **Live Monitor**.
   - Once all 17 sections reach `APPROVED_L2`, TM Section accepts compilation (Status -> **`TM_ACCEPTED`**).

5. **Stage 5 (Audit Certification - Internal Audit / ITA Section):**
   - ITA Officer logs in (e.g. `ita`), reviews historical trends and section submissions.
   - Types audit findings into **ITA Audit Remarks & Objections**, then clicks **Issue Certificate**. Status updates to **`ITA_REVIEWED`**.

6. **Stage 6 (Final Approval & Report Generation - AG / Dy. AG):**
   - Dy. AG / AG approves, and TM Section clicks **Compile Master KRA Report (.docx)**.
   - The cloud compiler generates the final official Word document formatted according to CAG guidelines.

---

## 🗄️ 3. Database Architecture (Google Sheets & Apps Script)

The backend database operates on Google Sheets, exposing a secure JSON API via Google Apps Script.

### Google Sheets Database Tabs

1. **`Submissions`**: Stores all 17 section form row data, FY, Quarter, Status, ReviewedBy, ReviewDate, ReturnComments.
2. **`Users`**: Stores Usernames, SHA-256 password hashes, Roles, Assigned Sections, Designations, Active Status.
3. **`Sessions`**: Stores active 24-hour authentication session tokens.
4. **`Quarter_Settings`**: Stores quarter lock status (`IsLocked`), locking user, deadline date.
5. **`ITA_Certificates`**: Stores official quarter certificate records, issued date, officer ID, and ITA Audit Remarks & Objections.
6. **`Audit_Log`**: Stores full audit logs of login attempts, data updates, status changes, and administrative actions.

---

## 🐍 4. Python Document Compilation Engine

The Master Report generation engine uses Python (`python-docx`) to compile data into the official Word document layout.

### Engine Workflow:
1. `cloud_server.py` receives a POST request from the web portal.
2. `consolidate.py` fetches the latest approved data across all 17 sections from Google Sheets.
3. `populate_template.py` loads the official master Word template (`Master_KRA_Template.docx`), populates header titles, annexures, summary tables, and section tables.
4. Returns the compiled `.docx` report for instant download.

---

## 📈 5. Trend Analytics Engine

The portal includes an interactive analytics module built with Chart.js:
- **Tracked Parameters:** 29 measurable indicators (MCA rendition delays, exclusions, suspense clearance %, voucher validation %, audit objections, broadsheet reconciliations).
- **Visualization Modes:** Bar charts, Line trends, Radar profiles, Pie distributions.

---

## 🔒 6. Security & Control Enforcements

- **Password Hashing:** Passwords are hashed client-side using SHA-256 (Web Crypto API) before transmission.
- **Section Boundary:** Section dropdowns and Excel templates are restricted by `visibleSections` logic.
- **Quarter Locks:** Locked quarters prevent any new data submissions or edits.
