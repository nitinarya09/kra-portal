# 👤 User Directory & Access Control Matrix
**Office of the Accountant General (A&E), Madhya Pradesh, Gwalior**

This document contains the complete official directory of all **59 pre-seeded user accounts**, role hierarchies, privilege levels, data boundaries, and security guidelines for the KRA Portal.

---

## 🛡️ Privilege Levels & Role Hierarchy

The portal defines **4 Privilege Levels** mapped to **7 Access Roles**:

```
Level 6: DEVELOPER     ──► System Administrator (Full System Control)
Level 5: ADMIN         ──► Accountant General / Dy. AG (Office Leadership)
Level 4: TM_ADMIN      ──► TM Section Admin (Quarter Control & Consolidation)
Level 4: ITA_OFFICER   ──► Internal Audit Officer (ITA Review & Certificate)
Level 3: SECTION_HEAD  ──► Sr. Accounts Officer (Section Sign-off)
Level 2: REVIEWER      ──► Asst. Accounts Officer (Section L1 Review & Return)
Level 1: OPERATOR      ──► Accountant / DEO (Section Data Entry)
Level 0: VIEWER        ──► View-Only Observer (Auditor / Monitoring)
```

### Privilege Matrix Summary

| Privilege Level | Access Scope | Accessible Tabs | Editable Data |
|:---|:---|:---|:---|
| 👑 **Full Control** | All 17 Sections | Live Monitor, Submit, Review, Trends, Templates, History, Administration | All sections, User CRUD, Quarter Lock, ITA Cert, Document Compile |
| 🛡️ **Standard User** | Assigned Section(s) | Live Monitor, Review, Trends, Templates, History | Review & Approve/Return submissions for assigned section |
| 📝 **Limited User** | Assigned Section(s) | Live Monitor (Assigned), Submit, Trends, Templates (Assigned), History | Webform & Excel data entry for assigned section only |
| 👁️ **View-Only** | All 17 Sections | Live Monitor, Trends, Templates, History | Read-only observation (No editing or submission rights) |

---

## 📋 1. Executive Leadership & Administrative Accounts

| Username | Password | Display Name | Privilege Level | Role | Designation | Assigned Sections |
|:---|:---|:---|:---|:---|:---|:---|
| `admin` | **`agmpadmin`** | System Administrator | Full Control | `DEVELOPER` | Developer | `ALL` |
| `agmp` | **`agmp`** | Accountant General MP | Full Control | `ADMIN` | Accountant General | `ALL` |
| `dagaccounts` | **`dagaccounts`** | Dy. Accountant General (Accounts) | Full Control | `ADMIN` | Dy. Accountant General | `ALL` |
| `tm_admin` | **`admin123`** | TM Section Admin | Full Control | `TM_ADMIN` | Sr. Accounts Officer | `ALL` |
| `kra_viewer` | **`viewer123`** | KRA View-Only Auditor | View-Only | `VIEWER` | View-Only Auditor | `ALL` |

---

## 🔍 2. ITA Section Accounts (Internal Audit & Certificate Control)

ITA Section officers have office-wide viewing rights and dedicated authority to issue official **ITA Audit Certificates** with custom audit remarks and objections.

| Username | Password | Display Name | Privilege Level | Role | Designation | Assigned Sections |
|:---|:---|:---|:---|:---|:---|:---|
| `ita` | **`ita`** | ITA Section Officer | Full Control | `ITA_OFFICER` | ITA Officer | `ALL` |
| `aaoita` | **`aaoita`** | Asst. Accounts Officer (ITA) | Full Control | `ITA_OFFICER` | Asst. Accounts Officer (ITA) | `ALL` |
| `sraoita` | **`sraoita`** | Sr. Accounts Officer (ITA) | Full Control | `ITA_OFFICER` | Sr. Accounts Officer (ITA) | `ALL` |

---

## 📂 3. Section Accounts (51 Accounts across 17 Sections)

Each of the 17 operational sections has 3 dedicated accounts:

### Section Accounts Reference Table

| Section Name | Code | Operator (Accountant / DEO) | Reviewer (AAO) | Section Head (Sr. AO) |
|:---|:---|:---|:---|:---|
| **Book-1** | `BOOK1` | **ID:** `book1`<br>**Pass:** `book1` | **ID:** `aaobook1`<br>**Pass:** `aaobook1` | **ID:** `sraobook1`<br>**Pass:** `sraobook1` |
| **Book-2** | `BOOK2` | **ID:** `book2`<br>**Pass:** `book2` | **ID:** `aaobook2`<br>**Pass:** `aaobook2` | **ID:** `sraobook2`<br>**Pass:** `sraobook2` |
| **Reconciliation Cell** | `RC_CELL` | **ID:** `rc_cell`<br>**Pass:** `rc_cell` | **ID:** `aaorc_cell`<br>**Pass:** `aaorc_cell` | **ID:** `sraorc_cell`<br>**Pass:** `sraorc_cell` |
| **Reserve Bank Deposits** | `RBD` | **ID:** `rbd`<br>**Pass:** `rbd` | **ID:** `aaorbd`<br>**Pass:** `aaorbd` | **ID:** `sraorbd`<br>**Pass:** `sraorbd` |
| **Checking & Validation Cell** | `CV_CELL` | **ID:** `cv_cell`<br>**Pass:** `cv_cell` | **ID:** `aaocv_cell`<br>**Pass:** `aaocv_cell` | **ID:** `sraocv_cell`<br>**Pass:** `sraocv_cell` |
| **Treasury Misc.** | `TM` | **ID:** `tm`<br>**Pass:** `tm` | **ID:** `aaotm`<br>**Pass:** `aaotm` | **ID:** `sraotm`<br>**Pass:** `sraotm` |
| **ACD Section** | `ACD` | **ID:** `acd`<br>**Pass:** `acd` | **ID:** `aaoacd`<br>**Pass:** `aaoacd` | **ID:** `sraoacd`<br>**Pass:** `sraoacd` |
| **Compilation & Treasury** | `CT` | **ID:** `ct`<br>**Pass:** `ct` | **ID:** `aaoct`<br>**Pass:** `aaoct` | **ID:** `sraoct`<br>**Pass:** `sraoct` |
| **Forest Compilation** | `FC` | **ID:** `fc`<br>**Pass:** `fc` | **ID:** `aaofc`<br>**Pass:** `aaofc` | **ID:** `sraofc`<br>**Pass:** `sraofc` |
| **Bhopal Branch** | `BHOPAL` | **ID:** `bhopal`<br>**Pass:** `bhopal` | **ID:** `aaobhopal`<br>**Pass:** `aaobhopal` | **ID:** `sraobhopal`<br>**Pass:** `sraobhopal` |
| **Loans Section** | `LOAN` | **ID:** `loan`<br>**Pass:** `loan` | **ID:** `aaoloan`<br>**Pass:** `aaoloan` | **ID:** `sraoloan`<br>**Pass:** `sraoloan` |
| **TI Cell** | `TI_CELL` | **ID:** `ti_cell`<br>**Pass:** `ti_cell` | **ID:** `aaoti_cell`<br>**Pass:** `aaoti_cell` | **ID:** `sraoti_cell`<br>**Pass:** `sraoti_cell` |
| **Fund Section** | `FUND` | **ID:** `fund`<br>**Pass:** `fund` | **ID:** `aaofund`<br>**Pass:** `aaofund` | **ID:** `sraofund`<br>**Pass:** `sraofund` |
| **Deposit Section** | `DEPOSIT` | **ID:** `deposit`<br>**Pass:** `deposit` | **ID:** `aaodeposit`<br>**Pass:** `aaodeposit` | **ID:** `sraodeposit`<br>**Pass:** `sraodeposit` |
| **Admin Section** | `ADMIN` | **ID:** `sec_admin`<br>**Pass:** `sec_admin` | **ID:** `aaosec_admin`<br>**Pass:** `aaosec_admin` | **ID:** `sraosec_admin`<br>**Pass:** `sraosec_admin` |
| **VLC Section** | `VLC` | **ID:** `vlc`<br>**Pass:** `vlc` | **ID:** `aaovlc`<br>**Pass:** `aaovlc` | **ID:** `sraovlc`<br>**Pass:** `sraovlc` |
| **Report Section** | `REPORT` | **ID:** `report`<br>**Pass:** `report` | **ID:** `aaoreport`<br>**Pass:** `aaoreport` | **ID:** `sraoreport`<br>**Pass:** `sraoreport` |

---

## 🔒 User Administration & Security Guidelines

1. **Changing Passwords:**
   - Users can change their password under account settings, or an Administrator can reset any user password via **Administration ➔ User Management ➔ 🔑 Reset Pass**.

2. **Editing User Permissions:**
   - Administrators can edit any existing user account via **Administration ➔ User Management ➔ ✏️ Edit**.
   - Modifications save immediately to the Google Sheets `Users` database tab.

3. **Re-Seeding Default Accounts:**
   - If user accounts are accidentally deleted or reset, an Administrator can click **`⚡ Seed All Official Accounts`** in User Management to automatically restore all 59 official accounts.
