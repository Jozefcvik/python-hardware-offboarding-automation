# IT Hardware Offboarding Automation

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)

Automates querying assigned IT hardware from a SQL Server database, generating per-employee reports, and emailing Excel attachments to employees with an inline summary.

---

## 📑 Table of Contents

1. [Python Modules Used](#🐍-python-modules-used)
2. [Configuration](#⚙️-configuration)
3. [How It Works](#📝-how-it-works)
   - [Helper Functions](#1-helper-functions)
   - [Main Workflow](#2-main-workflow-main)
   - [SQL Query Example](#3-sql-query-example)
   - [Email Structure](#4-email-structure)
4. [Features](#✅-features)
5. [Notes](#⚠️-notes)
6. [Usage](#📌-usage)
7. [Example Email Output](#📧-example-email-output)

---

## 🐍 Python Modules Used

| Module | Purpose |
|--------|---------|
| `pyodbc` | Connects to SQL Server and executes SQL queries. |
| `csv` | Reads employee CSV and writes per-employee/combined reports. |
| `os` | Handles file paths, directory creation, and file deletion. |
| `smtplib` | Sends emails via an SMTP server. |
| `pandas` | Reads CSVs and writes Excel files (.xlsx) with optional column autosizing. |
| `typing` | Provides type hints (`List`, `Tuple`) for clarity. |
| `email.mime.multipart` | Constructs multipart email messages (HTML + attachments). |
| `email.mime.text` | Adds HTML content to emails. |
| `email.mime.base` | Attaches binary files (Excel attachments). |
| `email.encoders` | Encodes attachments in base64 for email transport. |

---

## ⚙️ Configuration

Set these parameters at the top of the script:

- **Database Settings**: `server`, `database`, `username`, `password`
- **CSV Paths**: `INPUT_EMPLOYEES_CSV`, `COMBINED_OUTPUT_PATH`, `OUTPUT_DIR`
- **Email Settings**: `SMTP_HOST`, `SMTP_PORT`, `SENDER_EMAIL`, `RECIPIENT_DOMAIN`, `EMAIL_SUBJECT`, `ADDITIONAL_NOTE`

---

## 📝 How It Works

### 1. Helper Functions

- `_split_emails(raw: str)` – Splits multiple email addresses (`;` or `,`) into a list.  
- `read_employees(path: str)` – Reads employee CSV and returns `(GivenName, Surname, cc_emails_list)`.  
- `build_recipient_email(given: str, sur: str, domain: str)` – Constructs email address `given.surname@domain`.  
- `write_employee_csv(given, sur, col_names, rows)` – Writes per-employee CSV report.  
- `convert_csv_to_xlsx(csv_path, delimiter)` – Converts CSV to XLSX with column autosizing.  
- `append_to_combined_csv(col_names, rows, combined_path, first_write)` – Appends to combined audit CSV.  
- `build_html_table(col_names, rows)` – Generates email-friendly HTML table of hardware.  
- `send_email_with_attachment(to_email, subject, html_body, attachment_path, cc_emails)` – Sends email with XLSX attachment and optional CC.

---

### 2. Main Workflow (`main()`)

1. **Read Employees** from CSV.  
2. **Remove Previous Combined File** (optional).  
3. **Query SQL Server** for each employee's assigned hardware.  
4. **Generate Reports**:  
   - Per-employee CSV/XLSX  
   - Append to combined CSV audit file  
5. **Send Emails** with inline HTML table and XLSX attachment.  
6. **Print Summary** of emails sent and rows written.  

---

### 3. SQL Query Example

```sql
SELECT
    va.ManufacturerName,
    va.DeviceDescription,
    va.TypeDescription,
    a.Description,
    a.SerialNo,
    e.Surname,
    e.GivenName,
    e.Location,
    e.ManagerADLogin
FROM Assets AS a
JOIN AssetList AS va ON va.id = a.id
JOIN Employee AS e ON va.EmployeeId = e.id
WHERE LTRIM(RTRIM(e.GivenName)) = LTRIM(RTRIM(?))
  AND LTRIM(RTRIM(e.Surname)) = LTRIM(RTRIM(?));
``` 
- Fetches all hardware assigned to a given employee.
- Parameters `(GivenName, Surname)` are passed per employee.

---

### 4. Email Structure

- Greeting with employee name.
- Offboarding instructions.
- Inline HTML table preview of hardware.
- XLSX attachment with full details.
- Optional CC recipients from CSV.

---

### ✅ Features

- Modular and easy to customize.
- Supports multiple CC emails per employee.
- Converts CSV to Excel with readable formatting.
- Generates a combined audit CSV.
- Sends professional HTML emails with attachments.

--- 

### ⚠️ Notes

- Assumes internal SMTP relay (no authentication).
- CSV headers must include Surname, GivenName, and one of Email, EmailAddress, CC, or Cc.
- Works best with UTF-8 CSV files, including BOM-safe files from Excel.

---

### 📌 Usage

1. Configure all settings (DB, CSV paths, SMTP, etc.).
2. Place employeeInput.csv in the project folder.
3. Run the script:
```bash
python script.py
```
4. Check out/ for per-employee reports and hardwareOutput.csv for combined audit.

---

### 📧 Example Email Output

- Inline Table: Quick preview of hardware.
- Attachment: XLSX with all hardware details.
- CC Recipients: Pulled from CSV per employee.

