
import pyodbc
import csv
import os
import smtplib
import pandas as pd
from typing import List, Tuple
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ---------------------- CONFIG ----------------------

# SQL Server connection settings
server = 'DB_SERVER'
database = 'DB_NAME'
username = 'DB_USER'
password = 'DB_PASSWORD'


# Build the connection string (expects variables: server, database, username, password)
conn_str = (
    "DRIVER={ODBC Driver 18 for SQL Server};"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"UID={username};"
    f"PWD={password};"
    "Encrypt=no;"
)

# CSV input: headers must be Surname, GivenName, and Email (order doesn’t matter; names do)
INPUT_EMPLOYEES_CSV = "employeesInput.csv"
CSV_DELIMITER = ","  # set to ';' if your CSV uses semicolons

# Combined output (audit file for all employees)
COMBINED_OUTPUT_PATH = "hardwareOutput.csv"

# Per-employee attachments folder
OUTPUT_DIR = "out"

# SMTP (internal relay, no auth)
SMTP_HOST = "smtp.example.com"
SMTP_PORT = 25
SENDER_EMAIL = "noreply@example.com"
RECIPIENT_DOMAIN = "example.com"

EMAIL_SUBJECT = "IT Equipment Return (Offboarding)"
ADDITIONAL_NOTE = (
    "Below is the current list of IT hardware assigned to your account. "
    "Kindly review and let us know if any item is missing or should be returned."
)

# ---------------------- HELPERS ----------------------
def _split_emails(raw: str) -> List[str]:
    """
    Split a raw email string into a list. Supports ';' or ',' separators.
    Trims whitespace and ignores empty entries.
    """
    if not raw:
        return []
    parts = []
    for sep in (";", ","):
        if sep in raw:
            parts = [p.strip() for p in raw.split(sep)]
            break
    if not parts:  # no separator found, single email
        parts = [raw.strip()]
    return [p for p in parts if p]


def read_employees(path: str) -> List[Tuple[str, str, List[str]]]:
    """
    Read employees from a CSV with headers 'Surname', 'GivenName', and an Email/CC column.
    Returns a list of (GivenName, Surname, cc_emails_list) tuples.

    Accepted CC header names: 'Email', 'EmailAddress', 'CC', 'Cc'
    """
    employees: List[Tuple[str, str, List[str]]] = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f, delimiter=CSV_DELIMITER)
        fieldnames = {h.strip() for h in (reader.fieldnames or [])}

        # Required name columns must exist
        required = {"Surname", "GivenName"}
        missing = required.difference(fieldnames)
        if missing:
            raise ValueError(f"Input CSV missing required column(s): {', '.join(sorted(missing))}")

        # Find CC column (prefer 'Email'; fallback to common variants)
        cc_candidates = ["Email", "EmailAddress", "CC", "Cc"]
        cc_col = next((c for c in cc_candidates if c in fieldnames), None)
        if cc_col is None:
            raise ValueError(
                "Input CSV missing the third column for CC email. "
                "Please include one of: 'Email', 'EmailAddress', 'CC', or 'Cc'."
            )

        for row in reader:
            sur = (row.get("Surname") or "").strip()
            given = (row.get("GivenName") or "").strip()
            raw_cc = (row.get(cc_col) or "").strip()

            if not given or not sur:
                continue

            cc_emails = _split_emails(raw_cc)
            employees.append((given, sur, cc_emails))

    return employees


def build_recipient_email(given: str, sur: str, domain: str) -> str:
    """
    Build GivenName.Surname@domain (lowercase, no spaces).
    """
    local_part = f"{given}.{sur}".lower().replace(" ", "")
    return f"{local_part}@{domain}"


def write_employee_csv(given: str, sur: str, col_names: List[str], rows: List[tuple]) -> str:
    """
    Write a per-employee CSV (with header). Returns the CSV file path.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    safe_name = f"{given}_{sur}".replace(" ", "_")
    csv_path = os.path.join(OUTPUT_DIR, f"hardware_{safe_name}.csv")

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(col_names)
        for row in rows:
            writer.writerow(["" if v is None else v for v in row])

    return csv_path


def convert_csv_to_xlsx(csv_path: str, delimiter: str = CSV_DELIMITER) -> str:
    """
    Convert a CSV file to XLSX (Excel) and return the XLSX file path.
    """
    xlsx_path = os.path.splitext(csv_path)[0] + ".xlsx"
    # Read CSV (delimiter-aware, BOM-safe)
    df = pd.read_csv(csv_path, delimiter=delimiter, encoding="utf-8-sig")
    # Write XLSX
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        sheet_name = "Hardware"
        df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Optional: auto-size columns based on content length
        ws = writer.sheets[sheet_name]
        for idx, col in enumerate(df.columns, start=1):
            max_len = len(str(col))
            for cell in ws.iter_cols(min_col=idx, max_col=idx, min_row=2, max_row=ws.max_row):
                for c in cell:
                    val_len = len(str(c.value)) if c.value is not None else 0
                    if val_len > max_len:
                        max_len = val_len
            ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = min(max_len + 2, 80)
    return xlsx_path


def append_to_combined_csv(col_names: List[str], rows: List[tuple], combined_path: str, first_write: bool) -> int:
    """
    Append rows to the combined CSV file. Writes header only once (first_write=True).
    Returns number of rows appended.
    """
    if not rows:
        if first_write:
            with open(combined_path, "w", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow(col_names)
        return 0

    mode = "w" if first_write else "a"
    with open(combined_path, mode, newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if first_write:
            writer.writerow(col_names)
        for row in rows:
            writer.writerow(["" if v is None else v for v in row])

    return len(rows)


def build_html_table(col_names: List[str], rows: List[tuple]) -> str:
    """
    Build a simple HTML table (with inline styles) for embedding in email bodies.
    Email-client friendly: uses border-collapse and inline styles.
    """
    def esc(x):
        s = "" if x is None else str(x)
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    thead = "".join(
        f'<th style="background:#f2f2f2; padding:6px; border:1px solid #ccc; text-align:left;">{esc(c)}</th>'
        for c in col_names
    )
    tbody_rows = []
    for r in rows:
        # pyodbc rows can be Row objects; iterate as sequence
        tds = "".join(
            f'<td style="padding:6px; border:1px solid #ccc;">{esc(v)}</td>'
            for v in r
        )
        tbody_rows.append(f"<tr>{tds}</tr>")

    table = (
        '<table style="border-collapse:collapse; width:100%;">'
        f"<thead><tr>{thead}</tr></thead>"
        f"<tbody>{''.join(tbody_rows)}</tbody>"
        "</table>"
    )
    return table


def send_email_with_attachment(
    to_email: str,
    subject: str,
    html_body: str,
    attachment_path: str,
    cc_emails: List[str] | None = None
):
    """
    Send email via internal SMTP relay (no auth) with an Excel attachment.
    Adds CC recipients if provided.
    """
    cc_emails = cc_emails or []

    message = MIMEMultipart()
    message["From"] = SENDER_EMAIL
    message["To"] = to_email
    if cc_emails:
        message["Cc"] = ", ".join(cc_emails)
    message["Subject"] = subject

    # Body
    message.attach(MIMEText(html_body, "html"))

    # Attachment (.xlsx)
    with open(attachment_path, "rb") as f:
        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(attachment_path)}"')
    message.attach(part)

    # Send (no STARTTLS, no login)
    recipients = [to_email] + cc_emails
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        try:
            server.ehlo()
        except Exception:
            pass
        server.sendmail(SENDER_EMAIL, recipients, message.as_string())


# ---------------------- MAIN ----------------------
def main():
    # Read employees
    employees = read_employees(INPUT_EMPLOYEES_CSV)
    if not employees:
        print(f"No valid employees found in {INPUT_EMPLOYEES_CSV}. Nothing to do.")
        return

    # Remove previous combined file (optional)
    if os.path.exists(COMBINED_OUTPUT_PATH):
        os.remove(COMBINED_OUTPUT_PATH)

    # SQL query with explicit JOINs
    query = """
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
        JOIN AssetList AS va
          ON va.id = a.id
        JOIN Employee AS e
          ON va.EmployeeId = e.id
        WHERE
          LTRIM(RTRIM(e.GivenName)) = LTRIM(RTRIM(?))
          AND LTRIM(RTRIM(e.Surname))   = LTRIM(RTRIM(?));
    """

    conn = pyodbc.connect(conn_str)
    try:
        cursor = conn.cursor()

        total_written = 0
        first_write_combined = True

        for given, sur, cc_emails in employees:
            # Query DB
            cursor.execute(query, (given, sur))
            rows = cursor.fetchall()
            col_names = [desc[0] for desc in cursor.description]

            # Write per-employee CSV
            csv_path = write_employee_csv(given, sur, col_names, rows)

            # Convert CSV → XLSX for attachment
            xlsx_path = convert_csv_to_xlsx(csv_path, delimiter=CSV_DELIMITER)

            # Append to combined CSV (audit)
            written = append_to_combined_csv(col_names, rows, COMBINED_OUTPUT_PATH, first_write_combined)
            if written > 0 and first_write_combined:
                first_write_combined = False
            total_written += written

            # Build hardware HTML table (for email body)
            table_html = build_html_table(col_names, rows) if rows else ""

            # Recipient + body
            to_email = build_recipient_email(given, sur, RECIPIENT_DOMAIN)
            if rows:
                body_html = f"""
                <!DOCTYPE html>
                <html>
                  <body style="font-family:Segoe UI, Arial, sans-serif; font-size:18px; color:#333; line-height:1.5;">
                    <p>Dear {given} {sur},</p>
                
                    <p>
                      For your <strong>Offboarding</strong>, please make sure to
                      <u><strong>return all your IT Equipment as seen in the screenshot below.</strong></u>
                    </p>
                
                    <p>
                      <span style="background-color: yellow;"><strong><u>Important:</u></strong></span>
                    </p>
                
                    <hr style="border:none; border-top:1px solid #ddd; margin:20px 0;">
                    <p style="margin:0 0 8px 0;"><strong>Hardware assigned to your account (inline preview):</strong></p>
                
                    {table_html}
                
                    <p style="font-size:12px; color:#666; margin-top:10px;">
                      This inline table is for a quick preview. The full details are also attached as an Excel file.
                    </p>
                
                    <p>
                      Best regards,
                	    IT Team 
                    </p>
                
                  </body>
                </html>
                """
            else:
                body_html = f"""
                <!DOCTYPE html>
                <html>
                  <body style="font-family:Segoe UI, Arial, sans-serif; font-size:14px; color:#333; line-height:1.5;">
                    <p>Hi {given} {sur},</p>
                    <p>We could not find any hardware currently assigned to your account.</p>
                    <p>If this is unexpected, please contact IT Support.</p>
                    <p>Best regards,<br/>IT Support</p>
                  </body>
                </html>
                """

            # Send email with XLSX attachment (with CC from CSV)
            try:
                send_email_with_attachment(to_email, EMAIL_SUBJECT, body_html, xlsx_path, cc_emails=cc_emails)
                print(
                    f"Sent email to {to_email} (CC: {', '.join(cc_emails) if cc_emails else 'none'}) "
                    f"from {SENDER_EMAIL} via {SMTP_HOST}:{SMTP_PORT} with {len(rows)} row(s). "
                    f"Attachment: {xlsx_path}"
                )
            except Exception as mail_ex:
                print(f"ERROR sending to {to_email}: {mail_ex}")

        if total_written == 0:
            print("No rows matched any of the provided employees.")
        else:
            print(f"\nSaved {total_written} row(s) to combined file: {os.path.abspath(COMBINED_OUTPUT_PATH)}")

    finally:
        conn.close()


if __name__ == "__main__":
    main()
