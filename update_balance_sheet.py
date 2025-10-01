import requests
from openpyxl import load_workbook
from datetime import datetime
import json
import re
import os

# ================= CONFIG =================
CONFIG_PATH = r"C:\Users\Macheo\OneDrive\QuickBooksAuth\config.json"
EXCEL_PATH = r"C:\Users\Macheo\OneDrive\250730 - Balance sheet Items.xlsx"
LOG_PATH = r"C:\Users\Macheo\OneDrive\QuickBooksLogs\automation_updates.log"

TOKEN_URL = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"

# ================= ACCOUNTS =================
ACCOUNTS = [
    {"sheet": "Debtors Outside Macheo", "account": "2000 · Debtors Outside Macheo"},
    {"sheet": "Debtors Employees Macheo", "account": "2001 · Debtors Employees Macheo"},
    {"sheet": "Cash in hands Employees", "account": "2002 · Cash in hands Employees"},
    {"sheet": "Prepaid Expenses", "account": "2005 · Prepaid Expenses"},  # ← will be handled like working single-sheet
    {"sheet": "Creditors Employees Macheo", "account": "2007 · Creditors Employees Macheo"},
    {"sheet": "Creditors Outside Macheo", "account": "2008 · Creditors Outside Macheo"},
    {"sheet": "Statutory Payables", "account": "2009 · Statutory Payables"},
]

# ================= UTILITIES =================
def log_message(msg):
    timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    full_msg = f"{timestamp} {msg}"
    print(full_msg)
    os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(full_msg + "\n")

def load_config():
    with open(CONFIG_PATH, "r") as f:
        return json.load(f)

def save_config(config):
    with open(CONFIG_PATH, "w") as f:
        json.dump(config, f, indent=4)

def normalize(text):
    if not text:
        return ""
    return ' '.join(str(text).strip().lower().split())

def find_description_row(sheet, description):
    """Find existing row by description; return None if not found"""
    desc_norm = normalize(description)
    for row in range(2, sheet.max_row + 1):
        if normalize(sheet.cell(row=row, column=1).value) == desc_norm:
            return row
    return None  # Do not auto-create row here

def find_month_col(sheet, month_abbr):
    month_norm = normalize(month_abbr)
    for col in range(1, sheet.max_column + 1):
        val = sheet.cell(row=1, column=col).value
        if val and normalize(str(val)).startswith(month_norm):
            return col
    return None

def update_excel(sheet, name, month_abbr, total):
    """Update Excel; if description doesn't exist, append at the bottom"""
    row = find_description_row(sheet, name)
    if not row:
        row = sheet.max_row + 1
        sheet.cell(row=row, column=1).value = name

    col = find_month_col(sheet, month_abbr)
    if not col:
        print(f"⚠️ Month '{month_abbr}' not found in sheet, skipping '{name}'")
        return None

    prev_value = sheet.cell(row=row, column=col).value or 0
    sheet.cell(row=row, column=col, value=total)
    return f"{name} - {month_abbr}: {prev_value} → {total}"

def extract_staff_name(description):
    if not description:
        return "No Description"
    desc = normalize(description)
    clean = desc.replace("overspend", "").replace("overspent", "").replace("change", "").strip()
    clean = re.sub(r"[^a-zA-Z\s]", "", clean).strip()
    return clean.title() if clean else "No Description"

# ================= QUICKBOOKS =================
def refresh_access_token():
    config = load_config()
    auth = (config["client_id"], config["client_secret"])
    payload = {"grant_type": "refresh_token", "refresh_token": config["refresh_token"]}
    headers = {"Accept": "application/json", "Content-Type": "application/x-www-form-urlencoded"}
    resp = requests.post(TOKEN_URL, auth=auth, headers=headers, data=payload)
    resp.raise_for_status()
    tokens = resp.json()
    if "refresh_token" in tokens and tokens["refresh_token"] != config["refresh_token"]:
        config["refresh_token"] = tokens["refresh_token"]
        save_config(config)
    return tokens["access_token"]

def fetch_all_journal_entries(access_token, realm_id):
    url = f"https://quickbooks.api.intuit.com/v3/company/{realm_id}/query"
    all_entries, pos, batch_size = [], 1, 1000
    while True:
        query = f"SELECT * FROM JournalEntry STARTPOSITION {pos} MAXRESULTS {batch_size}"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json",
            "Content-Type": "application/text"
        }
        resp = requests.post(url, headers=headers, data=query)
        resp.raise_for_status()
        batch = resp.json().get("QueryResponse", {}).get("JournalEntry", [])
        if not batch:
            break
        all_entries.extend(batch)
        pos += batch_size
        if len(batch) < batch_size:
            break
    return all_entries

# ================= MAIN PROCESS =================
def main():
    log_message("🔄 Starting automation...")

    access_token = refresh_access_token()
    log_message("✅ Access token refreshed")

    config = load_config()
    realm_id = config["realm_id"]

    entries = fetch_all_journal_entries(access_token, realm_id)
    log_message(f"📥 Total JournalEntries fetched: {len(entries)}")

    wb = load_workbook(EXCEL_PATH)
    month_map = {
        "jan": "Jan", "feb": "Feb", "mar": "Mar",
        "apr": "Apr", "may": "May", "jun": "Jun",
        "jul": "Jul", "aug": "Aug", "sep": "Sep",
        "oct": "Oct", "nov": "Nov", "dec": "Dec"
    }

    current_year = datetime.now().year

    for acc in ACCOUNTS:
        sheet_name, account_name = acc["sheet"], acc["account"]

        if sheet_name not in wb.sheetnames:
            log_message(f"❌ Sheet '{sheet_name}' not found, skipping.")
            continue
        sheet = wb[sheet_name]

        accum = {}
        for je in entries:
            txn_date = je.get("TxnDate")
            if not txn_date:
                continue
            txn = datetime.strptime(txn_date, "%Y-%m-%d")
            if txn.year != current_year:
                continue

            month_abbr = txn.strftime("%b").lower()
            excel_month = month_map.get(month_abbr)
            if not excel_month:
                continue

            for line in je.get("Line", []):
                acct_name = line.get("JournalEntryLineDetail", {}).get("AccountRef", {}).get("name", "")
                if normalize(account_name) not in normalize(acct_name):
                    continue

                description = line.get("Description") or "No Description"
                staff_name = extract_staff_name(description)
                amount = float(line.get("Amount", 0))

                # Special logic for Cash in Hands
                if "cash in hands" in normalize(account_name):
                    desc_norm = normalize(description)
                    if "overspend" in desc_norm or "overspent" in desc_norm:
                        amount = -abs(amount)
                    elif "change" in desc_norm:
                        amount = abs(amount)

                # Use (description, month) as key
                accum[(staff_name, excel_month)] = accum.get((staff_name, excel_month), 0) + amount

        updates = []
        for (staff_name, month), total in accum.items():
            result = update_excel(sheet, staff_name, month, total)
            if result:
                updates.append(result)

        if updates:
            log_message(f"📊 Updates for '{sheet_name}':")
            for u in updates:
                log_message("   " + u)
        else:
            log_message(f"⚠️ No updates made for '{sheet_name}' this year.")

    wb.save(EXCEL_PATH)
    log_message("✅ Excel workbook updated successfully!\n")

if __name__ == "__main__":
    main()
