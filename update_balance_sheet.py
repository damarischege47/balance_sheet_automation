"""
QuickBooks Journal Entry Automation
-----------------------------------
This script retrieves journal entries from QuickBooks Online,
aggregates totals per account and month, and updates an Excel workbook.

All sensitive data (client IDs, secrets, paths) should be stored
in environment variables for security.
"""

import requests
from openpyxl import load_workbook
from datetime import datetime
import json
import re
import os


# ========================= CONFIG =========================
CONFIG_PATH = os.getenv("CONFIG_PATH", "config.json")
EXCEL_PATH = os.getenv("EXCEL_PATH", "BalanceSheetItems.xlsx")
LOG_PATH = os.getenv("LOG_PATH", "logs/automation_updates.log")

TOKEN_URL = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"

ACCOUNTS = [
    {"sheet": "Debtors Outside Macheo", "account": "2000 · Debtors Outside Macheo"},
    {"sheet": "Debtors Employees Macheo", "account": "2001 · Debtors Employees Macheo"},
    {"sheet": "Cash in hands Employees", "account": "2002 · Cash in hands Employees"},
    {"sheet": "Prepaid Expenses", "account": "2005 · Prepaid Expenses"},
    {"sheet": "Creditors Employees Macheo", "account": "2007 · Creditors Employees Macheo"},
    {"sheet": "Creditors Outside Macheo", "account": "2008 · Creditors Outside Macheo"},
    {"sheet": "Statutory Payables", "account": "2009 · Statutory Payables"},
]


# ========================= UTILITIES =========================
def log_message(msg: str):
    """Write a log entry with a timestamp."""
    timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    message = f"{timestamp} {msg}"
    print(message)
    os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(message + "\n")


def load_config():
    """Load QuickBooks OAuth credentials from JSON file."""
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(config):
    """Save updated tokens to config."""
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4)


def normalize(text):
    """Normalize text for consistent comparison."""
    if not text:
        return ""
    return " ".join(str(text).strip().lower().split())


def find_description_row(sheet, description):
    """Find a row by description."""
    target = normalize(description)
    for row in range(2, sheet.max_row + 1):
        if normalize(sheet.cell(row=row, column=1).value) == target:
            return row
    return None


def find_month_col(sheet, month_abbr):
    """Find the correct month column."""
    target = normalize(month_abbr)
    for col in range(1, sheet.max_column + 1):
        val = sheet.cell(row=1, column=col).value
        if val and normalize(str(val)).startswith(target):
            return col
    return None


def update_excel(sheet, name, month_abbr, total):
    """Update Excel cell or append a new row."""
    row = find_description_row(sheet, name)
    if not row:
        row = sheet.max_row + 1
        sheet.cell(row=row, column=1).value = name

    col = find_month_col(sheet, month_abbr)
    if not col:
        log_message(f"Month '{month_abbr}' not found — skipping '{name}'")
        return None

    previous = sheet.cell(row=row, column=col).value or 0
    sheet.cell(row=row, column=col, value=total)
    return f"{name} ({month_abbr}): {previous} → {total}"


def extract_staff_name(description):
    """Extract staff name from a description string."""
    if not description:
        return "No Description"
    desc = normalize(description)
    cleaned = desc.replace("overspend", "").replace("overspent", "").replace("change", "").strip()
    cleaned = re.sub(r"[^a-zA-Z\s]", "", cleaned).strip()
    return cleaned.title() if cleaned else "No Description"


# ========================= QUICKBOOKS =========================
def refresh_access_token():
    """Refresh QuickBooks access token using the refresh token."""
    config = load_config()
    client_id = os.getenv("QB_CLIENT_ID")
    client_secret = os.getenv("QB_CLIENT_SECRET")

    auth = (client_id, client_secret)
    payload = {"grant_type": "refresh_token", "refresh_token": config["refresh_token"]}
    headers = {"Accept": "application/json", "Content-Type": "application/x-www-form-urlencoded"}

    resp = requests.post(TOKEN_URL, auth=auth, headers=headers, data=payload)
    resp.raise_for_status()
    tokens = resp.json()

    if tokens.get("refresh_token") and tokens["refresh_token"] != config["refresh_token"]:
        config["refresh_token"] = tokens["refresh_token"]
        save_config(config)

    return tokens["access_token"]


def fetch_all_journal_entries(access_token, realm_id):
    """Retrieve all journal entries from QuickBooks Online."""
    url = f"https://quickbooks.api.intuit.com/v3/company/{realm_id}/query"
    entries, pos, batch_size = [], 1, 1000

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
        entries.extend(batch)
        pos += batch_size
        if len(batch) < batch_size:
            break

    return entries


# ========================= MAIN =========================
def main():
    log_message("Starting QuickBooks automation...")

    access_token = refresh_access_token()
    log_message("Access token refreshed successfully.")

    config = load_config()
    realm_id = config.get("realm_id")

    entries = fetch_all_journal_entries(access_token, realm_id)
    log_message(f"Retrieved {len(entries)} journal entries.")

    wb = load_workbook(EXCEL_PATH)
    month_map = {m.lower(): m.capitalize() for m in ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                                                     "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]}
    current_year = datetime.now().year

    for acc in ACCOUNTS:
        sheet_name, account_name = acc["sheet"], acc["account"]

        if sheet_name not in wb.sheetnames:
            log_message(f"Sheet '{sheet_name}' not found. Skipping.")
            continue

        sheet = wb[sheet_name]
        totals = {}

        for je in entries:
            txn_date = je.get("TxnDate")
            if not txn_date:
                continue
            txn = datetime.strptime(txn_date, "%Y-%m-%d")
            if txn.year != current_year:
                continue

            month_abbr = month_map.get(txn.strftime("%b").lower())
            if not month_abbr:
                continue

            for line in je.get("Line", []):
                acct_name = line.get("JournalEntryLineDetail", {}).get("AccountRef", {}).get("name", "")
                if normalize(account_name) not in normalize(acct_name):
                    continue

                description = line.get("Description") or "No Description"
                staff_name = extract_staff_name(description)
                amount = float(line.get("Amount", 0))

                # Handle special case for Cash in Hands
                if "cash in hands" in normalize(account_name):
                    desc_norm = normalize(description)
                    if "overspend" in desc_norm or "overspent" in desc_norm:
                        amount = -abs(amount)
                    elif "change" in desc_norm:
                        amount = abs(amount)

                totals[(staff_name, month_abbr)] = totals.get((staff_name, month_abbr), 0) + amount

        updates = []
        for (staff_name, month), total in totals.items():
            result = update_excel(sheet, staff_name, month, total)
            if result:
                updates.append(result)

        if updates:
            log_message(f"{sheet_name}: {len(updates)} updates.")
            for u in updates:
                log_message("   " + u)
        else:
            log_message(f"No updates made for '{sheet_name}'.")

    wb.save(EXCEL_PATH)
    log_message("Excel workbook updated successfully.\n")


if __name__ == "__main__":
    main()
