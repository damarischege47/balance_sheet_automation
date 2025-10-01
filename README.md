# QuickBooks → Excel Automation Tool

## Overview
I built an automated QuickBooks → Excel integration tool. It dynamically updates multiple balance sheet schedules (cash, debtors, creditors, prepaid expenses, etc.) in real-time using the QuickBooks API and Excel automation. The tool refreshes OAuth tokens, fetches journal entries with pagination, and writes directly into mapped balance sheet schedules. It runs daily via Task Scheduler, removing duplicate data entry and ensuring finance always has the latest numbers.

---

## Features
- Fetches journal entries from QuickBooks API with pagination.
- Updates multiple Excel sheets: cash, debtors, creditors, prepaid expenses, statutory payables.
- Handles special cases (e.g., cash overspends and changes).
- Automatic OAuth token refresh and secure configuration.
- Logs all updates and skipped entries for auditing.
- Ready for daily automation via Task Scheduler or cron.

---

## Tech Stack
- **Python 3.x**
- `requests` – API calls
- `openpyxl` – Excel manipulation
- `datetime`, `json`, `os`, `re` – utility handling
- QuickBooks Online API (OAuth2)

---

## Setup Instructions

1. **Clone the repository**
```bash
git clone https://github.com/YOUR_USERNAME/quickbooks-excel-automation.git
cd quickbooks-excel-automation
