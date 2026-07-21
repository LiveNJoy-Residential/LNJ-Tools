# LNJ Audit Bot — ResMan Concession & Revenue Integrity Audit System

Automated audit tool for **LiveNjoy Residential**. Reads CSV exports from ResMan, runs three independent audit engines, and produces an interactive Streamlit dashboard plus a timestamped Excel workbook.

---

## Audit Engines

| Engine | Owner | What It Checks |
|---|---|---|
| **Concession Audit** | John B. | 9 rules — unauthorized concessions, missing addenda, duplicate postings, over-applied credits |
| **Revenue Integrity** | Daniel Twito | 2-stage NER analysis — concession setups below market floor, double-discount structures, posted vs. recurring mismatches |
| **Fee Schedule Check** | — | Validates that recurring charges (pet fees, parking, etc.) match the approved fee schedule per property |

---

## Project Structure

```
audit_bot.py          ← Core engine: data ingestion, all rules, Excel export
app.py                ← Streamlit dashboard (7 tabs)
generate_frd.py       ← Converts FRD.md → FRD.docx (run once as needed)
FRD.md / FRD.docx     ← Functional Requirements Document
Fee Schedules.md      ← Approved fee schedule reference by property
Onboarding.md         ← New-user onboarding guide
requirements.txt      ← Python dependencies

data/
  transactions/       ← Transaction List Reports (Credit rows) — one CSV per property
  leases/             ← New & Renewed Leases — one CSV per property
  edits/              ← Edited Transactions by User — one CSV per property
  recurring/          ← Recurring Transaction Projections — one CSV per property
  rent_rolls/         ← Rent Rolls — one CSV per property
  activity/           ← Resident Activity reports — one CSV per property
  market rent schedule/ ← Market Rent Schedule Detail — one CSV per property

output/               ← Timestamped Excel reports written here
exports/              ← Raw ResMan export staging folder (not committed)
```

> **Note:** `data/` and `exports/` contain resident PII and are excluded from version control via `.gitignore`. Load CSV files locally before running.

---

## Setup

### 1 — Clone the repository

```powershell
git clone https://github.com/<org>/<repo>.git
cd <repo>
```

### 2 — Create and activate a virtual environment

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

### 3 — Install dependencies

```powershell
pip install -r requirements.txt
```

### 4 — Load data

Copy the 7-property ResMan CSV exports into the correct `data/` subfolders (one CSV per property per folder). File naming follows the ResMan export convention, e.g.:

```
data/transactions/Crossings at Irving Transaction List.csv
data/leases/Crossings at Irving New and Renewed Leases.csv
...
```

---

## Running the Bot

### Generate the Excel audit report

```powershell
.venv\Scripts\python.exe audit_bot.py
```

Output is written to `output/LNJ_Audit_<YYYYMMDD>_<HHMM>.xlsx`.

### Launch the Streamlit dashboard

```powershell
.venv\Scripts\streamlit.exe run app.py
```

Open [http://localhost:8501](http://localhost:8501) in your browser.

---

## Properties Covered

| Property |
|---|
| Crossings at Irving |
| Highland Park |
| La Prada |
| Parks on Taylor |
| Valencia Plaza |
| Village Green |
| Western Station |

---

## Data Sources (ResMan Exports)

All source files are exported from ResMan by a **full-access** account. Exports from limited-access accounts will be missing Credit transaction rows, causing the Concession Audit Engine to produce incomplete results.

| ResMan Report | `data/` Subfolder |
|---|---|
| Transaction List (Credits) | `data/transactions/` |
| New & Renewed Leases | `data/leases/` |
| Edited Transactions by User | `data/edits/` |
| Recurring Transaction Projection | `data/recurring/` |
| Rent Roll | `data/rent_rolls/` |
| Resident Activity | `data/activity/` |
| Market Rent Schedule Detail | `data/market rent schedule/` |

---

## Requirements

- Python 3.11+
- See `requirements.txt` for package versions
