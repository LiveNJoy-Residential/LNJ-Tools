# LNJ Audit Bot — Onboarding Overview

---

## What Is This?

The **LNJ Audit Bot** is an automated financial audit tool built for **LiveNjoy Residential**.

Every month, it reads data from the company's property management software (**ResMan**), checks that data against a set of rules, and produces a report showing anything that looks wrong — missing fees, unauthorized discounts, posting errors, etc.

Think of it as a **very fast, very thorough auditor that runs every month in minutes instead of days.**

---

## The Problem It Solves

Managing 7 apartment properties means thousands of individual rent and fee transactions every month. Auditing those manually is slow and things get missed.

Some examples of what can go wrong:
- A resident gets a rent discount they were never approved for
- A fee is posted at the wrong dollar amount
- A credit is posted after a lease has already ended
- A unit is showing $0 rent when it should not be

The bot catches all of these automatically, every month, across all 7 properties at once.

---

## The 7 Properties It Covers

| Property |
|---|
| Crossings at Irving |
| Parks on Taylor |
| Highland Park |
| La Prada |
| Village Green |
| Valencia Plaza |
| Western Station |

---

## How It Works — Step by Step

### Step 1 — Export the Data

Once a month, staff exports **49 CSV files** from ResMan (7 properties × 7 report types). These files contain things like rent rolls, lease records, transaction histories, and fee schedules.

> **CSV = a spreadsheet-style file that any software can read. ResMan exports them and the bot reads them.**

The 7 report types are:
- Rent Roll
- Transaction List
- Recurring Transaction Projection
- New & Renewed Leases
- Resident Activity
- Edited Transactions by User
- Market Rent Schedule

### Step 2 — Run the Bot

One command starts the audit:

```
python audit_bot.py
```

The bot reads all 49 files, runs three separate audit engines, and finishes in seconds.

### Step 3 — Review the Results

The bot produces two outputs:
1. **An Excel file** — saved with a timestamp (e.g. `LNJ_Audit_20260630_1535.xlsx`) for records and sharing
2. **An interactive dashboard** — opens in a web browser at `http://localhost:8501`

---

## The Three Audit Engines

### Engine 1 — Concession Audit (John's Rules)

Checks that any rent concessions (discounts) are legitimate and properly documented.

| What It Checks | Risk Level |
|---|---|
| Credit posted after lease end date | CRITICAL |
| Credit posted but no lease on file | HIGH |
| Large single credit ($700 or more) | CRITICAL |
| Concession on rent roll but never actually posted | HIGH |
| Amount on rent roll doesn't match what was actually posted | HIGH |

### Engine 2 — Revenue Integrity Audit (Daniel's Rules)

Checks the setup of recurring charges and what is actually posted.

**Part A — Recurring Charge Setup**

| What It Checks | Risk Level |
|---|---|
| A standard fee is missing from a unit that should have it | HIGH |
| A unit is being charged the wrong amount for a fee | HIGH / MEDIUM |
| A large concession (over $500) is running for multiple months | HIGH |
| A concession has no expiration date | MEDIUM |

**Part B — Net Effective Rent (NER)**

This checks whether the actual rent a resident pays (after discounts) falls below the minimum allowed rent for that property.

| What It Checks | Risk Level |
|---|---|
| Discount equals full rent for 2+ months (essentially free rent) | CRITICAL |
| Full-month free rent — verify paperwork exists | HIGH |
| Rent already below market AND an extra discount on top | CRITICAL |
| Final net rent is below the property minimum floor | CRITICAL |

**Part C — Posted Rent Roll Audit**

| What It Checks | Risk Level |
|---|---|
| Occupied unit showing $0 rent (not a new move-in) | CRITICAL |
| Unit showing negative rent | CRITICAL |
| A credit was manually posted with no matching setup in the system | HIGH |
| What was posted doesn't match what the system projected | HIGH |

### Engine 3 — Fee Schedule Check

Compares every fee charged to every unit against the official approved fee schedule. Flags any fee posted at the wrong dollar amount.

---

## What the Dashboard Looks Like

The dashboard opens in a web browser and has **7 tabs**, one per property. Each tab shows:
- A summary of all flags found
- Risk levels (CRITICAL / HIGH / MEDIUM)
- The specific unit, amount, and explanation for each flag
- Total financial exposure in dollars

---

## What the Output Numbers Look Like

From the **June 2026 audit run** as an example:

| Metric | Number |
|---|---|
| Properties audited | 7 |
| CSV files processed | 49 |
| Total flags found | 511 |
| CRITICAL flags | 23 |
| Estimated dollar exposure | $128,000+ |

---

## How It Is Connected

```
ResMan (Property Management Software)
        |
        | (manual CSV export, once a month)
        |
   49 CSV Files saved to the data/ folder
        |
        | (one command: python audit_bot.py)
        |
   Audit Bot runs 3 engines
        |
        |--- Excel Report (saved to output/ folder)
        |
        |--- Streamlit Dashboard (opens in browser)
```

There is **no live/direct connection to ResMan**. The bot reads files — it does not log in to anything or change any data. It is completely read-only.

---

## Who Uses It and How

| Person | Role |
|---|---|
| Staff (you) | Exports the CSV files from ResMan each month, drops them in the `data/` folder, runs the bot |
| Daniel Twito | Reviews the results, confirms or dismisses flags, defined the audit rules |
| John B. | Defined the concession audit rules (Engine 1) |
| Management / VP | Receives the Excel report and acts on CRITICAL findings |

---

## Tech Stack (Simple Version)

| Component | What It Is |
|---|---|
| Python | The programming language the bot is written in |
| pandas | A Python library for reading and processing spreadsheet data |
| Streamlit | A Python library that turns the results into a web dashboard |
| openpyxl | A Python library that creates the Excel output file |
| ResMan | The property management software that the CSV data comes from |

The bot runs **locally on your computer**. No cloud, no external servers, no internet required.

---

## Key Things to Know

- **It does not change anything.** It only reads data and produces reports.
- **It runs once a month.** Timing aligns with when ResMan data is available.
- **It covers all 7 properties in one run.** No need to audit each property separately.
- **The rules were defined by John and Daniel.** The bot enforces their rules automatically.
- **CRITICAL flags are the priority.** These represent the highest financial or compliance risk.
