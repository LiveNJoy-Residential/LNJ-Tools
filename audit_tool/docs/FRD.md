# Functional Requirements Document (FRD)
## LNJ Audit Bot — LiveNjoy Residential / ResMan Audit Automation
**Version:** 2.0  
**Authors:** John B. (Concession Rules), Daniel Twito (Revenue Integrity Rules)  
**Company:** LiveNjoy Residential  
**System:** ResMan Property Management  

---

## 1. Purpose

The LNJ Audit Bot is a Python/Streamlit system that automates the monthly financial audit of all seven LiveNjoy Residential properties managed through ResMan. Prior to this system, property accountants and regional managers had to manually cross-reference six to seven different CSV exports per property to detect concession fraud, improper postings, and fee schedule violations — a process involving upward of 49 files and thousands of rows each month.

The bot solves that problem by ingesting all ResMan CSV exports at once, running three audit engines in sequence, and producing a consolidated exception report with severity-rated flags and a calculated financial exposure figure. It also generates a timestamped Excel workbook and an interactive Streamlit dashboard, giving both field staff and leadership a single source of truth for every audit finding in the portfolio.

Specifically, the system:

- Identifies concession postings that are unauthorized, past lease term, or mismatched with the signed lease document.
- Validates that recurring charges in the Transaction Projection match the property's published fee schedule and appear consistently across all occupied units.
- Detects net effective rents that fall below property-specific NER floors (either hardcoded or dynamically derived from the Market Rent Schedule minus an approved discount).
- Flags manager-level transaction reversals and amount edits so that override abuse can be ranked and reviewed.
- Quantifies total financial exposure — deduped across engines — so leadership can prioritize remediation by dollar impact.

---

## 2. Scope

### 2.1 In Scope

The following capabilities are fully implemented in `audit_bot.py` and surfaced through the `app.py` Streamlit dashboard.

**Data Ingestion**

| Data Source | ResMan Report Name | `data/` Folder |
|---|---|---|
| Transaction List (Credits) | Transaction List Reports | `data/transactions/` |
| New & Renewed Leases | New and Renewed Leases | `data/leases/` |
| Edited Transactions by User | Edited Transactions by User | `data/edits/` |
| Recurring Projection | Transaction Projections | `data/recurring/` |
| Rent Roll | Rent Rolls | `data/rent_rolls/` |
| Resident Activity | Resident Activity | `data/activity/` |
| Market Rent Schedule | Market Rent Schedule Detail | `data/market rent schedule/` |

All seven properties are supported: Crossings at Irving, Highland Park, La Prada, Parks on Taylor, Valencia Plaza, Village Green, and Western Station.

**Concession Audit Engine (Rules R1–R7)**

- **R1 Post-Term Credit** — credit posted after the resident's lease end date.
- **R2 Missing Lease** — credit posted but the Rent Roll shows no active lease for that unit.
- **R3 Large Credit (≥ $700)** — any single credit meeting or exceeding the CRITICAL threshold.
- **R4 Non-Standard Description** — credit description does not match any approved concession keyword.
- **R5 Missing Addendum / No RR Setup** — credit posted in Transaction List but no corresponding concession row on the Rent Roll.
- **R6 Amount Mismatch** — Rent Roll concession amount differs from the Transaction List posted amount by more than $10 or 10%.
- **R7 Not Properly Posted** — Rent Roll shows a concession setup, but no matching credit appears in the Transaction List.

**Revenue Integrity Engine (2-Stage + NER)**

- **Stage 1 — Recurring Projection Audit:**
  - 90% Rule (Missing Standard Charge) — any charge category present on ≥ 90% of units must appear on all units.
  - Major and Minor Charge Amount Variance — projection amount deviates from the property's published fee schedule.
  - Recurring Concession Red Flags — concession > $500 for 2+ months; concession with no expiration date.

- **Stage 1B — NER Engine (Net Effective Rent):**
  - Full Rent Recurring Offset (CRITICAL) — concession equals rent across two or more months.
  - Future Month Full Offset (HIGH) — single upcoming month is fully offset; addendum required.
  - Double-Discount Setup (CRITICAL) — rent is already below market and a concession is layered on top.
  - Net Effective Rent Below Floor (CRITICAL) — NER (rent + concession) falls below the property-specific floor. Floors are either hardcoded (`PROPERTY_NER_FLOORS`) or computed dynamically as Market Rent Schedule total rent minus `PROPERTY_NER_DISCOUNT`.

- **Stage 2 — Posted Rent Roll Audit:**
  - Negative Net Rent (CRITICAL) — unit's net rent position is negative after charges.
  - $0 Net Rent — flagged CRITICAL for established residents, MEDIUM for recent move-ins (within 60 days).
  - Manual Posting Without Setup — concession posted on Rent Roll with no recurring projection row.
  - Posted vs Recurring Mismatch — gross rent on Rent Roll differs from projection by more than the proration threshold.
  - Misc Tenant Credit — miscellaneous tenant credit exists without a corresponding recurring setup.

**Fee Schedule Check Engine**

- Validates every recurring charge amount for all 7 properties against the hardcoded `PROPERTY_FEE_SCHEDULE` constants (sourced from official fee schedule documents).
- Flags any charge whose posted amount deviates by more than $1 from the published schedule amount.
- Optional charges (parking, pet rent, washer/dryer) are validated in amount only when present — they are not flagged as missing under the 90% Rule.

**Manager Override Audit**

- Parses Edited Transactions by User reports to identify reversals and amount changes.
- Calculates per-manager revenue impact (negative = revenue reduction).
- Ranks managers by total override impact and returns a raw override event log.

**Streamlit Dashboard (7 Tabs)**

| Tab | Content |
|---|---|
| Executive Summary | Portfolio KPI metrics, risk-level totals, filterable all-flags grid |
| Concession Audit Engine | Rule summary table, unit-level flag detail |
| Revenue Integrity Engine | Stage 1 and Stage 2 flag drilldowns |
| Manager Overrides | Manager ranking by impact, raw edit log |
| Exposure Drilldowns | Flags segmented by property, rule, and risk level |
| Risk Matrix | Heatmap of severity counts by property |
| Fee Schedule Check | Per-charge fee schedule violations |

**Excel Export**

- Timestamped workbook exported to `output/` on each audit run (e.g. `LNJ_Audit_20260630_1535.xlsx`).
- Includes formatted worksheets for each engine's flags, manager ranking, and exposure summary.
- Conditional formatting applied via openpyxl for risk-level color coding.

---

### 2.2 Out of Scope

The following are explicitly not performed by the current system:

- **No write-back to ResMan** — the bot is read-only. It does not create, modify, or delete any data in ResMan or any connected database.
- **No live/real-time data connection** — all data comes from manually exported CSV files copied into the `data/` subfolders. There is no API connection or scheduled pull from ResMan.
- **No PDF processing** — Resident Ledger PDFs (located in `exports/Resident Ledgers/`) are not parsed. They are documented as manual-review-only.
- **No user authentication or role-based access control** — the Streamlit app has no login system. Access is controlled by where the app is deployed.
- **No manual override or correction workflow within the UI** — users cannot mark flags as resolved, add notes, or approve exceptions from inside the dashboard.
- **No automated alert or email distribution** — the bot does not send notifications. Distribution of results (Excel file, dashboard link) is a manual step.
- **No approval workflow integration** — the bot does not interact with any ticketing, workflow, or ERP system.
- **No multi-month trend analysis** — the audit is scoped to a single `AUDIT_MONTH` per run. The multi-month projection data (`df_proj_full`) is used only for NER calculation, not for cross-period trend reporting.
- **No La Prada fee schedule hardcoding** — La Prada fee schedule amounts are not included in `PROPERTY_FEE_SCHEDULE` because no official fee sheet was provided; that property's charges are not validated against a fixed schedule.

---

## 3. Definitions

| Term | Definition |
|---|---|
| **ResMan** | The property management software used by LiveNjoy Residential to record lease terms, post charges and credits, and generate financial reports. All source data is exported from ResMan as CSV files. |
| **Audit Month** | The calendar month being audited, set via the `AUDIT_MONTH` constant (e.g. `"Jun 2026"`). One run covers one month. |
| **Transaction List (Credits)** | A ResMan export listing all credit-type postings for a property in a given period. Includes original credits and their reversals. The bot reads only rows in `Credit - ` sections, excluding Renters Insurance. |
| **Transaction Projection** | A ResMan export showing recurring charges scheduled per unit across multiple future months. Section 3 ("Recurring Transactions by Unit") is parsed to extract per-unit, per-category amounts. |
| **Rent Roll** | A ResMan export showing the current charge configuration per unit — market rent, actual rent, all recurring charge and concession rows, lease dates, and resident balances. |
| **New & Renewed Leases** | A ResMan export listing all leases executed in the current period, including approved recurring and one-time concession amounts from the signed lease document. |
| **Resident Activity** | A ResMan export with one row per resident showing move-in date, lease start/end, actual rent, and the leasing agent/manager name. |
| **Market Rent Schedule Detail** | A ResMan export listing base market rent, per-unit amenity premiums, and total market rent for each unit in the property. Used to compute dynamic NER floors. |
| **Edited Transactions by User** | A ResMan export logging all transaction reversals and amount edits, grouped by manager login. |
| **Concession Audit Engine** | The concession audit pipeline (`run_concession_audit_engine`) that validates every credit posting against R1–R7 rules using the Transaction List, Leases, and Rent Roll. |
| **Revenue Integrity Engine** | The revenue integrity audit pipeline (`run_revenue_integrity_engine`) that runs Stage 1 (Recurring Projection), Stage 1B (NER), and Stage 2 (Rent Roll) checks. |
| **NER (Net Effective Rent)** | The true economic rent paid by a resident: `recurring_rent + recurring_concession` (concession amounts are stored as negative in ResMan). NER below the property floor triggers a CRITICAL flag. |
| **NER Floor** | The minimum acceptable Net Effective Rent per property and bedroom type. Defined as a hardcoded dollar amount in `PROPERTY_NER_FLOORS` or computed dynamically as `market_rent - PROPERTY_NER_DISCOUNT`. |
| **90% Rule** | A charge category that appears on 90% or more of a property's units is considered a "standard charge" and must appear on all units. Units missing that charge receive a `Missing Standard Charge` flag. |
| **APPROVED_CODES** | The set of ResMan transaction codes accepted as valid concession types: `CONR`, `CRTCO`, `EMPL`, `MCCR`, `RRFee`. |
| **Amount Impact** | The dollar value of the exception — the financial exposure attributable to a single flag row. Used to calculate total and deduped portfolio exposure. |
| **Deduped Exposure** | The conservative exposure figure calculated by taking the maximum `Amount_Impact` per unit across all engines, preventing double-counting when a unit appears in both the Concession Audit Engine and Revenue Integrity Engine flags. |
| **RISK_CRITICAL / HIGH / MEDIUM** | The three risk tiers used to prioritize flags. Assigned per rule via the `RISK_MAP` constant. CRITICAL = immediate action required; HIGH = investigate within the period; MEDIUM = monitor. |
| **Proration Guard** | A logical safeguard in Stage 2 that skips the Posted vs Recurring Mismatch check when either the recurring or posted rent is less than 60% of the other, to avoid false positives on move-in or move-out proration periods. |
| **Streamlit** | The open-source Python framework (`streamlit`) used to build the interactive audit dashboard. The UI runs locally via `streamlit run app.py` and communicates with the engine through `st.session_state`. |
| **openpyxl** | The Python library used to write and format the Excel output workbook, including conditional fill colors for risk levels. |
| **derive_property()** | A helper function that maps any ResMan export filename (by short code: CAI, HP, POT, etc., or by keyword) to the full canonical property name used throughout the system. |
| **make_flag()** | The standardized factory function that builds every exception record with a fixed schema: Property, Unit, Resident, Rule, Risk_Level, Detail, Amount_Impact, Source_File. |

---

## 4. Actors

### 4.1 Human Actors

| Actor | Role | Interaction with the System |
|---|---|---|
| **Property Accountant** | Primary operator of the audit tool. Responsible for exporting CSV files from ResMan each month and running the bot. | Exports 7 report types × 7 properties = up to 49 CSV files from ResMan; copies them into the correct `data/` subfolders; clicks **Run Full Forensic Audit** in the Streamlit sidebar; reviews the flag tables and Excel output to identify units requiring follow-up. |
| **Regional Manager / Daniel Twito** | Audit rules owner and financial decision-maker. Defines acceptable NER floors, fee schedule amounts, and approved concession codes. | Reviews the Executive Summary KPIs and the Revenue Integrity Engine tab; receives the Excel workbook; validates CRITICAL and HIGH flags; approves or disputes findings; provides updated fee schedules and NER floor parameters to the developer when policy changes. |
| **Property Manager / Leasing Agent** | On-site staff whose postings are audited. Named in the Manager Override tab by their ResMan login. | Does not directly use the bot. Their ResMan activity (transaction edits, lease setups, concession postings) is the data source. Findings are communicated back to them by the Regional Manager or Accountant. |
| **Portfolio Leadership** | Executive audience for the Risk Matrix and Executive Summary. | Receives the Streamlit dashboard link or shared Excel file for a high-level view of portfolio exposure, CRITICAL flag counts, and manager override rankings. Does not interact with the bot directly. |
| **Developer / System Owner** | Maintains `audit_bot.py` and `app.py`. | Updates `AUDIT_MONTH`, `PROPERTY_FEE_SCHEDULE`, `PROPERTY_NER_FLOORS`, and `PROPERTY_NER_DISCOUNT` constants each audit cycle. Applies bug fixes to parsing logic when ResMan export formats change. |

### 4.2 System Actors

| System Actor | Component | Responsibility |
|---|---|---|
| **CSV Ingestion Layer** | `load_transaction_list()`, `load_leases()`, `load_edits()`, `load_transaction_projection()`, `load_rent_roll()`, `load_resident_activity()`, `load_market_rent_schedule()` | Reads, cleans, and normalizes each ResMan export format. Handles multi-encoding fallback (UTF-8-sig → CP1252 → Latin-1), section-header forward-fill, and unit/currency normalization. Outputs typed pandas DataFrames. |
| **Concession Audit Engine** | `run_concession_audit_engine()` in `audit_bot.py` | Applies R1–R7 concession rules using cross-referenced lookups across the Transaction List, Rent Roll, and Leases. Returns a DataFrame of standardized flag records. |
| **Revenue Integrity Engine** | `run_revenue_integrity_engine()` in `audit_bot.py` | Applies Stage 1 (Recurring Projection), Stage 1B (NER), and Stage 2 (Rent Roll) revenue integrity rules. Accepts the full multi-month projection for NER calculations and the Market Rent Schedule lookup for dynamic floor computation. Returns a DataFrame of standardized flag records. |
| **Fee Schedule Engine** | `run_fee_schedule_check()` in `audit_bot.py` | Compares every recurring charge amount in the Transaction Projection against the hardcoded `PROPERTY_FEE_SCHEDULE` for 6 of the 7 properties. Returns violation flag records. |
| **Manager Override Engine** | `run_manager_override_audit()` in `audit_bot.py` | Parses the Edited Transactions by User export to identify reversals and amount changes. Computes per-manager revenue impact and rank ordering. |
| **Exposure Calculator** | `calculate_exposure()` in `audit_bot.py` | Aggregates all flag DataFrames into portfolio-level, property-level, rule-level, and risk-level exposure summaries. Computes deduped exposure (max impact per unit) alongside the raw sum. |
| **Excel Exporter** | `export_to_excel()` in `audit_bot.py` | Writes the complete audit output to a timestamped `.xlsx` workbook in `output/` using openpyxl, with conditional risk-level formatting. |
| **Streamlit UI** | `app.py` | Hosts the 7-tab interactive dashboard. Triggers `run_full_audit()` on button click, stores results in `st.session_state`, and renders all tables with risk-level color styling. Provides filter controls (Risk Level, Property, Rule) on the Executive Summary tab. |
| **ResMan (External)** | Property management platform | Source of truth for all financial data. Not directly accessed by the bot — data is consumed via manually exported CSV files only. |

---

## 5. Functional Requirements

### 5.1 Data Ingestion & Normalization

| ID | Requirement Description |
|---|---|
| FR-1 | The system shall discover all `.csv` files within each configured `data/` subfolder by scanning the directory at runtime. If a folder does not exist or contains no CSV files, the system shall print a warning and return an empty DataFrame rather than raising an exception. |
| FR-2 | The system shall attempt to decode every CSV file using UTF-8-sig encoding first, then fall back to CP1252, then Latin-1, to handle the full range of Windows-based ResMan export encodings. |
| FR-3 | The system shall derive the canonical property name from every CSV filename by first checking a short-code prefix map (`CAI`, `HP`, `POT`, `LP`, `VG`, `VP`/`VPA`, `WST`) and then falling back to a keyword-in-filename scan (`crossing`, `irving`, `taylor`, `highland`, `prada`, `village`, `valencia`, `western`). The resolved name shall be one of the seven standard full property names used throughout the system. |
| FR-4 | The system shall load the **Transaction List** by skipping the first 6 header rows, then forward-filling ResMan section headers from the first column into a `_section` field. It shall retain only rows whose `_section` starts with `"Credit - "`, excluding `"Credit - Renters Insurance Premium Credit"`, and only rows where the `Unit` field is a numeric string. |
| FR-5 | The system shall flag Transaction List rows as reversals (`Is_Reversal = True`) when their `Amount` is negative. Original credits and their reversals shall both be retained in the DataFrame so that downstream engines can compute net concession amounts. |
| FR-6 | The system shall normalize all currency fields (Amount, Market Rent, Rec. Conc., etc.) by stripping `$`, `,`, `"`, and whitespace characters before casting to float. Any value that fails conversion shall be treated as `0.0`. |
| FR-7 | The system shall normalize all unit numbers by stripping leading zeros and splitting on `" - "` to discard resident name suffixes embedded in the unit field (e.g., `"101 - Inez Lee"` → `"101"`). Units that resolve to an empty string shall be set to `"0"`. |
| FR-8 | The system shall strip ResMan status marker characters (`*` for NTV, `**` for MTM) from all resident name fields across every loader using a regex substitution before storing names. |
| FR-9 | The system shall load the **New & Renewed Leases** report by skipping 5 header rows, retaining only rows where `Unit` is numeric, and normalizing `Rec. Conc.`, `One Time Conc.`, `Rent`, and `Market Rent` as currency fields and `Lease Start Date` / `Lease End Date` as timestamps. |
| FR-10 | The system shall load the **Edited Transactions by User** report by iterating rows and classifying each as either a manager name row (non-date, non-header first cell) or a data row (first cell matches `M/D/YYYY` format). It shall track the most recently seen manager name and attach it as `Manager_Login` to every subsequent data row. |
| FR-11 | The system shall classify every edit event as either a **Reversal** (Reversal Date is populated) or an **Amount Change** (Edited Amount differs from original Amount by more than $0.01). Rows that are neither shall be discarded. Reversals of $0 transactions shall also be discarded because they carry no revenue impact. |
| FR-12 | The system shall compute `Revenue_Impact` for each edit event as `-Original_Amount` for reversals and `Edited_Amount - Original_Amount` for amount changes. |
| FR-13 | The system shall load the **Transaction Projection** by locating the `"Recurring Transactions by Unit"` section marker within the CSV, reading the next row as column headers, and extracting the column whose header matches `AUDIT_MONTH` (case-insensitive). If no matching month column is found, it shall fall back to the fourth column (index 3). |
| FR-14 | The system shall load the **multi-month Transaction Projection** (`load_transaction_projection_all_months`) by extracting all columns whose headers match the pattern `"Mon YYYY"` (e.g., `"Jun 2026"`). It shall produce a long-format DataFrame with one row per `(Unit, Category, Month_Label)` combination. Concession amounts shall remain as stored in ResMan (negative values). |
| FR-15 | The system shall load the **Rent Roll** by skipping 6 header rows, then iterating rows to identify unit header rows (col 0 is numeric) and charge sub-rows (col 0 is blank, col 18 is a description). It shall carry forward `Unit`, `Residents`, `Unit_Type`, `Status`, `Market_Rent`, `Move_In`, `Lease_Start`, `Lease_End`, and `Balance` from each unit header row to all of its associated charge rows. |
| FR-16 | The system shall load the **Resident Activity** report by skipping 6 header rows, discarding the `"Adjusted Lease End Date"` overflow header row, and retaining only rows where the first column is a numeric unit number. The Manager name shall be resolved by scanning backward from the last column to find the first non-blank, non-NaN value after column 43. |
| FR-17 | The system shall load the **Market Rent Schedule** by skipping 5 header rows, retaining only rows where the first column is a 3–4 digit numeric unit number, and reading the 8th column (index 7, labelled `"Total Rent"`) as the per-unit total market rent including all amenity premiums. The result shall be stored as a dictionary keyed by `(property_name, unit_number_string)`. |
| FR-18 | The system shall parse all date fields using `pd.to_datetime` with `infer_datetime_format=True`. If parsing fails for any value, the field shall be set to `None` rather than raising an exception. |

---

### 5.2 Concession Audit Engine (Rules R1–R7)

| ID | Requirement Description |
|---|---|
| FR-19 | The system shall skip the Concession Audit Engine entirely only when **both** the Transaction List DataFrame and the Rent Roll DataFrame are empty. If the Transaction List is empty but the Rent Roll is present, the system shall skip rules R1–R4 with a printed warning and continue to execute R7 from the Rent Roll side. |
| FR-20 | **R1 — Post-Term Credit:** The system shall flag any credit row in the Transaction List whose `Date` is later than the `Lease_End` date for the same `(Property, Unit)` as found in the New & Renewed Leases report. The flag shall be rated CRITICAL. |
| FR-21 | **R2 — Missing Lease:** The system shall flag any credit-posting unit where the Rent Roll shows no active lease (i.e., `Lease_End` is null or in the past relative to the audit date). The system shall use the Rent Roll lease-end lookup — not just the New & Renewed Leases for the current month — so that all ongoing leases signed in prior months are correctly recognized as active. The flag shall be rated HIGH. |
| FR-22 | **R3 — Large Credit:** The system shall flag any non-reversal credit row in the Transaction List where `abs(Amount) >= $700` (`CONCESSION_CRITICAL_AMT`). The flag shall be rated HIGH (single month). Multi-month full offsets that implicitly produce large recurring credits are captured separately by the NER engine as CRITICAL. |
| FR-23 | **R4 — Non-Standard Description:** The system shall flag any credit row whose `Description` does not contain any of the approved concession keywords (`"concession"`, `"allowance"`, `"employee unit"`, `"courtesy officer"`, `"resident referral"`, `"referral"`, `"move in special"`, `"move-in special"`, `"reduce"`, `"special"`, `"discount"`). |
| FR-24 | **R5 — Missing Addendum / No RR Setup:** The system shall flag any `(Property, Unit)` that has a net-positive credit in the Transaction List (`Amount` summed across non-reversal rows > 0) but no concession row on the Rent Roll. Concession rows on the Rent Roll are identified by keywords (`"concession"`, `"special"`, `"reduce"`, `"employee"`, `"discount"`, `"free"`, `"allowance"`, `"courtesy"`, `"move in"`, etc.) combined with a negative amount (< -$0.01). Employee unit and resident referral credit sections in the Transaction List shall be excluded from R5 triggering. |
| FR-25 | **R6 — Amount Mismatch:** The system shall compare the Rent Roll concession amount (absolute value of the sum of all concession rows for the unit) against the net credit posted in the Transaction List (sum of positive credits minus reversals). If the variance exceeds $10 or 10% of the RR amount (whichever is smaller triggers the flag), the system shall emit an Amount Mismatch flag rated HIGH. |
| FR-26 | **R7 — Not Properly Posted:** The system shall flag any `(Property, Unit)` that has a concession row on the Rent Roll (negative amount + concession keyword) but no corresponding net credit in the Transaction List. This rule shall run even when the Transaction List is empty for the month (e.g., when the audit-month export contains no Credit sections), because the Rent Roll represents the approved setup that should have posted. The flag shall be rated HIGH. |
| FR-27 | The system shall build the net Transaction List credit per unit by summing all credit amounts (positive original credits plus negative reversals) so that a reversed-and-reposted concession is reflected as the final net amount rather than the gross sum of all positive rows. |
| FR-28 | All flags produced by the Concession Audit Engine shall be emitted as standardized flag records via `make_flag()`, containing: `Property`, `Unit`, `Resident`, `Rule`, `Risk_Level` (from `RISK_MAP`), `Detail` (human-readable explanation), `Amount_Impact`, and `Source_File`. |

---

### 5.3 Revenue Integrity Engine

#### 5.3.1 Stage 1 — Recurring Projection Audit

| ID | Requirement Description |
|---|---|
| FR-29 | **90% Rule — Missing Standard Charge:** For each `(Property, Category)` combination in the Transaction Projection, the system shall compute the percentage of occupied units that have a positive amount for that category. If the percentage is ≥ 90% (`STANDARD_CHARGE_THRESHOLD`), the category is classified as a standard charge, and every unit missing that charge shall receive a `Missing Standard Charge` flag rated HIGH. |
| FR-30 | The system shall exclude categories matching optional charge keywords (`"carport"`, `"parking"`, `"pet rent"`, `"pet fee"`, `"washer"`, `"dryer"`, `"first floor"`, `"1st floor"`) from the 90% Rule check, because these charges are unit-specific add-ons and not expected portfolio-wide. |
| FR-31 | **Major Charge Amount Variance:** The system shall group Transaction Projection rows by `(Property, Unit_Type, Category)` and compute the statistical mode amount for each group (using only positive-amount rows and requiring at least 3 data points). Any unit whose amount for that category deviates from the mode by ≥ 20% AND ≥ $5.00 shall receive a `Major Charge Amount Variance` flag rated HIGH. |
| FR-32 | **Minor Charge Amount Variance:** Any unit whose charge amount deviates from the mode by ≥ $1.00 but does not meet both thresholds for Major variance shall receive a `Minor Charge Amount Variance` flag rated MEDIUM. |
| FR-33 | The system shall exclude `"Rent"` category rows from amount variance checks. Rent legitimately varies between units due to different lease terms and negotiated rates; rent discrepancies against the Rent Roll are fully handled by the Stage 2 Posted vs Recurring Mismatch rule. Optional charge categories are also excluded from amount variance because their amount validation is owned by the Fee Schedule Check engine. |
| FR-34 | **Recurring Concession >$700:** The system shall scan all concession rows in the Transaction Projection (rows matching concession keywords: `"concession"`, `"conr"`, `"crtco"`, `"empl"`, `"mccr"`, `"rrfee"`, `"employee unit"`, `"resident referral"`, `"courtesy officer"`) for the audit month. If `abs(Amount) > $700` for any unit, it shall emit a `Recurring Concession >$700` flag rated HIGH. Concession amounts are stored as negative in the projection; `abs()` is used for all threshold comparisons. |
| FR-35 | **Concession >$500 for 2+ Months:** The system shall flag any unit whose recurring concession is exactly $500 (within $1.00) for the audit month, and any unit whose concession exceeds $500 and spans more than 2 months in the projection, both rated HIGH, since these arrangements require documented lease addenda. |

#### 5.3.2 Stage 1B — NER Engine

| ID | Requirement Description |
|---|---|
| FR-36 | The system shall execute the NER Engine only when the multi-month projection DataFrame (`df_proj_full`) is not empty. For each `(Property, Unit)`, it shall classify projection rows as rent rows (category contains `\brent\b` and not `"concession"`) and concession rows (category matches any concession keyword). It shall then compute `rent_by_month` and `conc_by_month` dictionaries for every month in the projection. |
| FR-37 | The system shall resolve the NER floor for each unit as follows: (1) check `PROPERTY_NER_FLOORS[property][bedroom_type]`; (2) if that is `None`, check `PROPERTY_NER_DISCOUNT[property]` and look up the unit's market rent in the Market Rent Schedule; if found, compute floor = `market_rent - discount`. If neither source yields a floor, NER checks are skipped for that unit. |
| FR-38 | The system shall derive the bedroom type (`1BR`, `2BR`, `3BR`) from the ResMan unit type code using the following logic: if the numeric suffix in the code is ≥ 3, classify as 3BR; otherwise use the letter prefix (A → 1BR, B → 2BR, C/D → 3BR). |
| FR-39 | **Full Rent Recurring Offset (CRITICAL):** The system shall identify months where `abs(rent + concession) < $1.00` (net rent is effectively zero). If two or more such months exist in the projection for a unit, the system shall emit a `Full Rent Recurring Offset` flag rated CRITICAL. The `Amount_Impact` shall be set to `rent_amount × number_of_full_offset_months` to represent total potential revenue loss. The detail shall include the month span, monthly rent, concession, and the instruction to correct the future concession setup immediately. |
| FR-40 | **Future Month Full Offset (HIGH):** If exactly one zero-NER month exists for a unit, the system shall emit a `Future Month Full Offset` flag rated HIGH. The detail message shall distinguish between the current audit month ("one-month-free concession in the current period — verify addendum") and a future month ("one-month-free set for an upcoming month — confirm addendum before it posts"). `Amount_Impact` shall be set to `$0.00` (the financial event has not materialized). |
| FR-41 | **Double-Discount Setup (CRITICAL):** For the current audit month, if a unit's NER is below its floor AND its recurring rent is already below market rent AND it also carries a concession, the system shall emit a `Double-Discount Setup` flag rated CRITICAL. The detail shall state the market rent, the recurring rent, the concession, the resulting NER, the floor, and the required action (obtain VP authorization or correct so NER ≥ floor). `Amount_Impact` shall be set to `floor - NER`. |
| FR-42 | **Net Effective Rent Below Floor (CRITICAL):** For the current audit month, if a unit's NER is below its floor but the setup does not qualify as a double-discount (rent is not below market), the system shall emit a `Net Effective Rent Below Floor` flag rated CRITICAL with `Amount_Impact = floor - NER`. |
| FR-43 | The system shall apply a proration guard to NER checks: if the unit's recurring rent for the audit month is less than 50% of its market rent (sourced from the Rent Roll), the NER check shall be skipped for that unit, as the low rent figure likely reflects a partial-month move-in or move-out proration rather than a recurring concession issue. |
| FR-44 | The system shall track flagged units in a per-run `_ner_flagged` set. Once a unit has been emitted as a Full Rent Recurring Offset, Double-Discount, or NER Below Floor flag, it shall not also be emitted for the remaining NER sub-rules for the same unit in the same run, preventing double-flagging of the same underlying condition. |

#### 5.3.3 Stage 2 — Posted Rent Roll Audit

| ID | Requirement Description |
|---|---|
| FR-45 | Stage 2 shall operate only on Rent Roll rows whose `Status` field is one of `"C"` (Current), `"MTM"` (Month-to-Month), or `"NTV"` (Notice to Vacate). Units with other statuses (vacant, applicant, etc.) shall be excluded. |
| FR-46 | The system shall build a move-in date lookup for Stage 2 by first consulting the Resident Activity report for each `(Property, Unit)`, then falling back to the `Lease Start` date from the New & Renewed Leases report if no activity record exists. |
| FR-47 | **Negative Net Rent (CRITICAL):** For each occupied unit on the Rent Roll, the system shall sum all charge rows whose `Description` contains `\brent\b` or `\bbase\b` (case-insensitive, regex) and whose `Amount` is positive. If the sum is negative, the system shall emit a `Negative Net Rent` flag rated CRITICAL with `Amount_Impact = abs(net_rent)`. |
| FR-48 | **$0 Net Rent — Recent Move-in (MEDIUM):** If a unit's computed net rent is exactly $0 and the resident's move-in date is within 60 days (`RECENT_MOVEIN_DAYS`) of the run date, the system shall emit a `$0 Net Rent (Recent Move-in)` flag rated MEDIUM. The detail shall instruct reviewers to verify first-month timing and confirm a free-month addendum is on file if applicable. |
| FR-49 | **$0 Net Rent — Not Recent (CRITICAL):** If a unit's computed net rent is exactly $0 and the resident is not a recent move-in, the system shall emit a `$0 Net Rent (Not Recent)` flag rated CRITICAL. The detail shall direct the reviewer to check whether the unit is a courtesy officer, employee, or model unit and confirm an approved addendum, or investigate for an unauthorized full-offset posting. |
| FR-50 | **Manual Posting Without Setup (HIGH):** The system shall compare every `(Property, Unit)` that has a net-positive credit in the Transaction List against the set of units present in the Transaction Projection. Any unit with a posted credit but no projection entry shall receive a `Manual Posting Without Setup` flag rated HIGH, indicating a credit was manually entered without a recurring concession setup. |
| FR-51 | **Posted vs Recurring Mismatch (HIGH):** For each occupied unit, the system shall compare the gross rent posted on the Rent Roll (sum of positive-amount rows whose `Description` matches `\brent\b` or `\bbase\b` and excludes concession rows) against the gross recurring rent in the Transaction Projection (rows matching `\brent\b` or `\bbase\b` excluding concession rows, with positive amounts). If the absolute variance exceeds $5.00, the system shall emit a `Posted vs Recurring Mismatch` flag rated HIGH. |
| FR-52 | The system shall apply a bidirectional proration guard to the Posted vs Recurring Mismatch check: if the smaller of `(posted_rent, recurring_rent)` is less than 60% of the larger, the comparison shall be skipped. This prevents false positives in both directions — a prorated new-move-in projection (small projection vs full Rent Roll) and a prorated move-out Rent Roll charge (small Rent Roll vs full projection). |
| FR-53 | The mismatch flag detail shall include a directional explanation: if the projection is higher than the Rent Roll, the message shall state the Rent Roll may reflect a rent reduction not yet updated in the recurring setup; if the Rent Roll is higher, it shall state a rent increase may have been applied to the Rent Roll but the recurring charge was not updated to match. |
| FR-54 | **Misc Tenant Credit (HIGH):** The system shall scan all non-reversal, positive-amount Transaction List rows whose `Description` contains any miscellaneous keyword (`"misc"`, `"miscellaneous"`, `"adjustment"`, `"write-off"`, `"write off"`, `"reclass"`, `"mccr"`) and emit a `Misc Tenant Credit` flag for each such row, rated HIGH, with the instruction to review individually. |

---

### 5.4 Fee Schedule Check Engine

| ID | Requirement Description |
|---|---|
| FR-55 | The system shall validate recurring charge amounts for all properties that have an entry in `PROPERTY_FEE_SCHEDULE`. Properties without a fee schedule (currently La Prada, which has no official fee sheet) shall be silently skipped with a printed informational message. |
| FR-56 | For each unit in the Transaction Projection, the system shall match each fee defined in the property's schedule to the unit's projection rows by checking whether any of the fee's `keywords` list appears in the `Category` field (case-insensitive substring match). |
| FR-57 | If no projection row matches a fee's keywords for a given unit, the system shall take no action. Missing charges are the responsibility of the 90% Missing Standard Charge rule, not the Fee Schedule Check. |
| FR-58 | If a matching projection row exists with a positive amount, the system shall compute `variance = abs(actual_amount - scheduled_amount)`. If `variance >= $1.00`, the system shall emit a `Fee Schedule Violation` flag rated HIGH with the fee name, scheduled amount, actual amount, variance, and the instruction to review the lease addendum. |
| FR-59 | For optional per-item fees (e.g., parking at $35/space), a resident may legitimately have multiple units of the same fee (e.g., 3 parking spaces = $105). The system shall skip the violation check when the actual amount is an exact whole-number multiple of the per-unit scheduled rate and the fee is marked `optional: True` in the schedule. |
| FR-60 | Concession and credit rows (negative amounts) in the Transaction Projection shall be excluded from all fee schedule comparisons. Only positive-amount rows are evaluated. |

---

### 5.5 Manager Override Audit

| ID | Requirement Description |
|---|---|
| FR-61 | The system shall aggregate the override event log by `(Property, Manager_Login)` to produce a Manager Ranking DataFrame with columns: `Total_Events`, `Reversals`, `Amount_Changes`, and `Total_Impact` (sum of `Revenue_Impact` across all events for that manager). |
| FR-62 | The Manager Ranking shall be sorted in ascending order by `Total_Impact` so that managers with the largest negative revenue impact (most reversals or reductions) appear at the top of the table. |
| FR-63 | The system shall return both the Manager Ranking summary DataFrame and the full Override Detail Log (all individual edit events) as separate outputs. The Override Detail Log shall retain all fields: `Property`, `Manager_Login`, `Unit`, `Resident`, `Category`, `Description`, `Original_Amount`, `Edited_Amount`, `Event_Type`, `Revenue_Impact`, `Date`, and `Source_File`. |

---

### 5.6 Financial Exposure Calculation

| ID | Requirement Description |
|---|---|
| FR-64 | The system shall aggregate all flags from all three engines (Concession Audit, Revenue Integrity, Fee Schedule) into a single `all_flags` DataFrame by concatenating the three engine outputs. |
| FR-65 | The system shall compute **Exposure by Property**: a DataFrame grouped by `(Property, Risk_Level)` showing `Exceptions` (count) and `Total_Exposure` (sum of `Amount_Impact`), sorted descending by exposure. |
| FR-66 | The system shall compute **Exposure by Rule**: a DataFrame grouped by `(Rule, Risk_Level)` showing `Count` and `Total_Exposure`, sorted descending by exposure. |
| FR-67 | The system shall compute **Exposure by Risk Level**: a DataFrame grouped by `Risk_Level` showing `Count` and `Total_Exposure`. |
| FR-68 | The system shall compute a **Portfolio Totals** summary row containing: `Total_Units_Audited` (count of distinct units across all flags), `Total_Exceptions` (total row count), `Total_Exposure` (raw sum of all `Amount_Impact` values), `Deduped_Exposure`, `Critical_Flags`, `High_Flags`, `Medium_Flags`, and `Avg_Flags_Per_Unit`. |
| FR-69 | The system shall compute **Deduped Exposure** by taking the maximum `Amount_Impact` per `(Property, Unit)` pair across all flags before summing. This prevents double-counting when the same unit is flagged by multiple engines for the same underlying financial event. The deduped figure represents the conservative floor; the raw `Total_Exposure` is the ceiling. |

---

### 5.7 Excel Export

| ID | Requirement Description |
|---|---|
| FR-70 | The system shall write the audit output to a timestamped Excel workbook in the `output/` directory using the filename format `LNJ_Audit_YYYYMMDD_HHMM.xlsx`. The `output/` directory shall be created if it does not already exist. |
| FR-71 | The workbook shall contain the following sheets when data exists: `Executive Summary`, `All Exceptions`, `Concession Audit Engine`, `Revenue Integrity Engine`, `Fee Schedule Violations`, `Exposure by Property`, `Exposure by Rule`, `Manager Ranking`, and `Override Detail Log`. Sheets with empty source DataFrames shall be omitted. |
| FR-72 | All three flag sheets (`All Exceptions`, `Concession Audit Engine`, `Revenue Integrity Engine`, `Fee Schedule Violations`) shall have two resolution workflow columns prepended: `Status` (defaulting to `"Open"`) and `Notes` (empty string). |
| FR-73 | The `All Exceptions` sheet shall be sorted by risk priority (CRITICAL first, then HIGH, then MEDIUM) and within each tier by `Amount_Impact` descending, so the most financially impactful critical findings appear at the top. |
| FR-74 | All flag sheets shall have a dropdown data validation applied to the `Status` column for every data row, restricted to the values `"Open"`, `"Reviewed"`, `"Cleared"`, and `"Escalated"`. The dropdown arrow shall be visible in Excel (via `showDropDown=False` in openpyxl's `DataValidation`). |
| FR-75 | All flag sheets shall apply row background fill colors based on the current `Status` value: no fill for Open, light blue (`BDD7EE`) for Reviewed, light green (`C6EFCE`) for Cleared, and light orange (`FFEB9C`) for Escalated. |
| FR-76 | All sheets shall have a dark blue header row (`1F4E79` fill, white bold 11pt font), center-aligned header text, and frozen panes set at row 2 so the header remains visible while scrolling. |
| FR-77 | The system shall auto-fit column widths across all sheets by sampling up to 200 data rows per column, capping individual cell contribution at 60 characters, and adding a 2-character buffer. Minimum column width shall be 10 characters. |
| FR-78 | A thin bottom border (`CCCCCC`) shall be applied to every data cell in all flag sheets to improve row readability. Text alignment shall be set to top-vertical with no word wrap on data cells. |

---

### 5.8 Streamlit Dashboard UI

| ID | Requirement Description |
|---|---|
| FR-79 | The Streamlit app shall render a sidebar containing: the LiveNjoy property icon, the system title and version, the list of approved concession codes, a **Run Full Forensic Audit** primary button, and a File Setup Guide explaining which ResMan exports map to which `data/` subfolders. |
| FR-80 | Clicking **Run Full Forensic Audit** shall trigger a `st.spinner` context while `run_full_audit()` executes. On success, results shall be stored in `st.session_state["results"]`. On failure, an error message shall be displayed and the app shall call `st.stop()`. |
| FR-81 | The app shall display an informational message and stop rendering if no audit results are present in `st.session_state` (i.e., the button has not yet been clicked). |
| FR-82 | The dashboard shall render seven tabs: **Executive Summary**, **Concession Audit Engine**, **Revenue Integrity Engine**, **Manager Overrides**, **Exposure Drilldowns**, **Risk Matrix**, and **Fee Schedule Check**. |
| FR-83 | **Tab 1 — Executive Summary:** The tab shall display five KPI metric tiles: Units Audited, Total Exceptions, Financial Exposure (deduped), Avg Flags / Unit, and Critical Flags. It shall also display three risk-level counts (CRITICAL, HIGH, MEDIUM) as separate metric tiles. |
| FR-84 | **Tab 1 — All Exceptions Grid:** Below the KPI tiles, the tab shall render the complete `all_flags` DataFrame with three interactive filter controls: a multiselect for Risk Level (defaulting to all three levels), a selectbox for Property (including "All"), and a selectbox for Rule (including "All"). The displayed row count shall be shown in a caption below the table. |
| FR-85 | **Tab 2 — Concession Audit:** The tab shall first display a rule summary table grouped by `(Rule, Risk_Level)` with count and total exposure columns, then the full unit-level flag table. A caption shall show the raw sum of the Concession Audit Engine exposure with a note that deduplication is applied in the Executive Summary. |
| FR-86 | **Tab 3 — Revenue Integrity:** The tab shall render two inner sub-tabs: **Stage 1 — Recurring Projection** and **Stage 2 — Posted Rent Roll**. Within Stage 1, it shall display sub-sections for Missing Standard Charges (90% Rule), Charge Amount Inconsistencies (variance rules), and Concession Red Flags. Within Stage 2, it shall display sub-sections for Net Rent Integrity Issues and Manual Concession / Invalid Code flags. |
| FR-87 | **Tab 4 — Manager Overrides:** The tab shall display the Manager Ranking DataFrame sorted by total revenue impact, followed by the raw Override Detail Log. |
| FR-88 | **Tab 5 — Exposure Drilldowns:** The tab shall display the Exposure by Property, Exposure by Rule, and Exposure by Risk Level summary tables. |
| FR-89 | **Tab 6 — Risk Matrix:** The tab shall render a heatmap visualization showing flag severity counts by property, providing a portfolio-level spatial overview of where risk is concentrated. |
| FR-90 | **Tab 7 — Fee Schedule Check:** The tab shall display the Fee Schedule Violations DataFrame from the Fee Schedule Check engine output. |
| FR-91 | All DataFrames displayed in the dashboard shall use the `styled_df()` helper, which applies `color_risk()` background styling to the `Risk_Level` column: red (`#FF4B4B`) for CRITICAL, orange (`#FFA500`) for HIGH, and gold (`#FFD700`) for MEDIUM, with black text and bold font weight. DataFrames without a `Risk_Level` column shall be rendered unstyled. |
| FR-92 | All DataFrames shall be rendered with `use_container_width=True` and `hide_index=True` to maximize readability and suppress the pandas row index. |

---

## 6. Business Rules (Exception Logic)

> **Severity key:** **Error** = definitive policy violation; produces an exception flag record in the audit output and requires corrective action. **Warning** = data quality guard or soft-logic threshold; either logs a console message without producing a flag record, or produces a flag whose finding may have an acceptable explanation and must be investigated before a classification is made.

### 6.1 Data Ingestion & Quality Guards

| ID | Rule | Severity | Requires |
|---|---|---|---|
| BR-1 | If a configured `data/` subfolder does not exist on disk, the system shall print `[WARN] Folder missing: <path>` and return an empty DataFrame. No exception shall be raised and processing of other folders shall continue normally. | Warning | Configured `DIRS` path entry; file system access |
| BR-2 | If a configured `data/` subfolder exists but contains no `.csv` files, the system shall print `[INFO] No CSVs in: <path>` and return an empty DataFrame. | Warning | `data/` subfolder present; `.csv` file extension |
| BR-3 | Every CSV file shall be decoded by attempting UTF-8-sig first, CP1252 second, and Latin-1 third. If all three encodings fail with `UnicodeDecodeError`, the system shall raise `ValueError` identifying the unreadable file. | Warning | CSV file on disk; one of three supported encodings |
| BR-4 | Every per-file loader is wrapped in a `try/except Exception` block. If any single file raises an unhandled exception during parsing, the system shall print `[ERROR] <filename>: <message>` and skip that file, allowing the remaining files to load normally. | Warning | Valid CSV structure; parseable content |
| BR-5 | Any value in a currency field (`Amount`, `Market Rent`, `Rec. Conc.`, etc.) that cannot be cast to float after stripping `$`, `,`, `"`, and whitespace shall be silently coerced to `0.0`. The rule applies universally via `clean_currency()`. | Warning | String-typed currency column present |
| BR-6 | Any value in a date field that cannot be parsed by `pd.to_datetime` shall be set to `None`. The rule applies universally via `parse_date()`. Downstream comparisons involving `None` dates shall be skipped using `pd.notna()` guards. | Warning | String-typed date column present |
| BR-7 | A unit number that is blank or `"nan"` after stripping shall be stored as `"UNKNOWN"`. A unit number that resolves to an empty string after stripping leading zeros shall be stored as `"0"`. | Warning | Unit column present in raw CSV |
| BR-8 | In the Transaction List loader, ResMan section headers (e.g., `"Credit - Concession - Rent"`, `"Payment - Payment"`) that appear in the first column shall be forward-filled into a `_section` field via `ffill()`. All non-`"Credit - "` section types — including Payment, Deposit, and Charge sections — must be captured by the forward-fill pattern so they correctly reset the section label and prevent Payment or Deposit rows from inheriting a prior Credit section label. | Warning | Multi-section ResMan Transaction List CSV |
| BR-9 | The `"Credit - Renters Insurance Premium Credit"` section shall be explicitly excluded from the Transaction List even though it begins with `"Credit - "`. These rows represent resident insurance payment activity, not concessions, and shall never appear in the engine inputs. | Warning | Transaction List containing insurance credit rows |
| BR-10 | In all loaders, only rows where the `Unit` column is a purely numeric string (matches `r"^\d+$"`) shall be retained as data rows. All other rows (section headers, totals, blank rows) shall be discarded. | Warning | Mixed header/data row format in ResMan export |
| BR-11 | In the Transaction Projection loader, if the string `"Recurring Transactions by Unit"` is not found as a section marker in the file, the system shall print `[WARN]` and skip that file entirely. If the section is found but no column header matches `AUDIT_MONTH` (case-insensitive), the system shall fall back to column index 3 as the amount column. | Warning | Transaction Projection CSV; correct `AUDIT_MONTH` value set |
| BR-12 | In the Market Rent Schedule loader, if a file does not have at least 8 columns after skipping header rows, the system shall print `[WARN]` and skip that file. The `"Total Rent"` value (index 7) must exist as a column. | Warning | Market Rent Schedule CSV with ≥ 8 columns |
| BR-13 | The Resident Activity loader shall discard the `"Adjusted Lease End Date"` overflow header row before processing data rows. This row is emitted by ResMan as a secondary header and must not be treated as a resident record. | Warning | Resident Activity CSV with overflow header row |
| BR-14 | In the Rent Roll loader, any row where `col[18]` (Description) equals `"Total"` shall be skipped. This row is a ResMan-generated subtotal row and does not represent a chargeable line item. | Warning | Rent Roll CSV with ResMan Total rows |
| BR-15 | In the Edited Transactions by User loader, a reversal event shall only be recorded when `Original_Amount != 0.0`. Reversals of $0 transactions carry no revenue impact and shall be discarded. | Warning | Edited Transactions CSV; `Reversal Date` populated; `Amount` field parseable |

---

### 6.2 Concession Audit Engine

| ID | Rule | Severity | Requires |
|---|---|---|---|
| BR-16 | **Engine Entry Guard:** The Concession Audit Engine shall be skipped entirely — with `[SKIP]` printed — only when **both** the Transaction List DataFrame and the Rent Roll DataFrame are empty. This dual-empty condition is the only valid reason to skip the engine. | Warning | `df_trans` and `df_rent_roll` both empty |
| BR-17 | **TX-Empty Guard:** If the Transaction List is empty but the Rent Roll is not, rules R1–R3 (which require Transaction List rows) shall be skipped with a `[WARN]` message. Rule R7 (Not Properly Posted) shall still execute because it operates entirely from the Rent Roll side. | Warning | `df_trans` empty; `df_rent_roll` non-empty |
| BR-18 | **Net Credit Computation:** The posted credit for a unit is computed as the **net sum** of all Transaction List rows for that unit — positive original credits plus negative reversal rows. A unit whose concession was posted and then fully reversed shall have a net posted amount of `$0`, not the gross positive total. This net figure is used in R6 and R7 comparisons. | Error | Transaction List with both original and reversal rows for same unit |
| BR-19 | **R1 — Post-Term Credit:** A credit row shall be flagged if its `Date` is strictly after the `Lease_End` date found for the same `(Property, Unit)` in the New & Renewed Leases lookup (sorted by Lease Start descending to use the most recent lease). Both the credit `Date` and the `Lease_End` must be non-null for the comparison to occur. Risk: HIGH. | Error | Transaction List credit row with non-null `Date`; New & Renewed Leases record with non-null `Lease_End` for same `(Property, Unit)` |
| BR-20 | **R2 — Missing Lease:** A unit with a net-positive credit posted (`net_actual > 0`) shall be flagged if its Rent Roll `Lease_End` is a date that is strictly before today's normalized date. A `Lease_End` that is `None`, `NaN`, or in the future is treated as an active lease and shall not trigger this rule. Risk: HIGH. | Error | Transaction List net-positive credit; Rent Roll record with `Lease_End` < today |
| BR-21 | **R3 — Large Credit Threshold:** Every individual non-reversal Transaction List credit row where `Amount >= $700` (`CONCESSION_CRITICAL_AMT`) shall produce a separate flag. This check operates on individual row amounts, not net totals, so that each discrete large credit is captured. Risk: HIGH. Note: multi-month recurring full-offset scenarios that implicitly produce large amounts are separately captured at CRITICAL by the NER engine. | Error | Transaction List row with `Amount >= 700` and `Is_Reversal = False` |
| BR-22 | **R4 — Non-Standard Description: DISABLED** (March 10, 2026). Confirmed that ResMan `Description` values are freeform identifiers with no enforced standard rule. The approved codes (`CONR`, `CRTCO`, `EMPL`, `MCCR`, `RRFee`) exist in a separate column not present in CSV exports. This rule generated false positives and was permanently disabled. No flags are produced. | N/A | N/A — rule is disabled |
| BR-23 | **R5 — Missing Addendum: DISABLED** (April 30, 2026). Lease addenda are stored as PDFs in ResMan's Documents section, which is not accessible from CSV exports. The presence of a Rent Roll concession row is not a reliable proxy for addendum existence — all flagged addenda were confirmed to exist on manual review. This rule generated only false positives and was permanently disabled. | N/A | N/A — rule is disabled |
| BR-24 | **R6 — Concession Amount Mismatch:** A mismatch flag is produced when a unit appears in **both** the Rent Roll concession lookup and the Transaction List net-credit lookup, and the absolute dollar delta between the two amounts exceeds **both** `$10` AND `10%` of the Rent Roll approved amount (`delta > 10 AND pct_diff > 0.10`). Both conditions must be true simultaneously; a large dollar delta with a small percentage, or a large percentage with a small dollar amount, shall not trigger the rule. Risk: HIGH. | Error | Rent Roll concession row (negative amount + keyword); Transaction List net credit > 0 for same `(Property, Unit)`; both `approved_amt` and `posted_amt` derivable |
| BR-25 | **R6 — All-Reversed Credit Detail:** When a unit's net posted amount is `$0` because all credits were reversed, the R6 detail message shall explicitly state "all posted credits were reversed — net posted: $0" and direct reviewers to confirm whether the non-posting was intentional. This case still triggers R6 because the Rent Roll shows an approved setup of `> $0` but nothing was effectively posted. | Error | Rent Roll concession > 0; Transaction List credits summing to net $0 after reversals; delta > $10 and > 10% |
| BR-26 | **R7 — Not Properly Posted:** A flag is produced when a unit appears in the Rent Roll concession lookup but **not** in the Transaction List net-credit lookup (i.e., `in_rr = True` and `in_tx = False`). The flag detail shall include the Rent Roll concession description and instruct reviewers to confirm whether this is a missed posting, a proration event, or an NTV situation. Risk: HIGH. | Error | Rent Roll concession row for `(Property, Unit)` with no corresponding Transaction List net credit |

---

### 6.3 Revenue Integrity Engine — Stage 1: Recurring Projection Audit

| ID | Rule | Severity | Requires |
|---|---|---|---|
| BR-27 | **90% Rule — Standard Charge Classification:** A charge category is classified as a standard charge for a property when the count of units with a **positive** amount for that category divided by the total distinct unit count for that property is `≥ 0.90` (`STANDARD_CHARGE_THRESHOLD`). The 90% threshold is evaluated per property independently; a charge standard at one property is not assumed standard at another. | Error | Transaction Projection DataFrame; at least 1 unit with positive amount for the category |
| BR-28 | **90% Rule — Optional Charge Exemption:** Any category whose name contains one of these keywords shall be permanently excluded from the 90% Rule: `"carport"`, `"parking"`, `"pet rent"`, `"pet fee"`, `"washer"`, `"dryer"`, `"first floor"`, `"1st floor"`. These are unit-specific add-ons per the revenue integrity configuration and are not expected portfolio-wide. | Warning | Category name matching any `OPTIONAL_CHARGE_KEYWORDS` entry |
| BR-29 | **90% Rule — Missing Unit Flag:** Every unit in the property that lacks a positive amount for a standard charge category shall receive a separate `Missing Standard Charge` flag. The `Amount_Impact` is set to the statistical mode amount for that charge at the property, representing the monthly revenue at risk. Risk: HIGH. | Error | Property with ≥ 1 unit missing a category that is standard at 90%+ of other units |
| BR-30 | **Amount Variance — Grouping Scope:** Amount consistency is evaluated within each `(Property, Unit_Type, Category)` group. Comparing within the same unit type prevents false positives when 1BR and 2BR units legitimately have different charge amounts for the same category. Groups with fewer than 3 active (positive-amount) rows are excluded from variance analysis to avoid statistical noise on small populations. | Warning | `(Property, Unit_Type, Category)` group with ≥ 3 positive-amount rows |
| BR-31 | **Amount Variance — Mode as Baseline:** The baseline expected amount for a `(Property, Unit_Type, Category)` group is the statistical mode of all positive-amount rows in that group. If the mode is `$0` or the mode computation yields an empty result, the group is skipped. | Warning | ≥ 3 positive-amount rows in group; mode not $0 |
| BR-32 | **Major Charge Amount Variance:** A unit's charge amount is flagged as a Major variance when `abs(actual - mode) >= $5.00` AND the percentage deviation `(delta / mode) >= 20%`. Both conditions must be true simultaneously. Risk: HIGH. | Error | `(Property, Unit_Type, Category)` group meeting sample size minimum; charge is not in `RENT_KEYWORDS` or `OPTIONAL_CHARGE_KEYWORDS` |
| BR-33 | **Minor Charge Amount Variance:** A unit's charge amount is flagged as a Minor variance when `abs(actual - mode) >= $1.00` and the variance does not also satisfy both thresholds for Major (i.e., is not simultaneously ≥ $5 and ≥ 20%). Risk: MEDIUM. | Warning | Same conditions as BR-32 |
| BR-34 | **Rent Exclusion from Variance:** Any category whose lowercased name is exactly `"rent"` shall be excluded from both Major and Minor variance checks. Rent amounts legitimately differ between units of the same type due to individually negotiated lease rates. Rent discrepancies are fully captured by Stage 2's Posted vs Recurring Mismatch rule. | Warning | Category name exactly `"rent"` (case-insensitive) |
| BR-35 | **Recurring Concession > $700:** A concession row in the Transaction Projection for the audit month where `abs(Amount) > $700` (`CONCESSION_CRITICAL_AMT`) shall produce a flag. Concession rows are identified by keyword match against the `Category` field: `"concession"`, `"conr"`, `"crtco"`, `"empl"`, `"mccr"`, `"rrfee"`, `"employee unit"`, `"resident referral"`, `"courtesy officer"`. Concession amounts are stored negative in the projection; `abs()` is applied before all threshold comparisons. Risk: HIGH. | Error | Transaction Projection with concession-category row; `abs(Amount) > 700` for audit month |
| BR-36 | **Concession Exactly $500:** A concession row where `abs(abs(Amount) - 500) < $1.00` (i.e., the amount is within $1 of exactly $500) shall be flagged independently, because an exact $500 concession is a specific signal that documentation and addendum verification is required. Risk: HIGH. | Error | Transaction Projection concession row; `abs(Amount)` within $1.00 of $500 |
| BR-37 | **Concession > $500 for More Than 2 Months:** A recurring concession setup where `abs(Amount) > $500` and the count of months with a positive concession amount exceeds `2` (`CONCESSION_HIGH_MONTHS`) shall be flagged. This detects multi-month concession arrangements that require documented lease addenda. Risk: HIGH. | Error | Transaction Projection concession row spanning > 2 months; `abs(Amount) > 500` |

---

### 6.4 Revenue Integrity Engine — Stage 1B: NER (Net Effective Rent) Engine

| ID | Rule | Severity | Requires |
|---|---|---|---|
| BR-38 | **NER Engine Activation Guard:** The NER engine executes only when the multi-month projection DataFrame (`df_proj_full`) is non-empty. If it is empty, all Stage 1B checks are silently skipped. | Warning | `df_proj_full` non-empty; multi-month Transaction Projection CSVs present |
| BR-39 | **Rent vs Concession Row Classification:** Within the multi-month projection, a row is classified as a **rent row** if its `Category` (lowercased) contains `\brent\b` via regex but does not contain `"concession"`. A row is classified as a **concession row** if its `Category` contains any concession keyword. A row may not qualify for both classifications. | Warning | Multi-month projection with clearly labeled Category values |
| BR-40 | **NER Floor — Hardcoded Priority:** The NER floor for a `(Property, bedroom_type)` pair is resolved first from `PROPERTY_NER_FLOORS`. If the value found is not `None`, it is used directly as the floor. Hardcoded floors: Parks on Taylor 1BR = $799, 2BR = $899; Highland Park 1BR = $799, 2BR = $999; Valencia Plaza 1BR = $999. | Error | `PROPERTY_NER_FLOORS` constant; bedroom type derivable from `Unit_Type` |
| BR-41 | **NER Floor — Dynamic Market-Relative Computation:** If `PROPERTY_NER_FLOORS[property][bedroom_type]` is `None`, the system checks `PROPERTY_NER_DISCOUNT` for the property. If a discount exists, the system looks up the unit's total market rent in the Market Rent Schedule dictionary (`(property, unit)` key). If found and `> 0`, the floor is computed as `market_rent - discount`. Dynamic floors: Village Green = market − $300; Crossings at Irving = market − $100; La Prada = market − $100. Western Station has no applicable rule (floor varies by floorplan) and is skipped. | Error | `PROPERTY_NER_DISCOUNT` constant; Market Rent Schedule loaded with entry for `(property, unit)` |
| BR-42 | **Bedroom Type Derivation:** A unit type code is mapped to a bedroom label (`1BR`, `2BR`, `3BR`) using two sequential checks: (1) if any numeric digit in the code is `≥ 3`, classify as `3BR`; (2) use the first letter of the code: `A` → `1BR`, `B` → `2BR`, `C`/`D` → `3BR`. Any code not matching these patterns returns `"Unknown"` and the NER floor lookup will yield `None` (no NER check performed). | Warning | `Unit_Type` value present in Transaction Projection row |
| BR-43 | **Full Rent Recurring Offset — CRITICAL:** For a given unit and month, the **Net Effective Rent** is computed as `rent_by_month + conc_by_month` (concession stored negative → NER = rent − |concession|). If `abs(NER) < $1.00` (net rent is effectively zero), that month is classified as a full-offset month. If **2 or more** full-offset months exist across the projection, the unit receives a `Full Rent Recurring Offset` flag. `Amount_Impact = rent_amount × number_of_full_offset_months`. Risk: CRITICAL. | Error | Multi-month projection; ≥ 2 months where `abs(rent + concession) < 1.00` |
| BR-44 | **Future Month Full Offset — HIGH:** If exactly **1** full-offset month exists for a unit, the unit receives a `Future Month Full Offset` flag. `Amount_Impact = $0.00` (the financial event has not yet materialized). The detail message distinguishes between the case where the zero-NER month is the current `AUDIT_MONTH` (described as "current audit month — verify addendum") and the case where it is a future month (described as "upcoming month — confirm addendum before it posts"). Risk: HIGH. | Warning | Multi-month projection; exactly 1 month where `abs(rent + concession) < 1.00` |
| BR-45 | **NER Proration Guard:** Before applying NER floor checks for the current audit month, the system retrieves the unit's `Market_Rent` from the Rent Roll. If `market_rent > 0` AND `rent_by_month[AUDIT_MONTH] < market_rent × 0.50`, the NER check is skipped for that unit. A recurring rent below 50% of market rent is interpreted as a partial-month move-in or move-out proration, not a structural concession problem. | Warning | Rent Roll `Market_Rent` > 0; `rent_by_month[AUDIT_MONTH]` < 50% of market rent |
| BR-46 | **Double-Discount Setup — CRITICAL:** For the current audit month, if (a) the unit's NER is below its floor, AND (b) the recurring rent is already below the Rent Roll market rent, AND (c) `abs(concession) > $1.00` (a meaningful concession is also in place), the unit receives a `Double-Discount Setup` flag. `Amount_Impact = floor − NER`. The detail shall state the market rent, recurring rent, concession, resulting NER, and floor, and instruct the reviewer to obtain VP authorization or correct the setup so NER ≥ floor. Risk: CRITICAL. | Error | NER < floor; `rent < market_rent`; `abs(conc) > 1.00`; NER floor resolvable |
| BR-47 | **Net Effective Rent Below Floor — CRITICAL:** For the current audit month, if the unit's NER is below its floor AND the setup does not qualify as a double-discount (i.e., recurring rent is not below market rent, or no concession is present), the unit receives a `Net Effective Rent Below Floor` flag. `Amount_Impact = floor − NER`. Risk: CRITICAL. | Error | NER > 0; NER < floor; does not meet both conditions for Double-Discount |
| BR-48 | **NER Per-Unit Deduplication:** A set (`_ner_flagged`) is maintained during each run. Once a `(Property, Unit)` has been emitted for any of the three CRITICAL NER sub-rules (Full Rent Recurring Offset, Double-Discount Setup, Net Effective Rent Below Floor), it is added to the set and excluded from subsequent NER sub-rule evaluation for the same run. This prevents the same underlying concession anomaly from producing multiple overlapping CRITICAL flags for the same unit. | Warning | Any prior NER flag emitted for the same `(Property, Unit)` in the current run |

---

### 6.5 Revenue Integrity Engine — Stage 2: Posted Rent Roll Audit

| ID | Rule | Severity | Requires |
|---|---|---|---|
| BR-49 | **Occupied Unit Filter:** Stage 2 evaluates only Rent Roll rows where `Status` is one of `"C"` (Current), `"MTM"` (Month-to-Month), or `"NTV"` (Notice to Vacate). All other status values (vacant, applicant, model, etc.) are excluded before any Stage 2 rule is applied. | Warning | Rent Roll `Status` column populated |
| BR-50 | **Move-In Date Lookup Chain:** The move-in date for $0 Net Rent classification is resolved in priority order: (1) `Move_In` from the Resident Activity report for the matching `(Property, Unit)`; (2) if absent, `Lease Start` from the New & Renewed Leases report. If neither source yields a date, the unit cannot be classified as a recent move-in and the CRITICAL $0 Net Rent (Not Recent) rule applies. | Warning | Resident Activity or New & Renewed Leases record for `(Property, Unit)` |
| BR-51 | **Net Rent Computation:** For each occupied unit on the Rent Roll, net rent is computed as the sum of `Amount` values for rows where `Description` (lowercased) matches `\brent\b` or `\bbase\b` (regex) AND `Amount > 0`. Concession rows — which contain `"concession"` in the description — are explicitly excluded from this sum even if they also contain `"rent"` (e.g., `"Concession - Rent"`). | Error | Rent Roll charge rows with `Description` matching rent/base keywords |
| BR-52 | **Negative Net Rent — CRITICAL:** If the computed net rent for an occupied unit is `< 0` (concession rows sum to more than rent rows), the unit receives a `Negative Net Rent` flag. `Amount_Impact = abs(net_rent)`. Risk: CRITICAL. | Error | Rent Roll occupied unit; `net_rent < 0` after positive-only rent row summation |
| BR-53 | **$0 Net Rent — Recent Move-In Threshold:** If net rent is exactly `$0` and the resolved move-in date is within `60 days` (`RECENT_MOVEIN_DAYS`) of the system run date (`pd.Timestamp.today()`), the unit receives a `$0 Net Rent (Recent Move-in)` flag. Risk: MEDIUM. The threshold of 60 days captures the first full calendar month after move-in, during which a legitimate one-month-free concession may still be in effect. | Warning | Occupied unit; `net_rent == 0`; move-in date resolvable and within 60 days of run date |
| BR-54 | **$0 Net Rent — Not Recent — CRITICAL:** If net rent is exactly `$0` and the resident is not within the 60-day recent-move-in window (either because no move-in date is available or the date is more than 60 days ago), the unit receives a `$0 Net Rent (Not Recent)` flag. The detail directs reviewers to confirm whether the unit is a courtesy officer, employee, or model unit, and if not, to investigate for an unauthorized full-offset posting. Risk: CRITICAL. | Error | Occupied unit; `net_rent == 0`; move-in date either missing or > 60 days before run date |
| BR-55 | **Manual Posting Without Setup:** The system identifies all `(Property, Unit)` pairs from the Transaction List with a net-positive non-reversal credit. It then checks whether each of those units has any row in the Transaction Projection. If a unit has posted credits but no projection entry at all, it receives a `Manual Posting Without Setup` flag. `Amount_Impact = net_credit_amount`. Risk: HIGH. | Error | Transaction List net-positive credit; Transaction Projection loaded; `(Property, Unit)` absent from projection |
| BR-56 | **Posted vs Recurring Mismatch — Positive-Only Filters:** Both sides of the comparison use positive-amount-only filters. `posted_rent` = sum of Rent Roll rows matching `\brent\b` or `\bbase\b` AND `Amount > 0`. `recurring_rent` = sum of Transaction Projection rows matching `\brent\b` or `\bbase\b` AND `Amount > 0` AND **not** containing `"concession"` in the Category. The exclusion of concession rows from the recurring side prevents a `"Concession - Rent"` entry in the projection from artificially reducing `recurring_rent` below `posted_rent`. | Error | Both Rent Roll and Transaction Projection non-empty; unit present in both datasets |
| BR-57 | **Posted vs Recurring Mismatch — Bidirectional Proration Guard:** Before comparing `posted_rent` and `recurring_rent`, the system applies: `if min(posted_rent, recurring_rent) < max(posted_rent, recurring_rent) × 0.60: skip`. This 60% threshold handles both directions of proration false positives — a new move-in where the projection carries a prorated small amount (projection << Rent Roll), and a move-out where the Rent Roll carries a prorated final charge (Rent Roll << projection). | Warning | Either `posted_rent` or `recurring_rent` is less than 60% of the other |
| BR-58 | **Posted vs Recurring Mismatch — Dollar Threshold:** After the proration guard, a mismatch flag is produced only when `abs(recurring_rent − posted_rent) > $5.00`. Variances of $5 or less are treated as rounding noise and suppressed. `Amount_Impact = abs(variance)`. Risk: HIGH. | Error | Both `posted_rent > 0` and `recurring_rent > 0`; proration guard not triggered; `abs(delta) > 5.00` |
| BR-59 | **Misc Tenant Credit — Keyword Detection:** Every non-reversal, positive-amount Transaction List row whose `Description` (lowercased) contains any of `"misc"`, `"miscellaneous"`, `"adjustment"`, `"write-off"`, `"write off"`, `"reclass"`, or `"mccr"` shall produce an individual `Misc Tenant Credit` flag. Each row produces a separate flag record; no aggregation is performed. Risk: HIGH. | Warning | Transaction List row with misc keyword in `Description`; `Is_Reversal = False`; `Amount > 0` |

---

### 6.6 Fee Schedule Check Engine

| ID | Rule | Severity | Requires |
|---|---|---|---|
| BR-60 | **Property Schedule Availability Guard:** If a property has no entry in `PROPERTY_FEE_SCHEDULE`, the system shall print `[SKIP] <property> — no fee schedule loaded` and skip all units for that property. Currently, La Prada is the only property without a fee schedule. | Warning | Property name present in `PROPERTY_FEE_SCHEDULE` constant |
| BR-61 | **Charge Matching by Keyword:** Each fee in the property's schedule is matched to projection rows using a substring search: the fee's `keywords` list is checked against the lowercased `Category` string for each unit row. If no row matches a fee's keywords, no action is taken (missing charges are the responsibility of the 90% Rule, not the Fee Schedule Check). | Warning | Transaction Projection row with `Category` matching fee keyword; `Amount > 0` |
| BR-62 | **Fee Amount Variance Threshold — $1:** A `Fee Schedule Violation` flag is produced when `abs(actual_amount − scheduled_amount) >= $1.00`. Variances below $1 are suppressed. This threshold is the official cutoff for charge amount comparisons. `Amount_Impact = variance`. Risk: HIGH. | Error | Matching projection row with `Amount > 0`; `abs(actual − scheduled) >= 1.00` |
| BR-63 | **Optional Fee Multi-Unit Exemption:** For fees marked `optional: True` in the schedule, if the resident's posted amount is an exact whole-number multiple of the per-unit scheduled rate (`actual_amount % scheduled_amount == 0.00`, rounded to 2 decimal places), the fee is not flagged. This allows residents with multiple optional units (e.g., 3 parking spaces at $35/each = $105) to pass validation without producing a false positive. | Warning | Fee marked `optional: True`; `actual_amount % scheduled_amount == 0.00` |
| BR-64 | **Negative Amount Exclusion:** Any Transaction Projection row with a negative `Amount` (a concession or credit entry) is excluded from all fee schedule comparisons. Only positive-amount charge rows are validated against the schedule. | Warning | Projection row with `Amount <= 0` |

---

### 6.7 Financial Exposure Calculation Rules

| ID | Rule | Severity | Requires |
|---|---|---|---|
| BR-65 | **Engine Output Concatenation:** All flags from the Concession Audit Engine, Revenue Integrity Engine, and the Fee Schedule Check Engine are concatenated into a single `all_flags` DataFrame before exposure is calculated. If all three engines produce empty DataFrames, exposure calculation returns a dict of empty DataFrames rather than raising an error. | Warning | At least one engine returning a non-empty flags DataFrame |
| BR-66 | **Deduped Exposure Computation:** The conservative (deduped) exposure figure is computed as: `df.groupby(["Property", "Unit"])["Amount_Impact"].max().sum()`. This takes the single highest-impact flag per unit across all engines before summing across all units, preventing the same financial event (e.g., one posted credit flagged by both Concession Audit R5 and Revenue Integrity Engine Manual Posting rules) from being double-counted in the reported exposure. | Error | `all_flags` non-empty; `Property`, `Unit`, and `Amount_Impact` columns present |
| BR-67 | **Raw Exposure vs Deduped Exposure Distinction:** The system retains both `Total_Exposure` (raw sum of all `Amount_Impact` values across every flag row) and `Deduped_Exposure` (per the rule above). `Total_Exposure` is the upper ceiling and `Deduped_Exposure` is the conservative floor. The Streamlit dashboard prominently displays `Deduped_Exposure` in the Financial Exposure KPI tile and labels `Total_Exposure` as the raw sum in the caption. | Warning | Both values present in `exposure["totals"]` row |
| BR-68 | **Risk-Level Sort Order for Excel:** All flag DataFrames written to the `All Exceptions` Excel sheet are sorted first by risk priority (`CRITICAL = 0`, `HIGH = 1`, `MEDIUM = 2`) and then by `Amount_Impact` descending within each tier. This ensures the highest-severity, highest-dollar-impact findings appear at the top of the workbook sheet that reviewers open first. | Warning | `all_flags` non-empty; `Risk_Level` and `Amount_Impact` columns present |
| BR-69 | **Status Default Value:** Every flag row written to an Excel flag sheet is prepended with `Status = "Open"` via `_add_review_columns()`. No flag shall be written to the workbook with any pre-set resolution status. All flags start as unreviewed. | Warning | Flag DataFrame non-empty before `export_to_excel()` is called |
| BR-70 | **`Amount_Impact` Coercion:** Before exposure aggregation, all `Amount_Impact` values are cast with `pd.to_numeric(..., errors="coerce").fillna(0)`. Any flag record whose `Amount_Impact` cannot be cast to a number is treated as $0 rather than causing an aggregation failure. | Warning | `all_flags` DataFrame with `Amount_Impact` column |

---

## 7. Non-Functional Requirements

### 7.1 Performance

- **End-to-end runtime target:** The full audit pipeline — ingesting all 49 CSV files across 7 properties, running all three engines, computing exposure, and writing the Excel workbook — shall complete in under 60 seconds on a standard developer workstation. Observed runtimes for the June 2026 full run (479 transaction rows, 7,283 rent roll rows, 7,262 projection rows, 1,135 market-rent-schedule units, 686 total flags) confirmed completion well within this threshold.
- **Pandas I/O efficiency:** All loaders iterate CSVs using a single `pd.read_csv()` call with `dtype=str` to minimise type-inference overhead. Encoding fallback (UTF-8-sig → CP1252 → Latin-1) retries the same file with the next codec rather than re-opening the filesystem entry unnecessarily.
- **openpyxl column-width sampling cap:** The `_format_flag_sheet()` function samples a maximum of 200 data rows per column when computing auto-fit widths (`for row_idx in range(2, min(max_row + 1, 200))`). This prevents the formatting pass from becoming a bottleneck on large workbooks while still producing accurate column widths for typical audit outputs.
- **In-memory processing only:** No intermediate files are written between pipeline stages. All DataFrames are held in memory and passed directly between functions within the same `run_full_audit()` call. There is no disk I/O between ingestion and export.
- **`warnings.filterwarnings("ignore")`:** Pandas performance and dtype warnings are suppressed globally to keep console output clean and scannable for `[OK]`, `[WARN]`, and `[ERROR]` log entries.
- **Streamlit session state:** The dashboard stores the full `run_full_audit()` result dict in `st.session_state` after the first run. Subsequent tab switches and filter interactions do not re-execute the engine; they re-slice the already-computed DataFrames in memory.

### 7.2 Data Freshness

- **Manual export cadence:** All source data is derived from ResMan CSV exports copied manually into `data/` subfolders before each audit run. There is no live database connection, scheduled data pull, or API integration. Data freshness is entirely dependent on when the operator copies the export files.
- **One run = one audit month:** The constant `AUDIT_MONTH` (e.g. `"Jun 2026"`) must be manually updated in `audit_bot.py` at the start of each audit cycle. The system does not auto-detect the current month. Running the bot without updating this constant will produce results scoped to the prior month's projection column.
- **ResMan account access requirement:** Transaction List CSVs must be exported from a **full-access ResMan account** (i.e., an account with permission to view Credit transaction sections). Exports from limited-access accounts omit Credit rows entirely, causing the Concession Audit Engine R1–R3 and Revenue Integrity Engine Stage 2 rules that depend on credit data to produce no flags for genuine violations. This was confirmed during the June 2026 audit when the initial set of Transaction List files from a limited-access account contained zero Credit sections.
- **Market Rent Schedule currency:** The Market Rent Schedule CSVs must reflect the fee rates in effect for the audit month. If a property has changed its market rents or amenity premiums mid-period, the NER floor computations for units on dynamic floors (Village Green, Crossings at Irving, La Prada) will reflect the most recent schedule loaded, not a historical snapshot.

### 7.3 Security

- **No API keys, tokens, or credentials in the codebase:** The system contains no secrets, authentication tokens, database connection strings, or environment variables. All `AUDIT_MONTH`, threshold, and fee schedule constants are hardcoded values representing business rules, not credentials.
- **No `.env` file or `secrets.toml`:** Neither a `.env` file nor a Streamlit `secrets.toml` is present in the repository. No secrets management tooling (e.g. python-dotenv, AWS Secrets Manager) is used or required for current functionality.
- **No external network calls at runtime:** The bot's audit pipeline makes zero network requests. The only external URL in the codebase is the sidebar property icon in `app.py` (`https://img.icons8.com/fluency/96/property.png`), which is a cosmetic image load by the browser rendering the Streamlit UI. The audit logic itself is entirely offline.
- **Local filesystem access only:** All input data is read from local `data/` subfolders and all output is written to the local `output/` directory. No data is transmitted to external services, cloud storage, or third-party APIs.
- **No Streamlit authentication layer:** The dashboard has no login screen, session management, or role-based access control. Access is controlled entirely by who can reach the machine running the Streamlit process (localhost:8501 by default). If the app is ever deployed beyond a local workstation, an authentication proxy or Streamlit Community Cloud authentication must be added.
- **Virtual environment isolation:** The Python runtime uses a `.venv` isolated virtual environment (`include-system-site-packages = false`) running Python 3.11.9. This prevents dependency conflicts with system-level Python packages and limits the installed package surface area to only what the project requires.
- **ResMan CSV data classification:** The CSV files loaded into the bot contain personally identifiable information (resident names) and financially sensitive data (rent amounts, concession details, lease dates). These files should be treated as confidential and should not be committed to version control repositories. The `data/` folder is populated manually from local downloads and should be excluded from `.gitignore`.
- **No SQL or shell injection risk:** The system performs no database queries and executes no shell commands. All data operations are Python/pandas in-memory operations on CSV content. There is no user-supplied input that is evaluated or passed to a system call.

### 7.4 Auditability

- **Timestamped Excel output:** Every run produces a uniquely named Excel workbook (`LNJ_Audit_YYYYMMDD_HHMM.xlsx`) in the `output/` directory. Multiple runs within the same audit cycle produce distinct files, preserving a full history of every engine execution without overwriting prior results.
- **Source file traceability:** Every flag record includes a `Source_File` field containing the exact CSV filename from which the flagging data was derived. This allows any exception to be traced back to the specific ResMan export file and row that triggered it.
- **Console audit log:** The `run_full_audit()` function prints a structured console log at each stage: file-level `[OK]` confirmations with row counts, `[WARN]` messages for missing folders or skipped files, `[ERROR]` messages for file parse failures, and a summary line with total units, exceptions, exposure, and CRITICAL flag count.
- **Resolution workflow in Excel:** All flag sheets include `Status` (dropdown: Open / Reviewed / Cleared / Escalated) and `Notes` columns. These allow the reviewing accountant to document the disposition of each finding directly in the workbook, creating a per-run audit trail of how exceptions were resolved.
- **Override Detail Log:** The full Edited Transactions by User event log — including manager login, event type (Reversal vs Amount Change), original amount, edited amount, and computed revenue impact — is preserved as a dedicated Excel sheet and surfaced in the Manager Overrides dashboard tab. This provides a complete record of all manager-initiated transaction modifications.
- **Engine flag attribution:** Every flag record's `Rule` field identifies exactly which audit rule fired. The `Risk_Level` field identifies severity. Together with `Source_File`, these three fields allow any finding to be attributed to a specific engine, rule, and input file without ambiguity.

### 7.5 Configurability

- **`AUDIT_MONTH`** (type: `str`, format: `"Mon YYYY"`): The single constant that scopes every run to a calendar month. Must be updated manually in `audit_bot.py` before each audit cycle. No command-line arguments or config file mechanism exists; the constant is the only supported configuration point for month selection.
- **`PROPERTY_FEE_SCHEDULE`** (type: `dict[str, list[dict]]`): The official monthly fee amounts per property. Updated by the developer when Daniel Twito provides revised fee sheet `.docx` files. Changes require editing the constant directly in `audit_bot.py`; there is no external fee schedule file or database table.
- **`PROPERTY_NER_FLOORS`** (type: `dict[str, dict[str, int | None]]`): Hardcoded minimum NER by property and bedroom type. `None` values indicate properties where NER floors are dynamic (computed from Market Rent Schedule) or not yet defined. Requires developer update when LiveNjoy policy changes.
- **`PROPERTY_NER_DISCOUNT`** (type: `dict[str, int]`): The maximum allowed discount from market rent for properties whose NER floor is market-relative. Currently set for Village Green ($300), Crossings at Irving ($100), and La Prada ($100). Requires developer update when policy changes.
- **`APPROVED_CODES`** (type: `set[str]`): The set of ResMan transaction codes accepted as valid concession types (`CONR`, `CRTCO`, `EMPL`, `MCCR`, `RRFee`). Hardcoded. No runtime override mechanism.
- **`STANDARD_CHARGE_THRESHOLD`** (type: `float`, value: `0.90`): The 90% occupancy threshold above which a charge category is classified as standard and required on all units. Hardcoded constant.
- **`RECENT_MOVEIN_DAYS`** (type: `int`, value: `60`): The window in calendar days within which a $0-net-rent unit is treated as a recent move-in (MEDIUM risk) rather than an unexplained zero-rent unit (CRITICAL). Hardcoded constant.
- **`CONCESSION_CRITICAL_AMT`** (type: `float`, value: `700.0`) and **`CONCESSION_HIGH_AMT`** (type: `float`, value: `500.0`)**: Dollar thresholds for Large Credit and multi-month concession flags. Hardcoded constants.
- **`DIRS`** (type: `dict[str, str]`): All seven input folder paths are derived at runtime from `BASE_DIR = os.path.dirname(os.path.abspath(__file__))`. This makes the project folder-portable — moving the project directory does not require updating any paths, as long as the relative `data/` subfolder structure is preserved.
- **No external configuration file:** There is no `config.yaml`, `.env`, `settings.ini`, or command-line argument parser in the project. All business rule parameters live as module-level constants in `audit_bot.py`. Changing any threshold requires a direct code edit and a rerun.

---

## 8. Assumptions & Open Items

### 8.1 Technical Assumptions Made During Development

- **ResMan CSV format stability:** The bot's seven loaders are built around specific ResMan export layouts — fixed skip-row counts (5–6 header rows), positional column indexing, and section-marker strings (e.g., `"Recurring Transactions by Unit"`). It is assumed that ResMan does not change these export formats without notice. A ResMan platform update that alters the column order, header row count, or section marker text would silently break one or more loaders without raising an exception, as the parsers use positional indexing rather than named-column lookups.
- **Concession storage convention:** It is assumed that LiveNjoy's ResMan instance stores approved concessions as **negative-amount line items on the Rent Roll** rather than in the `Rec. Conc.` field of the New & Renewed Leases report. This is a LiveNjoy-specific configuration; the `Rec. Conc.` column in the Leases export is `$0` for all rows in practice. The Concession Audit Engine and Revenue Integrity Engine Stage 2 both rely on Rent Roll concession rows as the source of truth for approved concession amounts. If LiveNjoy's ResMan configuration changes to populate `Rec. Conc.` directly, the concession lookup logic in both engines would need to be refactored.
- **Transaction List Credit exclusivity:** It is assumed that the Transaction List CSVs contain **only Credit-type sections** when the bot needs them, or that the section forward-fill logic correctly classifies and excludes all non-Credit sections (Payment, Deposit, Charge). The current regex `r"^(Credit|Charge|Payment|Deposit)"` is assumed to cover all section header prefixes ResMan may emit. A new section type introduced by ResMan that does not begin with one of these four words would be missed by the forward-fill and could leak into credit lookups.
- **Seven-property fixed portfolio:** The system assumes a fixed set of exactly seven properties. The `derive_property()` function maps filenames to a hardcoded property name via short-code and keyword maps. Adding or renaming a property requires updating `CODE_MAP`, `KEYWORD_MAP`, `PROPERTY_FEE_SCHEDULE`, `PROPERTY_NER_FLOORS`, and `PROPERTY_NER_DISCOUNT` simultaneously in the source code.
- **Unit number as the join key:** All cross-report lookups use `(Property, Unit)` as the composite key. It is assumed that unit numbers are stable identifiers within ResMan — i.e., a unit number refers to the same physical unit across all seven report types for the same audit period. Unit renaming or renumbering in ResMan would produce missed matches and incorrect flags without any error indication.
- **Occupied-status classification:** Revenue Integrity Engine Stage 2 filters to `Status` values of `"C"`, `"MTM"`, and `"NTV"`. It is assumed these are the only three status codes that represent revenue-generating occupied units in LiveNjoy's ResMan configuration. Courtesy officer units or employee units with a different status code would be silently excluded from Stage 2 checks.
- **Market Rent Schedule column position:** The Market Rent Schedule loader reads `"Total Rent"` from column index 7 (the 8th column). This is assumed to be a fixed positional column in the ResMan Market Rent Schedule Detail export. If ResMan adds or removes a column before index 7, the loader will silently read the wrong value.
- **Concession amounts are negative in projection CSVs:** It is assumed that ResMan always exports concession amounts as negative values in the Transaction Projection CSV. The NER engine and Stage 1 concession checks both apply `abs()` before threshold comparisons. If ResMan changes the sign convention, all concession-related calculations would invert silently.
- **`python-docx` dependency is dormant:** The `python-docx` package is listed as installed in the virtual environment (noted in CONTEXT.md) but is not imported or used anywhere in the current `audit_bot.py` or `app.py` codebase. It was likely installed during an earlier iteration when fee schedules were being parsed from `.docx` files. All fee schedule data has since been hardcoded into `PROPERTY_FEE_SCHEDULE`.

### 8.2 Pending Integration Dependencies

- **ResMan Document API (Lease Addenda):** Rule R5 (Missing Addendum) was **disabled on April 30, 2026** because lease addendum PDFs are stored in ResMan's Documents module, which is not accessible from CSV exports. There is no current data source that allows the bot to confirm whether a signed addendum exists for a given unit. **Open item:** If ResMan exposes a Documents API or if an addendum tracker spreadsheet is provided, R5 can be re-enabled using that source as the addendum lookup. Until then, the `if in_tx and not in_rr` block remains commented out.
- **Incorrect Frequency Setup rule (unimplemented):** A planned rule that checks whether a concession's recurring frequency setup matches the approved term (one-time, 12-month, MTM) was designed but never implemented. It requires a data source showing the approved concession term per unit — this is not present in any of the seven current ResMan CSV export types. **Open item:** A manual tracker for approved concession frequencies has not been received as of July 2026. The rule remains documented in CONTEXT.md as a future requirement.
- **NER Floors — Highland Park and Valencia Plaza confirmation:** The NER floors for Highland Park (1BR: $799, 2BR: $999) and Valencia Plaza (1BR: $999) are marked `⚠️ needs re-confirmation` in CONTEXT.md. These values were entered based on initial guidance but were flagged as potentially needing revision. **Open item:** These floors must be formally confirmed or corrected before the values in `PROPERTY_NER_FLOORS` can be treated as verified policy.
- **Western Station NER floor (not defined):** Western Station has no NER floor defined in either `PROPERTY_NER_FLOORS` or `PROPERTY_NER_DISCOUNT`. CONTEXT.md notes that "floor varies by floorplan — not a simple market-minus rule." **Open item:** A per-floorplan NER floor matrix for Western Station must be provided. Until received, all Western Station units are excluded from NER floor checks.
- **Live ResMan API integration:** The current architecture is entirely file-based. All seven CSV types must be manually exported from ResMan and copied into `data/` subfolders before each run. **Open item:** If LiveNjoy obtains access to the ResMan REST API or a scheduled data export mechanism, the seven `_csv_files()` / `_read_csv()` loader pairs could be replaced with API calls, eliminating the manual export step and enabling fully automated monthly runs.
- **Automated scheduling:** There is no scheduled or triggered execution of the audit. The bot is run manually via `python audit_bot.py` or via the Streamlit sidebar button. **Open item:** A Windows Task Scheduler job, a cloud scheduler (e.g. AWS EventBridge), or a CI/CD pipeline step could be configured to automate monthly execution once the data export step is also automated.
- **Cloud or shared deployment of the Streamlit dashboard:** The dashboard currently runs exclusively on `localhost:8501` on the developer's machine. There is no production deployment to a shared server, intranet host, or Streamlit Community Cloud. **Open item:** If leadership requires self-service access to the dashboard without running the bot locally, a deployment target (Streamlit Community Cloud, Azure App Service, or an internal server) must be selected, and an authentication mechanism must be added since the current app has no login protection.
- **`requirements.txt` / dependency pinning:** The project has no `requirements.txt` or `pyproject.toml` pinning package versions. The virtual environment was created with specific versions (e.g. Streamlit 1.54.0) but these are not documented in a reproducible lockfile. **Open item:** A `pip freeze > requirements.txt` or a `pyproject.toml` with pinned versions should be committed to the repository to ensure reproducible installs across machines and future environments.
- **Resident Ledger PDF parsing:** Resident Ledger PDFs are stored in `exports/Resident Ledgers/` but are explicitly documented as manual-review-only. The bot prints a note in the Streamlit sidebar: "⚠️ Resident Ledgers are PDFs — not processed by the bot." **Open item:** If automated ledger analysis is required in a future version, a PDF extraction library (e.g., `pdfplumber`, `PyMuPDF`) would need to be integrated, along with a parser for ResMan's specific ledger layout.
- **Multi-month trend reporting:** The current system audits one month per run. The multi-month projection data (`df_proj_full`) is loaded and used internally for NER calculations but is not surfaced as a trend view in the dashboard or Excel output. **Open item:** A future version could add a month-over-month comparison view showing which flags are new, which are recurring violations, and how total portfolio exposure has trended across audit periods.
