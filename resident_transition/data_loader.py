"""
CSV loaders for the three ResMan exports used by the Resident Transition Tool:
  1. New & Renewed Leases
  2. Resident Activity
  3. Rent Roll  (returns two DataFrames: unit-level and charge-level)

Column positions calibrated against June 2026 ResMan exports.
Standalone — no dependency on audit_bot.py.
"""

import os
import re
import numpy as np
import pandas as pd

from utils import (
    _csv_files, _read_csv, derive_property,
    clean_currency, clean_unit, parse_date, clean_name,
)

# ---------------------------------------------------------------------------
# Default data folder paths (shared with parent project)
# ---------------------------------------------------------------------------
_HERE   = os.path.dirname(os.path.abspath(__file__))
_PARENT = os.path.dirname(_HERE)
_DATA   = os.path.join(_PARENT, "data")

DIRS = {
    "leases":     os.path.join(_DATA, "leases"),
    "activity":   os.path.join(_DATA, "activity"),
    "rent_rolls": os.path.join(_DATA, "rent_rolls"),
}


# ===========================================================================
# 1. New & Renewed Leases
# ===========================================================================

def load_leases(folder: str = None) -> pd.DataFrame:
    """
    ResMan New & Renewed Leases:
      Rows 1-5  : property header block (skip)
      Row 6     : column headers
      Row 7+    : section labels ('New Leases', 'Renewed Leases') + data rows

    Returns one row per lease.
    Adds Lease_Type = 'New' | 'Renewed' to distinguish move-ins from renewals.
    """
    folder = folder or DIRS["leases"]
    all_data = []

    for fname in _csv_files(folder):
        src = os.path.join(folder, fname)
        try:
            df = _read_csv(src, skiprows=5, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]
            prop = derive_property(fname)

            lease_type = "New"
            records = []
            for _, row in df.iterrows():
                unit_raw = str(row.get("Unit", "")).strip()

                # Section header rows (not data)
                if "renewed" in unit_raw.lower():
                    lease_type = "Renewed"
                    continue
                if "new leases" in unit_raw.lower() or unit_raw == "":
                    continue
                if not re.match(r"^\d+", unit_raw):
                    continue

                records.append({
                    "Property":      prop,
                    "Unit":          clean_unit(unit_raw),
                    "Unit_Type":     str(row.get("Unit Type", "")).strip(),
                    "Residents":     clean_name(row.get("Residents", "Unknown")),
                    "Leasing_Agent": str(row.get("Leasing Agent", "")).strip(),
                    "App_Date":      parse_date(row.get("Application / Renewal Date", "")),
                    "Sign_Date":     parse_date(row.get("Lease Signed Date", "")),
                    "Lease_Start":   parse_date(row.get("Lease Start Date", "")),
                    "Lease_End":     parse_date(row.get("Lease End Date", "")),
                    "Prior_Rent":    clean_currency(row.get("Prior Rent", 0)),
                    "Market_Rent":   clean_currency(row.get("Market Rent", 0)),
                    "Rent":          clean_currency(row.get("Rent", 0)),
                    "Rec_Conc":      clean_currency(row.get("Rec. Conc.", 0)),
                    "One_Time_Conc": clean_currency(row.get("One Time Conc.", 0)),
                    "Lease_Type":    lease_type,
                    "Source_File":   fname,
                })

            if records:
                all_data.append(pd.DataFrame(records))
                print(f"  [OK] Leases: {fname}  ({len(records)} rows)")
        except Exception as e:
            print(f"  [ERROR] Leases {fname}: {e}")

    return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()


# ===========================================================================
# 2. Resident Activity
# ===========================================================================

def load_activity(folder: str = None) -> pd.DataFrame:
    """
    ResMan Resident Activity (very wide, ~73 columns):
      Rows 1-6  : header block (skip)
      Row 7     : sparse column headers
      Row 8     : 'Adjusted Lease End Date' overflow row — skip
      Rows 9+   : data rows (multiple rows per unit due to multi-staff logging)

    Key column positions (0-based after skiprows=6):
      [0]  Unit        [2]  Residents      [18] Unit Type
      [23] Actual Rent [29] Move In        [32] Initial Lease End
      [37] Lease Start [43] Lease End      (last non-blank > 43 = Manager)

    Returns one row per unit (most recent lease, deduplicated).
    """
    folder = folder or DIRS["activity"]
    all_data = []

    for fname in _csv_files(folder):
        src = os.path.join(folder, fname)
        try:
            raw  = _read_csv(src, skiprows=6, header=0, dtype=str)
            prop = derive_property(fname)

            # Drop the 'Adjusted Lease End Date' overflow header row
            raw = raw[~raw.iloc[:, 0].astype(str).str.contains("Adjusted", na=False)]
            # Keep only unit-number rows
            raw = raw[raw.iloc[:, 0].astype(str).str.strip().str.match(r"^\d+$")]

            while len(raw.columns) < 50:
                raw[f"_pad_{len(raw.columns)}"] = np.nan

            records = []
            for _, row in raw.iterrows():
                unit        = clean_unit(str(row.iloc[0]))
                residents   = clean_name(str(row.iloc[2]))
                unit_type   = str(row.iloc[18]).strip() if len(row) > 18 else ""
                actual_rent = clean_currency(row.iloc[23]) if len(row) > 23 else 0.0
                move_in     = parse_date(row.iloc[29]) if len(row) > 29 else None
                lease_start = parse_date(row.iloc[37]) if len(row) > 37 else None
                lease_end   = parse_date(row.iloc[43]) if len(row) > 43 else None

                manager = "Unknown"
                for i in range(len(row) - 1, 43, -1):
                    v = str(row.iloc[i]).strip()
                    if v not in ("", "nan"):
                        manager = v
                        break

                records.append({
                    "Property":    prop,
                    "Unit":        unit,
                    "Residents":   residents,
                    "Unit_Type":   unit_type,
                    "Actual_Rent": actual_rent,
                    "Move_In":     move_in,
                    "Lease_Start": lease_start,
                    "Lease_End":   lease_end,
                    "Manager":     manager,
                    "Source_File": fname,
                })

            if records:
                df_prop = pd.DataFrame(records)
                # Multiple rows per unit (different staff); keep most recent lease
                df_prop = df_prop.sort_values("Lease_Start", ascending=False, na_position="last")
                df_prop = df_prop.drop_duplicates(subset=["Property", "Unit"], keep="first")
                all_data.append(df_prop)
                print(f"  [OK] Activity: {fname}  ({len(df_prop)} units)")
        except Exception as e:
            print(f"  [ERROR] Activity {fname}: {e}")

    return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()


# ===========================================================================
# 3. Rent Roll
# ===========================================================================

def load_rent_roll(folder: str = None):
    """
    ResMan Rent Roll (hierarchical: unit header row + charge sub-rows + Total row):
      Rows 1-6  : header block (skip)
      Row 7     : sparse column headers
      Rows 8+   : UNIT HEADER rows + CHARGE sub-rows + TOTAL rows

    Key column positions (0-based after skiprows=6):
      [0]  Unit          [2]  Unit Type     [5]  Residents
      [10] Status        [12] Market Rent   [18] Description
      [21] Amount        [25] Move In       [26] Lease Start
      [27] Lease End     [31] Move Out      [34] Deposits
      [35] Balance

    Returns:
      df_units   — one row per unit (unit-level summary)
      df_charges — one row per charge line (flat, with unit info repeated)
    """
    folder = folder or DIRS["rent_rolls"]
    unit_records   = []
    charge_records = []

    for fname in _csv_files(folder):
        src = os.path.join(folder, fname)
        try:
            raw  = _read_csv(src, skiprows=6, header=0, dtype=str)
            prop = derive_property(fname)

            # Pad to at least 36 columns
            while len(raw.columns) < 36:
                raw[f"_pad_{len(raw.columns)}"] = np.nan

            current = {}

            for _, row in raw.iterrows():
                c0  = str(row.iloc[0]).strip()
                c18 = str(row.iloc[18]).strip()

                if c18.lower() == "total":
                    continue

                if re.match(r"^\d+$", c0):
                    # Unit header row
                    current = {
                        "Property":    prop,
                        "Unit":        clean_unit(c0),
                        "Unit_Type":   str(row.iloc[2]).strip(),
                        "Residents":   clean_name(str(row.iloc[5])),
                        "Status":      str(row.iloc[10]).strip(),
                        "Market_Rent": clean_currency(row.iloc[12]),
                        "Move_In":     parse_date(row.iloc[25]),
                        "Lease_Start": parse_date(row.iloc[26]),
                        "Lease_End":   parse_date(row.iloc[27]),
                        "Move_Out":    parse_date(row.iloc[31]) if len(row) > 31 else None,
                        "Deposits":    clean_currency(row.iloc[34]) if len(row) > 34 else 0.0,
                        "Balance":     clean_currency(row.iloc[35]) if len(row) > 35 else 0.0,
                        "Source_File": fname,
                    }
                    unit_records.append(current.copy())

                    # Some units have a charge on the same header row
                    if c18 not in ("", "nan"):
                        charge_records.append({
                            **current,
                            "Description": c18,
                            "Amount":      clean_currency(row.iloc[21]),
                        })

                elif current and c18 not in ("", "nan"):
                    # Charge sub-row — attach to current unit
                    charge_records.append({
                        **current,
                        "Description": c18,
                        "Amount":      clean_currency(row.iloc[21]),
                    })

            unit_count = sum(1 for r in unit_records if r.get("Source_File") == fname)
            print(f"  [OK] Rent Roll: {fname}  ({unit_count} units)")
        except Exception as e:
            print(f"  [ERROR] Rent Roll {fname}: {e}")

    df_units   = pd.DataFrame(unit_records)   if unit_records   else pd.DataFrame()
    df_charges = pd.DataFrame(charge_records) if charge_records else pd.DataFrame()
    return df_units, df_charges
