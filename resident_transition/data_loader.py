"""
CSV loaders for the three ResMan exports used by the Resident Activity Audit Tool:
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

try:
    from .utils import (
        _csv_files, _read_csv, derive_property,
        clean_currency, clean_unit, parse_date, clean_name,
    )
except ImportError:
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


def _iter_sources(files, folder, dir_key):
    """Yield (filename, source) for uploaded file objects or disk folder. files=None → disk."""
    if files is not None:
        for f in files:
            if hasattr(f, "seek"):
                f.seek(0)
            yield f.name, f
    else:
        folder = folder or DIRS.get(dir_key, "")
        for fname in _csv_files(folder):
            yield fname, os.path.join(folder, fname)


_KNOWN_PROPS = {
    "Crossings at Irving", "Highland Park", "La Prada",
    "Parks on Taylor", "Village Green", "Valencia Plaza", "Western Station",
}


def _get_prop(fname: str, src) -> str:
    """Derive property from filename; fall back to first CSV row when filename lacks it."""
    prop = derive_property(fname)
    if prop not in _KNOWN_PROPS:
        try:
            if hasattr(src, "seek"):
                src.seek(0)
            row1 = _read_csv(src, nrows=1, header=None, dtype=str)
            if hasattr(src, "seek"):
                src.seek(0)
            prop = derive_property(str(row1.iloc[0, 0]).strip()) or prop
        except Exception:
            pass
    return prop


def _iter_section_rows(src, skiprows: int, col_sentinel: str):
    """
    Yield (prop, row_dict) from single or multi-property ResMan CSVs.
    prop is None for single-property files — caller should use _get_prop() as fallback.
    """
    raw = _read_csv(src, skiprows=skiprows, header=None, dtype=str)
    prop = None
    col_names = None

    for _, row in raw.iterrows():
        c0 = str(row.iloc[0]).strip() if len(row) > 0 else ""

        if c0 in ("", "nan"):
            continue
        if "total" in c0.lower() and not re.match(r"^\d", c0):
            continue
        if c0 in ("Current", "Resident", "Holding Units"):
            continue

        if c0 == col_sentinel:
            col_names = [str(v).strip() for v in row.tolist()]
            continue

        maybe_prop = derive_property(c0)
        if maybe_prop in _KNOWN_PROPS:
            prop = maybe_prop
            col_names = None
            continue

        if col_names and re.match(r"^\d+", c0):
            yield prop, dict(zip(col_names, [str(v).strip() for v in row.tolist()]))


# ===========================================================================
# 1. New & Renewed Leases
# ===========================================================================

def load_leases(folder: str = None, files: list = None) -> pd.DataFrame:
    """
    ResMan New & Renewed Leases:
      Rows 1-5  : property header block (skip)
      Row 6     : column headers
      Row 7+    : section labels ('New Leases', 'Renewed Leases') + data rows

    Returns one row per lease.
    Adds Lease_Type = 'New' | 'Renewed' to distinguish move-ins from renewals.
    """
    all_data = []

    for fname, src in _iter_sources(files, folder, "leases"):
        try:
            raw = _read_csv(src, skiprows=5, header=None, dtype=str)
            fallback_prop = _get_prop(fname, src)
            prop = fallback_prop
            col_names = None
            lease_type = "New"
            records = []

            for _, row in raw.iterrows():
                c0 = str(row.iloc[0]).strip() if len(row) > 0 else ""
                if c0 in ("", "nan"):
                    continue
                if "total" in c0.lower() and not re.match(r"^\d", c0):
                    continue
                maybe_prop = derive_property(c0)
                if maybe_prop in _KNOWN_PROPS:
                    prop = maybe_prop
                    col_names = None
                    lease_type = "New"
                    continue
                if c0 == "Unit":
                    col_names = [str(v).strip() for v in row.tolist()]
                    continue
                if "renewed" in c0.lower():
                    lease_type = "Renewed"
                    continue
                if "new leases" in c0.lower():
                    lease_type = "New"
                    continue
                if not (col_names and re.match(r"^\d+", c0)):
                    continue
                r = dict(zip(col_names, [str(v).strip() for v in row.tolist()]))
                records.append({
                    "Property":      prop,
                    "Unit":          clean_unit(c0),
                    "Unit_Type":     r.get("Unit Type", ""),
                    "Residents":     clean_name(r.get("Residents", "Unknown")),
                    "Leasing_Agent": r.get("Leasing Agent", ""),
                    "App_Date":      parse_date(r.get("Application / Renewal Date", "")),
                    "Sign_Date":     parse_date(r.get("Lease Signed Date", "")),
                    "Lease_Start":   parse_date(r.get("Lease Start Date", "")),
                    "Lease_End":     parse_date(r.get("Lease End Date", "")),
                    "Prior_Rent":    clean_currency(r.get("Prior Rent", 0)),
                    "Market_Rent":   clean_currency(r.get("Market Rent", 0)),
                    "Rent":          clean_currency(r.get("Rent", 0)),
                    "Rec_Conc":      clean_currency(r.get("Rec. Conc.", 0)),
                    "One_Time_Conc": clean_currency(r.get("One Time Conc.", 0)),
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

def load_activity(folder: str = None, files: list = None) -> pd.DataFrame:
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
    all_data = []

    for fname, src in _iter_sources(files, folder, "activity"):
        try:
            raw  = _read_csv(src, skiprows=5, header=None, dtype=str)
            fallback_prop = _get_prop(fname, src)
            prop = fallback_prop

            # Drop the 'Adjusted Lease End Date' overflow header row
            raw = raw[~raw.iloc[:, 0].astype(str).str.contains("Adjusted", na=False)]

            while len(raw.columns) < 50:
                raw[f"_pad_{len(raw.columns)}"] = np.nan

            records = []
            for _, row in raw.iterrows():
                c0 = str(row.iloc[0]).strip()
                maybe_prop = derive_property(c0)
                if maybe_prop in _KNOWN_PROPS:
                    prop = maybe_prop
                    continue
                if not re.match(r"^\d+$", c0):
                    continue
                unit        = clean_unit(c0)
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

def load_rent_roll(folder: str = None, files: list = None):
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
    unit_records   = []
    charge_records = []

    for fname, src in _iter_sources(files, folder, "rent_rolls"):
        try:
            raw  = _read_csv(src, skiprows=5, header=None, dtype=str)
            fallback_prop = _get_prop(fname, src)
            prop = fallback_prop

            # Pad to at least 36 columns
            while len(raw.columns) < 36:
                raw[f"_pad_{len(raw.columns)}"] = np.nan

            current = {}

            for _, row in raw.iterrows():
                c0  = str(row.iloc[0]).strip()
                c18 = str(row.iloc[18]).strip() if len(row) > 18 else ""

                # Property name rows in multi-property portfolio files
                maybe_prop = derive_property(c0)
                if maybe_prop in _KNOWN_PROPS:
                    prop = maybe_prop
                    current = {}
                    continue

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


# ===========================================================================
# 4. Scheduled Move Ins
# ===========================================================================

def load_scheduled_move_ins(folder: str = None, files: list = None) -> pd.DataFrame:
    """ResMan Scheduled Move Ins: 5 header rows, then column headers, then data."""
    all_data = []
    for fname, src in _iter_sources(files, folder, "scheduled_move_ins"):
        try:
            fallback_prop = _get_prop(fname, src)
            records = []
            for _prop, r in _iter_section_rows(src, 5, "Unit"):
                prop = _prop or fallback_prop
                records.append({
                    "Property":      prop,
                    "Unit":          clean_unit(r.get("Unit", "")),
                    "Unit_Type":     r.get("Unit Type", ""),
                    "Residents":     clean_name(r.get("Residents", "")),
                    "App_Date":      parse_date(r.get("Application Date", "")),
                    "Approval_Date": parse_date(r.get("Approval Date", "")),
                    "Sign_Date":     parse_date(r.get("Lease Signed Date", "")),
                    "Move_In_Date":  parse_date(r.get("Move In Date", "")),
                    "Market_Rent":   clean_currency(r.get("Market Rent", 0)),
                    "Rent_Charges":  clean_currency(r.get("Rent Charges", 0)),
                    "Other_Charges": clean_currency(r.get("Other Charges", 0)),
                    "Credits":       clean_currency(r.get("Credits", 0)),
                    "Source_File":   fname,
                })

            if records:
                all_data.append(pd.DataFrame(records))
                print(f"  [OK] Scheduled Move Ins: {fname}  ({len(records)} rows)")
        except Exception as e:
            print(f"  [ERROR] Scheduled Move Ins {fname}: {e}")

    return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()


# ===========================================================================
# 5. Cancellations, Denials, and Move Outs
# ===========================================================================

def load_cancellations_move_outs(folder: str = None, files: list = None) -> pd.DataFrame:
    """
    ResMan Cancellations, Denials, and Move Outs:
      5 header rows + 1 blank, then three sections (Cancelled / Denied / Moved Out)
      each preceded by a repeated column header row and a section-label row.
    Parsed positionally since column headers repeat per section.
    """
    all_data = []
    for fname, src in _iter_sources(files, folder, "cancellations"):
        try:
            raw = _read_csv(src, skiprows=5, header=None, dtype=str)
            prop = _get_prop(fname, src)

            section = None
            records = []
            for _, row in raw.iterrows():
                c0 = str(row.iloc[0]).strip() if len(row) > 0 else ""

                if c0 in ("", "nan"):
                    continue
                maybe_prop = derive_property(c0)
                if maybe_prop in _KNOWN_PROPS:
                    prop = maybe_prop
                    section = None
                    continue
                if c0 in ("Cancelled", "Denied", "Moved Out"):
                    section = c0
                    continue
                if c0 == "Unit":        # repeated column header row
                    continue
                if "total" in c0.lower():
                    continue
                if not re.match(r"^\d+", c0):
                    continue

                records.append({
                    "Property":        prop,
                    "Unit":            clean_unit(c0),
                    "Flag":            str(row.iloc[1]).strip() if len(row) > 1 else "",
                    "Residents":       clean_name(row.iloc[2]) if len(row) > 2 else "",
                    "Days_Occupied":   str(row.iloc[3]).strip() if len(row) > 3 else "",
                    "Broke_Lease":     str(row.iloc[4]).strip() if len(row) > 4 else "",
                    "Days_Notice":     str(row.iloc[5]).strip() if len(row) > 5 else "",
                    "Date_Vacated":    parse_date(row.iloc[6]) if len(row) > 6 else None,
                    "Times_Late":      str(row.iloc[7]).strip() if len(row) > 7 else "",
                    "Times_NSF":       str(row.iloc[8]).strip() if len(row) > 8 else "",
                    "Market_Rent":     clean_currency(row.iloc[9]) if len(row) > 9 else 0.0,
                    "Actual_Rent":     clean_currency(row.iloc[10]) if len(row) > 10 else 0.0,
                    "Move_Out_Reason": str(row.iloc[11]).strip() if len(row) > 11 else "",
                    "Section":         section,
                    "Source_File":     fname,
                })

            if records:
                all_data.append(pd.DataFrame(records))
                print(f"  [OK] Cancellations: {fname}  ({len(records)} rows)")
        except Exception as e:
            print(f"  [ERROR] Cancellations {fname}: {e}")

    return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()


# ===========================================================================
# 6. Eviction Process
# ===========================================================================

def load_eviction_process(folder: str = None, files: list = None) -> pd.DataFrame:
    """ResMan Eviction Process: 6 header rows (property name appears twice), then columns."""
    all_data = []
    for fname, src in _iter_sources(files, folder, "evictions"):
        try:
            fallback_prop = _get_prop(fname, src)
            records = []
            for _prop, r in _iter_section_rows(src, 5, "Unit"):
                prop = _prop or fallback_prop
                records.append({
                    "Property":            prop,
                    "Unit":                clean_unit(r.get("Unit", "")),
                    "Residents":           clean_name(r.get("Resident Name", "")),
                    "Status":              r.get("Status", ""),
                    "Lease_Start":         parse_date(r.get("Lease Start", "")),
                    "Lease_End":           parse_date(r.get("Lease End", "")),
                    "Delinquency":         clean_currency(r.get("Delinquency", 0)),
                    "Rent_Delinquency":    clean_currency(r.get("Rent Delinquency", 0)),
                    "Eviction_Filed_Date": parse_date(r.get("Eviction Filed Date", "")),
                    "Move_Out_Date":       parse_date(r.get("Move Out Date", "")),
                    "Delinquency_Notes":   r.get("Delinquency Notes", ""),
                    "Source_File":         fname,
                })

            if records:
                all_data.append(pd.DataFrame(records))
                print(f"  [OK] Eviction Process: {fname}  ({len(records)} rows)")
        except Exception as e:
            print(f"  [ERROR] Eviction Process {fname}: {e}")

    return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()


# ===========================================================================
# 7. Pet Summary
# ===========================================================================

def load_pet_summary(folder: str = None, files: list = None) -> pd.DataFrame:
    """ResMan Pet Summary: 4 header rows, then column headers. One row per pet."""
    all_data = []
    for fname, src in _iter_sources(files, folder, "pet_summary"):
        try:
            raw = _read_csv(src, skiprows=4, header=None, dtype=str)
            fallback_prop = _get_prop(fname, src)
            prop = fallback_prop

            records = []
            for _, row in raw.iterrows():
                c0 = str(row.iloc[0]).strip() if len(row) > 0 else ""
                if c0 in ("", "nan", "Pet Name"):
                    continue
                maybe_prop = derive_property(c0)
                if maybe_prop in _KNOWN_PROPS:
                    prop = maybe_prop
                    continue
                unit_raw = str(row.iloc[1]).strip() if len(row) > 1 else ""
                if not re.match(r"^\d+", unit_raw):
                    continue
                pet_name = c0
                if "no pet" in pet_name.lower() or "nopet" in pet_name.lower():
                    continue

                records.append({
                    "Property":    prop,
                    "Unit":        clean_unit(unit_raw),
                    "Pet_Name":    pet_name,
                    "Owner":       clean_name(str(row.iloc[3]).strip() if len(row) > 3 else ""),
                    "Pet_Type":    str(row.iloc[5]).strip() if len(row) > 5 else "",
                    "Breed":       str(row.iloc[6]).strip() if len(row) > 6 else "",
                    "Reg_Type":    str(row.iloc[8]).strip() if len(row) > 8 else "",
                    "Source_File": fname,
                })

            if records:
                all_data.append(pd.DataFrame(records))
                print(f"  [OK] Pet Summary: {fname}  ({len(records)} rows)")
        except Exception as e:
            print(f"  [ERROR] Pet Summary {fname}: {e}")

    return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()


# ===========================================================================
# 8. Vehicles
# ===========================================================================

def load_vehicles(folder: str = None, files: list = None) -> pd.DataFrame:
    """
    ResMan Vehicles: 4 header rows + 1 'Resident' section label, then column headers.
    Only Status='C' (current) registrations are kept. Notes rows are filtered out.
    """
    all_data = []
    for fname, src in _iter_sources(files, folder, "vehicles"):
        try:
            fallback_prop = _get_prop(fname, src)
            records = []
            for _prop, r in _iter_section_rows(src, 4, "Unit"):
                prop   = _prop or fallback_prop
                status = r.get("Status", "").strip().upper()
                if status != "C":
                    continue

                records.append({
                    "Property":       prop,
                    "Unit":           clean_unit(r.get("Unit", "")),
                    "Resident":       clean_name(r.get("Resident", "")),
                    "Status":         status,
                    "Year":           r.get("Year", ""),
                    "Make":           r.get("Make", ""),
                    "Model":          r.get("Model", ""),
                    "License_Plate":  r.get("License Plate", ""),
                    "Permit_Number":  r.get("Permit Number", ""),
                    "Lease_End_Date": parse_date(r.get("Lease End Date", "")),
                    "Source_File":    fname,
                })

            if records:
                all_data.append(pd.DataFrame(records))
                print(f"  [OK] Vehicles: {fname}  ({len(records)} rows)")
        except Exception as e:
            print(f"  [ERROR] Vehicles {fname}: {e}")

    return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()


# ===========================================================================
# 9. Recurring Transaction Projections
# ===========================================================================

def load_recurring_projections(folder: str = None, files: list = None) -> pd.DataFrame:
    """
    ResMan Recurring Transaction Projection — Section 3 'Recurring Transactions by Unit'.
    Returns one row per (unit, charge category). Used for MI-6 auxiliary billing checks.
    Handles both single-property and multi-property portfolio files.
    """
    all_data = []
    for fname, src in _iter_sources(files, folder, "recurring"):
        try:
            raw = _read_csv(src, header=None, dtype=str)
            fallback_prop = _get_prop(fname, src)
            prop = fallback_prop
            in_unit_section = False
            col_names = None
            records = []

            for _, row in raw.iterrows():
                c0 = str(row.iloc[0]).strip() if len(row) > 0 else ""

                maybe_prop = derive_property(c0)
                if maybe_prop in _KNOWN_PROPS:
                    prop = maybe_prop
                    in_unit_section = False
                    col_names = None
                    continue

                if c0 == "Recurring Transactions by Unit":
                    in_unit_section = True
                    col_names = None
                    continue

                if c0.startswith("Recurring Transactions by"):
                    in_unit_section = False
                    col_names = None
                    continue

                if not in_unit_section:
                    continue

                if col_names is None and c0 not in ("", "nan"):
                    col_names = [str(v).strip() for v in row.tolist()]
                    continue

                if col_names and re.match(r"^\d", c0):
                    unit_num = clean_unit(c0)
                    category = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ""
                    if not category or category.lower() in ("nan", ""):
                        continue
                    records.append({
                        "Property":    prop,
                        "Unit":        unit_num,
                        "Description": category,
                        "Source_File": fname,
                    })

            if records:
                all_data.append(pd.DataFrame(records))
                print(f"  [OK] Recurring Projections: {fname}  ({len(records)} charge rows)")
        except Exception as e:
            print(f"  [ERROR] Recurring {fname}: {e}")

    return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()
