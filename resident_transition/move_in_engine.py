"""
Move-In Audit Engine — Phase 1
==============================
Rules:
  MI-1  Profile & Identity Discrepancy
  MI-2  Date & Timeline Mismatch
  MI-4  Initial Financial Setup Mismatch
  MI-5  Deposit Collection Alert

  MI-3 (Document Upload Compliance) is Phase 2 — requires ResMan Document Attachment export.

Source: Partha Balakrishnan — Move-In Audit Report Tool FRD v1.0 (July 22, 2026)
"""

import re
import pandas as pd


# ---------------------------------------------------------------------------
# Flag builder
# ---------------------------------------------------------------------------

def _flag(prop, unit, resident, rule_id, rule_name, field,
          resman_val, expected_val, detail, severity, source_file) -> dict:
    return {
        "Property":       prop,
        "Unit":           unit,
        "Resident":       resident,
        "Rule_ID":        rule_id,
        "Rule_Name":      rule_name,
        "Field":          field,
        "ResMan_Value":   str(resman_val),
        "Expected_Value": str(expected_val),
        "Detail":         detail,
        "Severity":       severity,
        "Status":         "Open",
        "Source_File":    source_file,
    }


def _normalize_name(name: str) -> str:
    """Lowercase, strip spaces, sort comma-separated names for comparison."""
    parts = [p.strip().lower() for p in str(name).split(",")]
    return ", ".join(sorted(p for p in parts if p))


# ===========================================================================
# MI-1: Profile & Identity Discrepancy
# ===========================================================================

def run_mi1(df_leases: pd.DataFrame, df_rr_units: pd.DataFrame) -> list:
    """
    FR-4.1 / Section 7: Unit number and resident name must match between
    New Leases and Rent Roll.

    Checks:
      - Unit present in Rent Roll
      - Resident name in Leases matches Rent Roll (normalized comparison)
    """
    flags = []
    new_leases = df_leases[df_leases["Lease_Type"] == "New"].copy()

    if new_leases.empty:
        return flags

    rr_lookup = (
        df_rr_units.set_index(["Property", "Unit"])["Residents"].to_dict()
        if not df_rr_units.empty else {}
    )

    for _, row in new_leases.iterrows():
        prop     = row["Property"]
        unit     = row["Unit"]
        resident = row["Residents"]
        src      = row["Source_File"]

        rr_resident = rr_lookup.get((prop, unit))

        if rr_resident is None:
            flags.append(_flag(
                prop, unit, resident,
                "MI-1", "Profile & Identity Discrepancy",
                "Unit in Rent Roll",
                "Not found", "Present",
                f"Unit {unit} appears in the New Leases report but is not found in the "
                "Rent Roll. Verify unit setup is complete in ResMan before move-in.",
                "HIGH", src,
            ))
        elif _normalize_name(resident) != _normalize_name(rr_resident):
            flags.append(_flag(
                prop, unit, resident,
                "MI-1", "Profile & Identity Discrepancy",
                "Resident Name",
                rr_resident, resident,
                f"Lease shows '{resident}' but Rent Roll shows '{rr_resident}'. "
                "Route to leasing office for correction. Exact match required.",
                "HIGH", src,
            ))

    return flags


# ===========================================================================
# MI-2: Date & Timeline Mismatch
# ===========================================================================

def run_mi2(df_leases: pd.DataFrame, df_activity: pd.DataFrame) -> list:
    """
    FR-4.1 / Section 7: Lease Start, Lease End, and Move-In dates must be
    consistent across the Leases report and Resident Activity.

    Checks:
      - Unit present in Resident Activity
      - Lease Signed Date is not after Lease Start Date
      - Lease Start in Leases matches Lease Start in Activity
      - Lease End in Leases matches Lease End in Activity
      - Move-In date is within 5 days of Lease Start (allows proration)
    """
    flags = []
    new_leases = df_leases[df_leases["Lease_Type"] == "New"].copy()

    if new_leases.empty:
        return flags

    act_lookup = (
        df_activity.set_index(["Property", "Unit"]).to_dict("index")
        if not df_activity.empty else {}
    )

    for _, row in new_leases.iterrows():
        prop      = row["Property"]
        unit      = row["Unit"]
        resident  = row["Residents"]
        src       = row["Source_File"]
        ls_start  = row["Lease_Start"]
        ls_end    = row["Lease_End"]
        sign_date = row["Sign_Date"]

        # Lease Signed Date must be ≤ Lease Start Date
        if sign_date and ls_start and sign_date > ls_start:
            flags.append(_flag(
                prop, unit, resident,
                "MI-2", "Date & Timeline Mismatch",
                "Lease Signed Date",
                sign_date.strftime("%m/%d/%Y"), ls_start.strftime("%m/%d/%Y"),
                f"Lease was signed on {sign_date.strftime('%m/%d/%Y')}, which is AFTER the "
                f"lease start date ({ls_start.strftime('%m/%d/%Y')}). "
                "Must be corrected before move-in.",
                "HIGH", src,
            ))

        act = act_lookup.get((prop, unit))

        if act is None:
            flags.append(_flag(
                prop, unit, resident,
                "MI-2", "Date & Timeline Mismatch",
                "Move-In Date in Activity Report",
                "Not found", "Present",
                f"Unit {unit} is in the New Leases report but not found in Resident Activity. "
                "Move-in may not have been recorded in ResMan.",
                "HIGH", src,
            ))
            continue

        act_move_in    = act.get("Move_In")
        act_lease_start = act.get("Lease_Start")
        act_lease_end   = act.get("Lease_End")

        # Move-In date should be close to Lease Start (allow up to 5 days for proration)
        if act_move_in and ls_start:
            delta = abs((act_move_in - ls_start).days)
            if delta > 5:
                flags.append(_flag(
                    prop, unit, resident,
                    "MI-2", "Date & Timeline Mismatch",
                    "Move-In Date vs Lease Start Date",
                    act_move_in.strftime("%m/%d/%Y"), ls_start.strftime("%m/%d/%Y"),
                    f"Move-In ({act_move_in.strftime('%m/%d/%Y')}) differs from Lease Start "
                    f"({ls_start.strftime('%m/%d/%Y')}) by {delta} days. "
                    "If proration applies, verify a proration addendum is on file.",
                    "MEDIUM", src,
                ))

        # Lease Start should match between Leases report and Activity
        if act_lease_start and ls_start and act_lease_start.date() != ls_start.date():
            flags.append(_flag(
                prop, unit, resident,
                "MI-2", "Date & Timeline Mismatch",
                "Lease Start Date",
                act_lease_start.strftime("%m/%d/%Y"), ls_start.strftime("%m/%d/%Y"),
                f"Lease Start in Activity ({act_lease_start.strftime('%m/%d/%Y')}) does not "
                f"match Lease Start in Leases report ({ls_start.strftime('%m/%d/%Y')}). "
                "Update ResMan or the lease document so dates match.",
                "HIGH", src,
            ))

        # Lease End should match between Leases report and Activity
        if act_lease_end and ls_end and act_lease_end.date() != ls_end.date():
            flags.append(_flag(
                prop, unit, resident,
                "MI-2", "Date & Timeline Mismatch",
                "Lease End Date",
                act_lease_end.strftime("%m/%d/%Y"), ls_end.strftime("%m/%d/%Y"),
                f"Lease End in Activity ({act_lease_end.strftime('%m/%d/%Y')}) does not "
                f"match Lease End in Leases report ({ls_end.strftime('%m/%d/%Y')}). "
                "Update ResMan or the lease document so dates match.",
                "HIGH", src,
            ))

    return flags


# ===========================================================================
# MI-4: Initial Financial Setup Mismatch
# ===========================================================================

def run_mi4(df_leases: pd.DataFrame, df_rr_charges: pd.DataFrame) -> list:
    """
    FR-4.1 / Section 7: Rent charge in the Rent Roll must match the lease amount.

    Checks:
      - Rent charge row exists in Rent Roll for the unit
      - Rent Roll rent amount matches Leases rent amount (within $1.00 tolerance)
    """
    flags = []
    new_leases = df_leases[df_leases["Lease_Type"] == "New"].copy()

    if new_leases.empty or df_rr_charges.empty:
        return flags

    rr_by_unit = (
        df_rr_charges.groupby(["Property", "Unit"])
        if not df_rr_charges.empty else {}
    )

    for _, row in new_leases.iterrows():
        prop       = row["Property"]
        unit       = row["Unit"]
        resident   = row["Residents"]
        src        = row["Source_File"]
        lease_rent = row["Rent"]

        try:
            unit_charges = rr_by_unit.get_group((prop, unit))
        except KeyError:
            continue  # Unit not in Rent Roll — already caught by MI-1

        # Locate the Rent charge row (positive, not a concession)
        rent_rows = unit_charges[
            unit_charges["Description"].str.lower().str.contains(r"\brent\b", na=False, regex=True) &
            ~unit_charges["Description"].str.lower().str.contains("concession", na=False) &
            (unit_charges["Amount"] > 0)
        ]

        if rent_rows.empty:
            flags.append(_flag(
                prop, unit, resident,
                "MI-4", "Initial Financial Setup Mismatch",
                "Rent Charge",
                "Not set up", f"${lease_rent:,.2f}",
                f"No rent charge row found in Rent Roll for unit {unit}. "
                "Rent must be set up in ResMan before move-in.",
                "CRITICAL", src,
            ))
        else:
            rr_rent = rent_rows["Amount"].iloc[0]
            if abs(rr_rent - lease_rent) > 1.00:
                flags.append(_flag(
                    prop, unit, resident,
                    "MI-4", "Initial Financial Setup Mismatch",
                    "Rent Charge Amount",
                    f"${rr_rent:,.2f}", f"${lease_rent:,.2f}",
                    f"Rent Roll shows ${rr_rent:,.2f} but the lease says ${lease_rent:,.2f} "
                    f"(difference: ${abs(rr_rent - lease_rent):,.2f}). "
                    "Must be corrected before move-in.",
                    "HIGH", src,
                ))

    return flags


# ===========================================================================
# MI-5: Deposit Collection Alert
# ===========================================================================

def run_mi5(df_leases: pd.DataFrame, df_rr_units: pd.DataFrame) -> list:
    """
    FR-4.2 / Section 7: A deposit must be recorded in the Rent Roll for all
    new move-ins. $0 deposit is flagged for immediate correction.
    """
    flags = []
    new_leases = df_leases[df_leases["Lease_Type"] == "New"].copy()

    if new_leases.empty or df_rr_units.empty:
        return flags

    rr_deposit_lookup = df_rr_units.set_index(["Property", "Unit"])["Deposits"].to_dict()

    for _, row in new_leases.iterrows():
        prop     = row["Property"]
        unit     = row["Unit"]
        resident = row["Residents"]
        src      = row["Source_File"]

        deposits = rr_deposit_lookup.get((prop, unit))

        if deposits is None:
            continue  # Unit not in Rent Roll — already flagged by MI-1

        if deposits == 0.0:
            flags.append(_flag(
                prop, unit, resident,
                "MI-5", "Deposit Collection Alert",
                "Deposit Paid",
                "$0.00", "Amount per lease agreement",
                f"No deposit is recorded in the Rent Roll for unit {unit}. "
                "Verify the deposit was collected and posted in ResMan. "
                "Correct immediately — required before move-in.",
                "CRITICAL", src,
            ))

    return flags


# ===========================================================================
# Main runner
# ===========================================================================

def run_move_in_audit(df_leases: pd.DataFrame,
                      df_activity: pd.DataFrame,
                      df_rr_units: pd.DataFrame,
                      df_rr_charges: pd.DataFrame) -> pd.DataFrame:
    """
    Run all Phase 1 Move-In audit rules against the loaded DataFrames.
    Returns a DataFrame of all exceptions found.
    """
    if df_leases.empty:
        print("  [WARN] No lease data loaded — Move-In audit skipped.")
        return pd.DataFrame()

    new_count = (df_leases["Lease_Type"] == "New").sum() if "Lease_Type" in df_leases else 0
    print(f"\n--- Move-In Audit ({new_count} new leases) ---")

    flags = []
    flags.extend(run_mi1(df_leases, df_rr_units))
    flags.extend(run_mi2(df_leases, df_activity))
    flags.extend(run_mi4(df_leases, df_rr_charges))
    flags.extend(run_mi5(df_leases, df_rr_units))

    if not flags:
        print("  No Move-In exceptions found.")
        return pd.DataFrame()

    df_out = pd.DataFrame(flags)
    print(
        f"  Move-In Audit complete: {len(df_out)} exceptions — "
        f"{(df_out['Severity'] == 'CRITICAL').sum()} CRITICAL  "
        f"{(df_out['Severity'] == 'HIGH').sum()} HIGH  "
        f"{(df_out['Severity'] == 'MEDIUM').sum()} MEDIUM"
    )
    return df_out
