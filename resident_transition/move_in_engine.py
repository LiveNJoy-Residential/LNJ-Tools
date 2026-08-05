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


MI4_RENT_TOLERANCE = 25.00
MI4_CRITICAL_DIFF = 100.00
MI2_ACTIVITY_MISSING_GRACE_DAYS = 3


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


def _is_vacant_placeholder(name: str) -> bool:
    """Detect standard Rent Roll placeholders used before physical move-in."""
    s = str(name).strip().lower()
    return s in {"vacant unit", "vacant-unleased", "vacant unleased", "vacant"} or s.startswith("vacant")


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
    new_leases = df_leases[df_leases["Lease_Type"] == "New"].drop_duplicates(subset=["Property", "Unit"], keep="first").copy()

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
        elif _is_vacant_placeholder(rr_resident):
            # Pre-move-in placeholder; do not treat as identity mismatch noise.
            continue
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
    new_leases = df_leases[df_leases["Lease_Type"] == "New"].drop_duplicates(subset=["Property", "Unit"], keep="first").copy()

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
        if pd.notna(sign_date) and pd.notna(ls_start) and sign_date > ls_start:
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
            # Skip pending move-ins — Activity won't have them until they physically arrive
            if pd.notna(ls_start) and ls_start > pd.Timestamp.today().normalize():
                continue

            sev = "HIGH"
            if pd.notna(ls_start):
                days_since_start = (pd.Timestamp.today().normalize() - ls_start.normalize()).days
                if days_since_start >= MI2_ACTIVITY_MISSING_GRACE_DAYS:
                    sev = "CRITICAL"
            flags.append(_flag(
                prop, unit, resident,
                "MI-2", "Date & Timeline Mismatch",
                "Move-In Date in Activity Report",
                "Not found", "Present",
                f"Unit {unit} is in the New Leases report but not found in Resident Activity. "
                "Move-in may not have been recorded in ResMan.",
                sev, src,
            ))
            continue

        act_move_in    = act.get("Move_In")
        act_lease_start = act.get("Lease_Start")
        act_lease_end   = act.get("Lease_End")

        # Move-In date should be close to Lease Start (allow up to 5 days for proration)
        if pd.notna(act_move_in) and pd.notna(ls_start):
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
        if pd.notna(act_lease_start) and pd.notna(ls_start) and act_lease_start.date() != ls_start.date():
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
        if pd.notna(act_lease_end) and pd.notna(ls_end) and act_lease_end.date() != ls_end.date():
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
    new_leases = df_leases[df_leases["Lease_Type"] == "New"].drop_duplicates(subset=["Property", "Unit"], keep="first").copy()

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
            if abs(rr_rent - lease_rent) > MI4_RENT_TOLERANCE:
                # If there is a Concession - Rent in the RR, the effective rent may equal the lease
                conc_rows = unit_charges[
                    unit_charges["Description"].str.lower().str.contains("concession", na=False) &
                    (unit_charges["Amount"] < 0)
                ]
                effective_rent = rr_rent + conc_rows["Amount"].sum()
                if abs(effective_rent - lease_rent) <= MI4_RENT_TOLERANCE:
                    continue  # Concession explains the gap — setup is correct

                gap = abs(rr_rent - lease_rent)
                sev = "CRITICAL" if gap >= MI4_CRITICAL_DIFF else "HIGH"
                flags.append(_flag(
                    prop, unit, resident,
                    "MI-4", "Initial Financial Setup Mismatch",
                    "Rent Charge Amount",
                    f"${rr_rent:,.2f}", f"${lease_rent:,.2f}",
                    f"Rent Roll shows ${rr_rent:,.2f} but the lease says ${lease_rent:,.2f} "
                    f"(difference: ${gap:,.2f}). "
                    "Must be corrected before move-in.",
                    sev, src,
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
    new_leases = df_leases[df_leases["Lease_Type"] == "New"].drop_duplicates(subset=["Property", "Unit"], keep="first").copy()

    if new_leases.empty or df_rr_units.empty:
        return flags

    rr_deposit_lookup = df_rr_units.set_index(["Property", "Unit"])["Deposits"].to_dict()

    for _, row in new_leases.iterrows():
        prop      = row["Property"]
        unit      = row["Unit"]
        resident  = row["Residents"]
        src       = row["Source_File"]
        ls_start  = row["Lease_Start"]

        # Skip future move-ins — deposit isn't expected before the resident physically arrives
        if pd.notna(ls_start) and ls_start > pd.Timestamp.today().normalize():
            continue

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
                    "MEDIUM", src,
            ))

    return flags


# ===========================================================================
# Main runner
# ===========================================================================

# ===========================================================================
# MI-6: Auxiliary Billing Verification (Pets & Vehicles)
# ===========================================================================

def run_mi6(df_pets: pd.DataFrame, df_vehicles: pd.DataFrame,
            df_recurring: pd.DataFrame, df_rr_units: pd.DataFrame = None,
            df_rr_charges: pd.DataFrame = None,
            df_leases: pd.DataFrame = None) -> list:
    """
    FR-4.1: Units with pets must have pet rent charged (checked against Rent Roll
    charges for current accuracy; falls back to Recurring Projection if RR unavailable).
    Units with assigned carport/reserved permits must have a parking charge.
    """
    flags = []
    if df_recurring.empty:
        return flags

    PET_KEYWORDS  = ["pet rent", "pet fee", "animal fee"]
    PARK_KEYWORDS = ["parking", "garage", "carport", "car port", "reserved"]

    # MI-6 belongs to the move-in workflow: only evaluate units from New leases.
    move_in_set: set = set()
    if df_leases is not None and not df_leases.empty and "Lease_Type" in df_leases.columns:
        _new = df_leases[df_leases["Lease_Type"] == "New"].drop_duplicates(subset=["Property", "Unit"], keep="first")
        move_in_set = set(zip(_new["Property"], _new["Unit"]))

    if not df_pets.empty:
        # Build set of currently-occupied units from Rent Roll — skip ghost/vacated records
        rr_current_set: set = set()
        if df_rr_units is not None and not df_rr_units.empty:
            _today = pd.Timestamp.today().normalize()
            _status = df_rr_units["Status"].astype(str).str.strip().str.lower()
            _current = df_rr_units[
                (df_rr_units["Move_Out"].isna() | (df_rr_units["Move_Out"] > _today)) &
                df_rr_units["Status"].notna() &
                (~_status.isin(["", "nan", "none", "nat"]))
            ]
            rr_current_set = set(zip(_current["Property"], _current["Unit"]))

        _ESA_KEYWORDS = ["esa", "emotional support", "service animal"]
        for (prop, unit), grp in df_pets.groupby(["Property", "Unit"]):
            if move_in_set and (prop, unit) not in move_in_set:
                continue
            # Skip units not in the current Rent Roll (former residents, unclean records)
            if rr_current_set and (prop, unit) not in rr_current_set:
                continue
            src      = grp["Source_File"].iloc[0]
            resident = grp["Owner"].iloc[0]
            pets = ", ".join(sorted(set(str(p).strip() for p in grp["Pet_Name"].tolist() if str(p).strip())))

            # Skip units where every pet is an ESA / emotional support animal
            if "Reg_Type" in grp.columns:
                reg_types = grp["Reg_Type"].fillna("").astype(str).str.lower()
            else:
                reg_types = pd.Series([""] * len(grp), index=grp.index)
            if "Pet_Type" in grp.columns:
                pet_types = grp["Pet_Type"].fillna("").astype(str).str.lower()
            else:
                pet_types = pd.Series([""] * len(grp), index=grp.index)
            pet_names_l = grp["Pet_Name"].fillna("").astype(str).str.lower()

            esa_markers = reg_types + " " + pet_types + " " + pet_names_l
            all_esa = esa_markers.apply(lambda r: any(kw in r for kw in _ESA_KEYWORDS)).all()
            if all_esa:
                continue
            # Use Rent Roll charges (current) when available; fall back to Recurring Projection (future)
            if df_rr_charges is not None and not df_rr_charges.empty:
                _unit_ch = df_rr_charges[
                    (df_rr_charges["Property"] == prop) & (df_rr_charges["Unit"] == unit)
                ]
                has_pet = _unit_ch["Description"].str.lower().apply(
                    lambda d: any(kw in d for kw in PET_KEYWORDS)
                ).any()
            else:
                unit_rec = df_recurring[
                    (df_recurring["Property"] == prop) & (df_recurring["Unit"] == unit)
                ]
                has_pet = (
                    not unit_rec.empty and
                    unit_rec["Description"].str.lower().apply(
                        lambda d: any(kw in d for kw in PET_KEYWORDS)
                    ).any()
                )
            if not has_pet:
                flags.append(_flag(
                    prop, unit, resident,
                    "MI-6", "Auxiliary Billing Verification",
                    "Pet Rent Charge",
                    "Not in Recurring Projection", "Required",
                    f"Unit {unit} has pet(s) on file ({pets}) but no pet rent charge is in "
                    "the Recurring Transaction Projection. Verify pet rent is set up in ResMan.",
                    "MEDIUM", src,
                ))

    if not df_vehicles.empty:
        assigned = df_vehicles[
            df_vehicles["Permit_Number"].notna() &
            (df_vehicles["Permit_Number"].str.strip() != "") &
            (~df_vehicles["Permit_Number"].str.upper().isin(["OPEN", "NAN", ""]))
        ]
        if not assigned.empty:
            # Only check properties where at least one unit already has a parking charge.
            # If no units pay for parking in the recurring data, the property doesn't bill for it.
            props_with_parking = set()
            if not df_recurring.empty:
                parking_mask = df_recurring["Description"].str.lower().apply(
                    lambda d: any(kw in d for kw in PARK_KEYWORDS)
                )
                props_with_parking = set(df_recurring[parking_mask]["Property"].unique())

            # Permit prefixes that indicate a paid carport / reserved spot
            _PAID_PREFIXES = ("cp-", "cp ", "reserved", "covered", "grg")

            for (prop, unit), grp in assigned.groupby(["Property", "Unit"]):
                if move_in_set and (prop, unit) not in move_in_set:
                    continue
                if prop not in props_with_parking:
                    continue  # property doesn't charge for parking

                # Only check permits that look like a paid carport/reserved spot
                paid_permits = grp[
                    grp["Permit_Number"].str.lower().apply(
                        lambda p: any(p.startswith(pfx) for pfx in _PAID_PREFIXES) or
                                  any(kw in p for kw in ("reserved", "covered", "carport"))
                    )
                ]
                if paid_permits.empty:
                    continue  # only free surface-parking permit numbers — skip
                src      = paid_permits["Source_File"].iloc[0]
                resident = paid_permits["Resident"].iloc[0]
                permits  = ", ".join(paid_permits["Permit_Number"].tolist())
                unit_rec = df_recurring[
                    (df_recurring["Property"] == prop) & (df_recurring["Unit"] == unit)
                ]
                has_parking = (
                    not unit_rec.empty and
                    unit_rec["Description"].str.lower().apply(
                        lambda d: any(kw in d for kw in PARK_KEYWORDS)
                    ).any()
                )
                if not has_parking:
                    flags.append(_flag(
                        prop, unit, resident,
                        "MI-6", "Auxiliary Billing Verification",
                        "Parking Charge",
                        "Not in Recurring Projection", "Required",
                        f"Unit {unit} has assigned permit(s) ({permits}) but no parking charge "
                        "is in the Recurring Transaction Projection. Verify parking fee is set "
                        "up in ResMan.",
                        "MEDIUM", src,
                    ))

    return flags


# ===========================================================================
# Main runner
# ===========================================================================

def run_move_in_audit(df_leases: pd.DataFrame,
                      df_activity: pd.DataFrame,
                      df_rr_units: pd.DataFrame,
                      df_rr_charges: pd.DataFrame,
                      df_pets: pd.DataFrame = None,
                      df_vehicles: pd.DataFrame = None,
                      df_recurring: pd.DataFrame = None) -> pd.DataFrame:
    """
    Run all Phase 1 Move-In audit rules against the loaded DataFrames.
    Returns a DataFrame of all exceptions found.
    """
    if df_leases.empty:
        print("  [WARN] No lease data loaded — Move-In audit skipped.")
        return pd.DataFrame()

    new_count = (df_leases["Lease_Type"] == "New").sum() if "Lease_Type" in df_leases else 0
    print(f"\n--- Move-In Audit ({new_count} new leases) ---")

    _pets      = df_pets      if df_pets      is not None else pd.DataFrame()
    _vehicles  = df_vehicles  if df_vehicles  is not None else pd.DataFrame()
    _recurring = df_recurring if df_recurring is not None else pd.DataFrame()

    flags = []
    flags.extend(run_mi1(df_leases, df_rr_units))
    flags.extend(run_mi2(df_leases, df_activity))
    flags.extend(run_mi4(df_leases, df_rr_charges))
    flags.extend(run_mi5(df_leases, df_rr_units))
    flags.extend(run_mi6(
        _pets,
        _vehicles,
        _recurring,
        df_rr_units=df_rr_units,
        df_rr_charges=df_rr_charges,
        df_leases=df_leases,
    ))

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
