"""
Eviction Audit Engine — Phase 1
================================
Rules:
  EV-1  Active Eviction with Outstanding Balance
  EV-2  Stalled Eviction — No Move-Out Recorded After 30 Days
  EV-3  Evicted Resident with Balance Ready for Collections

Source: Partha Balakrishnan — FRD v2.0 (July 2026)
"""

import pandas as pd
from datetime import datetime, date, timedelta

STALE_DAYS      = 30   # days after filing before flagging as stalled
COLLECTION_DAYS = 30   # days post move-out before balance should go to collections


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


# ===========================================================================
# EV-1: Active Eviction with Outstanding Balance
# ===========================================================================

def run_ev1(df_ev: pd.DataFrame, audit_date: date = None) -> list:
    """
    Active eviction (UE status, or no/future move-out date) AND delinquency > 0.
    Resident is still in the unit — legal action pending, money at risk.
    """
    if audit_date is None:
        audit_date = datetime.today().date()
    flags = []
    if df_ev.empty:
        return flags

    def _is_active(row):
        status = str(row.get("Status", "")).strip().upper()
        if status == "UE":
            return True
        mo = row["Move_Out_Date"]
        if not pd.notna(mo):
            return True
        mo_dt = mo.date() if hasattr(mo, "date") else mo
        return mo_dt > audit_date

    active = df_ev[df_ev.apply(_is_active, axis=1)].copy()

    for _, row in active.iterrows():
        prop   = row["Property"]
        unit   = row["Unit"]
        res    = row["Residents"]
        src    = row["Source_File"]
        delinq = float(row.get("Delinquency", 0) or 0)
        rent_d = float(row.get("Rent_Delinquency", 0) or 0)
        total  = max(delinq, rent_d)

        if total <= 0:
            continue

        flags.append(_flag(
            prop, unit, res,
            "EV-1", "Active Eviction — Outstanding Balance",
            "Delinquency",
            f"${total:,.2f}", "$0.00",
            f"Unit {unit} has an active eviction with no move-out date and an outstanding "
            f"balance of ${total:,.2f}. Coordinate with legal counsel — do NOT release "
            "hold on collections.",
            "CRITICAL", src,
        ))

    return flags


# ===========================================================================
# EV-2: Stalled Eviction — No Move-Out After 30 Days
# ===========================================================================

def run_ev2(df_ev: pd.DataFrame, audit_date: date = None) -> list:
    """
    Eviction filing older than 30 days, resident still in unit (UE status or no/future move-out).
    Process may be stalled — requires follow-up with legal counsel.
    """
    if audit_date is None:
        audit_date = datetime.today().date()
    flags = []
    if df_ev.empty:
        return flags

    def _is_active(row):
        status = str(row.get("Status", "")).strip().upper()
        if status == "UE":
            return True
        mo = row["Move_Out_Date"]
        if not pd.notna(mo):
            return True
        mo_dt = mo.date() if hasattr(mo, "date") else mo
        return mo_dt > audit_date

    active = df_ev[df_ev.apply(_is_active, axis=1)].copy()

    for _, row in active.iterrows():
        prop       = row["Property"]
        unit       = row["Unit"]
        res        = row["Residents"]
        src        = row["Source_File"]
        filed_date = row["Eviction_Filed_Date"]

        if not pd.notna(filed_date):
            continue

        filed_dt  = filed_date.date() if hasattr(filed_date, "date") else filed_date
        days_open = (audit_date - filed_dt).days

        if days_open < STALE_DAYS:
            continue

        flags.append(_flag(
            prop, unit, res,
            "EV-2", "Stalled Eviction Process",
            "Days Since Filing",
            f"{days_open} days (filed {filed_dt.strftime('%m/%d/%Y')})",
            f"Move-out within {STALE_DAYS} days of filing",
            f"Eviction for unit {unit} was filed {days_open} days ago "
            f"({filed_dt.strftime('%m/%d/%Y')}) and no move-out has been recorded. "
            "Follow up with legal counsel to confirm the process is active and on track.",
            "HIGH", src,
        ))

    return flags


# ===========================================================================
# EV-3: Evicted Resident with Balance Ready for Collections
# ===========================================================================

def run_ev3(df_ev: pd.DataFrame, audit_date: date = None) -> list:
    """
    Evicted and physically vacated (past move-out date, status E) with unpaid balance.
    Should be submitted to collections if no dispute is pending.
    """
    if audit_date is None:
        audit_date = datetime.today().date()
    flags = []
    if df_ev.empty:
        return flags

    def _is_vacated(row):
        status = str(row.get("Status", "")).strip().upper()
        if status == "UE":
            return False
        mo = row["Move_Out_Date"]
        if not pd.notna(mo):
            return False
        mo_dt = mo.date() if hasattr(mo, "date") else mo
        return mo_dt <= audit_date

    vacated = df_ev[df_ev.apply(_is_vacated, axis=1)].copy()

    for _, row in vacated.iterrows():
        prop     = row["Property"]
        unit     = row["Unit"]
        res      = row["Residents"]
        src      = row["Source_File"]
        delinq   = float(row.get("Delinquency", 0) or 0)
        rent_d   = float(row.get("Rent_Delinquency", 0) or 0)
        total    = max(delinq, rent_d)
        move_out = row["Move_Out_Date"]

        if total <= 0:
            continue

        mo_dt      = move_out.date() if hasattr(move_out, "date") else move_out
        days_since = (audit_date - mo_dt).days

        if days_since < COLLECTION_DAYS:
            continue

        flags.append(_flag(
            prop, unit, res,
            "EV-3", "Eviction Balance Ready for Collections",
            "Outstanding Balance Post Move-Out",
            f"${total:,.2f} ({days_since} days since move-out)",
            "Submit to collections",
            f"Unit {unit} vacated {days_since} days ago ({mo_dt.strftime('%m/%d/%Y')}) "
            f"with an unpaid balance of ${total:,.2f}. Verify no active dispute and submit "
            "to collections. Confirm hold-on-collections status in ResMan.",
            "HIGH", src,
        ))

    return flags


# ===========================================================================
# Main runner
# ===========================================================================

def run_eviction_audit(df_ev: pd.DataFrame, audit_date: date = None) -> pd.DataFrame:
    """Run all Phase 1 Eviction audit rules. Returns a DataFrame of exceptions."""
    if df_ev is None or df_ev.empty:
        print("  [WARN] No eviction data — Eviction audit skipped.")
        return pd.DataFrame()

    print(f"\n--- Eviction Audit ({len(df_ev)} eviction records) ---")

    flags = []
    flags.extend(run_ev1(df_ev))
    flags.extend(run_ev2(df_ev, audit_date))
    flags.extend(run_ev3(df_ev, audit_date))

    if not flags:
        print("  No eviction exceptions found.")
        return pd.DataFrame()

    df_out = pd.DataFrame(flags)
    print(
        f"  Eviction Audit complete: {len(df_out)} exceptions — "
        f"{(df_out['Severity'] == 'CRITICAL').sum()} CRITICAL  "
        f"{(df_out['Severity'] == 'HIGH').sum()} HIGH"
    )
    return df_out
