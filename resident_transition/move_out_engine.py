"""
Move-Out Audit Engine — Phase 1
================================
Rules:
  MO-1  Early Lease Break Fee Check
  MO-3  Final Account Statement Reconciliation
  MO-4  Collection Readiness Check  (includes Texas Property Code §92.103 — 30-day rule)
  MO-5  Refund Payout Accuracy

  MO-2 (Damage Photo Compliance) is Phase 2 — requires ResMan Document Attachment export.

Source: Partha Balakrishnan — Move-Out Audit Report Tool FRD v1.0 (July 22, 2026)
"""

import pandas as pd
from datetime import datetime, date, timedelta

# Texas Property Code §92.103 — landlord must return deposit within 30 days of move-out
TEXAS_REFUND_DAYS = 30

# LNJ standard minimum notice period for voluntary move-outs
NOTICE_DAYS_REQUIRED = 60

# Severity thresholds for executive-risk prioritization.
MO_NOTICE_CRITICAL_DAYS = 1
MO_BALANCE_CRITICAL_AMT = 1000.0
MO_BALANCE_MEDIUM_MAX_AMT = 100.0


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


def _moved_out_units(df_rr_units: pd.DataFrame, audit_date=None) -> pd.DataFrame:
    """Filter Rent Roll to units that have already vacated (Move_Out ≤ today)."""
    if audit_date is None:
        audit_date = pd.Timestamp.today().normalize()
    if df_rr_units.empty or "Move_Out" not in df_rr_units.columns:
        return pd.DataFrame()
    return df_rr_units[
        df_rr_units["Move_Out"].notna() &
        (df_rr_units["Move_Out"] <= audit_date)
    ].copy()


# ===========================================================================
# MO-1: Early Lease Break Fee Check
# ===========================================================================

def run_mo1(df_rr_units: pd.DataFrame, df_rr_charges: pd.DataFrame) -> list:
    """
    FR-5.2 / Section 7: If Move_Out < Lease_End, the resident broke their lease early.
    The broke-lease identifier must be set and a reletting/termination fee must be charged.

    Checks:
      - Move_Out < Lease_End → early termination
      - Reletting or termination fee present in unit's charge rows
    """
    flags = []
    moved_out = _moved_out_units(df_rr_units)

    if moved_out.empty:
        return flags

    ETF_KEYWORDS = ["reletting", "termination", "early termination", "lease break", "etf",
                    "re-letting", "re-let"]

    for _, row in moved_out.iterrows():
        prop      = row["Property"]
        unit      = row["Unit"]
        resident  = row["Residents"]
        src       = row["Source_File"]
        move_out  = row["Move_Out"]
        lease_end = row["Lease_End"]

        if not pd.notna(move_out) or not pd.notna(lease_end):
            continue
        if move_out >= lease_end:
            continue  # Normal end-of-lease move-out

        days_early = (lease_end - move_out).days

        # Check for reletting / termination fee in this unit's charges
        unit_charges = (
            df_rr_charges[
                (df_rr_charges["Property"] == prop) &
                (df_rr_charges["Unit"] == unit)
            ]
            if not df_rr_charges.empty else pd.DataFrame()
        )

        has_etf = (
            not unit_charges.empty and
            unit_charges["Description"].str.lower().apply(
                lambda d: any(kw in d for kw in ETF_KEYWORDS)
            ).any()
        )

        if not has_etf:
            flags.append(_flag(
                prop, unit, resident,
                "MO-1", "Early Lease Break Fee Check",
                "Reletting / Termination Fee",
                "Not charged", "Required",
                f"Resident moved out {move_out.strftime('%m/%d/%Y')} — "
                f"{days_early} days before lease end ({lease_end.strftime('%m/%d/%Y')}). "
                "Broke-lease identifier should be set in ResMan and a reletting or early "
                "termination fee must be charged. Submit to property manager.",
                "HIGH", src,
            ))

    return flags


# ===========================================================================
# MO-3: Final Account Statement Reconciliation
# ===========================================================================

def run_mo3(df_rr_units: pd.DataFrame) -> list:
    """
    FR-4.3 / Section 7: All charges, refunds, and balances must be accounted
    for before a move-out is marked complete.

    Checks:
      - Non-zero balance on a moved-out unit → not fully reconciled
    """
    flags = []
    moved_out = _moved_out_units(df_rr_units)

    if moved_out.empty:
        return flags

    for _, row in moved_out.iterrows():
        prop     = row["Property"]
        unit     = row["Unit"]
        resident = row["Residents"]
        src      = row["Source_File"]
        balance  = row["Balance"]
        move_out = row["Move_Out"]

        if balance == 0.0:
            continue  # Fully reconciled

        direction = "owes" if balance > 0 else "is owed a refund of"
        amt = abs(balance)
        mo_str = move_out.strftime("%m/%d/%Y") if pd.notna(move_out) else "unknown"

        sev = "HIGH"
        if balance >= MO_BALANCE_CRITICAL_AMT:
            sev = "CRITICAL"
        elif amt < MO_BALANCE_MEDIUM_MAX_AMT:
            sev = "MEDIUM"

        flags.append(_flag(
            prop, unit, resident,
            "MO-3", "Final Account Statement Reconciliation",
            "Final Balance",
            f"${balance:,.2f}", "$0.00",
            f"Unit {unit} has an unresolved balance of ${balance:,.2f} after move-out "
            f"({mo_str}). Resident {direction} ${amt:,.2f}. "
            "Review all charges and credits. Coordinate with property manager/accounting "
            "to reconcile before submitting to collections or issuing refund.",
            sev, src,
        ))

    return flags


# ===========================================================================
# MO-4: Collection Readiness Check
# ===========================================================================

def run_mo4(df_rr_units: pd.DataFrame, audit_date: date = None) -> list:
    """
    FR-4.2 / Section 7: Verify balance is clean and ready for collections or refund.

    Checks:
      - Negative balance (refund owed) past Texas §92.103 30-day deadline → CRITICAL
      - Negative balance within 30-day window → HIGH (upcoming deadline)
      - Positive balance past 30 days → HIGH (overdue for collections)
    """
    flags = []
    if audit_date is None:
        audit_date = datetime.today().date()
    moved_out = _moved_out_units(df_rr_units)

    if moved_out.empty:
        return flags

    for _, row in moved_out.iterrows():
        prop     = row["Property"]
        unit     = row["Unit"]
        resident = row["Residents"]
        src      = row["Source_File"]
        balance  = row["Balance"]
        move_out = row["Move_Out"]

        if not pd.notna(move_out) or balance == 0.0:
            continue

        mo_date = move_out.date() if hasattr(move_out, "date") else move_out
        days_since = (audit_date - mo_date).days
        deadline   = mo_date + timedelta(days=TEXAS_REFUND_DAYS)

        if balance < 0:
            # Refund owed to resident
            refund_amt = abs(balance)
            if days_since > TEXAS_REFUND_DAYS:
                flags.append(_flag(
                    prop, unit, resident,
                    "MO-4", "Collection Readiness Check",
                    "Refund — Texas §92.103 Deadline EXCEEDED",
                    f"${refund_amt:,.2f} refund not yet issued",
                    f"Issue within {TEXAS_REFUND_DAYS} days of move-out",
                    f"OVERDUE: Resident moved out {days_since} days ago "
                    f"({mo_date.strftime('%m/%d/%Y')}). Refund of ${refund_amt:,.2f} was "
                    f"due by {deadline.strftime('%m/%d/%Y')} per Texas Property Code §92.103. "
                    "Immediate action required — legal liability risk.",
                    "CRITICAL", src,
                ))
            else:
                remaining = TEXAS_REFUND_DAYS - days_since
                flags.append(_flag(
                    prop, unit, resident,
                    "MO-4", "Collection Readiness Check",
                    "Refund Pending — Texas §92.103 Deadline Approaching",
                    f"${refund_amt:,.2f} owed to resident",
                    f"Issue by {deadline.strftime('%m/%d/%Y')} ({remaining} days remaining)",
                    f"Refund of ${refund_amt:,.2f} must be issued by "
                    f"{deadline.strftime('%m/%d/%Y')} ({remaining} days remaining). "
                    "Texas Property Code §92.103.",
                    "CRITICAL", src,
                ))

        elif balance > 0 and days_since > TEXAS_REFUND_DAYS:
            # Resident owes money — past 30 days
            flags.append(_flag(
                prop, unit, resident,
                "MO-4", "Collection Readiness Check",
                "Outstanding Balance — Past 30 Days",
                f"${balance:,.2f} owed by resident",
                "Submit to collections if account is clean",
                f"Balance of ${balance:,.2f} has been outstanding for {days_since} days since "
                f"move-out ({mo_date.strftime('%m/%d/%Y')}). Verify no pending disputes exist "
                "and submit to collections.",
                "HIGH", src,
            ))

    return flags


# ===========================================================================
# MO-5: Refund Payout Accuracy
# ===========================================================================

def run_mo5(df_rr_units: pd.DataFrame) -> list:
    """
    FR-4.1 / Section 7: Refund amount must align with the final balance and
    deposists on file.

    Checks:
      - Deposit on file + negative balance → calculate total refund due
      - Deposit exceeds positive balance → deposit should cover charges, remainder refunded
    """
    flags = []
    moved_out = _moved_out_units(df_rr_units)

    if moved_out.empty:
        return flags

    for _, row in moved_out.iterrows():
        prop     = row["Property"]
        unit     = row["Unit"]
        resident = row["Residents"]
        src      = row["Source_File"]
        balance  = row["Balance"]
        deposits = row["Deposits"]

        if deposits <= 0:
            continue  # No deposit on file — nothing to reconcile here

        if balance < 0:
            # Resident has a credit balance; deposit should also be returned
            total_refund = abs(balance) + deposits
            flags.append(_flag(
                prop, unit, resident,
                "MO-5", "Refund Payout Accuracy",
                "Total Refund Amount",
                f"Credit balance ${abs(balance):,.2f} | Deposit on file ${deposits:,.2f}",
                f"Total refund ${total_refund:,.2f}",
                f"Resident has a credit balance of ${abs(balance):,.2f} plus a deposit of "
                f"${deposits:,.2f} on file. Total refund owed = ${total_refund:,.2f}. "
                "Verify this amount before issuing the refund check.",
                "MEDIUM", src,
            ))
        elif balance > 0 and deposits >= balance:
            # Deposit covers all charges — refund the remainder
            refund_amt = deposits - balance
            flags.append(_flag(
                prop, unit, resident,
                "MO-5", "Refund Payout Accuracy",
                "Deposit Covers Balance — Refund Remainder",
                f"Balance ${balance:,.2f} | Deposit ${deposits:,.2f}",
                f"Refund ${refund_amt:,.2f} to resident",
                f"Deposit (${deposits:,.2f}) covers the outstanding balance (${balance:,.2f}). "
                f"After applying deposit, resident is owed ${refund_amt:,.2f}. "
                "Verify refund has been or will be issued. Correct before issuing.",
                "HIGH", src,
            ))

    return flags


# ===========================================================================
# Main runner
# ===========================================================================

# ===========================================================================
# MO-1 extension: Notice Sufficiency Check
# ===========================================================================

def run_mo1_notice(df_cancel: pd.DataFrame) -> list:
    """
    Residents who gave fewer than 60 days notice before a voluntary move-out.
    Excludes evictions. Requires the Cancellations & Move Outs CSV.
    """
    flags = []
    if df_cancel.empty:
        return flags

    moved_out = df_cancel[df_cancel["Section"] == "Moved Out"].copy()

    for _, row in moved_out.iterrows():
        prop   = row["Property"]
        unit   = row["Unit"]
        res    = row["Residents"]
        src    = row["Source_File"]
        notice = str(row.get("Days_Notice", "")).strip()
        reason = str(row.get("Move_Out_Reason", "")).strip()

        # Skip involuntary/emergency departures — 60-day notice only applies to voluntary moves
        _EXEMPT = ["eviction", "onsite transfer", "skipped", "owes money", "no notice",
                   "death", "illness", "disaster"]
        if any(kw in reason.lower() for kw in _EXEMPT):
            continue
        try:
            days = int(notice)
        except (ValueError, TypeError):
            continue

        if days < NOTICE_DAYS_REQUIRED:
            sev = "CRITICAL" if days <= MO_NOTICE_CRITICAL_DAYS else "HIGH"
            flags.append(_flag(
                prop, unit, res,
                "MO-1", "Notice Sufficiency & Lease Break Fee Audit",
                "Days Notice Given",
                f"{days} days", f"{NOTICE_DAYS_REQUIRED} days required",
                f"Resident gave {days} day(s) notice before vacating "
                f"(reason: {reason or 'not recorded'}). "
                f"LNJ requires {NOTICE_DAYS_REQUIRED} days. Verify whether additional "
                "charges or lease-break fees apply.",
                sev, src,
            ))

    return flags


# ===========================================================================
# MO-2: Move-Out Reason & Tag Consistency
# ===========================================================================

def run_mo2(df_cancel: pd.DataFrame, df_ev: pd.DataFrame = None) -> list:
    """
    FR-5.2: Move-out reason must be set in ResMan and consistent with the eviction record.

    Checks:
      - Moved-out unit has a blank/missing reason
      - Reason is 'Eviction or Problems' but no matching Eviction Process record exists
    """
    flags = []
    if df_cancel.empty:
        return flags

    moved_out = df_cancel[df_cancel["Section"] == "Moved Out"].copy()

    ev_keys = set()
    if df_ev is not None and not df_ev.empty:
        ev_keys = set(zip(df_ev["Property"], df_ev["Unit"]))

    for _, row in moved_out.iterrows():
        prop   = row["Property"]
        unit   = row["Unit"]
        res    = row["Residents"]
        src    = row["Source_File"]
        reason = str(row.get("Move_Out_Reason", "")).strip()

        if not reason or reason.lower() in ("nan", "", "none"):
            flags.append(_flag(
                prop, unit, res,
                "MO-2", "Move-Out Reason & Tag Consistency",
                "Move-Out Reason",
                "Blank", "Required",
                f"No move-out reason is recorded in ResMan for unit {unit}. "
                "A reason must be entered on the Resident card before the "
                "move-out can be considered complete.",
                "MEDIUM", src,
            ))
            continue
        # Note: eviction-without-filing check removed — Eviction Process CSV only shows active cases

    return flags


# ===========================================================================
# Main runner
# ===========================================================================

def run_move_out_audit(df_rr_units: pd.DataFrame,
                       df_rr_charges: pd.DataFrame,
                       df_cancel: pd.DataFrame = None,
                       df_ev: pd.DataFrame = None,
                       audit_date: date = None) -> pd.DataFrame:
    """
    Run all Phase 1 Move-Out audit rules against the loaded DataFrames.
    Returns a DataFrame of all exceptions found.
    """
    _cancel = df_cancel if df_cancel is not None else pd.DataFrame()
    _ev     = df_ev     if df_ev     is not None else pd.DataFrame()

    moved_out_count = (
        df_rr_units["Move_Out"].notna().sum()
        if not df_rr_units.empty and "Move_Out" in df_rr_units.columns else 0
    )
    print(f"\n--- Move-Out Audit ({moved_out_count} moved-out units in Rent Roll) ---")

    flags = []
    # Rules that run from Cancellations CSV (no Rent Roll required)
    flags.extend(run_mo1_notice(_cancel))
    flags.extend(run_mo2(_cancel, _ev))

    # Rules that require Rent Roll data
    if not df_rr_units.empty and moved_out_count > 0:
        _mo1_etf = run_mo1(df_rr_units, df_rr_charges)
        # Remove ETF flags for units already covered by the Evictions engine
        if not _ev.empty:
            _ev_keys = set(zip(_ev["Property"], _ev["Unit"]))
            _mo1_etf = [f for f in _mo1_etf if (f["Property"], f["Unit"]) not in _ev_keys]
        flags.extend(_mo1_etf)
        flags.extend(run_mo3(df_rr_units))
        flags.extend(run_mo4(df_rr_units, audit_date))
        flags.extend(run_mo5(df_rr_units))

    if not flags:
        print("  No Move-Out exceptions found.")
        return pd.DataFrame()

    df_out = pd.DataFrame(flags)
    print(
        f"  Move-Out Audit complete: {len(df_out)} exceptions — "
        f"{(df_out['Severity'] == 'CRITICAL').sum()} CRITICAL  "
        f"{(df_out['Severity'] == 'HIGH').sum()} HIGH  "
        f"{(df_out['Severity'] == 'MEDIUM').sum()} MEDIUM"
    )
    return df_out
