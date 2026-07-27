"""
Revenue & Concession Audit Tool
LiveNjoy Residential  |  Built per John B. & Daniel Twito specifications

Part of the LiveNjoy Automation Suite.
Click "← Hub" in the sidebar to return to the tool hub.
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "audit_tool"))

from datetime import datetime
from pathlib import Path

import pandas as pd
import numpy as np
import streamlit as st
import streamlit.components.v1 as components

from audit_bot import (
    run_full_audit,
    APPROVED_CODES,
    AUDIT_MONTH,
    RISK_CRITICAL, RISK_HIGH, RISK_MEDIUM, RISK_VERIFY,
    PROPERTY_FEE_SCHEDULE,
)

st.set_page_config(
    page_title="LNJ Revenue & Concession Audit",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Hide Streamlit auto-generated page-navigation links in sidebar
st.markdown(
    '<style>[data-testid="stSidebarNav"],[data-testid="stSidebarNavContainer"],[data-testid="stSidebarNavItems"],[data-testid="stSidebarNavLink"]{display:none !important;}</style>',
    unsafe_allow_html=True,
)

RISK_COLORS = {
    RISK_CRITICAL: "#FF4B4B",
    RISK_HIGH:     "#FFA500",
    RISK_MEDIUM:   "#FFD700",
    RISK_VERIFY:   "#A8D8A8",   # light green — verify only
}

def color_risk(val):
    color = RISK_COLORS.get(val, "#FFFFFF")
    return f"background-color: {color}; color: black; font-weight: bold;"

def styled_df(df, risk_col="Risk_Level"):
    if df.empty:
        return df
    if risk_col in df.columns:
        return df.style.map(color_risk, subset=[risk_col])
    return df

def _norm(name: str) -> str:
    return name.lower().replace(" ", "").replace("-", "").replace("_", "")

def _valid(files, keyword: str) -> bool:
    return bool(files) and all(keyword in _norm(f.name) for f in files)

def derive_file_type(filename: str) -> str | None:
    """Route a ResMan CSV filename to its data category."""
    n = _norm(filename)
    if "marketrent" in n:            return "market_rent_schedule"
    if "editedtransactions" in n:    return "edits"
    if "recurringtransaction" in n:  return "recurring"
    if "transactionlist" in n:       return "transactions"
    if "rentroll" in n:              return "rent_rolls"
    if "newandrenewed" in n:         return "leases"
    if "residentactivity" in n:      return "activity"
    return None

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
<style>
/* ── Compact sidebar spacing ── */
section[data-testid="stSidebar"] hr {
    margin-top: 0.35rem !important;
    margin-bottom: 0.35rem !important;
}
section[data-testid="stSidebar"] .stTextInput {
    margin-bottom: 0 !important;
}
/* Hide the "200MB per file • CSV" hint to save height */
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] > div > small {
    display: none !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
    padding: 0.4rem 0.75rem !important;
    min-height: unset !important;
}
/* Tighten gap between stacked file uploader widgets */
section[data-testid="stSidebar"] [data-testid="stFileUploader"] {
    margin-bottom: 0.1rem !important;
}
</style>
""", unsafe_allow_html=True)
    if st.button("\u2190  Go back to Hub", key="back_to_hub", use_container_width=True):
        st.switch_page("app.py")
    st.divider()

    audit_month_input = st.session_state.get("audit_month_auto", "Upload CSV Files to Auto-Detect")
    st.text_input(
        "Audit Month",
        value=audit_month_input,
        disabled=True,
        help="Auto-detected from your uploaded Transaction List CSV",
    )

    st.divider()
    with st.expander("📂 Upload ResMan CSVs", expanded=False):
        st.caption(
            "Drop all ResMan CSV exports here — files are auto-detected by filename. "
            "Upload any mix of properties and report types."
        )
        _raw_uploads = st.file_uploader(
            "ResMan CSV Exports",
            type="csv",
            accept_multiple_files=True,
            help="Transaction Lists · Leases · Rent Rolls · Recurring Projections · Edited Transactions · Resident Activity · Market Rent Schedule",
        )

        # ── Route each file to its bucket by filename ──────────────────────
        up_transactions, up_leases, up_edits   = [], [], []
        up_recurring, up_rent_rolls, up_activity, up_market_rent = [], [], [], []
        _unrecognized = []
        _seen = set()
        for _f in (_raw_uploads or []):
            if _f.name in _seen:
                continue
            _seen.add(_f.name)
            _ftype = derive_file_type(_f.name)
            if   _ftype == "transactions":        up_transactions.append(_f)
            elif _ftype == "leases":              up_leases.append(_f)
            elif _ftype == "edits":               up_edits.append(_f)
            elif _ftype == "recurring":           up_recurring.append(_f)
            elif _ftype == "rent_rolls":          up_rent_rolls.append(_f)
            elif _ftype == "activity":            up_activity.append(_f)
            elif _ftype == "market_rent_schedule": up_market_rent.append(_f)
            else:                                 _unrecognized.append(_f.name)

        # ── Detection summary ───────────────────────────────────────────────
        if _raw_uploads:
            _cats = [
                ("Transaction Lists",  up_transactions),
                ("Leases",             up_leases),
                ("Rent Rolls",         up_rent_rolls),
                ("Recurring",          up_recurring),
                ("Edits",              up_edits),
                ("Activity",           up_activity),
                ("Market Rent",        up_market_rent),
            ]
            st.markdown(
                "  \n".join(
                    f"{'✅' if b else '❌'} **{lbl}** ({len(b)})"
                    for lbl, b in _cats
                )
            )
            if _unrecognized:
                st.warning(f"⚠️ {len(_unrecognized)} unrecognized file(s) — check filenames: {', '.join(_unrecognized)}")

    # ── Auto-detect audit month from uploaded Transaction List ──────────────────────
    if up_transactions:
        try:
            import io as _io
            _content = up_transactions[0].read()
            up_transactions[0].seek(0)
            _raw = pd.read_csv(_io.BytesIO(_content), skiprows=6, dtype=str, header=0, nrows=300)
            _dates = pd.to_datetime(_raw.iloc[:, 0].astype(str), errors="coerce").dropna()
            if not _dates.empty:
                _period = _dates.dt.to_period("M").value_counts().index[0]
                st.session_state["audit_month_auto"] = _period.to_timestamp().strftime("%b %Y")
        except Exception:
            pass
    elif "audit_month_auto" not in st.session_state:
        st.session_state["audit_month_auto"] = AUDIT_MONTH
    audit_month_input = st.session_state["audit_month_auto"]

    st.divider()
    with st.expander("📄 Audit Report History", expanded=False):
        _out = Path("output")
        _reports = sorted(
            _out.glob("LNJ_Audit_*.xlsx"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        ) if _out.exists() else []

        if not _reports:
            st.caption("No saved reports yet. Run an audit to generate one.")
        else:
            def _fmt_report(p):
                try:
                    parts = p.stem.split("_")  # LNJ Audit [MonYYYY] YYYYMMDD HHMM
                    if len(parts) >= 5 and not parts[2].isdigit():
                        # New format: LNJ_Audit_Jun2026_20260722_1430
                        month_lbl = parts[2][:3] + " " + parts[2][3:]  # "Jun2026" -> "Jun 2026"
                        dt = datetime.strptime(parts[3] + parts[4], "%Y%m%d%H%M")
                        return f"{month_lbl}  ·  run {dt.strftime('%b %d, %Y  %I:%M %p')}"
                    else:
                        # Old format: LNJ_Audit_20260604_0122
                        dt = datetime.strptime(parts[2] + parts[3], "%Y%m%d%H%M")
                        return dt.strftime("%b %d, %Y  ·  %I:%M %p")
                except Exception:
                    return p.name

            _labels = [_fmt_report(f) for f in _reports]
            _sel_idx = st.selectbox(
                "Select a past report",
                range(len(_labels)),
                format_func=lambda i: _labels[i],
                label_visibility="collapsed",
            )
            with open(_reports[_sel_idx], "rb") as _fh:
                st.download_button(
                    label="📥  Download Selected Report",
                    data=_fh.read(),
                    file_name=_reports[_sel_idx].name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                )

    total_files = sum(len(x or []) for x in [
        up_transactions, up_leases, up_edits,
        up_recurring, up_rent_rolls, up_activity, up_market_rent,
    ])

    # All 7 categories must have at least one file
    all_uploaded = all([
        up_transactions, up_leases, up_edits,
        up_recurring, up_rent_rolls, up_activity, up_market_rent,
    ])

    # Every uploaded file must match its expected ResMan filename keyword
    all_valid = all([
        _valid(up_transactions, "transactionlist"),
        _valid(up_leases,       "newandrenewed"),
        _valid(up_edits,        "editedtransactions"),
        _valid(up_recurring,    "recurringtransaction"),
        _valid(up_rent_rolls,   "rentroll"),
        _valid(up_activity,     "residentactivity"),
        _valid(up_market_rent,  "marketrent"),
    ])

    run_ready = all_uploaded and all_valid

    cats_filled = sum(1 for x in [
        up_transactions, up_leases, up_edits,
        up_recurring, up_rent_rolls, up_activity, up_market_rent,
    ] if x)

    if run_ready:
        btn_label = f"🚀  Run Audit  ({total_files} file{'s' if total_files != 1 else ''} ready)"
    elif all_uploaded and not all_valid:
        btn_label = "🚀  Run Audit — check file names"
    elif cats_filled:
        btn_label = f"🚀  Run Audit  ({cats_filled} / 7 categories)"
    else:
        btn_label = "🚀  Run Audit"

    st.divider()
    run_btn = st.button(
        btn_label,
        type="primary",
        width='stretch',
        disabled=not run_ready,
    )

    if st.session_state.get("excel_bytes"):
        ts_label = st.session_state.get("run_ts", datetime.now().strftime("%Y%m%d_%H%M"))
        st.download_button(
            label="📥  Download Excel Report",
            data=st.session_state["excel_bytes"],
            file_name=f"LNJ_Audit_{ts_label}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch',
        )


# ─── HEADER ───────────────────────────────────────────────────────────────────
st.markdown("""
<div style="display:flex;align-items:center;gap:14px;margin-bottom:0.25rem;">
<div style="width:48px;height:48px;background:#1A2744;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
<span style="color:#FFFFFF;font-size:1.4rem;font-weight:700;">$</span>
</div>
<h1 style="margin:0;padding:0;font-size:2rem;font-weight:700;line-height:1.2;">Revenue &amp; Concession Audit Tool</h1>
</div>
""", unsafe_allow_html=True)
st.markdown(
    "Automated **Concession Audit** (Post-Term, Missing Addendum, Amount Mismatch, Not Properly Posted, Large Credit, Non-Standard Description) + "
    "**Recurring Revenue Integrity Audit**"
)
st.divider()


# ─── RUN ENGINE ───────────────────────────────────────────────────────────────
if run_btn:
    uploaded_files = {
        "transactions":         up_transactions  or [],
        "leases":               up_leases        or [],
        "edits":                up_edits         or [],
        "recurring":            up_recurring     or [],
        "rent_rolls":           up_rent_rolls    or [],
        "activity":             up_activity      or [],
        "market_rent_schedule": up_market_rent   or [],
    }
    with st.spinner("Ingesting CSVs and running audit engines…"):
        try:
            results = run_full_audit(
                uploaded_files=uploaded_files,
                audit_month=audit_month_input.strip() or None,
            )
            st.session_state["results"]     = results
            st.session_state["excel_bytes"] = results.get("excel_bytes")
            _month_slug = audit_month_input.replace(" ", "")
            st.session_state["run_ts"] = f"{_month_slug}_{datetime.now().strftime('%Y%m%d_%H%M')}"
            st.success("✅  Audit complete — use the **Download Excel Report** button in the sidebar.")
            st.rerun()
        except Exception as exc:
            st.error(f"❌  Engine error: {exc}")
            st.stop()

if "results" not in st.session_state:
    st.info("👈  Upload your ResMan CSV exports in the sidebar, then click **Run Audit** to begin.")
    st.stop()

# Unpack results
R               = st.session_state["results"]
concession_flags     = R["concession_flags"]
revenue_integrity_flags   = R["revenue_integrity_flags"]
fee_flags       = R.get("fee_flags", pd.DataFrame())
all_flags       = R["all_flags"]
manager_ranking = R["manager_ranking"]
override_log    = R["override_log"]
exposure        = R["exposure"]

totals          = exposure.get("totals", pd.DataFrame())
by_prop         = exposure.get("by_property", pd.DataFrame())
by_rule         = exposure.get("by_rule", pd.DataFrame())
by_risk         = exposure.get("by_risk", pd.DataFrame())


# ─── TABS ─────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📈 Executive Summary",
    "🔍 Concession Audit Engine",
    "⚙️  Revenue Integrity Engine",
    "👤 Manager Overrides",
    "💰 Exposure Drilldowns",
    "🗂️  Risk Matrix",
    "📋 Fee Schedule Check",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — EXECUTIVE SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.header("Portfolio Health Snapshot")

    if totals.empty:
        st.warning("No audit data available.")
    else:
        row = totals.iloc[0]

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Units Audited",        int(row.get("Total_Units_Audited", 0)))
        c2.metric("Total Exceptions",     int(row.get("Total_Exceptions", 0)))
        c3.metric("Financial Exposure",   f"${row.get('Deduped_Exposure', row.get('Total_Exposure', 0)):,.2f}")
        c4.metric("Avg Flags / Unit",     f"{row.get('Avg_Flags_Per_Unit', 0):.1f}")
        c5.metric("Critical Flags",       int(row.get("Critical_Flags", 0)))
        st.caption(
            f"Financial Exposure shows the conservative deduped figure "
            f"(max impact per unit across all engines: "
            f"**${row.get('Deduped_Exposure', row.get('Total_Exposure', 0)):,.2f}**). "
            f"Raw sum across all flags: **${row.get('Total_Exposure', 0):,.2f}**."
        )

        st.divider()
        st.subheader("Flags by Risk Level")
        rc1, rc2, rc3, rc4 = st.columns(4)
        rc1.metric("🔴 CRITICAL", int(row.get("Critical_Flags", 0)))
        rc2.metric("🟠 HIGH",     int(row.get("High_Flags", 0)))
        rc3.metric("🟡 MEDIUM",   int(row.get("Medium_Flags", 0)))
        rc4.metric("🟢 VERIFY",   int(row.get("Verify_Flags", 0)))

    st.divider()
    st.subheader("All Exceptions")
    if not all_flags.empty:
        # Filters
        cols = st.columns(3)
        with cols[0]:
            risk_filter = st.multiselect("Filter by Risk",
                [RISK_CRITICAL, RISK_HIGH, RISK_MEDIUM, RISK_VERIFY],
                default=[RISK_CRITICAL, RISK_HIGH, RISK_MEDIUM, RISK_VERIFY])
        with cols[1]:
            prop_options = ["All"] + sorted(all_flags["Property"].dropna().unique().tolist())
            prop_filter = st.selectbox("Filter by Property", prop_options)
        with cols[2]:
            rule_options = ["All"] + sorted(all_flags["Rule"].dropna().unique().tolist())
            rule_filter = st.selectbox("Filter by Rule", rule_options)

        view = all_flags[all_flags["Risk_Level"].isin(risk_filter)]
        if prop_filter != "All":
            view = view[view["Property"] == prop_filter]
        if rule_filter != "All":
            view = view[view["Rule"] == rule_filter]

        st.dataframe(styled_df(view), width='stretch', hide_index=True)
        st.caption(f"{len(view):,} records shown")
    else:
        st.success("No exceptions found — clean portfolio!")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — CONCESSION AUDIT ENGINE
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.header("Concession Audit Engine")
    st.markdown(
        "Validates every concession/credit posting against the "
        "approved codes (`CONR`, `CRTCO`, `EMPL`, `MCCR`, `RRFee`) "
        "and the legal lease document."
    )

    if concession_flags.empty:
        st.success("✅  No concession violations detected.")
    else:
        # Summary by rule
        rule_summary = (
            concession_flags.groupby(["Rule", "Risk_Level"])
            .agg(Count=("Rule", "count"), Exposure=("Amount_Impact", "sum"))
            .reset_index()
            .sort_values("Exposure", ascending=False)
        )
        st.subheader("Rule Summary")
        st.dataframe(styled_df(rule_summary), width='stretch', hide_index=True)

        st.divider()
        st.subheader("Unit-Level Flags")
        st.dataframe(styled_df(concession_flags), width='stretch', hide_index=True)
        st.caption(
            f"Concession Audit Engine exposure (raw sum): **${concession_flags['Amount_Impact'].sum():,.2f}**. "
            "Note: some units may also appear in the Revenue Integrity Engine — the Executive Summary "
            "deduplicates by taking the max impact per unit."
        )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — REVENUE INTEGRITY ENGINE AUDIT
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.header("Revenue Integrity Engine — 2-Stage Audit")

    s1_tab, s2_tab = st.tabs(["Stage 1 — Recurring Projection", "Stage 2 — Posted Rent Roll"])

    # Stage 1 flags (identified by rules from that stage)
    stage1_rules = {
        "Missing Standard Charge", "Major Charge Amount Variance",
        "Minor Charge Amount Variance", "Recurring Concession >$700",
        "Concession >$500 for 2+ Months", "Concession No Expiration",
        "Post-Term Credit",
    }
    stage2_rules = {
        "Negative Net Rent", "$0 Net Rent (Recent Move-in)", "$0 Net Rent (Not Recent)",
        "Manual Posting Without Setup", "Invalid Credit Code",
        "Posted vs Recurring Mismatch", "Misc Tenant Credit",
    }

    stage1_flags = revenue_integrity_flags[revenue_integrity_flags["Rule"].isin(stage1_rules)] if not revenue_integrity_flags.empty else pd.DataFrame()
    stage2_flags = revenue_integrity_flags[revenue_integrity_flags["Rule"].isin(stage2_rules)] if not revenue_integrity_flags.empty else pd.DataFrame()

    with s1_tab:
        st.subheader("What should post every month vs what is configured")
        if stage1_flags.empty:
            st.success("✅  No recurring projection issues found.")
        else:
            st.markdown(f"**{len(stage1_flags)} flags** — Total Exposure: **${stage1_flags['Amount_Impact'].sum():,.2f}**")

            # 90% rule violations
            missing_charges = stage1_flags[stage1_flags["Rule"] == "Missing Standard Charge"]
            if not missing_charges.empty:
                st.markdown("#### Missing Standard Charges (90% Rule)")
                st.dataframe(missing_charges, width='stretch', hide_index=True)

            # Amount variance
            variances = stage1_flags[stage1_flags["Rule"].str.contains("Variance")]
            if not variances.empty:
                st.markdown("#### Charge Amount Inconsistencies")
                st.dataframe(styled_df(variances), width='stretch', hide_index=True)

            # Concession red flags
            conc_flags = stage1_flags[stage1_flags["Rule"].str.contains("Concession|concession")]
            if not conc_flags.empty:
                st.markdown("#### Concession Red Flags")
                st.dataframe(styled_df(conc_flags), width='stretch', hide_index=True)

    with s2_tab:
        st.subheader("What managers actually posted this month")
        if stage2_flags.empty:
            st.success("✅  No posted rent roll issues found.")
        else:
            st.markdown(f"**{len(stage2_flags)} flags** — Total Exposure: **${stage2_flags['Amount_Impact'].sum():,.2f}**")

            # Net rent integrity
            net_flags = stage2_flags[stage2_flags["Rule"].str.contains("Net Rent")]
            if not net_flags.empty:
                st.markdown("#### Net Rent Integrity Issues")
                st.dataframe(styled_df(net_flags), width='stretch', hide_index=True)

            # Manual concessions
            manual_flags = stage2_flags[stage2_flags["Rule"].isin(
                {"Manual Posting Without Setup", "Invalid Credit Code"}
            )]
            if not manual_flags.empty:
                st.markdown("#### Manual Concession / Invalid Code")
                st.dataframe(styled_df(manual_flags), width='stretch', hide_index=True)

            # Mismatch + misc
            other = stage2_flags[~stage2_flags["Rule"].isin(
                net_flags["Rule"].tolist() + manual_flags["Rule"].tolist()
            )]
            if not other.empty:
                st.markdown("#### Other Posted Flags")
                st.dataframe(styled_df(other), width='stretch', hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — MANAGER OVERRIDES
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.header("Manager Override Analysis")
    st.markdown(
        "Tracks every manual ledger edit from the **Edited Transactions** report. "
        "Ranks managers by total revenue impact of their overrides."
    )

    left, right = st.columns([1, 2])

    with left:
        st.subheader("📋 Manager Leaderboard")
        if manager_ranking.empty:
            st.success("No manual overrides detected.")
        else:
            st.dataframe(manager_ranking, width='stretch', hide_index=True)
            worst = manager_ranking.iloc[0]
            st.warning(
                f"⚠️ Highest impact: **{worst['Manager_Login']}** "
                f"at **{worst['Property']}** — "
                f"${worst['Total_Impact']:,.2f} across "
                f"{int(worst['Total_Events'])} events"
            )

    with right:
        st.subheader("📝 Raw Override Log")
        if override_log.empty:
            st.info("No override detail available.")
        else:
            # Filter by manager
            managers = ["All"] + sorted(override_log["Manager_Login"].unique().tolist())
            selected_mgr = st.selectbox("Filter by Manager", managers)
            view = override_log if selected_mgr == "All" else override_log[override_log["Manager_Login"] == selected_mgr]
            st.dataframe(view, width='stretch', hide_index=True)
            if not view.empty:
                st.caption(f"Revenue impact shown: **${view['Revenue_Impact'].sum():,.2f}**")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — EXPOSURE DRILLDOWNS
# ══════════════════════════════════════════════════════════════════════════════
with tab5:
    st.header("Financial Exposure Drilldowns")

    d1, d2, d3 = st.columns(3)

    with d1:
        st.subheader("By Property")
        if not by_prop.empty:
            st.dataframe(by_prop, width='stretch', hide_index=True)
        else:
            st.info("No property data.")

    with d2:
        st.subheader("By Rule / Charge Type")
        if not by_rule.empty:
            st.dataframe(by_rule, width='stretch', hide_index=True)
        else:
            st.info("No rule data.")

    with d3:
        st.subheader("By Risk Level")
        if not by_risk.empty:
            st.dataframe(styled_df(by_risk, "Risk_Level"), width='stretch', hide_index=True)
        else:
            st.info("No risk data.")

    st.divider()
    st.subheader("Exposure by Manager (from Override Log)")
    if not override_log.empty:
        mgr_exposure = (
            override_log.groupby(["Property", "Manager_Login"])["Revenue_Impact"]
            .agg(Edits="count", Total_Impact="sum")
            .reset_index()
            .sort_values("Total_Impact", ascending=True)
        )
        st.dataframe(mgr_exposure, width='stretch', hide_index=True)
    else:
        st.info("No manager exposure data (no edited transactions loaded).")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — RISK MATRIX
# ══════════════════════════════════════════════════════════════════════════════
with tab6:
    st.header("Risk Matrix — Severity by Property")

    if all_flags.empty:
        st.success("No flags — nothing to display.")
    else:
        pivot = (
            all_flags.groupby(["Property", "Risk_Level"])
            .agg(Count=("Rule", "count"), Exposure=("Amount_Impact", "sum"))
            .reset_index()
            .pivot_table(index="Property",
                         columns="Risk_Level",
                         values=["Count", "Exposure"],
                         fill_value=0)
        )
        pivot.columns = [f"{v}_{c}" for v, c in pivot.columns]
        pivot = pivot.reset_index()
        st.dataframe(pivot, width='stretch')

        st.divider()
        st.subheader("Resident-Level Drilldown")
        props = sorted(all_flags["Property"].dropna().unique().tolist())
        selected_prop = st.selectbox("Select Property", ["All"] + props)
        drilldown = all_flags if selected_prop == "All" else all_flags[all_flags["Property"] == selected_prop]
        drilldown_sorted = drilldown.sort_values(
            ["Risk_Level", "Amount_Impact"],
            key=lambda col: col.map({RISK_CRITICAL: 0, RISK_HIGH: 1, RISK_MEDIUM: 2, RISK_VERIFY: 3}) if col.name == "Risk_Level" else col,
            ascending=[True, False]
        )
        st.dataframe(styled_df(drilldown_sorted), width='stretch', hide_index=True)
        st.caption(
            f"{len(drilldown_sorted):,} exceptions | "
            f"Exposure: **${drilldown_sorted['Amount_Impact'].sum():,.2f}**"
        )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 — FEE SCHEDULE CHECK
# ══════════════════════════════════════════════════════════════════════════════
with tab7:
    st.header("Fee Schedule Violations")
    st.markdown(
        "Compares each unit's **Recurring Transaction Projection** against the "
        "official fee sheet amounts provided by Daniel Twito. "
        "Flags any charge that differs from the fee schedule by **≥ $1**. "
        "\n\n> **La Prada** is excluded — no fee sheet provided. "
        "**Parking, pet fees, and washer/dryer** are marked optional and only "
        "flagged when the charge exists with the wrong amount."
    )

    if fee_flags.empty:
        st.success("✅  All recurring charges match the official fee schedule.")
    else:
        # Summary by property
        summary = (
            fee_flags.groupby(["Property"])
            .agg(Units=("Unit", "nunique"), Flags=("Rule", "count"),
                 Total_Variance=("Amount_Impact", "sum"))
            .reset_index()
            .sort_values("Total_Variance", ascending=False)
        )
        st.subheader("Summary by Property")
        st.dataframe(summary, width='stretch', hide_index=True)

        st.divider()

        # Filter by property
        props_fs = ["All"] + sorted(fee_flags["Property"].dropna().unique().tolist())
        sel_prop = st.selectbox("Filter by Property", props_fs, key="fee_prop_filter")
        view_fee = fee_flags if sel_prop == "All" else fee_flags[fee_flags["Property"] == sel_prop]

        st.subheader(f"Unit-Level Detail ({len(view_fee)} flags)")
        st.dataframe(styled_df(view_fee), width='stretch', hide_index=True)
        st.caption(f"Total variance exposure: **${view_fee['Amount_Impact'].sum():,.2f}**")

        st.divider()
        st.subheader("Official Fee Schedule Reference")
        fee_rows = []
        for prop_name, fees in PROPERTY_FEE_SCHEDULE.items():
            for f in fees:
                fee_rows.append({
                    "Property":   prop_name,
                    "Fee Name":   f["name"],
                    "Expected $": f"${f['amount']:.2f}",
                    "Optional":   "Yes" if f["optional"] else "No",
                })
        st.dataframe(pd.DataFrame(fee_rows), width='stretch', hide_index=True)


# ─── FOOTER ───────────────────────────────────────────────────────────────────
st.divider()
st.caption("LiveNjoy Residential · ResMan Audit Bot · Built per John B. & Daniel Twito specifications")
