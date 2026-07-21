"""
LiveNjoy ResMan Audit — Streamlit Dashboard  (app.py)
======================================================
Imports the engine from audit_bot.py and builds an interactive web dashboard.

Tabs
----
  1  Executive Summary       — KPIs + exposure totals
  2  Concession Audit        — John's 9-rule flags
  3  Revenue Integrity       — Daniel's 2-stage flags
  4  Manager Overrides       — Ranking + raw edit log
  5  Exposure Drilldowns     — By property / rule / risk
  6  Risk Matrix             — Heatmap of severity by property
"""

import streamlit as st
import pandas as pd
import numpy as np

from audit_bot import (
    run_full_audit,
    APPROVED_CODES,
    RISK_CRITICAL, RISK_HIGH, RISK_MEDIUM,
    PROPERTY_FEE_SCHEDULE,
)

# ─── PAGE CONFIG ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="LiveNjoy Audit Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── STYLING ──────────────────────────────────────────────────────────────────
RISK_COLORS = {
    RISK_CRITICAL: "#FF4B4B",
    RISK_HIGH:     "#FFA500",
    RISK_MEDIUM:   "#FFD700",
}

def color_risk(val):
    color = RISK_COLORS.get(val, "#FFFFFF")
    return f"background-color: {color}; color: black; font-weight: bold;"

def styled_df(df: pd.DataFrame, risk_col: str = "Risk_Level") -> object:
    if df.empty:
        return df
    if risk_col in df.columns:
        return df.style.map(color_risk, subset=[risk_col])
    return df


# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/property.png", width=60)
    st.title("LiveNjoy Audit Bot")
    st.markdown("**Version:** 2.0  \n**Engine:** John + Daniel Rules")
    st.divider()
    st.markdown("**Approved Concession Codes**")
    for code in sorted(APPROVED_CODES):
        st.markdown(f"- `{code}`")
    st.divider()
    run_btn = st.button("🚀  Run Full Forensic Audit", type="primary", use_container_width=True)
    st.divider()
    st.markdown("**File Setup Guide**")
    st.markdown(
        "Copy your exports into these folders:"
        "\n```"
        "\nexports/Transaction List Reports/"
        "\n  → data/transactions/"
        "\nexports/New & Renewed Leases/"
        "\n  → data/leases/"
        "\nexports/Edited Transactions by User/"
        "\n  → data/edits/"
        "\nexports/Transaction Projections/"
        "\n  → data/recurring/"
        "\nexports/Rent Rolls/"
        "\n  → data/rent_rolls/"
        "\nexports/Resident Activity/"
        "\n  → data/activity/"
        "\n```"
        "\n⚠️ Resident Ledgers are PDFs — not processed by the bot."
    )


# ─── HEADER ───────────────────────────────────────────────────────────────────
st.title("📊 LiveNjoy — Master Audit Dashboard")
st.markdown(
    "Automated **Concession Audit** (John's Rules: Post-Term, Missing Addendum, Amount Mismatch, Not Properly Posted, Large Credit, Non-Standard Description) + "
    "**Recurring Revenue Integrity Audit** (Daniel's 2-Stage Engine)"
)
st.divider()


# ─── RUN ENGINE ───────────────────────────────────────────────────────────────
if run_btn:
    with st.spinner("Ingesting CSVs and running audit engines…"):
        try:
            results = run_full_audit()
            st.session_state["results"] = results
            st.success("✅  Audit complete — dashboard updated.")
        except Exception as exc:
            st.error(f"❌  Engine error: {exc}")
            st.stop()

if "results" not in st.session_state:
    st.info("👈  Click **Run Full Forensic Audit** in the sidebar to begin.")
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
        rc1, rc2, rc3 = st.columns(3)
        rc1.metric("🔴 CRITICAL", int(row.get("Critical_Flags", 0)))
        rc2.metric("🟠 HIGH",     int(row.get("High_Flags", 0)))
        rc3.metric("🟡 MEDIUM",   int(row.get("Medium_Flags", 0)))

    st.divider()
    st.subheader("All Exceptions")
    if not all_flags.empty:
        # Filters
        cols = st.columns(3)
        with cols[0]:
            risk_filter = st.multiselect("Filter by Risk",
                [RISK_CRITICAL, RISK_HIGH, RISK_MEDIUM],
                default=[RISK_CRITICAL, RISK_HIGH, RISK_MEDIUM])
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

        st.dataframe(styled_df(view), use_container_width=True, hide_index=True)
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
        st.dataframe(styled_df(rule_summary), use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("Unit-Level Flags")
        st.dataframe(styled_df(concession_flags), use_container_width=True, hide_index=True)
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
                st.dataframe(missing_charges, use_container_width=True, hide_index=True)

            # Amount variance
            variances = stage1_flags[stage1_flags["Rule"].str.contains("Variance")]
            if not variances.empty:
                st.markdown("#### Charge Amount Inconsistencies")
                st.dataframe(styled_df(variances), use_container_width=True, hide_index=True)

            # Concession red flags
            conc_flags = stage1_flags[stage1_flags["Rule"].str.contains("Concession|concession")]
            if not conc_flags.empty:
                st.markdown("#### Concession Red Flags")
                st.dataframe(styled_df(conc_flags), use_container_width=True, hide_index=True)

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
                st.dataframe(styled_df(net_flags), use_container_width=True, hide_index=True)

            # Manual concessions
            manual_flags = stage2_flags[stage2_flags["Rule"].isin(
                {"Manual Posting Without Setup", "Invalid Credit Code"}
            )]
            if not manual_flags.empty:
                st.markdown("#### Manual Concession / Invalid Code")
                st.dataframe(styled_df(manual_flags), use_container_width=True, hide_index=True)

            # Mismatch + misc
            other = stage2_flags[~stage2_flags["Rule"].isin(
                net_flags["Rule"].tolist() + manual_flags["Rule"].tolist()
            )]
            if not other.empty:
                st.markdown("#### Other Posted Flags")
                st.dataframe(styled_df(other), use_container_width=True, hide_index=True)


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
            st.dataframe(manager_ranking, use_container_width=True, hide_index=True)
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
            st.dataframe(view, use_container_width=True, hide_index=True)
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
            st.dataframe(by_prop, use_container_width=True, hide_index=True)
        else:
            st.info("No property data.")

    with d2:
        st.subheader("By Rule / Charge Type")
        if not by_rule.empty:
            st.dataframe(by_rule, use_container_width=True, hide_index=True)
        else:
            st.info("No rule data.")

    with d3:
        st.subheader("By Risk Level")
        if not by_risk.empty:
            st.dataframe(styled_df(by_risk, "Risk_Level"), use_container_width=True, hide_index=True)
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
        st.dataframe(mgr_exposure, use_container_width=True, hide_index=True)
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
        st.dataframe(pivot, use_container_width=True)

        st.divider()
        st.subheader("Resident-Level Drilldown")
        props = sorted(all_flags["Property"].dropna().unique().tolist())
        selected_prop = st.selectbox("Select Property", ["All"] + props)
        drilldown = all_flags if selected_prop == "All" else all_flags[all_flags["Property"] == selected_prop]
        drilldown_sorted = drilldown.sort_values(
            ["Risk_Level", "Amount_Impact"],
            key=lambda col: col.map({RISK_CRITICAL: 0, RISK_HIGH: 1, RISK_MEDIUM: 2}) if col.name == "Risk_Level" else col,
            ascending=[True, False]
        )
        st.dataframe(styled_df(drilldown_sorted), use_container_width=True, hide_index=True)
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
        st.dataframe(summary, use_container_width=True, hide_index=True)

        st.divider()

        # Filter by property
        props_fs = ["All"] + sorted(fee_flags["Property"].dropna().unique().tolist())
        sel_prop = st.selectbox("Filter by Property", props_fs, key="fee_prop_filter")
        view_fee = fee_flags if sel_prop == "All" else fee_flags[fee_flags["Property"] == sel_prop]

        st.subheader(f"Unit-Level Detail ({len(view_fee)} flags)")
        st.dataframe(styled_df(view_fee), use_container_width=True, hide_index=True)
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
        st.dataframe(pd.DataFrame(fee_rows), use_container_width=True, hide_index=True)


# ─── FOOTER ───────────────────────────────────────────────────────────────────
st.divider()
st.caption("LiveNjoy Residential · ResMan Audit Bot · Built per John B. & Daniel Twito specifications")
