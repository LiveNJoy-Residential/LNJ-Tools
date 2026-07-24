"""
Resident Activity Audit Tool — Streamlit Dashboard
LiveNjoy Residential | Move-In & Move-Out Audit | Phase 1

Run: .venv\Scripts\streamlit.exe run resident_transition/transition_app.py
"""

import sys
import os
import pandas as pd
import streamlit as st

# Allow direct imports from this folder
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'resident_transition'))

from data_loader import load_leases, load_activity, load_rent_roll
from move_in_engine import run_move_in_audit
from move_out_engine import run_move_out_audit

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="LNJ Resident Transitions",
    page_icon="🏠",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.markdown("""
<div style="display:flex;align-items:center;gap:14px;margin-bottom:0.25rem;">
<div style="width:48px;height:48px;background:#1A2744;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
<span style="color:#FFFFFF;font-size:1.4rem;font-weight:700;">&#x21C4;</span>
</div>
<h1 style="margin:0;padding:0;font-size:2rem;font-weight:700;line-height:1.2;">Resident Activity Audit Tool</h1>
</div>
""", unsafe_allow_html=True)
st.markdown("Automated **Move-In Audit** + **Move-Out Audit**")
st.caption("Phase 1 · Currently in development")


# Hide Streamlit auto-generated page-navigation links in sidebar
st.markdown(
    '<style>[data-testid="stSidebarNav"],[data-testid="stSidebarNavContainer"],[data-testid="stSidebarNavItems"],[data-testid="stSidebarNavLink"]{display:none !important;}</style>',
    unsafe_allow_html=True,
)

st.divider()

# ---------------------------------------------------------------------------
# Data loader (cached)
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Loading ResMan CSV data...")
def load_all():
    print("\n=== Loading Transition Tool Data ===")
    df_leases                = load_leases()
    df_activity              = load_activity()
    df_rr_units, df_rr_charges = load_rent_roll()
    return df_leases, df_activity, df_rr_units, df_rr_charges


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    if st.button("\u2190  Go back to Hub", use_container_width=True): st.switch_page("app.py")
    st.divider()
    st.header("Controls")
    if st.button("🔄  Reload Data", use_container_width=True):
        st.cache_data.clear()
        for key in ["df_mi", "df_mo"]:
            st.session_state.pop(key, None)
        st.rerun()

    st.divider()
    st.markdown("**Phase 1 Rules**")
    st.markdown("Move-In: MI-1 · MI-2 · MI-4 · MI-5")
    st.markdown("Move-Out: MO-1 · MO-3 · MO-4 · MO-5")
    st.caption("MI-3 & MO-2 (Document Upload) are Phase 2.")
    st.divider()
    st.markdown("**Data sources**")
    st.markdown("• New & Renewed Leases")
    st.markdown("• Resident Activity")
    st.markdown("• Rent Roll")


# ---------------------------------------------------------------------------
# Severity color helper
# ---------------------------------------------------------------------------
SEVERITY_COLOR = {"CRITICAL": "🔴", "HIGH": "🟠", "MEDIUM": "🟡"}

def _severity_icon(sev: str) -> str:
    return SEVERITY_COLOR.get(sev, "⚪")


# ---------------------------------------------------------------------------
# Shared display helper
# ---------------------------------------------------------------------------
def render_results(df: pd.DataFrame, key_prefix: str):
    if df.empty:
        st.success("✅  No exceptions found.")
        return

    # Summary metrics
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Exceptions", len(df))
    c2.metric("🔴 CRITICAL", int((df["Severity"] == "CRITICAL").sum()))
    c3.metric("🟠 HIGH",     int((df["Severity"] == "HIGH").sum()))
    c4.metric("🟡 MEDIUM",   int((df["Severity"] == "MEDIUM").sum()))

    st.divider()

    # Filters
    props = ["All"] + sorted(df["Property"].dropna().unique().tolist())
    sevs  = ["All", "CRITICAL", "HIGH", "MEDIUM"]
    stats = ["All", "Open", "Reviewed", "Cleared", "Escalated"]

    f1, f2, f3 = st.columns(3)
    f_prop = f1.selectbox("Property",  props, key=f"{key_prefix}_prop")
    f_sev  = f2.selectbox("Severity",  sevs,  key=f"{key_prefix}_sev")
    f_stat = f3.selectbox("Status",    stats, key=f"{key_prefix}_stat")

    mask = pd.Series([True] * len(df), index=df.index)
    if f_prop != "All":
        mask &= df["Property"] == f_prop
    if f_sev != "All":
        mask &= df["Severity"] == f_sev
    if f_stat != "All":
        mask &= df["Status"] == f_stat

    filtered = df[mask].copy()
    st.caption(f"Showing {len(filtered)} of {len(df)} exceptions")

    display_cols = [
        "Property", "Unit", "Resident", "Rule_ID", "Rule_Name",
        "Field", "ResMan_Value", "Expected_Value", "Severity", "Status", "Detail",
    ]
    # Only include columns that exist
    display_cols = [c for c in display_cols if c in filtered.columns]

    st.dataframe(
        filtered[display_cols].reset_index(drop=True),
        use_container_width=True,
        height=450,
        column_config={
            "Rule_ID":        st.column_config.TextColumn("Rule", width="small"),
            "Severity":       st.column_config.TextColumn("Severity", width="small"),
            "ResMan_Value":   st.column_config.TextColumn("ResMan Value", width="medium"),
            "Expected_Value": st.column_config.TextColumn("Expected Value", width="medium"),
            "Detail":         st.column_config.TextColumn("Detail", width="large"),
        },
    )

    # Download
    csv_bytes = filtered.to_csv(index=False).encode("utf-8")
    ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M")
    st.download_button(
        "⬇  Download CSV",
        data=csv_bytes,
        file_name=f"{key_prefix}_exceptions_{ts}.csv",
        mime="text/csv",
    )


# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tab_mi, tab_mo = st.tabs(["🔑  Move-In Audit", "📦  Move-Out Audit"])

# ── Move-In ─────────────────────────────────────────────────────────────────
with tab_mi:
    st.subheader("Move-In Audit")

    col_desc, col_btn = st.columns([4, 1])
    with col_desc:
        st.markdown(
            "Audits all **new move-ins** from the Leases report. "
            "Checks identity, dates, financial setup, "
            "and deposit collection."
        )
        st.info(
            "Document Upload Compliance is **Phase 2** — "
            "requires the ResMan Document Attachment export.",
            icon="ℹ️",
        )
    with col_btn:
        run_mi = st.button("▶  Run Audit", type="primary", key="btn_run_mi",
                           use_container_width=True)

    if run_mi:
        with st.spinner("Running Move-In audit..."):
            df_leases, df_activity, df_rr_units, df_rr_charges = load_all()
            result = run_move_in_audit(df_leases, df_activity, df_rr_units, df_rr_charges)
            st.session_state["df_mi"] = result

    if "df_mi" in st.session_state:
        render_results(st.session_state["df_mi"], "move_in")

# ── Move-Out ─────────────────────────────────────────────────────────────────
with tab_mo:
    st.subheader("Move-Out Audit")

    col_desc, col_btn = st.columns([4, 1])
    with col_desc:
        st.markdown(
            "Audits all **moved-out units** in the Rent Roll. "
            "Checks early lease break fees, final account reconciliation, "
            "collection readiness including the **Texas §92.103 30-day rule**, "
            "and refund payout accuracy."
        )
        st.info(
            "Damage Photo Compliance is **Phase 2** — "
            "requires the ResMan Document Attachment export.",
            icon="ℹ️",
        )
    with col_btn:
        run_mo = st.button("▶  Run Audit", type="primary", key="btn_run_mo",
                           use_container_width=True)

    if run_mo:
        with st.spinner("Running Move-Out audit..."):
            _, _, df_rr_units, df_rr_charges = load_all()
            result = run_move_out_audit(df_rr_units, df_rr_charges)
            st.session_state["df_mo"] = result

    if "df_mo" in st.session_state:
        render_results(st.session_state["df_mo"], "move_out")
