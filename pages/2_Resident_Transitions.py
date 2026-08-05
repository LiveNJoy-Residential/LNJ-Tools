"""
Resident Activity Audit Tool — Streamlit Dashboard
Upload-driven: all processing is in-memory, nothing stored server-side.
FRD v2.0 | July 2026
"""

import sys
import os
import io
import pandas as pd
import streamlit as st

_ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from resident_transition.utils import route_uploaded_files
from resident_transition.data_loader import (
    load_leases, load_activity, load_rent_roll,
    load_scheduled_move_ins, load_cancellations_move_outs,
    load_eviction_process, load_pet_summary, load_vehicles,
    load_recurring_projections,
)
from resident_transition.move_in_engine import run_move_in_audit
from resident_transition.move_out_engine import run_move_out_audit
from resident_transition.eviction_engine import run_eviction_audit

st.set_page_config(
    page_title="LNJ Resident Activity Audit",
    page_icon="\U0001f3e0",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    "<style>[data-testid='stSidebarNav'],[data-testid='stSidebarNavContainer'],"
    "[data-testid='stSidebarNavItems'],[data-testid='stSidebarNavLink']"
    "{display:none !important;}</style>",
    unsafe_allow_html=True,
)

# ─── SIDEBAR ───────────────────────────────────────────────────────────────────────────
_raw_uploads = []
routed: dict = {}

with st.sidebar:
    st.markdown("""
<style>
section[data-testid="stSidebar"] hr {
    margin-top: 0.35rem !important;
    margin-bottom: 0.35rem !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] > div > small {
    display: none !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
    padding: 0.4rem 0.75rem !important;
    min-height: unset !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploader"] {
    margin-bottom: 0.1rem !important;
}
</style>
""", unsafe_allow_html=True)

    if st.button("\u2190  Go back to Hub", key="back_to_hub", use_container_width=True):
        st.switch_page("app.py")
    st.divider()

    with st.expander("\U0001f4c2 Upload ResMan CSVs", expanded=False):
        st.caption(
            "Drop all ResMan CSV exports here \u2014 files are auto-detected by filename. "
            "Upload any mix of properties and report types."
        )
        _raw_uploads = st.file_uploader(
            "ResMan CSV Exports",
            type="csv",
            accept_multiple_files=True,
            help=(
                "Scheduled Move Ins \u00b7 New & Renewed Leases \u00b7 Cancellations & Move Outs \u00b7 "
                "Eviction Process \u00b7 Resident Activity \u00b7 Pet Summary \u00b7 Vehicles \u00b7 "
                "Rent Roll \u00b7 Recurring Transaction Projections"
            ),
        ) or []

        routed = route_uploaded_files(_raw_uploads)

        if _raw_uploads:
            _CATS = [
                ("Scheduled Move Ins",    "scheduled_move_ins"),
                ("New & Renewed Leases",  "leases"),
                ("Cancellations & MO",    "cancellations"),
                ("Eviction Process",      "evictions"),
                ("Resident Activity",     "activity"),
                ("Pet Summary",           "pet_summary"),
                ("Vehicles",              "vehicles"),
                ("Rent Roll",             "rent_rolls"),
                ("Recurring Projections", "recurring"),
            ]
            _ok, _no = "\u2705", "\u274c"
            st.markdown(
                "  \n".join(
                    f"{_ok if routed.get(k) else _no} **{lbl}** ({len(routed.get(k, []))})"
                    for lbl, k in _CATS
                )
            )
            _all_routed = {f.name for bucket in routed.values() for f in bucket}
            _unrouted = [f.name for f in _raw_uploads if f.name not in _all_routed]
            if _unrouted:
                st.warning(f"\u26a0\ufe0f {len(_unrouted)} unrecognized file(s) \u2014 check filenames: {', '.join(_unrouted)}")

    st.divider()

    _total_files = sum(len(routed.get(k, [])) for k in [
        "scheduled_move_ins", "leases", "cancellations", "evictions",
        "activity", "pet_summary", "vehicles", "rent_rolls", "recurring",
    ])
    _btn_label = (
        f"\U0001f680  Run Audit  ({_total_files} file{'s' if _total_files != 1 else ''} ready)"
        if _raw_uploads else "\U0001f680  Run Audit"
    )
    run_btn = st.button(
        _btn_label,
        type="primary",
        use_container_width=True,
        disabled=not bool(_raw_uploads),
    )

    if st.session_state.get("excel_bytes_rt"):
        _ts = st.session_state.get("run_ts_rt", pd.Timestamp.now().strftime("%Y%m%d_%H%M"))
        st.download_button(
            label="\U0001f4e5  Download Excel Report",
            data=st.session_state["excel_bytes_rt"],
            file_name=f"LNJ_ResidentAudit_{_ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# ─── HEADER ───────────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="display:flex;align-items:center;gap:14px;margin-bottom:0.25rem;">
<div style="width:48px;height:48px;background:#1A2744;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
<span style="color:#FFFFFF;font-size:1.4rem;font-weight:700;">&#x21C4;</span>
</div>
<h1 style="margin:0;padding:0;font-size:2rem;font-weight:700;line-height:1.2;">Resident Activity Audit Tool</h1>
</div>
""", unsafe_allow_html=True)
st.markdown(
    "Automated **Move-In \u00b7 Move-Out \u00b7 Eviction** lifecycle audit"
)
st.divider()


# ─── HELPERS ───────────────────────────────────────────────────────────────────────────
def _build_excel(df_mi, df_mo, df_ev) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for df, sheet in [
            (df_mi, "Move-In Exceptions"),
            (df_mo, "Move-Out Exceptions"),
            (df_ev, "Evictions"),
        ]:
            (df if df is not None and not df.empty else pd.DataFrame()).to_excel(
                writer, sheet_name=sheet, index=False)
    return buf.getvalue()


def render_results(df: pd.DataFrame, key_prefix: str):
    if df is None or df.empty:
        st.success("\u2705  No exceptions found.")
        return

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Exceptions", len(df))
    c2.metric("\U0001f534 CRITICAL", int((df["Severity"] == "CRITICAL").sum()))
    c3.metric("\U0001f7e0 HIGH",     int((df["Severity"] == "HIGH").sum()))
    c4.metric("\U0001f7e1 MEDIUM",   int((df["Severity"] == "MEDIUM").sum()))
    st.divider()

    props = ["All"] + sorted(df["Property"].dropna().unique().tolist())
    f1, f2 = st.columns(2)
    sel_prop = f1.selectbox("Property", props,                               key=f"{key_prefix}_prop")
    sel_sev  = f2.selectbox("Severity", ["All", "CRITICAL", "HIGH", "MEDIUM"], key=f"{key_prefix}_sev")

    mask = pd.Series([True] * len(df), index=df.index)
    if sel_prop != "All":
        mask &= df["Property"] == sel_prop
    if sel_sev != "All":
        mask &= df["Severity"] == sel_sev

    filtered = df[mask].copy()
    st.caption(f"Showing {len(filtered)} of {len(df)} exceptions")

    # Rule_Name omitted \u2014 Rule_ID is the compact identifier shown
    display_cols = [c for c in [
        "Property", "Unit", "Resident", "Rule_ID",
        "Field", "ResMan_Value", "Expected_Value", "Severity", "Detail",
    ] if c in filtered.columns]

    st.dataframe(
        filtered[display_cols].reset_index(drop=True),
        use_container_width=True,
        height=450,
        column_config={
            "Rule_ID":        st.column_config.TextColumn("Rule",           width="small"),
            "Severity":       st.column_config.TextColumn("Severity",       width="small"),
            "ResMan_Value":   st.column_config.TextColumn("ResMan Value",   width="medium"),
            "Expected_Value": st.column_config.TextColumn("Expected Value", width="medium"),
            "Detail":         st.column_config.TextColumn("Detail",         width="large"),
        },
    )

    st.download_button(
        "\u2b07  Download CSV",
        data=filtered.to_csv(index=False).encode("utf-8"),
        file_name=f"{key_prefix}_exceptions_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv",
        key=f"{key_prefix}_csv_dl",
    )


# ─── RUN ENGINE ───────────────────────────────────────────────────────────────────────────
if run_btn:
    with st.spinner("Ingesting CSVs and running audit engines\u2026"):
        try:
            df_leases             = load_leases(files=routed.get("leases", []))
            df_activity           = load_activity(files=routed.get("activity", []))
            df_rr_units, df_rr_ch = load_rent_roll(files=routed.get("rent_rolls", []))
            df_cancel             = load_cancellations_move_outs(files=routed.get("cancellations", []))
            df_ev                 = load_eviction_process(files=routed.get("evictions", []))
            df_pets               = load_pet_summary(files=routed.get("pet_summary", []))
            df_veh                = load_vehicles(files=routed.get("vehicles", []))
            df_recurring          = load_recurring_projections(files=routed.get("recurring", []))

            df_mi_result = run_move_in_audit(
                df_leases, df_activity, df_rr_units, df_rr_ch,
                df_pets=df_pets, df_vehicles=df_veh, df_recurring=df_recurring,
            )
            df_mo_result = run_move_out_audit(
                df_rr_units, df_rr_ch,
                df_cancel=df_cancel, df_ev=df_ev,
            )
            df_ev_flags  = run_eviction_audit(df_ev)

            st.session_state["df_mi"]          = df_mi_result
            st.session_state["df_mo"]          = df_mo_result
            st.session_state["df_ev"]          = df_ev_flags
            st.session_state["excel_bytes_rt"] = _build_excel(df_mi_result, df_mo_result, df_ev_flags)
            st.session_state["run_ts_rt"]      = pd.Timestamp.now().strftime("%Y%m%d_%H%M")
        except Exception as exc:
            st.error(f"\u274c  Engine error: {exc}")
            st.stop()

    st.success("\u2705  Audit complete \u2014 use the **Download Excel Report** button in the sidebar.")
    st.rerun()

if "df_mi" not in st.session_state:
    st.info("\U0001f448  Upload your ResMan CSV exports in the sidebar, then click **Run Audit** to begin.")
    st.stop()


# ─── RESULTS TABS ───────────────────────────────────────────────────────────────────────────
tab_mi, tab_mo, tab_ev = st.tabs(["\U0001f511  Move-In Audit", "\U0001f4e6  Move-Out Audit", "\u2696\ufe0f  Evictions"])

with tab_mi:
    st.subheader("Move-In Audit")
    st.markdown(
        "Audits all **new move-ins** from the Leases and Scheduled Move-Ins reports. "
        "Checks identity match, date timelines, financial setup, deposit collection, and auxiliary billing."
    )
    render_results(st.session_state.get("df_mi"), "move_in")

with tab_mo:
    st.subheader("Move-Out Audit")
    st.markdown(
        "Audits all **moved-out units**. Checks early lease break fees, "
        "final account reconciliation, Texas \u00a792.103 30-day refund rule, and refund accuracy."
    )
    render_results(st.session_state.get("df_mo"), "move_out")

with tab_ev:
    st.subheader("Evictions")
    st.markdown("Automated audit of all **Status E / UE** units from the Eviction Process report.")
    render_results(st.session_state.get("df_ev"), "eviction")
