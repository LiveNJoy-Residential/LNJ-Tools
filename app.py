"""
LiveNjoy Automation Suite — Tool Hub  (app.py)
===============================================
Entry point. Navigates to the Revenue & Concession Audit Tool
and the Resident Activity Audit Tool.

Run: .venv\Scripts\streamlit.exe run app.py
"""

import streamlit as st
import streamlit.components.v1 as components

# ─── PAGE CONFIG ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="LiveNJoy Tools",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── NAVIGATION STATE ────────────────────────────────────────────────────────
if "page" not in st.session_state:
    st.session_state["page"] = "home"
_page = st.session_state["page"]


# ─── HOME SCREEN ──────────────────────────────────────────────────────────────
if _page == "home":
    st.markdown("""
<style>
/* ── Hub colour tokens — light (default) ── */
:root {
  --hub-outer:          #0D1B2E;
  --hub-inner:          #FFFFFF;
  --hub-border:         #E5E9EF;
  --hub-txt1:           #1A2744;
  --hub-txt2:           #6B7280;
  --hub-txt3:           #374151;
  --hub-txt-muted:      #9CA3AF;
  --hub-logo:           #1A2744;
  --hub-btn-hover:      #243357;
  --hub-btn-dis-bg:     #F0F0F0;
  --hub-btn-dis-txt:    #AAAAAA;
  --hub-btn-dis-bdr:    #E0E0E0;
}
/* ── Dark — OS/system preference ── */
@media (prefers-color-scheme: dark) { :root {
  --hub-outer:          #060C17;
  --hub-inner:          #111827;
  --hub-border:         #1F2D42;
  --hub-txt1:           #E2E8F0;
  --hub-txt2:           #94A3B8;
  --hub-txt3:           #CBD5E1;
  --hub-txt-muted:      #4B5563;
  --hub-logo:           #2C4A8A;
  --hub-btn-hover:      #3A5EA8;
  --hub-btn-dis-bg:     #1A2438;
  --hub-btn-dis-txt:    #4B5E78;
  --hub-btn-dis-bdr:    #1F2D42;
}}
/* ── Structural overrides ── */
[data-testid="stSidebar"]        { display: none !important; }
[data-testid="collapsedControl"] { display: none !important; }
.stApp { background-color: var(--hub-outer) !important; }
section[data-testid="stMain"] {
    padding: 1.25rem !important;
    background-color: var(--hub-outer) !important;
}
.block-container {
    background: var(--hub-inner) !important;
    border-radius: 16px !important;
    max-width: 100% !important;
    width: 100% !important;
    min-height: calc(100vh - 2.5rem) !important;
    padding: 2rem 2.5rem 4rem !important;
    margin: 0 !important;
    box-sizing: border-box !important;
    position: relative !important;
}
/* ── Buttons ── */
div[data-testid="stButton"] > button {
    border-radius: 8px !important;
    font-size: 0.95rem !important;
    font-weight: 500 !important;
    padding: 0.75rem 1rem !important;
    width: 100% !important;
    border: none !important;
}
div[data-testid="stButton"] > button[kind="primary"] {
    background-color: var(--hub-logo) !important;
    color: #FFFFFF !important;
}
div[data-testid="stButton"] > button[kind="primary"]:hover {
    background-color: var(--hub-btn-hover) !important;
}
div[data-testid="stButton"] > button[disabled] {
    background-color: var(--hub-btn-dis-bg) !important;
    color: var(--hub-btn-dis-txt) !important;
    border: 1px solid var(--hub-btn-dis-bdr) !important;
    cursor: not-allowed !important;
}
[data-testid="column"] [data-testid="stButton"] {
    border-left: 1.5px solid var(--hub-border) !important;
    border-right: 1.5px solid var(--hub-border) !important;
    border-bottom: 1.5px solid var(--hub-border) !important;
    border-radius: 0 0 14px 14px !important;
    overflow: hidden !important;
    margin-top: 0 !important;
    padding: 0 0 0.1rem !important;
}
[data-testid="column"] [data-testid="stButton"] > button {
    border-radius: 0 !important;
    margin: 0 !important;
}
div[data-testid="stLinkButton"] > a {
    border-radius: 8px !important;
    font-size: 0.95rem !important;
    font-weight: 500 !important;
    padding: 0.75rem 1rem !important;
    width: 100% !important;
    background-color: var(--hub-logo) !important;
    color: #FFFFFF !important;
    text-decoration: none !important;
    display: flex !important;
    justify-content: center !important;
    align-items: center !important;
    border: none !important;
}
div[data-testid="stLinkButton"] > a:hover {
    background-color: var(--hub-btn-hover) !important;
}
[data-testid="column"] [data-testid="stLinkButton"] {
    border-left: 1.5px solid var(--hub-border) !important;
    border-right: 1.5px solid var(--hub-border) !important;
    border-bottom: 1.5px solid var(--hub-border) !important;
    border-radius: 0 0 14px 14px !important;
    overflow: hidden !important;
    margin-top: 0 !important;
    padding: 0 0 0.1rem !important;
}
[data-testid="column"] [data-testid="stLinkButton"] > a {
    border-radius: 0 !important;
    margin: 0 !important;
}
/* ── Status badges ── */
.hub-badge-active {
    background:#F0FDF4; color:#16A34A; border:1.5px solid #16A34A;
    padding:0.2rem 0.65rem; border-radius:20px; font-size:0.68rem;
    font-weight:700; letter-spacing:0.1em; text-transform:uppercase;
    white-space:nowrap; flex-shrink:0; margin-left:8px; display:inline-block;
}
.hub-badge-dev {
    background:#FFFBEB; color:#B45309; border:1.5px solid #B45309;
    padding:0.2rem 0.65rem; border-radius:20px; font-size:0.68rem;
    font-weight:700; letter-spacing:0.1em; text-transform:uppercase;
    white-space:nowrap; flex-shrink:0; margin-left:8px; display:inline-block;
}
@media (prefers-color-scheme: dark) {
    .hub-badge-active { background:#052E16; color:#4ADE80; border-color:#4ADE80; }
    .hub-badge-dev    { background:#1C1400; color:#FBB042; border-color:#FBB042; }
}
[data-hub-theme="dark"] .hub-badge-active  { background:#052E16; color:#4ADE80; border-color:#4ADE80; }
[data-hub-theme="dark"] .hub-badge-dev     { background:#1C1400; color:#FBB042; border-color:#FBB042; }
[data-hub-theme="light"] .hub-badge-active { background:#F0FDF4; color:#16A34A; border-color:#16A34A; }
[data-hub-theme="light"] .hub-badge-dev    { background:#FFFBEB; color:#B45309; border-color:#B45309; }
</style>
""", unsafe_allow_html=True)

    st.markdown("""
<div style="display:flex;justify-content:space-between;align-items:center;padding:0.25rem 0 1.25rem;">
<div style="display:flex;align-items:center;gap:12px;">
<div style="width:44px;height:44px;background:var(--hub-logo);border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
<span style="color:#FFFFFF;font-size:1.1rem;font-weight:900;letter-spacing:-0.03em;">LJ</span>
</div>
<div>
<div style="font-size:0.95rem;font-weight:700;color:var(--hub-txt1);letter-spacing:0.01em;line-height:1.3;">LiveNJoy Automation Suite</div>
<div style="font-size:0.68rem;color:var(--hub-txt-muted);letter-spacing:0.14em;text-transform:uppercase;line-height:1;">Enterprise Operations Hub</div>
</div>
</div>
<div style="display:flex;align-items:center;gap:8px;">
<span style="width:10px;height:10px;background:#22C55E;border-radius:50%;display:inline-block;flex-shrink:0;"></span>
<div style="text-align:right;line-height:1.4;">
<div style="font-size:0.75rem;font-weight:700;color:var(--hub-txt1);letter-spacing:0.1em;text-transform:uppercase;">Connected</div>
<div style="font-size:0.67rem;color:var(--hub-txt-muted);letter-spacing:0.1em;text-transform:uppercase;">ResMan Environment</div>
</div>
</div>
</div>
<hr style="border:none;border-top:1px solid var(--hub-border);margin:0 0 1.5rem;">
""", unsafe_allow_html=True)

    col1, col2 = st.columns(2, gap="medium")

    with col1:
        st.markdown("""
<div style="border:1.5px solid var(--hub-border);border-bottom:none;border-radius:14px 14px 0 0;padding:1.5rem 1.5rem 1.25rem;background:var(--hub-inner);margin-bottom:0;">
<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:0.9rem;">
<div style="display:flex;align-items:center;gap:12px;">
<div style="width:44px;height:44px;background:var(--hub-logo);border-radius:9px;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
<span style="color:#FFFFFF;font-size:1.25rem;font-weight:700;">$</span>
</div>
<div style="font-size:1.12rem;font-weight:700;color:var(--hub-txt1);line-height:1.35;">Revenue &amp; Concession<br>Audit Tool</div>
</div>
<span class="hub-badge-active">Active</span>
</div>
<p style="color:var(--hub-txt2);font-size:0.85rem;margin:0 0 0.85rem;line-height:1.4;">ResMan Revenue Integrity &amp; Exception Analytics</p>
<ul style="margin:0;padding-left:1.1rem;color:var(--hub-txt3);font-size:0.84rem;line-height:2.0;list-style-type:disc;">
<li>Automated Across 7 ResMan CSV Reports</li>
<li>Exposure Mapping (Critical, High, Medium Severity Flags)</li>
<li>Automated Excel Report Generation</li>
</ul>
</div>
""", unsafe_allow_html=True)
        if st.button("Open Audit Tool  \u2192", key="go_audit", type="primary", width="stretch"):
            st.switch_page("pages/1_Audit_Tool.py")

    with col2:
        st.markdown("""
<div style="border:1.5px solid var(--hub-border);border-bottom:none;border-radius:14px 14px 0 0;padding:1.5rem 1.5rem 1.25rem;background:var(--hub-inner);margin-bottom:0;">
<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:0.9rem;">
<div style="display:flex;align-items:center;gap:12px;">
<div style="width:44px;height:44px;background:var(--hub-logo);border-radius:9px;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
<span style="color:#FFFFFF;font-size:1.25rem;font-weight:700;">&#x21C4;</span>
</div>
<div style="font-size:1.12rem;font-weight:700;color:var(--hub-txt1);line-height:1.35;">Resident Activity<br>Audit Tool</div>
</div>
<span class="hub-badge-dev">Under Development</span>
</div>
<p style="color:var(--hub-txt2);font-size:0.85rem;margin:0 0 0.85rem;line-height:1.4;">Move-In &amp; Move-Out Workflow Automation</p>
<ul style="margin:0;padding-left:1.1rem;color:var(--hub-txt3);font-size:0.84rem;line-height:2.0;list-style-type:disc;">
<li>Lease Addendum Verification &amp; Unit Inspection Checks</li>
<li>Resident Onboarding Checklist &amp; Ledger Reconciliation</li>
<li>Streamlined Deposit Refund &amp; Unit Turn Tracking</li>
</ul>
</div>
""", unsafe_allow_html=True)
        if st.button("Open Transition Tool  \u2192", key="go_resident", type="primary", width="stretch"):
            st.switch_page("pages/2_Resident_Transitions.py")

    st.markdown("""
<div style="position:fixed;bottom:1.25rem;left:1.25rem;right:1.25rem;display:flex;justify-content:space-between;align-items:center;border-top:1px solid var(--hub-border);padding:0.9rem 2.5rem 0;background:var(--hub-inner);z-index:999;">
<span style="font-size:0.67rem;color:var(--hub-txt-muted);letter-spacing:0.1em;text-transform:uppercase;">LiveNJoy Residential LLC &copy; &middot; 2026 &middot; Internal Use Only</span>
<span style="font-size:0.67rem;color:var(--hub-txt-muted);letter-spacing:0.1em;text-transform:uppercase;">System Status &middot; <span style="color:#22C55E;font-weight:700;">Operational</span></span>
</div>
""", unsafe_allow_html=True)

    # ── Theme sync: reads Streamlit's actual theme and applies CSS vars as
    #    inline styles on <html> — inline styles beat any stylesheet rule.
    components.html("""
<script>
(function() {
    var doc = window.parent.document;
    var root = doc.documentElement;
    var LIGHT = {
        '--hub-outer':'#0D1B2E','--hub-inner':'#FFFFFF',
        '--hub-border':'#E5E9EF','--hub-txt1':'#1A2744',
        '--hub-txt2':'#6B7280','--hub-txt3':'#374151',
        '--hub-txt-muted':'#9CA3AF','--hub-logo':'#1A2744',
        '--hub-btn-hover':'#243357','--hub-btn-dis-bg':'#F0F0F0',
        '--hub-btn-dis-txt':'#AAAAAA','--hub-btn-dis-bdr':'#E0E0E0'
    };
    var DARK = {
        '--hub-outer':'#060C17','--hub-inner':'#111827',
        '--hub-border':'#1F2D42','--hub-txt1':'#E2E8F0',
        '--hub-txt2':'#94A3B8','--hub-txt3':'#CBD5E1',
        '--hub-txt-muted':'#4B5563','--hub-logo':'#2C4A8A',
        '--hub-btn-hover':'#3A5EA8','--hub-btn-dis-bg':'#1A2438',
        '--hub-btn-dis-txt':'#4B5E78','--hub-btn-dis-bdr':'#1F2D42'
    };
    function getTheme() {
        // Priority 1 – Streamlit's data-theme attribute (v1.36+)
        var attr = root.getAttribute('data-theme');
        if (attr === 'light' || attr === 'dark') return attr;
        // Priority 2 – Streamlit's injected --background-color CSS variable
        var bg = window.parent.getComputedStyle(root)
            .getPropertyValue('--background-color').trim().toLowerCase();
        if (bg === '#ffffff' || bg === 'white' || bg === 'rgb(255, 255, 255)') return 'light';
        if (bg) return 'dark';
        // Priority 3 – OS preference
        return window.parent.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
    }
    function apply() {
        var theme = getTheme();
        var tokens = theme === 'light' ? LIGHT : DARK;
        for (var k in tokens) root.style.setProperty(k, tokens[k]);
        root.setAttribute('data-hub-theme', theme);
    }
    apply();
    // Watch for Streamlit theme changes
    new MutationObserver(apply).observe(root, {
        attributes: true, attributeFilter: ['data-theme', 'class']
    });
    setInterval(apply, 500);
})();
</script>
""", height=0, scrolling=False)

    st.stop()
