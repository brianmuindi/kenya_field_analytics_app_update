"""
Home page — Kenya Field Analytics
"""

import streamlit as st
import sys, os
_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _HERE)
from engines import GLOBAL_CSS, THEME, inject_sidebar_toggle, inject_theme_toggle
from auth import require_login, logout

st.set_page_config(
    page_title="Kenya Field Analytics",
    page_icon="🇰🇪",
    layout="wide",
)

require_login()  # ← stops here and shows login if not authenticated

st.markdown(GLOBAL_CSS, unsafe_allow_html=True)
inject_sidebar_toggle()
inject_theme_toggle()

# ── Logout button in sidebar ──────────────────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    user = st.session_state.get("username", "user")
    st.markdown(f"<div style='font-size:12px;color:var(--muted);padding:0 8px 4px;'>Signed in as <strong style='color:var(--text)'>{user}</strong></div>", unsafe_allow_html=True)
    if st.button("🚪  Sign Out", key="logout_home", use_container_width=True):
        logout()

st.markdown("""
<style>
.home-hero { text-align:center; padding:50px 20px 20px; }
.home-hero .eyebrow { font-family:'DM Mono',monospace; font-size:11px; letter-spacing:3px; color:#4f8ef7; text-transform:uppercase; margin-bottom:10px; }
.home-hero h1 { font-size:42px; font-weight:700; color:#e8eaf0; margin:0 0 12px 0; line-height:1.1; }
.home-hero p  { font-size:16px; color:#6b7280; max-width:520px; margin:0 auto 32px; }
.legend-row  { display:flex; gap:16px; justify-content:center; margin-top:32px; flex-wrap:wrap; }
.legend-pill { display:flex; align-items:center; gap:8px; background:#1a1d27; border:1px solid #2a2d3a; border-radius:20px; padding:8px 16px; font-size:13px; color:#c8cad4; }
.dot { width:10px; height:10px; border-radius:50%; flex-shrink:0; }

section[data-testid="column"] div[data-testid="stButton"] > button {
    background:#1a1d27 !important; border:1.5px solid #2a2d3a !important;
    border-radius:16px !important; padding:32px 24px 24px !important;
    text-align:center !important; width:100% !important;
    color:#e8eaf0 !important; font-size:15px !important; font-weight:500 !important;
    height:auto !important; line-height:1.6 !important; white-space:pre-wrap !important;
}
section[data-testid="column"] div[data-testid="stButton"] > button:hover {
    background:rgba(79,142,247,.08) !important; border-color:#4f8ef7 !important; transform:none !important;
}

/* logout button in sidebar */
section[data-testid="stSidebar"] div[data-testid="stButton"] > button {
    background:rgba(248,113,113,.1) !important; border:1px solid rgba(248,113,113,.3) !important;
    color:#fca5a5 !important; font-size:13px !important; padding:8px !important;
    border-radius:8px !important;
}
section[data-testid="stSidebar"] div[data-testid="stButton"] > button:hover {
    background:rgba(248,113,113,.2) !important;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="home-hero">
  <div class="eyebrow">Kenya Field Analytics</div>
  <h1>🇰🇪 Report Generator</h1>
  <p>Upload your Kenya OSA data, visualise performance against MHSKU must-haves + MBQ targets, and download formatted Excel reports — no coding required.</p>
</div>
""", unsafe_allow_html=True)

c1, c2, c3 = st.columns(3)
with c1:
    if st.button("🟢\n\nOSA Analytics\n\nOn-Shelf Availability — scored against MHSKU must-haves & MBQ, by outlet category, region & key account.\n\n→ Open", key="nav_osa", use_container_width=True):
        st.switch_page("pages/1_OSA_Analytics.py")
with c2:
    if st.button("⬇️\n\nReport Generator\n\nDownload a fully formatted Kenya OSA Excel report instantly.\n\n→ Open", key="nav_rpt", use_container_width=True):
        st.switch_page("pages/2_Report_Generator.py")
with c3:
    if st.button("📋\n\nMHSKU Reference\n\nMust-Have SKU list with MBQ targets by outlet category — shared persistent memory.\n\n→ Open", key="nav_mhsku", use_container_width=True):
        st.switch_page("pages/3_MHSKU_Reference.py")

st.markdown("""
<div class="legend-row">
  <div class="legend-pill"><div class="dot" style="background:#34c97b"></div> On Target ≥ 95%</div>
  <div class="legend-pill"><div class="dot" style="background:#f59e0b"></div> Near Target 80–94%</div>
  <div class="legend-pill"><div class="dot" style="background:#f87171"></div> Below Target &lt; 80%</div>
</div>
<div class="footer" style="margin-top:24px">Click a card above — or use the sidebar on the left.</div>
""", unsafe_allow_html=True)
