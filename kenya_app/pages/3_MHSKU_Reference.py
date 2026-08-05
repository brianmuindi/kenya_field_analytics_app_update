"""
MHSKU Reference — Kenya Field Analytics
========================================
Persistent store for the Kimfay MHSKU workbook (Hyper / Large SPMKT /
Small SPMKT / Express / LMT Teir1 / LMT Teir2 / LMT Teir3) — the shared
must-have SKU + MBQ pressure target reference used by OSA Analytics and
the Report Generator.
"""

import streamlit as st
import sys, os
import pandas as pd
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from engines import (
    GLOBAL_CSS, inject_sidebar_toggle, inject_theme_toggle,
    load_ke_mhsku, ke_mhsku_save, ke_mhsku_load, ke_mhsku_clear,
    KE_MHSKU_SHEET_MAP,
)
from auth import require_login, logout, is_admin

st.set_page_config(page_title="MHSKU Reference", page_icon="📋", layout="wide")
require_login()
st.markdown(GLOBAL_CSS, unsafe_allow_html=True)
inject_sidebar_toggle()
inject_theme_toggle()

with st.sidebar:
    st.markdown("---")
    user = st.session_state.get("username", "user")
    role_badge = "🔑 Admin" if is_admin() else "👤 Analyst"
    st.markdown(
        f"<div style='font-size:12px;color:var(--muted);padding:0 8px 4px;'>"
        f"Signed in as <strong style='color:var(--text)'>{user}</strong> "
        f"<span style='color:#4f8ef7;font-size:11px'>({role_badge})</span></div>",
        unsafe_allow_html=True,
    )
    if st.button("🚪  Sign Out", key="logout_mhsku", use_container_width=True):
        logout()

st.markdown("""
<style>
.sku-meta   { font-size:12px; color:var(--muted); margin-bottom:24px; }
.admin-zone {
    background:rgba(248,113,113,.06); border:1px solid rgba(248,113,113,.2);
    border-radius:12px; padding:20px 24px; margin-top:32px;
}
.admin-zone h4 { color:#f87171; margin:0 0 6px 0; font-size:14px; }
.admin-zone p  { color:#9ca3af; font-size:13px; margin:0 0 16px 0; }
.empty-state {
    text-align:center; padding:60px 20px;
    color:var(--muted); border:1.5px dashed var(--border);
    border-radius:16px; margin-top:12px;
}
.empty-state .icon { font-size:40px; margin-bottom:12px; }
.empty-state h3    { font-size:16px; font-weight:600; color:var(--text); margin-bottom:6px; }
.upload-card {
    background:var(--surface); border:1px solid var(--border);
    border-radius:16px; padding:24px 28px; margin-bottom:24px;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="pg-title">
  <div class="eyebrow">Kenya Field Analytics</div>
  <h1>📋 MHSKU Reference</h1>
  <p>Upload the compiled MHSKU workbook to persist it as the shared must-have SKU + MBQ reference —
  used by OSA Analytics and the Report Generator. Sheets expected: Hyper, Large SPMKT, Small SPMKT,
  Express, LMT Teir1, LMT Teir2, LMT Teir3.</p>
</div>
""", unsafe_allow_html=True)

mhsku_map, last_updated, last_updated_by, last_filename = ke_mhsku_load()

# ── Upload card ───────────────────────────────────────────────────────────────
st.markdown('<div class="upload-card">', unsafe_allow_html=True)
st.markdown("#### 📤 Update MHSKU Reference")
st.markdown(
    "<p style='font-size:13px;color:var(--muted);margin-bottom:16px'>"
    "Upload the latest compiled MHSKU .xlsx to refresh the shared reference. "
    "The new workbook replaces the previous one.</p>",
    unsafe_allow_html=True,
)

uploaded = st.file_uploader(
    "Upload MHSKU .xlsx file", type=["xlsx"], key="mhsku_upload",
    label_visibility="collapsed",
    help="One sheet per outlet category, each with an 'EPR Code' column and an MBQ column.",
)

if uploaded:
    file_bytes = uploaded.read()
    try:
        parsed = load_ke_mhsku(file_bytes)
        n_skus = sum(len(v) for v in parsed.values())
        st.success(f"✅  Parsed **{n_skus:,} SKUs** across **{len(parsed)} outlet categories** from `{uploaded.name}` — ready to save.")

        with st.expander(f"Preview ({len(parsed)} categories)", expanded=False):
            for sheet, codes in parsed.items():
                st.markdown(f"**{sheet}** — {len(codes)} SKUs")
                prev_df = pd.DataFrame([
                    {"EPR Code": code, "Description": rec["description"],
                     "Brand": rec["brand"], "Category": rec["category"], "MBQ": rec["mbq"]}
                    for code, rec in list(codes.items())[:10]
                ])
                st.dataframe(prev_df, use_container_width=True, hide_index=True)

        if st.button("💾  Save to Reference Memory", key="save_mhsku", use_container_width=True):
            ke_mhsku_save(parsed, uploaded_by=st.session_state.get("username", "unknown"), filename=uploaded.name)
            st.success("✅  Reference saved! OSA Analytics and Report Generator will now use this MHSKU list.")
            st.rerun()
    except Exception as e:
        st.error(f"❌  Could not parse file: {e}")

st.markdown("</div>", unsafe_allow_html=True)

# ── Current reference view ────────────────────────────────────────────────────
st.markdown("---")

if last_updated:
    try:
        ts = datetime.fromisoformat(last_updated.replace("Z", "+00:00"))
        ts_fmt = ts.strftime("%d %b %Y, %H:%M UTC")
    except Exception:
        ts_fmt = last_updated
    total_skus = sum(len(v) for v in mhsku_map.values()) if mhsku_map else 0
    st.markdown(
        f"<div class='sku-meta'>Last updated: <strong style='color:var(--text)'>{ts_fmt}</strong>"
        f" &nbsp;·&nbsp; by <strong style='color:var(--text)'>{last_updated_by or 'unknown'}</strong>"
        f" &nbsp;·&nbsp; file <strong style='color:var(--text)'>{last_filename or '—'}</strong>"
        f" &nbsp;·&nbsp; <strong style='color:var(--text)'>{total_skus:,}</strong> SKUs stored</div>",
        unsafe_allow_html=True,
    )

if not mhsku_map:
    st.markdown("""
    <div class="empty-state">
      <div class="icon">📭</div>
      <h3>No MHSKU reference uploaded yet</h3>
      <p>Upload the compiled MHSKU .xlsx file above to create the shared reference.</p>
    </div>
    """, unsafe_allow_html=True)
else:
    tab_labels = [f"{sheet} ({len(codes)})" for sheet, codes in mhsku_map.items()]
    tabs = st.tabs(tab_labels)
    for tab, (sheet, codes) in zip(tabs, mhsku_map.items()):
        with tab:
            display_df = pd.DataFrame([
                {"EPR Code": code, "Description": rec["description"], "Brand": rec["brand"],
                 "Category": rec["category"], "MBQ (Pressure Target)": rec["mbq"]}
                for code, rec in codes.items()
            ])
            st.dataframe(display_df, use_container_width=True, hide_index=True,
                         column_config={"MBQ (Pressure Target)": st.column_config.NumberColumn(format="%d PCS")})
            st.caption(f"{len(display_df)} must-have SKUs — outlet categories mapped from CUSTOMER CATEGORY: "
                       f"{[k for k, v in KE_MHSKU_SHEET_MAP.items() if v == sheet]}")

# ── Admin zone — clear ────────────────────────────────────────────────────────
if is_admin():
    st.markdown("""
    <div class="admin-zone">
      <h4>🔑 Admin Controls</h4>
      <p>Clear the entire MHSKU reference store. This cannot be undone — you will need to re-upload the workbook.</p>
    </div>
    """, unsafe_allow_html=True)

    col_confirm, col_btn = st.columns([3, 1])
    with col_confirm:
        confirm_text = st.text_input(
            "Type CLEAR to confirm", placeholder="Type CLEAR to enable the button",
            key="clear_confirm", label_visibility="collapsed",
        )
    with col_btn:
        can_clear = confirm_text.strip().upper() == "CLEAR"
        if st.button("🗑️  Clear Memory", key="clear_mhsku", disabled=not can_clear, use_container_width=True):
            ke_mhsku_clear()
            st.success("✅  MHSKU reference cleared.")
            st.rerun()
else:
    st.markdown(
        "<div style='margin-top:24px;font-size:12px;color:var(--muted);text-align:center'>"
        "🔒 Clearing the reference requires admin access.</div>",
        unsafe_allow_html=True,
    )
