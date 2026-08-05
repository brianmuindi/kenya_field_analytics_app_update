"""
Report Generator — Kenya Field Analytics
Upload OSA data + MHSKU reference → download a fully formatted Excel report.
"""

import streamlit as st
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from engines import (
    GLOBAL_CSS, process_osa_kenya, build_osa_excel_kenya,
    ke_mhsku_load, inject_sidebar_toggle, inject_theme_toggle,
    osa_data_load, KE_KEY_ACCOUNTS,
    apply_ke_osa_filters, KE_FILTER_KEYS, KE_EXCLUDE_ACCOUNT_KEYWORDS,
)
from auth import require_login, logout
from datetime import datetime

st.set_page_config(page_title="Report Generator", page_icon="⬇️", layout="centered")
require_login()
st.markdown(GLOBAL_CSS, unsafe_allow_html=True)
inject_sidebar_toggle()
inject_theme_toggle()

with st.sidebar:
    st.markdown("---")
    user = st.session_state.get("username", "user")
    st.markdown(
        f"<div style='font-size:12px;color:var(--muted);padding:0 8px 4px;'>"
        f"Signed in as <strong style='color:var(--text)'>{user}</strong></div>",
        unsafe_allow_html=True,
    )
    if st.button("🚪  Sign Out", key="logout_rpt", use_container_width=True):
        logout()

st.markdown("""
<style>
.block-container { max-width: 720px !important; }
.rg-card {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: 16px; padding: 24px 28px 20px; margin-bottom: 20px;
}
.rg-card h3 { font-size:16px; font-weight:600; color:var(--text); margin:0 0 4px 0; }
.rg-card p  { font-size:13px; color:var(--muted); margin:0 0 18px 0; line-height:1.5; }
.info-grid  { display:grid; grid-template-columns:1fr 1fr; gap:12px; margin-top:16px; }
.info-box   { background:var(--bg); border:1px solid var(--border); border-radius:10px; padding:14px; }
.info-box .lbl { font-family:'DM Mono',monospace; font-size:10px; letter-spacing:1.5px;
                 text-transform:uppercase; color:var(--muted); margin-bottom:6px; }
.info-box .val { font-size:13px; color:var(--text); line-height:1.55; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="pg-title">
  <div class="eyebrow">Kenya Field Analytics</div>
  <h1>⬇️ Report Generator</h1>
  <p>Upload your OSA export and download a fully formatted Excel report instantly — scored against MHSKU / MBQ targets.</p>
</div>
""", unsafe_allow_html=True)

# ── MHSKU reference status ────────────────────────────────────────────────────
mhsku_map, mh_ts, mh_by, mh_fname = ke_mhsku_load()
if mhsku_map:
    n_skus = sum(len(v) for v in mhsku_map.values())
    st.markdown(
        f"<div class='result-box success'>✅ Using saved MHSKU reference — "
        f"<strong>{n_skus:,}</strong> must-have SKUs across {len(mhsku_map)} outlet categories "
        f"({mh_fname or 'workbook'}, saved by {mh_by or '?'}). "
        f"Go to <strong>MHSKU Reference</strong> to update it.</div>",
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        "<div class='result-box info'>⚠️ No MHSKU reference saved yet — upload one on the "
        "<strong>MHSKU Reference</strong> page first, or the report will fall back to a flat "
        "default pressure target instead of per-category MBQ.</div>",
        unsafe_allow_html=True,
    )

st.markdown("""
<div class="rg-card">
  <h3>📄 Kenya OSA Report</h3>
  <p>Upload an OSA export (.csv, .xls, .xlsx). Kenya regions only — Kampala/Uganda rows and
  "Not Listed" SKUs are dropped automatically. Only Must-Have SKUs are scored, against each
  outlet category's MBQ pressure target.</p>
</div>
""", unsafe_allow_html=True)

saved_df, saved_meta = osa_data_load()
use_saved = False
if saved_df is not None:
    col_use, _ = st.columns([3, 1])
    use_saved = col_use.checkbox(
        f"Use saved data — {saved_meta.get('filename','file')} "
        f"({saved_meta.get('rows',0):,} rows, saved {saved_meta.get('saved_at','')[:16].replace('T',' ')})",
        value=True, key="rpt_use_saved",
    )

uploaded = None
if not use_saved:
    uploaded = st.file_uploader("Upload OSA file", type=["xls", "xlsx", "csv"], key="rpt_osa_upload")

must_have_only = st.toggle("Must-Have SKUs only (recommended)", value=True, key="rpt_mh_only")

# ── Carry over the filters chosen on the OSA Analytics page ───────────────────
# Both pages share the same session_state keys (KE_FILTER_KEYS), so whatever
# outlets/regions/accounts you narrowed to on the analytics page can flow
# straight into this report.
_sel = {name: (st.session_state.get(key) or None) for name, key in KE_FILTER_KEYS.items()}
_active = {
    'Months': _sel['months'], 'Regions': _sel['regions'],
    'Outlet Categories': _sel['ocats'], 'Brands': _sel['brands'],
    'Key Accounts': _sel['accts'],
}
_active = {k: v for k, v in _active.items() if v}

apply_filters = False
if _active:
    apply_filters = st.toggle(
        "Match the OSA Analytics filters (same outlets/accounts as on-screen)",
        value=True, key="rpt_apply_filters",
        help="Builds the report from exactly the rows currently selected on the OSA Analytics page.",
    )
    with st.expander("Filters carried from OSA Analytics", expanded=False):
        for k, v in _active.items():
            shown = ", ".join(map(str, v[:12])) + (" …" if len(v) > 12 else "")
            st.markdown(
                f"<div style='font-size:12px;margin-bottom:4px'>"
                f"<strong>{k}:</strong> <span style='color:var(--muted)'>{shown}</span></div>",
                unsafe_allow_html=True,
            )
else:
    st.caption("No analysis filters set yet — open the OSA Analytics page and narrow the "
               "filters there to carry them into this report.")

if st.button("🛠️  Generate Report", key="rpt_generate", use_container_width=True,
             disabled=(uploaded is None and not use_saved)):
    try:
        with st.spinner("Processing OSA data and building the report…"):
            if use_saved and saved_df is not None:
                df = saved_df
                # Defensive cleanup for older saved snapshots: drop excluded
                # accounts/outlets (Bestmart, MUHINDI MWEUSI WITEITHIE) that may
                # still be baked into the saved file. Fresh uploads are already
                # filtered inside process_osa_kenya.
                _excl = '|'.join(KE_EXCLUDE_ACCOUNT_KEYWORDS)
                for _col in ('CUSTOMER NAME', 'ACCOUNT'):
                    if _col in df.columns:
                        _norm = df[_col].astype(str).str.strip().str.lower()
                        df = df[~_norm.str.contains(_excl, regex=True, na=False)].copy()
            else:
                raw = uploaded.read()
                df = process_osa_kenya(raw, uploaded.name, mhsku_map=mhsku_map, must_have_only=must_have_only)
            if df.empty:
                st.error("❌ No rows left after filtering — check the MHSKU reference matches this file's PRODUCT CODEs.")
                st.stop()
            if apply_filters and _active:
                df = apply_ke_osa_filters(
                    df, months=_sel['months'], regions=_sel['regions'],
                    ocats=_sel['ocats'], brands=_sel['brands'], accts=_sel['accts'],
                )
                if df.empty:
                    st.error("❌ No rows left after applying the OSA Analytics filters. "
                             "Widen the filters on the analytics page or turn off filter matching.")
                    st.stop()
            xlsx_bytes = build_osa_excel_kenya(df)
        st.success(f"✅ Report ready — {len(df):,} rows, {df['Month'].nunique()} month(s).")
        st.download_button(
            "⬇️ Download Kenya OSA Report (.xlsx)",
            data=xlsx_bytes,
            file_name=f"Kenya_OSA_Report_{datetime.utcnow().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        with st.expander("Report contents"):
            st.markdown(f"""
<div class="info-grid">
  <div class="info-box"><div class="lbl">Rows scored</div><div class="val">{len(df):,}</div></div>
  <div class="info-box"><div class="lbl">Months</div><div class="val">{', '.join(sorted(df['Month'].dropna().unique()))}</div></div>
  <div class="info-box"><div class="lbl">Outlet categories</div><div class="val">{', '.join(sorted(df['MHSKU_SHEET'].dropna().unique())) if 'MHSKU_SHEET' in df.columns else '—'}</div></div>
  <div class="info-box"><div class="lbl">Key accounts tracked</div><div class="val">{', '.join(c for c, _ in KE_KEY_ACCOUNTS)}</div></div>
</div>
""", unsafe_allow_html=True)
    except Exception as e:
        st.error(f"❌ {e}")

st.markdown('<div class="footer">Kenya Field Analytics — Report Generator</div>', unsafe_allow_html=True)
