"""OSA Analytics — Kenya Field Analytics — On-Shelf Availability against MHSKU / MBQ targets"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from engines import (
    GLOBAL_CSS, get_theme, get_plotly_layout, osa_color,
    process_osa_kenya, load_ke_mhsku, ke_mhsku_save, ke_mhsku_load, ke_mhsku_clear,
    inject_sidebar_toggle, inject_theme_toggle,
    osa_data_save, osa_data_load, osa_data_clear,
    KE_MHSKU_SHEET_MAP, KE_OOS_REASONS, KE_OOS_REASON_GUIDANCE,
    apply_ke_osa_filters, KE_EXCLUDE_ACCOUNT_KEYWORDS,
)
from auth import require_login, logout

st.set_page_config(page_title="OSA Analytics", page_icon="🇰🇪", layout="wide")
require_login()
st.markdown(GLOBAL_CSS, unsafe_allow_html=True)
inject_sidebar_toggle()
inject_theme_toggle()

with st.sidebar:
    st.markdown("---")
    user = st.session_state.get("username", "user")
    st.markdown(f"<div style='font-size:12px;color:var(--muted);padding:0 8px 4px;'>Signed in as <strong style='color:var(--text)'>{user}</strong></div>", unsafe_allow_html=True)
    if st.button("🚪  Sign Out", key="logout_ke_osa", use_container_width=True):
        logout()

# ── helpers ───────────────────────────────────────────────────────────────────
def T(): return get_theme()

def apply_layout(fig, title="", height=380):
    th = T()
    fig.update_layout(**get_plotly_layout(height=height, title=title))
    fig.update_layout(hoverlabel=dict(bgcolor=th["surface"], font_color=th["text"], bordercolor=th["border"]))
    fig.update_xaxes(tickfont_color=th["text"], title_font_color=th["text"])
    fig.update_yaxes(tickfont_color=th["text"], title_font_color=th["text"])
    return fig

def color_bars(values): return [osa_color(v) for v in values]

def synced_multiselect(container, label, options, key):
    """A multiselect that defaults to "all selected" and re-defaults
    whenever the available options actually change (new file uploaded,
    upstream filter narrowed the data) — without fighting the user if
    they deliberately deselect everything on the *current* options.

    Streamlit ignores `default=` once a key already exists in
    session_state, so on its own a stale selection from a previous
    dataset can silently persist and make every downstream `.isin()`
    filter return zero rows ("No data matches the selected filters"
    even though data was clearly loaded). We track a snapshot of the
    options list alongside the selection; only when the options
    themselves change do we reset the selection back to "all".
    """
    options = list(options)
    snap_key = f"{key}__opts_snapshot"
    if st.session_state.get(snap_key) != options:
        st.session_state[key] = options
        st.session_state[snap_key] = options
    return container.multiselect(label, options, key=key)

PALETTE = lambda th: [th['blue'], th['green'], th['amber'], th['purple'],
                       th['cyan'], th['red'], '#fb923c', '#e879f9', '#a3e635', '#38bdf8']

# ── page header ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="pg-title">
  <div class="eyebrow">Kenya Field Analytics</div>
  <h1>🇰🇪 Kenya OSA Analytics</h1>
  <p>On-Shelf Availability scored against Must-Have SKUs (MHSKU) and their per-outlet-category MBQ pressure targets. Kampala/Uganda rows, Bestmart accounts, and "Not Listed" SKUs are dropped automatically — we don't track those.</p>
</div>
""", unsafe_allow_html=True)

# ── MHSKU reference workbook (persistent) ────────────────────────────────────
saved_map, mh_ts, mh_by, mh_fname = ke_mhsku_load()

with st.expander("📋  MHSKU Reference (Must-Haves + MBQ by outlet category)", expanded=(not saved_map)):
    if saved_map:
        n_skus = sum(len(v) for v in saved_map.values())
        st.markdown(
            f"<div style='font-size:13px;color:#34c97b;background:rgba(52,201,123,.08);"
            f"border:1px solid rgba(52,201,123,.2);border-radius:8px;padding:8px 14px;margin-bottom:8px'>"
            f"✅ <strong>{mh_fname or 'MHSKU workbook'}</strong> — {n_skus:,} must-have SKUs across "
            f"{len(saved_map)} outlet categories · saved {mh_ts[:16].replace('T',' ')} by <strong>{mh_by or '?'}</strong></div>",
            unsafe_allow_html=True,
        )
        cnts = ", ".join(f"{k}: {len(v)}" for k, v in saved_map.items())
        st.caption(cnts)
        cu, cc = st.columns([3, 1])
        use_saved_mh = cu.checkbox("Use saved MHSKU reference (skip upload)", value=True, key="ke_mh_use_saved")
        if cc.button("🗑️ Clear", key="ke_mh_clear"):
            ke_mhsku_clear()
            st.rerun()
    else:
        st.info("No MHSKU reference saved yet. Upload the compiled MHSKU workbook below "
                "(sheets: Hyper, Large SPMKT, Small SPMKT, Express, LMT Teir1, LMT Teir2, LMT Teir3).")
        use_saved_mh = False

    mh_upload = st.file_uploader("Upload MHSKU workbook (.xlsx)", type=["xlsx"], key="ke_mh_upload")
    if mh_upload is not None:
        try:
            mh_bytes = mh_upload.read()
            new_map = load_ke_mhsku(mh_bytes)
            ke_mhsku_save(new_map, uploaded_by=st.session_state.get("username", "user"),
                          filename=mh_upload.name)
            st.success(f"✅ MHSKU reference loaded — "
                       f"{sum(len(v) for v in new_map.values()):,} SKUs across {len(new_map)} categories.")
            st.rerun()
        except Exception as e:
            st.error(f"❌ Could not parse MHSKU workbook: {e}")

mhsku_map = saved_map if (saved_map and use_saved_mh) else {}
if not mhsku_map:
    st.markdown('<div class="result-box info">⬆️ Upload (or reuse) the MHSKU reference workbook above to score OSA against must-haves.</div>', unsafe_allow_html=True)

# ── OSA data upload ───────────────────────────────────────────────────────────
saved_df, saved_meta = osa_data_load()

with st.expander("💾  Saved OSA Data", expanded=(saved_df is None)):
    if saved_df is not None:
        st.markdown(
            f"<div style='font-size:13px;color:#34c97b;background:rgba(52,201,123,.08);"
            f"border:1px solid rgba(52,201,123,.2);border-radius:8px;padding:8px 14px;margin-bottom:8px'>"
            f"✅ <strong>{saved_meta.get('filename','file')}</strong> — "
            f"{saved_meta.get('rows',0):,} rows · saved {saved_meta.get('saved_at','')[:16].replace('T',' ')}"
            f" by <strong>{saved_meta.get('uploaded_by','?')}</strong></div>",
            unsafe_allow_html=True,
        )
        col_use, col_clear = st.columns([3, 1])
        use_saved = col_use.checkbox("Use saved data (skip upload)", value=True, key="ke_osa_use_saved")
        if col_clear.button("🗑️ Clear saved data", key="ke_osa_clear"):
            osa_data_clear()
            st.rerun()
    else:
        st.info("No saved data yet. Upload an OSA export below and it will be saved automatically.")
        use_saved = False

col_main, col_toggle = st.columns([3, 2])
with col_main:
    st.markdown('<div class="sec-label">OSA Data File (.csv, .xls, .xlsx)</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload OSA file", type=["xls", "xlsx", "csv"],
                                 label_visibility="collapsed", key="ke_osa_upload")
with col_toggle:
    st.markdown('<div class="sec-label">Scope</div>', unsafe_allow_html=True)
    must_have_only = st.toggle("Must-Have SKUs only (recommended)", value=True, key="ke_osa_mh_only",
                                help="When on, only rows whose SKU is a must-have for that outlet's category are kept, and the pressure target used is that category's MBQ.")

@st.cache_data(show_spinner="Processing Kenya OSA data…", max_entries=3, ttl=3600)
def load_osa_ke(file_bytes: bytes, filename: str, mhsku_map: dict, must_have_only: bool):
    from engines import process_osa_kenya
    return process_osa_kenya(file_bytes, filename, mhsku_map=mhsku_map, must_have_only=must_have_only)

if uploaded is not None:
    try:
        raw = uploaded.read()
        df = load_osa_ke(raw, uploaded.name, mhsku_map, must_have_only)
        del raw
        osa_data_save(df, filename=uploaded.name, uploaded_by=st.session_state.get("username", "user"))
        st.success(f"✅ **{uploaded.name}** loaded and saved — {len(df):,} rows "
                   f"({'must-have SKUs only' if must_have_only else 'all rows'}).")
    except Exception as e:
        st.error(f"❌ {e}")
        st.stop()
elif saved_df is not None and use_saved:
    df = saved_df.copy()
else:
    st.markdown('<div class="result-box info">⬆️ Upload your Kenya OSA export to continue, or enable saved data above.</div>', unsafe_allow_html=True)
    st.stop()

# Defensive cleanup for any persisted/stale saved snapshot: drop excluded
# accounts/outlets (Bestmart, MUHINDI MWEUSI WITEITHIE) that may still be baked
# into an older saved file, while leaving 99MART in the analysis universe.
_excl_pattern = '|'.join(KE_EXCLUDE_ACCOUNT_KEYWORDS)
if 'CUSTOMER NAME' in df.columns:
    customer_norm = df['CUSTOMER NAME'].astype(str).str.strip().str.lower()
    df = df[~customer_norm.str.contains(_excl_pattern, regex=True, na=False)].copy()
if 'ACCOUNT' in df.columns:
    acct_norm = df['ACCOUNT'].astype(str).str.strip().str.lower()
    df = df[~acct_norm.str.contains(_excl_pattern, regex=True, na=False)].copy()

if df.empty:
    st.warning("No rows left after filtering — check that the MHSKU reference matches this file's PRODUCT CODEs.")
    st.stop()

# ── banners ───────────────────────────────────────────────────────────────────
kampala_note = ("🇰🇪 <strong>Kenya regions only</strong> — Kampala / Uganda rows and Bestmart accounts dropped automatically. "
                f"<strong>{len(df):,}</strong> rows loaded.")
st.markdown(
    "<div style='font-size:12px;color:#34c97b;background:rgba(52,201,123,.08);"
    "border:1px solid rgba(52,201,123,.2);border-radius:8px;padding:8px 14px;margin-bottom:8px'>"
    f"{kampala_note}</div>", unsafe_allow_html=True,
)
if 'MUST HAVE' in df.columns and mhsku_map:
    n_mh = int(df['MUST HAVE'].sum()) if not must_have_only else len(df)
    st.markdown(
        "<div style='font-size:12px;color:#4f8ef7;background:rgba(79,142,247,.08);"
        "border:1px solid rgba(79,142,247,.2);border-radius:8px;padding:6px 14px;margin-bottom:8px'>"
        f"📋 <strong>MHSKU / MBQ scoring active</strong> — pressure target = MBQ for the outlet's category "
        f"({', '.join(KE_MHSKU_SHEET_MAP.values())}).</div>", unsafe_allow_html=True,
    )

# ── filter lists ──────────────────────────────────────────────────────────────
months  = sorted(df['Month'].dropna().unique(), key=lambda m: pd.to_datetime(m, format='%B').month)
regions = sorted(df['REGION'].dropna().unique()) if 'REGION' in df.columns else []
outlet_cats = sorted(df['MHSKU_SHEET'].dropna().unique()) if 'MHSKU_SHEET' in df.columns else []
brands  = sorted(df['BRAND NAME'].dropna().unique()) if 'BRAND NAME' in df.columns else []
accts   = sorted(df['ACCOUNT'].dropna().unique()) if 'ACCOUNT' in df.columns else []

with st.expander("🔍  Filters", expanded=True):
    fc1, fc2, fc3 = st.columns(3)
    sel_months  = synced_multiselect(fc1, "Month", months, "ke_osa_months")
    sel_regions = synced_multiselect(fc2, "Region", regions, "ke_osa_regions") if regions else []
    sel_ocats   = synced_multiselect(fc3, "Outlet Category", outlet_cats, "ke_osa_ocats") if outlet_cats else []
    fc4, fc5 = st.columns(2)
    sel_brands = synced_multiselect(fc4, "Brand", brands, "ke_osa_brands") if brands else []
    sel_accts  = synced_multiselect(fc5, "Key Account", accts, "ke_osa_accts") if accts else []

# Shared with the Report Generator — same session_state keys, same logic, so a
# selection here carries straight through to the downloaded Excel report.
dff = apply_ke_osa_filters(df, sel_months, sel_regions, sel_ocats, sel_brands, sel_accts)

if dff.empty:
    st.warning("No data matches the selected filters.")
    st.stop()

th = T()
pal = PALETTE(th)

# ── KPIs ──────────────────────────────────────────────────────────────────────
overall_osa = dff['OSA'].mean()
kc = osa_color(overall_osa)
st.markdown(f"""
<div class="metric-row">
  <div class="metric-card">
    <div class="m-label">Overall OSA%</div>
    <div class="m-value" style="color:{kc}">{overall_osa:.1f}%</div>
    <div class="m-sub">mean across selection</div>
  </div>
  <div class="metric-card">
    <div class="m-label">On Target (100%)</div>
    <div class="m-value" style="color:{th['green']}">{(dff['OSA']==100).mean()*100:.1f}%</div>
    <div class="m-sub">of records</div>
  </div>
  <div class="metric-card">
    <div class="m-label">Outlet Categories</div>
    <div class="m-value">{dff['MHSKU_SHEET'].nunique() if 'MHSKU_SHEET' in dff.columns else '—'}</div>
    <div class="m-sub">in selection</div>
  </div>
  <div class="metric-card">
    <div class="m-label">Accounts</div>
    <div class="m-value">{dff['ACCOUNT'].nunique() if 'ACCOUNT' in dff.columns else '—'}</div>
    <div class="m-sub">key groups</div>
  </div>
  <div class="metric-card">
    <div class="m-label">Must-Have SKUs</div>
    <div class="m-value">{dff['PRODUCT CODE'].nunique() if 'PRODUCT CODE' in dff.columns else '—'}</div>
    <div class="m-sub">unique products</div>
  </div>
</div>
""", unsafe_allow_html=True)
st.markdown("---")

# ── Row 1: OSA% by Account ───────────────────────────────────────────────────
acct_osa_bar = (dff.groupby('ACCOUNT', observed=True)['OSA'].mean().round(1)
                   .sort_values().reset_index())
acct_osa_bar.columns = ['Account', 'OSA']
fig1 = go.Figure(go.Bar(
    x=acct_osa_bar['OSA'], y=acct_osa_bar['Account'], orientation='h',
    marker_color=color_bars(acct_osa_bar['OSA']),
    text=acct_osa_bar['OSA'].map(lambda v: f"{v:.1f}%"),
    textposition='outside', textfont=dict(size=10, color=th['text']),
))
fig1.add_vline(x=95, line_dash="dot", line_color=th['green'],
               annotation_text="95% target", annotation_font_color=th['green'], annotation_font_size=10)
apply_layout(fig1, "OSA% by Account (MBQ target)", height=max(320, len(acct_osa_bar)*30))
fig1.update_xaxes(range=[0, 115], ticksuffix='%')
st.plotly_chart(fig1, use_container_width=True)

# ── Account branch deep-dive (modal popup, triggered from the cards below) ────
@st.dialog("🔎 Account Branch Deep-Dive", width="large")
def account_deep_dive(acct: str):
    thc = T()
    pal_d = PALETTE(thc)
    a_all = dff[dff['ACCOUNT'] == acct].copy()
    a_gap = a_all[a_all['OSA'] == 0].copy()
    if 'REASON' in a_gap.columns:
        a_gap['REASON'] = a_gap['REASON'].astype(str).str.strip()
        a_gap = a_gap[a_gap['REASON'].isin(KE_OOS_REASONS)]

    acct_osa = a_all['OSA'].mean()
    n_branches = a_all['CUSTOMER NAME'].nunique() if 'CUSTOMER NAME' in a_all.columns else 0
    n_gap = len(a_gap)
    n_reps = a_gap['REP NAME'].nunique() if 'REP NAME' in a_gap.columns else 0
    dom_reason = a_gap['REASON'].value_counts().index[0] if n_gap else '—'
    dom_label = KE_OOS_REASON_GUIDANCE.get(dom_reason, {}).get('label', dom_reason)

    st.markdown(f"## {acct}")
    st.markdown(
        f"<span style='color:{osa_color(acct_osa)};font-weight:700;font-size:15px'>{acct_osa:.1f}% OSA</span>"
        f"&nbsp;·&nbsp;dominant challenge: <strong style='color:{thc['amber']}'>{dom_label}</strong>",
        unsafe_allow_html=True,
    )

    # per-branch OSA + region
    branch_osa = (a_all.groupby('CUSTOMER NAME', observed=True)
                       .agg(OSA=('OSA', 'mean'), Rows=('OSA', 'size')).reset_index())
    branch_osa['OSA'] = branch_osa['OSA'].round(1)
    weakest = branch_osa.sort_values('OSA').iloc[0] if len(branch_osa) else None

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Branches", f"{n_branches:,}")
    m2.metric("OOS gap rows", f"{n_gap:,}")
    m3.metric("Reps to follow up", f"{n_reps:,}")
    if weakest is not None:
        m4.metric("Weakest branch", f"{weakest['OSA']:.0f}%", help=str(weakest['CUSTOMER NAME']))

    if n_gap == 0:
        st.success("No recorded OOS gaps for this account on the current selection — nothing to chase. 🎉")
        return

    # ── build per-branch follow-up detail ────────────────────────────────────
    rows = []
    for br, bdf in a_gap.groupby('CUSTOMER NAME', observed=True):
        rc = bdf.groupby('REASON', observed=True).size().sort_values(ascending=False)
        top_reason = KE_OOS_REASON_GUIDANCE.get(rc.index[0], {}).get('label', rc.index[0]) if len(rc) else ''
        region = (bdf['REGION'].dropna().mode().iloc[0]
                  if 'REGION' in bdf.columns and bdf['REGION'].notna().any() else '')
        if 'REP NAME' in bdf.columns and bdf['REP NAME'].notna().any():
            rep = bdf.groupby('REP NAME', observed=True).size().sort_values(ascending=False).index[0]
        else:
            rep = ''
        boa = branch_osa.loc[branch_osa['CUSTOMER NAME'] == br, 'OSA']
        rows.append({
            'Branch': str(br), 'Region': str(region),
            'OSA %': float(boa.iloc[0]) if len(boa) else None,
            'Gap rows': int(len(bdf)), 'Top reason': top_reason, 'Filled by (rep)': str(rep),
        })
    follow = pd.DataFrame(rows).sort_values('Gap rows', ascending=False).reset_index(drop=True)

    dc1, dc2 = st.columns(2)
    # top branches by gap volume, colored by that branch's dominant reason
    with dc1:
        topb = follow.head(10).sort_values('Gap rows')
        figb = go.Figure(go.Bar(
            x=topb['Gap rows'], y=topb['Branch'], orientation='h',
            marker_color=thc['amber'],
            customdata=topb[['Region', 'Top reason']].values,
            text=topb['Gap rows'].map(lambda v: f"{v:,}"),
            textposition='outside', textfont=dict(size=9, color=thc['text']),
            hovertemplate="<b>%{y}</b><br>%{x:,} gap rows<br>%{customdata[0]} · %{customdata[1]}<extra></extra>",
        ))
        figb.update_layout(**get_plotly_layout(height=max(280, len(topb)*28), title="Worst branches by OOS gap rows"))
        figb.update_xaxes(tickfont_color=thc['text']); figb.update_yaxes(tickfont=dict(size=9, color=thc['text']))
        st.plotly_chart(figb, use_container_width=True)

    # reason mix for the account
    with dc2:
        rmix = a_gap.groupby('REASON', observed=True).size().sort_values().reset_index(name='Count')
        rmix['Label'] = rmix['REASON'].map(lambda r: KE_OOS_REASON_GUIDANCE.get(r, {}).get('label', r))
        figr = go.Figure(go.Bar(
            x=rmix['Count'], y=rmix['Label'], orientation='h',
            marker_color=[thc['red'] if r == dom_reason else thc['blue'] for r in rmix['REASON']],
            text=rmix['Count'].map(lambda v: f"{v:,}"),
            textposition='outside', textfont=dict(size=9, color=thc['text']),
        ))
        figr.update_layout(**get_plotly_layout(height=max(280, len(rmix)*30), title="Why — reason mix"))
        figr.update_xaxes(tickfont_color=thc['text']); figr.update_yaxes(tickfont=dict(size=9, color=thc['text']))
        st.plotly_chart(figr, use_container_width=True)

    # ── affected assortment — Category / Brand / SKU breakdown ───────────────
    # combined SKU label (code — description) for the SKU views
    if 'PRODUCT CODE' in a_gap.columns and 'DESCRIPTION' in a_gap.columns:
        a_gap['_SKU'] = (a_gap['PRODUCT CODE'].astype(str).str.strip()
                         + ' — ' + a_gap['DESCRIPTION'].astype(str).str.strip())
    elif 'DESCRIPTION' in a_gap.columns:
        a_gap['_SKU'] = a_gap['DESCRIPTION'].astype(str).str.strip()
    elif 'PRODUCT CODE' in a_gap.columns:
        a_gap['_SKU'] = a_gap['PRODUCT CODE'].astype(str).str.strip()

    def _affected_bar(col, title, color, top=10):
        if col not in a_gap.columns:
            return
        d = (a_gap.groupby(col, observed=True).size()
                  .sort_values(ascending=False).head(top).sort_values().reset_index(name='Gap rows'))
        fig = go.Figure(go.Bar(
            x=d['Gap rows'], y=d[col].astype(str), orientation='h', marker_color=color,
            text=d['Gap rows'].map(lambda v: f"{v:,}"),
            textposition='outside', textfont=dict(size=9, color=thc['text']),
        ))
        fig.update_layout(**get_plotly_layout(height=max(240, len(d)*26), title=title))
        fig.update_xaxes(tickfont_color=thc['text']); fig.update_yaxes(tickfont=dict(size=9, color=thc['text']))
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("<div style='font-size:12px;color:var(--muted);margin:8px 0 2px'>"
                "Affected assortment — what's out of stock (by gap rows)</div>", unsafe_allow_html=True)
    ac1, ac2 = st.columns(2)
    with ac1:
        _affected_bar('PRODUCT CATEGORY', 'Product categories most affected', thc['purple'])
    with ac2:
        _affected_bar('BRAND NAME', 'Brands most affected', thc['cyan'])
    _affected_bar('_SKU', 'Top affected SKUs', thc['amber'], top=15)

    st.markdown("<div style='font-size:12px;color:var(--muted);margin:6px 0 4px'>"
                "Branch follow-up list — who filled the OSA & where to chase</div>", unsafe_allow_html=True)
    st.dataframe(follow, use_container_width=True, hide_index=True,
                 column_config={"OSA %": st.column_config.NumberColumn(format="%.0f%%")})
    st.download_button(
        f"⬇️ {acct} branch follow-up (CSV)",
        data=follow.to_csv(index=False).encode('utf-8'),
        file_name=f"{acct}_branch_followup.csv", mime="text/csv", use_container_width=True,
    )

    # ── SKU-level worklist: affected category × brand × SKU per branch ────────
    wcols = [c for c in ['CUSTOMER NAME', 'REGION', 'PRODUCT CATEGORY', 'BRAND NAME',
                         '_SKU', 'REASON', 'REP NAME'] if c in a_gap.columns]
    if wcols:
        work = (a_gap.groupby(wcols, observed=True).size().reset_index(name='Gap rows')
                     .sort_values('Gap rows', ascending=False))
        work = work.rename(columns={'CUSTOMER NAME': 'Branch', 'REGION': 'Region',
                                    'PRODUCT CATEGORY': 'Category', 'BRAND NAME': 'Brand',
                                    '_SKU': 'SKU', 'REASON': 'Reason', 'REP NAME': 'Filled by (rep)'})
        with st.expander(f"📋 SKU-level worklist — {len(work):,} rows (branch × category × brand × SKU)"):
            st.dataframe(work, use_container_width=True, hide_index=True)
            st.download_button(
                f"⬇️ {acct} SKU worklist (CSV)",
                data=work.to_csv(index=False).encode('utf-8'),
                file_name=f"{acct}_sku_worklist.csv", mime="text/csv", use_container_width=True,
            )


# ── Row 1.5: Why OSA Is Affected — root cause (REASON) ───────────────────────
st.markdown("---")
st.markdown('<div class="sec-label">Why OSA Is Affected — Root Cause</div>', unsafe_allow_html=True)

if 'REASON' in dff.columns and 'STOCK LEVEL' in dff.columns:
    gap = dff[dff['OSA'] == 0].copy()
    gap['REASON'] = gap['REASON'].astype(str).str.strip()
    gap = gap[gap['REASON'].isin(KE_OOS_REASONS)]
    gap = gap[gap['ACCOUNT'].notna()].copy() if 'ACCOUNT' in gap.columns else gap

    if gap.empty:
        st.info("No OOS-reason data on the current selection — either everything is on target, "
                "or the file doesn't carry a REASON column for the gaps.")
    else:
        n_total_gaps = int((dff['OSA'] == 0).sum())
        n_explained = len(gap)
        reason_counts = gap['REASON'].value_counts()
        top_reason, top_n = reason_counts.index[0], reason_counts.iloc[0]
        top_meta = KE_OOS_REASON_GUIDANCE.get(top_reason, {})
        top_label = top_meta.get('label', top_reason)
        top_challenge = top_meta.get('challenge', 'Unknown challenge area')
        top_hint = top_meta.get('hint', 'No extra guidance available.')

        st.markdown(f"""
<div style="font-size:13px;color:var(--muted);background:var(--surface);border:1px solid var(--border);
            border-radius:10px;padding:10px 16px;margin-bottom:12px">
  Of <strong style="color:var(--text)">{n_total_gaps:,}</strong> must-have SKU × outlet gaps
  (quantity below MBQ), <strong style="color:var(--text)">{n_explained:,}</strong>
  ({n_explained / max(n_total_gaps,1) * 100:.0f}%) carry a recorded reason.
  <br><strong style="color:{th['amber']}">{top_label}</strong> is the leading challenge
  ({top_n:,} rows, {top_n / max(n_explained,1) * 100:.0f}% of explained gaps).
  <br><strong style="color:var(--text)">{top_challenge}</strong> — {top_hint}
</div>
""", unsafe_allow_html=True)

        reason_by_account = (gap.groupby(['ACCOUNT', 'REASON']).size().reset_index(name='Count'))
        account_totals = (reason_by_account.groupby('ACCOUNT', as_index=False)['Count'].sum()
                              .sort_values('Count', ascending=False)
                              .head(5))
        account_totals.columns = ['Account', 'Gap Rows']
        account_branches = gap.groupby(['ACCOUNT', 'CUSTOMER NAME']).size().reset_index(name='Branch Rows') if 'CUSTOMER NAME' in gap.columns else None

        st.markdown("<div style='font-size:12px;color:var(--muted);margin-bottom:10px'>"
                    "Top 5 accounts by affected branches — open a card for the full branch deep-dive</div>",
                    unsafe_allow_html=True)
        for _, account_row in account_totals.iterrows():
            acct = account_row['Account']
            acct_reason = reason_by_account[reason_by_account['ACCOUNT'] == acct].sort_values('Count', ascending=False).iloc[0]
            acct_reason_label = KE_OOS_REASON_GUIDANCE.get(acct_reason['REASON'], {}).get('label', acct_reason['REASON'])
            branch_text = ""
            if account_branches is not None:
                top_branches = (account_branches[account_branches['ACCOUNT'] == acct]
                                  .sort_values('Branch Rows', ascending=False)
                                  .head(5))
                top_branch_list = [f"{br} ({int(cnt):,})" for br, cnt in zip(top_branches['CUSTOMER NAME'], top_branches['Branch Rows'])]
                branch_text = f"<br><span style='color:var(--muted)'>Top branches: {', '.join(top_branch_list)}</span>"
            st.markdown(f"""
<div style="font-size:12px;color:var(--text);background:rgba(79,142,247,.06);
            border:1px solid rgba(79,142,247,.2);border-bottom:none;
            border-radius:8px 8px 0 0;padding:8px 10px 6px;margin:0">
  <strong>{acct}</strong> — <strong style="color:{th['amber']}">{acct_reason_label}</strong> ({int(acct_reason['Count']):,} rows)<br>
  <span style="color:var(--muted)">Total affected gaps: {int(account_row['Gap Rows']):,}</span>{branch_text}
</div>
""", unsafe_allow_html=True)
            if st.button(f"🔍  Branch deep-dive — {acct}", key=f"dd_{acct}", use_container_width=True):
                account_deep_dive(acct)
            st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

        rc1, rc2 = st.columns(2)
        with rc1:
            rc = reason_counts.reset_index()
            rc.columns = ['Reason', 'Count']
            rc['Reason Label'] = rc['Reason'].map(lambda r: KE_OOS_REASON_GUIDANCE.get(r, {}).get('label', r))
            rc = rc.sort_values('Count')
            fig_reason = go.Figure(go.Bar(
                x=rc['Count'], y=rc['Reason Label'], orientation='h',
                marker_color=[th['red'] if r == top_reason else th['amber'] for r in rc['Reason']],
                text=rc['Count'].map(lambda v: f"{v:,}"),
                textposition='outside', textfont=dict(size=10, color=th['text']),
            ))
            apply_layout(fig_reason, "OOS Reason — All Gaps", height=max(280, len(rc) * 40))
            st.plotly_chart(fig_reason, use_container_width=True)

        with rc2:
            top_acc = (reason_by_account.groupby('ACCOUNT', observed=True)['Count'].sum()
                          .sort_values(ascending=False).head(10).index.tolist())
            rba = reason_by_account[reason_by_account['ACCOUNT'].isin(top_acc)]
            acc_order = (rba.groupby('ACCOUNT', observed=True)['Count'].sum()
                            .sort_values().index.tolist())
            fig_reason_acc = go.Figure()
            for i, reason in enumerate(reason_counts.index):
                sub = (rba[rba['REASON'] == reason].set_index('ACCOUNT')
                          .reindex(acc_order)['Count'].fillna(0))
                fig_reason_acc.add_trace(go.Bar(
                    name=KE_OOS_REASON_GUIDANCE.get(reason, {}).get('label', reason),
                    y=list(acc_order), x=list(sub.values), orientation='h',
                    marker_color=pal[i % len(pal)],
                ))
            fig_reason_acc.update_layout(barmode='stack')
            apply_layout(fig_reason_acc, "OOS Reason — by Account (top 10)", height=max(340, len(acc_order)*30))
            st.plotly_chart(fig_reason_acc, use_container_width=True)
else:
    st.info("This file doesn't have a REASON / STOCK LEVEL column, so a root-cause breakdown isn't available.")

# ── Row 2: Weekly trend — Outlet Category ────────────────────────────────────
st.markdown("---")
st.markdown('<div class="sec-label">Weekly Trend — by Outlet Category</div>', unsafe_allow_html=True)
week_order = [f'Week {i}' for i in range(1, 7)]
week_oc = (dff.groupby(['WEEK_LABEL','MHSKU_SHEET'])['OSA'].mean().round(1).reset_index())
week_oc.columns = ['Week','Outlet Category','OSA']
wo = [w for w in week_order if w in week_oc['Week'].values]
fig3 = go.Figure()
for i, oc in enumerate(sorted(week_oc['Outlet Category'].unique())):
    sub = week_oc[week_oc['Outlet Category']==oc].set_index('Week').reindex(wo).reset_index()
    fig3.add_trace(go.Scatter(
        x=sub['Week'], y=sub['OSA'], name=oc, mode='lines+markers',
        line=dict(color=pal[i % len(pal)], width=2), marker=dict(size=7),
        hovertemplate=f"<b>{oc}</b><br>%{{x}}: %{{y:.1f}}%<extra></extra>",
    ))
fig3.add_hline(y=95, line_dash="dot", line_color=th['green'], annotation_text="95%", annotation_font_color=th['green'])
fig3.add_hrect(y0=80, y1=95, fillcolor=th['amber'], opacity=0.05, line_width=0)
apply_layout(fig3, "", height=360)
fig3.update_yaxes(range=[0, 110], ticksuffix='%')
st.plotly_chart(fig3, use_container_width=True)

# ── Row 3: Heatmap — Account × Outlet Category ───────────────────────────────
st.markdown("---")
colorscale = [[0.0,'#f87171'],[0.8,'#fbbf24'],[0.95,'#34c97b'],[1.0,'#34c97b']]

def _heatmap(pivot, title):
    fig = go.Figure(go.Heatmap(
        z=pivot.values, x=list(pivot.columns), y=list(pivot.index),
        colorscale=colorscale, zmin=0, zmax=100,
        text=np.where(np.isnan(pivot.values), '', np.round(pivot.values, 0).astype(int).astype(str) + '%'),
        texttemplate='%{text}', textfont=dict(size=9, color='#0f1117'),
        hovertemplate='<b>%{y}</b> × <b>%{x}</b><br>OSA: %{z:.1f}%<extra></extra>',
        showscale=True,
        colorbar=dict(ticksuffix='%', tickfont=dict(color=T()['text']), outlinecolor=T()['border'], outlinewidth=1),
    ))
    apply_layout(fig, title, height=max(380, len(pivot)*26))
    fig.update_xaxes(tickangle=-35, tickfont=dict(size=9))
    fig.update_yaxes(tickfont=dict(size=9))
    return fig

st.markdown('<div class="sec-label">Account × Outlet Category Heatmap</div>', unsafe_allow_html=True)
heat = (dff.groupby(['ACCOUNT','MHSKU_SHEET'], observed=True)['OSA'].mean().round(1).reset_index()
           .pivot(index='ACCOUNT', columns='MHSKU_SHEET', values='OSA'))
st.plotly_chart(_heatmap(heat, "Account × Outlet Category OSA (%)"), use_container_width=True)

# ── Row 3.5: Account × Product Category  &  Account × Brand ───────────────────
st.markdown("---")
st.markdown('<div class="sec-label">OSA by Account × Product Category & Brand</div>', unsafe_allow_html=True)

def _worst_combos(data, dim, min_records=20, n=6):
    g = (data.groupby(['ACCOUNT', dim], observed=True)
             .agg(OSA=('OSA', 'mean'), Records=('OSA', 'size')).reset_index())
    g = g[g['Records'] >= min_records].copy()
    g['OSA'] = g['OSA'].round(1)
    return g.sort_values('OSA').head(n)

def _combo_chips(worst, dim, caption):
    if worst.empty:
        return
    items = "".join(
        f"<div style='font-size:12px;color:var(--text);background:rgba(248,113,113,.08);"
        f"border:1px solid rgba(248,113,113,.25);border-radius:8px;padding:6px 10px;margin-bottom:5px'>"
        f"<strong>{row['ACCOUNT']}</strong> × {row[dim]} — "
        f"<strong style='color:{th['red']}'>{row['OSA']:.0f}%</strong> "
        f"<span style='color:var(--muted)'>({int(row['Records']):,} records)</span></div>"
        for _, row in worst.iterrows()
    )
    st.markdown(f"<div style='font-size:12px;color:var(--muted);margin:4px 0 6px'>{caption}</div>{items}",
                unsafe_allow_html=True)

# Account × Product Category
if 'PRODUCT CATEGORY' in dff.columns:
    st.markdown("<div style='font-size:13px;color:var(--text);font-weight:600;margin:6px 0'>Account × Product Category</div>", unsafe_allow_html=True)
    hpc1, hpc2 = st.columns([3, 2])
    with hpc1:
        heat_pc = (dff.groupby(['ACCOUNT', 'PRODUCT CATEGORY'], observed=True)['OSA'].mean().round(1).reset_index()
                      .pivot(index='ACCOUNT', columns='PRODUCT CATEGORY', values='OSA'))
        fig_pc = _heatmap(heat_pc, "Account × Product Category OSA (%)")
        fig_pc.update_xaxes(tickfont=dict(size=8))
        st.plotly_chart(fig_pc, use_container_width=True)
    with hpc2:
        _combo_chips(_worst_combos(dff, 'PRODUCT CATEGORY'), 'PRODUCT CATEGORY',
                     "Weakest account × product-category combos (≥20 records)")
    with st.expander("📋 Full Account × Product Category table"):
        st.dataframe(heat_pc.round(1), use_container_width=True)

# Account × Brand
if 'BRAND NAME' in dff.columns:
    st.markdown("<div style='font-size:13px;color:var(--text);font-weight:600;margin:12px 0 6px'>Account × Brand</div>", unsafe_allow_html=True)
    hb1, hb2 = st.columns([3, 2])
    with hb1:
        heat_br = (dff.groupby(['ACCOUNT', 'BRAND NAME'], observed=True)['OSA'].mean().round(1).reset_index()
                      .pivot(index='ACCOUNT', columns='BRAND NAME', values='OSA'))
        fig_br = _heatmap(heat_br, "Account × Brand OSA (%)")
        fig_br.update_xaxes(tickfont=dict(size=8))
        st.plotly_chart(fig_br, use_container_width=True)
    with hb2:
        _combo_chips(_worst_combos(dff, 'BRAND NAME'), 'BRAND NAME',
                     "Weakest account × brand combos (≥20 records)")
    with st.expander("📋 Full Account × Brand table"):
        st.dataframe(heat_br.round(1), use_container_width=True)

# ── Row 4: Top/Bottom accounts ────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="sec-label">Account Performance Ranking</div>', unsafe_allow_html=True)
acct_osa = dff.groupby('ACCOUNT')['OSA'].mean().round(1).reset_index()
acct_osa.columns = ['Account','OSA']
acct_osa['is_99mart'] = acct_osa['Account'].astype(str).str.upper().str.contains('99MART', na=False)
acct_osa = acct_osa.sort_values(['is_99mart', 'OSA'], ascending=[True, False]).reset_index(drop=True)

ra1, ra2 = st.columns(2)
with ra1:
    top10 = acct_osa.head(10).sort_values('OSA', ascending=True)
    fig7 = go.Figure(go.Bar(
        x=top10['OSA'], y=top10['Account'], orientation='h',
        marker_color=[th['green']]*len(top10),
        text=top10['OSA'].map(lambda v: f"{v:.1f}%"),
        textposition='outside', textfont=dict(size=10, color=th['text']),
    ))
    apply_layout(fig7, "🏆 Top 10 Accounts", height=340)
    fig7.update_xaxes(range=[0, 115], ticksuffix='%')
    st.plotly_chart(fig7, use_container_width=True)

with ra2:
    bot10 = acct_osa.tail(10).sort_values('OSA', ascending=True)
    fig8 = go.Figure(go.Bar(
        x=bot10['OSA'], y=bot10['Account'], orientation='h',
        marker_color=color_bars(bot10['OSA']),
        text=bot10['OSA'].map(lambda v: f"{v:.1f}%"),
        textposition='outside', textfont=dict(size=10, color=th['text']),
    ))
    apply_layout(fig8, "⚠️ Bottom 10 Accounts", height=340)
    fig8.update_xaxes(range=[0, 115], ticksuffix='%')
    st.plotly_chart(fig8, use_container_width=True)

# ── Row 5: SKU-wise OSA ───────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="sec-label">OSA% by Must-Have SKU</div>', unsafe_allow_html=True)

sku_group_cols = [c for c in ['PRODUCT CODE', 'DESCRIPTION'] if c in dff.columns]
if not sku_group_cols:
    st.info("No SKU / description column found in this file.")
else:
    sku_ctrl1, sku_ctrl2 = st.columns([2, 1])
    with sku_ctrl1:
        sku_search = st.text_input("🔎 Search SKU (code or description)", value="", key="ke_osa_sku_search",
                                    placeholder="Type part of a code or description to filter…")
    with sku_ctrl2:
        sku_n = st.number_input("Show top/bottom N", min_value=5, max_value=100, value=20, step=5, key="ke_osa_sku_n")

    has_code = 'PRODUCT CODE' in dff.columns
    has_desc = 'DESCRIPTION' in dff.columns
    sku_osa = (dff.groupby(sku_group_cols).agg(OSA=('OSA', 'mean'), Records=('OSA', 'size')).reset_index())
    sku_osa['OSA'] = sku_osa['OSA'].round(1)

    if has_code and has_desc:
        sku_osa['SKU'] = (sku_osa['PRODUCT CODE'].astype(str).str.strip() + ' — ' + sku_osa['DESCRIPTION'].astype(str).str.strip())
        sku_osa = sku_osa.rename(columns={'PRODUCT CODE': 'Code', 'DESCRIPTION': 'Description'})
        sku_osa = sku_osa[['SKU', 'Code', 'Description', 'OSA', 'Records']]
    else:
        sku_osa = sku_osa.rename(columns={sku_group_cols[0]: 'SKU'})

    if sku_search:
        search_mask = sku_osa['SKU'].str.contains(sku_search, case=False, na=False)
        if has_code and has_desc:
            search_mask = search_mask | sku_osa['Code'].str.contains(sku_search, case=False, na=False)
        sku_osa = sku_osa[search_mask]

    if sku_osa.empty:
        st.warning("No SKUs match that search.")
    else:
        sku_osa = sku_osa.sort_values('OSA', ascending=False)
        sk1, sk2 = st.columns(2)
        with sk1:
            top_sku = sku_osa.head(int(sku_n)).sort_values('OSA')
            fig_sku_top = go.Figure(go.Bar(
                x=top_sku['OSA'], y=top_sku['SKU'], orientation='h',
                marker_color=[th['green']] * len(top_sku),
                text=top_sku['OSA'].map(lambda v: f"{v:.1f}%"),
                textposition='outside', textfont=dict(size=9, color=th['text']),
            ))
            apply_layout(fig_sku_top, f"🏆 Top {len(top_sku)} SKUs by OSA%", height=max(340, len(top_sku) * 24))
            fig_sku_top.update_xaxes(range=[0, 115], ticksuffix='%')
            fig_sku_top.update_yaxes(tickfont=dict(size=9))
            st.plotly_chart(fig_sku_top, use_container_width=True)
        with sk2:
            bottom_sku = sku_osa.tail(int(sku_n)).sort_values('OSA')
            fig_sku_bot = go.Figure(go.Bar(
                x=bottom_sku['OSA'], y=bottom_sku['SKU'], orientation='h',
                marker_color=color_bars(bottom_sku['OSA']),
                text=bottom_sku['OSA'].map(lambda v: f"{v:.1f}%"),
                textposition='outside', textfont=dict(size=9, color=th['text']),
            ))
            apply_layout(fig_sku_bot, f"⚠️ Bottom {len(bottom_sku)} SKUs by OSA%", height=max(340, len(bottom_sku) * 24))
            fig_sku_bot.update_xaxes(range=[0, 115], ticksuffix='%')
            fig_sku_bot.update_yaxes(tickfont=dict(size=9))
            st.plotly_chart(fig_sku_bot, use_container_width=True)

        with st.expander(f"📋 Full SKU table ({len(sku_osa):,} SKUs)"):
            st.dataframe(sku_osa.rename(columns={'OSA': 'OSA %'}), use_container_width=True, hide_index=True)

# ── Export ────────────────────────────────────────────────────────────────────
st.markdown("---")
csv_bytes = dff.drop(columns=['DATE_PARSED'], errors='ignore').to_csv(index=False).encode('utf-8')
st.download_button("⬇️ Download filtered data (CSV)", data=csv_bytes,
                   file_name="kenya_osa_filtered.csv", mime="text/csv")

st.markdown('<div class="footer">Kenya OSA Analytics — Field Data scored against MHSKU / MBQ targets</div>', unsafe_allow_html=True)
