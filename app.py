"""
app.py  —  Creative Audit Report Generator
Streamlit web front-end for the Captivate Insights PPTX pipeline.
"""

import io
import gc
import logging
from pathlib import Path

import streamlit as st

st.set_page_config(
    page_title="Creative Audit | Report Generator",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

import report_engine as engine

log = logging.getLogger(__name__)

st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }
[data-testid="stSidebar"] { background: #F7F5FF; border-right: 1px solid #E0E0E0; }
[data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,[data-testid="stSidebar"] h3 { color: #3C2A8C; }
[data-testid="stSidebar"] label { font-weight: 600; font-size: 0.78rem; color: #3C2A8C; text-transform: uppercase; letter-spacing: 0.05em; }
.captivate-header { display: flex; align-items: center; gap: 14px; padding: 18px 0 10px 0; border-bottom: 2px solid #3C2A8C; margin-bottom: 24px; }
.captivate-header .logo-ring { width: 44px; height: 44px; border-radius: 50%; background: conic-gradient(#3C2A8C 0deg 180deg, #E08C00 180deg 360deg); display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
.captivate-header .logo-inner { width: 28px; height: 28px; border-radius: 50%; background: white; display: flex; align-items: center; justify-content: center; }
.captivate-header .logo-core { width: 14px; height: 14px; border-radius: 50%; background: #3C2A8C; }
.captivate-header h1 { margin: 0; font-size: 1.5rem; color: #3C2A8C; font-weight: 700; }
.captivate-header p { margin: 2px 0 0 0; font-size: 0.85rem; color: #888; }
.stat-row { display: flex; gap: 12px; margin-bottom: 20px; }
.stat-card { flex: 1; background: white; border: 1px solid #E0E0E0; border-left: 4px solid #3C2A8C; border-radius: 6px; padding: 14px 16px; }
.stat-card .stat-label { font-size: 0.72rem; font-weight: 600; color: #888; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 4px; }
.stat-card .stat-value { font-size: 1.6rem; font-weight: 700; color: #3C2A8C; line-height: 1; }
.stat-card .stat-sub { font-size: 0.78rem; color: #888; margin-top: 3px; }
.stat-teal { border-left-color: #009999 !important; }
.stat-teal .stat-value { color: #009999; }
.stat-red { border-left-color: #C00000 !important; }
.stat-red .stat-value { color: #C00000; }
.stat-amber { border-left-color: #E08C00 !important; }
.stat-amber .stat-value { color: #E08C00; }
.filter-section { font-size: 0.7rem; font-weight: 700; color: #3C2A8C; text-transform: uppercase; letter-spacing: 0.06em; padding: 8px 0 4px 0; border-top: 1px solid #E0E0E0; margin-top: 6px; }
div[data-testid="stButton"] > button { background: #3C2A8C; color: white; border: none; border-radius: 6px; font-weight: 600; font-size: 0.95rem; padding: 10px 0; width: 100%; transition: background 0.15s; }
div[data-testid="stButton"] > button:hover { background: #2A1B6E; }
div[data-testid="stDownloadButton"] > button { background: #009999; color: white; border: none; border-radius: 6px; font-weight: 600; font-size: 0.95rem; padding: 10px 0; width: 100%; }
div[data-testid="stDownloadButton"] > button:hover { background: #007777; }
.preview-table { width: 100%; border-collapse: collapse; font-size: 0.82rem; }
.preview-table th { background: #3C2A8C; color: white; padding: 8px 10px; text-align: left; font-weight: 600; font-size: 0.75rem; text-transform: uppercase; letter-spacing: 0.04em; }
.preview-table td { padding: 7px 10px; border-bottom: 1px solid #F0F0F0; color: #1A1A1A; }
.preview-table tr:hover td { background: #F7F5FF; }
.badge { display: inline-block; padding: 2px 8px; border-radius: 20px; font-size: 0.72rem; font-weight: 700; }
.badge-unmissable { background: #E6F7F7; color: #009999; }
.badge-wallpaper { background: #FFF4E0; color: #E08C00; }
.badge-missable { background: #FDEAEA; color: #C00000; }
.stProgress > div > div { background: #3C2A8C; }
[data-testid="stExpander"] summary { font-weight: 600; color: #3C2A8C; }
</style>
""", unsafe_allow_html=True)


def score_badge(score_str):
    try:
        s = float(score_str.replace("%", ""))
    except Exception:
        return score_str
    if s >= 85:
        return f'<span class="badge badge-unmissable">{score_str} UNMISSABLE</span>'
    if s >= 55:
        return f'<span class="badge badge-wallpaper">{score_str} WALLPAPER</span>'
    return f'<span class="badge badge-missable">{score_str} MISSABLE</span>'


def month_sort_key(m):
    try:
        return int(m.strip().split(".")[0].strip())
    except (ValueError, IndexError):
        pass
    m_lower = m.lower()
    for idx, short in enumerate(engine.MONTH_ORDER):
        if short.lower() in m_lower:
            return idx
    return 99


@st.cache_resource(show_spinner="Connecting to Smartsheet…")
def get_sheet_data():
    client = engine.connect_to_smartsheet()
    sheet, col_map = engine.load_sheet(client)
    unique = engine.extract_unique_values(sheet, col_map)
    all_rows_meta = []
    for row in sheet.rows:
        cm = {cell.column_id: cell for cell in row.cells}
        def _v(col_name, _cm=cm):
            cid = col_map.get(col_name)
            cell = _cm.get(cid) if cid else None
            raw = str(cell.value).strip() if (cell and cell.value is not None) else ""
            if col_name == engine.COL_YEAR:
                try:
                    raw = str(int(float(raw)))
                except (ValueError, TypeError):
                    pass
            elif col_name != engine.COL_MONTH:
                raw = raw.title()
            return raw
        all_rows_meta.append({
            "bizgroup": _v(engine.COL_BIZGROUP),
            "category": _v(engine.COL_CATEGORY),
            "brand":    _v(engine.COL_BRAND),
            "itemtype": _v(engine.COL_ITEMTYPE),
            "bg1ul":    _v(engine.COL_BG1UL),
            "cluster":  _v(engine.COL_CLUSTER),
            "country":  _v(engine.COL_COUNTRY),
            "year":     _v(engine.COL_YEAR),
            "month":    _v(engine.COL_MONTH),
        })
    return client, sheet, col_map, unique, all_rows_meta


def get_cascade_options(field, upstream, all_rows_meta):
    vals = set()
    for r in all_rows_meta:
        match = True
        for k, v in upstream.items():
            if not v:
                continue
            allowed = v if isinstance(v, list) else [v]
            if r.get(k, "") not in allowed:
                match = False
                break
        if match and r.get(field, ""):
            vals.add(r[field])
    if field == "month":
        return sorted(vals, key=month_sort_key)
    elif field == "year":
        return sorted(vals, key=lambda y: int(y) if y.isdigit() else 0)
    else:
        return sorted(vals, key=str.upper)


try:
    client, sheet, col_map, unique, all_rows_meta = get_sheet_data()
except Exception as e:
    st.error(f"Could not connect to Smartsheet: {e}")
    st.stop()

st.markdown("""
<div class="captivate-header">
  <div class="logo-ring"><div class="logo-inner"><div class="logo-core"></div></div></div>
  <div>
    <h1>Creative Audit Report Generator</h1>
    <p>Select filters, preview matching assets, then generate your PPTX report.</p>
  </div>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## Filters")
    st.caption("Leave any filter empty to include all values for that dimension.")

    def ms(label, options, key):
        prior = st.session_state.get(key, [])
        valid_prior = [v for v in prior if v in options]
        return st.multiselect(label, options=options, default=valid_prior, key=key)

    st.markdown('<div class="filter-section">Scope</div>', unsafe_allow_html=True)
    bizgroup_opts = get_cascade_options("bizgroup", {}, all_rows_meta)
    sel_bizgroup  = ms("Business Group", bizgroup_opts, "bizgroup")
    category_opts = get_cascade_options("category", {"bizgroup": sel_bizgroup}, all_rows_meta)
    sel_category  = ms("Category", category_opts, "category")

    st.markdown('<div class="filter-section">Brand</div>', unsafe_allow_html=True)
    brand_opts = get_cascade_options("brand", {"bizgroup": sel_bizgroup, "category": sel_category}, all_rows_meta)
    sel_brand  = ms("Brand", brand_opts, "brand")
    itemtype_opts = get_cascade_options("itemtype", {"bizgroup": sel_bizgroup, "category": sel_category, "brand": sel_brand}, all_rows_meta)
    sel_itemtype  = ms("Item Type", itemtype_opts, "itemtype")
    bg1ul_opts = get_cascade_options("bg1ul", {"bizgroup": sel_bizgroup, "category": sel_category, "brand": sel_brand, "itemtype": sel_itemtype}, all_rows_meta)
    sel_bg1ul  = ms("BG / 1UL", bg1ul_opts, "bg1ul")

    st.markdown('<div class="filter-section">Geography & Time</div>', unsafe_allow_html=True)
    cluster_opts = get_cascade_options("cluster", {"bizgroup": sel_bizgroup, "category": sel_category, "brand": sel_brand, "itemtype": sel_itemtype, "bg1ul": sel_bg1ul}, all_rows_meta)
    sel_cluster  = ms("BU / Cluster", cluster_opts, "cluster")
    country_opts = get_cascade_options("country", {"bizgroup": sel_bizgroup, "category": sel_category, "brand": sel_brand, "itemtype": sel_itemtype, "bg1ul": sel_bg1ul, "cluster": sel_cluster}, all_rows_meta)
    sel_country  = ms("Country", country_opts, "country")
    year_opts = get_cascade_options("year", {"bizgroup": sel_bizgroup, "category": sel_category, "brand": sel_brand, "itemtype": sel_itemtype, "bg1ul": sel_bg1ul, "cluster": sel_cluster, "country": sel_country}, all_rows_meta)
    sel_year  = ms("Year", year_opts, "year")
    month_opts_raw = get_cascade_options("month", {"bizgroup": sel_bizgroup, "category": sel_category, "brand": sel_brand, "itemtype": sel_itemtype, "bg1ul": sel_bg1ul, "cluster": sel_cluster, "country": sel_country, "year": sel_year}, all_rows_meta)
    month_opts = ["(All)"] + month_opts_raw
    col1, col2 = st.columns(2)
    with col1:
        month_start = st.selectbox("Month from", month_opts, key="month_start")
    with col2:
        month_end = st.selectbox("Month to", month_opts, key="month_end")

    st.markdown('<div class="filter-section">Score</div>', unsafe_allow_html=True)
    score_range = st.slider("Score range (%)", 0, 120, (0, 120), key="score_range")

    st.markdown('<div class="filter-section">Report Options</div>', unsafe_allow_html=True)
    analysis_mode = st.radio("Chart mode", options=["average", "volume"], format_func=lambda x: "Average score" if x == "average" else "Volume", horizontal=True, key="analysis_mode")
    st.caption("**Average** — how well do assets score?  \n**Volume** — how many assets per brand/country?")

    st.divider()
    preview_btn  = st.button("🔍  Preview matching assets", use_container_width=True)
    generate_btn = st.button("▶  Generate Report", use_container_width=True)


def build_filters():
    f = {}
    def _add(key, sel):
        if len(sel) == 1:
            f[key] = sel[0]
        elif len(sel) > 1:
            f[key] = sel
    _add("bizgroup", sel_bizgroup)
    _add("category", sel_category)
    _add("brand",    sel_brand)
    _add("itemtype", sel_itemtype)
    _add("bg1ul",    sel_bg1ul)
    _add("cluster",  sel_cluster)
    _add("country",  sel_country)
    _add("year",     sel_year)
    ms_val = month_start if month_start != "(All)" else None
    me_val = month_end   if month_end   != "(All)" else None
    if ms_val and me_val:
        f["month_start"] = ms_val
        f["month_end"]   = me_val
        f["month"] = ms_val if ms_val == me_val else f"{ms_val}–{me_val}"
    elif ms_val:
        f["month_start"] = f["month_end"] = f["month"] = ms_val
    elif me_val:
        f["month_start"] = f["month_end"] = f["month"] = me_val
    s_min, s_max = score_range
    if s_min > 0:   f["score_min"] = s_min
    if s_max < 120: f["score_max"] = s_max
    f["analysis_mode"] = analysis_mode
    return f


filters = build_filters()
header_text = engine.build_header_text(filters)

if st.session_state.get("last_filters") != filters:
    st.session_state.pop("rows_preview", None)
    st.session_state["last_filters"] = filters

st.markdown(f"**Scope:** `{header_text}`")

if preview_btn:
    with st.spinner("Fetching matching assets…"):
        rows = engine.get_filtered_rows(sheet, col_map, filters)
    st.session_state["rows_preview"] = rows

rows = st.session_state.get("rows_preview", [])

if rows:
    scores = []
    for r in rows:
        try:
            scores.append(float(r["score"].replace("%", "")))
        except Exception:
            pass
    avg  = sum(scores) / len(scores) if scores else 0
    unm  = sum(1 for s in scores if s >= 85)
    wall = sum(1 for s in scores if 55 <= s < 85)
    miss = sum(1 for s in scores if s < 55)

    st.markdown(f"""
    <div class="stat-row">
      <div class="stat-card"><div class="stat-label">Assets Found</div><div class="stat-value">{len(rows)}</div></div>
      <div class="stat-card"><div class="stat-label">Avg Score</div><div class="stat-value">{avg:.0f}%</div></div>
      <div class="stat-card stat-teal"><div class="stat-label">Unmissable</div><div class="stat-value">{unm}</div><div class="stat-sub">≥85%</div></div>
      <div class="stat-card stat-amber"><div class="stat-label">Wallpaper</div><div class="stat-value">{wall}</div><div class="stat-sub">55–84%</div></div>
      <div class="stat-card stat-red"><div class="stat-label">Missable</div><div class="stat-value">{miss}</div><div class="stat-sub">&lt;55%</div></div>
    </div>
    """, unsafe_allow_html=True)

    with st.expander(f"Preview — {len(rows)} matching asset(s)", expanded=True):
        rows_sorted = sorted(rows, key=lambda r: float(r["score"].replace("%","")) if r["score"].replace("%","").replace(".","").isdigit() else 0, reverse=True)
        table_rows = ""
        for r in rows_sorted[:100]:
            table_rows += f"<tr><td>{r.get('brand','—')}</td><td>{r.get('country','—')}</td><td>{r.get('category','—')}</td><td>{r.get('cluster','—')}</td><td>{r.get('itemtype','—')}</td><td>{r.get('bg1ul','—')}</td><td>{r.get('ref_code','—')}</td><td>{score_badge(r.get('score','—'))}</td></tr>"
        st.markdown(f"""<table class="preview-table"><thead><tr><th>Brand</th><th>Country</th><th>Category</th><th>BU / Cluster</th><th>Item Type</th><th>BG / 1UL</th><th>Ref Code</th><th>Score</th></tr></thead><tbody>{table_rows}</tbody></table>{"<p style='font-size:0.78rem;color:#888;margin-top:6px;'>Showing first 100 rows. All rows will be included in the report.</p>" if len(rows) > 100 else ""}""", unsafe_allow_html=True)

elif preview_btn:
    st.warning("No matching assets found for the selected filters.")

if generate_btn:
    if not rows:
        with st.spinner("Fetching matching assets…"):
            rows = engine.get_filtered_rows(sheet, col_map, filters)
        st.session_state["rows_preview"] = rows

    if not rows:
        st.warning("No matching assets found for the selected filters.")
    else:
        progress_bar = st.progress(0, text="Starting…")
        status_text  = st.empty()

        try:
            import tempfile
            with tempfile.TemporaryDirectory() as tmp_dir:
                engine.DOWNLOAD_DIR = Path(tmp_dir)
                from pptx import Presentation as _Prs
                prs = _Prs()
                prs.slide_width  = engine.SLIDE_W
                prs.slide_height = engine.SLIDE_H

                rows_sorted = sorted(rows, key=lambda r: float(r["score"].replace("%","")) if r["score"].replace("%","").replace(".","").isdigit() else 0, reverse=True)

                progress_bar.progress(1, text="Building title slide…")
                engine.build_title_slide(prs, header_text)

                _fail_counts = [0] * 15
                for _row in rows_sorted:
                    for _i in range(15):
                        if _row.get(f"crit_{_i+1}", "").strip().upper() == "N":
                            _fail_counts[_i] += 1

                progress_bar.progress(10, text="Building summary slide…")
                engine.build_summary_slide(prs, rows_sorted, header_text, analysis_mode=filters.get("analysis_mode", "average"))
                gc.collect()

                progress_bar.progress(20, text="Building analytics slide…")
                engine.build_analytics_slide(prs, rows_sorted, header_text, analysis_mode=filters.get("analysis_mode", "average"))
                gc.collect()

                progress_bar.progress(30, text="Generating AI insights (this may take ~20s)…")
                engine.build_synopsis_slide(prs, rows_sorted, filters, _fail_counts, header_text)
                gc.collect()

                for i, row in enumerate(rows_sorted, 1):
                    pct = 30 + int(i / len(rows_sorted) * 65)
                    progress_bar.progress(pct, text=f"Asset {i}/{len(rows_sorted)} — {row.get('brand','')} {row.get('ref_code','')}")
                    images = engine.fetch_images_for_row(client, row["row_id"])
                    engine.build_slide(prs, row, images, header_text)
                    for img_path in images:
                        try: img_path.unlink(missing_ok=True)
                        except Exception: pass
                    if i % 10 == 0:
                        gc.collect()

                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)

            progress_bar.progress(100, text="Done!")
            status_text.empty()

            label_parts = []
            for k in ["bizgroup","category","brand","itemtype","bg1ul","cluster","country","year","month"]:
                v = filters.get(k)
                if v:
                    label_parts.append(v if isinstance(v, str) else "+".join(v[:2]))
            file_label = "_".join(label_parts).replace(" ","_").replace("/","-") if label_parts else "All_Data"

            st.success(f"✅ Report ready — {len(rows_sorted)} asset slide(s) generated.")
            st.download_button(
                label="⬇  Download report (.pptx)",
                data=buf,
                file_name=f"{file_label}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )

        except Exception as e:
            progress_bar.empty()
            status_text.empty()
            st.error(f"Report generation failed: {e}")
            log.exception("Report generation failed")

st.divider()
st.caption("Captivate Insights · Creative Audit Report Generator")
