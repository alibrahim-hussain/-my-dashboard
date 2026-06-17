import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import io
from copy import deepcopy

st.set_page_config(page_title="One-Stop Clinic Dashboard", page_icon="🏥",
                   layout="wide", initial_sidebar_state="expanded")

TARGETS = {
    "Number of Enrolled Patients": None,
    "Average Number of Visits Prior Procedure": 3,
    "Average Waiting Time from Eligibility to 1st Visit": 14,
    "Average Waiting Time from Last Clinic Visit to Surgery": 28,
}
KPI_SHORT = {
    "Number of Enrolled Patients": "Enrolled Patients",
    "Average Number of Visits Prior Procedure": "Avg Visits",
    "Average Waiting Time from Eligibility to 1st Visit": "Wait → 1st Visit",
    "Average Waiting Time from Last Clinic Visit to Surgery": "Wait → Surgery",
}
KPI_UNITS = {
    "Number of Enrolled Patients": "",
    "Average Number of Visits Prior Procedure": " visits",
    "Average Waiting Time from Eligibility to 1st Visit": " days",
    "Average Waiting Time from Last Clinic Visit to Surgery": " days",
}
MONTH_ORDER = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
PALETTE = ["#38bdf8","#818cf8","#34d399","#fb923c","#f472b6","#a78bfa","#fbbf24"]

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;}
section[data-testid="stSidebar"]{background:#0a0f1e!important;border-right:1px solid #1e2a45;}
section[data-testid="stSidebar"] *{color:#cbd5e1!important;}
section[data-testid="stSidebar"] label{color:#94a3b8!important;font-size:11px;text-transform:uppercase;letter-spacing:1px;}
.main{background:#060d1f;}
.block-container{padding:1.5rem 2rem!important;}
.kpi-card{background:linear-gradient(135deg,#0f172a 0%,#1e293b 100%);border:1px solid #1e3a5f;border-radius:16px;padding:20px 24px;position:relative;overflow:hidden;transition:transform .2s,box-shadow .2s;}
.kpi-card:hover{transform:translateY(-2px);box-shadow:0 8px 32px rgba(56,189,248,.15);}
.kpi-card::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg,#38bdf8,#818cf8);}
.kpi-label{font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1.2px;margin-bottom:8px;}
.kpi-value{font-size:36px;font-weight:700;color:#f1f5f9;line-height:1;}
.kpi-sub{font-size:12px;color:#475569;margin-top:6px;}
.kpi-badge{display:inline-block;font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;margin-top:8px;}
.badge-good{background:rgba(34,197,94,.15);color:#4ade80;}
.badge-warn{background:rgba(251,191,36,.15);color:#fbbf24;}
.badge-bad{background:rgba(239,68,68,.15);color:#f87171;}
.section-header{font-size:12px;font-weight:600;color:#38bdf8;text-transform:uppercase;letter-spacing:2px;border-left:3px solid #38bdf8;padding-left:10px;margin:20px 0 12px;}
.dash-title{font-size:28px;font-weight:700;background:linear-gradient(135deg,#38bdf8,#818cf8);-webkit-background-clip:text;-webkit-text-fill-color:transparent;}
.dash-sub{font-size:13px;color:#475569;margin-top:2px;}
.stTabs [data-baseweb="tab-list"]{background:#0a0f1e;border-radius:12px;padding:4px;gap:2px;border:1px solid #1e2a45;}
.stTabs [data-baseweb="tab"]{background:transparent;color:#64748b;border-radius:8px;font-size:13px;font-weight:500;padding:8px 20px;}
.stTabs [aria-selected="true"]{background:linear-gradient(135deg,#1e3a5f,#1e2a45)!important;color:#38bdf8!important;}
.stSelectbox>div>div{background:#0f172a;border:1px solid #1e3a5f;border-radius:10px;color:#f1f5f9;}
.stMultiSelect>div>div{background:#0f172a;border:1px solid #1e3a5f;border-radius:10px;}
.stButton button{background:linear-gradient(135deg,#1e3a5f,#1e2a45);color:#38bdf8;border:1px solid #1e3a5f;border-radius:10px;font-weight:500;font-size:13px;transition:all .2s;}
.stButton button:hover{border-color:#38bdf8;box-shadow:0 0 16px rgba(56,189,248,.2);}
.sidebar-brand{text-align:center;padding:20px 0 8px;border-bottom:1px solid #1e2a45;margin-bottom:16px;}
.sidebar-brand h2{font-size:15px;font-weight:700;color:#38bdf8!important;margin:0;}
.sidebar-brand p{font-size:10px;color:#475569!important;margin:2px 0 0;}
</style>
""", unsafe_allow_html=True)


def parse_excel(file):
    wb_df = pd.read_excel(file, header=None, sheet_name=0)
    rows = wb_df.values.tolist()
    records = []
    i = 0
    while i < len(rows):
        row = rows[i]
        clinic_names, clinic_cols = [], []
        for c, val in enumerate(row):
            if isinstance(val, str) and val.strip() and val.strip() not in ('Month','KPI','Value'):
                clinic_names.append(val.strip())
                clinic_cols.append(c)
        if clinic_names:
            j = i + 1
            while j < len(rows) and all(v is None for v in rows[j]):
                j += 1
            if j < len(rows) and any(str(v) in ('Month','KPI','Value') for v in rows[j] if v):
                j += 1
            current_month = {c: None for c in clinic_cols}
            while j < len(rows):
                r = rows[j]
                if all(v is None for v in r):
                    j += 1
                    if j < len(rows) and all(v is None for v in rows[j]):
                        break
                    continue
                for ci, cc in enumerate(clinic_cols):
                    month_val = r[cc]   if cc < len(r) else None
                    kpi_val   = r[cc+1] if cc+1 < len(r) else None
                    val       = r[cc+2] if cc+2 < len(r) else None
                    if isinstance(month_val, str) and month_val.strip():
                        current_month[cc] = month_val.strip()
                    if isinstance(kpi_val, str) and kpi_val.strip() and val is not None:
                        try:
                            records.append({"Clinic":clinic_names[ci],"Month":current_month[cc],
                                            "KPI":kpi_val.strip(),"Value":float(val)})
                        except (TypeError, ValueError):
                            pass
                j += 1
            i = j
        else:
            i += 1
    df = pd.DataFrame(records)
    if df.empty:
        return df
    df["Month"] = pd.Categorical(df["Month"], categories=MONTH_ORDER, ordered=True)
    df["Facility"]   = df["Clinic"].apply(lambda x: x.split()[-1] if x else x)
    df["ClinicType"] = df["Clinic"].apply(lambda x: " ".join(x.split()[:-1]) if x else x)
    return df.sort_values("Month")


def load_sample():
    data = []
    clinics = {
        "Optha DGESH":         {"Jan":[81,2.5,5,22],"Feb":[64,2.1,12,28],"Mar":[45,2.3,0.02,35.9]},
        "HerniaLapCholen DMC": {"Jan":[18,2.3,0,8.6],"Feb":[22,1.9,8,8.1],"Mar":[14,1.9,19.1,14.7]},
        "HerniaLapCholen JHN": {"Jan":[16,1,7.1,2.6],"Feb":[4,1,7,7],"Mar":[3,1,5.3,3]},
        "Ortho Ras Tanura":    {"Jan":[18,1,5.9,8.1],"Feb":[8,1,6.3,8.7],"Mar":[2,1,2.5,9]},
        "HerniaLapCholen QHN": {"Jan":[28,1.1,6.5,4.6],"Feb":[26,1.9,13.5,6.9],"Mar":[28,2,24,24]},
    }
    kpis = list(TARGETS.keys())
    for clinic, months in clinics.items():
        for month, vals in months.items():
            for ki, kpi in enumerate(kpis):
                data.append({"Clinic":clinic,"Month":month,"KPI":kpi,"Value":vals[ki]})
    df = pd.DataFrame(data)
    df["Month"] = pd.Categorical(df["Month"], categories=MONTH_ORDER, ordered=True)
    df["Facility"]   = df["Clinic"].apply(lambda x: x.split()[-1])
    df["ClinicType"] = df["Clinic"].apply(lambda x: " ".join(x.split()[:-1]))
    return df.sort_values("Month")


DARK_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,23,42,0.6)",
    font=dict(family="Inter", color="#94a3b8"),
    xaxis=dict(gridcolor="#1e2a45", linecolor="#1e2a45", tickfont=dict(size=11)),
    yaxis=dict(gridcolor="#1e2a45", linecolor="#1e2a45", tickfont=dict(size=11)),
    legend=dict(bgcolor="rgba(10,15,30,0.8)", bordercolor="#1e2a45", borderwidth=1,
                font=dict(size=11), orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    margin=dict(l=40, r=20, t=60, b=40),
    hovermode="x unified",
)


def line_chart(sub, kpi, title, legend, target, cidx):
    fig = go.Figure()
    x = [str(m) for m in sorted(sub["Month"].unique(), key=lambda m: MONTH_ORDER.index(str(m)) if str(m) in MONTH_ORDER else 99)]
    y = sub.groupby("Month", observed=True)["Value"].mean().reindex(x)
    c = PALETTE[cidx % len(PALETTE)]
    r,g,b = int(c[1:3],16), int(c[3:5],16), int(c[5:7],16)
    fig.add_trace(go.Scatter(x=x, y=y, mode="lines+markers", name=legend,
        line=dict(color=c, width=2.5),
        marker=dict(size=8, line=dict(color="#0f172a", width=2)),
        fill="tozeroy", fillcolor=f"rgba({r},{g},{b},0.06)",
        hovertemplate=f"<b>%{{x}}</b><br>{legend}: %{{y:.1f}}{KPI_UNITS.get(kpi,'')}<extra></extra>"))
    if target:
        fig.add_hline(y=target, line_dash="dot", line_color="#fbbf24", line_width=1.5,
                      annotation_text=f"Target: {target}", annotation_position="right",
                      annotation_font=dict(color="#fbbf24", size=11))
    layout = deepcopy(DARK_LAYOUT)
    layout["title"] = dict(text=title, font=dict(size=13, color="#e2e8f0"), x=0.02)
    fig.update_layout(**layout)
    return fig


def multi_line_chart(df, kpi, title, target, clinics):
    fig = go.Figure()
    for ci, clinic in enumerate(clinics):
        sub = df[(df["Clinic"]==clinic) & (df["KPI"]==kpi)]
        if sub.empty: continue
        x = [str(m) for m in sorted(sub["Month"].unique(), key=lambda m: MONTH_ORDER.index(str(m)) if str(m) in MONTH_ORDER else 99)]
        y = sub.groupby("Month", observed=True)["Value"].mean().reindex(x)
        c = PALETTE[ci % len(PALETTE)]
        fig.add_trace(go.Scatter(x=x, y=y, mode="lines+markers", name=clinic,
            line=dict(color=c, width=2), marker=dict(size=7, line=dict(color="#0f172a",width=2)),
            hovertemplate=f"<b>%{{x}}</b><br>{clinic}: %{{y:.1f}}{KPI_UNITS.get(kpi,'')}<extra></extra>"))
    if target:
        fig.add_hline(y=target, line_dash="dot", line_color="#fbbf24", line_width=1.5,
                      annotation_text=f"Target: {target}", annotation_position="right",
                      annotation_font=dict(color="#fbbf24", size=11))
    layout = deepcopy(DARK_LAYOUT)
    layout["title"] = dict(text=title, font=dict(size=13, color="#e2e8f0"), x=0.02)
    fig.update_layout(**layout)
    return fig


# Session state
if "df" not in st.session_state:
    st.session_state.df = load_sample()
if "chart_titles" not in st.session_state:
    st.session_state.chart_titles = {k: k for k in TARGETS}
if "chart_legends" not in st.session_state:
    st.session_state.chart_legends = {k: KPI_SHORT[k] for k in TARGETS}

df_main = st.session_state.df.copy()

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="sidebar-brand"><h2>🏥 One-Stop Clinic</h2><p>MoC Dashboard · 2026</p></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-header">📂 Upload Data</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Excel file", type=["xlsx","xls"], label_visibility="collapsed")
    if uploaded:
        try:
            nd = parse_excel(uploaded)
            if not nd.empty:
                st.session_state.df = nd
                df_main = nd.copy()
                st.success(f"✅ {len(nd)} records loaded")
            else:
                st.warning("⚠️ No data found — check format")
        except Exception as e:
            st.error(str(e))

    st.divider()
    st.markdown('<div class="section-header">🔍 Filters</div>', unsafe_allow_html=True)

    all_fac = sorted(df_main["Facility"].dropna().unique())
    sel_fac = st.multiselect("Facility", all_fac, default=all_fac)
    df_f = df_main[df_main["Facility"].isin(sel_fac)] if sel_fac else df_main

    all_clin = sorted(df_f["Clinic"].dropna().unique())
    sel_clin = st.multiselect("Clinic / Program", all_clin, default=all_clin)
    df_f = df_f[df_f["Clinic"].isin(sel_clin)] if sel_clin else df_f

    all_mon = [m for m in MONTH_ORDER if m in df_f["Month"].astype(str).unique()]
    sel_mon = st.multiselect("Month", all_mon, default=all_mon)
    df_f = df_f[df_f["Month"].astype(str).isin(sel_mon)] if sel_mon else df_f

    st.divider()
    st.markdown('<div class="section-header">🎛️ Chart Controls</div>', unsafe_allow_html=True)
    ctrl_kpi = st.selectbox("Chart to edit", list(TARGETS.keys()), format_func=lambda x: KPI_SHORT[x])
    new_title = st.text_input("Title", value=st.session_state.chart_titles[ctrl_kpi])
    new_legend = st.text_input("Legend label", value=st.session_state.chart_legends[ctrl_kpi])
    if st.button("✅ Apply", use_container_width=True):
        st.session_state.chart_titles[ctrl_kpi] = new_title
        st.session_state.chart_legends[ctrl_kpi] = new_legend
        st.rerun()

# ── HEADER ────────────────────────────────────────────────────────────────────
fac_str  = ", ".join(sel_fac)  if sel_fac  else "All"
mon_str  = ", ".join(sel_mon)  if sel_mon  else "All"
clin_str = f"{len(sel_clin)} program(s)" if sel_clin else "All"
st.markdown(f'<div class="dash-title">One-Stop Clinic Dashboard</div>'
            f'<div class="dash-sub">Ministry of Health · {fac_str} · {clin_str} · {mon_str}</div>', unsafe_allow_html=True)
st.markdown("---")

tab1, tab2, tab3 = st.tabs(["📊  Overview", "📋  Raw Data", "🔬  Compare Clinics"])

# ── TAB 1: OVERVIEW ──────────────────────────────────────────────────────────
with tab1:
    cols = st.columns(4)
    for i, kpi in enumerate(TARGETS):
        sub = df_f[df_f["KPI"]==kpi]["Value"]
        val = sub.mean() if not sub.empty else 0
        target = TARGETS[kpi]
        with cols[i]:
            if target:
                pct = (val/target)*100
                if pct <= 80:   bcls, btxt = "badge-good", f"✓ {pct:.0f}% of target"
                elif pct <= 100: bcls, btxt = "badge-warn", f"⚠ {pct:.0f}% of target"
                else:            bcls, btxt = "badge-bad",  f"✕ {pct:.0f}% of target"
                badge  = f'<span class="kpi-badge {bcls}">{btxt}</span>'
                tsub   = f'<div class="kpi-sub">Target: {target}{KPI_UNITS.get(kpi,"")}</div>'
            else:
                badge, tsub = "", '<div class="kpi-sub">Total across selection</div>'

            if kpi == "Number of Enrolled Patients":
                disp = f"{int(df_f[df_f['KPI']==kpi]['Value'].sum()):,}"
            else:
                disp = f"{val:.1f}{KPI_UNITS.get(kpi,'')}"

            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-label">{KPI_SHORT[kpi]}</div>
                <div class="kpi-value">{disp}</div>
                {tsub}{badge}
            </div>""", unsafe_allow_html=True)

    st.markdown("")
    c1, c2 = st.columns(2)
    chart_defs = [
        ("Number of Enrolled Patients", c1, 0),
        ("Average Number of Visits Prior Procedure", c2, 1),
        ("Average Waiting Time from Eligibility to 1st Visit", c1, 2),
        ("Average Waiting Time from Last Clinic Visit to Surgery", c2, 3),
    ]
    for kpi, col, cidx in chart_defs:
        sub = df_f[df_f["KPI"]==kpi]
        fig = line_chart(sub, kpi, st.session_state.chart_titles[kpi],
                         st.session_state.chart_legends[kpi], TARGETS[kpi], cidx)
        with col:
            st.plotly_chart(fig, use_container_width=True, key=f"ch{cidx}")
            try:
                buf = io.BytesIO()
                fig.write_image(buf, format="png", width=900, height=450, scale=2)
                buf.seek(0)
                st.download_button("⬇ Export PNG", data=buf,
                    file_name=f"{KPI_SHORT[kpi].replace(' ','_')}.png",
                    mime="image/png", key=f"dl{cidx}", use_container_width=True)
            except Exception:
                st.caption("Install kaleido for PNG export: `pip install kaleido`")

# ── TAB 2: RAW DATA ──────────────────────────────────────────────────────────
with tab2:
    st.markdown('<div class="section-header">📋 Raw Data — Edit & Export</div>', unsafe_allow_html=True)
    disp = df_f[["Clinic","Facility","ClinicType","Month","KPI","Value"]].copy()
    disp["Month"] = disp["Month"].astype(str)
    edited = st.data_editor(disp, use_container_width=True, num_rows="dynamic",
        column_config={
            "Value": st.column_config.NumberColumn("Value", format="%.2f"),
            "Month": st.column_config.SelectboxColumn("Month", options=MONTH_ORDER),
            "KPI":   st.column_config.SelectboxColumn("KPI",   options=list(TARGETS.keys())),
        }, hide_index=True, key="de")
    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("💾 Apply Edits", use_container_width=True):
            st.success("✅ Edits noted — re-upload or refresh to persist")
    with col_b:
        csv = disp.to_csv(index=False).encode()
        st.download_button("⬇ Export CSV", csv, "osc_data.csv", "text/csv", use_container_width=True)
    st.markdown(f'<div style="font-size:12px;color:#475569;margin-top:8px">{len(disp):,} rows · {disp["Clinic"].nunique()} clinics · {disp["Month"].nunique()} months</div>', unsafe_allow_html=True)

# ── TAB 3: COMPARE ───────────────────────────────────────────────────────────
with tab3:
    st.markdown('<div class="section-header">🔬 Clinic Comparison</div>', unsafe_allow_html=True)
    cc1, cc2 = st.columns(2)
    with cc1:
        cmp_kpi = st.selectbox("KPI", list(TARGETS.keys()), format_func=lambda x: KPI_SHORT[x], key="ck")
    with cc2:
        chart_type = st.selectbox("Chart Type", ["Line","Bar","Radar"], key="ct")
    all_c = sorted(df_f["Clinic"].dropna().unique())
    cmp_clinics = st.multiselect("Select Clinics", all_c, default=all_c[:min(4,len(all_c))], key="cc")

    if cmp_clinics:
        sub = df_f[(df_f["KPI"]==cmp_kpi) & (df_f["Clinic"].isin(cmp_clinics))]
        if chart_type == "Line":
            fig = multi_line_chart(df_f, cmp_kpi, f"{KPI_SHORT[cmp_kpi]} — Clinic Comparison", TARGETS[cmp_kpi], cmp_clinics)
        elif chart_type == "Bar":
            agg = sub.groupby(["Clinic","Month"], observed=True)["Value"].mean().reset_index()
            agg["Month"] = agg["Month"].astype(str)
            fig = px.bar(agg, x="Month", y="Value", color="Clinic", barmode="group",
                         template="plotly_dark", color_discrete_sequence=PALETTE,
                         title=f"{KPI_SHORT[cmp_kpi]} — By Clinic & Month")
            layout = deepcopy(DARK_LAYOUT)
            layout["title"] = dict(text=f"{KPI_SHORT[cmp_kpi]} — By Clinic", font=dict(size=13,color="#e2e8f0"), x=0.02)
            fig.update_layout(**layout)
            if TARGETS[cmp_kpi]:
                fig.add_hline(y=TARGETS[cmp_kpi], line_dash="dot", line_color="#fbbf24",
                              annotation_text=f"Target: {TARGETS[cmp_kpi]}")
        else:
            agg = sub.groupby("Clinic", observed=True)["Value"].mean().reset_index()
            cats = agg["Clinic"].tolist() + [agg["Clinic"].iloc[0]]
            vals = agg["Value"].tolist() + [agg["Value"].iloc[0]]
            fig = go.Figure(go.Scatterpolar(r=vals, theta=cats, fill="toself",
                fillcolor="rgba(56,189,248,.15)", line=dict(color="#38bdf8")))
            layout = deepcopy(DARK_LAYOUT)
            layout["polar"] = dict(bgcolor="#0f172a",
                radialaxis=dict(gridcolor="#1e2a45"), angularaxis=dict(gridcolor="#1e2a45"))
            layout["title"] = dict(text=f"{KPI_SHORT[cmp_kpi]} — Radar", x=0.02, font=dict(size=13,color="#e2e8f0"))
            fig.update_layout(**layout)

        st.plotly_chart(fig, use_container_width=True)
        try:
            buf = io.BytesIO()
            fig.write_image(buf, format="png", width=1200, height=600, scale=2)
            buf.seek(0)
            st.download_button("⬇ Export Chart PNG", buf,
                f"compare_{KPI_SHORT[cmp_kpi].replace(' ','_')}.png", "image/png")
        except Exception:
            st.caption("Install kaleido for PNG export")
    else:
        st.info("Select at least one clinic above.")
