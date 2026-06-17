import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import io
import base64
import zipfile
from copy import deepcopy
from PIL import Image

st.set_page_config(page_title="One-Stop Clinic 2026", page_icon="🏥",
                   layout="wide", initial_sidebar_state="collapsed")

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

# Chart layout: 5 charts — big left + 2 right stacked + 2 bottom
CHART_ORDER = [
    ("Average Waiting Time from Eligibility to 1st Visit", "big"),
    ("Number of Enrolled Patients", "small"),
    ("Average Number of Visits Prior Procedure", "small"),
    ("Average Waiting Time from Last Clinic Visit to Surgery", "bottom"),
    ("Average Waiting Time from Last Clinic Visit to Surgery", "bottom"),
]

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;}
.main{background:#f0f4f8;}
.block-container{padding:0!important;max-width:100%!important;}

/* Header */
.osc-header{
    background:linear-gradient(90deg,#1a4f8a 0%,#1e5fa0 100%);
    padding:10px 24px;display:flex;align-items:center;
    justify-content:space-between;
}
.osc-header-title{color:#fff;font-size:18px;font-weight:600;letter-spacing:.3px;}
.osc-logo-text{color:#a8d4f0;font-size:12px;text-align:right;line-height:1.5;}
.osc-logo-name{font-size:14px;font-weight:600;color:#fff;}

/* Filter bar */
.filter-bar{
    background:#dce8f5;padding:10px 24px;
    display:flex;align-items:center;gap:16px;
    border-bottom:1px solid #b8cfe8;
}

/* Info cards row */
.info-cards{display:flex;gap:12px;padding:14px 24px 0;}
.info-card{
    background:#fff;border:1.5px solid #1a4f8a;border-radius:4px;
    padding:8px 20px;text-align:center;min-width:140px;
}
.info-card-label{font-size:11px;color:#1a4f8a;font-weight:500;}
.info-card-value{font-size:15px;font-weight:700;color:#1a4f8a;}

/* KPI Table */
.kpi-table{
    background:#fff;border:1px solid #d0dcea;border-radius:4px;
    overflow:hidden;min-width:320px;
}
.kpi-table-header{
    background:#1a4f8a;color:#fff;
    display:flex;justify-content:space-between;
    padding:8px 14px;font-size:12px;font-weight:600;
}
.kpi-row-item{
    display:flex;justify-content:space-between;align-items:center;
    padding:7px 14px;border-bottom:1px solid #eef2f7;font-size:12px;
}
.kpi-row-item:last-child{border-bottom:none;}
.kpi-row-item:nth-child(even){background:#f7fafd;}
.kpi-target{font-weight:700;color:#1a4f8a;}

/* Chart container */
.charts-area{padding:14px 24px;}

/* Section tabs */
.stTabs [data-baseweb="tab-list"]{
    background:#dce8f5;padding:4px 24px 0;gap:4px;
    border-bottom:2px solid #b8cfe8;
}
.stTabs [data-baseweb="tab"]{
    background:transparent;color:#4a6a8a;
    font-size:13px;font-weight:500;padding:8px 20px;
    border-radius:6px 6px 0 0;
}
.stTabs [aria-selected="true"]{
    background:#fff!important;color:#1a4f8a!important;
    border:1px solid #b8cfe8;border-bottom:2px solid #fff;
}

/* Select boxes */
.stSelectbox>div>div{
    background:#fff;border:1px solid #b8cfe8;
    border-radius:6px;color:#1a4f8a;font-size:13px;
}

/* Multiselect tags — blue not red */
span[data-baseweb="tag"]{
    background:#1a4f8a!important;
}

/* Export buttons */
.stDownloadButton button{
    background:#1a4f8a;color:#fff;border:none;
    border-radius:6px;font-size:12px;padding:4px 14px;
    font-weight:500;
}
.stDownloadButton button:hover{background:#1e5fa0;}
.stButton button{
    background:#fff;color:#1a4f8a;
    border:1.5px solid #1a4f8a;border-radius:6px;
    font-size:12px;padding:4px 14px;font-weight:500;
}

/* Data editor */
.stDataEditor{border-radius:8px;overflow:hidden;}
</style>
""", unsafe_allow_html=True)


# ── Data helpers ──────────────────────────────────────────────────────────────
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
                    month_val = r[cc]   if cc   < len(r) else None
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


# ── Chart builder ─────────────────────────────────────────────────────────────
BLUE   = "#1665c4"
GOLD   = "#e8a020"
RED    = "#c03020"
FILL   = "rgba(22,101,196,0.07)"
BGCOL  = "rgba(255,255,255,1)"
GRIDC  = "#e8eef5"

DARK_LAYOUT = dict(
    paper_bgcolor=BGCOL,
    plot_bgcolor=BGCOL,
    font=dict(family="Inter", color="#2a3a4a", size=11),
    xaxis=dict(gridcolor=GRIDC, linecolor=GRIDC, tickfont=dict(size=10,color="#4a6a8a"), zeroline=False),
    yaxis=dict(gridcolor=GRIDC, linecolor="#dce8f5", tickfont=dict(size=10,color="#4a6a8a"), zeroline=True, zerolinecolor=GRIDC),
    legend=dict(bgcolor="rgba(255,255,255,0.9)", bordercolor="#dce8f5", borderwidth=1,
                font=dict(size=10,color="#4a6a8a"), orientation="v",
                yanchor="bottom", y=0.02, xanchor="right", x=0.99),
    margin=dict(l=45,r=20,t=45,b=40),
    hovermode="x unified",
)


def make_chart(sub, kpi, title, legend_name, target, line_color=BLUE):
    fig = go.Figure()
    x = [str(m) for m in sorted(sub["Month"].unique(),
         key=lambda m: MONTH_ORDER.index(str(m)) if str(m) in MONTH_ORDER else 99)]
    y = sub.groupby("Month", observed=True)["Value"].mean().reindex(x)

    fig.add_trace(go.Scatter(
        x=x, y=y, mode="lines+markers", name=legend_name,
        line=dict(color=line_color, width=2),
        marker=dict(size=7, color=line_color, line=dict(color="#fff", width=1.5)),
        fill="tozeroy", fillcolor=FILL,
        hovertemplate=f"<b>%{{x}}</b><br>{legend_name}: %{{y:.1f}}{KPI_UNITS.get(kpi,'')}<extra></extra>"
    ))
    if target:
        fig.add_hline(y=target, line_dash="solid", line_color=GOLD, line_width=1.5,
                      annotation_text=f"Target", annotation_position="bottom right",
                      annotation_font=dict(color=GOLD, size=9))

    layout = deepcopy(DARK_LAYOUT)
    layout["title"] = dict(text=f"<b>{title}</b>", font=dict(size=12, color="#1a4f8a"), x=0.5, xanchor="center")
    fig.update_layout(**layout)
    return fig


def fig_to_html_bytes(fig):
    """Export figure as standalone HTML (works without kaleido)."""
    html_str = fig.to_html(include_plotlyjs="cdn", full_html=True)
    return html_str.encode("utf-8"), "html"


# ── Session state ─────────────────────────────────────────────────────────────
if "df" not in st.session_state:
    st.session_state.df = load_sample()
if "chart_titles" not in st.session_state:
    st.session_state.chart_titles = {k: k for k in TARGETS}
if "chart_legends" not in st.session_state:
    st.session_state.chart_legends = {
        "Number of Enrolled Patients": "Enrolled Patients",
        "Average Number of Visits Prior Procedure": "No. of Visits",
        "Average Waiting Time from Eligibility to 1st Visit": "Waiting Time",
        "Average Waiting Time from Last Clinic Visit to Surgery": "Waiting Time",
    }
if "logo_b64" not in st.session_state:
    st.session_state.logo_b64 = None

df_main = st.session_state.df.copy()

# ── HEADER ────────────────────────────────────────────────────────────────────
hcol1, hcol2, hcol3 = st.columns([2, 3, 2])
with hcol1:
    if st.session_state.logo_b64:
        st.markdown(f'<img src="data:image/png;base64,{st.session_state.logo_b64}" style="height:52px;margin-top:4px"/>', unsafe_allow_html=True)
    else:
        st.markdown('''<div style="background:#1a4f8a;border-radius:8px;padding:8px 14px;display:inline-block">
            <div style="color:#fff;font-size:13px;font-weight:600">تجمع الشرقية الصحي</div>
            <div style="color:#a8d4f0;font-size:10px">شركة الصحة القابضة</div>
        </div>''', unsafe_allow_html=True)
with hcol2:
    st.markdown('''<div style="text-align:center;padding:6px 0">
        <div style="font-size:20px;font-weight:700;color:#1a4f8a">One-Stop Clinic 2026</div>
        <div style="font-size:11px;color:#4a6a8a">Ministry of Health Dashboard</div>
    </div>''', unsafe_allow_html=True)
with hcol3:
    pass
st.markdown('<hr style="border:none;border-top:3px solid #1a4f8a;margin:4px 0 12px"/>', unsafe_allow_html=True)

# ── TABS ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📊  Overview", "📋  Raw Data", "🔬  Compare Clinics"])


# ════════════════════ TAB 1 — OVERVIEW ═══════════════════════════════════════
with tab1:
    # ── Upload + Filters row ─────────────────────────────────────────────────
    fc1, fc2, fc3, fc4, fc5 = st.columns([2, 2, 2, 2, 2])

    with fc1:
        uploaded = st.file_uploader("📂 Upload Excel", type=["xlsx","xls"], label_visibility="collapsed")
        if uploaded:
            try:
                nd = parse_excel(uploaded)
                if not nd.empty:
                    st.session_state.df = nd
                    df_main = nd.copy()
                    st.success(f"✅ {len(nd)} records")
                else:
                    st.warning("⚠️ No data")
            except Exception as e:
                st.error(str(e))

    with fc2:
        logo_up = st.file_uploader("🖼 Logo", type=["png","jpg","jpeg"], label_visibility="collapsed")
        if logo_up:
            img_bytes = logo_up.read()
            st.session_state.logo_b64 = base64.b64encode(img_bytes).decode()
            st.rerun()

    # Facility filter — single select
    all_fac = sorted(df_main["Facility"].dropna().unique())
    with fc3:
        sel_fac = st.selectbox("Facility", ["All"] + all_fac, index=0)

    df_f = df_main if sel_fac == "All" else df_main[df_main["Facility"] == sel_fac]

    # Clinic filter — single select
    all_clin = sorted(df_f["Clinic"].dropna().unique())
    with fc4:
        sel_clin = st.selectbox("Clinic / Program", ["All"] + all_clin, index=0)

    df_f = df_f if sel_clin == "All" else df_f[df_f["Clinic"] == sel_clin]

    # Month filter
    all_mon = [m for m in MONTH_ORDER if m in df_f["Month"].astype(str).unique()]
    with fc5:
        sel_mon = st.multiselect("Month", all_mon, default=all_mon)
    df_f = df_f[df_f["Month"].astype(str).isin(sel_mon)] if sel_mon else df_f

    st.markdown("")

    # ── Info cards + KPI table ────────────────────────────────────────────────
    inf1, inf2, inf3, inf4, kpi_col = st.columns([1.2, 1.2, 1.2, 1.2, 2.5])

    fac_display  = sel_fac  if sel_fac  != "All" else "All"
    clin_display = sel_clin if sel_clin != "All" else "All"
    enrolled_total = int(df_f[df_f["KPI"]=="Number of Enrolled Patients"]["Value"].sum())

    with inf1:
        st.markdown(f"""<div class="info-card">
            <div class="info-card-label">Facility</div>
            <div class="info-card-value">{fac_display}</div>
        </div>""", unsafe_allow_html=True)
    with inf2:
        st.markdown(f"""<div class="info-card">
            <div class="info-card-label">Clinic</div>
            <div class="info-card-value">{clin_display}</div>
        </div>""", unsafe_allow_html=True)
    with inf3:
        st.markdown(f"""<div class="info-card" style="background:#1a4f8a;">
            <div class="info-card-label" style="color:#a8d4f0;">Total Enrolled Patients</div>
            <div class="info-card-value" style="color:#fff;font-size:22px;">{enrolled_total}</div>
        </div>""", unsafe_allow_html=True)
    with inf4:
        # Chart controls
        with st.expander("🎛️ Chart Controls"):
            ctrl_kpi = st.selectbox("Chart", list(TARGETS.keys()), format_func=lambda x: KPI_SHORT[x])
            new_title  = st.text_input("Title",  value=st.session_state.chart_titles[ctrl_kpi])
            new_legend = st.text_input("Legend", value=st.session_state.chart_legends[ctrl_kpi])
            if st.button("✅ Apply"):
                st.session_state.chart_titles[ctrl_kpi]  = new_title
                st.session_state.chart_legends[ctrl_kpi] = new_legend
                st.rerun()

    with kpi_col:
        rows_html = ""
        for kpi, target in TARGETS.items():
            tval = str(target) if target else "—"
            rows_html += f"""<div class="kpi-row-item">
                <span>{kpi}</span>
                <span class="kpi-target">{tval}</span>
            </div>"""
        st.markdown(f"""<div class="kpi-table">
            <div class="kpi-table-header"><span>KPI</span><span>Target ▲</span></div>
            {rows_html}
        </div>""", unsafe_allow_html=True)

    st.markdown("---")

    # ── 5 Charts layout ──────────────────────────────────────────────────────
    # Row 1: big left (Wait → 1st Visit) | right column with 2 stacked
    row1_left, row1_right = st.columns([1.4, 1])

    with row1_left:
        kpi = "Average Waiting Time from Eligibility to 1st Visit"
        sub = df_f[df_f["KPI"]==kpi]
        fig = make_chart(sub, kpi, st.session_state.chart_titles[kpi],
                         st.session_state.chart_legends[kpi], TARGETS[kpi])
        fig.update_layout(height=420)
        st.plotly_chart(fig, use_container_width=True, key="c0")
        img_bytes, fmt = fig_to_html_bytes(fig)
        st.download_button("⬇ Export HTML", img_bytes,
            f"chart_wait_1st.{fmt}", "text/html", key="dl0", use_container_width=True)

    with row1_right:
        for idx, kpi in enumerate(["Number of Enrolled Patients",
                                    "Average Number of Visits Prior Procedure"]):
            sub = df_f[df_f["KPI"]==kpi]
            fig = make_chart(sub, kpi, st.session_state.chart_titles[kpi],
                             st.session_state.chart_legends[kpi], TARGETS[kpi])
            fig.update_layout(height=195, margin=dict(l=45,r=20,t=35,b=30))
            st.plotly_chart(fig, use_container_width=True, key=f"c{idx+1}")
            img_bytes, fmt = fig_to_html_bytes(fig)
            st.download_button("⬇ Export HTML", img_bytes,
                f"chart_{idx+1}.{fmt}", "text/html", key=f"dl{idx+1}", use_container_width=True)

    # Row 2: 2 bottom charts side by side (both Wait → Surgery, different clinics)
    bot1, bot2 = st.columns(2)
    kpi_bot = "Average Waiting Time from Last Clinic Visit to Surgery"

    # Determine which clinics to show in bottom charts
    available_clinics = sorted(df_f["Clinic"].dropna().unique())
    clinic_a = available_clinics[0] if len(available_clinics) > 0 else None
    clinic_b = available_clinics[1] if len(available_clinics) > 1 else clinic_a

    for i, (col, clinic) in enumerate([(bot1, clinic_a), (bot2, clinic_b)]):
        with col:
            sub = df_f[(df_f["KPI"]==kpi_bot) & (df_f["Clinic"]==clinic)] if clinic else df_f[df_f["KPI"]==kpi_bot]
            title = f"{st.session_state.chart_titles[kpi_bot]}"
            fig = make_chart(sub, kpi_bot, title,
                             st.session_state.chart_legends[kpi_bot], TARGETS[kpi_bot])
            fig.update_layout(height=300)
            st.plotly_chart(fig, use_container_width=True, key=f"cbot{i}")
            img_bytes, fmt = fig_to_html_bytes(fig)
            st.download_button("⬇ Export HTML", img_bytes,
                f"chart_surgery_{i+1}.{fmt}", "text/html", key=f"dlbot{i}", use_container_width=True)

    # ── Export All ────────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("#### Export All — per Clinic")

    all_clinics_list = sorted(df_main["Clinic"].dropna().unique())
    if st.button("📦 Export All Clinics as PNG (ZIP)", use_container_width=False):
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w") as zf:
            for clinic in all_clinics_list:
                df_clinic = df_main[df_main["Clinic"]==clinic]
                for kpi in TARGETS:
                    sub = df_clinic[df_clinic["KPI"]==kpi]
                    if sub.empty:
                        continue
                    fig = make_chart(sub, kpi,
                                     st.session_state.chart_titles[kpi],
                                     st.session_state.chart_legends[kpi],
                                     TARGETS[kpi])
                    fig.update_layout(height=400,
                        title_text=f"<b>{kpi}</b><br><span style='font-size:10px;color:#4a6a8a'>{clinic}</span>")
                    try:
                        html_str = fig.to_html(include_plotlyjs="cdn", full_html=True)
                        safe_clinic = clinic.replace(" ","_").replace("/","_")
                        safe_kpi    = KPI_SHORT[kpi].replace(" ","_").replace("→","to")
                        zf.writestr(f"{safe_clinic}/{safe_kpi}.html", html_str.encode("utf-8"))
                    except Exception:
                        pass
        zip_buf.seek(0)
        st.download_button("⬇ Download ZIP (HTML charts)", zip_buf.read(),
                           "osc_all_clinics.zip", "application/zip")


# ════════════════════ TAB 2 — RAW DATA ══════════════════════════════════════
with tab2:
    st.markdown("#### Raw Data — View & Edit")
    disp = df_main[["Clinic","Facility","ClinicType","Month","KPI","Value"]].copy()
    disp["Month"] = disp["Month"].astype(str)

    # Filters for raw data
    rc1, rc2 = st.columns(2)
    with rc1:
        raw_fac = st.selectbox("Facility", ["All"] + sorted(disp["Facility"].unique()), key="rf")
    with rc2:
        raw_clin = st.selectbox("Clinic", ["All"] + sorted(disp["Clinic"].unique()), key="rc")

    if raw_fac  != "All": disp = disp[disp["Facility"]==raw_fac]
    if raw_clin != "All": disp = disp[disp["Clinic"]==raw_clin]

    edited = st.data_editor(disp, use_container_width=True, num_rows="dynamic",
        column_config={
            "Value": st.column_config.NumberColumn("Value", format="%.2f"),
            "Month": st.column_config.SelectboxColumn("Month", options=MONTH_ORDER),
            "KPI":   st.column_config.SelectboxColumn("KPI",   options=list(TARGETS.keys())),
        }, hide_index=True, key="de")

    csv = disp.to_csv(index=False).encode()
    st.download_button("⬇ Export CSV", csv, "osc_data.csv", "text/csv")
    st.caption(f"{len(disp):,} rows · {disp['Clinic'].nunique()} clinics")


# ════════════════════ TAB 3 — COMPARE ══════════════════════════════════════
with tab3:
    st.markdown("#### Clinic Comparison")
    cc1, cc2 = st.columns(2)
    with cc1:
        cmp_kpi = st.selectbox("KPI", list(TARGETS.keys()), format_func=lambda x: KPI_SHORT[x])
    with cc2:
        chart_type = st.selectbox("Chart Type", ["Line","Bar","Radar"])

    all_c = sorted(df_main["Clinic"].dropna().unique())
    cmp_clinics = st.multiselect("Clinics to Compare", all_c, default=all_c[:min(4,len(all_c))])

    if cmp_clinics:
        PALETTE = [BLUE,"#2a9d8f","#e76f51","#a78bfa","#fbbf24","#f472b6"]
        sub_all = df_main[(df_main["KPI"]==cmp_kpi) & (df_main["Clinic"].isin(cmp_clinics))]
        fig = go.Figure()

        if chart_type == "Line":
            for ci, clinic in enumerate(cmp_clinics):
                sub = sub_all[sub_all["Clinic"]==clinic]
                if sub.empty: continue
                x = [str(m) for m in sorted(sub["Month"].unique(),
                     key=lambda m: MONTH_ORDER.index(str(m)) if str(m) in MONTH_ORDER else 99)]
                y = sub.groupby("Month",observed=True)["Value"].mean().reindex(x)
                fig.add_trace(go.Scatter(x=x,y=y,mode="lines+markers",name=clinic,
                    line=dict(color=PALETTE[ci%len(PALETTE)],width=2),
                    marker=dict(size=7,line=dict(color="#fff",width=1.5))))
            if TARGETS[cmp_kpi]:
                fig.add_hline(y=TARGETS[cmp_kpi],line_dash="solid",line_color=GOLD,
                    annotation_text="Target",annotation_position="bottom right")
            layout = deepcopy(DARK_LAYOUT)
            layout["title"] = dict(text=f"<b>{KPI_SHORT[cmp_kpi]} — Clinic Comparison</b>",
                                   font=dict(size=13,color="#1a4f8a"),x=0.5,xanchor="center")
            layout["legend"] = dict(bgcolor="rgba(255,255,255,0.9)",bordercolor="#dce8f5",
                                    borderwidth=1,font=dict(size=10),
                                    orientation="h",yanchor="bottom",y=1.02,xanchor="left",x=0)
            fig.update_layout(**layout)

        elif chart_type == "Bar":
            agg = sub_all.groupby(["Clinic","Month"],observed=True)["Value"].mean().reset_index()
            agg["Month"] = agg["Month"].astype(str)
            fig = px.bar(agg,x="Month",y="Value",color="Clinic",barmode="group",
                         color_discrete_sequence=PALETTE,
                         title=f"{KPI_SHORT[cmp_kpi]} — By Clinic & Month")
            layout = deepcopy(DARK_LAYOUT)
            layout["title"] = dict(text=f"<b>{KPI_SHORT[cmp_kpi]}</b>",
                                   font=dict(size=13,color="#1a4f8a"),x=0.5,xanchor="center")
            fig.update_layout(**layout)
            if TARGETS[cmp_kpi]:
                fig.add_hline(y=TARGETS[cmp_kpi],line_dash="solid",line_color=GOLD)

        else:  # Radar
            agg = sub_all.groupby("Clinic",observed=True)["Value"].mean().reset_index()
            cats = agg["Clinic"].tolist()+[agg["Clinic"].iloc[0]]
            vals = agg["Value"].tolist()+[agg["Value"].iloc[0]]
            fig.add_trace(go.Scatterpolar(r=vals,theta=cats,fill="toself",
                fillcolor=f"rgba(22,101,196,0.15)",line=dict(color=BLUE,width=2)))
            layout = deepcopy(DARK_LAYOUT)
            layout["polar"] = dict(bgcolor="#f7fafd",
                radialaxis=dict(gridcolor=GRIDC),angularaxis=dict(gridcolor=GRIDC))
            layout["title"] = dict(text=f"<b>{KPI_SHORT[cmp_kpi]}</b>",
                                   font=dict(size=13,color="#1a4f8a"),x=0.5,xanchor="center")
            fig.update_layout(**layout)

        st.plotly_chart(fig, use_container_width=True)
        img_bytes, fmt = fig_to_html_bytes(fig)
        st.download_button("⬇ Export Chart", img_bytes,
            f"compare_{KPI_SHORT[cmp_kpi].replace(' ','_').replace('→','to')}.{fmt}",
            "text/html")
    else:
        st.info("Select at least one clinic.")
