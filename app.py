import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
from datetime import datetime

st.set_page_config(
    page_title="SeminarIQ — Command Center",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ══════════════════════════════════════════════════════════════════════════════
# DESIGN SYSTEM
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=JetBrains+Mono:wght@300;400;500&family=Outfit:wght@300;400;500;600&display=swap');

*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
html,body,[class*="css"],.stApp{background:#06080F !important;font-family:'Outfit',sans-serif;}
.block-container{padding:0 !important;max-width:100% !important;}
section[data-testid="stSidebar"]{display:none;}
#MainMenu,footer,header,.stDeployButton,div[data-testid="stToolbar"]{display:none !important;visibility:hidden;}

/* ── TOPBAR ── */
.topbar{
    display:flex;align-items:center;justify-content:space-between;
    padding:.85rem 2rem;background:#06080F;
    border-bottom:1px solid rgba(255,255,255,0.05);
    position:sticky;top:0;z-index:100;
    backdrop-filter:blur(10px);
}
.logo{font-family:'Syne',sans-serif;font-weight:800;font-size:16px;color:#fff;letter-spacing:-.5px;}
.logo em{color:#00C896;font-style:normal;}
.topbar-sep{width:1px;height:18px;background:rgba(255,255,255,0.08);margin:0 14px;}
.topbar-sub{font-family:'JetBrains Mono',monospace;font-size:10px;color:#2D3748;letter-spacing:.1em;}
.topbar-right{display:flex;align-items:center;gap:8px;}
.dot{width:5px;height:5px;border-radius:50%;background:#00C896;box-shadow:0 0 6px #00C896;animation:blink 2s infinite;}
@keyframes blink{0%,100%{opacity:1}50%{opacity:.3}}
.live{font-family:'JetBrains Mono',monospace;font-size:9px;color:#00C896;letter-spacing:.15em;}
.clock{font-family:'JetBrains Mono',monospace;font-size:9px;color:#1A202C;margin-left:10px;}

/* ── PAGE WRAP ── */
.page{padding:1.6rem 2rem 3rem 2rem;}

/* ── SECTION HEADER ── */
.sh{display:flex;align-items:center;gap:10px;margin:1.8rem 0 1rem 0;}
.sh-line{flex:1;height:1px;background:linear-gradient(90deg,rgba(0,200,150,.2),transparent);}
.sh-label{font-family:'JetBrains Mono',monospace;font-size:9px;font-weight:500;
           color:#00C896;letter-spacing:.2em;text-transform:uppercase;white-space:nowrap;}
.sh-note{font-family:'JetBrains Mono',monospace;font-size:9px;color:#1A202C;}

/* ── KPI STRIP ── */
.kpi-strip{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));
           gap:1px;background:rgba(255,255,255,0.04);
           border:1px solid rgba(255,255,255,0.05);border-radius:6px;overflow:hidden;margin-bottom:2px;}
.kpi-cell{background:#0B0F1A;padding:1rem 1.1rem;position:relative;}
.kpi-cell:hover{background:#0F1628;}
.kpi-cell::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;}
.kpi-cell.g::before{background:#00C896;}
.kpi-cell.r::before{background:#F75F7A;}
.kpi-cell.b::before{background:#4C6EF5;}
.kpi-cell.a::before{background:#F6B93B;}
.kpi-cell.t::before{background:#22D3EE;}
.kpi-cell.p::before{background:#A78BFA;}
.kpi-lbl{font-family:'JetBrains Mono',monospace;font-size:8.5px;color:#2D3748;
          letter-spacing:.12em;text-transform:uppercase;margin-bottom:7px;}
.kpi-val{font-family:'Syne',sans-serif;font-size:22px;font-weight:700;color:#F1F5F9;line-height:1;margin-bottom:5px;}
.kpi-val.sm{font-size:17px;}
.kpi-sub{font-family:'JetBrains Mono',monospace;font-size:9px;}
.kpi-sub.g{color:#00C896;} .kpi-sub.r{color:#F75F7A;} .kpi-sub.a{color:#F6B93B;} .kpi-sub.n{color:#2D3748;}

/* ── CHART SHELL ── */
.cshell{background:#0B0F1A;border:1px solid rgba(255,255,255,0.05);
        border-radius:6px;padding:1.1rem 1.2rem;position:relative;overflow:hidden;}
.cshell::after{content:'';position:absolute;top:0;left:0;width:2px;height:100%;
               background:linear-gradient(180deg,#00C896 0%,transparent 100%);}
.ctitle{font-family:'JetBrains Mono',monospace;font-size:8.5px;font-weight:500;
        color:#2D3748;letter-spacing:.18em;text-transform:uppercase;
        margin-bottom:.9rem;padding-left:8px;}

/* ── FILTER PANEL ── */
.stExpander{border:1px solid rgba(255,255,255,0.05) !important;
            border-left:2px solid #00C896 !important;
            border-radius:0 6px 6px 0 !important;background:#0B0F1A !important;}
.stExpander summary p{font-family:'JetBrains Mono',monospace !important;
    font-size:9px !important;color:#00C896 !important;letter-spacing:.15em !important;text-transform:uppercase !important;}

/* ── WIDGETS ── */
div[data-baseweb="select"]>div,div[data-baseweb="input"]>div{
    background:#0B0F1A !important;border-color:rgba(255,255,255,0.06) !important;
    border-radius:4px !important;font-family:'Outfit',sans-serif !important;font-size:12.5px !important;}
div[data-baseweb="select"]>div:hover,div[data-baseweb="input"]>div:hover{border-color:rgba(0,200,150,.35) !important;}
label[data-testid="stWidgetLabel"]{font-family:'JetBrains Mono',monospace !important;
    font-size:8.5px !important;color:#2D3748 !important;letter-spacing:.12em !important;text-transform:uppercase !important;}
.stMultiSelect [data-baseweb="tag"]{background:rgba(0,200,150,.1) !important;
    border:1px solid rgba(0,200,150,.25) !important;border-radius:2px !important;}
button[data-testid="stBaseButton-secondary"]{background:#0B0F1A !important;
    border:1px solid rgba(0,200,150,.25) !important;color:#00C896 !important;
    border-radius:4px !important;font-family:'JetBrains Mono',monospace !important;
    font-size:10px !important;letter-spacing:.05em !important;transition:all .15s !important;}
button[data-testid="stBaseButton-secondary"]:hover{background:rgba(0,200,150,.08) !important;border-color:#00C896 !important;}
.stTabs [data-baseweb="tab-list"]{background:transparent !important;border-bottom:1px solid rgba(255,255,255,0.05) !important;gap:0;}
.stTabs [data-baseweb="tab"]{font-family:'JetBrains Mono',monospace !important;font-size:9px !important;
    letter-spacing:.14em !important;text-transform:uppercase !important;color:#2D3748 !important;
    padding:.6rem 1.5rem !important;border-bottom:2px solid transparent !important;background:transparent !important;}
.stTabs [aria-selected="true"]{color:#00C896 !important;border-bottom-color:#00C896 !important;}
div[data-testid="metric-container"]{background:#0B0F1A !important;
    border:1px solid rgba(255,255,255,0.05) !important;border-radius:6px !important;padding:.8rem 1rem !important;}
div[data-testid="stMetricValue"]{font-family:'Syne',sans-serif !important;font-size:20px !important;font-weight:700 !important;color:#F1F5F9 !important;}
div[data-testid="stMetricLabel"]{font-family:'JetBrains Mono',monospace !important;font-size:8.5px !important;letter-spacing:.1em !important;color:#2D3748 !important;text-transform:uppercase !important;}
div[data-testid="stMetricDelta"]{font-family:'JetBrains Mono',monospace !important;font-size:9.5px !important;}
.stSuccess{border-left:3px solid #00C896 !important;background:rgba(0,200,150,.06) !important;border-radius:3px !important;}
.stInfo{border-left:3px solid #4C6EF5 !important;background:rgba(76,110,245,.06) !important;border-radius:3px !important;}
.stWarning{border-left:3px solid #F6B93B !important;background:rgba(246,185,59,.06) !important;border-radius:3px !important;}
::-webkit-scrollbar{width:4px;height:4px;}
::-webkit-scrollbar-track{background:#06080F;}
::-webkit-scrollbar-thumb{background:rgba(0,200,150,.2);border-radius:2px;}

/* ── DIVIDER ── */
.div{height:1px;background:rgba(255,255,255,0.04);margin:1.4rem 0;}

/* ── STATS ROW ── */
.stats-row{display:flex;gap:1px;background:rgba(255,255,255,0.04);border-radius:4px;overflow:hidden;margin-bottom:1rem;}
.stats-cell{flex:1;background:#0B0F1A;padding:.6rem 1rem;}
.stats-lbl{font-family:'JetBrains Mono',monospace;font-size:8px;color:#2D3748;letter-spacing:.1em;text-transform:uppercase;margin-bottom:3px;}
.stats-val{font-family:'Syne',sans-serif;font-size:13px;font-weight:600;color:#E2E8F0;}

/* ── INSIGHT BADGE ── */
.insight{display:inline-flex;align-items:center;gap:6px;
         background:#0B0F1A;border:1px solid rgba(255,255,255,0.06);
         border-radius:3px;padding:4px 10px;margin:2px;
         font-family:'JetBrains Mono',monospace;font-size:9px;color:#4A5568;}
.insight strong{color:#E2E8F0;font-weight:500;}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
BG   = "#0B0F1A"
GRID = "rgba(255,255,255,0.04)"
TXT  = "#4A5568"
TXT2 = "#718096"
WHT  = "#F1F5F9"

C_GREEN  = "#00C896"
C_RED    = "#F75F7A"
C_BLUE   = "#4C6EF5"
C_AMBER  = "#F6B93B"
C_TEAL   = "#22D3EE"
C_PURPLE = "#A78BFA"
C_GRAY   = "#2D3748"

BASE = dict(
    paper_bgcolor=BG, plot_bgcolor=BG,
    font=dict(family="JetBrains Mono", color=TXT, size=10),
    margin=dict(l=0, r=10, t=4, b=0),
)

PM_MAP = {"mode1":"Full Payment","mode2":"Instalment","mode3":"EMI","mode4":"Partial","mode13":"Scholarship"}

def fmt(v):
    if pd.isna(v) or v==0: return "₹0"
    if v>=1e7: return f"₹{v/1e7:.2f}Cr"
    if v>=1e5: return f"₹{v/1e5:.1f}L"
    return f"₹{v:,.0f}"

def ch(t):
    return f'<div class="ctitle">{t}</div>'

def kpi(lbl, val, sub, accent="g", sub_cls="n", small=False):
    sm = " sm" if small else ""
    return f"""<div class="kpi-cell {accent}">
<div class="kpi-lbl">{lbl}</div>
<div class="kpi-val{sm}">{val}</div>
<div class="kpi-sub {sub_cls}">{sub}</div></div>"""

# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
if "master" not in st.session_state: st.session_state.master = None
if "loaded" not in st.session_state: st.session_state.loaded = False

# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADER
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def load(sb, cb):
    sem  = pd.read_csv(io.BytesIO(sb))
    conv = pd.read_excel(io.BytesIO(cb))
    sem.columns = sem.columns.str.strip()
    for c in ['Is Attended ?','Is Converted ?','Session','TRADER',
              'Is our Student ?','Trainer / Presenter','Place']:
        sem[c] = sem[c].astype(str).str.strip().str.upper()
    sem['Seminar Date'] = pd.to_datetime(sem['Seminar Date'], errors='coerce', dayfirst=True)
    sem['Amount Paid']  = pd.to_numeric(sem['Amount Paid'], errors='coerce').fillna(0)
    sem['mob']          = sem['Mobile'].astype(str).str.replace(r'\D','',regex=True).str[-10:]

    att = sem[sem['Is Attended ?']=='YES'].copy().reset_index(drop=True)
    att['Conv Status'] = att['Is Converted ?'].apply(
        lambda x: 'Converted' if x in ['CONVERTED','YES'] else 'Not Converted')

    for c in ['payment_received','total_amount','total_due','total_gst']:
        conv[c] = pd.to_numeric(conv[c], errors='coerce').fillna(0)
    conv['order_date']   = pd.to_datetime(conv['order_date'], errors='coerce', utc=True).dt.tz_localize(None)
    conv['phone_clean']  = conv['phone'].astype(str).str.replace(r'\D','',regex=True).str[-10:]
    conv['PM Label']     = conv['payment_mode'].map(PM_MAP).fillna(conv['payment_mode'])
    conv['service_name'] = conv['service_name'].astype(str).str.strip()
    conv['sales_rep_name']= conv['sales_rep_name'].astype(str).str.strip()
    conv['trainer_name'] = conv['trainer'].astype(str).str.split(' - ').str[-1].str.strip()

    mg = att.merge(
        conv[['phone_clean','orderID','order_date','service_code','service_name',
              'payment_received','total_amount','total_due','total_gst',
              'payment_mode','PM Label','status','sales_rep_name',
              'trainer','trainer_name','student_invid','batch_date']],
        left_on='mob', right_on='phone_clean', how='left'
    )
    def due_tag(r):
        if pd.isna(r.get('total_due')): return 'No Order'
        if r['total_due'] <= 0:          return 'Fully Paid'
        if r['total_amount']>0 and r['total_due']<r['total_amount']: return 'Partially Paid'
        return 'Fully Due'
    mg['Due Status'] = mg.apply(due_tag, axis=1)
    return mg

# ══════════════════════════════════════════════════════════════════════════════
# TOP BAR
# ══════════════════════════════════════════════════════════════════════════════
now = datetime.now()
st.markdown(f"""
<div class="topbar">
  <div style="display:flex;align-items:center">
    <div class="logo">◈ SEMINAR<em>IQ</em></div>
    <div class="topbar-sep"></div>
    <div class="topbar-sub">INTELLIGENCE COMMAND CENTER</div>
  </div>
  <div class="topbar-right">
    <div class="dot"></div>
    <div class="live">LIVE</div>
    <div class="clock">{now.strftime("%d %b %Y  %H:%M")}</div>
  </div>
</div>
<div class="page">
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
t_up, t_dash, t_rec = st.tabs(["◈  DATA UPLOAD", "▣  ANALYTICS", "≡  RECORDS & EXPORT"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
with t_up:
    st.markdown("""<div class="sh"><div class="sh-label">Data Ingestion</div><div class="sh-line"></div></div>""", unsafe_allow_html=True)
    st.markdown("""
    <div style="background:#0B0F1A;border:1px solid rgba(0,200,150,0.1);border-radius:6px;padding:1.1rem 1.4rem;margin-bottom:1.4rem;">
      <div style="font-family:'Outfit',sans-serif;font-size:12px;color:#4A5568;line-height:1.9;">
        Upload both files. The system joins on <span style="color:#00C896;font-family:'JetBrains Mono',monospace;">phone number</span>
        to enrich every attendee row with their course, payment, due amount, sales rep and trainer from the conversion list.
      </div>
    </div>""", unsafe_allow_html=True)

    uc1, uc2 = st.columns(2, gap="medium")
    with uc1:
        st.markdown("""<div style="font-family:'JetBrains Mono',monospace;font-size:8.5px;color:#00C896;letter-spacing:.15em;text-transform:uppercase;margin-bottom:6px;">◈ FILE 01 — Seminar Updated Sheet (.csv)</div>""", unsafe_allow_html=True)
        f_sem = st.file_uploader("sem", type=["csv"], key="up_sem", label_visibility="collapsed")
    with uc2:
        st.markdown("""<div style="font-family:'JetBrains Mono',monospace;font-size:8.5px;color:#4C6EF5;letter-spacing:.15em;text-transform:uppercase;margin-bottom:6px;">◈ FILE 02 — Conversion List (.xlsx)</div>""", unsafe_allow_html=True)
        f_conv = st.file_uploader("conv", type=["xlsx","xls"], key="up_conv", label_visibility="collapsed")

    if f_sem and f_conv:
        with st.spinner("Merging files…"):
            mg = load(f_sem.read(), f_conv.read())
        st.session_state.master = mg
        st.session_state.loaded = True
        total=len(mg); matched=mg['orderID'].notna().sum(); conv_n=(mg['Conv Status']=='Converted').sum()
        st.success(f"✓  {total:,} attendees · {matched:,} matched to orders · {conv_n:,} converted")
        m1,m2,m3,m4,m5 = st.columns(5)
        m1.metric("Total Attendees",   f"{total:,}")
        m2.metric("Matched to Orders", f"{matched:,}", f"{matched/total*100:.1f}%")
        m3.metric("Converted",         f"{conv_n:,}",  f"{conv_n/total*100:.1f}%")
        m4.metric("Locations",         f"{mg['Place'].nunique()}")
        m5.metric("Seminar Dates",     f"{mg['Seminar Date'].nunique()}")
    elif f_sem or f_conv:
        st.info("Waiting for: " + ("Conversion List" if f_sem else "Seminar Sheet"))

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — ANALYTICS
# ══════════════════════════════════════════════════════════════════════════════
with t_dash:
    if not st.session_state.loaded:
        st.markdown("""<div style="padding:3rem;text-align:center;color:#1A202C;font-family:'JetBrains Mono',monospace;font-size:11px;letter-spacing:.1em;">◈  UPLOAD FILES FIRST</div>""", unsafe_allow_html=True)
    else:
        dfa = st.session_state.master.copy()

        # ── FILTERS ───────────────────────────────────────────────────────────
        with st.expander("▼  FILTERS", expanded=True):
            fa,fb,fc,fd = st.columns(4)
            fe,ff,fg,fh = st.columns(4)
            fi,fj,fk,_  = st.columns(4)
            sel_place   = fa.multiselect("Location",          sorted(dfa['Place'].dropna().unique()),                          key="d_pl")
            sel_trainer = fb.multiselect("Trainer",           sorted(dfa['Trainer / Presenter'].dropna().unique()),            key="d_tr")
            sel_date    = fc.multiselect("Seminar Date",      [d.strftime("%d %b %Y") for d in sorted(dfa['Seminar Date'].dropna().unique())], key="d_dt")
            sel_sess    = fd.multiselect("Session",           ["MORNING","EVENING"],                                          key="d_se")
            sel_conv    = fe.multiselect("Conversion Status", ["Converted","Not Converted"],                                  key="d_cv")
            sel_due     = ff.multiselect("Due Status",        sorted(dfa['Due Status'].dropna().unique()),                    key="d_du")
            sel_course  = fg.multiselect("Course",            sorted(dfa['service_name'].dropna().unique()),                  key="d_co")
            sel_rep     = fh.multiselect("Sales Rep",         sorted(dfa['sales_rep_name'].dropna().unique()),                key="d_re")
            sel_status  = fi.multiselect("Order Status",      sorted(dfa['status'].dropna().unique()),                       key="d_st")
            sel_trader  = fj.selectbox("Is Trader?",          ["All","Yes","No"],                                            key="d_td")
            sel_ourstu  = fk.selectbox("Existing Student?",   ["All","Yes","No"],                                            key="d_os")

        df = dfa.copy()
        if sel_place:   df = df[df['Place'].isin(sel_place)]
        if sel_trainer: df = df[df['Trainer / Presenter'].isin(sel_trainer)]
        if sel_date:
            fmts = [pd.to_datetime(d, format="%d %b %Y") for d in sel_date]
            df = df[df['Seminar Date'].isin(fmts)]
        if sel_sess:    df = df[df['Session'].isin(sel_sess)]
        if sel_conv:    df = df[df['Conv Status'].isin(sel_conv)]
        if sel_due:     df = df[df['Due Status'].isin(sel_due)]
        if sel_course:  df = df[df['service_name'].isin(sel_course)]
        if sel_rep:     df = df[df['sales_rep_name'].isin(sel_rep)]
        if sel_status:  df = df[df['status'].isin(sel_status)]
        if sel_trader!="All": df = df[df['TRADER'].isin(['YES','Y','TYES'] if sel_trader=="Yes" else ['NO','N'])]
        if sel_ourstu!="All": df = df[df['Is our Student ?'].isin(['YES','STUDENT'] if sel_ourstu=="Yes" else ['NO'])]

        N      = len(df)
        conv_n = (df['Conv Status']=='Converted').sum()
        nconv  = N - conv_n
        conv_r = conv_n/N*100 if N else 0
        rcvd   = df['payment_received'].sum()
        due    = df['total_due'].sum()
        rev    = df['total_amount'].sum()
        fp     = (df['Due Status']=='Fully Paid').sum()
        hd     = df[df['Due Status'].isin(['Partially Paid','Fully Due'])].shape[0]

        # ── KPI STRIP ─────────────────────────────────────────────────────────
        st.markdown("""<div class="sh"><div class="sh-label">Key Metrics</div><div class="sh-line"></div></div>""", unsafe_allow_html=True)
        st.markdown(f"""<div class="kpi-strip">
          {kpi("TOTAL ATTENDED",  f"{N:,}",          f"of {len(dfa):,} total",            "b","n")}
          {kpi("CONVERTED",       f"{conv_n:,}",     f"↑ {conv_r:.1f}% rate",              "g","g")}
          {kpi("NOT CONVERTED",   f"{nconv:,}",      f"{100-conv_r:.1f}% of attended",     "r","r")}
          {kpi("GROSS REVENUE",   fmt(rev),          "Total order value",                  "p","n", True)}
          {kpi("COLLECTED",       fmt(rcvd),         f"Due outstanding: {fmt(due)}",       "t","g", True)}
          {kpi("FULLY PAID",      f"{fp:,}",         f"Has balance due: {hd:,}",           "a","a")}
        </div>""", unsafe_allow_html=True)

        # ══════════════════════════════════════════════════════════════════════
        # SECTION A — ATTENDANCE & CONVERSION
        # Chart choices:
        # 1. Grouped bar — Location: compare Attended vs Converted counts side by side
        # 2. Line + Area — Seminar Date trend: shows change over time correctly
        # 3. Grouped bar — Session Morning vs Evening: two categories, two measures
        # ══════════════════════════════════════════════════════════════════════
        st.markdown("""<div class="sh"><div class="sh-label">Attendance & Conversion Analysis</div><div class="sh-line"></div></div>""", unsafe_allow_html=True)

        a1, a2 = st.columns([5, 4], gap="small")

        with a1:
            # GROUPED BAR — Location: attended vs converted (correct for comparing two measures across categories)
            loc = df.groupby('Place').agg(
                Attended=('NAME','count'),
                Converted=('Conv Status', lambda x:(x=='Converted').sum())
            ).reset_index().sort_values('Attended', ascending=False)
            loc['Conv Rate'] = (loc['Converted']/loc['Attended']*100).round(1)

            fig = go.Figure()
            fig.add_trace(go.Bar(
                name='Attended', x=loc['Place'], y=loc['Attended'],
                marker_color=C_BLUE, marker_line_width=0, opacity=0.55,
                hovertemplate='<b>%{x}</b><br>Attended: %{y}<extra></extra>'
            ))
            fig.add_trace(go.Bar(
                name='Converted', x=loc['Place'], y=loc['Converted'],
                marker_color=C_GREEN, marker_line_width=0,
                hovertemplate='<b>%{x}</b><br>Converted: %{y}<extra></extra>'
            ))
            fig.update_layout(**BASE, height=270, barmode='group', bargap=0.25, bargroupgap=0.06,
                legend=dict(orientation="h", x=0, y=1.08, font=dict(size=9), bgcolor="rgba(0,0,0,0)"),
                xaxis=dict(showgrid=False, tickangle=-30, tickfont=dict(size=8.5, color=TXT2)),
                yaxis=dict(gridcolor=GRID, tickfont=dict(size=9, color=C_GRAY), title=None))
            st.markdown(ch("ATTENDED vs CONVERTED — BY LOCATION"), unsafe_allow_html=True)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

        with a2:
            # HORIZONTAL BAR — Conversion Rate % per location (single metric, ranked)
            loc_s = loc.sort_values('Conv Rate')
            colors = [C_RED if r < 10 else (C_AMBER if r < 20 else C_GREEN) for r in loc_s['Conv Rate']]

            fig2 = go.Figure(go.Bar(
                x=loc_s['Conv Rate'], y=loc_s['Place'], orientation='h',
                marker_color=colors, marker_line_width=0,
                text=[f"{r}%" for r in loc_s['Conv Rate']],
                textposition='outside', textfont=dict(size=9, color=TXT2),
                hovertemplate='<b>%{y}</b><br>Conv Rate: %{x:.1f}%<extra></extra>'
            ))
            fig2.add_vline(x=loc_s['Conv Rate'].mean(), line_dash="dot",
                           line_color=C_AMBER, line_width=1,
                           annotation_text=f"avg {loc_s['Conv Rate'].mean():.1f}%",
                           annotation_font=dict(size=8, color=C_AMBER),
                           annotation_position="top right")
            fig2.update_layout(**BASE, height=270, bargap=0.3,
                xaxis=dict(showgrid=False, showticklabels=False, zeroline=False, range=[0, loc_s['Conv Rate'].max()*1.22]),
                yaxis=dict(showgrid=False, tickfont=dict(size=8.5, color=TXT2)))
            st.markdown(ch("CONVERSION RATE % — BY LOCATION  (red < 10%  ·  amber < 20%  ·  green ≥ 20%)"), unsafe_allow_html=True)
            st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar":False})

        # ══════════════════════════════════════════════════════════════════════
        # SECTION B — TIME SERIES + SESSION + FUNNEL
        # Chart choices:
        # 1. Dual-axis line/area — Date trend: time series needs lines not bars
        # 2. Stacked bar — Session: part-of-whole across two categories
        # 3. Horizontal funnel — correct for pipeline stages
        # ══════════════════════════════════════════════════════════════════════
        st.markdown("""<div class="div"></div>""", unsafe_allow_html=True)
        b1, b2, b3 = st.columns([5, 2, 2], gap="small")

        with b1:
            # AREA + LINE — time series (correct for showing trend over dates)
            date_g = df.groupby('Seminar Date').agg(
                Attended=('NAME','count'),
                Converted=('Conv Status', lambda x:(x=='Converted').sum())
            ).reset_index().dropna(subset=['Seminar Date']).sort_values('Seminar Date')
            date_g['Label'] = date_g['Seminar Date'].dt.strftime("%d %b")
            date_g['Conv Rate'] = (date_g['Converted']/date_g['Attended']*100).round(1)

            fig3 = make_subplots(specs=[[{"secondary_y":True}]])
            fig3.add_trace(go.Scatter(
                x=date_g['Label'], y=date_g['Attended'],
                name='Attended', mode='lines+markers',
                line=dict(color=C_BLUE, width=2),
                marker=dict(size=5, color=C_BLUE, line=dict(color=BG, width=1.5)),
                fill='tozeroy', fillcolor='rgba(76,110,245,0.07)',
                hovertemplate='%{x}<br>Attended: %{y}<extra></extra>'
            ), secondary_y=False)
            fig3.add_trace(go.Scatter(
                x=date_g['Label'], y=date_g['Conv Rate'],
                name='Conv Rate %', mode='lines+markers',
                line=dict(color=C_GREEN, width=2, dash='solid'),
                marker=dict(size=5, color=C_GREEN, line=dict(color=BG, width=1.5)),
                hovertemplate='%{x}<br>Conv Rate: %{y:.1f}%<extra></extra>'
            ), secondary_y=True)
            fig3.update_layout(**BASE, height=230,
                legend=dict(orientation="h", x=0, y=1.1, font=dict(size=9), bgcolor="rgba(0,0,0,0)"),
                xaxis=dict(showgrid=False, tickfont=dict(size=9, color=TXT2)),
                yaxis=dict(gridcolor=GRID, tickfont=dict(size=9, color=C_BLUE), title=None, zeroline=False),
                yaxis2=dict(showgrid=False, tickfont=dict(size=9, color=C_GREEN), ticksuffix="%", title=None))
            st.markdown(ch("ATTENDANCE TREND & CONVERSION RATE — BY SEMINAR DATE"), unsafe_allow_html=True)
            st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar":False})

        with b2:
            # STACKED BAR — session (shows part-of-whole + absolute counts)
            sess = df.groupby('Session').agg(
                Attended=('NAME','count'),
                Converted=('Conv Status', lambda x:(x=='Converted').sum())
            ).reset_index()
            sess = sess[sess['Session'].isin(['MORNING','EVENING'])]
            sess['Not Conv'] = sess['Attended'] - sess['Converted']

            fig4 = go.Figure()
            fig4.add_trace(go.Bar(
                name='Converted',     x=sess['Session'], y=sess['Converted'],
                marker_color=C_GREEN, marker_line_width=0,
                hovertemplate='%{x}<br>Converted: %{y}<extra></extra>'
            ))
            fig4.add_trace(go.Bar(
                name='Not Converted', x=sess['Session'], y=sess['Not Conv'],
                marker_color='rgba(76,110,245,0.3)', marker_line_width=0,
                hovertemplate='%{x}<br>Not Converted: %{y}<extra></extra>'
            ))
            fig4.update_layout(**BASE, height=230, barmode='stack', bargap=0.4,
                legend=dict(orientation="h", x=0, y=1.1, font=dict(size=8.5), bgcolor="rgba(0,0,0,0)"),
                xaxis=dict(showgrid=False, tickfont=dict(size=9, color=TXT2)),
                yaxis=dict(gridcolor=GRID, tickfont=dict(size=9, color=C_GRAY)))
            st.markdown(ch("MORNING vs EVENING — STACKED"), unsafe_allow_html=True)
            st.plotly_chart(fig4, use_container_width=True, config={"displayModeBar":False})

        with b3:
            # HORIZONTAL FUNNEL — pipeline stages (correct for stage-to-stage drop-off)
            fig5 = go.Figure(go.Funnel(
                y=["TOTAL REG", "ATTENDED", "CONVERTED"],
                x=[len(dfa), N, conv_n],
                textinfo="value+percent initial",
                textfont=dict(family="JetBrains Mono", size=9, color=WHT),
                marker=dict(color=[C_BLUE, C_PURPLE, C_GREEN],
                            line=dict(color=BG, width=2)),
                connector=dict(line=dict(color=GRID, width=1)),
            ))
            fig5.update_layout(**BASE, height=230)
            st.markdown(ch("CONVERSION FUNNEL"), unsafe_allow_html=True)
            st.plotly_chart(fig5, use_container_width=True, config={"displayModeBar":False})

        # ══════════════════════════════════════════════════════════════════════
        # SECTION C — PAYMENT & REVENUE
        # Chart choices:
        # 1. Donut — Due Status: part-of-whole with 4 categories, no time component
        # 2. Grouped bar — Revenue by location: two measures (collected + due)
        # 3. Horizontal bar — Payment mode: ranked single metric
        # ══════════════════════════════════════════════════════════════════════
        st.markdown("""<div class="sh"><div class="sh-label">Payment & Revenue Intelligence</div><div class="sh-line"></div></div>""", unsafe_allow_html=True)

        c1, c2, c3 = st.columns([2, 4, 3], gap="small")

        with c1:
            # DONUT — Due Status (correct for part-of-whole composition)
            due_c = df['Due Status'].value_counts()
            cmap  = {'Fully Paid':C_GREEN, 'Partially Paid':C_AMBER, 'Fully Due':C_RED, 'No Order':C_GRAY}
            fig6 = go.Figure(go.Pie(
                labels=due_c.index, values=due_c.values, hole=0.60,
                marker=dict(colors=[cmap.get(l, C_BLUE) for l in due_c.index],
                            line=dict(color=BG, width=3)),
                textinfo='percent', textfont=dict(size=9, family="JetBrains Mono"),
                hovertemplate='<b>%{label}</b><br>Count: %{value}<br>%{percent}<extra></extra>',
            ))
            fig6.update_layout(**BASE, height=220, showlegend=True,
                legend=dict(orientation="v", x=1.02, y=0.5, font=dict(size=8.5), bgcolor="rgba(0,0,0,0)"),
                annotations=[dict(text="DUE<br>STATUS", x=0.5, y=0.5, showarrow=False,
                                  font=dict(size=8, color=C_GRAY, family="JetBrains Mono"))])
            st.markdown(ch("PAYMENT DUE STATUS"), unsafe_allow_html=True)
            st.plotly_chart(fig6, use_container_width=True, config={"displayModeBar":False})

        with c2:
            # GROUPED BAR — Revenue collected vs due per location (two measures, compare across categories)
            rev_l = df.groupby('Place').agg(
                Collected=('payment_received','sum'), Due=('total_due','sum')
            ).reset_index().sort_values('Collected', ascending=False)

            fig7 = go.Figure()
            fig7.add_trace(go.Bar(
                name='Collected', x=rev_l['Place'], y=rev_l['Collected'],
                marker_color=C_GREEN, marker_line_width=0, opacity=0.85,
                hovertemplate='<b>%{x}</b><br>Collected: ₹%{y:,.0f}<extra></extra>'
            ))
            fig7.add_trace(go.Bar(
                name='Due', x=rev_l['Place'], y=rev_l['Due'],
                marker_color=C_RED, marker_line_width=0, opacity=0.85,
                hovertemplate='<b>%{x}</b><br>Due: ₹%{y:,.0f}<extra></extra>'
            ))
            fig7.update_layout(**BASE, height=220, barmode='group', bargap=0.25, bargroupgap=0.06,
                legend=dict(orientation="h", x=0, y=1.1, font=dict(size=9), bgcolor="rgba(0,0,0,0)"),
                xaxis=dict(showgrid=False, tickangle=-30, tickfont=dict(size=8.5, color=TXT2)),
                yaxis=dict(gridcolor=GRID, tickformat=",.0f", tickfont=dict(size=9, color=C_GRAY),
                           tickprefix="₹"))
            st.markdown(ch("COLLECTED vs DUE — BY LOCATION"), unsafe_allow_html=True)
            st.plotly_chart(fig7, use_container_width=True, config={"displayModeBar":False})

        with c3:
            # HORIZONTAL BAR — Payment mode: ranked single metric, easier to read labels
            pm = df[df['Conv Status']=='Converted']['PM Label'].value_counts()
            pm_colors = [C_GREEN, C_BLUE, C_PURPLE, C_AMBER, C_TEAL]
            fig8 = go.Figure(go.Bar(
                x=pm.values, y=pm.index, orientation='h',
                marker_color=pm_colors[:len(pm)], marker_line_width=0,
                text=pm.values, textposition='outside',
                textfont=dict(size=9, color=TXT2),
                hovertemplate='<b>%{y}</b><br>Count: %{x}<extra></extra>'
            ))
            fig8.update_layout(**BASE, height=220, bargap=0.35,
                xaxis=dict(showgrid=False, showticklabels=False, zeroline=False,
                           range=[0, pm.max()*1.22]),
                yaxis=dict(showgrid=False, tickfont=dict(size=9, color=TXT2)))
            st.markdown(ch("PAYMENT MODE — CONVERTED STUDENTS"), unsafe_allow_html=True)
            st.plotly_chart(fig8, use_container_width=True, config={"displayModeBar":False})

        # ══════════════════════════════════════════════════════════════════════
        # SECTION D — COURSE · SALES REP · TRAINER
        # Chart choices:
        # 1. Horizontal bar — Course by count: ranked single metric, long labels need horizontal
        # 2. Scatter — Sales rep: two metrics (conversions + revenue) → scatter reveals relationship
        # 3. Grouped bar — Trainer: attended vs converted, two measures
        # ══════════════════════════════════════════════════════════════════════
        st.markdown("""<div class="sh"><div class="sh-label">Course · Sales Representative · Trainer Performance</div><div class="sh-line"></div></div>""", unsafe_allow_html=True)

        d1, d2, d3 = st.columns(3, gap="small")

        with d1:
            # HORIZONTAL BAR — courses (long names → horizontal always better)
            cr = df[df['Conv Status']=='Converted'].groupby('service_name').agg(
                Count=('NAME','count'), Revenue=('total_amount','sum')
            ).reset_index().sort_values('Count')
            cr['Short'] = cr['service_name'].apply(lambda x: x[:30]+'…' if len(x)>30 else x)
            # Color encode by revenue magnitude
            max_rev = cr['Revenue'].max()
            cr['opacity'] = 0.4 + 0.6*(cr['Revenue']/max_rev) if max_rev > 0 else 0.7

            fig9 = go.Figure(go.Bar(
                x=cr['Count'], y=cr['Short'], orientation='h',
                marker=dict(
                    color=cr['Count'],
                    colorscale=[[0, C_PURPLE], [1, C_GREEN]],
                    line=dict(width=0),
                    showscale=False
                ),
                text=cr['Count'], textposition='outside',
                textfont=dict(size=9, color=TXT2),
                hovertemplate='<b>%{y}</b><br>Conversions: %{x}<br>Revenue: ₹%{customdata:,.0f}<extra></extra>',
                customdata=cr['Revenue']
            ))
            fig9.update_layout(**BASE, height=280, bargap=0.3,
                xaxis=dict(showgrid=False, showticklabels=False, zeroline=False,
                           range=[0, cr['Count'].max()*1.22]),
                yaxis=dict(showgrid=False, tickfont=dict(size=8.5, color=TXT2)))
            st.markdown(ch("TOP COURSES BY CONVERSIONS"), unsafe_allow_html=True)
            st.plotly_chart(fig9, use_container_width=True, config={"displayModeBar":False})

        with d2:
            # SCATTER PLOT — Sales rep: conversions (x) vs revenue (y), bubble = avg deal size
            # Scatter is correct here: reveals which reps do high-volume vs high-value
            rp = df[df['Conv Status']=='Converted'].groupby('sales_rep_name').agg(
                Conv=('NAME','count'), Rev=('total_amount','sum')
            ).reset_index()
            rp['Avg Deal'] = rp['Rev']/rp['Conv']
            rp = rp[rp['Conv'] > 0]

            fig10 = go.Figure(go.Scatter(
                x=rp['Conv'], y=rp['Rev'],
                mode='markers+text',
                marker=dict(
                    size=rp['Avg Deal']/800,
                    color=rp['Conv'],
                    colorscale=[[0,C_BLUE],[0.5,C_PURPLE],[1,C_GREEN]],
                    line=dict(color=BG, width=1.5),
                    sizemin=8, sizemode='diameter',
                ),
                text=rp['sales_rep_name'].apply(lambda x: x.split()[0]),
                textfont=dict(size=8, color=TXT2),
                textposition='top center',
                hovertemplate='<b>%{text}</b><br>Conversions: %{x}<br>Revenue: ₹%{y:,.0f}<br>Avg deal: ₹%{customdata:,.0f}<extra></extra>',
                customdata=rp['Avg Deal']
            ))
            # Average lines
            fig10.add_hline(y=rp['Rev'].mean(), line_dash="dot", line_color=C_AMBER, line_width=1,
                            annotation_text="avg rev", annotation_font=dict(size=7.5, color=C_AMBER))
            fig10.add_vline(x=rp['Conv'].mean(), line_dash="dot", line_color=C_AMBER, line_width=1,
                            annotation_text="avg conv", annotation_font=dict(size=7.5, color=C_AMBER))
            fig10.update_layout(**BASE, height=280,
                xaxis=dict(gridcolor=GRID, title=dict(text="Conversions", font=dict(size=8.5, color=TXT)),
                           tickfont=dict(size=8.5, color=C_GRAY), zeroline=False),
                yaxis=dict(gridcolor=GRID, title=dict(text="Revenue (₹)", font=dict(size=8.5, color=TXT)),
                           tickformat=",.0f", tickfont=dict(size=8.5, color=C_GRAY), zeroline=False))
            st.markdown(ch("SALES REP — CONVERSIONS vs REVENUE  (bubble = avg deal size)"), unsafe_allow_html=True)
            st.plotly_chart(fig10, use_container_width=True, config={"displayModeBar":False})

        with d3:
            # GROUPED BAR — Trainer: attended vs converted (correct for two-measure comparison)
            tr = df.groupby('Trainer / Presenter').agg(
                Attended=('NAME','count'),
                Converted=('Conv Status', lambda x:(x=='Converted').sum())
            ).reset_index()
            tr['Short'] = tr['Trainer / Presenter'].apply(lambda x: x[:22]+'…' if len(x)>22 else x)
            tr['Rate']  = (tr['Converted']/tr['Attended']*100).round(1)
            tr = tr.sort_values('Attended', ascending=True)

            fig11 = go.Figure()
            fig11.add_trace(go.Bar(
                name='Attended', y=tr['Short'], x=tr['Attended'],
                orientation='h', marker_color=C_BLUE, opacity=0.45, marker_line_width=0,
                hovertemplate='<b>%{y}</b><br>Attended: %{x}<extra></extra>'
            ))
            fig11.add_trace(go.Bar(
                name='Converted', y=tr['Short'], x=tr['Converted'],
                orientation='h', marker_color=C_GREEN, marker_line_width=0,
                hovertemplate='<b>%{y}</b><br>Converted: %{x}<extra></extra>'
            ))
            fig11.update_layout(**BASE, height=280, barmode='overlay', bargap=0.28,
                legend=dict(orientation="h", x=0, y=1.1, font=dict(size=9), bgcolor="rgba(0,0,0,0)"),
                xaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
                yaxis=dict(showgrid=False, tickfont=dict(size=8, color=TXT2)))
            st.markdown(ch("TRAINER — ATTENDED vs CONVERTED (overlap bars)"), unsafe_allow_html=True)
            st.plotly_chart(fig11, use_container_width=True, config={"displayModeBar":False})

        # ══════════════════════════════════════════════════════════════════════
        # SECTION E — PROFILE MIX
        # Chart choices:
        # 1. Donut — Trader mix: 2 categories, part of whole
        # 2. Donut — Existing student: 2 categories, part of whole
        # ══════════════════════════════════════════════════════════════════════
        st.markdown("""<div class="sh"><div class="sh-label">Attendee Profile</div><div class="sh-line"></div></div>""", unsafe_allow_html=True)

        e1, e2, e3, e4 = st.columns(4, gap="small")

        with e1:
            trader = df['TRADER'].map(lambda x:'Trader' if x in ['YES','Y','TYES'] else 'Non-Trader').value_counts()
            fig12 = go.Figure(go.Pie(
                labels=trader.index, values=trader.values, hole=0.62,
                marker=dict(colors=[C_TEAL, C_PURPLE], line=dict(color=BG, width=3)),
                textinfo='label+percent', textfont=dict(size=9, family="JetBrains Mono"),
            ))
            fig12.update_layout(**BASE, height=200, showlegend=False,
                annotations=[dict(text="TRADER<br>MIX", x=0.5, y=0.5, showarrow=False,
                                  font=dict(size=8, color=C_GRAY, family="JetBrains Mono"))])
            st.markdown(ch("TRADER vs NON-TRADER"), unsafe_allow_html=True)
            st.plotly_chart(fig12, use_container_width=True, config={"displayModeBar":False})

        with e2:
            stu = df['Is our Student ?'].map(
                lambda x:'Existing' if x in ['YES','STUDENT'] else 'New Lead'
            ).value_counts()
            fig13 = go.Figure(go.Pie(
                labels=stu.index, values=stu.values, hole=0.62,
                marker=dict(colors=[C_AMBER, C_BLUE], line=dict(color=BG, width=3)),
                textinfo='label+percent', textfont=dict(size=9, family="JetBrains Mono"),
            ))
            fig13.update_layout(**BASE, height=200, showlegend=False,
                annotations=[dict(text="STUDENT<br>MIX", x=0.5, y=0.5, showarrow=False,
                                  font=dict(size=8, color=C_GRAY, family="JetBrains Mono"))])
            st.markdown(ch("EXISTING STUDENT vs NEW LEAD"), unsafe_allow_html=True)
            st.plotly_chart(fig13, use_container_width=True, config={"displayModeBar":False})

        with e3:
            # HORIZONTAL BAR — Conversion rate: trader vs non-trader
            t_conv = df.copy()
            t_conv['Trader Label'] = t_conv['TRADER'].map(lambda x:'Trader' if x in ['YES','Y','TYES'] else 'Non-Trader')
            tc = t_conv.groupby('Trader Label').agg(
                Attended=('NAME','count'),
                Converted=('Conv Status', lambda x:(x=='Converted').sum())
            ).reset_index()
            tc['Rate'] = (tc['Converted']/tc['Attended']*100).round(1)

            fig14 = go.Figure(go.Bar(
                x=tc['Rate'], y=tc['Trader Label'], orientation='h',
                marker_color=[C_TEAL, C_PURPLE], marker_line_width=0,
                text=[f"{r}%" for r in tc['Rate']], textposition='outside',
                textfont=dict(size=11, color=TXT2),
            ))
            fig14.update_layout(**BASE, height=200, bargap=0.45,
                xaxis=dict(showgrid=False, showticklabels=False, zeroline=False, range=[0,tc['Rate'].max()*1.3]),
                yaxis=dict(showgrid=False, tickfont=dict(size=10, color=TXT2)))
            st.markdown(ch("CONVERSION RATE — TRADER vs NON-TRADER"), unsafe_allow_html=True)
            st.plotly_chart(fig14, use_container_width=True, config={"displayModeBar":False})

        with e4:
            # HORIZONTAL BAR — Conversion rate: existing student vs new
            sc = t_conv.copy()
            sc['Stu Label'] = sc['Is our Student ?'].map(
                lambda x:'Existing' if x in ['YES','STUDENT'] else 'New Lead')
            sc2 = sc.groupby('Stu Label').agg(
                Attended=('NAME','count'),
                Converted=('Conv Status', lambda x:(x=='Converted').sum())
            ).reset_index()
            sc2['Rate'] = (sc2['Converted']/sc2['Attended']*100).round(1)

            fig15 = go.Figure(go.Bar(
                x=sc2['Rate'], y=sc2['Stu Label'], orientation='h',
                marker_color=[C_AMBER, C_BLUE], marker_line_width=0,
                text=[f"{r}%" for r in sc2['Rate']], textposition='outside',
                textfont=dict(size=11, color=TXT2),
            ))
            fig15.update_layout(**BASE, height=200, bargap=0.45,
                xaxis=dict(showgrid=False, showticklabels=False, zeroline=False, range=[0, sc2['Rate'].max()*1.3]),
                yaxis=dict(showgrid=False, tickfont=dict(size=10, color=TXT2)))
            st.markdown(ch("CONVERSION RATE — EXISTING vs NEW LEAD"), unsafe_allow_html=True)
            st.plotly_chart(fig15, use_container_width=True, config={"displayModeBar":False})

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — RECORDS
# ══════════════════════════════════════════════════════════════════════════════
with t_rec:
    if not st.session_state.loaded:
        st.markdown("""<div style="padding:3rem;text-align:center;color:#1A202C;font-family:'JetBrains Mono',monospace;font-size:11px;">◈  UPLOAD FILES FIRST</div>""", unsafe_allow_html=True)
    else:
        dfa = st.session_state.master.copy()
        st.markdown("""<div class="sh"><div class="sh-label">Search & Filter Records</div><div class="sh-line"></div></div>""", unsafe_allow_html=True)

        tf1,tf2,tf3,tf4 = st.columns(4)
        tf5,tf6,tf7,tf8 = st.columns(4)
        tf9,tf10,tf11,tf12 = st.columns(4)

        t_place   = tf1.multiselect("Location",          sorted(dfa['Place'].dropna().unique()),                               key="r_pl")
        t_trainer = tf2.multiselect("Trainer",           sorted(dfa['Trainer / Presenter'].dropna().unique()),                 key="r_tr")
        t_date    = tf3.multiselect("Seminar Date",      [d.strftime("%d %b %Y") for d in sorted(dfa['Seminar Date'].dropna().unique())], key="r_dt")
        t_sess    = tf4.multiselect("Session",           ["MORNING","EVENING"],                                               key="r_se")
        t_conv    = tf5.multiselect("Conversion Status", ["Converted","Not Converted"],                                       key="r_cv")
        t_due     = tf6.multiselect("Due Status",        sorted(dfa['Due Status'].dropna().unique()),                         key="r_du")
        t_course  = tf7.multiselect("Course",            sorted(dfa['service_name'].dropna().unique()),                       key="r_co")
        t_rep     = tf8.multiselect("Sales Rep",         sorted(dfa['sales_rep_name'].dropna().unique()),                     key="r_re")
        t_status  = tf9.multiselect("Order Status",      sorted(dfa['status'].dropna().unique()),                             key="r_st")
        t_trader  = tf10.selectbox("Is Trader?",         ["All","Yes","No"],                                                  key="r_td")
        t_ourstu  = tf11.selectbox("Existing Student?",  ["All","Yes","No"],                                                  key="r_os")
        t_search  = tf12.text_input("Search Name / Phone / Email", placeholder="Type to search…",                            key="r_sr")

        dt = dfa.copy()
        if t_place:   dt = dt[dt['Place'].isin(t_place)]
        if t_trainer: dt = dt[dt['Trainer / Presenter'].isin(t_trainer)]
        if t_date:
            fmts = [pd.to_datetime(d, format="%d %b %Y") for d in t_date]
            dt = dt[dt['Seminar Date'].isin(fmts)]
        if t_sess:    dt = dt[dt['Session'].isin(t_sess)]
        if t_conv:    dt = dt[dt['Conv Status'].isin(t_conv)]
        if t_due:     dt = dt[dt['Due Status'].isin(t_due)]
        if t_course:  dt = dt[dt['service_name'].isin(t_course)]
        if t_rep:     dt = dt[dt['sales_rep_name'].isin(t_rep)]
        if t_status:  dt = dt[dt['status'].isin(t_status)]
        if t_trader!="All": dt = dt[dt['TRADER'].isin(['YES','Y','TYES'] if t_trader=="Yes" else ['NO','N'])]
        if t_ourstu!="All": dt = dt[dt['Is our Student ?'].isin(['YES','STUDENT'] if t_ourstu=="Yes" else ['NO'])]
        if t_search:
            mask = pd.Series([False]*len(dt))
            for c in ['NAME','Mobile','email']:
                if c in dt.columns:
                    mask = mask | dt[c].astype(str).str.lower().str.contains(t_search.lower(), na=False)
            dt = dt[mask]

        N_t=len(dt); conv_t=(dt['Conv Status']=='Converted').sum()
        rcvd_t=dt['payment_received'].sum(); due_t=dt['total_due'].sum()
        st.markdown(f"""
        <div class="stats-row">
          <div class="stats-cell"><div class="stats-lbl">Records</div><div class="stats-val">{N_t:,}</div></div>
          <div class="stats-cell"><div class="stats-lbl">Converted</div><div class="stats-val" style="color:#00C896">{conv_t:,}</div></div>
          <div class="stats-cell"><div class="stats-lbl">Not Converted</div><div class="stats-val" style="color:#F75F7A">{N_t-conv_t:,}</div></div>
          <div class="stats-cell"><div class="stats-lbl">Collected</div><div class="stats-val" style="color:#00C896">{fmt(rcvd_t)}</div></div>
          <div class="stats-cell"><div class="stats-lbl">Total Due</div><div class="stats-val" style="color:#F75F7A">{fmt(due_t)}</div></div>
          <div class="stats-cell"><div class="stats-lbl">Conv Rate</div><div class="stats-val">{conv_t/N_t*100:.1f}% if N_t else 0%</div></div>
        </div>""", unsafe_allow_html=True)

        disp = ['NAME','Mobile','Place','Seminar Date','Session','Trainer / Presenter',
                'Conv Status','Amount Paid','service_name','total_amount',
                'payment_received','total_due','Due Status','status',
                'sales_rep_name','trainer_name','TRADER','Is our Student ?','PM Label','Remarks']
        disp = [c for c in disp if c in dt.columns]
        ds = dt[disp].copy()
        ds['Seminar Date'] = ds['Seminar Date'].dt.strftime("%d %b %Y")

        st.dataframe(ds, use_container_width=True, height=460, hide_index=True,
            column_config={
                "total_amount":        st.column_config.NumberColumn("Course Amt",    format="₹%.0f"),
                "payment_received":    st.column_config.NumberColumn("Collected",     format="₹%.0f"),
                "total_due":           st.column_config.NumberColumn("Due",           format="₹%.0f"),
                "Amount Paid":         st.column_config.NumberColumn("Sem Paid",      format="₹%.0f"),
                "service_name":        st.column_config.TextColumn("Course"),
                "sales_rep_name":      st.column_config.TextColumn("Sales Rep"),
                "trainer_name":        st.column_config.TextColumn("Trainer (Order)"),
                "Trainer / Presenter": st.column_config.TextColumn("Seminar Trainer"),
                "PM Label":            st.column_config.TextColumn("Payment Mode"),
            })

        st.markdown("""<div class="sh" style="margin-top:1.5rem"><div class="sh-label">Location Summary</div><div class="sh-line"></div></div>""", unsafe_allow_html=True)
        loc_s = dt.groupby('Place').agg(
            Attended=('NAME','count'), Converted=('Conv Status',lambda x:(x=='Converted').sum()),
            Revenue=('total_amount','sum'), Collected=('payment_received','sum'), Due=('total_due','sum'),
        ).reset_index()
        loc_s['Conv Rate'] = (loc_s['Converted']/loc_s['Attended']*100).round(1).astype(str)+'%'
        loc_s['Revenue']   = loc_s['Revenue'].apply(fmt)
        loc_s['Collected'] = loc_s['Collected'].apply(fmt)
        loc_s['Due']       = loc_s['Due'].apply(fmt)
        st.dataframe(loc_s, use_container_width=True, hide_index=True, height=260)

        st.markdown("""<div class="sh" style="margin-top:1.5rem"><div class="sh-label">Export</div><div class="sh-line"></div></div>""", unsafe_allow_html=True)
        def to_xl(d):
            buf=io.BytesIO(); d.to_excel(buf,index=False); return buf.getvalue()
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        ex1,ex2,ex3,ex4 = st.columns(4)
        ex1.download_button("⬇  All Filtered",      to_xl(dt[disp]), file_name=f"filtered_{ts}.xlsx",    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        ex2.download_button("⬇  Converted Only",    to_xl(dt[dt['Conv Status']=='Converted'][disp]),     file_name=f"converted_{ts}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        ex3.download_button("⬇  Has Due Balance",   to_xl(dt[dt['total_due']>0][disp]),                  file_name=f"has_due_{ts}.xlsx",   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        ex4.download_button("⬇  Location Summary",  to_xl(loc_s),                                        file_name=f"loc_summary_{ts}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

st.markdown("</div>", unsafe_allow_html=True)
