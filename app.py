import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
from datetime import datetime

st.set_page_config(
    page_title="Seminar Intelligence — Command Center",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ══════════════════════════════════════════════════════════════════════════════
# DESIGN SYSTEM — Luxury Fintech / Bloomberg Terminal Aesthetic
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;500;600;700;800&family=JetBrains+Mono:wght@300;400;500&family=Outfit:wght@300;400;500;600&display=swap');

/* ── Reset & Base ── */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
html, body, [class*="css"], .stApp { background: #080C14 !important; font-family: 'Outfit', sans-serif; }
.block-container { padding: 0 !important; max-width: 100% !important; }
section[data-testid="stSidebar"] { display: none; }

/* ── Hide Streamlit Chrome ── */
#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }
div[data-testid="stToolbar"] { display: none; }

/* ── Main wrapper ── */
.cmd-wrap { padding: 0 2.2rem 3rem 2.2rem; }

/* ── Top Bar ── */
.topbar {
    display: flex; align-items: center; justify-content: space-between;
    padding: 1rem 2.2rem;
    background: #080C14;
    border-bottom: 1px solid rgba(0, 212, 170, 0.15);
    position: sticky; top: 0; z-index: 100;
}
.topbar-left { display: flex; align-items: center; gap: 14px; }
.topbar-logo {
    font-family: 'Syne', sans-serif; font-weight: 800; font-size: 17px;
    color: #fff; letter-spacing: -0.5px;
}
.topbar-logo span { color: #00D4AA; }
.topbar-sep { width: 1px; height: 22px; background: rgba(255,255,255,0.1); }
.topbar-sub { font-family: 'JetBrains Mono', monospace; font-size: 11px; color: #4A5568; letter-spacing: 0.08em; }
.topbar-right { display: flex; align-items: center; gap: 10px; }
.status-dot { width: 6px; height: 6px; border-radius: 50%; background: #00D4AA;
              box-shadow: 0 0 8px #00D4AA; animation: pulse 2s infinite; }
@keyframes pulse { 0%,100%{opacity:1;} 50%{opacity:0.4;} }
.status-label { font-family: 'JetBrains Mono', monospace; font-size: 10px; color: #00D4AA; letter-spacing: 0.1em; }
.topbar-time { font-family: 'JetBrains Mono', monospace; font-size: 10px; color: #2D3748; }

/* ── Section Headers ── */
.sec-header {
    display: flex; align-items: center; gap: 10px;
    margin: 2rem 0 1rem 0;
}
.sec-line { flex: 1; height: 1px; background: linear-gradient(90deg, rgba(0,212,170,0.3), transparent); }
.sec-title { font-family: 'Syne', sans-serif; font-size: 10px; font-weight: 600;
             color: #00D4AA; letter-spacing: 0.2em; text-transform: uppercase; white-space: nowrap; }
.sec-count { font-family: 'JetBrains Mono', monospace; font-size: 10px; color: #2D3748; }

/* ── KPI Cards ── */
.kpi-grid { display: grid; grid-template-columns: repeat(6, minmax(0,1fr)); gap: 1px;
            background: rgba(0,212,170,0.08); border: 1px solid rgba(0,212,170,0.12);
            border-radius: 4px; overflow: hidden; margin-bottom: 1px; }
.kpi-card {
    background: #0B1120;
    padding: 1.1rem 1.2rem;
    position: relative;
    transition: background 0.2s;
}
.kpi-card:hover { background: #0F1928; }
.kpi-card::after {
    content: ''; position: absolute; top: 0; left: 0; right: 0; height: 2px;
}
.kpi-card.accent-green::after  { background: #00D4AA; }
.kpi-card.accent-red::after    { background: #FF4D6D; }
.kpi-card.accent-amber::after  { background: #F6B93B; }
.kpi-card.accent-blue::after   { background: #4361EE; }
.kpi-card.accent-purple::after { background: #7B2FBE; }
.kpi-card.accent-teal::after   { background: #0EA5E9; }

.kpi-label {
    font-family: 'JetBrains Mono', monospace; font-size: 9px; font-weight: 400;
    color: #2D3748; letter-spacing: 0.12em; text-transform: uppercase; margin-bottom: 8px;
}
.kpi-value {
    font-family: 'Syne', sans-serif; font-size: 22px; font-weight: 700;
    color: #F7FAFC; line-height: 1; margin-bottom: 6px;
}
.kpi-value.small { font-size: 18px; }
.kpi-delta {
    font-family: 'JetBrains Mono', monospace; font-size: 10px; font-weight: 400;
}
.kpi-delta.up     { color: #00D4AA; }
.kpi-delta.down   { color: #FF4D6D; }
.kpi-delta.warn   { color: #F6B93B; }
.kpi-delta.neu    { color: #4A5568; }
.kpi-delta.blue   { color: #4361EE; }

/* ── Filter Panel ── */
.filter-panel {
    background: #0B1120;
    border: 1px solid rgba(255,255,255,0.05);
    border-left: 2px solid #00D4AA;
    border-radius: 0 4px 4px 0;
    padding: 1rem 1.4rem;
    margin-bottom: 1.5rem;
}
.filter-panel-title {
    font-family: 'JetBrains Mono', monospace; font-size: 9px; color: #00D4AA;
    letter-spacing: 0.18em; text-transform: uppercase; margin-bottom: 0.8rem;
}

/* ── Chart Cards ── */
.chart-card {
    background: #0B1120;
    border: 1px solid rgba(255,255,255,0.05);
    border-radius: 4px;
    padding: 1.2rem;
    position: relative;
    overflow: hidden;
}
.chart-card::before {
    content: ''; position: absolute; top: 0; left: 0; width: 3px; height: 100%;
    background: linear-gradient(180deg, #00D4AA, transparent);
}
.chart-title {
    font-family: 'JetBrains Mono', monospace; font-size: 9px; font-weight: 500;
    color: #4A5568; letter-spacing: 0.15em; text-transform: uppercase;
    margin-bottom: 0.9rem; padding-left: 8px;
}

/* ── Upload Zone ── */
.upload-zone {
    border: 1px dashed rgba(0,212,170,0.25);
    border-radius: 4px; background: #0B1120;
    padding: 2.5rem; text-align: center;
}
.upload-zone-title { font-family: 'Syne', sans-serif; font-size: 14px; font-weight: 600; color: #E2E8F0; margin-bottom: 6px; }
.upload-zone-sub   { font-family: 'JetBrains Mono', monospace; font-size: 11px; color: #2D3748; }

/* ── Data Table ── */
.stDataFrame { border-radius: 4px !important; }
div[data-testid="stDataFrame"] > div { border-radius: 4px !important; }

/* ── Streamlit widget overrides ── */
div[data-baseweb="select"] > div,
div[data-baseweb="input"] > div {
    background: #0B1120 !important;
    border-color: rgba(255,255,255,0.07) !important;
    border-radius: 3px !important;
    font-family: 'Outfit', sans-serif !important;
    font-size: 13px !important;
}
div[data-baseweb="select"] > div:hover,
div[data-baseweb="input"] > div:hover {
    border-color: rgba(0,212,170,0.4) !important;
}
label[data-testid="stWidgetLabel"] {
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 9px !important;
    color: #4A5568 !important;
    letter-spacing: 0.12em !important;
    text-transform: uppercase !important;
}
.stMultiSelect [data-baseweb="tag"] {
    background: rgba(0,212,170,0.15) !important;
    border: 1px solid rgba(0,212,170,0.3) !important;
    border-radius: 2px !important;
}
button[data-testid="stBaseButton-secondary"] {
    background: #0B1120 !important;
    border: 1px solid rgba(0,212,170,0.3) !important;
    color: #00D4AA !important;
    border-radius: 3px !important;
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 11px !important;
    letter-spacing: 0.05em !important;
    transition: all 0.15s !important;
}
button[data-testid="stBaseButton-secondary"]:hover {
    background: rgba(0,212,170,0.1) !important;
    border-color: #00D4AA !important;
}
.stFileUploader {
    border: 1px dashed rgba(0,212,170,0.2) !important;
    border-radius: 4px !important;
    background: #0B1120 !important;
    padding: 0.5rem !important;
}
.stTabs [data-baseweb="tab-list"] {
    background: transparent !important;
    border-bottom: 1px solid rgba(255,255,255,0.06) !important;
    gap: 0 !important;
}
.stTabs [data-baseweb="tab"] {
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 10px !important;
    letter-spacing: 0.12em !important;
    text-transform: uppercase !important;
    color: #4A5568 !important;
    padding: 0.6rem 1.4rem !important;
    border-bottom: 2px solid transparent !important;
    background: transparent !important;
}
.stTabs [aria-selected="true"] {
    color: #00D4AA !important;
    border-bottom-color: #00D4AA !important;
}
.stExpander {
    border: 1px solid rgba(255,255,255,0.05) !important;
    border-left: 2px solid rgba(0,212,170,0.5) !important;
    border-radius: 0 4px 4px 0 !important;
    background: #0B1120 !important;
}
.stExpander summary {
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 10px !important;
    letter-spacing: 0.1em !important;
    color: #00D4AA !important;
    text-transform: uppercase !important;
}
div[data-testid="metric-container"] {
    background: #0B1120 !important;
    border: 1px solid rgba(255,255,255,0.05) !important;
    border-radius: 4px !important;
    padding: 0.8rem 1rem !important;
}
div[data-testid="stMetricValue"] {
    font-family: 'Syne', sans-serif !important;
    font-size: 20px !important;
    font-weight: 700 !important;
    color: #F7FAFC !important;
}
div[data-testid="stMetricLabel"] {
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 9px !important;
    letter-spacing: 0.1em !important;
    color: #4A5568 !important;
    text-transform: uppercase !important;
}
div[data-testid="stMetricDelta"] {
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 10px !important;
}
.stSuccess, .stInfo, .stWarning {
    border-radius: 3px !important;
    font-family: 'Outfit', sans-serif !important;
    font-size: 13px !important;
}
.stSuccess { border-left: 3px solid #00D4AA !important; background: rgba(0,212,170,0.07) !important; }
.stInfo    { border-left: 3px solid #4361EE !important; background: rgba(67,97,238,0.07) !important; }
.stWarning { border-left: 3px solid #F6B93B !important; background: rgba(246,185,59,0.07) !important; }

/* ── Divider ── */
.cmd-divider { height: 1px; background: rgba(255,255,255,0.04); margin: 1.5rem 0; }

/* ── Tag badges ── */
.tag { display: inline-block; font-family: 'JetBrains Mono', monospace; font-size: 9px;
       padding: 2px 8px; border-radius: 2px; letter-spacing: 0.08em; font-weight: 500; }
.tag-green  { background: rgba(0,212,170,0.1);  color: #00D4AA;  border: 1px solid rgba(0,212,170,0.2); }
.tag-red    { background: rgba(255,77,109,0.1); color: #FF4D6D;  border: 1px solid rgba(255,77,109,0.2); }
.tag-amber  { background: rgba(246,185,59,0.1); color: #F6B93B;  border: 1px solid rgba(246,185,59,0.2); }
.tag-blue   { background: rgba(67,97,238,0.1);  color: #4361EE;  border: 1px solid rgba(67,97,238,0.2); }
.tag-neu    { background: rgba(74,85,104,0.2);  color: #718096;  border: 1px solid rgba(74,85,104,0.2); }

/* ── Row Stats Bar ── */
.stats-bar {
    display: flex; gap: 1px; background: rgba(255,255,255,0.04);
    border-radius: 3px; overflow: hidden; margin-bottom: 1rem;
}
.stats-bar-item {
    flex: 1; background: #0B1120; padding: 0.6rem 1rem;
    font-family: 'JetBrains Mono', monospace;
}
.stats-bar-label { font-size: 8px; color: #2D3748; letter-spacing: 0.1em; text-transform: uppercase; margin-bottom: 3px; }
.stats-bar-value { font-size: 13px; color: #E2E8F0; font-weight: 500; }

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 4px; height: 4px; }
::-webkit-scrollbar-track { background: #080C14; }
::-webkit-scrollbar-thumb { background: rgba(0,212,170,0.2); border-radius: 2px; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CHART CONFIG
# ══════════════════════════════════════════════════════════════════════════════
BG      = "#0B1120"
GRID    = "rgba(255,255,255,0.04)"
TEXT    = "#4A5568"
TEXT_L  = "#718096"
WHITE   = "#F7FAFC"
GREEN   = "#00D4AA"
RED     = "#FF4D6D"
AMBER   = "#F6B93B"
BLUE    = "#4361EE"
PURPLE  = "#7B2FBE"
TEAL    = "#0EA5E9"
GRAY    = "#2D3748"

BASE = dict(
    paper_bgcolor=BG, plot_bgcolor=BG,
    font=dict(family="JetBrains Mono", color=TEXT, size=10),
    margin=dict(l=0, r=0, t=4, b=0),
)

def ch(title):
    return f'<div class="chart-title">{title}</div>'

def kpi_card(label, value, delta, accent="green", delta_type="neu", small=False):
    size_cls = " small" if small else ""
    return f"""
<div class="kpi-card accent-{accent}">
  <div class="kpi-label">{label}</div>
  <div class="kpi-value{size_cls}">{value}</div>
  <div class="kpi-delta {delta_type}">{delta}</div>
</div>"""

def fmt_inr(v):
    if pd.isna(v) or v == 0: return "₹0"
    if v >= 1e7: return f"₹{v/1e7:.2f}Cr"
    if v >= 1e5: return f"₹{v/1e5:.1f}L"
    return f"₹{v:,.0f}"

PM_MAP = {
    "mode1":"Full Payment","mode2":"Instalment","mode3":"EMI",
    "mode4":"Partial","mode13":"Scholarship","mode5":"Other",
}

# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
if "master" not in st.session_state: st.session_state.master = None
if "loaded" not in st.session_state: st.session_state.loaded = False

# ══════════════════════════════════════════════════════════════════════════════
# DATA PROCESSING
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def load_merge(sem_b, conv_b):
    sem  = pd.read_csv(io.BytesIO(sem_b))
    conv = pd.read_excel(io.BytesIO(conv_b))

    sem.columns = sem.columns.str.strip()
    for c in ['Is Attended ?','Is Converted ?','Session','TRADER','Is our Student ?',
               'Trainer / Presenter','Place']:
        sem[c] = sem[c].astype(str).str.strip().str.upper()
    sem['Seminar Date']  = pd.to_datetime(sem['Seminar Date'], errors='coerce', dayfirst=True)
    sem['Amount Paid']   = pd.to_numeric(sem['Amount Paid'], errors='coerce').fillna(0)
    sem['mobile_clean']  = sem['Mobile'].astype(str).str.replace(r'\D','',regex=True).str[-10:]

    att = sem[sem['Is Attended ?'] == 'YES'].copy().reset_index(drop=True)
    att['Conv Status'] = att['Is Converted ?'].apply(
        lambda x: 'Converted' if x in ['CONVERTED','YES'] else 'Not Converted')

    conv['order_date']       = pd.to_datetime(conv['order_date'], errors='coerce', utc=True).dt.tz_localize(None)
    conv['payment_received'] = pd.to_numeric(conv['payment_received'], errors='coerce').fillna(0)
    conv['total_amount']     = pd.to_numeric(conv['total_amount'],     errors='coerce').fillna(0)
    conv['total_due']        = pd.to_numeric(conv['total_due'],        errors='coerce').fillna(0)
    conv['phone_clean']      = conv['phone'].astype(str).str.replace(r'\D','',regex=True).str[-10:]
    conv['PM Label']         = conv['payment_mode'].map(PM_MAP).fillna(conv['payment_mode'])
    conv['service_name']     = conv['service_name'].astype(str).str.strip()
    conv['sales_rep_name']   = conv['sales_rep_name'].astype(str).str.strip()
    conv['trainer_name']     = conv['trainer'].astype(str).str.split(' - ').str[-1].str.strip()

    mg = att.merge(
        conv[['phone_clean','orderID','order_date','service_code','service_name',
              'payment_received','total_amount','total_due','total_gst',
              'payment_mode','PM Label','status','sales_rep_name',
              'trainer','trainer_name','student_invid','batch_date']],
        left_on='mobile_clean', right_on='phone_clean', how='left'
    )

    def due_tag(r):
        if pd.isna(r.get('total_due')): return 'No Order'
        if r['total_due'] <= 0:         return 'Fully Paid'
        if r['total_amount'] > 0 and r['total_due'] < r['total_amount']: return 'Partially Paid'
        return 'Fully Due'
    mg['Due Status'] = mg.apply(due_tag, axis=1)
    return mg

# ══════════════════════════════════════════════════════════════════════════════
# TOP BAR
# ══════════════════════════════════════════════════════════════════════════════
now = datetime.now()
st.markdown(f"""
<div class="topbar">
  <div class="topbar-left">
    <div class="topbar-logo">◈ SEMINAR<span>IQ</span></div>
    <div class="topbar-sep"></div>
    <div class="topbar-sub">INTELLIGENCE COMMAND CENTER</div>
  </div>
  <div class="topbar-right">
    <div class="status-dot"></div>
    <div class="status-label">LIVE</div>
    <div class="topbar-sep"></div>
    <div class="topbar-time">{now.strftime("%d %b %Y  %H:%M")}</div>
  </div>
</div>
<div class="cmd-wrap">
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
t_up, t_dash, t_records = st.tabs([
    "◈  DATA UPLOAD",
    "▣  ANALYTICS DASHBOARD",
    "≡  RECORDS & EXPORT"
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
with t_up:
    st.markdown("""
    <div class="sec-header" style="margin-top:1.5rem">
      <div class="sec-title">Data Ingestion</div>
      <div class="sec-line"></div>
    </div>""", unsafe_allow_html=True)

    st.markdown("""
    <div style="background:#0B1120;border:1px solid rgba(0,212,170,0.1);border-radius:4px;
                padding:1.2rem 1.5rem;margin-bottom:1.5rem;">
      <div style="font-family:'Syne',sans-serif;font-size:13px;color:#E2E8F0;font-weight:600;margin-bottom:6px;">
        How it works
      </div>
      <div style="font-family:'Outfit',sans-serif;font-size:12px;color:#4A5568;line-height:1.8;">
        Upload your <span style="color:#00D4AA;font-family:'JetBrains Mono',monospace;">Seminar Sheet (CSV)</span>
        and <span style="color:#00D4AA;font-family:'JetBrains Mono',monospace;">Conversion List (XLSX)</span>.
        The system joins both on <b style="color:#718096;">phone number</b> to enrich every attendee
        with their course, payment, due amount, sales rep and trainer details — instantly.
      </div>
    </div>""", unsafe_allow_html=True)

    uc1, uc2 = st.columns(2, gap="medium")
    with uc1:
        st.markdown("""
        <div style="font-family:'JetBrains Mono',monospace;font-size:9px;color:#00D4AA;
                    letter-spacing:0.15em;text-transform:uppercase;margin-bottom:8px;">
          ◈ FILE 01 — Seminar Updated Sheet
        </div>""", unsafe_allow_html=True)
        f_sem = st.file_uploader("sem", type=["csv"], key="up_sem", label_visibility="collapsed")
        st.markdown("""<div style="font-family:'JetBrains Mono',monospace;font-size:9px;color:#2D3748;margin-top:6px;">
          Columns used: NAME · Mobile · Place · Trainer · Seminar Date · Session · Is Attended · Is Converted · Amount Paid
        </div>""", unsafe_allow_html=True)

    with uc2:
        st.markdown("""
        <div style="font-family:'JetBrains Mono',monospace;font-size:9px;color:#4361EE;
                    letter-spacing:0.15em;text-transform:uppercase;margin-bottom:8px;">
          ◈ FILE 02 — Conversion List
        </div>""", unsafe_allow_html=True)
        f_conv = st.file_uploader("conv", type=["xlsx","xls"], key="up_conv", label_visibility="collapsed")
        st.markdown("""<div style="font-family:'JetBrains Mono',monospace;font-size:9px;color:#2D3748;margin-top:6px;">
          Columns used: phone · service_name · total_amount · total_due · payment_received · status · sales_rep_name · trainer
        </div>""", unsafe_allow_html=True)

    if f_sem and f_conv:
        with st.spinner(""):
            mg = load_merge(f_sem.read(), f_conv.read())
        st.session_state.master = mg
        st.session_state.loaded = True

        total   = len(mg)
        matched = mg['orderID'].notna().sum()
        conv_n  = (mg['Conv Status']=='Converted').sum()
        locs    = mg['Place'].nunique()

        st.markdown("""<div class="sec-header" style="margin-top:1.5rem">
          <div class="sec-title">Ingestion Summary</div><div class="sec-line"></div>
        </div>""", unsafe_allow_html=True)

        m1,m2,m3,m4,m5 = st.columns(5)
        m1.metric("Total Attendees",   f"{total:,}")
        m2.metric("Matched to Orders", f"{matched:,}",  f"{matched/total*100:.1f}% match rate")
        m3.metric("Converted",         f"{conv_n:,}",   f"{conv_n/total*100:.1f}% conv rate")
        m4.metric("Locations",         f"{locs}")
        m5.metric("Seminar Dates",     f"{mg['Seminar Date'].nunique()}")

        st.success(f"✓  Both files merged successfully. Navigate to **Analytics Dashboard** to explore.")
    elif f_sem or f_conv:
        missing = "Conversion List (XLSX)" if f_sem else "Seminar Sheet (CSV)"
        st.info(f"Waiting for: **{missing}**")
    else:
        st.markdown("""
        <div class="upload-zone">
          <div class="upload-zone-title">No files uploaded</div>
          <div class="upload-zone-sub" style="margin-top:6px;">Upload both files above to activate the dashboard</div>
        </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
with t_dash:
    if not st.session_state.loaded:
        st.markdown("""<div style="padding:3rem;text-align:center;color:#2D3748;
            font-family:'JetBrains Mono',monospace;font-size:12px;letter-spacing:0.1em;">
            ◈  UPLOAD FILES FIRST TO ACTIVATE DASHBOARD
        </div>""", unsafe_allow_html=True)
    else:
        dfa = st.session_state.master.copy()

        # ── FILTER PANEL ──────────────────────────────────────────────────────
        with st.expander("▼  FILTERS — CLICK TO EXPAND", expanded=True):
            fa,fb,fc,fd = st.columns(4)
            fe,ff,fg,fh = st.columns(4)
            fi,fj,fk,_  = st.columns(4)

            all_places   = sorted(dfa['Place'].dropna().unique())
            all_trainers = sorted(dfa['Trainer / Presenter'].dropna().unique())
            all_dates    = sorted(dfa['Seminar Date'].dropna().unique())
            all_sess     = ["MORNING","EVENING"]
            all_conv     = ["Converted","Not Converted"]
            all_due      = sorted(dfa['Due Status'].dropna().unique())
            all_courses  = sorted(dfa['service_name'].dropna().unique())
            all_reps     = sorted(dfa['sales_rep_name'].dropna().unique())
            all_status   = sorted(dfa['status'].dropna().unique())

            sel_place   = fa.multiselect("Location",         all_places,   key="d_place")
            sel_trainer = fb.multiselect("Trainer / Presenter", all_trainers, key="d_trainer")
            sel_date    = fc.multiselect("Seminar Date",
                [d.strftime("%d %b %Y") for d in all_dates], key="d_date")
            sel_sess    = fd.multiselect("Session",          all_sess,     key="d_sess")
            sel_conv    = fe.multiselect("Conversion Status",all_conv,     key="d_conv")
            sel_due     = ff.multiselect("Due Status",       all_due,      key="d_due")
            sel_course  = fg.multiselect("Course",           all_courses,  key="d_course")
            sel_rep     = fh.multiselect("Sales Rep",        all_reps,     key="d_rep")
            sel_status  = fi.multiselect("Order Status",     all_status,   key="d_status")
            sel_trader  = fj.selectbox("Is Trader?",         ["All","Yes","No"], key="d_trader")
            sel_ourstu  = fk.selectbox("Existing Student?",  ["All","Yes","No"], key="d_ourstu")

        # Apply filters
        df = dfa.copy()
        if sel_place:   df = df[df['Place'].isin(sel_place)]
        if sel_trainer: df = df[df['Trainer / Presenter'].isin(sel_trainer)]
        if sel_date:
            fmt = [pd.to_datetime(d, format="%d %b %Y") for d in sel_date]
            df = df[df['Seminar Date'].isin(fmt)]
        if sel_sess:    df = df[df['Session'].isin(sel_sess)]
        if sel_conv:    df = df[df['Conv Status'].isin(sel_conv)]
        if sel_due:     df = df[df['Due Status'].isin(sel_due)]
        if sel_course:  df = df[df['service_name'].isin(sel_course)]
        if sel_rep:     df = df[df['sales_rep_name'].isin(sel_rep)]
        if sel_status:  df = df[df['status'].isin(sel_status)]
        if sel_trader != "All":
            df = df[df['TRADER'].isin(['YES','Y','TYES'] if sel_trader=="Yes" else ['NO','N'])]
        if sel_ourstu != "All":
            df = df[df['Is our Student ?'].isin(['YES','STUDENT'] if sel_ourstu=="Yes" else ['NO'])]

        N       = len(df)
        conv_n  = (df['Conv Status']=='Converted').sum()
        nconv_n = N - conv_n
        conv_r  = conv_n/N*100 if N else 0
        rcvd    = df['payment_received'].sum()
        due     = df['total_due'].sum()
        rev     = df['total_amount'].sum()
        fp      = (df['Due Status']=='Fully Paid').sum()
        hd      = df[df['Due Status'].isin(['Partially Paid','Fully Due'])].shape[0]

        # ── KPI ROW ───────────────────────────────────────────────────────────
        st.markdown("""<div class="sec-header" style="margin-top:1.2rem">
          <div class="sec-title">Key Performance Indicators</div>
          <div class="sec-line"></div>
          <div class="sec-count">FILTERED VIEW</div>
        </div>""", unsafe_allow_html=True)

        st.markdown(f"""
        <div class="kpi-grid">
          {kpi_card("TOTAL ATTENDED",  f"{N:,}",         f"of {len(dfa):,} total",      "blue",   "blue")}
          {kpi_card("CONVERTED",       f"{conv_n:,}",    f"↑ {conv_r:.1f}% rate",       "green",  "up"   if conv_r>15 else "warn")}
          {kpi_card("NOT CONVERTED",   f"{nconv_n:,}",   f"{100-conv_r:.1f}% of base",  "red",    "down")}
          {kpi_card("GROSS REVENUE",   fmt_inr(rev),     "Total order value",            "purple", "up")}
          {kpi_card("COLLECTED",       fmt_inr(rcvd),    f"Due: {fmt_inr(due)}",         "teal",   "up")}
          {kpi_card("FULLY PAID",      f"{fp:,}",        f"Has due: {hd:,}",             "amber",  "up" if fp>hd else "warn")}
        </div>""", unsafe_allow_html=True)

        st.markdown('<div style="height:1.5rem"></div>', unsafe_allow_html=True)

        # ── ROW A: Location conversion + Funnel ───────────────────────────────
        st.markdown("""<div class="sec-header">
          <div class="sec-title">Attendance & Conversion</div><div class="sec-line"></div>
        </div>""", unsafe_allow_html=True)

        ra1, ra2, ra3 = st.columns([4,3,2], gap="small")

        with ra1:
            loc = df.groupby('Place').agg(
                Attended=('NAME','count'),
                Converted=('Conv Status', lambda x:(x=='Converted').sum())
            ).reset_index().sort_values('Attended')
            loc['Rate'] = (loc['Converted']/loc['Attended']*100).round(1)

            fig = go.Figure()
            fig.add_trace(go.Bar(
                name='Attended', y=loc['Place'], x=loc['Attended'],
                orientation='h', marker_color=BLUE, marker_line_width=0, opacity=0.5,
                hovertemplate='%{y}: %{x} attended<extra></extra>'
            ))
            fig.add_trace(go.Bar(
                name='Converted', y=loc['Place'], x=loc['Converted'],
                orientation='h', marker_color=GREEN, marker_line_width=0,
                hovertemplate='%{y}: %{x} converted<extra></extra>'
            ))
            fig.update_layout(**BASE, height=max(260, len(loc)*36),
                barmode='overlay', bargap=0.32,
                legend=dict(orientation="h",x=0,y=1.06,font=dict(size=9),
                            bgcolor="rgba(0,0,0,0)"),
                xaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
                yaxis=dict(showgrid=False, tickfont=dict(size=10, color=TEXT_L)))
            st.markdown(ch("CONVERSION BY LOCATION"), unsafe_allow_html=True)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

        with ra2:
            sess = df.groupby('Session').agg(
                Attended=('NAME','count'),
                Converted=('Conv Status', lambda x:(x=='Converted').sum())
            ).reset_index()
            sess = sess[sess['Session'].isin(['MORNING','EVENING'])]

            fig2 = go.Figure()
            fig2.add_trace(go.Bar(
                name='Attended', x=sess['Session'], y=sess['Attended'],
                marker_color=BLUE, opacity=0.5, marker_line_width=0
            ))
            fig2.add_trace(go.Bar(
                name='Converted', x=sess['Session'], y=sess['Converted'],
                marker_color=GREEN, marker_line_width=0
            ))
            fig2.update_layout(**BASE, height=160, barmode='group', bargap=0.38,
                legend=dict(orientation="h",x=0,y=1.1,font=dict(size=9),bgcolor="rgba(0,0,0,0)"),
                xaxis=dict(showgrid=False, tickfont=dict(size=10,color=TEXT_L)),
                yaxis=dict(gridcolor=GRID, tickfont=dict(size=9,color=GRAY)))
            st.markdown(ch("MORNING vs EVENING"), unsafe_allow_html=True)
            st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar":False})

            st.markdown(ch("TRADER MIX"), unsafe_allow_html=True)
            trader = df['TRADER'].map(lambda x:'Trader' if x in ['YES','Y','TYES'] else 'Non-Trader').value_counts()
            fig3 = go.Figure(go.Pie(
                labels=trader.index, values=trader.values, hole=0.65,
                marker=dict(colors=[TEAL,PURPLE], line=dict(color=BG, width=3)),
                textfont=dict(size=9), textinfo='label+percent',
            ))
            fig3.update_layout(**BASE, height=130,
                showlegend=False,
                annotations=[dict(text="TRADER",x=0.5,y=0.5,showarrow=False,
                                  font=dict(size=8,color=GRAY,family="JetBrains Mono"))])
            st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar":False})

        with ra3:
            st.markdown(ch("CONVERSION FUNNEL"), unsafe_allow_html=True)
            fig_f = go.Figure(go.Funnel(
                y=["REGISTERED","ATTENDED","CONVERTED"],
                x=[len(dfa), N, conv_n],
                textinfo="value+percent initial",
                textfont=dict(family="JetBrains Mono",size=9,color=WHITE),
                marker=dict(color=[BLUE, PURPLE, GREEN],
                            line=dict(color=BG, width=2)),
                connector=dict(line=dict(color=GRID,width=1)),
            ))
            fig_f.update_layout(**BASE, height=220)
            st.plotly_chart(fig_f, use_container_width=True, config={"displayModeBar":False})

            st.markdown(ch("EXISTING STUDENT MIX"), unsafe_allow_html=True)
            stu = df['Is our Student ?'].map(
                lambda x:'Existing' if x in ['YES','STUDENT'] else 'New Lead'
            ).value_counts()
            fig_s = go.Figure(go.Pie(
                labels=stu.index, values=stu.values, hole=0.65,
                marker=dict(colors=[AMBER,TEAL], line=dict(color=BG,width=3)),
                textfont=dict(size=9), textinfo='label+percent',
            ))
            fig_s.update_layout(**BASE, height=130, showlegend=False,
                annotations=[dict(text="STUDENT",x=0.5,y=0.5,showarrow=False,
                                  font=dict(size=8,color=GRAY,family="JetBrains Mono"))])
            st.plotly_chart(fig_s, use_container_width=True, config={"displayModeBar":False})

        st.markdown('<div class="cmd-divider"></div>', unsafe_allow_html=True)

        # ── ROW B: Date trend + Due Status + Payment Mode ─────────────────────
        st.markdown("""<div class="sec-header">
          <div class="sec-title">Time Series & Payment Intelligence</div><div class="sec-line"></div>
        </div>""", unsafe_allow_html=True)

        rb1, rb2, rb3 = st.columns([4,2,2], gap="small")

        with rb1:
            date_g = df.groupby('Seminar Date').agg(
                Attended=('NAME','count'),
                Converted=('Conv Status', lambda x:(x=='Converted').sum())
            ).reset_index().dropna(subset=['Seminar Date']).sort_values('Seminar Date')
            date_g['Label'] = date_g['Seminar Date'].dt.strftime("%d %b")

            fig4 = make_subplots(specs=[[{"secondary_y":True}]])
            fig4.add_trace(go.Bar(
                name='Attended', x=date_g['Label'], y=date_g['Attended'],
                marker_color=BLUE, marker_line_width=0, opacity=0.6,
                hovertemplate='%{x}: %{y} attended<extra></extra>'
            ), secondary_y=False)
            fig4.add_trace(go.Scatter(
                name='Converted', x=date_g['Label'], y=date_g['Converted'],
                line=dict(color=GREEN,width=2,dash='solid'),
                mode='lines+markers',
                marker=dict(size=6,color=GREEN,line=dict(color=BG,width=2)),
                hovertemplate='%{x}: %{y} converted<extra></extra>'
            ), secondary_y=True)
            fig4.update_layout(**BASE, height=220,
                bargap=0.35,
                legend=dict(orientation="h",x=0,y=1.08,font=dict(size=9),bgcolor="rgba(0,0,0,0)"),
                xaxis=dict(showgrid=False, tickfont=dict(size=9,color=TEXT_L)),
                yaxis=dict(gridcolor=GRID, tickfont=dict(size=9,color=GRAY), title=None),
                yaxis2=dict(showgrid=False, tickfont=dict(size=9,color=GREEN), title=None))
            st.markdown(ch("ATTENDANCE & CONVERSION BY SEMINAR DATE"), unsafe_allow_html=True)
            st.plotly_chart(fig4, use_container_width=True, config={"displayModeBar":False})

        with rb2:
            due_c = df['Due Status'].value_counts()
            col_m = {'Fully Paid':GREEN,'Partially Paid':AMBER,'Fully Due':RED,'No Order':GRAY}
            fig5 = go.Figure(go.Pie(
                labels=due_c.index, values=due_c.values, hole=0.62,
                marker=dict(colors=[col_m.get(l,BLUE) for l in due_c.index],
                            line=dict(color=BG,width=3)),
                textfont=dict(size=9,family="JetBrains Mono"),
                textinfo='label+percent',
                hovertemplate='%{label}: %{value}<extra></extra>',
            ))
            fig5.update_layout(**BASE, height=220, showlegend=False,
                annotations=[dict(text="DUE",x=0.5,y=0.5,showarrow=False,
                                  font=dict(size=9,color=GRAY,family="JetBrains Mono"))])
            st.markdown(ch("PAYMENT DUE STATUS"), unsafe_allow_html=True)
            st.plotly_chart(fig5, use_container_width=True, config={"displayModeBar":False})

        with rb3:
            pm = df[df['Conv Status']=='Converted']['PM Label'].value_counts()
            fig6 = go.Figure(go.Pie(
                labels=pm.index, values=pm.values, hole=0.62,
                marker=dict(colors=[BLUE,PURPLE,GREEN,AMBER,RED,TEAL,GRAY],
                            line=dict(color=BG,width=3)),
                textfont=dict(size=9,family="JetBrains Mono"),
                textinfo='label+percent',
            ))
            fig6.update_layout(**BASE, height=220, showlegend=False,
                annotations=[dict(text="MODE",x=0.5,y=0.5,showarrow=False,
                                  font=dict(size=9,color=GRAY,family="JetBrains Mono"))])
            st.markdown(ch("PAYMENT MODE (CONVERTED)"), unsafe_allow_html=True)
            st.plotly_chart(fig6, use_container_width=True, config={"displayModeBar":False})

        st.markdown('<div class="cmd-divider"></div>', unsafe_allow_html=True)

        # ── ROW C: Courses + Sales Rep + Trainer ──────────────────────────────
        st.markdown("""<div class="sec-header">
          <div class="sec-title">Course · Sales Rep · Trainer Performance</div><div class="sec-line"></div>
        </div>""", unsafe_allow_html=True)

        rc1, rc2, rc3 = st.columns(3, gap="small")

        with rc1:
            cr_df = df[df['Conv Status']=='Converted'].groupby('service_name').agg(
                Count=('NAME','count'), Revenue=('total_amount','sum')
            ).reset_index().sort_values('Count').tail(8)
            cr_df['Short'] = cr_df['service_name'].apply(lambda x: x[:28]+'…' if len(x)>28 else x)
            fig7 = go.Figure(go.Bar(
                x=cr_df['Count'], y=cr_df['Short'], orientation='h',
                marker=dict(color=cr_df['Count'],colorscale=[[0,PURPLE],[1,TEAL]],
                            line=dict(width=0)),
                text=cr_df['Count'], textposition='outside',
                textfont=dict(size=9,color=TEXT_L),
                hovertemplate='%{y}: %{x} conversions<extra></extra>',
            ))
            fig7.update_layout(**BASE, height=280,
                xaxis=dict(showgrid=False,showticklabels=False,zeroline=False),
                yaxis=dict(showgrid=False, tickfont=dict(size=9,color=TEXT_L)), bargap=0.3)
            st.markdown(ch("TOP COURSES BY CONVERSIONS"), unsafe_allow_html=True)
            st.plotly_chart(fig7, use_container_width=True, config={"displayModeBar":False})

        with rc2:
            rp_df = df[df['Conv Status']=='Converted'].groupby('sales_rep_name').agg(
                Conv=('NAME','count'), Rev=('total_amount','sum')
            ).reset_index().sort_values('Conv').tail(10)
            fig8 = go.Figure(go.Bar(
                x=rp_df['Conv'], y=rp_df['sales_rep_name'], orientation='h',
                marker=dict(color=rp_df['Conv'],colorscale=[[0,BLUE],[1,GREEN]],
                            line=dict(width=0)),
                text=rp_df['Conv'], textposition='outside',
                textfont=dict(size=9,color=TEXT_L),
            ))
            fig8.update_layout(**BASE, height=280,
                xaxis=dict(showgrid=False,showticklabels=False,zeroline=False),
                yaxis=dict(showgrid=False, tickfont=dict(size=9,color=TEXT_L)), bargap=0.3)
            st.markdown(ch("SALES REP — CONVERSION COUNT"), unsafe_allow_html=True)
            st.plotly_chart(fig8, use_container_width=True, config={"displayModeBar":False})

        with rc3:
            tr_df = df.groupby('Trainer / Presenter').agg(
                Attended=('NAME','count'),
                Conv=('Conv Status', lambda x:(x=='Converted').sum())
            ).reset_index().sort_values('Attended')
            tr_df['Short'] = tr_df['Trainer / Presenter'].apply(lambda x: x[:26]+'…' if len(x)>26 else x)
            fig9 = go.Figure()
            fig9.add_trace(go.Bar(
                name='Attended', y=tr_df['Short'], x=tr_df['Attended'],
                orientation='h', marker_color=BLUE, opacity=0.4, marker_line_width=0,
            ))
            fig9.add_trace(go.Bar(
                name='Converted', y=tr_df['Short'], x=tr_df['Conv'],
                orientation='h', marker_color=GREEN, marker_line_width=0,
            ))
            fig9.update_layout(**BASE, height=280,
                barmode='overlay', bargap=0.3,
                legend=dict(orientation="h",x=0,y=1.1,font=dict(size=9),bgcolor="rgba(0,0,0,0)"),
                xaxis=dict(showgrid=False,showticklabels=False),
                yaxis=dict(showgrid=False, tickfont=dict(size=9,color=TEXT_L)))
            st.markdown(ch("TRAINER — ATTENDANCE vs CONVERSION"), unsafe_allow_html=True)
            st.plotly_chart(fig9, use_container_width=True, config={"displayModeBar":False})

        st.markdown('<div class="cmd-divider"></div>', unsafe_allow_html=True)

        # ── ROW D: Revenue ────────────────────────────────────────────────────
        st.markdown("""<div class="sec-header">
          <div class="sec-title">Revenue Intelligence</div><div class="sec-line"></div>
        </div>""", unsafe_allow_html=True)

        rd1, rd2 = st.columns([3,2], gap="small")

        with rd1:
            rev_loc = df.groupby('Place').agg(
                Collected=('payment_received','sum'),
                Due=('total_due','sum')
            ).reset_index().sort_values('Collected')
            fig10 = go.Figure()
            fig10.add_trace(go.Bar(
                name='Collected', y=rev_loc['Place'], x=rev_loc['Collected'],
                orientation='h', marker_color=GREEN, marker_line_width=0, opacity=0.85,
                hovertemplate='%{y} Collected: ₹%{x:,.0f}<extra></extra>'
            ))
            fig10.add_trace(go.Bar(
                name='Due', y=rev_loc['Place'], x=rev_loc['Due'],
                orientation='h', marker_color=RED, marker_line_width=0, opacity=0.85,
                hovertemplate='%{y} Due: ₹%{x:,.0f}<extra></extra>'
            ))
            fig10.update_layout(**BASE, height=max(240,len(rev_loc)*36),
                barmode='stack', bargap=0.3,
                legend=dict(orientation="h",x=0,y=1.08,font=dict(size=9),bgcolor="rgba(0,0,0,0)"),
                xaxis=dict(gridcolor=GRID, tickformat=",.0f", tickfont=dict(size=9,color=GRAY)),
                yaxis=dict(showgrid=False, tickfont=dict(size=9,color=TEXT_L)))
            st.markdown(ch("COLLECTED vs DUE BY LOCATION"), unsafe_allow_html=True)
            st.plotly_chart(fig10, use_container_width=True, config={"displayModeBar":False})

        with rd2:
            rep_rev = df[df['Conv Status']=='Converted'].groupby('sales_rep_name').agg(
                Rev=('total_amount','sum'), Due=('total_due','sum')
            ).reset_index().sort_values('Rev', ascending=False).head(8)
            rep_rev['Short'] = rep_rev['sales_rep_name']

            fig11 = go.Figure()
            fig11.add_trace(go.Bar(
                name='Revenue', x=rep_rev['Rev'], y=rep_rev['Short'],
                orientation='h', marker=dict(
                    color=rep_rev['Rev'],
                    colorscale=[[0,PURPLE],[0.5,BLUE],[1,TEAL]],
                    line=dict(width=0)),
                hovertemplate='%{y}: ₹%{x:,.0f}<extra></extra>',
                textfont=dict(size=9,color=TEXT_L),
            ))
            fig11.update_layout(**BASE, height=max(240, len(rep_rev)*32),
                showlegend=False, bargap=0.3,
                xaxis=dict(showgrid=False,showticklabels=False,zeroline=False),
                yaxis=dict(showgrid=False, tickfont=dict(size=9,color=TEXT_L)))
            st.markdown(ch("SALES REP — REVENUE GENERATED"), unsafe_allow_html=True)
            st.plotly_chart(fig11, use_container_width=True, config={"displayModeBar":False})

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — RECORDS & EXPORT
# ══════════════════════════════════════════════════════════════════════════════
with t_records:
    if not st.session_state.loaded:
        st.markdown("""<div style="padding:3rem;text-align:center;color:#2D3748;
            font-family:'JetBrains Mono',monospace;font-size:12px;letter-spacing:0.1em;">
            ◈  UPLOAD FILES FIRST TO ACTIVATE RECORDS
        </div>""", unsafe_allow_html=True)
    else:
        dfa = st.session_state.master.copy()

        st.markdown("""<div class="sec-header" style="margin-top:1.2rem">
          <div class="sec-title">Search & Filter Records</div><div class="sec-line"></div>
        </div>""", unsafe_allow_html=True)

        # Filters
        tf1,tf2,tf3,tf4 = st.columns(4)
        tf5,tf6,tf7,tf8 = st.columns(4)
        tf9,tf10,tf11,tf12 = st.columns(4)

        t_place   = tf1.multiselect("Location",         sorted(dfa['Place'].dropna().unique()), key="r_place")
        t_trainer = tf2.multiselect("Trainer",          sorted(dfa['Trainer / Presenter'].dropna().unique()), key="r_trainer")
        t_date    = tf3.multiselect("Seminar Date",
                       [d.strftime("%d %b %Y") for d in sorted(dfa['Seminar Date'].dropna().unique())], key="r_date")
        t_sess    = tf4.multiselect("Session",          ["MORNING","EVENING"], key="r_sess")
        t_conv    = tf5.multiselect("Conversion Status",["Converted","Not Converted"], key="r_conv")
        t_due     = tf6.multiselect("Due Status",       sorted(dfa['Due Status'].dropna().unique()), key="r_due")
        t_course  = tf7.multiselect("Course",           sorted(dfa['service_name'].dropna().unique()), key="r_course")
        t_rep     = tf8.multiselect("Sales Rep",        sorted(dfa['sales_rep_name'].dropna().unique()), key="r_rep")
        t_status  = tf9.multiselect("Order Status",     sorted(dfa['status'].dropna().unique()), key="r_status")
        t_trader  = tf10.selectbox("Is Trader?",        ["All","Yes","No"], key="r_trader")
        t_ourstu  = tf11.selectbox("Existing Student?", ["All","Yes","No"], key="r_ourstu")
        t_search  = tf12.text_input("Search Name / Phone / Email", key="r_search",
                                    placeholder="Type to search…")

        dt = dfa.copy()
        if t_place:   dt = dt[dt['Place'].isin(t_place)]
        if t_trainer: dt = dt[dt['Trainer / Presenter'].isin(t_trainer)]
        if t_date:
            fmt = [pd.to_datetime(d, format="%d %b %Y") for d in t_date]
            dt = dt[dt['Seminar Date'].isin(fmt)]
        if t_sess:    dt = dt[dt['Session'].isin(t_sess)]
        if t_conv:    dt = dt[dt['Conv Status'].isin(t_conv)]
        if t_due:     dt = dt[dt['Due Status'].isin(t_due)]
        if t_course:  dt = dt[dt['service_name'].isin(t_course)]
        if t_rep:     dt = dt[dt['sales_rep_name'].isin(t_rep)]
        if t_status:  dt = dt[dt['status'].isin(t_status)]
        if t_trader != "All":
            dt = dt[dt['TRADER'].isin(['YES','Y','TYES'] if t_trader=="Yes" else ['NO','N'])]
        if t_ourstu != "All":
            dt = dt[dt['Is our Student ?'].isin(['YES','STUDENT'] if t_ourstu=="Yes" else ['NO'])]
        if t_search:
            mask = pd.Series([False]*len(dt))
            for c in ['NAME','Mobile','email']:
                if c in dt.columns:
                    mask = mask | dt[c].astype(str).str.lower().str.contains(t_search.lower(), na=False)
            dt = dt[mask]

        # Stats bar
        N_t    = len(dt)
        conv_t = (dt['Conv Status']=='Converted').sum()
        rcvd_t = dt['payment_received'].sum()
        due_t  = dt['total_due'].sum()
        st.markdown(f"""
        <div class="stats-bar">
          <div class="stats-bar-item"><div class="stats-bar-label">Records</div><div class="stats-bar-value">{N_t:,}</div></div>
          <div class="stats-bar-item"><div class="stats-bar-label">Converted</div><div class="stats-bar-value" style="color:#00D4AA">{conv_t:,}</div></div>
          <div class="stats-bar-item"><div class="stats-bar-label">Not Converted</div><div class="stats-bar-value" style="color:#FF4D6D">{N_t-conv_t:,}</div></div>
          <div class="stats-bar-item"><div class="stats-bar-label">Collected</div><div class="stats-bar-value" style="color:#00D4AA">{fmt_inr(rcvd_t)}</div></div>
          <div class="stats-bar-item"><div class="stats-bar-label">Total Due</div><div class="stats-bar-value" style="color:#FF4D6D">{fmt_inr(due_t)}</div></div>
        </div>""", unsafe_allow_html=True)

        # Display cols
        disp = ['NAME','Mobile','Place','Seminar Date','Session',
                'Trainer / Presenter','Conv Status','Amount Paid',
                'service_name','total_amount','payment_received','total_due',
                'Due Status','status','sales_rep_name','trainer_name',
                'TRADER','Is our Student ?','PM Label','Remarks']
        disp = [c for c in disp if c in dt.columns]

        ds = dt[disp].copy()
        ds['Seminar Date'] = ds['Seminar Date'].dt.strftime("%d %b %Y")

        st.dataframe(
            ds, use_container_width=True, height=480, hide_index=True,
            column_config={
                "total_amount":        st.column_config.NumberColumn("Course Amt",  format="₹%.0f"),
                "payment_received":    st.column_config.NumberColumn("Collected",   format="₹%.0f"),
                "total_due":           st.column_config.NumberColumn("Due",         format="₹%.0f"),
                "Amount Paid":         st.column_config.NumberColumn("Sem Paid",    format="₹%.0f"),
                "service_name":        st.column_config.TextColumn("Course"),
                "sales_rep_name":      st.column_config.TextColumn("Sales Rep"),
                "trainer_name":        st.column_config.TextColumn("Trainer (Order)"),
                "Trainer / Presenter": st.column_config.TextColumn("Seminar Trainer"),
                "PM Label":            st.column_config.TextColumn("Payment Mode"),
                "Conv Status":         st.column_config.TextColumn("Conversion"),
                "Due Status":          st.column_config.TextColumn("Due Status"),
            }
        )

        # Location summary
        st.markdown("""<div class="sec-header" style="margin-top:1.5rem">
          <div class="sec-title">Location Summary</div><div class="sec-line"></div>
        </div>""", unsafe_allow_html=True)

        loc_s = dt.groupby('Place').agg(
            Attended=('NAME','count'),
            Converted=('Conv Status', lambda x:(x=='Converted').sum()),
            Revenue=('total_amount','sum'),
            Collected=('payment_received','sum'),
            Due=('total_due','sum'),
        ).reset_index()
        loc_s['Conv Rate'] = (loc_s['Converted']/loc_s['Attended']*100).round(1).astype(str)+'%'
        loc_s['Revenue']   = loc_s['Revenue'].apply(fmt_inr)
        loc_s['Collected'] = loc_s['Collected'].apply(fmt_inr)
        loc_s['Due']       = loc_s['Due'].apply(fmt_inr)
        st.dataframe(loc_s, use_container_width=True, hide_index=True, height=260)

        # Export
        st.markdown("""<div class="sec-header" style="margin-top:1.5rem">
          <div class="sec-title">Export Data</div><div class="sec-line"></div>
        </div>""", unsafe_allow_html=True)

        def to_xl(df):
            buf = io.BytesIO(); df.to_excel(buf, index=False); return buf.getvalue()
        ts = datetime.now().strftime("%Y%m%d_%H%M")

        ex1,ex2,ex3,ex4 = st.columns(4)
        ex1.download_button("⬇  Filtered Records",    to_xl(dt[disp]),
            file_name=f"seminar_filtered_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
        ex2.download_button("⬇  Converted Only",
            to_xl(dt[dt['Conv Status']=='Converted'][disp]),
            file_name=f"converted_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
        ex3.download_button("⬇  Has Due Amount",
            to_xl(dt[dt['total_due']>0][disp]),
            file_name=f"has_due_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
        ex4.download_button("⬇  Location Summary",    to_xl(loc_s),
            file_name=f"location_summary_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

# Close main wrapper
st.markdown("</div>", unsafe_allow_html=True)
