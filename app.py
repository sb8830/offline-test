"""
Seminar Intelligence Dashboard — Invesmate UI Style
Includes: Seminar CSV + Conversion List + Leads File
"""
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
from datetime import datetime

st.set_page_config(
    page_title="Seminar Intelligence",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
  #MainMenu,footer,header{visibility:hidden}
  .block-container{padding:0!important;max-width:100%!important}
  .stApp{background:#060910}
  div[data-testid="stToolbar"]{display:none}
  section[data-testid="stSidebar"]{display:none}
  div[data-testid="stDecoration"]{display:none}
  div[data-testid="stStatusWidget"]{display:none}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# DESIGN SYSTEM
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@400;500;600&display=swap" rel="stylesheet">
<style>
.im-nav{background:linear-gradient(180deg,#0d1119 0%,#080b12 100%);
  border-bottom:1px solid rgba(255,255,255,.07);padding:0 24px;height:60px;
  display:flex;align-items:center;justify-content:space-between;
  position:sticky;top:0;z-index:9999}
.im-brand{font-family:'Syne',sans-serif;font-size:16px;font-weight:800;color:#eceef5;letter-spacing:-.3px;line-height:1.1}
.im-brand-sub{font-size:9px;color:#4a5068;text-transform:uppercase;letter-spacing:.9px}
.im-pill{background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.08);
  border-radius:20px;padding:4px 14px 4px 10px;display:flex;align-items:center;gap:8px;font-size:12px;color:#8a90aa}
.im-dot{width:7px;height:7px;background:#4fce8f;border-radius:50%;animation:imdot 2s infinite;flex-shrink:0}
@keyframes imdot{0%,100%{opacity:1}50%{opacity:.3}}

.sg{display:grid;gap:11px}
.sc{background:#111520;border:1px solid rgba(255,255,255,.06);border-radius:12px;
  padding:14px 16px;position:relative;overflow:hidden}
.sc::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:var(--c,#4f8ef7)}
.sv{font-family:'Syne',sans-serif;font-size:26px;font-weight:800;color:#eceef5;line-height:1}
.sv.sm{font-size:18px}
.sl{font-size:10px;color:#4a5068;text-transform:uppercase;letter-spacing:.5px;margin-top:5px}
.sd{font-size:10px;margin-top:3px}
.sd.g{color:#4fce8f}.sd.r{color:#f76f4f}.sd.a{color:#f7c948}.sd.n{color:#4a5068}

.asec{background:#0c1018;border:1px solid rgba(255,255,255,.07);border-radius:14px;padding:20px 22px;margin-bottom:16px}
.asec-t{font-family:'Syne',sans-serif;font-size:10px;font-weight:700;color:#f7c948;
  margin-bottom:14px;text-transform:uppercase;letter-spacing:.9px}
.ct{font-family:'Syne',sans-serif;font-size:10px;font-weight:700;
  color:#f7c948;text-transform:uppercase;letter-spacing:.9px;margin-bottom:12px}

.sbar{display:flex;gap:1px;background:rgba(255,255,255,.04);border-radius:10px;overflow:hidden;margin-bottom:14px}
.sbar-cell{flex:1;background:#111520;padding:.7rem 1rem}
.sbar-lbl{font-size:9px;color:#4a5068;text-transform:uppercase;letter-spacing:.5px;margin-bottom:3px}
.sbar-val{font-family:'Syne',sans-serif;font-size:14px;font-weight:700;color:#eceef5}

.file-card{background:#0c1018;border:1px solid rgba(255,255,255,.07);border-radius:12px;padding:14px;margin-bottom:8px}
.file-card-title{font-family:'Syne',sans-serif;font-size:11px;font-weight:700;color:#eceef5;margin:5px 0 2px}
.file-card-sub{font-size:10px;color:#4a5068}

div[data-baseweb="select"]>div{background:#111520!important;border-color:rgba(255,255,255,.08)!important;border-radius:8px!important;font-family:'DM Sans',sans-serif!important;font-size:13px!important}
div[data-baseweb="select"]>div:hover{border-color:rgba(79,206,143,.3)!important}
div[data-baseweb="input"]>div{background:#111520!important;border-color:rgba(255,255,255,.08)!important;border-radius:8px!important}
label[data-testid="stWidgetLabel"]{font-family:'DM Sans',sans-serif!important;font-size:11px!important;color:#4a5068!important;letter-spacing:.02em!important}
.stMultiSelect [data-baseweb="tag"]{background:rgba(79,206,143,.12)!important;border:1px solid rgba(79,206,143,.25)!important;border-radius:6px!important}
.stTabs [data-baseweb="tab-list"]{background:transparent!important;border-bottom:1px solid rgba(255,255,255,.07)!important;gap:0}
.stTabs [data-baseweb="tab"]{font-family:'DM Sans',sans-serif!important;font-size:12px!important;font-weight:500!important;color:#4a5068!important;padding:.6rem 1.3rem!important;border-bottom:2px solid transparent!important;background:transparent!important}
.stTabs [aria-selected="true"]{color:#4fce8f!important;border-bottom-color:#4fce8f!important}
.stExpander{border:1px solid rgba(255,255,255,.07)!important;border-radius:12px!important;background:#0c1018!important}
.stExpander summary p{font-family:'DM Sans',sans-serif!important;font-size:12px!important;color:#8a90aa!important}
button[data-testid="stBaseButton-secondary"]{background:#111520!important;border:1px solid rgba(255,255,255,.1)!important;color:#8a90aa!important;border-radius:8px!important;font-family:'DM Sans',sans-serif!important;font-size:12px!important;transition:all .15s!important}
button[data-testid="stBaseButton-secondary"]:hover{border-color:rgba(79,206,143,.4)!important;color:#eceef5!important}
button[data-testid="stBaseButton-primary"]{background:linear-gradient(135deg,#4fce8f,#3ab87a)!important;border:none!important;color:#060910!important;font-weight:700!important;border-radius:8px!important;font-family:'DM Sans',sans-serif!important}
div[data-testid="metric-container"]{background:#111520!important;border:1px solid rgba(255,255,255,.07)!important;border-radius:12px!important;padding:.9rem 1rem!important}
div[data-testid="stMetricValue"]{font-family:'Syne',sans-serif!important;font-size:20px!important;font-weight:800!important;color:#eceef5!important}
div[data-testid="stMetricLabel"]{font-size:10px!important;color:#4a5068!important;text-transform:uppercase!important;letter-spacing:.5px!important}
div[data-testid="stMetricDelta"] svg{display:none!important}
.stSuccess{background:rgba(79,206,143,.07)!important;border-left:3px solid #4fce8f!important;border-radius:8px!important}
.stInfo{background:rgba(79,142,247,.07)!important;border-left:3px solid #4f8ef7!important;border-radius:8px!important}
.stWarning{background:rgba(247,201,72,.07)!important;border-left:3px solid #f7c948!important;border-radius:8px!important}
.stDataFrame{border-radius:10px!important}
::-webkit-scrollbar{width:4px;height:4px}
::-webkit-scrollbar-track{background:#060910}
::-webkit-scrollbar-thumb{background:rgba(79,206,143,.2);border-radius:2px}
.page{padding:0 22px 60px 22px}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CHART PALETTE
# ══════════════════════════════════════════════════════════════════════════════
BG   = "#0c1018"
GRID = "rgba(255,255,255,0.04)"
TXT  = "#4a5068"
TXT2 = "#8a90aa"
C_GRN  = "#4fce8f"; C_BLUE = "#4f8ef7"; C_RED  = "#f76f4f"
C_AMB  = "#f7c948"; C_PURP = "#b44fe7"; C_TEAL = "#4fd8f7"; C_GRAY = "#2d3748"

BASE = dict(paper_bgcolor=BG, plot_bgcolor=BG,
            font=dict(family="DM Sans", color=TXT, size=11),
            margin=dict(l=0, r=8, t=4, b=0))

PM_MAP = {"mode1":"Full Payment","mode2":"Instalment","mode3":"EMI","mode4":"Partial","mode13":"Scholarship"}

def fmt(v):
    if pd.isna(v) or v==0: return "₹0"
    if v>=1e7: return f"₹{v/1e7:.2f}Cr"
    if v>=1e5: return f"₹{v/1e5:.1f}L"
    return f"₹{v:,.0f}"

def ksc(label, val, delta="", c="#4f8ef7", dc="n", sm=False):
    return f"""<div class="sc" style="--c:{c}">
<div class="sv{'  sm' if sm else ''}">{val}</div>
<div class="sl">{label}</div>
{"" if not delta else f'<div class="sd {dc}">{delta}</div>'}
</div>"""

def ct(t): return f'<div class="ct">{t}</div>'

def empty_state(icon, title, sub):
    return f"""<div style="padding:60px;text-align:center">
<div style="font-size:36px;margin-bottom:16px">{icon}</div>
<div style="font-family:'Syne',sans-serif;font-size:18px;font-weight:800;color:#eceef5;margin-bottom:8px">{title}</div>
<div style="color:#4a5068;font-size:13px">{sub}</div></div>"""

# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
for k in ["sem_conv","leads","loaded_sem","loaded_leads"]:
    if k not in st.session_state: st.session_state[k] = None
for k in ["loaded_sem","loaded_leads"]:
    if k not in st.session_state: st.session_state[k] = False

# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADERS
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def load_sem_conv(sb, cb):
    sem  = pd.read_csv(io.BytesIO(sb))
    conv = pd.read_excel(io.BytesIO(cb))
    sem.columns = sem.columns.str.strip()
    for c in ['Is Attended ?','Is Converted ?','Session','TRADER','Is our Student ?','Trainer / Presenter','Place']:
        sem[c] = sem[c].astype(str).str.strip().str.upper()
    sem['Seminar Date'] = pd.to_datetime(sem['Seminar Date'], errors='coerce', dayfirst=True)
    sem['Amount Paid']  = pd.to_numeric(sem['Amount Paid'], errors='coerce').fillna(0)
    sem['mob']          = sem['Mobile'].astype(str).str.replace(r'\D','',regex=True).str[-10:]
    att = sem[sem['Is Attended ?']=='YES'].copy().reset_index(drop=True)
    att['Conv Status'] = att['Is Converted ?'].apply(
        lambda x: 'Converted' if x in ['CONVERTED','YES'] else 'Not Converted')
    for c in ['payment_received','total_amount','total_due']:
        conv[c] = pd.to_numeric(conv[c], errors='coerce').fillna(0)
    conv['order_date']    = pd.to_datetime(conv['order_date'], errors='coerce', utc=True).dt.tz_localize(None)
    conv['phone_clean']   = conv['phone'].astype(str).str.replace(r'\D','',regex=True).str[-10:]
    conv['PM Label']      = conv['payment_mode'].map(PM_MAP).fillna(conv['payment_mode'])
    conv['service_name']  = conv['service_name'].astype(str).str.strip()
    conv['sales_rep_name']= conv['sales_rep_name'].astype(str).str.strip()
    conv['trainer_name']  = conv['trainer'].astype(str).str.split(' - ').str[-1].str.strip()
    mg = att.merge(
        conv[['phone_clean','orderID','order_date','service_code','service_name',
              'payment_received','total_amount','total_due','total_gst',
              'payment_mode','PM Label','status','sales_rep_name',
              'trainer','trainer_name','student_invid','batch_date']],
        left_on='mob', right_on='phone_clean', how='left')
    def due_tag(r):
        if pd.isna(r.get('total_due')): return 'No Order'
        if r['total_due']<=0:           return 'Fully Paid'
        if r['total_amount']>0 and r['total_due']<r['total_amount']: return 'Partially Paid'
        return 'Fully Due'
    mg['Due Status'] = mg.apply(due_tag, axis=1)
    return mg

@st.cache_data(show_spinner=False)
def load_leads(lb):
    df = pd.read_excel(io.BytesIO(lb), sheet_name='Sheet 1')
    df['leaddate']    = pd.to_datetime(df['leaddate'],    errors='coerce')
    df['updatedate']  = pd.to_datetime(df['updatedate'],  errors='coerce')
    # Normalise key columns
    for c in ['converted_from','leadsource','campaign_name','leadstatus',
              'stage_name','leadownername','state','Attempted/Unattempted',
              'servicename','servicetype']:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
            df[c] = df[c].replace('nan', 'Unknown')
    df['phone_clean'] = df['phone'].astype(str).str.replace(r'\D','',regex=True).str[-10:]
    # Derived
    df['Lead Month'] = df['leaddate'].dt.to_period('M').astype(str)
    df['Is Converted'] = df['leadstatus'].str.lower().str.contains('converted', na=False)
    return df

# ══════════════════════════════════════════════════════════════════════════════
# NAV
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="im-nav">
  <div style="display:flex;align-items:center;gap:11px">
    <div style="width:36px;height:36px;border-radius:50%;
      background:linear-gradient(135deg,#4fce8f,#4f8ef7);
      display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0">📊</div>
    <div>
      <div class="im-brand">Invesmate</div>
      <div class="im-brand-sub">Seminar Intelligence Hub</div>
    </div>
  </div>
  <div class="im-pill">
    <div class="im-dot"></div>
    <span>{datetime.now().strftime("%d %b %Y  %H:%M")}</span>
  </div>
</div>
<div class="page">
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
t_up, t_sem, t_leads, t_rec = st.tabs([
    "📁  Upload Files",
    "📊  Seminar Analytics",
    "🎯  Leads Analytics",
    "📋  Records & Export"
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
with t_up:
    st.markdown("""
<div style="text-align:center;padding:36px 20px 24px">
  <div style="font-size:36px">📊</div>
  <div style="font-family:'Syne',sans-serif;font-size:32px;font-weight:800;
    color:#eceef5;margin:10px 0 6px;letter-spacing:-1px">Seminar Intelligence Hub</div>
  <div style="color:#4a5068;font-size:12px;text-transform:uppercase;letter-spacing:.8px">
    Upload · Merge · Analyse · Export
  </div>
</div>""", unsafe_allow_html=True)

    # ── FILE SET 1: Seminar + Conversion ─────────────────────────────────────
    st.markdown('<div class="asec"><div class="asec-t">📋 File Set 1 — Seminar Attendance & Conversion</div>', unsafe_allow_html=True)
    s1c1, s1c2 = st.columns(2, gap="medium")
    with s1c1:
        st.markdown("""<div class="file-card">
          <span style="font-size:18px">📝</span>
          <div class="file-card-title">Seminar Updated Sheet (.csv)</div>
          <div class="file-card-sub">NAME · Mobile · Place · Trainer · Date · Session · Attended · Converted · Amount Paid</div>
        </div>""", unsafe_allow_html=True)
        f_sem = st.file_uploader("sem", type=["csv"], key="up_sem", label_visibility="collapsed")

    with s1c2:
        st.markdown("""<div class="file-card">
          <span style="font-size:18px">📋</span>
          <div class="file-card-title">Conversion List (.xlsx)</div>
          <div class="file-card-sub">phone · service_name · total_amount · total_due · payment_received · status · sales_rep_name · trainer</div>
        </div>""", unsafe_allow_html=True)
        f_conv = st.file_uploader("conv", type=["xlsx","xls"], key="up_conv", label_visibility="collapsed")

    _, b1, _ = st.columns([1,2,1])
    with b1:
        if f_sem and f_conv:
            if st.button("🚀  Load Seminar Data", use_container_width=True, type="primary", key="load_sem"):
                with st.spinner("Merging seminar + conversion files…"):
                    mg = load_sem_conv(f_sem.read(), f_conv.read())
                st.session_state.sem_conv    = mg
                st.session_state.loaded_sem  = True
                total = len(mg); matched = mg['orderID'].notna().sum()
                conv_n = (mg['Conv Status']=='Converted').sum()
                st.success(f"✅ Seminar data ready — {total:,} attendees · {matched:,} matched · {conv_n:,} converted")
        elif f_sem or f_conv:
            st.info(f"⏳ Waiting for: **{'Conversion List' if f_sem else 'Seminar Sheet'}**")
        else:
            st.markdown("""<div style="text-align:center;padding:12px;
              background:rgba(79,206,143,.04);border:1px dashed rgba(79,206,143,.2);
              border-radius:10px;color:#4a5068;font-size:13px">
              Upload both files above to load seminar analytics
            </div>""", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # ── FILE SET 2: Leads ─────────────────────────────────────────────────────
    st.markdown('<div class="asec"><div class="asec-t">🎯 File Set 2 — Leads Data</div>', unsafe_allow_html=True)
    l1, l2 = st.columns([2,3], gap="medium")
    with l1:
        st.markdown("""<div class="file-card">
          <span style="font-size:18px">🎯</span>
          <div class="file-card-title">Leads Report (.xlsx)</div>
          <div class="file-card-sub">
            converted_from (Webinar / Non Webinar) · leadsource · campaign_name ·
            leadstatus · stage_name · leadownername · servicename · state · Attempted/Unattempted
          </div>
        </div>""", unsafe_allow_html=True)
        f_leads = st.file_uploader("leads", type=["xlsx","xls"], key="up_leads", label_visibility="collapsed")

    with l2:
        st.markdown("""<div style="background:#111520;border:1px solid rgba(255,255,255,.06);
          border-radius:10px;padding:14px 16px;">
          <div style="font-family:'Syne',sans-serif;font-size:10px;font-weight:700;
            color:#f7c948;text-transform:uppercase;letter-spacing:.9px;margin-bottom:10px">Key Filters Available in Leads Tab</div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:11px;color:#8a90aa">
            <div>✅ Converted From <span style="color:#4fce8f">(Webinar / Non-Webinar)</span></div>
            <div>📢 Lead Source</div>
            <div>📣 Campaign Name</div>
            <div>📊 Lead Status / Stage</div>
            <div>👤 Lead Owner (Sales Rep)</div>
            <div>🗓 Lead Date Range</div>
            <div>📍 State</div>
            <div>📞 Attempted / Unattempted</div>
          </div>
        </div>""", unsafe_allow_html=True)

    _, b2, _ = st.columns([1,2,1])
    with b2:
        if f_leads:
            if st.button("🚀  Load Leads Data", use_container_width=True, type="primary", key="load_leads"):
                with st.spinner("Parsing leads file…"):
                    ldf = load_leads(f_leads.read())
                st.session_state.leads         = ldf
                st.session_state.loaded_leads  = True
                st.success(f"✅ Leads data ready — {len(ldf):,} leads · {ldf['Is Converted'].sum():,} converted · {ldf['converted_from'].value_counts().get('Webinar',0)} from Webinar")
        else:
            st.markdown("""<div style="text-align:center;padding:12px;
              background:rgba(79,142,247,.04);border:1px dashed rgba(79,142,247,.2);
              border-radius:10px;color:#4a5068;font-size:13px">
              Upload leads file above to load leads analytics
            </div>""", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # Status summary
    s1, s2, s3 = st.columns(3)
    s1.metric("Seminar Data", "✅ Loaded" if st.session_state.loaded_sem  else "⏳ Not loaded",
              f"{len(st.session_state.sem_conv):,} attendees" if st.session_state.loaded_sem else "")
    s2.metric("Leads Data",   "✅ Loaded" if st.session_state.loaded_leads else "⏳ Not loaded",
              f"{len(st.session_state.leads):,} leads" if st.session_state.loaded_leads else "")
    s3.metric("Tabs Unlocked", f"{int(st.session_state.loaded_sem) + int(st.session_state.loaded_leads) + 1} / 4", "")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — SEMINAR ANALYTICS
# ══════════════════════════════════════════════════════════════════════════════
with t_sem:
    if not st.session_state.loaded_sem:
        st.markdown(empty_state("📊","No Seminar Data","Upload Seminar Sheet + Conversion List in the Upload Files tab"), unsafe_allow_html=True)
    else:
        dfa = st.session_state.sem_conv.copy()

        with st.expander("🔽  Seminar Filters", expanded=True):
            fa,fb,fc,fd = st.columns(4)
            fe,ff,fg,fh = st.columns(4)
            fi,fj,fk,_  = st.columns(4)
            sel_place   = fa.multiselect("📍 Location",         sorted(dfa['Place'].dropna().unique()),                               key="s_pl")
            sel_trainer = fb.multiselect("🎤 Trainer",          sorted(dfa['Trainer / Presenter'].dropna().unique()),                 key="s_tr")
            sel_date    = fc.multiselect("📅 Seminar Date",     [d.strftime("%d %b %Y") for d in sorted(dfa['Seminar Date'].dropna().unique())], key="s_dt")
            sel_sess    = fd.multiselect("🕐 Session",          ["MORNING","EVENING"],                                               key="s_se")
            sel_conv    = fe.multiselect("✅ Conversion",        ["Converted","Not Converted"],                                       key="s_cv")
            sel_due     = ff.multiselect("💳 Due Status",        sorted(dfa['Due Status'].dropna().unique()),                         key="s_du")
            sel_course  = fg.multiselect("📚 Course",            sorted(dfa['service_name'].dropna().unique()),                       key="s_co")
            sel_rep     = fh.multiselect("👤 Sales Rep",         sorted(dfa['sales_rep_name'].dropna().unique()),                     key="s_re")
            sel_status  = fi.multiselect("📌 Order Status",      sorted(dfa['status'].dropna().unique()),                            key="s_st")
            sel_trader  = fj.selectbox("🔄 Is Trader?",          ["All","Yes","No"],                                                 key="s_td")
            sel_ourstu  = fk.selectbox("🎓 Existing Student?",   ["All","Yes","No"],                                                 key="s_os")

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

        N=len(df); conv_n=(df['Conv Status']=='Converted').sum(); nconv=N-conv_n
        conv_r=conv_n/N*100 if N else 0; rcvd=df['payment_received'].sum()
        due=df['total_due'].sum(); rev=df['total_amount'].sum()
        fp=(df['Due Status']=='Fully Paid').sum(); hd=df[df['Due Status'].isin(['Partially Paid','Fully Due'])].shape[0]

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f"""<div class="sg" style="grid-template-columns:repeat(8,minmax(0,1fr))">
  {ksc("Total Attended", f"{N:,}",       f"of {len(dfa):,} total", "#4f8ef7","n")}
  {ksc("Converted",      f"{conv_n:,}",  f"↑ {conv_r:.1f}%",      "#4fce8f","g")}
  {ksc("Not Converted",  f"{nconv:,}",   f"{100-conv_r:.1f}%",    "#f76f4f","r")}
  {ksc("Gross Revenue",  fmt(rev),       "Order value",            "#b44fe7","n",True)}
  {ksc("Collected",      fmt(rcvd),      f"Due: {fmt(due)}",       "#4fce8f","g",True)}
  {ksc("Fully Paid",     f"{fp:,}",      f"Has bal: {hd:,}",      "#f7c948","a" if fp>hd else "r")}
  {ksc("Locations",      f"{df['Place'].nunique()}", "Venues",     "#4fd8f7","n")}
  {ksc("Sales Reps",     f"{df['sales_rep_name'].dropna().nunique()}", "Active", "#4f8ef7","n")}
</div>""", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # ── Attendance & Conversion ───────────────────────────────────────────
        st.markdown('<div class="asec"><div class="asec-t">📍 Attendance & Conversion by Location</div>', unsafe_allow_html=True)
        loc = df.groupby('Place').agg(Attended=('NAME','count'),Converted=('Conv Status',lambda x:(x=='Converted').sum())).reset_index().sort_values('Attended',ascending=False)
        loc['Rate'] = (loc['Converted']/loc['Attended']*100).round(1)
        a1,a2 = st.columns([3,2],gap="medium")
        with a1:
            st.markdown(ct("Attended vs Converted — by Location"), unsafe_allow_html=True)
            fig=go.Figure()
            fig.add_trace(go.Bar(name='Attended',x=loc['Place'],y=loc['Attended'],marker_color=C_BLUE,marker_line_width=0,opacity=0.5,hovertemplate='<b>%{x}</b><br>Attended:%{y}<extra></extra>'))
            fig.add_trace(go.Bar(name='Converted',x=loc['Place'],y=loc['Converted'],marker_color=C_GRN,marker_line_width=0,hovertemplate='<b>%{x}</b><br>Converted:%{y}<extra></extra>'))
            fig.update_layout(**BASE,height=250,barmode='group',bargap=0.25,bargroupgap=0.06,
                legend=dict(orientation="h",x=0,y=1.12,font=dict(size=11),bgcolor="rgba(0,0,0,0)"),
                xaxis=dict(showgrid=False,tickangle=-30,tickfont=dict(size=10,color=TXT2)),
                yaxis=dict(gridcolor=GRID,tickfont=dict(size=10,color=TXT)))
            st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})
        with a2:
            st.markdown(ct("Conversion Rate % — by Location"), unsafe_allow_html=True)
            loc_s=loc.sort_values('Rate'); bar_colors=[C_RED if r<10 else(C_AMB if r<20 else C_GRN) for r in loc_s['Rate']]
            fig2=go.Figure(go.Bar(x=loc_s['Rate'],y=loc_s['Place'],orientation='h',marker_color=bar_colors,marker_line_width=0,text=[f"{r}%" for r in loc_s['Rate']],textposition='outside',textfont=dict(size=10,color=TXT2),hovertemplate='<b>%{y}</b><br>%{x:.1f}%<extra></extra>'))
            if len(loc_s): fig2.add_vline(x=loc_s['Rate'].mean(),line_dash="dot",line_color=C_AMB,line_width=1,annotation_text=f"avg {loc_s['Rate'].mean():.1f}%",annotation_font=dict(size=9,color=C_AMB),annotation_position="top right")
            fig2.update_layout(**BASE,height=250,bargap=0.3,xaxis=dict(showgrid=False,showticklabels=False,zeroline=False,range=[0,loc_s['Rate'].max()*1.3 if len(loc_s) else 10]),yaxis=dict(showgrid=False,tickfont=dict(size=10,color=TXT2)))
            st.plotly_chart(fig2,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

        # ── Date Trend + Session + Funnel ─────────────────────────────────────
        st.markdown('<div class="asec"><div class="asec-t">📅 Trend · Session · Funnel</div>',unsafe_allow_html=True)
        b1,b2,b3=st.columns([4,2,2],gap="medium")
        with b1:
            st.markdown(ct("Attendance Trend & Conversion Rate — by Seminar Date"),unsafe_allow_html=True)
            dg=df.groupby('Seminar Date').agg(Attended=('NAME','count'),Converted=('Conv Status',lambda x:(x=='Converted').sum())).reset_index().dropna(subset=['Seminar Date']).sort_values('Seminar Date')
            dg['Label']=dg['Seminar Date'].dt.strftime("%d %b"); dg['CR']=(dg['Converted']/dg['Attended']*100).round(1)
            fig3=make_subplots(specs=[[{"secondary_y":True}]])
            fig3.add_trace(go.Bar(name='Attended',x=dg['Label'],y=dg['Attended'],marker_color=C_BLUE,marker_line_width=0,opacity=0.55),secondary_y=False)
            fig3.add_trace(go.Scatter(name='Conv Rate %',x=dg['Label'],y=dg['CR'],line=dict(color=C_GRN,width=2.5,shape='spline'),mode='lines+markers',marker=dict(size=6,color=C_GRN,line=dict(color=BG,width=1.5))),secondary_y=True)
            fig3.update_layout(**BASE,height=220,bargap=0.3,legend=dict(orientation="h",x=0,y=1.12,font=dict(size=11),bgcolor="rgba(0,0,0,0)"),xaxis=dict(showgrid=False,tickfont=dict(size=10,color=TXT2)),yaxis=dict(gridcolor=GRID,tickfont=dict(size=10,color=TXT),zeroline=False),yaxis2=dict(showgrid=False,tickfont=dict(size=10,color=C_GRN),ticksuffix="%"))
            st.plotly_chart(fig3,use_container_width=True,config={"displayModeBar":False})
        with b2:
            st.markdown(ct("Morning vs Evening"),unsafe_allow_html=True)
            sess=df.groupby('Session').agg(Attended=('NAME','count'),Converted=('Conv Status',lambda x:(x=='Converted').sum())).reset_index()
            sess=sess[sess['Session'].isin(['MORNING','EVENING'])].copy(); sess['NotConv']=sess['Attended']-sess['Converted']
            fig4=go.Figure()
            fig4.add_trace(go.Bar(name='Converted',x=sess['Session'],y=sess['Converted'],marker_color=C_GRN,marker_line_width=0))
            fig4.add_trace(go.Bar(name='Not Converted',x=sess['Session'],y=sess['NotConv'],marker_color=C_BLUE,opacity=0.3,marker_line_width=0))
            fig4.update_layout(**BASE,height=220,barmode='stack',bargap=0.4,legend=dict(orientation="h",x=0,y=1.12,font=dict(size=10),bgcolor="rgba(0,0,0,0)"),xaxis=dict(showgrid=False,tickfont=dict(size=11,color=TXT2)),yaxis=dict(gridcolor=GRID,tickfont=dict(size=10,color=TXT)))
            st.plotly_chart(fig4,use_container_width=True,config={"displayModeBar":False})
        with b3:
            st.markdown(ct("Conversion Funnel"),unsafe_allow_html=True)
            fig5=go.Figure(go.Funnel(y=["Registered","Attended","Converted"],x=[len(dfa),N,conv_n],textinfo="value+percent initial",textfont=dict(family="DM Sans",size=11),marker=dict(color=[C_BLUE,C_PURP,C_GRN],line=dict(color=BG,width=2)),connector=dict(line=dict(color=GRID,width=1))))
            fig5.update_layout(**BASE,height=220)
            st.plotly_chart(fig5,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

        # ── Payment + Course + Rep ────────────────────────────────────────────
        st.markdown('<div class="asec"><div class="asec-t">💳 Payment · Course · Sales Rep · Trainer</div>',unsafe_allow_html=True)
        c1,c2,c3,c4=st.columns(4,gap="medium")
        with c1:
            st.markdown(ct("Due Status"),unsafe_allow_html=True)
            duc=df['Due Status'].value_counts(); cmap={'Fully Paid':C_GRN,'Partially Paid':C_AMB,'Fully Due':C_RED,'No Order':C_GRAY}
            fig6=go.Figure(go.Pie(labels=duc.index,values=duc.values,hole=0.62,marker=dict(colors=[cmap.get(l,C_BLUE) for l in duc.index],line=dict(color=BG,width=3)),textinfo='percent',textfont=dict(size=10,family="DM Sans"),hovertemplate='<b>%{label}</b><br>%{value}<extra></extra>'))
            fig6.update_layout(**BASE,height=210,legend=dict(orientation="v",x=1.02,y=0.5,font=dict(size=9)),annotations=[dict(text="Due",x=0.5,y=0.5,showarrow=False,font=dict(size=11,color=TXT2,family="DM Sans"))])
            st.plotly_chart(fig6,use_container_width=True,config={"displayModeBar":False})
        with c2:
            st.markdown(ct("Payment Mode"),unsafe_allow_html=True)
            pm=df[df['Conv Status']=='Converted']['PM Label'].value_counts()
            fig7=go.Figure(go.Bar(x=pm.values,y=pm.index,orientation='h',marker_color=[C_GRN,C_BLUE,C_PURP,C_AMB,C_RED][:len(pm)],marker_line_width=0,text=pm.values,textposition='outside',textfont=dict(size=10,color=TXT2)))
            fig7.update_layout(**BASE,height=210,bargap=0.35,xaxis=dict(showgrid=False,showticklabels=False,zeroline=False,range=[0,pm.max()*1.3] if len(pm) else [0,10]),yaxis=dict(showgrid=False,tickfont=dict(size=10,color=TXT2)))
            st.plotly_chart(fig7,use_container_width=True,config={"displayModeBar":False})
        with c3:
            st.markdown(ct("Top Courses"),unsafe_allow_html=True)
            cr=df[df['Conv Status']=='Converted'].groupby('service_name').agg(Count=('NAME','count')).reset_index().sort_values('Count')
            cr['Short']=cr['service_name'].apply(lambda x:x[:26]+'…' if len(x)>26 else x)
            fig8=go.Figure(go.Bar(x=cr['Count'],y=cr['Short'],orientation='h',marker=dict(color=cr['Count'],colorscale=[[0,C_PURP],[1,C_TEAL]],line=dict(width=0)),text=cr['Count'],textposition='outside',textfont=dict(size=10,color=TXT2)))
            fig8.update_layout(**BASE,height=210,bargap=0.3,xaxis=dict(showgrid=False,showticklabels=False,zeroline=False,range=[0,cr['Count'].max()*1.3] if len(cr) else [0,10]),yaxis=dict(showgrid=False,tickfont=dict(size=9,color=TXT2)))
            st.plotly_chart(fig8,use_container_width=True,config={"displayModeBar":False})
        with c4:
            st.markdown(ct("Trainer Performance"),unsafe_allow_html=True)
            tr=df.groupby('Trainer / Presenter').agg(Attended=('NAME','count'),Converted=('Conv Status',lambda x:(x=='Converted').sum())).reset_index()
            tr['Short']=tr['Trainer / Presenter'].apply(lambda x:x[:20]+'…' if len(x)>20 else x); tr=tr.sort_values('Attended')
            fig9=go.Figure()
            fig9.add_trace(go.Bar(name='Attended',y=tr['Short'],x=tr['Attended'],orientation='h',marker_color=C_BLUE,opacity=0.35,marker_line_width=0))
            fig9.add_trace(go.Bar(name='Converted',y=tr['Short'],x=tr['Converted'],orientation='h',marker_color=C_GRN,marker_line_width=0))
            fig9.update_layout(**BASE,height=210,barmode='overlay',bargap=0.28,legend=dict(orientation="h",x=0,y=1.12,font=dict(size=9),bgcolor="rgba(0,0,0,0)"),xaxis=dict(showgrid=False,showticklabels=False,zeroline=False),yaxis=dict(showgrid=False,tickfont=dict(size=8,color=TXT2)))
            st.plotly_chart(fig9,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

        # ── Revenue by Location ───────────────────────────────────────────────
        st.markdown('<div class="asec"><div class="asec-t">💰 Revenue — Collected vs Due by Location</div>',unsafe_allow_html=True)
        rev_l=df.groupby('Place').agg(Collected=('payment_received','sum'),Due=('total_due','sum')).reset_index().sort_values('Collected',ascending=False)
        fig10=go.Figure()
        fig10.add_trace(go.Bar(name='Collected',x=rev_l['Place'],y=rev_l['Collected'],marker_color=C_GRN,marker_line_width=0,opacity=0.85,hovertemplate='<b>%{x}</b><br>Collected:₹%{y:,.0f}<extra></extra>'))
        fig10.add_trace(go.Bar(name='Due',x=rev_l['Place'],y=rev_l['Due'],marker_color=C_RED,marker_line_width=0,opacity=0.85,hovertemplate='<b>%{x}</b><br>Due:₹%{y:,.0f}<extra></extra>'))
        fig10.update_layout(**BASE,height=230,barmode='group',bargap=0.25,bargroupgap=0.06,
            legend=dict(orientation="h",x=0,y=1.12,font=dict(size=11),bgcolor="rgba(0,0,0,0)"),
            xaxis=dict(showgrid=False,tickangle=-30,tickfont=dict(size=10,color=TXT2)),
            yaxis=dict(gridcolor=GRID,tickformat=",.0f",tickprefix="₹",tickfont=dict(size=10,color=TXT)))
        st.plotly_chart(fig10,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — LEADS ANALYTICS
# ══════════════════════════════════════════════════════════════════════════════
with t_leads:
    if not st.session_state.loaded_leads:
        st.markdown(empty_state("🎯","No Leads Data","Upload the Leads Report in the Upload Files tab"), unsafe_allow_html=True)
    else:
        ldf_all = st.session_state.leads.copy()

        # ── FILTERS ───────────────────────────────────────────────────────────
        with st.expander("🔽  Leads Filters", expanded=True):
            lf1,lf2,lf3,lf4 = st.columns(4)
            lf5,lf6,lf7,lf8 = st.columns(4)

            # Converted From — THE KEY FILTER
            conv_from_opts = [x for x in sorted(ldf_all['converted_from'].unique()) if x != 'Unknown']
            sel_conv_from  = lf1.multiselect("✅ Converted From",   ["Webinar","Non Webinar"],    key="l_cf")

            # Lead Source
            ls_opts = sorted([x for x in ldf_all['leadsource'].unique() if x != 'Unknown'])
            sel_ls  = lf2.multiselect("📢 Lead Source",             ls_opts,                      key="l_ls")

            # Campaign Name
            cp_opts = sorted([x for x in ldf_all['campaign_name'].unique() if x != 'Unknown'])
            sel_cp  = lf3.multiselect("📣 Campaign Name",           cp_opts,                      key="l_cp")

            # Lead Status
            lst_opts = sorted([x for x in ldf_all['leadstatus'].unique() if x != 'Unknown'])
            sel_lst  = lf4.multiselect("📊 Lead Status",            lst_opts,                     key="l_lst")

            # Stage
            stg_opts = sorted([x for x in ldf_all['stage_name'].unique() if x != 'Unknown'])
            sel_stg  = lf5.multiselect("🏷 Stage",                  stg_opts,                     key="l_stg")

            # Lead Owner
            lo_opts  = sorted([x for x in ldf_all['leadownername'].unique() if x != 'Unknown'])
            sel_lo   = lf6.multiselect("👤 Lead Owner",             lo_opts,                      key="l_lo")

            # State
            st_opts  = sorted([x for x in ldf_all['state'].unique() if x != 'Unknown'])
            sel_st   = lf7.multiselect("📍 State",                  st_opts,                      key="l_st")

            # Attempted
            sel_att  = lf8.selectbox("📞 Attempted?",               ["All","Attempted","Unattempted"], key="l_att")

            # Date range
            dr1, dr2 = st.columns(2)
            min_date = ldf_all['leaddate'].min().date() if ldf_all['leaddate'].notna().any() else None
            max_date = ldf_all['leaddate'].max().date() if ldf_all['leaddate'].notna().any() else None
            if min_date and max_date:
                date_from = dr1.date_input("📅 Lead Date From", value=min_date, min_value=min_date, max_value=max_date, key="l_df")
                date_to   = dr2.date_input("📅 Lead Date To",   value=max_date, min_value=min_date, max_value=max_date, key="l_dt")

        # Apply filters
        ldf = ldf_all.copy()
        if sel_conv_from: ldf = ldf[ldf['converted_from'].isin(sel_conv_from)]
        if sel_ls:        ldf = ldf[ldf['leadsource'].isin(sel_ls)]
        if sel_cp:        ldf = ldf[ldf['campaign_name'].isin(sel_cp)]
        if sel_lst:       ldf = ldf[ldf['leadstatus'].isin(sel_lst)]
        if sel_stg:       ldf = ldf[ldf['stage_name'].isin(sel_stg)]
        if sel_lo:        ldf = ldf[ldf['leadownername'].isin(sel_lo)]
        if sel_st:        ldf = ldf[ldf['state'].isin(sel_st)]
        if sel_att != "All": ldf = ldf[ldf['Attempted/Unattempted'] == sel_att]
        if min_date and max_date:
            ldf = ldf[(ldf['leaddate'].dt.date >= date_from) & (ldf['leaddate'].dt.date <= date_to)]

        NL       = len(ldf)
        conv_l   = ldf['Is Converted'].sum()
        conv_lr  = conv_l/NL*100 if NL else 0
        web_l    = (ldf['converted_from']=='Webinar').sum()
        nonweb_l = (ldf['converted_from']=='Non Webinar').sum()
        att_l    = (ldf['Attempted/Unattempted']=='Attempted').sum()
        seat_l   = (ldf['leadstatus']=='Seat Booked').sum()

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f"""<div class="sg" style="grid-template-columns:repeat(7,minmax(0,1fr))">
  {ksc("Total Leads",       f"{NL:,}",       f"of {len(ldf_all):,} total",     "#4f8ef7","n")}
  {ksc("Converted",         f"{conv_l:,}",   f"↑ {conv_lr:.1f}% rate",         "#4fce8f","g")}
  {ksc("Webinar Leads",     f"{web_l:,}",    f"{web_l/NL*100:.1f}% of total" if NL else "", "#4fd8f7","n")}
  {ksc("Non-Webinar Leads", f"{nonweb_l:,}", f"{nonweb_l/NL*100:.1f}% of total" if NL else "", "#b44fe7","n")}
  {ksc("Attempted",         f"{att_l:,}",    f"{att_l/NL*100:.1f}% contact rate" if NL else "", "#4fce8f","g")}
  {ksc("Seat Booked",       f"{seat_l:,}",   "In pipeline",                   "#f7c948","a")}
  {ksc("Campaigns",         f"{ldf['campaign_name'].nunique()}", "Active",     "#4f8ef7","n")}
</div>""", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # ── Webinar vs Non-Webinar ─────────────────────────────────────────────
        st.markdown('<div class="asec"><div class="asec-t">✅ Webinar vs Non-Webinar Conversion</div>',unsafe_allow_html=True)
        w1,w2,w3 = st.columns([2,2,3],gap="medium")
        with w1:
            st.markdown(ct("Converted From — Source Mix"),unsafe_allow_html=True)
            cf=ldf['converted_from'].value_counts()
            cf=cf[cf.index.isin(['Webinar','Non Webinar','Unknown'])]
            fig_cf=go.Figure(go.Pie(labels=cf.index,values=cf.values,hole=0.62,
                marker=dict(colors=[C_GRN,C_BLUE,C_GRAY],line=dict(color=BG,width=3)),
                textinfo='label+percent',textfont=dict(size=11,family="DM Sans")))
            fig_cf.update_layout(**BASE,height=220,showlegend=False,
                annotations=[dict(text="Source",x=0.5,y=0.5,showarrow=False,font=dict(size=11,color=TXT2,family="DM Sans"))])
            st.plotly_chart(fig_cf,use_container_width=True,config={"displayModeBar":False})
        with w2:
            st.markdown(ct("Conversion Rate — Webinar vs Non-Webinar"),unsafe_allow_html=True)
            wc=ldf[ldf['converted_from'].isin(['Webinar','Non Webinar'])].groupby('converted_from').agg(
                Total=('_id','count'), Conv=('Is Converted','sum')).reset_index()
            wc['Rate']=(wc['Conv']/wc['Total']*100).round(1)
            fig_wc=go.Figure(go.Bar(x=wc['Rate'],y=wc['converted_from'],orientation='h',
                marker_color=[C_GRN if r>wc['Rate'].mean() else C_BLUE for r in wc['Rate']],
                marker_line_width=0,text=[f"{r}%" for r in wc['Rate']],
                textposition='outside',textfont=dict(size=13,color=TXT2)))
            fig_wc.update_layout(**BASE,height=220,bargap=0.45,
                xaxis=dict(showgrid=False,showticklabels=False,zeroline=False,range=[0,wc['Rate'].max()*1.3] if len(wc) else [0,10]),
                yaxis=dict(showgrid=False,tickfont=dict(size=12,color=TXT2)))
            st.plotly_chart(fig_wc,use_container_width=True,config={"displayModeBar":False})
        with w3:
            st.markdown(ct("Lead Status Breakdown — Webinar vs Non-Webinar"),unsafe_allow_html=True)
            wls=ldf[ldf['converted_from'].isin(['Webinar','Non Webinar'])].groupby(['converted_from','leadstatus']).size().reset_index(name='Count')
            top_statuses=ldf['leadstatus'].value_counts().head(6).index.tolist()
            wls=wls[wls['leadstatus'].isin(top_statuses)]
            fig_wls=go.Figure()
            for i,(src,col) in enumerate([('Webinar',C_GRN),('Non Webinar',C_BLUE)]):
                sub=wls[wls['converted_from']==src]
                fig_wls.add_trace(go.Bar(name=src,x=sub['leadstatus'],y=sub['Count'],marker_color=col,opacity=0.85,marker_line_width=0,hovertemplate=f'{src}<br>%{{x}}: %{{y}}<extra></extra>'))
            fig_wls.update_layout(**BASE,height=220,barmode='group',bargap=0.2,bargroupgap=0.05,
                legend=dict(orientation="h",x=0,y=1.12,font=dict(size=11),bgcolor="rgba(0,0,0,0)"),
                xaxis=dict(showgrid=False,tickangle=-25,tickfont=dict(size=9,color=TXT2)),
                yaxis=dict(gridcolor=GRID,tickfont=dict(size=10,color=TXT)))
            st.plotly_chart(fig_wls,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

        # ── Lead Source + Campaign ─────────────────────────────────────────────
        st.markdown('<div class="asec"><div class="asec-t">📢 Lead Source & Campaign Performance</div>',unsafe_allow_html=True)
        ls1,ls2 = st.columns(2,gap="medium")
        with ls1:
            st.markdown(ct("Top Lead Sources — Volume & Conversion"),unsafe_allow_html=True)
            lsrc=ldf.groupby('leadsource').agg(Total=('_id','count'),Conv=('Is Converted','sum')).reset_index()
            lsrc['Rate']=(lsrc['Conv']/lsrc['Total']*100).round(1)
            lsrc=lsrc.sort_values('Total',ascending=False).head(12)
            fig_ls=make_subplots(specs=[[{"secondary_y":True}]])
            fig_ls.add_trace(go.Bar(name='Total Leads',x=lsrc['leadsource'],y=lsrc['Total'],marker_color=C_BLUE,marker_line_width=0,opacity=0.6,hovertemplate='<b>%{x}</b><br>Leads:%{y}<extra></extra>'),secondary_y=False)
            fig_ls.add_trace(go.Scatter(name='Conv Rate %',x=lsrc['leadsource'],y=lsrc['Rate'],mode='lines+markers',line=dict(color=C_GRN,width=2.5),marker=dict(size=7,color=C_GRN,line=dict(color=BG,width=1.5)),hovertemplate='<b>%{x}</b><br>%{y:.1f}%<extra></extra>'),secondary_y=True)
            fig_ls.update_layout(**BASE,height=270,bargap=0.3,legend=dict(orientation="h",x=0,y=1.12,font=dict(size=11),bgcolor="rgba(0,0,0,0)"),xaxis=dict(showgrid=False,tickangle=-35,tickfont=dict(size=9,color=TXT2)),yaxis=dict(gridcolor=GRID,tickfont=dict(size=10,color=TXT)),yaxis2=dict(showgrid=False,tickfont=dict(size=10,color=C_GRN),ticksuffix="%"))
            st.plotly_chart(fig_ls,use_container_width=True,config={"displayModeBar":False})
        with ls2:
            st.markdown(ct("Top Campaigns — Lead Volume"),unsafe_allow_html=True)
            cmp=ldf.groupby('campaign_name').agg(Total=('_id','count'),Conv=('Is Converted','sum')).reset_index()
            cmp['Rate']=(cmp['Conv']/cmp['Total']*100).round(1)
            cmp=cmp.sort_values('Total',ascending=False).head(12)
            cmp['Short']=cmp['campaign_name'].apply(lambda x:x[:30]+'…' if len(x)>30 else x)
            fig_cp=go.Figure(go.Bar(x=cmp['Total'],y=cmp['Short'],orientation='h',
                marker=dict(color=cmp['Conv'],colorscale=[[0,C_PURP],[1,C_GRN]],line=dict(width=0)),
                text=cmp['Total'],textposition='outside',textfont=dict(size=9,color=TXT2),
                customdata=cmp['Rate'],hovertemplate='<b>%{y}</b><br>Leads:%{x}<br>Conv Rate:%{customdata:.1f}%<extra></extra>'))
            fig_cp.update_layout(**BASE,height=270,bargap=0.25,xaxis=dict(showgrid=False,showticklabels=False,zeroline=False,range=[0,cmp['Total'].max()*1.25] if len(cmp) else [0,10]),yaxis=dict(showgrid=False,tickfont=dict(size=9,color=TXT2)))
            st.plotly_chart(fig_cp,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

        # ── Lead Status + Stage + Owner ───────────────────────────────────────
        st.markdown('<div class="asec"><div class="asec-t">📊 Status · Stage · Lead Owner</div>',unsafe_allow_html=True)
        o1,o2,o3=st.columns(3,gap="medium")
        with o1:
            st.markdown(ct("Lead Status Distribution"),unsafe_allow_html=True)
            lst=ldf['leadstatus'].value_counts().head(10)
            fig_lst=go.Figure(go.Bar(x=lst.values,y=lst.index,orientation='h',
                marker=dict(color=lst.values,colorscale=[[0,C_BLUE],[0.5,C_PURP],[1,C_GRN]],line=dict(width=0)),
                text=lst.values,textposition='outside',textfont=dict(size=9,color=TXT2)))
            fig_lst.update_layout(**BASE,height=280,bargap=0.3,xaxis=dict(showgrid=False,showticklabels=False,zeroline=False,range=[0,lst.max()*1.25] if len(lst) else [0,10]),yaxis=dict(showgrid=False,tickfont=dict(size=9,color=TXT2)))
            st.plotly_chart(fig_lst,use_container_width=True,config={"displayModeBar":False})
        with o2:
            st.markdown(ct("Stage Pipeline"),unsafe_allow_html=True)
            stg=ldf['stage_name'].value_counts()
            stg=stg[stg.index!='Unknown']
            stg_colors={'Sales Closed':C_GRN,'Registered':C_TEAL,'Pipeline':C_BLUE,'Prospect':C_AMB,
                        'High Priority':C_AMB,'Not Converted':C_RED,'Poor Lead':C_RED,
                        'Not Connected':C_GRAY,'Partially Converted':C_PURP}
            fig_stg=go.Figure(go.Pie(labels=stg.index,values=stg.values,hole=0.58,
                marker=dict(colors=[stg_colors.get(l,C_BLUE) for l in stg.index],line=dict(color=BG,width=2)),
                textinfo='percent',textfont=dict(size=9,family="DM Sans")))
            fig_stg.update_layout(**BASE,height=280,legend=dict(orientation="v",x=1.02,y=0.5,font=dict(size=9)),
                annotations=[dict(text="Stage",x=0.5,y=0.5,showarrow=False,font=dict(size=10,color=TXT2,family="DM Sans"))])
            st.plotly_chart(fig_stg,use_container_width=True,config={"displayModeBar":False})
        with o3:
            st.markdown(ct("Lead Owner — Volume & Conversions"),unsafe_allow_html=True)
            lo=ldf.groupby('leadownername').agg(Total=('_id','count'),Conv=('Is Converted','sum')).reset_index()
            lo=lo.sort_values('Total',ascending=True).tail(10)
            fig_lo=go.Figure()
            fig_lo.add_trace(go.Bar(name='Total',y=lo['leadownername'],x=lo['Total'],orientation='h',marker_color=C_BLUE,opacity=0.4,marker_line_width=0))
            fig_lo.add_trace(go.Bar(name='Converted',y=lo['leadownername'],x=lo['Conv'],orientation='h',marker_color=C_GRN,marker_line_width=0))
            fig_lo.update_layout(**BASE,height=280,barmode='overlay',bargap=0.28,legend=dict(orientation="h",x=0,y=1.12,font=dict(size=10),bgcolor="rgba(0,0,0,0)"),xaxis=dict(showgrid=False,showticklabels=False,zeroline=False),yaxis=dict(showgrid=False,tickfont=dict(size=9,color=TXT2)))
            st.plotly_chart(fig_lo,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

        # ── Attempted Analysis ────────────────────────────────────────────────
        st.markdown('<div class="asec"><div class="asec-t">📞 Attempt Rate & Service Analysis</div>',unsafe_allow_html=True)
        at1,at2,at3=st.columns([2,2,3],gap="medium")
        with at1:
            st.markdown(ct("Attempted vs Unattempted"),unsafe_allow_html=True)
            att=ldf['Attempted/Unattempted'].value_counts()
            fig_att=go.Figure(go.Pie(labels=att.index,values=att.values,hole=0.62,
                marker=dict(colors=[C_GRN,C_RED],line=dict(color=BG,width=3)),
                textinfo='label+percent',textfont=dict(size=11,family="DM Sans")))
            fig_att.update_layout(**BASE,height=200,showlegend=False,
                annotations=[dict(text="Attempt",x=0.5,y=0.5,showarrow=False,font=dict(size=10,color=TXT2,family="DM Sans"))])
            st.plotly_chart(fig_att,use_container_width=True,config={"displayModeBar":False})
        with at2:
            st.markdown(ct("Conversion Rate — Attempted vs Not"),unsafe_allow_html=True)
            atc=ldf.groupby('Attempted/Unattempted').agg(Total=('_id','count'),Conv=('Is Converted','sum')).reset_index()
            atc['Rate']=(atc['Conv']/atc['Total']*100).round(1)
            fig_atc=go.Figure(go.Bar(x=atc['Rate'],y=atc['Attempted/Unattempted'],orientation='h',
                marker_color=[C_GRN,C_RED],marker_line_width=0,
                text=[f"{r}%" for r in atc['Rate']],textposition='outside',textfont=dict(size=13,color=TXT2)))
            fig_atc.update_layout(**BASE,height=200,bargap=0.45,xaxis=dict(showgrid=False,showticklabels=False,zeroline=False,range=[0,atc['Rate'].max()*1.35] if len(atc) else [0,10]),yaxis=dict(showgrid=False,tickfont=dict(size=11,color=TXT2)))
            st.plotly_chart(fig_atc,use_container_width=True,config={"displayModeBar":False})
        with at3:
            st.markdown(ct("Top Services Enrolled"),unsafe_allow_html=True)
            svc=ldf[ldf['servicename']!='Unknown']['servicename'].value_counts().head(8)
            svc.index=[s[:32]+'…' if len(s)>32 else s for s in svc.index]
            fig_svc=go.Figure(go.Bar(x=svc.values,y=svc.index,orientation='h',
                marker=dict(color=svc.values,colorscale=[[0,C_BLUE],[1,C_TEAL]],line=dict(width=0)),
                text=svc.values,textposition='outside',textfont=dict(size=9,color=TXT2)))
            fig_svc.update_layout(**BASE,height=200,bargap=0.3,xaxis=dict(showgrid=False,showticklabels=False,zeroline=False,range=[0,svc.max()*1.25] if len(svc) else [0,10]),yaxis=dict(showgrid=False,tickfont=dict(size=9,color=TXT2)))
            st.plotly_chart(fig_svc,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — RECORDS & EXPORT
# ══════════════════════════════════════════════════════════════════════════════
with t_rec:
    rec_tab = st.radio("View records from:", ["Seminar Attendance","Leads"], horizontal=True, key="rec_src")

    if rec_tab == "Seminar Attendance":
        if not st.session_state.loaded_sem:
            st.markdown(empty_state("📊","No Seminar Data","Upload files in the Upload Files tab"), unsafe_allow_html=True)
        else:
            dfa = st.session_state.sem_conv.copy()
            st.markdown('<br><div class="asec"><div class="asec-t">🔍 Filter Seminar Records</div>',unsafe_allow_html=True)
            rf1,rf2,rf3,rf4=st.columns(4)
            rf5,rf6,rf7,rf8=st.columns(4)
            rf9,rf10,rf11,rf12=st.columns(4)
            t_place   = rf1.multiselect("📍 Location",         sorted(dfa['Place'].dropna().unique()),                               key="r_pl")
            t_trainer = rf2.multiselect("🎤 Trainer",          sorted(dfa['Trainer / Presenter'].dropna().unique()),                 key="r_tr")
            t_date    = rf3.multiselect("📅 Seminar Date",     [d.strftime("%d %b %Y") for d in sorted(dfa['Seminar Date'].dropna().unique())], key="r_dt")
            t_sess    = rf4.multiselect("🕐 Session",          ["MORNING","EVENING"],                                               key="r_se")
            t_conv    = rf5.multiselect("✅ Conversion",        ["Converted","Not Converted"],                                       key="r_cv")
            t_due     = rf6.multiselect("💳 Due Status",        sorted(dfa['Due Status'].dropna().unique()),                         key="r_du")
            t_course  = rf7.multiselect("📚 Course",            sorted(dfa['service_name'].dropna().unique()),                       key="r_co")
            t_rep     = rf8.multiselect("👤 Sales Rep",         sorted(dfa['sales_rep_name'].dropna().unique()),                     key="r_re")
            t_status  = rf9.multiselect("📌 Order Status",      sorted(dfa['status'].dropna().unique()),                            key="r_st")
            t_trader  = rf10.selectbox("🔄 Is Trader?",         ["All","Yes","No"],                                                 key="r_td")
            t_ourstu  = rf11.selectbox("🎓 Existing Student?",  ["All","Yes","No"],                                                 key="r_os")
            t_search  = rf12.text_input("🔍 Search Name/Phone", placeholder="Type to search…",                                     key="r_sr")
            st.markdown('</div>',unsafe_allow_html=True)

            dt=dfa.copy()
            if t_place:   dt=dt[dt['Place'].isin(t_place)]
            if t_trainer: dt=dt[dt['Trainer / Presenter'].isin(t_trainer)]
            if t_date:
                fmts=[pd.to_datetime(d,format="%d %b %Y") for d in t_date]; dt=dt[dt['Seminar Date'].isin(fmts)]
            if t_sess:    dt=dt[dt['Session'].isin(t_sess)]
            if t_conv:    dt=dt[dt['Conv Status'].isin(t_conv)]
            if t_due:     dt=dt[dt['Due Status'].isin(t_due)]
            if t_course:  dt=dt[dt['service_name'].isin(t_course)]
            if t_rep:     dt=dt[dt['sales_rep_name'].isin(t_rep)]
            if t_status:  dt=dt[dt['status'].isin(t_status)]
            if t_trader!="All": dt=dt[dt['TRADER'].isin(['YES','Y','TYES'] if t_trader=="Yes" else ['NO','N'])]
            if t_ourstu!="All": dt=dt[dt['Is our Student ?'].isin(['YES','STUDENT'] if t_ourstu=="Yes" else ['NO'])]
            if t_search:
                mask=pd.Series([False]*len(dt))
                for c in ['NAME','Mobile','email']:
                    if c in dt.columns: mask=mask|dt[c].astype(str).str.lower().str.contains(t_search.lower(),na=False)
                dt=dt[mask]

            NT=len(dt); ct_=( dt['Conv Status']=='Converted').sum(); rt_=dt['payment_received'].sum(); dt_=dt['total_due'].sum()
            if NT:
                st.markdown(f"""<div class="sbar">
  <div class="sbar-cell"><div class="sbar-lbl">Records</div><div class="sbar-val">{NT:,}</div></div>
  <div class="sbar-cell"><div class="sbar-lbl">Converted</div><div class="sbar-val" style="color:#4fce8f">{ct_:,}</div></div>
  <div class="sbar-cell"><div class="sbar-lbl">Not Converted</div><div class="sbar-val" style="color:#f76f4f">{NT-ct_:,}</div></div>
  <div class="sbar-cell"><div class="sbar-lbl">Collected</div><div class="sbar-val" style="color:#4fce8f">{fmt(rt_)}</div></div>
  <div class="sbar-cell"><div class="sbar-lbl">Total Due</div><div class="sbar-val" style="color:#f76f4f">{fmt(dt_)}</div></div>
  <div class="sbar-cell"><div class="sbar-lbl">Conv Rate</div><div class="sbar-val">{ct_/NT*100:.1f}%</div></div>
</div>""",unsafe_allow_html=True)

            disp=['NAME','Mobile','Place','Seminar Date','Session','Trainer / Presenter',
                  'Conv Status','Amount Paid','service_name','total_amount',
                  'payment_received','total_due','Due Status','status',
                  'sales_rep_name','trainer_name','TRADER','Is our Student ?','PM Label','Remarks']
            disp=[c for c in disp if c in dt.columns]
            ds=dt[disp].copy(); ds['Seminar Date']=ds['Seminar Date'].dt.strftime("%d %b %Y")

            st.markdown('<div class="asec"><div class="asec-t">📄 Attendee Records</div>',unsafe_allow_html=True)
            st.dataframe(ds,use_container_width=True,height=440,hide_index=True,
                column_config={"total_amount":st.column_config.NumberColumn("Course Amt",format="₹%.0f"),"payment_received":st.column_config.NumberColumn("Collected",format="₹%.0f"),"total_due":st.column_config.NumberColumn("Due",format="₹%.0f"),"Amount Paid":st.column_config.NumberColumn("Sem Paid",format="₹%.0f"),"service_name":st.column_config.TextColumn("Course"),"sales_rep_name":st.column_config.TextColumn("Sales Rep"),"trainer_name":st.column_config.TextColumn("Trainer (Order)"),"Trainer / Presenter":st.column_config.TextColumn("Seminar Trainer"),"PM Label":st.column_config.TextColumn("Payment Mode")})
            st.markdown('</div>',unsafe_allow_html=True)

            def to_xl(d): buf=io.BytesIO(); d.to_excel(buf,index=False); return buf.getvalue()
            ts=datetime.now().strftime("%Y%m%d_%H%M")
            st.markdown('<div class="asec"><div class="asec-t">⬇️ Export</div>',unsafe_allow_html=True)
            e1,e2,e3=st.columns(3)
            e1.download_button("⬇️ All Filtered",    to_xl(dt[disp]), file_name=f"seminar_{ts}.xlsx",   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            e2.download_button("⬇️ Converted Only",  to_xl(dt[dt['Conv Status']=='Converted'][disp]),   file_name=f"converted_{ts}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            e3.download_button("⬇️ Has Due Balance", to_xl(dt[dt['total_due']>0][disp]),                file_name=f"has_due_{ts}.xlsx",  mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            st.markdown('</div>',unsafe_allow_html=True)

    else:  # Leads records
        if not st.session_state.loaded_leads:
            st.markdown(empty_state("🎯","No Leads Data","Upload the Leads Report in the Upload Files tab"), unsafe_allow_html=True)
        else:
            ldf_all=st.session_state.leads.copy()
            st.markdown('<br><div class="asec"><div class="asec-t">🔍 Filter Lead Records</div>',unsafe_allow_html=True)
            lrf1,lrf2,lrf3,lrf4=st.columns(4)
            lrf5,lrf6,lrf7,lrf8=st.columns(4)
            lr_cf    = lrf1.multiselect("✅ Converted From",  ["Webinar","Non Webinar"],                                             key="lr_cf")
            lr_ls    = lrf2.multiselect("📢 Lead Source",     sorted([x for x in ldf_all['leadsource'].unique() if x!='Unknown']),  key="lr_ls")
            lr_cp    = lrf3.multiselect("📣 Campaign",        sorted([x for x in ldf_all['campaign_name'].unique() if x!='Unknown']),key="lr_cp")
            lr_lst   = lrf4.multiselect("📊 Lead Status",     sorted([x for x in ldf_all['leadstatus'].unique() if x!='Unknown']), key="lr_lst")
            lr_stg   = lrf5.multiselect("🏷 Stage",           sorted([x for x in ldf_all['stage_name'].unique() if x!='Unknown']), key="lr_stg")
            lr_lo    = lrf6.multiselect("👤 Lead Owner",      sorted([x for x in ldf_all['leadownername'].unique() if x!='Unknown']),key="lr_lo")
            lr_st    = lrf7.multiselect("📍 State",           sorted([x for x in ldf_all['state'].unique() if x!='Unknown']),      key="lr_st")
            lr_att   = lrf8.selectbox("📞 Attempted?",        ["All","Attempted","Unattempted"],                                    key="lr_att")
            lr_search= st.text_input("🔍 Search Name / Phone / Email", placeholder="Type to search…",                              key="lr_sr")
            st.markdown('</div>',unsafe_allow_html=True)

            ld=ldf_all.copy()
            if lr_cf:  ld=ld[ld['converted_from'].isin(lr_cf)]
            if lr_ls:  ld=ld[ld['leadsource'].isin(lr_ls)]
            if lr_cp:  ld=ld[ld['campaign_name'].isin(lr_cp)]
            if lr_lst: ld=ld[ld['leadstatus'].isin(lr_lst)]
            if lr_stg: ld=ld[ld['stage_name'].isin(lr_stg)]
            if lr_lo:  ld=ld[ld['leadownername'].isin(lr_lo)]
            if lr_st:  ld=ld[ld['state'].isin(lr_st)]
            if lr_att!="All": ld=ld[ld['Attempted/Unattempted']==lr_att]
            if lr_search:
                mask=pd.Series([False]*len(ld))
                for c in ['name','phone','email']:
                    if c in ld.columns: mask=mask|ld[c].astype(str).str.lower().str.contains(lr_search.lower(),na=False)
                ld=ld[mask]

            NLR=len(ld); cLR=ld['Is Converted'].sum()
            st.markdown(f"""<div class="sbar">
  <div class="sbar-cell"><div class="sbar-lbl">Records</div><div class="sbar-val">{NLR:,}</div></div>
  <div class="sbar-cell"><div class="sbar-lbl">Converted</div><div class="sbar-val" style="color:#4fce8f">{cLR:,}</div></div>
  <div class="sbar-cell"><div class="sbar-lbl">From Webinar</div><div class="sbar-val" style="color:#4fd8f7">{(ld['converted_from']=='Webinar').sum():,}</div></div>
  <div class="sbar-cell"><div class="sbar-lbl">From Non-Webinar</div><div class="sbar-val" style="color:#b44fe7">{(ld['converted_from']=='Non Webinar').sum():,}</div></div>
  <div class="sbar-cell"><div class="sbar-lbl">Conv Rate</div><div class="sbar-val">{cLR/NLR*100:.1f}%</div></div>
</div>""" if NLR else "",unsafe_allow_html=True)

            lead_disp_cols=['name','phone','email','leaddate','converted_from','leadsource',
                            'campaign_name','leadstatus','stage_name','leadownername',
                            'servicename','state','Attempted/Unattempted','remarks']
            lead_disp_cols=[c for c in lead_disp_cols if c in ld.columns]
            ld_show=ld[lead_disp_cols].copy()
            ld_show['leaddate']=ld_show['leaddate'].dt.strftime("%d %b %Y")

            st.markdown('<div class="asec"><div class="asec-t">📄 Lead Records</div>',unsafe_allow_html=True)
            st.dataframe(ld_show,use_container_width=True,height=440,hide_index=True,
                column_config={"name":st.column_config.TextColumn("Name"),"leaddate":st.column_config.TextColumn("Lead Date"),"converted_from":st.column_config.TextColumn("Converted From"),"leadsource":st.column_config.TextColumn("Lead Source"),"campaign_name":st.column_config.TextColumn("Campaign"),"leadstatus":st.column_config.TextColumn("Status"),"stage_name":st.column_config.TextColumn("Stage"),"leadownername":st.column_config.TextColumn("Lead Owner"),"servicename":st.column_config.TextColumn("Service"),"Attempted/Unattempted":st.column_config.TextColumn("Attempted?")})
            st.markdown('</div>',unsafe_allow_html=True)

            def to_xl(d): buf=io.BytesIO(); d.to_excel(buf,index=False); return buf.getvalue()
            ts=datetime.now().strftime("%Y%m%d_%H%M")
            st.markdown('<div class="asec"><div class="asec-t">⬇️ Export</div>',unsafe_allow_html=True)
            e1,e2,e3=st.columns(3)
            e1.download_button("⬇️ All Filtered Leads",       to_xl(ld[lead_disp_cols]), file_name=f"leads_{ts}.xlsx",         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            e2.download_button("⬇️ Webinar Leads Only",       to_xl(ld[ld['converted_from']=='Webinar'][lead_disp_cols]),      file_name=f"webinar_leads_{ts}.xlsx",    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            e3.download_button("⬇️ Non-Webinar Leads Only",   to_xl(ld[ld['converted_from']=='Non Webinar'][lead_disp_cols]), file_name=f"non_webinar_leads_{ts}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            st.markdown('</div>',unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)
