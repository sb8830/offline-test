"""
Invesmate Seminar Intelligence — Master Dashboard
Files: Seminar CSV + Conversion List XLSX + Leads XLSX
"""
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io, warnings
from datetime import datetime
warnings.filterwarnings('ignore')

st.set_page_config(page_title="Seminar Intelligence", page_icon="📊", layout="wide",
                   initial_sidebar_state="collapsed")

st.markdown("""
<style>
  #MainMenu,footer,header{visibility:hidden}
  .block-container{padding:0!important;max-width:100%!important}
  .stApp{background:#060910}
  div[data-testid="stToolbar"],div[data-testid="stDecoration"],
  div[data-testid="stStatusWidget"],section[data-testid="stSidebar"]{display:none!important}
</style>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# DESIGN SYSTEM — Invesmate exact style
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@400;500;600&display=swap" rel="stylesheet">
<style>
/* NAV */
.im-nav{background:linear-gradient(180deg,#0d1119,#080b12);border-bottom:1px solid rgba(255,255,255,.07);
  padding:0 24px;height:60px;display:flex;align-items:center;justify-content:space-between;
  position:sticky;top:0;z-index:9999}
.im-brand{font-family:'Syne',sans-serif;font-size:16px;font-weight:800;color:#eceef5;letter-spacing:-.3px;line-height:1.1}
.im-brand-sub{font-size:9px;color:#4a5068;text-transform:uppercase;letter-spacing:.9px}
.im-pill{background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.08);border-radius:20px;
  padding:4px 14px 4px 10px;display:flex;align-items:center;gap:8px;font-size:12px;color:#8a90aa}
.im-dot{width:7px;height:7px;background:#4fce8f;border-radius:50%;animation:blink 2s infinite;flex-shrink:0}
@keyframes blink{0%,100%{opacity:1}50%{opacity:.3}}

/* STAT CARDS */
.sg{display:grid;gap:10px}
.sc{background:#111520;border:1px solid rgba(255,255,255,.06);border-radius:12px;
  padding:14px 16px;position:relative;overflow:hidden}
.sc::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:var(--c,#4f8ef7)}
.sv{font-family:'Syne',sans-serif;font-size:24px;font-weight:800;color:#eceef5;line-height:1}
.sv.sm{font-size:17px}.sl{font-size:10px;color:#4a5068;text-transform:uppercase;letter-spacing:.5px;margin-top:5px}
.sd{font-size:10px;margin-top:3px}
.sd.g{color:#4fce8f}.sd.r{color:#f76f4f}.sd.a{color:#f7c948}.sd.n{color:#4a5068}.sd.b{color:#4f8ef7}

/* SECTIONS */
.asec{background:#0c1018;border:1px solid rgba(255,255,255,.07);border-radius:14px;padding:20px 22px;margin-bottom:14px}
.asec-t{font-family:'Syne',sans-serif;font-size:10px;font-weight:700;color:#f7c948;
  margin-bottom:14px;text-transform:uppercase;letter-spacing:.9px}
.ct{font-family:'Syne',sans-serif;font-size:9px;font-weight:700;color:#f7c948;
  text-transform:uppercase;letter-spacing:.9px;margin-bottom:10px}

/* STATS BAR */
.sbar{display:flex;gap:1px;background:rgba(255,255,255,.04);border-radius:10px;overflow:hidden;margin-bottom:12px}
.sbar-cell{flex:1;background:#111520;padding:.65rem .9rem}
.sbar-lbl{font-size:8.5px;color:#4a5068;text-transform:uppercase;letter-spacing:.5px;margin-bottom:3px}
.sbar-val{font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#eceef5}

/* FILE CARDS */
.fcard{background:#0c1018;border:1px solid rgba(255,255,255,.07);border-radius:12px;padding:13px;margin-bottom:8px}
.fcard-t{font-family:'Syne',sans-serif;font-size:11px;font-weight:700;color:#eceef5;margin:5px 0 2px}
.fcard-s{font-size:10px;color:#4a5068;line-height:1.5}

/* WIDGETS */
div[data-baseweb="select"]>div{background:#111520!important;border-color:rgba(255,255,255,.08)!important;border-radius:8px!important;font-family:'DM Sans',sans-serif!important;font-size:12px!important}
div[data-baseweb="select"]>div:hover{border-color:rgba(79,206,143,.3)!important}
div[data-baseweb="input"]>div{background:#111520!important;border-color:rgba(255,255,255,.08)!important;border-radius:8px!important}
label[data-testid="stWidgetLabel"]{font-family:'DM Sans',sans-serif!important;font-size:10.5px!important;color:#4a5068!important}
.stMultiSelect [data-baseweb="tag"]{background:rgba(79,206,143,.12)!important;border:1px solid rgba(79,206,143,.25)!important;border-radius:5px!important}
.stTabs [data-baseweb="tab-list"]{background:transparent!important;border-bottom:1px solid rgba(255,255,255,.07)!important;gap:0}
.stTabs [data-baseweb="tab"]{font-family:'DM Sans',sans-serif!important;font-size:12px!important;font-weight:500!important;
  color:#4a5068!important;padding:.6rem 1.2rem!important;border-bottom:2px solid transparent!important;background:transparent!important}
.stTabs [aria-selected="true"]{color:#4fce8f!important;border-bottom-color:#4fce8f!important}
.stExpander{border:1px solid rgba(255,255,255,.07)!important;border-radius:12px!important;background:#0c1018!important}
.stExpander summary p{font-family:'DM Sans',sans-serif!important;font-size:11.5px!important;color:#8a90aa!important}
button[data-testid="stBaseButton-secondary"]{background:#111520!important;border:1px solid rgba(255,255,255,.1)!important;
  color:#8a90aa!important;border-radius:8px!important;font-family:'DM Sans',sans-serif!important;font-size:12px!important}
button[data-testid="stBaseButton-secondary"]:hover{border-color:rgba(79,206,143,.4)!important;color:#eceef5!important}
button[data-testid="stBaseButton-primary"]{background:linear-gradient(135deg,#4fce8f,#3ab87a)!important;
  border:none!important;color:#060910!important;font-weight:700!important;border-radius:8px!important;font-family:'DM Sans',sans-serif!important}
div[data-testid="metric-container"]{background:#111520!important;border:1px solid rgba(255,255,255,.07)!important;border-radius:12px!important;padding:.8rem 1rem!important}
div[data-testid="stMetricValue"]{font-family:'Syne',sans-serif!important;font-size:20px!important;font-weight:800!important;color:#eceef5!important}
div[data-testid="stMetricLabel"]{font-size:9.5px!important;color:#4a5068!important;text-transform:uppercase!important;letter-spacing:.5px!important}
div[data-testid="stMetricDelta"] svg{display:none!important}
.stSuccess{background:rgba(79,206,143,.07)!important;border-left:3px solid #4fce8f!important;border-radius:8px!important}
.stInfo{background:rgba(79,142,247,.07)!important;border-left:3px solid #4f8ef7!important;border-radius:8px!important}
.stWarning{background:rgba(247,201,72,.07)!important;border-left:3px solid #f7c948!important;border-radius:8px!important}
::-webkit-scrollbar{width:4px;height:4px}::-webkit-scrollbar-track{background:#060910}
::-webkit-scrollbar-thumb{background:rgba(79,206,143,.2);border-radius:2px}
.page{padding:0 20px 60px 20px}
</style>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
BG="#0c1018"; GRID="rgba(255,255,255,0.04)"; TXT="#4a5068"; TXT2="#8a90aa"
G="#4fce8f"; B="#4f8ef7"; R="#f76f4f"; A="#f7c948"; P="#b44fe7"; T="#4fd8f7"; GR="#2d3748"
BASE=dict(paper_bgcolor=BG,plot_bgcolor=BG,font=dict(family="DM Sans",color=TXT,size=11),margin=dict(l=0,r=8,t=4,b=0))
PM={"mode1":"Full Payment","mode2":"Instalment","mode3":"EMI","mode4":"Partial","mode13":"Scholarship","mode5":"Other"}

def fmt(v):
    if pd.isna(v) or v==0: return "₹0"
    if v>=1e7: return f"₹{v/1e7:.2f}Cr"
    if v>=1e5: return f"₹{v/1e5:.1f}L"
    return f"₹{v:,.0f}"

def sc(label,val,delta="",c=B,dc="n",sm=False):
    return f"""<div class="sc" style="--c:{c}"><div class="sv{' sm' if sm else ''}">{val}</div>
<div class="sl">{label}</div>{"" if not delta else f'<div class="sd {dc}">{delta}</div>'}</div>"""

def ct(t): return f'<div class="ct">{t}</div>'

def nodata(icon,title,sub):
    return f"""<div style="padding:60px;text-align:center">
<div style="font-size:40px;margin-bottom:14px">{icon}</div>
<div style="font-family:'Syne',sans-serif;font-size:18px;font-weight:800;color:#eceef5;margin-bottom:8px">{title}</div>
<div style="color:#4a5068;font-size:13px">{sub}</div></div>"""


def ensure_columns(df, defaults):
    for col, default in defaults.items():
        if col not in df.columns:
            df[col] = default() if callable(default) else default
    return df


def safe_read_excel_bytes(file_bytes, preferred_sheet=None):
    bio = io.BytesIO(file_bytes)
    if preferred_sheet is not None:
        try:
            return pd.read_excel(bio, sheet_name=preferred_sheet)
        except Exception:
            bio.seek(0)
    return pd.read_excel(bio)


def safe_date_bounds(series):
    valid = pd.to_datetime(series, errors='coerce').dropna()
    today = pd.Timestamp.today().date()
    if valid.empty:
        return today, today
    return valid.min().date(), valid.max().date()

def bar_h(x,y,colors=None,single_color=B,text=None,height=260,gap=0.3,rng_mult=1.25):
    c=colors if colors else single_color
    fig=go.Figure(go.Bar(x=x,y=y,orientation='h',marker_color=c,marker_line_width=0,
        text=text,textposition='outside' if text else 'none',textfont=dict(size=9,color=TXT2),
        hovertemplate='<b>%{y}</b><br>%{x}<extra></extra>'))
    mx=max(x) if len(x) else 10
    fig.update_layout(**BASE,height=height,bargap=gap,
        xaxis=dict(showgrid=False,showticklabels=False,zeroline=False,range=[0,mx*rng_mult]),
        yaxis=dict(showgrid=False,tickfont=dict(size=9,color=TXT2)))
    return fig

def donut(labels,values,colors,text="",height=220):
    fig=go.Figure(go.Pie(labels=labels,values=values,hole=0.62,
        marker=dict(colors=colors,line=dict(color=BG,width=3)),
        textinfo='percent',textfont=dict(size=10,family="DM Sans"),
        hovertemplate='<b>%{label}</b><br>%{value}<br>%{percent}<extra></extra>'))
    fig.update_layout(**BASE,height=height,showlegend=True,
        legend=dict(orientation="v",x=1.02,y=0.5,font=dict(size=9),bgcolor="rgba(0,0,0,0)"),
        annotations=[dict(text=text,x=0.5,y=0.5,showarrow=False,font=dict(size=11,color=TXT2,family="DM Sans"))] if text else [])
    return fig

# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
for k in ["sem_df","conv_df","leads_df","master","loaded"]:
    if k not in st.session_state: st.session_state[k]=None
if "loaded" not in st.session_state: st.session_state.loaded=False

# ══════════════════════════════════════════════════════════════════════════════
# DATA PROCESSORS
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def process_all(sb,cb,lb):
    # ── SEMINAR ──────────────────────────────────────────────────────────────
    sem=pd.read_csv(io.BytesIO(sb)); sem.columns=sem.columns.str.strip()
    sem = ensure_columns(sem, {
        'NAME':'Unknown','Mobile':'','Place':'Unknown','Trainer / Presenter':'Unknown',
        'Seminar Date':pd.NaT,'Session':'Unknown','Is Attended ?':'NO','Is Converted ?':'NO',
        'Amount Paid':0,'Mode of Payment':'Unknown','TRADER':'NO','Is our Student ?':'NO','Remarks':''
    })
    for c in ['Is Attended ?','Is Converted ?','Session','TRADER','Is our Student ?','Trainer / Presenter','Place']:
        sem[c]=sem[c].astype(str).str.strip().str.upper()
    sem['Seminar Date']=pd.to_datetime(sem['Seminar Date'],errors='coerce',dayfirst=True)
    sem['Amount Paid']=pd.to_numeric(sem['Amount Paid'],errors='coerce').fillna(0)
    sem['Mode of Payment']=sem['Mode of Payment'].astype(str).str.strip().str.upper()
    sem['mob']=sem['Mobile'].astype(str).str.replace(r'\D','',regex=True).str[-10:]
    # Normalise trainer names
    trainer_map={'HIRAMNOY LAHERI/PRATIM CHAKRABORTY':'HIRANMOY LAHIRI & PRATIM CHAKRABORTY',
                 'PRATIM KUMAR CHAKRABORTY/MIHIR KANTI CHAKRABORTY':'PRATIM CHAKRABORTY & MIHIR KANTI CHAKRABORTY',
                 'PRATIM KUMAR CHAKRABORTY & AKASH MISHRA':'PRATIM CHAKRABORTY & AKASH MISHRA',
                 'PRFATIM CHAKRABORTY & AKASH MISHRA':'PRATIM CHAKRABORTY & AKASH MISHRA',
                 'PRATIM CHAKRABORTY & MIHIR KANTI CHAKRABORTYPARIJAT':'PRATIM CHAKRABORTY & MIHIR KANTI CHAKRABORTY'}
    sem['Trainer Norm']=sem['Trainer / Presenter'].replace(trainer_map)
    att=sem[sem['Is Attended ?']=='YES'].copy().reset_index(drop=True)
    att['Conv Status']=att['Is Converted ?'].apply(lambda x:'Converted' if x in ['CONVERTED','YES'] else 'Not Converted')

    # ── CONVERSION ───────────────────────────────────────────────────────────
    conv=safe_read_excel_bytes(cb)
    conv.columns=conv.columns.str.strip()
    conv = ensure_columns(conv, {
        '_id': lambda: range(1, len(conv)+1), 'phone':'', 'payment_received':0, 'total_amount':0,
        'total_due':0, 'total_gst':0, 'order_date':pd.NaT, 'payment_mode':'Unknown',
        'service_name':'Unknown', 'sales_rep_name':'Unknown', 'trainer':'Unknown', 'status':'Unknown',
        'student_name':'Unknown', 'email':'', 'batch_date':pd.NaT, 'is_refunded':False, 'is_shortClosed':False
    })
    for c in ['payment_received','total_amount','total_due','total_gst']:
        conv[c]=pd.to_numeric(conv[c],errors='coerce').fillna(0)
    conv['order_date']=pd.to_datetime(conv['order_date'],errors='coerce',utc=True).dt.tz_localize(None)
    conv['phone_clean']=conv['phone'].astype(str).str.replace(r'\D','',regex=True).str[-10:]
    conv['PM Label']=conv['payment_mode'].map(PM).fillna(conv['payment_mode'])
    conv['service_name']=conv['service_name'].astype(str).str.strip()
    conv['sales_rep_name']=conv['sales_rep_name'].astype(str).str.strip()
    conv['trainer_name']=conv['trainer'].astype(str).str.split(' - ').str[-1].str.strip()
    conv['month']=conv['order_date'].dt.to_period('M').astype(str)
    conv['is_refunded']=conv['is_refunded'].fillna(False)
    conv['is_shortClosed']=conv['is_shortClosed'].fillna(False)

    # ── LEADS ─────────────────────────────────────────────────────────────────
    leads=safe_read_excel_bytes(lb, preferred_sheet='Sheet 1')
    leads.columns=leads.columns.str.strip()
    leads = ensure_columns(leads, {
        '_id': lambda: range(1, len(leads)+1), 'leaddate':pd.NaT, 'converted_from':'Unknown',
        'leadsource':'Unknown', 'campaign_name':'Unknown', 'leadstatus':'Unknown', 'stage_name':'Unknown',
        'leadownername':'Unknown', 'state':'Unknown', 'Attempted/Unattempted':'Unknown', 'servicename':'Unknown',
        'phone':'', 'name':'Unknown', 'email':'', 'remarks':''
    })
    leads['leaddate']=pd.to_datetime(leads['leaddate'],errors='coerce')
    leads['Lead Month']=leads['leaddate'].dt.to_period('M').astype(str)
    for c in ['converted_from','leadsource','campaign_name','leadstatus','stage_name','leadownername','state','Attempted/Unattempted','servicename']:
        leads[c]=leads[c].astype(str).str.strip().replace('nan','Unknown')
    leads['converted_from'] = leads['converted_from'].replace({
        'Non-Webinar':'Non Webinar', 'NON WEBINAR':'Non Webinar', 'WEBINAR':'Webinar', 'NONWEBINAR':'Non Webinar'
    })
    leads['Attempted/Unattempted'] = leads['Attempted/Unattempted'].replace({
        'ATTEMPTED':'Attempted', 'UNATTEMPTED':'Unattempted'
    })
    leads['phone_clean']=leads['phone'].astype(str).str.replace(r'\D','',regex=True).str[-10:]
    leads['Is Converted']=leads['leadstatus'].str.lower().str.contains('converted',na=False)

    # ── MERGE: Attendance → Conversion ───────────────────────────────────────
    mg=att.merge(conv[['phone_clean','orderID','order_date','service_code','service_name',
                        'payment_received','total_amount','total_due','total_gst',
                        'payment_mode','PM Label','status','sales_rep_name',
                        'trainer','trainer_name','student_invid','batch_date','month']],
                 left_on='mob',right_on='phone_clean',how='left')
    def due_tag(r):
        if pd.isna(r.get('total_due')): return 'No Order'
        if r['total_due']<=0:           return 'Fully Paid'
        if r['total_amount']>0 and r['total_due']<r['total_amount']: return 'Partially Paid'
        return 'Fully Due'
    mg['Due Status']=mg.apply(due_tag,axis=1)

    # ── MERGE: Attendance → Leads (for cross-enrichment) ──────────────────────
    mg=mg.merge(leads[['phone_clean','converted_from','leadsource','campaign_name',
                        'leadstatus','stage_name','leadownername']].drop_duplicates('phone_clean'),
                left_on='mob',right_on='phone_clean',how='left',suffixes=('','_lead'))
    mg['converted_from']=mg['converted_from'].fillna('Unknown')
    mg['leadsource']=mg['leadsource'].fillna('Unknown')

    return {'att':mg,'conv':conv,'leads':leads,'sem':sem}

# ══════════════════════════════════════════════════════════════════════════════
# NAV BAR
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="im-nav">
  <div style="display:flex;align-items:center;gap:11px">
    <div style="width:36px;height:36px;border-radius:50%;
      background:linear-gradient(135deg,#4fce8f,#4f8ef7);
      display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0">📊</div>
    <div><div class="im-brand">Invesmate</div>
    <div class="im-brand-sub">Seminar Intelligence Hub</div></div>
  </div>
  <div class="im-pill"><div class="im-dot"></div>
  <span>{datetime.now().strftime("%d %b %Y  %H:%M")}</span></div>
</div><div class="page">""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
t_up,t_overview,t_sem,t_leads,t_conv,t_records=st.tabs([
    "📁 Upload","🏠 Overview","🎯 Seminar","📣 Leads","💰 Conversion","📋 Records"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
with t_up:
    st.markdown("""<div style="text-align:center;padding:32px 20px 22px">
<div style="font-size:34px">📊</div>
<div style="font-family:'Syne',sans-serif;font-size:30px;font-weight:800;
  color:#eceef5;margin:10px 0 6px;letter-spacing:-1px">Seminar Intelligence Hub</div>
<div style="color:#4a5068;font-size:11px;text-transform:uppercase;letter-spacing:.8px">
  Upload 3 files · Auto-merge · Deep KPI Analysis</div></div>""", unsafe_allow_html=True)

    st.markdown('<div class="asec"><div class="asec-t">📂 Upload Your 3 Data Files</div>', unsafe_allow_html=True)
    u1,u2,u3=st.columns(3,gap="medium")
    with u1:
        st.markdown("""<div class="fcard"><span style="font-size:20px">📝</span>
<div class="fcard-t">Seminar Updated Sheet (.csv)</div>
<div class="fcard-s">NAME · Mobile · Place · Trainer · Seminar Date · Session · Is Attended · Is Converted · Amount Paid · Mode of Payment</div></div>""",unsafe_allow_html=True)
        f_sem=st.file_uploader("sem",type=["csv"],key="up_sem",label_visibility="collapsed")
    with u2:
        st.markdown("""<div class="fcard"><span style="font-size:20px">📋</span>
<div class="fcard-t">Conversion List (.xlsx)</div>
<div class="fcard-s">phone · service_name · total_amount · total_due · payment_received · status · sales_rep_name · trainer · payment_mode · order_date</div></div>""",unsafe_allow_html=True)
        f_conv=st.file_uploader("conv",type=["xlsx","xls"],key="up_conv",label_visibility="collapsed")
    with u3:
        st.markdown("""<div class="fcard"><span style="font-size:20px">🎯</span>
<div class="fcard-t">Leads Report (.xlsx)</div>
<div class="fcard-s">converted_from (Webinar/Non-Webinar) · leadsource · campaign_name · leadstatus · stage_name · leadownername · state · Attempted/Unattempted</div></div>""",unsafe_allow_html=True)
        f_leads=st.file_uploader("leads",type=["xlsx","xls"],key="up_leads",label_visibility="collapsed")
    st.markdown('</div>',unsafe_allow_html=True)

    _,cb,_=st.columns([1,2,1])
    with cb:
        all3=f_sem and f_conv and f_leads
        if all3:
            if st.button("🚀  Generate Full Dashboard",use_container_width=True,type="primary",key="gen"):
                with st.spinner("Processing and merging all 3 files…"):
                    data=process_all(f_sem.read(),f_conv.read(),f_leads.read())
                st.session_state.master=data; st.session_state.loaded=True
                att=data['att']; conv=data['conv']; leads=data['leads']
                st.success(f"✅ Ready — {len(att):,} attendees · {(att['Conv Status']=='Converted').sum():,} converted · {len(conv):,} orders · {len(leads):,} leads")
                m1,m2,m3,m4,m5,m6=st.columns(6)
                m1.metric("Attendees",f"{len(att):,}")
                m2.metric("Converted",(att['Conv Status']=='Converted').sum())
                m3.metric("Orders",f"{len(conv):,}")
                m4.metric("Leads",f"{len(leads):,}")
                m5.metric("Gross Revenue",fmt(conv['total_amount'].sum()))
                m6.metric("Locations",att['Place'].nunique())
        else:
            missing=[n for n,f in [("Seminar CSV",f_sem),("Conversion List",f_conv),("Leads Report",f_leads)] if not f]
            st.markdown(f"""<div style="text-align:center;padding:14px;background:rgba(79,206,143,.04);
  border:1px dashed rgba(79,206,143,.2);border-radius:10px;color:#4a5068;font-size:13px">
  Waiting for: <strong style="color:#8a90aa">{' · '.join(missing)}</strong></div>""",unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SHARED FILTER HELPER
# ══════════════════════════════════════════════════════════════════════════════
def apply_sem_filters(df,prefix):
    with st.expander("🔽  Filters",expanded=True):
        c1,c2,c3,c4=st.columns(4); c5,c6,c7,c8=st.columns(4); c9,c10,c11,_=st.columns(4)
        sp=prefix
        sel_place  =c1.multiselect("📍 Location",      sorted(df['Place'].dropna().unique()),                                        key=f"{sp}pl")
        sel_trainer=c2.multiselect("🎤 Trainer",        sorted(df['Trainer Norm'].dropna().unique()),                                 key=f"{sp}tr")
        sel_date   =c3.multiselect("📅 Seminar Date",   [d.strftime("%d %b %Y") for d in sorted(df['Seminar Date'].dropna().unique())],key=f"{sp}dt")
        sel_sess   =c4.multiselect("🕐 Session",        ["MORNING","EVENING"],                                                       key=f"{sp}se")
        sel_conv   =c5.multiselect("✅ Conversion",      ["Converted","Not Converted"],                                               key=f"{sp}cv")
        sel_due    =c6.multiselect("💳 Due Status",      sorted(df['Due Status'].dropna().unique()),                                  key=f"{sp}du")
        sel_course =c7.multiselect("📚 Course",          sorted(df['service_name'].dropna().unique()),                                key=f"{sp}co")
        sel_rep    =c8.multiselect("👤 Sales Rep",       sorted(df['sales_rep_name'].dropna().unique()),                             key=f"{sp}re")
        sel_st     =c9.multiselect("📌 Order Status",    sorted(df['status'].dropna().unique()),                                     key=f"{sp}os")
        sel_trader =c10.selectbox("🔄 Trader?",         ["All","Yes","No"],                                                         key=f"{sp}td")
        sel_ourstu =c11.selectbox("🎓 Existing Student?",["All","Yes","No"],                                                         key=f"{sp}es")
    out=df.copy()
    if sel_place:   out=out[out['Place'].isin(sel_place)]
    if sel_trainer: out=out[out['Trainer Norm'].isin(sel_trainer)]
    if sel_date:
        fmts=[pd.to_datetime(d,format="%d %b %Y") for d in sel_date]; out=out[out['Seminar Date'].isin(fmts)]
    if sel_sess:    out=out[out['Session'].isin(sel_sess)]
    if sel_conv:    out=out[out['Conv Status'].isin(sel_conv)]
    if sel_due:     out=out[out['Due Status'].isin(sel_due)]
    if sel_course:  out=out[out['service_name'].isin(sel_course)]
    if sel_rep:     out=out[out['sales_rep_name'].isin(sel_rep)]
    if sel_st:      out=out[out['status'].isin(sel_st)]
    if sel_trader!="All": out=out[out['TRADER'].isin(['YES','TYES'] if sel_trader=="Yes" else ['NO'])]
    if sel_ourstu!="All": out=out[out['Is our Student ?'].isin(['YES','STUDENT'] if sel_ourstu=="Yes" else ['NO'])]
    return out

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — OVERVIEW (all 3 combined top-level KPIs)
# ══════════════════════════════════════════════════════════════════════════════
with t_overview:
    if not st.session_state.loaded:
        st.markdown(nodata("📊","No Data Loaded","Upload all 3 files in the Upload tab and click Generate"),unsafe_allow_html=True)
    else:
        D=st.session_state.master
        att=D['att']; conv=D['conv']; leads=D['leads']
        conv_n=(att['Conv Status']=='Converted').sum()
        N=len(att)
        conv_r=conv_n/N*100 if N else 0
        rcvd=conv['payment_received'].sum(); rev=conv['total_amount'].sum(); due=conv['total_due'].sum()
        lead_conv=leads['Is Converted'].sum(); lead_n=len(leads)
        web_leads=(leads['converted_from']=='Webinar').sum()
        nonweb_leads=(leads['converted_from']=='Non Webinar').sum()
        active_orders=(conv['status']=='Active').sum()
        fp=(att['Due Status']=='Fully Paid').sum()

        st.markdown("<br>",unsafe_allow_html=True)
        # ── Row 1: 8 master KPIs ─────────────────────────────────────────────
        st.markdown(f"""<div class="sg" style="grid-template-columns:repeat(8,minmax(0,1fr))">
{sc("Seminar Attendees",f"{N:,}",f"of {len(D['sem']):,} registered",B,"b")}
{sc("Seminar Converted",f"{conv_n:,}",f"↑ {conv_r:.1f}% rate",G,"g")}
{sc("Not Converted",f"{N-conv_n:,}",f"{100-conv_r:.1f}% of attended",R,"r")}
{sc("Gross Revenue",fmt(rev),"Conversion list",P,"n",True)}
{sc("Collected",fmt(rcvd),f"Due: {fmt(due)}",G,"g",True)}
{sc("Total Leads",f"{lead_n:,}",f"{lead_conv:,} converted ({lead_conv/lead_n*100:.0f}%)",T,"b")}
{sc("Webinar Leads",f"{web_leads:,}",f"{web_leads/lead_n*100:.0f}% of leads",G,"g")}
{sc("Non-Webinar Leads",f"{nonweb_leads:,}",f"{nonweb_leads/lead_n*100:.0f}% of leads",P,"n")}
</div>""",unsafe_allow_html=True)
        st.markdown("<br>",unsafe_allow_html=True)

        # ── Row 2: 6 more KPIs ───────────────────────────────────────────────
        st.markdown(f"""<div class="sg" style="grid-template-columns:repeat(6,minmax(0,1fr))">
{sc("Active Orders",f"{active_orders:,}",f"Inactive: {(conv['status']=='Inactive').sum():,}",G,"g")}
{sc("Fully Paid",f"{fp:,}",f"Has balance: {(att['Due Status'].isin(['Partially Paid','Fully Due'])).sum():,}",A,"a")}
{sc("Seminar Locations",f"{att['Place'].nunique()}","Offline venues",B,"n")}
{sc("Seminar Amount",fmt(att[att['Conv Status']=='Converted']['Amount Paid'].sum()),"Collected at venue",G,"g",True)}
{sc("Active Sales Reps",f"{conv['sales_rep_name'].nunique()}","In conversion list",T,"n")}
{sc("Lead Owners",f"{leads['leadownername'].nunique()}","Following leads",P,"n")}
</div>""",unsafe_allow_html=True)
        st.markdown("<br>",unsafe_allow_html=True)

        # ── Overview Charts Row 1 ─────────────────────────────────────────────
        st.markdown('<div class="asec"><div class="asec-t">🔀 Cross-File Intelligence</div>',unsafe_allow_html=True)
        ov1,ov2,ov3=st.columns(3,gap="medium")

        with ov1:
            st.markdown(ct("Seminar Conversion Funnel"),unsafe_allow_html=True)
            fig=go.Figure(go.Funnel(y=["Registered","Attended","Converted"],x=[len(D['sem']),N,conv_n],
                textinfo="value+percent initial",textfont=dict(family="DM Sans",size=11),
                marker=dict(color=[B,P,G],line=dict(color=BG,width=2)),
                connector=dict(line=dict(color=GRID,width=1))))
            fig.update_layout(**BASE,height=230)
            st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})

        with ov2:
            st.markdown(ct("Lead Stage Pipeline"),unsafe_allow_html=True)
            stg=leads['stage_name'].value_counts(); stg=stg[stg.index!='Unknown']
            stg_c={'Sales Closed':G,'Registered':T,'Pipeline':B,'Prospect':A,'High Priority':A,
                   'Not Converted':R,'Poor Lead':R,'Not Connected':GR,'Partially Converted':P}
            fig2=go.Figure(go.Pie(labels=stg.index,values=stg.values,hole=0.58,
                marker=dict(colors=[stg_c.get(l,B) for l in stg.index],line=dict(color=BG,width=2)),
                textinfo='percent',textfont=dict(size=9,family="DM Sans")))
            fig2.update_layout(**BASE,height=230,showlegend=True,
                legend=dict(orientation="v",x=1.02,y=0.5,font=dict(size=8.5),bgcolor="rgba(0,0,0,0)"),
                annotations=[dict(text="Stage",x=0.5,y=0.5,showarrow=False,font=dict(size=10,color=TXT2,family="DM Sans"))])
            st.plotly_chart(fig2,use_container_width=True,config={"displayModeBar":False})

        with ov3:
            st.markdown(ct("Monthly Revenue — Conversion List"),unsafe_allow_html=True)
            mo=conv.groupby('month').agg(n=('_id','count'),rev=('total_amount','sum')).reset_index().sort_values('month').tail(6)
            fig3=make_subplots(specs=[[{"secondary_y":True}]])
            fig3.add_trace(go.Bar(name='Orders',x=mo['month'],y=mo['n'],marker_color=B,marker_line_width=0,opacity=0.6),secondary_y=False)
            fig3.add_trace(go.Scatter(name='Revenue',x=mo['month'],y=mo['rev'],line=dict(color=G,width=2.5,shape='spline'),mode='lines+markers',marker=dict(size=6,color=G,line=dict(color=BG,width=1.5))),secondary_y=True)
            fig3.update_layout(**BASE,height=230,bargap=0.3,
                legend=dict(orientation="h",x=0,y=1.12,font=dict(size=10),bgcolor="rgba(0,0,0,0)"),
                xaxis=dict(showgrid=False,tickfont=dict(size=9,color=TXT2)),
                yaxis=dict(gridcolor=GRID,tickfont=dict(size=9,color=TXT),zeroline=False),
                yaxis2=dict(showgrid=False,tickfont=dict(size=9,color=G),tickformat=",.0f"))
            st.plotly_chart(fig3,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

        # ── Overview Charts Row 2 ─────────────────────────────────────────────
        st.markdown('<div class="asec"><div class="asec-t">📊 Performance Snapshot</div>',unsafe_allow_html=True)
        ov4,ov5,ov6,ov7=st.columns(4,gap="medium")

        with ov4:
            st.markdown(ct("Order Status"),unsafe_allow_html=True)
            os=conv['status'].value_counts()
            fig4=go.Figure(go.Pie(labels=os.index,values=os.values,hole=0.62,
                marker=dict(colors=[G,R,A,GR],line=dict(color=BG,width=3)),
                textinfo='percent',textfont=dict(size=10,family="DM Sans")))
            fig4.update_layout(**BASE,height=200,legend=dict(orientation="v",x=1.02,y=0.5,font=dict(size=9),bgcolor="rgba(0,0,0,0)"),
                annotations=[dict(text="Status",x=0.5,y=0.5,showarrow=False,font=dict(size=10,color=TXT2,family="DM Sans"))])
            st.plotly_chart(fig4,use_container_width=True,config={"displayModeBar":False})

        with ov5:
            st.markdown(ct("Due Status (Attendees)"),unsafe_allow_html=True)
            duc=att['Due Status'].value_counts(); cmap={'Fully Paid':G,'Partially Paid':A,'Fully Due':R,'No Order':GR}
            fig5=go.Figure(go.Pie(labels=duc.index,values=duc.values,hole=0.62,
                marker=dict(colors=[cmap.get(l,B) for l in duc.index],line=dict(color=BG,width=3)),
                textinfo='percent',textfont=dict(size=10,family="DM Sans")))
            fig5.update_layout(**BASE,height=200,legend=dict(orientation="v",x=1.02,y=0.5,font=dict(size=9),bgcolor="rgba(0,0,0,0)"),
                annotations=[dict(text="Due",x=0.5,y=0.5,showarrow=False,font=dict(size=10,color=TXT2,family="DM Sans"))])
            st.plotly_chart(fig5,use_container_width=True,config={"displayModeBar":False})

        with ov6:
            st.markdown(ct("Webinar vs Non-Webinar Leads"),unsafe_allow_html=True)
            cf=leads['converted_from'].value_counts(); cf=cf[cf.index.isin(['Webinar','Non Webinar'])]
            fig6=go.Figure(go.Pie(labels=cf.index,values=cf.values,hole=0.62,
                marker=dict(colors=[G,B],line=dict(color=BG,width=3)),
                textinfo='label+percent',textfont=dict(size=10,family="DM Sans")))
            fig6.update_layout(**BASE,height=200,showlegend=False,
                annotations=[dict(text="Source",x=0.5,y=0.5,showarrow=False,font=dict(size=10,color=TXT2,family="DM Sans"))])
            st.plotly_chart(fig6,use_container_width=True,config={"displayModeBar":False})

        with ov7:
            st.markdown(ct("Attempted vs Unattempted"),unsafe_allow_html=True)
            at=leads['Attempted/Unattempted'].value_counts()
            fig7=go.Figure(go.Pie(labels=at.index,values=at.values,hole=0.62,
                marker=dict(colors=[G,R],line=dict(color=BG,width=3)),
                textinfo='label+percent',textfont=dict(size=10,family="DM Sans")))
            fig7.update_layout(**BASE,height=200,showlegend=False,
                annotations=[dict(text="Attempt",x=0.5,y=0.5,showarrow=False,font=dict(size=10,color=TXT2,family="DM Sans"))])
            st.plotly_chart(fig7,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — SEMINAR ANALYTICS
# ══════════════════════════════════════════════════════════════════════════════
with t_sem:
    if not st.session_state.loaded:
        st.markdown(nodata("🎯","No Data","Upload files first"),unsafe_allow_html=True)
    else:
        att_all=st.session_state.master['att']
        df=apply_sem_filters(att_all,"sm_")
        N=len(df); conv_n=(df['Conv Status']=='Converted').sum()
        conv_r=conv_n/N*100 if N else 0
        rcvd=df['payment_received'].sum(); due=df['total_due'].sum(); rev=df['total_amount'].sum()
        fp=(df['Due Status']=='Fully Paid').sum(); hd=df[df['Due Status'].isin(['Partially Paid','Fully Due'])].shape[0]
        sem_amt=df[df['Conv Status']=='Converted']['Amount Paid'].sum()
        avg_amt=df[df['Conv Status']=='Converted']['Amount Paid'].mean()

        st.markdown("<br>",unsafe_allow_html=True)
        st.markdown(f"""<div class="sg" style="grid-template-columns:repeat(8,minmax(0,1fr))">
{sc("Attended",f"{N:,}",f"of {len(att_all):,}",B,"b")}
{sc("Converted",f"{conv_n:,}",f"↑ {conv_r:.1f}%",G,"g")}
{sc("Not Converted",f"{N-conv_n:,}",f"{100-conv_r:.1f}%",R,"r")}
{sc("Sem Amount Coll.",fmt(sem_amt),f"Avg ₹{avg_amt:,.0f}",G,"g",True)}
{sc("Order Revenue",fmt(rev),"Gross value",P,"n",True)}
{sc("Collected",fmt(rcvd),f"Due: {fmt(due)}",G,"g",True)}
{sc("Fully Paid",f"{fp:,}",f"Has bal: {hd:,}",A,"a")}
{sc("Locations",f"{df['Place'].nunique()}","Venues",T,"n")}
</div>""",unsafe_allow_html=True)
        st.markdown("<br>",unsafe_allow_html=True)

        # Location analysis
        st.markdown('<div class="asec"><div class="asec-t">📍 Location Intelligence</div>',unsafe_allow_html=True)
        loc=df.groupby('Place').agg(Attended=('NAME','count'),Converted=('Conv Status',lambda x:(x=='Converted').sum()),Revenue=('total_amount','sum'),Collected=('payment_received','sum'),Due=('total_due','sum'),SemAmt=('Amount Paid','sum')).reset_index()
        loc['Rate']=(loc['Converted']/loc['Attended']*100).round(1)
        l1,l2=st.columns([3,2],gap="medium")
        with l1:
            st.markdown(ct("Attended vs Converted — by Location"),unsafe_allow_html=True)
            loc_s=loc.sort_values('Attended',ascending=False)
            fig=go.Figure()
            fig.add_trace(go.Bar(name='Attended',x=loc_s['Place'],y=loc_s['Attended'],marker_color=B,marker_line_width=0,opacity=0.5,hovertemplate='<b>%{x}</b><br>Attended: %{y}<extra></extra>'))
            fig.add_trace(go.Bar(name='Converted',x=loc_s['Place'],y=loc_s['Converted'],marker_color=G,marker_line_width=0,hovertemplate='<b>%{x}</b><br>Converted: %{y}<extra></extra>'))
            fig.update_layout(**BASE,height=240,barmode='group',bargap=0.22,bargroupgap=0.05,
                legend=dict(orientation="h",x=0,y=1.12,font=dict(size=10),bgcolor="rgba(0,0,0,0)"),
                xaxis=dict(showgrid=False,tickangle=-30,tickfont=dict(size=9,color=TXT2)),
                yaxis=dict(gridcolor=GRID,tickfont=dict(size=9,color=TXT)))
            st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})
        with l2:
            st.markdown(ct("Conversion Rate % — Color coded"),unsafe_allow_html=True)
            lrs=loc.sort_values('Rate'); bc=[R if r<10 else(A if r<20 else G) for r in lrs['Rate']]
            fig2=go.Figure(go.Bar(x=lrs['Rate'],y=lrs['Place'],orientation='h',marker_color=bc,marker_line_width=0,text=[f"{r}%" for r in lrs['Rate']],textposition='outside',textfont=dict(size=9,color=TXT2),hovertemplate='<b>%{y}</b><br>%{x:.1f}%<extra></extra>'))
            if len(lrs): fig2.add_vline(x=lrs['Rate'].mean(),line_dash="dot",line_color=A,line_width=1,annotation_text=f"avg {lrs['Rate'].mean():.1f}%",annotation_font=dict(size=8,color=A),annotation_position="top right")
            fig2.update_layout(**BASE,height=240,bargap=0.28,xaxis=dict(showgrid=False,showticklabels=False,zeroline=False,range=[0,lrs['Rate'].max()*1.3 if len(lrs) else 10]),yaxis=dict(showgrid=False,tickfont=dict(size=9,color=TXT2)))
            st.plotly_chart(fig2,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

        # Date + Session + Trainer
        st.markdown('<div class="asec"><div class="asec-t">📅 Date · Session · Trainer Analysis</div>',unsafe_allow_html=True)
        d1,d2,d3=st.columns([4,2,3],gap="medium")
        with d1:
            st.markdown(ct("Attendance & Conversion Rate — Seminar Date Trend"),unsafe_allow_html=True)
            dg=df.groupby('Seminar Date').agg(Att=('NAME','count'),Conv=('Conv Status',lambda x:(x=='Converted').sum())).reset_index().dropna(subset=['Seminar Date']).sort_values('Seminar Date')
            dg['Label']=dg['Seminar Date'].dt.strftime("%d %b"); dg['CR']=(dg['Conv']/dg['Att']*100).round(1)
            fig3=make_subplots(specs=[[{"secondary_y":True}]])
            fig3.add_trace(go.Bar(name='Attended',x=dg['Label'],y=dg['Att'],marker_color=B,marker_line_width=0,opacity=0.55,hovertemplate='%{x}<br>Attended: %{y}<extra></extra>'),secondary_y=False)
            fig3.add_trace(go.Scatter(name='Conv Rate %',x=dg['Label'],y=dg['CR'],line=dict(color=G,width=2.5,shape='spline'),mode='lines+markers',marker=dict(size=6,color=G,line=dict(color=BG,width=1.5)),hovertemplate='%{x}<br>%{y:.1f}%<extra></extra>'),secondary_y=True)
            fig3.update_layout(**BASE,height=230,bargap=0.3,legend=dict(orientation="h",x=0,y=1.12,font=dict(size=10),bgcolor="rgba(0,0,0,0)"),xaxis=dict(showgrid=False,tickfont=dict(size=9,color=TXT2)),yaxis=dict(gridcolor=GRID,tickfont=dict(size=9,color=TXT),zeroline=False),yaxis2=dict(showgrid=False,tickfont=dict(size=9,color=G),ticksuffix="%"))
            st.plotly_chart(fig3,use_container_width=True,config={"displayModeBar":False})
        with d2:
            st.markdown(ct("Morning vs Evening — Stacked"),unsafe_allow_html=True)
            ses=df.groupby('Session').agg(Att=('NAME','count'),Conv=('Conv Status',lambda x:(x=='Converted').sum())).reset_index()
            ses=ses[ses['Session'].isin(['MORNING','EVENING'])].copy(); ses['NC']=ses['Att']-ses['Conv']
            fig4=go.Figure()
            fig4.add_trace(go.Bar(name='Converted',x=ses['Session'],y=ses['Conv'],marker_color=G,marker_line_width=0))
            fig4.add_trace(go.Bar(name='Not Conv.',x=ses['Session'],y=ses['NC'],marker_color=B,opacity=0.3,marker_line_width=0))
            fig4.update_layout(**BASE,height=230,barmode='stack',bargap=0.4,legend=dict(orientation="h",x=0,y=1.12,font=dict(size=9),bgcolor="rgba(0,0,0,0)"),xaxis=dict(showgrid=False,tickfont=dict(size=10,color=TXT2)),yaxis=dict(gridcolor=GRID,tickfont=dict(size=9,color=TXT)))
            st.plotly_chart(fig4,use_container_width=True,config={"displayModeBar":False})
        with d3:
            st.markdown(ct("Trainer — Attended vs Converted"),unsafe_allow_html=True)
            tr=df.groupby('Trainer Norm').agg(Att=('NAME','count'),Conv=('Conv Status',lambda x:(x=='Converted').sum())).reset_index()
            tr['Short']=tr['Trainer Norm'].apply(lambda x:x[:24]+'…' if len(x)>24 else x); tr=tr.sort_values('Att')
            fig5=go.Figure()
            fig5.add_trace(go.Bar(name='Attended',y=tr['Short'],x=tr['Att'],orientation='h',marker_color=B,opacity=0.35,marker_line_width=0))
            fig5.add_trace(go.Bar(name='Converted',y=tr['Short'],x=tr['Conv'],orientation='h',marker_color=G,marker_line_width=0))
            fig5.update_layout(**BASE,height=230,barmode='overlay',bargap=0.28,legend=dict(orientation="h",x=0,y=1.12,font=dict(size=9),bgcolor="rgba(0,0,0,0)"),xaxis=dict(showgrid=False,showticklabels=False,zeroline=False),yaxis=dict(showgrid=False,tickfont=dict(size=8,color=TXT2)))
            st.plotly_chart(fig5,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

        # Payment + Profile + Revenue
        st.markdown('<div class="asec"><div class="asec-t">💳 Payment · Profile · Revenue</div>',unsafe_allow_html=True)
        p1,p2,p3,p4=st.columns(4,gap="medium")
        with p1:
            st.markdown(ct("Seminar Payment Mode"),unsafe_allow_html=True)
            conv_att=df[df['Conv Status']=='Converted']
            pm_sem=conv_att['Mode of Payment'].str.upper().str.strip()
            pm_sem=pm_sem.replace({'QR CODE':'QR','QR ':'QR','CASH ':'CASH','LINK':'Link'}).value_counts()
            pm_sem=pm_sem[pm_sem.index!='NAN']
            fig6=go.Figure(go.Pie(labels=pm_sem.index,values=pm_sem.values,hole=0.62,marker=dict(colors=[G,B,A,P,T,R],line=dict(color=BG,width=3)),textinfo='label+percent',textfont=dict(size=9,family="DM Sans")))
            fig6.update_layout(**BASE,height=210,showlegend=False,annotations=[dict(text="Mode",x=0.5,y=0.5,showarrow=False,font=dict(size=10,color=TXT2,family="DM Sans"))])
            st.plotly_chart(fig6,use_container_width=True,config={"displayModeBar":False})
        with p2:
            st.markdown(ct("Due Status"),unsafe_allow_html=True)
            duc=df['Due Status'].value_counts(); cmap={'Fully Paid':G,'Partially Paid':A,'Fully Due':R,'No Order':GR}
            fig7=go.Figure(go.Pie(labels=duc.index,values=duc.values,hole=0.62,marker=dict(colors=[cmap.get(l,B) for l in duc.index],line=dict(color=BG,width=3)),textinfo='percent',textfont=dict(size=10,family="DM Sans")))
            fig7.update_layout(**BASE,height=210,legend=dict(orientation="v",x=1.02,y=0.5,font=dict(size=9),bgcolor="rgba(0,0,0,0)"),annotations=[dict(text="Due",x=0.5,y=0.5,showarrow=False,font=dict(size=10,color=TXT2,family="DM Sans"))])
            st.plotly_chart(fig7,use_container_width=True,config={"displayModeBar":False})
        with p3:
            st.markdown(ct("Trader vs Non-Trader"),unsafe_allow_html=True)
            trd=df['TRADER'].map(lambda x:'Trader' if x in ['YES','TYES'] else 'Non-Trader').value_counts()
            fig8=go.Figure(go.Pie(labels=trd.index,values=trd.values,hole=0.62,marker=dict(colors=[T,P],line=dict(color=BG,width=3)),textinfo='label+percent',textfont=dict(size=10,family="DM Sans")))
            fig8.update_layout(**BASE,height=210,showlegend=False,annotations=[dict(text="Trader",x=0.5,y=0.5,showarrow=False,font=dict(size=10,color=TXT2,family="DM Sans"))])
            st.plotly_chart(fig8,use_container_width=True,config={"displayModeBar":False})
        with p4:
            st.markdown(ct("Existing Student vs New"),unsafe_allow_html=True)
            stu=df['Is our Student ?'].map(lambda x:'Existing' if x in ['YES','STUDENT'] else 'New Lead').value_counts()
            fig9=go.Figure(go.Pie(labels=stu.index,values=stu.values,hole=0.62,marker=dict(colors=[A,B],line=dict(color=BG,width=3)),textinfo='label+percent',textfont=dict(size=10,family="DM Sans")))
            fig9.update_layout(**BASE,height=210,showlegend=False,annotations=[dict(text="Student",x=0.5,y=0.5,showarrow=False,font=dict(size=10,color=TXT2,family="DM Sans"))])
            st.plotly_chart(fig9,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

        # Revenue by location
        st.markdown('<div class="asec"><div class="asec-t">💰 Revenue — Collected vs Due by Location</div>',unsafe_allow_html=True)
        rev_l=df.groupby('Place').agg(Collected=('payment_received','sum'),Due=('total_due','sum')).reset_index().sort_values('Collected',ascending=False)
        fig10=go.Figure()
        fig10.add_trace(go.Bar(name='Collected',x=rev_l['Place'],y=rev_l['Collected'],marker_color=G,marker_line_width=0,opacity=0.85,hovertemplate='<b>%{x}</b><br>Collected: ₹%{y:,.0f}<extra></extra>'))
        fig10.add_trace(go.Bar(name='Due',x=rev_l['Place'],y=rev_l['Due'],marker_color=R,marker_line_width=0,opacity=0.85,hovertemplate='<b>%{x}</b><br>Due: ₹%{y:,.0f}<extra></extra>'))
        fig10.update_layout(**BASE,height=220,barmode='group',bargap=0.22,bargroupgap=0.05,
            legend=dict(orientation="h",x=0,y=1.12,font=dict(size=10),bgcolor="rgba(0,0,0,0)"),
            xaxis=dict(showgrid=False,tickangle=-30,tickfont=dict(size=9,color=TXT2)),
            yaxis=dict(gridcolor=GRID,tickformat=",.0f",tickprefix="₹",tickfont=dict(size=9,color=TXT)))
        st.plotly_chart(fig10,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — LEADS ANALYTICS
# ══════════════════════════════════════════════════════════════════════════════
with t_leads:
    if not st.session_state.loaded:
        st.markdown(nodata("📣","No Data","Upload files first"),unsafe_allow_html=True)
    else:
        leads_all=st.session_state.master['leads']

        with st.expander("🔽  Leads Filters",expanded=True):
            lf1,lf2,lf3,lf4=st.columns(4); lf5,lf6,lf7,lf8=st.columns(4); lf9,lf10,_,_=st.columns(4)
            sel_cf  =lf1.multiselect("✅ Converted From",    ["Webinar","Non Webinar"],                                                    key="lf_cf")
            sel_ls  =lf2.multiselect("📢 Lead Source",       sorted([x for x in leads_all['leadsource'].unique() if x!='Unknown']),        key="lf_ls")
            sel_cp  =lf3.multiselect("📣 Campaign",          sorted([x for x in leads_all['campaign_name'].unique() if x!='Unknown']),     key="lf_cp")
            sel_lst =lf4.multiselect("📊 Lead Status",       sorted([x for x in leads_all['leadstatus'].unique() if x!='Unknown']),        key="lf_lst")
            sel_stg =lf5.multiselect("🏷 Stage",             sorted([x for x in leads_all['stage_name'].unique() if x!='Unknown']),        key="lf_stg")
            sel_lo  =lf6.multiselect("👤 Lead Owner",        sorted([x for x in leads_all['leadownername'].unique() if x!='Unknown']),     key="lf_lo")
            sel_st  =lf7.multiselect("📍 State",             sorted([x for x in leads_all['state'].unique() if x!='Unknown']),             key="lf_st")
            sel_att =lf8.selectbox("📞 Attempted?",          ["All","Attempted","Unattempted"],                                            key="lf_att")
            min_d, max_d = safe_date_bounds(leads_all['leaddate'])
            d_from=lf9.date_input("📅 Date From",value=min_d,min_value=min_d,max_value=max_d,key="lf_df")
            d_to  =lf10.date_input("📅 Date To",  value=max_d,min_value=min_d,max_value=max_d,key="lf_dt")

        ld=leads_all.copy()
        if sel_cf:  ld=ld[ld['converted_from'].isin(sel_cf)]
        if sel_ls:  ld=ld[ld['leadsource'].isin(sel_ls)]
        if sel_cp:  ld=ld[ld['campaign_name'].isin(sel_cp)]
        if sel_lst: ld=ld[ld['leadstatus'].isin(sel_lst)]
        if sel_stg: ld=ld[ld['stage_name'].isin(sel_stg)]
        if sel_lo:  ld=ld[ld['leadownername'].isin(sel_lo)]
        if sel_st:  ld=ld[ld['state'].isin(sel_st)]
        if sel_att!="All":
            ld=ld[ld['Attempted/Unattempted'].astype(str).str.strip().str.lower()==sel_att.lower()]
        ld=ld[(ld['leaddate'].dt.date>=d_from)&(ld['leaddate'].dt.date<=d_to)]

        NL=len(ld); cl=ld['Is Converted'].sum(); cl_r=cl/NL*100 if NL else 0
        web=( ld['converted_from']=='Webinar').sum(); nweb=(ld['converted_from']=='Non Webinar').sum()
        seat=(ld['leadstatus']=='Seat Booked').sum(); att_n=(ld['Attempted/Unattempted']=='Attempted').sum()

        st.markdown("<br>",unsafe_allow_html=True)
        st.markdown(f"""<div class="sg" style="grid-template-columns:repeat(8,minmax(0,1fr))">
{sc("Total Leads",f"{NL:,}",f"of {len(leads_all):,}",B,"b")}
{sc("Converted",f"{cl:,}",f"↑ {cl_r:.1f}%",G,"g")}
{sc("Not Converted",f"{NL-cl:,}",f"{100-cl_r:.1f}%",R,"r")}
{sc("Webinar Leads",f"{web:,}",f"{web/NL*100:.0f}%" if NL else "",T,"b")}
{sc("Non-Webinar",f"{nweb:,}",f"{nweb/NL*100:.0f}%" if NL else "",P,"n")}
{sc("Seat Booked",f"{seat:,}","In pipeline",A,"a")}
{sc("Attempted",f"{att_n:,}",f"{att_n/NL*100:.0f}% contact rate" if NL else "",G,"g")}
{sc("Campaigns",f"{ld['campaign_name'].nunique()}","Active",B,"n")}
</div>""",unsafe_allow_html=True)
        st.markdown("<br>",unsafe_allow_html=True)

        # Webinar vs Non-Webinar deep dive
        st.markdown('<div class="asec"><div class="asec-t">✅ Webinar vs Non-Webinar Deep Analysis</div>',unsafe_allow_html=True)
        w1,w2,w3=st.columns([2,2,4],gap="medium")
        with w1:
            st.markdown(ct("Source Mix"),unsafe_allow_html=True)
            cf_v=ld['converted_from'].value_counts(); cf_v=cf_v[cf_v.index.isin(['Webinar','Non Webinar'])]
            fig=go.Figure(go.Pie(labels=cf_v.index,values=cf_v.values,hole=0.62,marker=dict(colors=[G,B],line=dict(color=BG,width=3)),textinfo='label+percent',textfont=dict(size=11,family="DM Sans")))
            fig.update_layout(**BASE,height=210,showlegend=False,annotations=[dict(text="Source",x=0.5,y=0.5,showarrow=False,font=dict(size=10,color=TXT2,family="DM Sans"))])
            st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})
        with w2:
            st.markdown(ct("Conversion Rate Comparison"),unsafe_allow_html=True)
            wc=ld[ld['converted_from'].isin(['Webinar','Non Webinar'])].groupby('converted_from').agg(Total=('_id','count'),Conv=('Is Converted','sum')).reset_index()
            wc['Rate']=(wc['Conv']/wc['Total']*100).round(1)
            fig2=go.Figure(go.Bar(x=wc['Rate'],y=wc['converted_from'],orientation='h',marker_color=[G if i==wc['Rate'].idxmax() else B for i in wc.index],marker_line_width=0,text=[f"{r}%" for r in wc['Rate']],textposition='outside',textfont=dict(size=13,color=TXT2)))
            fig2.update_layout(**BASE,height=210,bargap=0.45,xaxis=dict(showgrid=False,showticklabels=False,zeroline=False,range=[0,wc['Rate'].max()*1.35] if len(wc) else [0,10]),yaxis=dict(showgrid=False,tickfont=dict(size=11,color=TXT2)))
            st.plotly_chart(fig2,use_container_width=True,config={"displayModeBar":False})
        with w3:
            st.markdown(ct("Lead Status — Webinar vs Non-Webinar (Top 6 statuses)"),unsafe_allow_html=True)
            wls=ld[ld['converted_from'].isin(['Webinar','Non Webinar'])].groupby(['converted_from','leadstatus']).size().reset_index(name='Count')
            top6=ld['leadstatus'].value_counts().head(6).index.tolist()
            wls=wls[wls['leadstatus'].isin(top6)]
            fig3=go.Figure()
            for src,col in [('Webinar',G),('Non Webinar',B)]:
                sub=wls[wls['converted_from']==src]
                fig3.add_trace(go.Bar(name=src,x=sub['leadstatus'],y=sub['Count'],marker_color=col,opacity=0.85,marker_line_width=0,hovertemplate=f'{src}<br>%{{x}}: %{{y}}<extra></extra>'))
            fig3.update_layout(**BASE,height=210,barmode='group',bargap=0.2,bargroupgap=0.05,legend=dict(orientation="h",x=0,y=1.12,font=dict(size=10),bgcolor="rgba(0,0,0,0)"),xaxis=dict(showgrid=False,tickangle=-20,tickfont=dict(size=9,color=TXT2)),yaxis=dict(gridcolor=GRID,tickfont=dict(size=9,color=TXT)))
            st.plotly_chart(fig3,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

        # Lead Source + Campaign
        st.markdown('<div class="asec"><div class="asec-t">📢 Lead Source & Campaign Performance</div>',unsafe_allow_html=True)
        ls1,ls2=st.columns(2,gap="medium")
        with ls1:
            st.markdown(ct("Lead Source — Volume & Conversion Rate"),unsafe_allow_html=True)
            lsrc=ld.groupby('leadsource').agg(Total=('_id','count'),Conv=('Is Converted','sum')).reset_index()
            lsrc['Rate']=(lsrc['Conv']/lsrc['Total']*100).round(1); lsrc=lsrc[lsrc['leadsource']!='Unknown'].sort_values('Total',ascending=False).head(12)
            fig4=make_subplots(specs=[[{"secondary_y":True}]])
            fig4.add_trace(go.Bar(name='Total Leads',x=lsrc['leadsource'],y=lsrc['Total'],marker_color=B,marker_line_width=0,opacity=0.6),secondary_y=False)
            fig4.add_trace(go.Scatter(name='Conv Rate %',x=lsrc['leadsource'],y=lsrc['Rate'],mode='lines+markers',line=dict(color=G,width=2.5),marker=dict(size=7,color=G,line=dict(color=BG,width=1.5))),secondary_y=True)
            fig4.update_layout(**BASE,height=260,bargap=0.3,legend=dict(orientation="h",x=0,y=1.12,font=dict(size=10),bgcolor="rgba(0,0,0,0)"),xaxis=dict(showgrid=False,tickangle=-35,tickfont=dict(size=9,color=TXT2)),yaxis=dict(gridcolor=GRID,tickfont=dict(size=9,color=TXT)),yaxis2=dict(showgrid=False,tickfont=dict(size=9,color=G),ticksuffix="%"))
            st.plotly_chart(fig4,use_container_width=True,config={"displayModeBar":False})
        with ls2:
            st.markdown(ct("Top Campaigns — Volume (color = conversions)"),unsafe_allow_html=True)
            cmp=ld.groupby('campaign_name').agg(Total=('_id','count'),Conv=('Is Converted','sum')).reset_index()
            cmp=cmp[cmp['campaign_name']!='Unknown'].sort_values('Total',ascending=False).head(12)
            cmp['Short']=cmp['campaign_name'].apply(lambda x:x[:30]+'…' if len(x)>30 else x)
            fig5=go.Figure(go.Bar(x=cmp['Total'],y=cmp['Short'],orientation='h',marker=dict(color=cmp['Conv'],colorscale=[[0,P],[1,G]],line=dict(width=0)),text=cmp['Total'],textposition='outside',textfont=dict(size=9,color=TXT2),customdata=cmp['Conv'],hovertemplate='<b>%{y}</b><br>Leads: %{x}<br>Converted: %{customdata}<extra></extra>'))
            fig5.update_layout(**BASE,height=260,bargap=0.25,xaxis=dict(showgrid=False,showticklabels=False,zeroline=False,range=[0,cmp['Total'].max()*1.25] if len(cmp) else [0,10]),yaxis=dict(showgrid=False,tickfont=dict(size=8.5,color=TXT2)))
            st.plotly_chart(fig5,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

        # Status + Stage + Owner + Attempted
        st.markdown('<div class="asec"><div class="asec-t">📊 Status · Stage · Owner · Attempt Analysis</div>',unsafe_allow_html=True)
        o1,o2,o3,o4=st.columns(4,gap="medium")
        with o1:
            st.markdown(ct("Lead Status Top 10"),unsafe_allow_html=True)
            lst_v=ld['leadstatus'].value_counts().head(10)
            fig6=go.Figure(go.Bar(x=lst_v.values,y=lst_v.index,orientation='h',marker=dict(color=lst_v.values,colorscale=[[0,B],[0.5,P],[1,G]],line=dict(width=0)),text=lst_v.values,textposition='outside',textfont=dict(size=9,color=TXT2)))
            fig6.update_layout(**BASE,height=280,bargap=0.28,xaxis=dict(showgrid=False,showticklabels=False,zeroline=False,range=[0,lst_v.max()*1.25] if len(lst_v) else [0,10]),yaxis=dict(showgrid=False,tickfont=dict(size=8.5,color=TXT2)))
            st.plotly_chart(fig6,use_container_width=True,config={"displayModeBar":False})
        with o2:
            st.markdown(ct("Stage Pipeline"),unsafe_allow_html=True)
            stg_v=ld['stage_name'].value_counts(); stg_v=stg_v[stg_v.index!='Unknown']
            stg_c={'Sales Closed':G,'Registered':T,'Pipeline':B,'Prospect':A,'High Priority':A,'Not Converted':R,'Poor Lead':R,'Not Connected':GR,'Partially Converted':P}
            fig7=go.Figure(go.Pie(labels=stg_v.index,values=stg_v.values,hole=0.58,marker=dict(colors=[stg_c.get(l,B) for l in stg_v.index],line=dict(color=BG,width=2)),textinfo='percent',textfont=dict(size=9,family="DM Sans")))
            fig7.update_layout(**BASE,height=280,legend=dict(orientation="v",x=1.02,y=0.5,font=dict(size=8.5),bgcolor="rgba(0,0,0,0)"),annotations=[dict(text="Stage",x=0.5,y=0.5,showarrow=False,font=dict(size=10,color=TXT2,family="DM Sans"))])
            st.plotly_chart(fig7,use_container_width=True,config={"displayModeBar":False})
        with o3:
            st.markdown(ct("Lead Owner — Volume & Conversions"),unsafe_allow_html=True)
            lo=ld[ld['leadownername']!='Unknown'].groupby('leadownername').agg(Total=('_id','count'),Conv=('Is Converted','sum')).reset_index()
            lo=lo.sort_values('Total',ascending=True).tail(10)
            fig8=go.Figure()
            fig8.add_trace(go.Bar(name='Total',y=lo['leadownername'],x=lo['Total'],orientation='h',marker_color=B,opacity=0.35,marker_line_width=0))
            fig8.add_trace(go.Bar(name='Converted',y=lo['leadownername'],x=lo['Conv'],orientation='h',marker_color=G,marker_line_width=0))
            fig8.update_layout(**BASE,height=280,barmode='overlay',bargap=0.28,legend=dict(orientation="h",x=0,y=1.12,font=dict(size=9),bgcolor="rgba(0,0,0,0)"),xaxis=dict(showgrid=False,showticklabels=False,zeroline=False),yaxis=dict(showgrid=False,tickfont=dict(size=8.5,color=TXT2)))
            st.plotly_chart(fig8,use_container_width=True,config={"displayModeBar":False})
        with o4:
            st.markdown(ct("Attempt Rate & Conv Rate"),unsafe_allow_html=True)
            atc=ld.groupby('Attempted/Unattempted').agg(Total=('_id','count'),Conv=('Is Converted','sum')).reset_index()
            atc['Rate']=(atc['Conv']/atc['Total']*100).round(1)
            fig9a=go.Figure(go.Pie(labels=atc['Attempted/Unattempted'],values=atc['Total'],hole=0.55,marker=dict(colors=[G,R],line=dict(color=BG,width=3)),textinfo='label+percent',textfont=dict(size=10,family="DM Sans")))
            fig9a.update_layout(paper_bgcolor=BG, plot_bgcolor=BG, font=dict(family="DM Sans", color=TXT, size=11), height=150, showlegend=False, margin=dict(l=0,r=0,t=4,b=4))
            st.plotly_chart(fig9a,use_container_width=True,config={"displayModeBar":False})
            fig9b=go.Figure(go.Bar(x=atc['Rate'],y=atc['Attempted/Unattempted'],orientation='h',marker_color=[G,R],marker_line_width=0,text=[f"{r}%" for r in atc['Rate']],textposition='outside',textfont=dict(size=11,color=TXT2)))
            fig9b.update_layout(paper_bgcolor=BG, plot_bgcolor=BG, font=dict(family="DM Sans", color=TXT, size=11), height=115, bargap=0.45, margin=dict(l=0,r=40,t=4,b=0), xaxis=dict(showgrid=False,showticklabels=False,zeroline=False,range=[0,atc['Rate'].max()*1.4] if len(atc) else [0,10]), yaxis=dict(showgrid=False,tickfont=dict(size=10,color=TXT2)))
            st.plotly_chart(fig9b,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — CONVERSION ANALYTICS
# ══════════════════════════════════════════════════════════════════════════════
with t_conv:
    if not st.session_state.loaded:
        st.markdown(nodata("💰","No Data","Upload files first"),unsafe_allow_html=True)
    else:
        conv_all=st.session_state.master['conv']

        with st.expander("🔽  Conversion Filters",expanded=True):
            cf1,cf2,cf3,cf4=st.columns(4)
            sel_cst =cf1.multiselect("📌 Status",        sorted(conv_all['status'].dropna().unique()),            key="cv_st")
            sel_csvc=cf2.multiselect("📚 Service",        sorted(conv_all['service_name'].dropna().unique()),      key="cv_sv")
            sel_crep=cf3.multiselect("👤 Sales Rep",      sorted(conv_all['sales_rep_name'].dropna().unique()),    key="cv_re")
            sel_cpm =cf4.multiselect("💳 Payment Mode",   sorted(conv_all['PM Label'].dropna().unique()),          key="cv_pm")
            cf5,cf6,_,_=st.columns(4)
            sel_cmth=cf5.multiselect("📅 Month",          sorted(conv_all['month'].dropna().unique()),             key="cv_mt")
            sel_ctr =cf6.multiselect("🎤 Trainer",        sorted(conv_all['trainer_name'].dropna().unique()),      key="cv_tr")

        cv=conv_all.copy()
        if sel_cst:  cv=cv[cv['status'].isin(sel_cst)]
        if sel_csvc: cv=cv[cv['service_name'].isin(sel_csvc)]
        if sel_crep: cv=cv[cv['sales_rep_name'].isin(sel_crep)]
        if sel_cpm:  cv=cv[cv['PM Label'].isin(sel_cpm)]
        if sel_cmth: cv=cv[cv['month'].isin(sel_cmth)]
        if sel_ctr:  cv=cv[cv['trainer_name'].isin(sel_ctr)]

        NCV=len(cv); rev_cv=cv['total_amount'].sum(); rcvd_cv=cv['payment_received'].sum()
        due_cv=cv['total_due'].sum(); active_cv=(cv['status']=='Active').sum()
        refunded_cv=cv['is_refunded'].sum(); sc_cv=cv['is_shortClosed'].sum()

        st.markdown("<br>",unsafe_allow_html=True)
        st.markdown(f"""<div class="sg" style="grid-template-columns:repeat(8,minmax(0,1fr))">
{sc("Total Orders",f"{NCV:,}",f"of {len(conv_all):,}",B,"b")}
{sc("Active",f"{active_cv:,}",f"{active_cv/NCV*100:.0f}%" if NCV else "",G,"g")}
{sc("Inactive",f"{(cv['status']=='Inactive').sum():,}","",R,"r")}
{sc("Gross Revenue",fmt(rev_cv),"Order value",P,"n",True)}
{sc("Collected",fmt(rcvd_cv),f"Due: {fmt(due_cv)}",G,"g",True)}
{sc("Avg Order Value",fmt(cv[cv['total_amount']>0]['total_amount'].mean()),"Non-zero",T,"n",True)}
{sc("Refunded",f"{int(refunded_cv)}","",R,"r")}
{sc("Short Closed",f"{int(sc_cv)}","",A,"a")}
</div>""",unsafe_allow_html=True)
        st.markdown("<br>",unsafe_allow_html=True)

        st.markdown('<div class="asec"><div class="asec-t">📈 Revenue Trend & Order Analysis</div>',unsafe_allow_html=True)
        cv1,cv2=st.columns([3,2],gap="medium")
        with cv1:
            st.markdown(ct("Monthly Orders & Revenue"),unsafe_allow_html=True)
            mo=cv.groupby('month').agg(n=('_id','count'),rev=('total_amount','sum'),coll=('payment_received','sum')).reset_index().tail(6)
            fig=make_subplots(specs=[[{"secondary_y":True}]])
            fig.add_trace(go.Bar(name='Orders',x=mo['month'],y=mo['n'],marker_color=B,marker_line_width=0,opacity=0.6),secondary_y=False)
            fig.add_trace(go.Scatter(name='Revenue',x=mo['month'],y=mo['rev'],line=dict(color=G,width=2.5,shape='spline'),mode='lines+markers',marker=dict(size=6,color=G,line=dict(color=BG,width=1.5)),fill='tozeroy',fillcolor='rgba(79,206,143,0.07)'),secondary_y=True)
            fig.add_trace(go.Scatter(name='Collected',x=mo['month'],y=mo['coll'],line=dict(color=T,width=2,dash='dot'),mode='lines+markers',marker=dict(size=5,color=T)),secondary_y=True)
            fig.update_layout(**BASE,height=250,bargap=0.3,legend=dict(orientation="h",x=0,y=1.12,font=dict(size=10),bgcolor="rgba(0,0,0,0)"),xaxis=dict(showgrid=False,tickfont=dict(size=9,color=TXT2)),yaxis=dict(gridcolor=GRID,tickfont=dict(size=9,color=TXT),zeroline=False),yaxis2=dict(showgrid=False,tickfont=dict(size=9,color=G),tickformat=",.0f"))
            st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})
        with cv2:
            st.markdown(ct("Order Status Breakdown"),unsafe_allow_html=True)
            ost=cv['status'].value_counts()
            fig2=go.Figure(go.Pie(labels=ost.index,values=ost.values,hole=0.62,marker=dict(colors=[G,R,A,GR],line=dict(color=BG,width=3)),textinfo='label+percent',textfont=dict(size=10,family="DM Sans")))
            fig2.update_layout(**BASE,height=250,showlegend=False,annotations=[dict(text="Status",x=0.5,y=0.5,showarrow=False,font=dict(size=10,color=TXT2,family="DM Sans"))])
            st.plotly_chart(fig2,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

        st.markdown('<div class="asec"><div class="asec-t">🏆 Course · Sales Rep · Trainer · Payment Mode</div>',unsafe_allow_html=True)
        cv3,cv4,cv5,cv6=st.columns(4,gap="medium")
        with cv3:
            st.markdown(ct("Top Courses by Orders"),unsafe_allow_html=True)
            svc=cv.groupby('service_name').agg(Count=('_id','count'),Rev=('total_amount','sum')).reset_index().sort_values('Count').tail(8)
            svc['Short']=svc['service_name'].apply(lambda x:x[:26]+'…' if len(x)>26 else x)
            fig3=go.Figure(go.Bar(x=svc['Count'],y=svc['Short'],orientation='h',marker=dict(color=svc['Count'],colorscale=[[0,P],[1,T]],line=dict(width=0)),text=svc['Count'],textposition='outside',textfont=dict(size=9,color=TXT2),customdata=svc['Rev'],hovertemplate='<b>%{y}</b><br>Orders: %{x}<br>Rev: ₹%{customdata:,.0f}<extra></extra>'))
            fig3.update_layout(**BASE,height=260,bargap=0.3,xaxis=dict(showgrid=False,showticklabels=False,zeroline=False,range=[0,svc['Count'].max()*1.25] if len(svc) else [0,10]),yaxis=dict(showgrid=False,tickfont=dict(size=8.5,color=TXT2)))
            st.plotly_chart(fig3,use_container_width=True,config={"displayModeBar":False})
        with cv4:
            st.markdown(ct("Sales Rep — Conv vs Revenue (bubble=avg deal)"),unsafe_allow_html=True)
            rp=cv.groupby('sales_rep_name').agg(Conv=('_id','count'),Rev=('total_amount','sum')).reset_index()
            rp['Avg']=(rp['Rev']/rp['Conv']).fillna(0); rp=rp[rp['Conv']>0]
            bubble_size = rp['Avg'].clip(lower=5000).div(1500).clip(upper=60)
            fig4=go.Figure(go.Scatter(x=rp['Conv'],y=rp['Rev'],mode='markers+text',marker=dict(size=bubble_size,color=rp['Conv'],colorscale=[[0,B],[.5,P],[1,G]],line=dict(color=BG,width=1.5),sizemin=6,sizemode='diameter'),text=rp['sales_rep_name'].apply(lambda x:str(x).split()[0] if str(x).split() else 'Rep'),textfont=dict(size=8.5,color=TXT2),textposition='top center',customdata=rp['Avg'],hovertemplate='<b>%{text}</b><br>Orders:%{x}<br>Rev:₹%{y:,.0f}<br>Avg:₹%{customdata:,.0f}<extra></extra>'))
            if len(rp):
                fig4.add_hline(y=rp['Rev'].mean(),line_dash="dot",line_color=A,line_width=1,annotation_text="avg rev",annotation_font=dict(size=8,color=A))
                fig4.add_vline(x=rp['Conv'].mean(),line_dash="dot",line_color=A,line_width=1,annotation_text="avg",annotation_font=dict(size=8,color=A))
            fig4.update_layout(**BASE,height=260,xaxis=dict(gridcolor=GRID,title=dict(text="Orders",font=dict(size=10,color=TXT)),tickfont=dict(size=9,color=TXT),zeroline=False),yaxis=dict(gridcolor=GRID,title=dict(text="Revenue",font=dict(size=10,color=TXT)),tickformat=",.0f",tickfont=dict(size=9,color=TXT),zeroline=False))
            st.plotly_chart(fig4,use_container_width=True,config={"displayModeBar":False})
        with cv5:
            st.markdown(ct("Top Trainers by Orders"),unsafe_allow_html=True)
            tn=cv[cv['trainer_name']!='nan'].groupby('trainer_name').agg(Count=('_id','count'),Rev=('total_amount','sum')).reset_index().sort_values('Count').tail(10)
            fig5=go.Figure(go.Bar(x=tn['Count'],y=tn['trainer_name'],orientation='h',marker=dict(color=tn['Count'],colorscale=[[0,B],[1,G]],line=dict(width=0)),text=tn['Count'],textposition='outside',textfont=dict(size=9,color=TXT2)))
            fig5.update_layout(**BASE,height=260,bargap=0.28,xaxis=dict(showgrid=False,showticklabels=False,zeroline=False,range=[0,tn['Count'].max()*1.25] if len(tn) else [0,10]),yaxis=dict(showgrid=False,tickfont=dict(size=8.5,color=TXT2)))
            st.plotly_chart(fig5,use_container_width=True,config={"displayModeBar":False})
        with cv6:
            st.markdown(ct("Payment Mode Distribution"),unsafe_allow_html=True)
            pm_cv=cv['PM Label'].value_counts()
            pm_colors = ([G,B,P,A,R,T,GR] * ((len(pm_cv) // 7) + 1))[:len(pm_cv)]
            fig6=go.Figure(go.Bar(x=pm_cv.values,y=pm_cv.index,orientation='h',marker_color=pm_colors,marker_line_width=0,text=pm_cv.values,textposition='outside',textfont=dict(size=9,color=TXT2),hovertemplate='<b>%{y}</b><br>%{x}<extra></extra>'))
            fig6.update_layout(**BASE,height=260,bargap=0.3,xaxis=dict(showgrid=False,showticklabels=False,zeroline=False,range=[0,pm_cv.max()*1.25] if len(pm_cv) else [0,10]),yaxis=dict(showgrid=False,tickfont=dict(size=9,color=TXT2)))
            st.plotly_chart(fig6,use_container_width=True,config={"displayModeBar":False})
        st.markdown('</div>',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — RECORDS & EXPORT
# ══════════════════════════════════════════════════════════════════════════════
with t_records:
    if not st.session_state.loaded:
        st.markdown(nodata("📋","No Data","Upload files first"),unsafe_allow_html=True)
    else:
        D=st.session_state.master
        rec_src=st.radio("View records from:",["📝 Seminar Attendance","🎯 Leads","💰 Conversion Orders"],horizontal=True,key="rs")
        def to_xl(d):
            buf=io.BytesIO(); d.to_excel(buf,index=False); return buf.getvalue()
        ts=datetime.now().strftime("%Y%m%d_%H%M")

        if rec_src=="📝 Seminar Attendance":
            att_all=D['att']; df=apply_sem_filters(att_all,"rsa_")
            NT=len(df); ct_=(df['Conv Status']=='Converted').sum()
            rt_=df['payment_received'].sum(); dt_=df['total_due'].sum()
            if NT:
                st.markdown(f"""<div class="sbar">
<div class="sbar-cell"><div class="sbar-lbl">Records</div><div class="sbar-val">{NT:,}</div></div>
<div class="sbar-cell"><div class="sbar-lbl">Converted</div><div class="sbar-val" style="color:#4fce8f">{ct_:,}</div></div>
<div class="sbar-cell"><div class="sbar-lbl">Not Converted</div><div class="sbar-val" style="color:#f76f4f">{NT-ct_:,}</div></div>
<div class="sbar-cell"><div class="sbar-lbl">Collected</div><div class="sbar-val" style="color:#4fce8f">{fmt(rt_)}</div></div>
<div class="sbar-cell"><div class="sbar-lbl">Due</div><div class="sbar-val" style="color:#f76f4f">{fmt(dt_)}</div></div>
<div class="sbar-cell"><div class="sbar-lbl">Conv Rate</div><div class="sbar-val">{ct_/NT*100:.1f}%</div></div>
<div class="sbar-cell"><div class="sbar-lbl">Sem Amount</div><div class="sbar-val">{fmt(df[df['Conv Status']=='Converted']['Amount Paid'].sum())}</div></div>
</div>""",unsafe_allow_html=True)
            cols=['NAME','Mobile','Place','Seminar Date','Session','Trainer Norm','Conv Status',
                  'Amount Paid','Mode of Payment','service_name','total_amount','payment_received',
                  'total_due','Due Status','status','sales_rep_name','trainer_name',
                  'TRADER','Is our Student ?','PM Label','converted_from','leadsource','Remarks']
            cols=[c for c in cols if c in df.columns]
            ds=df[cols].copy(); ds['Seminar Date']=ds['Seminar Date'].dt.strftime("%d %b %Y")
            st.markdown('<div class="asec"><div class="asec-t">📄 Attendee Records</div>',unsafe_allow_html=True)
            st.dataframe(ds,use_container_width=True,height=440,hide_index=True,
                column_config={"total_amount":st.column_config.NumberColumn("Course Amt",format="₹%.0f"),"payment_received":st.column_config.NumberColumn("Collected",format="₹%.0f"),"total_due":st.column_config.NumberColumn("Due",format="₹%.0f"),"Amount Paid":st.column_config.NumberColumn("Sem Paid",format="₹%.0f"),"service_name":st.column_config.TextColumn("Course"),"sales_rep_name":st.column_config.TextColumn("Sales Rep"),"trainer_name":st.column_config.TextColumn("Trainer"),"Trainer Norm":st.column_config.TextColumn("Seminar Trainer"),"PM Label":st.column_config.TextColumn("Payment Mode"),"converted_from":st.column_config.TextColumn("Lead Source Type")})
            st.markdown('</div>',unsafe_allow_html=True)
            # Location summary
            st.markdown('<div class="asec"><div class="asec-t">📍 Location Summary</div>',unsafe_allow_html=True)
            ls=df.groupby('Place').agg(Attended=('NAME','count'),Converted=('Conv Status',lambda x:(x=='Converted').sum()),Revenue=('total_amount','sum'),Collected=('payment_received','sum'),Due=('total_due','sum'),SemAmt=('Amount Paid','sum')).reset_index()
            ls['Conv Rate']=(ls['Converted']/ls['Attended']*100).round(1).astype(str)+'%'
            for c in ['Revenue','Collected','Due','SemAmt']: ls[c]=ls[c].apply(fmt)
            st.dataframe(ls,use_container_width=True,hide_index=True,height=280)
            st.markdown('</div>',unsafe_allow_html=True)
            st.markdown('<div class="asec"><div class="asec-t">⬇️ Export</div>',unsafe_allow_html=True)
            e1,e2,e3,e4=st.columns(4)
            e1.download_button("⬇️ All Filtered",to_xl(df[cols]),file_name=f"seminar_{ts}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            e2.download_button("⬇️ Converted Only",to_xl(df[df['Conv Status']=='Converted'][cols]),file_name=f"converted_{ts}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            e3.download_button("⬇️ Has Due",to_xl(df[df['total_due']>0][cols]),file_name=f"has_due_{ts}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            e4.download_button("⬇️ Location Summary",to_xl(ls),file_name=f"loc_{ts}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            st.markdown('</div>',unsafe_allow_html=True)

        elif rec_src=="🎯 Leads":
            leads_all=D['leads']
            st.markdown('<div class="asec"><div class="asec-t">🔍 Filter Lead Records</div>',unsafe_allow_html=True)
            lrf1,lrf2,lrf3,lrf4=st.columns(4); lrf5,lrf6,lrf7,lrf8=st.columns(4)
            lr_cf  =lrf1.multiselect("✅ Converted From",["Webinar","Non Webinar"],key="lr_cf")
            lr_ls  =lrf2.multiselect("📢 Lead Source",sorted([x for x in leads_all['leadsource'].unique() if x!='Unknown']),key="lr_ls")
            lr_cp  =lrf3.multiselect("📣 Campaign",sorted([x for x in leads_all['campaign_name'].unique() if x!='Unknown']),key="lr_cp")
            lr_lst =lrf4.multiselect("📊 Lead Status",sorted([x for x in leads_all['leadstatus'].unique() if x!='Unknown']),key="lr_lst")
            lr_stg =lrf5.multiselect("🏷 Stage",sorted([x for x in leads_all['stage_name'].unique() if x!='Unknown']),key="lr_stg")
            lr_lo  =lrf6.multiselect("👤 Lead Owner",sorted([x for x in leads_all['leadownername'].unique() if x!='Unknown']),key="lr_lo")
            lr_st  =lrf7.multiselect("📍 State",sorted([x for x in leads_all['state'].unique() if x!='Unknown']),key="lr_st")
            lr_att =lrf8.selectbox("📞 Attempted?",["All","Attempted","Unattempted"],key="lr_att")
            lr_search=st.text_input("🔍 Search Name/Phone/Email",placeholder="Type to search…",key="lr_sr")
            st.markdown('</div>',unsafe_allow_html=True)
            ld=leads_all.copy()
            if lr_cf:  ld=ld[ld['converted_from'].isin(lr_cf)]
            if lr_ls:  ld=ld[ld['leadsource'].isin(lr_ls)]
            if lr_cp:  ld=ld[ld['campaign_name'].isin(lr_cp)]
            if lr_lst: ld=ld[ld['leadstatus'].isin(lr_lst)]
            if lr_stg: ld=ld[ld['stage_name'].isin(lr_stg)]
            if lr_lo:  ld=ld[ld['leadownername'].isin(lr_lo)]
            if lr_st:  ld=ld[ld['state'].isin(lr_st)]
            if lr_att!="All":
                ld=ld[ld['Attempted/Unattempted'].astype(str).str.strip().str.lower()==lr_att.lower()]
            if lr_search:
                mask=pd.Series([False]*len(ld))
                for c in ['name','phone','email']:
                    if c in ld.columns: mask=mask|ld[c].astype(str).str.lower().str.contains(lr_search.lower(),na=False)
                ld=ld[mask]
            NLR=len(ld); cLR=ld['Is Converted'].sum()
            st.markdown(f"""<div class="sbar">
<div class="sbar-cell"><div class="sbar-lbl">Records</div><div class="sbar-val">{NLR:,}</div></div>
<div class="sbar-cell"><div class="sbar-lbl">Converted</div><div class="sbar-val" style="color:#4fce8f">{cLR:,}</div></div>
<div class="sbar-cell"><div class="sbar-lbl">Webinar</div><div class="sbar-val" style="color:#4fd8f7">{(ld['converted_from']=='Webinar').sum():,}</div></div>
<div class="sbar-cell"><div class="sbar-lbl">Non-Webinar</div><div class="sbar-val" style="color:#b44fe7">{(ld['converted_from']=='Non Webinar').sum():,}</div></div>
<div class="sbar-cell"><div class="sbar-lbl">Conv Rate</div><div class="sbar-val">{cLR/NLR*100:.1f}%</div></div>
</div>""" if NLR else "",unsafe_allow_html=True)
            ld_cols=['name','phone','email','leaddate','converted_from','leadsource','campaign_name','leadstatus','stage_name','leadownername','servicename','state','Attempted/Unattempted','remarks']
            ld_cols=[c for c in ld_cols if c in ld.columns]; ld_show=ld[ld_cols].copy()
            ld_show['leaddate']=ld_show['leaddate'].dt.strftime("%d %b %Y")
            st.markdown('<div class="asec"><div class="asec-t">📄 Lead Records</div>',unsafe_allow_html=True)
            st.dataframe(ld_show,use_container_width=True,height=440,hide_index=True)
            st.markdown('</div>',unsafe_allow_html=True)
            st.markdown('<div class="asec"><div class="asec-t">⬇️ Export</div>',unsafe_allow_html=True)
            e1,e2,e3=st.columns(3)
            e1.download_button("⬇️ All Filtered",to_xl(ld[ld_cols]),file_name=f"leads_{ts}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            e2.download_button("⬇️ Webinar Only",to_xl(ld[ld['converted_from']=='Webinar'][ld_cols]),file_name=f"webinar_{ts}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            e3.download_button("⬇️ Non-Webinar Only",to_xl(ld[ld['converted_from']=='Non Webinar'][ld_cols]),file_name=f"nonwebinar_{ts}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            st.markdown('</div>',unsafe_allow_html=True)

        else:  # Conversion Orders
            cv_all=D['conv']
            st.markdown('<div class="asec"><div class="asec-t">🔍 Filter Conversion Records</div>',unsafe_allow_html=True)
            crf1,crf2,crf3,crf4=st.columns(4)
            cr_st =crf1.multiselect("📌 Status",   sorted(cv_all['status'].dropna().unique()),         key="cr_st")
            cr_sv =crf2.multiselect("📚 Service",  sorted(cv_all['service_name'].dropna().unique()),   key="cr_sv")
            cr_rep=crf3.multiselect("👤 Sales Rep",sorted(cv_all['sales_rep_name'].dropna().unique()),key="cr_re")
            cr_pm =crf4.multiselect("💳 PM Mode",  sorted(cv_all['PM Label'].dropna().unique()),       key="cr_pm")
            cr_search=st.text_input("🔍 Search Name/Phone",placeholder="Type to search…",key="cr_sr")
            st.markdown('</div>',unsafe_allow_html=True)
            cr=cv_all.copy()
            if cr_st:  cr=cr[cr['status'].isin(cr_st)]
            if cr_sv:  cr=cr[cr['service_name'].isin(cr_sv)]
            if cr_rep: cr=cr[cr['sales_rep_name'].isin(cr_rep)]
            if cr_pm:  cr=cr[cr['PM Label'].isin(cr_pm)]
            if cr_search:
                mask=pd.Series([False]*len(cr))
                for c in ['student_name','phone']:
                    if c in cr.columns: mask=mask|cr[c].astype(str).str.lower().str.contains(cr_search.lower(),na=False)
                cr=cr[mask]
            NCR=len(cr)
            st.markdown(f"""<div class="sbar">
<div class="sbar-cell"><div class="sbar-lbl">Orders</div><div class="sbar-val">{NCR:,}</div></div>
<div class="sbar-cell"><div class="sbar-lbl">Revenue</div><div class="sbar-val" style="color:#4fce8f">{fmt(cr['total_amount'].sum())}</div></div>
<div class="sbar-cell"><div class="sbar-lbl">Collected</div><div class="sbar-val" style="color:#4fce8f">{fmt(cr['payment_received'].sum())}</div></div>
<div class="sbar-cell"><div class="sbar-lbl">Due</div><div class="sbar-val" style="color:#f76f4f">{fmt(cr['total_due'].sum())}</div></div>
<div class="sbar-cell"><div class="sbar-lbl">Active</div><div class="sbar-val">{(cr['status']=='Active').sum():,}</div></div>
</div>""" if NCR else "",unsafe_allow_html=True)
            cr_cols=['student_name','phone','email','order_date','service_name','total_amount','payment_received','total_due','status','PM Label','sales_rep_name','trainer_name','batch_date']
            cr_cols=[c for c in cr_cols if c in cr.columns]; cr_show=cr[cr_cols].copy()
            cr_show['order_date']=cr_show['order_date'].dt.strftime("%d %b %Y")
            st.markdown('<div class="asec"><div class="asec-t">📄 Order Records</div>',unsafe_allow_html=True)
            st.dataframe(cr_show,use_container_width=True,height=440,hide_index=True,
                column_config={"total_amount":st.column_config.NumberColumn("Order Amt",format="₹%.0f"),"payment_received":st.column_config.NumberColumn("Collected",format="₹%.0f"),"total_due":st.column_config.NumberColumn("Due",format="₹%.0f"),"service_name":st.column_config.TextColumn("Course"),"sales_rep_name":st.column_config.TextColumn("Sales Rep"),"trainer_name":st.column_config.TextColumn("Trainer"),"PM Label":st.column_config.TextColumn("Payment Mode")})
            st.markdown('</div>',unsafe_allow_html=True)
            st.markdown('<div class="asec"><div class="asec-t">⬇️ Export</div>',unsafe_allow_html=True)
            e1,e2,e3=st.columns(3)
            e1.download_button("⬇️ All Orders",to_xl(cr[cr_cols]),file_name=f"orders_{ts}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            e2.download_button("⬇️ Active Only",to_xl(cr[cr['status']=='Active'][cr_cols]),file_name=f"active_{ts}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            e3.download_button("⬇️ Has Due Balance",to_xl(cr[cr['total_due']>0][cr_cols]),file_name=f"due_{ts}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            st.markdown('</div>',unsafe_allow_html=True)

st.markdown("</div>",unsafe_allow_html=True)
