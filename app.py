import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
from datetime import datetime

st.set_page_config(page_title="Seminar Attendee Dashboard", page_icon="🎯", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif;}
.main{background:#0d0f14;}
.block-container{padding:1.5rem 2rem 2rem 2rem;max-width:1600px;}

.dash-header{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:1.2rem;padding-bottom:1rem;border-bottom:1px solid rgba(255,255,255,0.07);}
.dash-title{font-size:21px;font-weight:600;color:#f0f2f7;letter-spacing:-0.3px;}
.dash-sub{font-size:12px;color:#6b7280;margin-top:3px;font-family:'DM Mono',monospace;}
.live-pill{background:rgba(79,142,247,0.15);color:#4f8ef7;border:1px solid rgba(79,142,247,0.3);padding:4px 12px;border-radius:99px;font-size:11px;font-weight:500;}

.metric-card{background:#161922;border:1px solid rgba(255,255,255,0.07);border-radius:12px;padding:1rem 1.15rem;position:relative;overflow:hidden;margin-bottom:12px;}
.metric-card::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,#4f8ef7,#7c5af7);opacity:.7;}
.m-label{font-size:11px;color:#6b7280;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px;}
.m-value{font-size:24px;font-weight:600;color:#f0f2f7;line-height:1;font-family:'DM Mono',monospace;}
.m-badge{display:inline-block;font-size:11px;padding:2px 8px;border-radius:99px;margin-top:7px;font-weight:500;}
.badge-up{background:rgba(52,211,153,0.12);color:#34d399;}
.badge-down{background:rgba(248,113,113,0.12);color:#f87171;}
.badge-warn{background:rgba(251,191,36,0.12);color:#fbbf24;}
.badge-neu{background:rgba(107,114,128,0.12);color:#9ca3af;}
.badge-blue{background:rgba(79,142,247,0.12);color:#4f8ef7;}

.filter-label{font-size:11px;font-weight:500;color:#6b7280;text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px;}
.section-head{font-size:11px;font-weight:500;color:#6b7280;text-transform:uppercase;letter-spacing:.07em;margin-bottom:10px;}
.section-divider{border:none;border-top:1px solid rgba(255,255,255,0.07);margin:1.2rem 0;}
.filter-bar{background:#161922;border:1px solid rgba(255,255,255,0.07);border-radius:12px;padding:1rem 1.2rem;margin-bottom:1.2rem;}
</style>
""", unsafe_allow_html=True)

# ── Palette ───────────────────────────────────────────────────────────────────
BG   = "#161922"; GRID = "rgba(255,255,255,0.05)"; TEXT = "#9ca3af"
BLUE = "#4f8ef7"; PURP = "#7c5af7"; GRN  = "#34d399"
AMB  = "#fbbf24"; RED  = "#f87171"; GRAY = "#6b7280"; TEAL = "#2dd4bf"

BASE_LAYOUT = dict(paper_bgcolor=BG, plot_bgcolor=BG,
                   font=dict(family="DM Sans", color=TEXT, size=12),
                   margin=dict(l=0, r=0, t=8, b=0))

def kpi(label, value, badge, kind="neu"):
    bc = {"up":"badge-up","down":"badge-down","warn":"badge-warn","neu":"badge-neu","blue":"badge-blue"}[kind]
    return f"""<div class="metric-card"><div class="m-label">{label}</div>
<div class="m-value">{value}</div><span class="m-badge {bc}">{badge}</span></div>"""

def fmt_inr(v):
    if v >= 1e7:  return f"₹{v/1e7:.2f} Cr"
    if v >= 1e5:  return f"₹{v/1e5:.1f} L"
    return f"₹{v:,.0f}"

PAYMENT_MODE_MAP = {
    "mode1":"Full Payment","mode2":"Installment","mode3":"EMI",
    "mode4":"Partial","mode5":"Other","mode6":"Other",
    "mode8":"Other","mode10":"Other","mode11":"Other",
    "mode12":"Other","mode13":"Scholarship/Discount"
}

# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
if "master" not in st.session_state: st.session_state.master = None
if "loaded" not in st.session_state: st.session_state.loaded = False

# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADER
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def load_and_merge(sem_bytes, conv_bytes):
    df_sem  = pd.read_csv(io.BytesIO(sem_bytes))
    df_conv = pd.read_excel(io.BytesIO(conv_bytes))

    # ── Normalise Seminar CSV ─────────────────────────────────────────────────
    df_sem.columns = df_sem.columns.str.strip()
    df_sem['Is Attended ?']        = df_sem['Is Attended ?'].astype(str).str.strip().str.upper()
    df_sem['Is Converted ?']       = df_sem['Is Converted ?'].astype(str).str.strip().str.upper()
    df_sem['Session']              = df_sem['Session'].astype(str).str.strip().str.upper()
    df_sem['TRADER']               = df_sem['TRADER'].astype(str).str.strip().str.upper()
    df_sem['Is our Student ?']     = df_sem['Is our Student ?'].astype(str).str.strip().str.upper()
    df_sem['Trainer / Presenter']  = df_sem['Trainer / Presenter'].astype(str).str.strip().str.upper()
    df_sem['Place']                = df_sem['Place'].astype(str).str.strip().str.upper()
    df_sem['Seminar Date']         = pd.to_datetime(df_sem['Seminar Date'], errors='coerce', dayfirst=True)
    df_sem['Amount Paid']          = pd.to_numeric(df_sem['Amount Paid'], errors='coerce').fillna(0)
    df_sem['mobile_clean']         = df_sem['Mobile'].astype(str).str.replace(r'\D','',regex=True).str[-10:]

    # Keep only attended
    attended = df_sem[df_sem['Is Attended ?'] == 'YES'].copy().reset_index(drop=True)

    # Standardise conversion status
    attended['Conversion Status'] = attended['Is Converted ?'].apply(
        lambda x: 'Converted' if x in ['CONVERTED','YES'] else ('Not Converted' if 'NOT' in str(x) else 'Not Converted')
    )

    # ── Normalise Conversion List ─────────────────────────────────────────────
    df_conv['order_date']       = pd.to_datetime(df_conv['order_date'], errors='coerce', utc=True).dt.tz_localize(None)
    df_conv['payment_received'] = pd.to_numeric(df_conv['payment_received'], errors='coerce').fillna(0)
    df_conv['total_amount']     = pd.to_numeric(df_conv['total_amount'],     errors='coerce').fillna(0)
    df_conv['total_due']        = pd.to_numeric(df_conv['total_due'],        errors='coerce').fillna(0)
    df_conv['total_gst']        = pd.to_numeric(df_conv['total_gst'],        errors='coerce').fillna(0)
    df_conv['phone_clean']      = df_conv['phone'].astype(str).str.replace(r'\D','',regex=True).str[-10:]
    df_conv['Payment Mode Label']= df_conv['payment_mode'].map(PAYMENT_MODE_MAP).fillna(df_conv['payment_mode'])
    df_conv['service_name']     = df_conv['service_name'].astype(str).str.strip()
    df_conv['sales_rep_name']   = df_conv['sales_rep_name'].astype(str).str.strip()
    df_conv['trainer_clean']    = df_conv['trainer'].astype(str).str.split(' - ').str[-1].str.strip()

    # ── Merge on phone ────────────────────────────────────────────────────────
    merged = attended.merge(
        df_conv[['phone_clean','orderID','order_date','service_code','service_name',
                 'payment_received','total_amount','total_due','total_gst',
                 'payment_mode','Payment Mode Label','status',
                 'sales_rep_name','trainer','trainer_clean','student_invid',
                 'is_refunded','is_shortClosed','batch_date','coupon_code']],
        left_on='mobile_clean', right_on='phone_clean', how='left'
    )

    # ── Payment due status label ──────────────────────────────────────────────
    def due_label(row):
        if pd.isna(row.get('total_due')): return 'No Order'
        if row['total_due'] <= 0:         return 'Fully Paid'
        if row['total_due'] > 0 and row['total_amount'] > 0:
            pct = row['total_due'] / row['total_amount']
            return 'Partially Paid' if pct < 1 else 'Fully Due'
        return 'No Order'
    merged['Due Status'] = merged.apply(due_label, axis=1)

    return merged

# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="dash-header">
  <div>
    <div class="dash-title">🎯 Seminar Attendee Intelligence Dashboard</div>
    <div class="dash-sub">Attendance · Conversion · Payment · Course · Sales Rep · Trainer · {datetime.now().strftime("%d %b %Y")}</div>
  </div>
  <div class="live-pill">SEMINAR FOCUSED</div>
</div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tab_upload, tab_dash, tab_detail = st.tabs(["📁  Upload Files", "📊  Dashboard & Analytics", "📋  Detailed Records"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
with tab_upload:
    st.markdown("#### Upload your 2 source files")
    st.markdown("The dashboard merges them automatically on **phone number** to link seminar attendance with order/payment data.")
    st.markdown("<hr style='border-color:rgba(255,255,255,0.07)'>", unsafe_allow_html=True)

    uc1, uc2 = st.columns(2)
    with uc1:
        st.markdown("**📝 Seminar Updated Sheet** (`.csv`)")
        st.caption("Contains: Name, Mobile, Place, Trainer, Seminar Date, Session, Attended, Converted, Amount Paid")
        f_sem = st.file_uploader("Seminar CSV", type=["csv"], key="up_sem", label_visibility="collapsed")
    with uc2:
        st.markdown("**📋 Conversion List** (`.xlsx`)")
        st.caption("Contains: Order details, Course, Payment, Due, Status, Sales Rep, Trainer")
        f_conv = st.file_uploader("Conversion XLSX", type=["xlsx","xls"], key="up_conv", label_visibility="collapsed")

    if f_sem and f_conv:
        with st.spinner("Merging and processing data…"):
            df = load_and_merge(f_sem.read(), f_conv.read())
        st.session_state.master = df
        st.session_state.loaded = True

        total  = len(df)
        matched= df['orderID'].notna().sum()
        conv   = (df['Conversion Status'] == 'Converted').sum()
        st.success(f"✅ Done! **{total:,}** attendees loaded · **{matched:,}** matched to orders · **{conv:,}** converted")

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Total Attended",    f"{total:,}")
        c2.metric("Matched to Orders", f"{matched:,}", f"{matched/total*100:.1f}%")
        c3.metric("Converted",         f"{conv:,}",    f"{conv/total*100:.1f}%")
        c4.metric("Seminar Locations", str(df['Place'].nunique()))
    elif f_sem or f_conv:
        st.info("⏳ Please upload both files to proceed.")
    else:
        st.markdown("""
        <div style="background:#161922;border:1.5px dashed rgba(79,142,247,0.3);border-radius:12px;padding:2rem;text-align:center;">
          <div style="font-size:15px;font-weight:500;color:#f0f2f7;margin-bottom:6px;">📂 No files uploaded yet</div>
          <div style="font-size:13px;color:#6b7280;">Upload both files above to unlock the full dashboard</div>
        </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
with tab_dash:
    if not st.session_state.loaded:
        st.warning("⚠️ Upload both files in the **Upload Files** tab first.")
    else:
        df_all = st.session_state.master.copy()

        # ══════════════════════════════════════════════════════════════════════
        # FILTER BAR
        # ══════════════════════════════════════════════════════════════════════
        with st.expander("🔽  Filters — click to expand / collapse", expanded=True):
            fa, fb, fc, fd = st.columns(4)
            fe, ff, fg, fh = st.columns(4)

            all_places   = sorted(df_all['Place'].dropna().unique())
            all_trainers = sorted(df_all['Trainer / Presenter'].dropna().unique())
            all_dates    = sorted(df_all['Seminar Date'].dropna().unique())
            all_sessions = ["MORNING","EVENING"]
            all_conv     = ["Converted","Not Converted"]
            all_due      = sorted(df_all['Due Status'].dropna().unique())
            all_courses  = sorted(df_all['service_name'].dropna().unique())
            all_reps     = sorted(df_all['sales_rep_name'].dropna().unique())
            all_status   = sorted(df_all['status'].dropna().unique())

            sel_place    = fa.multiselect("📍 Location",         all_places,   key="f_place")
            sel_trainer  = fb.multiselect("🎤 Trainer",          all_trainers, key="f_trainer")
            sel_date     = fc.multiselect("📅 Seminar Date",     [d.strftime("%d %b %Y") if pd.notna(d) else "" for d in all_dates], key="f_date")
            sel_session  = fd.multiselect("🕐 Session",          all_sessions, key="f_session")
            sel_conv     = fe.multiselect("✅ Conversion Status", all_conv,    key="f_conv")
            sel_due      = ff.multiselect("💳 Due Status",        all_due,     key="f_due")
            sel_course   = fg.multiselect("📚 Course",            all_courses, key="f_course")
            sel_rep      = fh.multiselect("👤 Sales Rep",         all_reps,    key="f_rep")

            fr1, fr2, fr3 = st.columns(3)
            sel_status   = fr1.multiselect("📌 Order Status", all_status, key="f_status")
            is_trader    = fr2.selectbox("🔄 Is Trader?",   ["All","Yes","No"], key="f_trader")
            is_our_stu   = fr3.selectbox("🎓 Existing Student?", ["All","Yes","No"], key="f_ourstu")

        # ── Apply filters ─────────────────────────────────────────────────────
        df = df_all.copy()
        if sel_place:    df = df[df['Place'].isin(sel_place)]
        if sel_trainer:  df = df[df['Trainer / Presenter'].isin(sel_trainer)]
        if sel_date:
            fmt_dates = [pd.to_datetime(d, format="%d %b %Y") for d in sel_date]
            df = df[df['Seminar Date'].isin(fmt_dates)]
        if sel_session:  df = df[df['Session'].isin(sel_session)]
        if sel_conv:     df = df[df['Conversion Status'].isin(sel_conv)]
        if sel_due:      df = df[df['Due Status'].isin(sel_due)]
        if sel_course:   df = df[df['service_name'].isin(sel_course)]
        if sel_rep:      df = df[df['sales_rep_name'].isin(sel_rep)]
        if sel_status:   df = df[df['status'].isin(sel_status)]
        if is_trader != "All":
            df = df[df['TRADER'].str.upper().isin(['YES'] if is_trader=="Yes" else ['NO'])]
        if is_our_stu != "All":
            df = df[df['Is our Student ?'].str.upper().isin(['YES','STUDENT'] if is_our_stu=="Yes" else ['NO'])]

        total_f  = len(df)
        conv_f   = (df['Conversion Status']=='Converted').sum()
        nconv_f  = total_f - conv_f
        conv_r   = conv_f/total_f*100 if total_f else 0
        matched_f= df['orderID'].notna().sum()
        amt_coll = df['payment_received'].sum()
        amt_due  = df['total_due'].sum()
        amt_tot  = df['total_amount'].sum()
        fully_paid = (df['Due Status']=='Fully Paid').sum()
        has_due    = (df['Due Status'].isin(['Partially Paid','Fully Due'])).sum()

        # ── KPI Row ───────────────────────────────────────────────────────────
        st.markdown("##### Summary KPIs — filtered view")
        k1,k2,k3,k4,k5,k6 = st.columns(6)
        with k1: st.markdown(kpi("Total Attended", f"{total_f:,}", f"{len(df_all):,} total", "blue"), unsafe_allow_html=True)
        with k2: st.markdown(kpi("Converted", f"{conv_f:,}", f"{conv_r:.1f}% rate", "up" if conv_r>15 else "warn"), unsafe_allow_html=True)
        with k3: st.markdown(kpi("Not Converted", f"{nconv_f:,}", f"{100-conv_r:.1f}% of attended", "down"), unsafe_allow_html=True)
        with k4: st.markdown(kpi("Course Revenue", fmt_inr(amt_tot), "Gross order value", "up"), unsafe_allow_html=True)
        with k5: st.markdown(kpi("Collected", fmt_inr(amt_coll), f"Due: {fmt_inr(amt_due)}", "up"), unsafe_allow_html=True)
        with k6: st.markdown(kpi("Fully Paid", f"{fully_paid:,}", f"Has due: {has_due:,}", "up" if fully_paid > has_due else "warn"), unsafe_allow_html=True)

        st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

        # ── Row 1: Conversion by Location + Session ───────────────────────────
        r1c1, r1c2 = st.columns([3,2])

        with r1c1:
            st.markdown('<div class="section-head">Conversion rate by location</div>', unsafe_allow_html=True)
            loc_grp = df.groupby('Place').agg(
                Attended=('NAME','count'),
                Converted=('Conversion Status', lambda x: (x=='Converted').sum())
            ).reset_index()
            loc_grp['Rate'] = (loc_grp['Converted']/loc_grp['Attended']*100).round(1)
            loc_grp = loc_grp.sort_values('Attended', ascending=True)
            fig = go.Figure()
            fig.add_trace(go.Bar(name='Attended',  y=loc_grp['Place'], x=loc_grp['Attended'],
                                 orientation='h', marker_color=BLUE, opacity=0.7, marker_line_width=0))
            fig.add_trace(go.Bar(name='Converted', y=loc_grp['Place'], x=loc_grp['Converted'],
                                 orientation='h', marker_color=GRN, marker_line_width=0))
            fig.update_layout(**BASE_LAYOUT, height=max(280, len(loc_grp)*38),
                barmode='overlay', bargap=0.3,
                legend=dict(orientation="h",x=0,y=1.08,font=dict(size=11)),
                xaxis=dict(showgrid=False), yaxis=dict(showgrid=False))
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

        with r1c2:
            st.markdown('<div class="section-head">Conversion funnel</div>', unsafe_allow_html=True)
            fig2 = go.Figure(go.Funnel(
                y=["Registered", "Attended", "Converted"],
                x=[len(df_all), total_f, conv_f],
                textinfo="value+percent initial",
                marker=dict(color=[BLUE, PURP, GRN]),
                connector=dict(line=dict(color=BG, width=3)),
            ))
            fig2.update_layout(**BASE_LAYOUT, height=200)
            st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar":False})

            st.markdown('<div class="section-head" style="margin-top:12px">Session split</div>', unsafe_allow_html=True)
            sess = df.groupby('Session').agg(
                Attended=('NAME','count'),
                Converted=('Conversion Status', lambda x: (x=='Converted').sum())
            ).reset_index()
            fig3 = go.Figure()
            fig3.add_trace(go.Bar(name='Attended',  x=sess['Session'], y=sess['Attended'],
                                  marker_color=BLUE, opacity=0.7, marker_line_width=0))
            fig3.add_trace(go.Bar(name='Converted', x=sess['Session'], y=sess['Converted'],
                                  marker_color=GRN, marker_line_width=0))
            fig3.update_layout(**BASE_LAYOUT, height=180, barmode='group', bargap=0.35,
                legend=dict(orientation="h",x=0,y=1.12,font=dict(size=11)),
                xaxis=dict(showgrid=False), yaxis=dict(gridcolor=GRID))
            st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar":False})

        # ── Row 2: Date trend + Due status ────────────────────────────────────
        r2c1, r2c2 = st.columns([3,2])

        with r2c1:
            st.markdown('<div class="section-head">Attendance & conversion by seminar date</div>', unsafe_allow_html=True)
            date_grp = df.groupby('Seminar Date').agg(
                Attended=('NAME','count'),
                Converted=('Conversion Status', lambda x: (x=='Converted').sum())
            ).reset_index().dropna(subset=['Seminar Date']).sort_values('Seminar Date')
            date_grp['Label'] = date_grp['Seminar Date'].dt.strftime("%d %b")
            fig4 = go.Figure()
            fig4.add_trace(go.Bar(name='Attended',  x=date_grp['Label'], y=date_grp['Attended'],
                                  marker_color=BLUE, opacity=0.7, marker_line_width=0))
            fig4.add_trace(go.Scatter(name='Converted', x=date_grp['Label'], y=date_grp['Converted'],
                                      line=dict(color=GRN,width=2.5), mode='lines+markers',
                                      marker=dict(size=7,color=GRN,line=dict(color=BG,width=2)),
                                      yaxis='y2'))
            fig4.update_layout(**BASE_LAYOUT, height=240,
                bargap=0.3,
                legend=dict(orientation="h",x=0,y=1.1,font=dict(size=11)),
                xaxis=dict(showgrid=False),
                yaxis=dict(gridcolor=GRID, title="Attended"),
                yaxis2=dict(overlaying='y', side='right', showgrid=False, title="Converted"))
            st.plotly_chart(fig4, use_container_width=True, config={"displayModeBar":False})

        with r2c2:
            st.markdown('<div class="section-head">Payment due status</div>', unsafe_allow_html=True)
            due_counts = df['Due Status'].value_counts()
            color_map = {'Fully Paid':GRN, 'Partially Paid':AMB, 'Fully Due':RED, 'No Order':GRAY}
            fig5 = go.Figure(go.Pie(
                labels=due_counts.index, values=due_counts.values, hole=0.62,
                marker=dict(colors=[color_map.get(l,BLUE) for l in due_counts.index],
                            line=dict(color=BG,width=2)),
                textfont=dict(size=11),
                hovertemplate="%{label}: %{value}<extra></extra>",
            ))
            fig5.update_layout(**BASE_LAYOUT, height=240,
                legend=dict(orientation="h",x=0.5,xanchor="center",y=-0.2,font=dict(size=11)),
                annotations=[dict(text="Due<br>Status",x=0.5,y=0.5,showarrow=False,
                                  font=dict(size=11,color=TEXT))])
            st.plotly_chart(fig5, use_container_width=True, config={"displayModeBar":False})

        # ── Row 3: Course + Sales Rep ─────────────────────────────────────────
        r3c1, r3c2 = st.columns(2)

        with r3c1:
            st.markdown('<div class="section-head">Top courses by converted students</div>', unsafe_allow_html=True)
            course_df = df[df['Conversion Status']=='Converted'].groupby('service_name').agg(
                Count=('NAME','count'), Revenue=('total_amount','sum')
            ).reset_index().sort_values('Count', ascending=True).tail(10)
            course_df['Short'] = course_df['service_name'].apply(lambda x: x[:32]+'…' if len(x)>32 else x)
            fig6 = go.Figure(go.Bar(
                x=course_df['Count'], y=course_df['Short'], orientation='h',
                marker_color=PURP, marker_line_width=0,
                text=course_df['Count'], textposition='outside', textfont=dict(size=11, color=TEXT),
            ))
            fig6.update_layout(**BASE_LAYOUT, height=280,
                xaxis=dict(showgrid=False,showticklabels=False), yaxis=dict(showgrid=False), bargap=0.3)
            st.plotly_chart(fig6, use_container_width=True, config={"displayModeBar":False})

        with r3c2:
            st.markdown('<div class="section-head">Sales rep performance (conversions)</div>', unsafe_allow_html=True)
            rep_df = df[df['Conversion Status']=='Converted'].groupby('sales_rep_name').agg(
                Converted=('NAME','count'),
                Revenue=('total_amount','sum'),
                Due=('total_due','sum')
            ).reset_index().sort_values('Converted', ascending=True).tail(10)
            fig7 = go.Figure()
            fig7.add_trace(go.Bar(name='Conversions', y=rep_df['sales_rep_name'], x=rep_df['Converted'],
                                  orientation='h', marker_color=TEAL, marker_line_width=0,
                                  text=rep_df['Converted'], textposition='outside', textfont=dict(size=11,color=TEXT)))
            fig7.update_layout(**BASE_LAYOUT, height=280,
                xaxis=dict(showgrid=False,showticklabels=False), yaxis=dict(showgrid=False),
                bargap=0.3, showlegend=False)
            st.plotly_chart(fig7, use_container_width=True, config={"displayModeBar":False})

        # ── Row 4: Trainer + Payment mode ─────────────────────────────────────
        r4c1, r4c2 = st.columns(2)

        with r4c1:
            st.markdown('<div class="section-head">Trainer wise attendance & conversion</div>', unsafe_allow_html=True)
            tr_df = df.groupby('Trainer / Presenter').agg(
                Attended=('NAME','count'),
                Converted=('Conversion Status', lambda x: (x=='Converted').sum())
            ).reset_index()
            tr_df['Rate'] = (tr_df['Converted']/tr_df['Attended']*100).round(1)
            tr_df = tr_df.sort_values('Attended', ascending=True)
            tr_df['Short'] = tr_df['Trainer / Presenter'].apply(lambda x: x[:30]+'…' if len(x)>30 else x)
            fig8 = go.Figure()
            fig8.add_trace(go.Bar(name='Attended',  y=tr_df['Short'], x=tr_df['Attended'],
                                  orientation='h', marker_color=BLUE, opacity=0.6, marker_line_width=0))
            fig8.add_trace(go.Bar(name='Converted', y=tr_df['Short'], x=tr_df['Converted'],
                                  orientation='h', marker_color=GRN, marker_line_width=0))
            fig8.update_layout(**BASE_LAYOUT, height=300,
                barmode='overlay', bargap=0.3,
                legend=dict(orientation="h",x=0,y=1.08,font=dict(size=11)),
                xaxis=dict(showgrid=False), yaxis=dict(showgrid=False))
            st.plotly_chart(fig8, use_container_width=True, config={"displayModeBar":False})

        with r4c2:
            st.markdown('<div class="section-head">Payment mode of converted students</div>', unsafe_allow_html=True)
            pm_df = df[df['Conversion Status']=='Converted']['Payment Mode Label'].value_counts()
            fig9 = go.Figure(go.Pie(
                labels=pm_df.index, values=pm_df.values, hole=0.58,
                marker=dict(colors=[BLUE,PURP,GRN,AMB,RED,TEAL,GRAY], line=dict(color=BG,width=2)),
                textfont=dict(size=11),
            ))
            fig9.update_layout(**BASE_LAYOUT, height=300,
                legend=dict(orientation="v",x=1.02,y=0.5,font=dict(size=11)),
                annotations=[dict(text="Mode",x=0.5,y=0.5,showarrow=False,font=dict(size=12,color=TEXT))])
            st.plotly_chart(fig9, use_container_width=True, config={"displayModeBar":False})

        # ── Row 5: Revenue by location + Trader mix ───────────────────────────
        r5c1, r5c2 = st.columns([3,2])

        with r5c1:
            st.markdown('<div class="section-head">Revenue collected vs due by location</div>', unsafe_allow_html=True)
            rev_loc = df.groupby('Place').agg(
                Collected=('payment_received','sum'), Due=('total_due','sum')
            ).reset_index().sort_values('Collected', ascending=True)
            fig10 = go.Figure()
            fig10.add_trace(go.Bar(name='Collected', y=rev_loc['Place'], x=rev_loc['Collected'],
                                   orientation='h', marker_color=GRN, opacity=0.8, marker_line_width=0))
            fig10.add_trace(go.Bar(name='Due',       y=rev_loc['Place'], x=rev_loc['Due'],
                                   orientation='h', marker_color=RED, opacity=0.8, marker_line_width=0))
            fig10.update_layout(**BASE_LAYOUT, height=max(280, len(rev_loc)*38),
                barmode='stack', bargap=0.3,
                legend=dict(orientation="h",x=0,y=1.08,font=dict(size=11)),
                xaxis=dict(gridcolor=GRID,tickformat=",.0f"), yaxis=dict(showgrid=False))
            st.plotly_chart(fig10, use_container_width=True, config={"displayModeBar":False})

        with r5c2:
            st.markdown('<div class="section-head">Trader vs non-trader (attended)</div>', unsafe_allow_html=True)
            trader_counts = df['TRADER'].map(
                lambda x: 'Trader' if x in ['YES','Y','TYES'] else 'Non-Trader'
            ).value_counts()
            fig11 = go.Figure(go.Pie(
                labels=trader_counts.index, values=trader_counts.values, hole=0.6,
                marker=dict(colors=[TEAL, PURP], line=dict(color=BG,width=2)),
                textfont=dict(size=12),
            ))
            fig11.update_layout(**BASE_LAYOUT, height=180,
                legend=dict(orientation="h",x=0.5,xanchor="center",y=-0.2,font=dict(size=11)),
                annotations=[dict(text="Mix",x=0.5,y=0.5,showarrow=False,font=dict(size=12,color=TEXT))])
            st.plotly_chart(fig11, use_container_width=True, config={"displayModeBar":False})

            st.markdown('<div class="section-head" style="margin-top:10px">Existing student vs new</div>', unsafe_allow_html=True)
            stu_counts = df['Is our Student ?'].map(
                lambda x: 'Existing Student' if x in ['YES','STUDENT'] else 'New Lead'
            ).value_counts()
            fig12 = go.Figure(go.Pie(
                labels=stu_counts.index, values=stu_counts.values, hole=0.6,
                marker=dict(colors=[AMB, BLUE], line=dict(color=BG,width=2)),
                textfont=dict(size=12),
            ))
            fig12.update_layout(**BASE_LAYOUT, height=180,
                legend=dict(orientation="h",x=0.5,xanchor="center",y=-0.2,font=dict(size=11)),
                annotations=[dict(text="Student",x=0.5,y=0.5,showarrow=False,font=dict(size=12,color=TEXT))])
            st.plotly_chart(fig12, use_container_width=True, config={"displayModeBar":False})

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — DETAILED RECORDS
# ══════════════════════════════════════════════════════════════════════════════
with tab_detail:
    if not st.session_state.loaded:
        st.warning("⚠️ Upload both files in the **Upload Files** tab first.")
    else:
        df_all = st.session_state.master.copy()

        st.markdown("#### Attendee Records — Full Filterable Table")
        st.markdown("Every seminar attendee with their payment, course, and sales details. Use filters below.")

        # ── Inline filters ────────────────────────────────────────────────────
        tf1,tf2,tf3,tf4 = st.columns(4)
        tf5,tf6,tf7,tf8 = st.columns(4)

        t_place    = tf1.multiselect("📍 Location",         sorted(df_all['Place'].dropna().unique()), key="t_place")
        t_trainer  = tf2.multiselect("🎤 Trainer",          sorted(df_all['Trainer / Presenter'].dropna().unique()), key="t_trainer")
        t_date     = tf3.multiselect("📅 Seminar Date",
                                     [d.strftime("%d %b %Y") for d in sorted(df_all['Seminar Date'].dropna().unique())], key="t_date")
        t_session  = tf4.multiselect("🕐 Session",          ["MORNING","EVENING"], key="t_session")
        t_conv     = tf5.multiselect("✅ Conversion",        ["Converted","Not Converted"], key="t_conv")
        t_due      = tf6.multiselect("💳 Due Status",        sorted(df_all['Due Status'].dropna().unique()), key="t_due")
        t_course   = tf7.multiselect("📚 Course",            sorted(df_all['service_name'].dropna().unique()), key="t_course")
        t_rep      = tf8.multiselect("👤 Sales Rep",         sorted(df_all['sales_rep_name'].dropna().unique()), key="t_rep")

        tg1, tg2, tg3 = st.columns(3)
        t_status   = tg1.multiselect("📌 Order Status", sorted(df_all['status'].dropna().unique()), key="t_status")
        t_trader   = tg2.selectbox("🔄 Trader?",          ["All","Yes","No"], key="t_trader")
        t_search   = tg3.text_input("🔍 Search name / phone / email", key="t_search")

        df_t = df_all.copy()
        if t_place:    df_t = df_t[df_t['Place'].isin(t_place)]
        if t_trainer:  df_t = df_t[df_t['Trainer / Presenter'].isin(t_trainer)]
        if t_date:
            fmt_d = [pd.to_datetime(d, format="%d %b %Y") for d in t_date]
            df_t = df_t[df_t['Seminar Date'].isin(fmt_d)]
        if t_session:  df_t = df_t[df_t['Session'].isin(t_session)]
        if t_conv:     df_t = df_t[df_t['Conversion Status'].isin(t_conv)]
        if t_due:      df_t = df_t[df_t['Due Status'].isin(t_due)]
        if t_course:   df_t = df_t[df_t['service_name'].isin(t_course)]
        if t_rep:      df_t = df_t[df_t['sales_rep_name'].isin(t_rep)]
        if t_status:   df_t = df_t[df_t['status'].isin(t_status)]
        if t_trader != "All":
            df_t = df_t[df_t['TRADER'].isin(['YES','Y','TYES'] if t_trader=="Yes" else ['NO','N'])]
        if t_search:
            mask = pd.Series([False]*len(df_t))
            for c in ['NAME','Mobile','email']:
                if c in df_t.columns:
                    mask = mask | df_t[c].astype(str).str.lower().str.contains(t_search.lower(), na=False)
            df_t = df_t[mask]

        st.markdown(f"**{len(df_t):,}** records shown")

        # Display columns
        display_cols = ['NAME','Mobile','Place','Seminar Date','Session',
                        'Trainer / Presenter','Conversion Status','Amount Paid',
                        'service_name','total_amount','payment_received','total_due',
                        'Due Status','status','sales_rep_name','trainer_clean',
                        'TRADER','Is our Student ?','Payment Mode Label','Remarks']
        display_cols = [c for c in display_cols if c in df_t.columns]

        df_show = df_t[display_cols].copy()
        df_show['Seminar Date'] = df_show['Seminar Date'].dt.strftime("%d %b %Y")

        st.dataframe(
            df_show,
            use_container_width=True,
            height=520,
            hide_index=True,
            column_config={
                "total_amount":     st.column_config.NumberColumn("Course Amount",    format="₹%.0f"),
                "payment_received": st.column_config.NumberColumn("Collected",        format="₹%.0f"),
                "total_due":        st.column_config.NumberColumn("Due Amount",       format="₹%.0f"),
                "Amount Paid":      st.column_config.NumberColumn("Seminar Paid",     format="₹%.0f"),
                "service_name":     st.column_config.TextColumn("Course"),
                "sales_rep_name":   st.column_config.TextColumn("Sales Rep"),
                "trainer_clean":    st.column_config.TextColumn("Trainer (Conv)"),
                "Trainer / Presenter": st.column_config.TextColumn("Seminar Trainer"),
                "Payment Mode Label": st.column_config.TextColumn("Payment Mode"),
            }
        )

        st.markdown("<hr style='border-color:rgba(255,255,255,0.07)'>", unsafe_allow_html=True)

        # ── Summary table by location ─────────────────────────────────────────
        st.markdown("**Summary by Location**")
        loc_sum = df_t.groupby('Place').agg(
            Attended=('NAME','count'),
            Converted=('Conversion Status', lambda x: (x=='Converted').sum()),
            Total_Revenue=('total_amount','sum'),
            Collected=('payment_received','sum'),
            Due=('total_due','sum'),
        ).reset_index()
        loc_sum['Conv Rate'] = (loc_sum['Converted']/loc_sum['Attended']*100).round(1).astype(str) + '%'
        loc_sum['Total_Revenue'] = loc_sum['Total_Revenue'].apply(fmt_inr)
        loc_sum['Collected']     = loc_sum['Collected'].apply(fmt_inr)
        loc_sum['Due']           = loc_sum['Due'].apply(fmt_inr)
        st.dataframe(loc_sum, use_container_width=True, hide_index=True, height=280)

        # ── Export ────────────────────────────────────────────────────────────
        st.markdown("<hr style='border-color:rgba(255,255,255,0.07)'>", unsafe_allow_html=True)
        def to_xl(df):
            buf = io.BytesIO(); df.to_excel(buf, index=False); return buf.getvalue()

        ts = datetime.now().strftime("%Y%m%d_%H%M")
        ex1,ex2,ex3,ex4 = st.columns(4)
        ex1.download_button("⬇️ Export filtered records",     to_xl(df_t[display_cols]),
            file_name=f"seminar_attendees_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        ex2.download_button("⬇️ Export converted only",
            to_xl(df_t[df_t['Conversion Status']=='Converted'][display_cols]),
            file_name=f"converted_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        ex3.download_button("⬇️ Export with due amount",
            to_xl(df_t[df_t['total_due']>0][display_cols]),
            file_name=f"has_due_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        ex4.download_button("⬇️ Export location summary",     to_xl(loc_sum),
            file_name=f"location_summary_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="border-top:1px solid rgba(255,255,255,0.06);margin-top:2rem;padding-top:0.8rem;
     display:flex;justify-content:space-between">
  <span style="font-size:11px;color:#374151;font-family:'DM Mono',monospace">seminar attendee dashboard · {datetime.now().strftime("%B %Y").lower()}</span>
  <span style="font-size:11px;color:#374151">merged on phone number · upload fresh files to refresh</span>
</div>""", unsafe_allow_html=True)
