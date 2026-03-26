import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
from datetime import datetime

st.set_page_config(page_title="Seminar Intelligence Dashboard", page_icon="📊", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif;}
.main{background:#0d0f14;}
.block-container{padding:1.5rem 2rem 2rem 2rem;max-width:1500px;}
.dash-header{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:1.5rem;padding-bottom:1rem;border-bottom:1px solid rgba(255,255,255,0.07);}
.dash-title{font-size:22px;font-weight:600;color:#f0f2f7;letter-spacing:-0.3px;}
.dash-sub{font-size:13px;color:#6b7280;margin-top:3px;font-family:'DM Mono',monospace;}
.live-pill{background:rgba(29,158,117,0.15);color:#34d399;border:1px solid rgba(52,211,153,0.25);padding:4px 12px;border-radius:99px;font-size:11px;font-weight:500;letter-spacing:.05em;}
.metric-card{background:#161922;border:1px solid rgba(255,255,255,0.07);border-radius:12px;padding:1.1rem 1.2rem;position:relative;overflow:hidden;margin-bottom:14px;}
.metric-card::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,#4f8ef7,#7c5af7);opacity:.6;}
.m-label{font-size:11px;color:#6b7280;text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px;}
.m-value{font-size:26px;font-weight:600;color:#f0f2f7;line-height:1;font-family:'DM Mono',monospace;}
.m-badge{display:inline-block;font-size:11px;padding:3px 8px;border-radius:99px;margin-top:8px;font-weight:500;}
.badge-up{background:rgba(29,158,117,0.15);color:#34d399;}
.badge-down{background:rgba(226,75,74,0.15);color:#f87171;}
.badge-neu{background:rgba(107,114,128,0.15);color:#9ca3af;}
.badge-warn{background:rgba(251,191,36,0.15);color:#fbbf24;}
.section-head{font-size:11px;font-weight:500;color:#6b7280;text-transform:uppercase;letter-spacing:.07em;margin-bottom:12px;}
.section-divider{border:none;border-top:1px solid rgba(255,255,255,0.07);margin:1.5rem 0;}
.tab-head{font-size:16px;font-weight:500;color:#f0f2f7;margin-bottom:4px;}
.tab-sub{font-size:13px;color:#6b7280;margin-bottom:1rem;}
.upload-box{background:#161922;border:1.5px dashed rgba(79,142,247,0.35);border-radius:14px;padding:1.5rem;text-align:center;margin-bottom:1rem;}
</style>
""", unsafe_allow_html=True)

# ── Palette ───────────────────────────────────────────────────────────────────
BG   = "#161922"; GRID = "rgba(255,255,255,0.05)"; TEXT = "#9ca3af"
BLUE = "#4f8ef7"; PURP = "#7c5af7"; GRN  = "#34d399"
AMB  = "#fbbf24"; RED  = "#f87171"; GRAY = "#6b7280"; TEAL = "#2dd4bf"

CHART_LAYOUT = dict(paper_bgcolor=BG, plot_bgcolor=BG,
                    font=dict(family="DM Sans", color=TEXT, size=12),
                    margin=dict(l=0, r=0, t=10, b=0))

def kpi_card(label, value, badge, kind="up"):
    bc = {"up":"badge-up","down":"badge-down","neu":"badge-neu","warn":"badge-warn"}[kind]
    return f"""<div class="metric-card"><div class="m-label">{label}</div>
    <div class="m-value">{value}</div>
    <div><span class="m-badge {bc}">{badge}</span></div></div>"""

def fmt_cr(v): return f"₹{v/1e7:.2f}Cr" if v >= 1e7 else (f"₹{v/1e5:.1f}L" if v >= 1e5 else f"₹{v:,.0f}")

# ── Session state ─────────────────────────────────────────────────────────────
for k in ["conv","seminar","indepth","report","loaded"]:
    if k not in st.session_state:
        st.session_state[k] = None
if "loaded" not in st.session_state:
    st.session_state.loaded = False

# ══════════════════════════════════════════════════════════════════════════════
# DATA PROCESSING
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def process_all(conv_bytes, sem_bytes, indepth_bytes, report_bytes):
    results = {}

    # ── Conversion List ───────────────────────────────────────────────────────
    df_conv = pd.read_excel(io.BytesIO(conv_bytes))
    df_conv['order_date'] = pd.to_datetime(df_conv['order_date'], errors='coerce', utc=True)
    df_conv['order_date'] = df_conv['order_date'].dt.tz_localize(None)
    df_conv['month'] = df_conv['order_date'].dt.to_period('M').astype(str)
    df_conv['payment_received'] = pd.to_numeric(df_conv['payment_received'], errors='coerce').fillna(0)
    df_conv['total_amount']     = pd.to_numeric(df_conv['total_amount'],     errors='coerce').fillna(0)
    df_conv['total_due']        = pd.to_numeric(df_conv['total_due'],        errors='coerce').fillna(0)
    results['conv'] = df_conv

    # ── Seminar Updated CSV ───────────────────────────────────────────────────
    df_sem = pd.read_csv(io.BytesIO(sem_bytes))
    df_sem['Is Attended ?']  = df_sem['Is Attended ?'].astype(str).str.strip().str.upper()
    df_sem['Is Converted ?'] = df_sem['Is Converted ?'].astype(str).str.strip().str.upper()
    df_sem['Amount Paid']    = pd.to_numeric(df_sem['Amount Paid'], errors='coerce').fillna(0)
    df_sem['Session']        = df_sem['Session'].astype(str).str.strip().str.upper()
    results['seminar'] = df_sem

    # ── Offline Indepth ───────────────────────────────────────────────────────
    xl = pd.ExcelFile(io.BytesIO(indepth_bytes))
    dfs = []
    for s in xl.sheet_names:
        try:
            df = pd.read_excel(io.BytesIO(indepth_bytes), sheet_name=s)
            if 'student_name' in df.columns:
                df['location'] = s
                dfs.append(df)
        except: pass
    indepth = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
    for col in ['payment_received','total_amount','total_due','total_gst']:
        indepth[col] = pd.to_numeric(indepth.get(col, 0), errors='coerce').fillna(0)
    results['indepth'] = indepth

    # ── Seminar Report ────────────────────────────────────────────────────────
    df_rep = pd.read_excel(io.BytesIO(report_bytes), sheet_name='Offline Report', header=1)
    df_rep = df_rep.dropna(subset=['Sr No']).copy()
    for col in ['Total\nAttended','Actual Expenses','Expected Revenue',
                'Actual Revenue(W/O GST)\nAttendees','Total Revenue\n(W/O GST)\nAttendees',
                'Surplus or Deficit','Targeted\n','Total\nSeat\nBooked\n(in Seminar)',
                'Targeted to Attended (%)']:
        df_rep[col] = pd.to_numeric(df_rep[col], errors='coerce').fillna(0)
    results['report'] = df_rep

    return results

# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="dash-header">
  <div>
    <div class="dash-title">📊 Seminar Intelligence Dashboard</div>
    <div class="dash-sub">Conversion · Attendance · Revenue · Payment · {datetime.now().strftime("%d %B %Y")}</div>
  </div>
  <div class="live-pill">LIVE OVERVIEW</div>
</div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tab_upload, tab_overview, tab_seminar, tab_conversion, tab_revenue, tab_students = st.tabs([
    "📁 Upload Files",
    "🏠 Overview",
    "🎯 Seminar Analytics",
    "💰 Conversion Insights",
    "📈 Revenue & Finance",
    "👥 Student Payment Manager"
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
with tab_upload:
    st.markdown('<div class="tab-head">Upload Your 4 Data Files</div>', unsafe_allow_html=True)
    st.markdown('<div class="tab-sub">Upload all 4 files below. The dashboard will auto-process and populate all tabs.</div>', unsafe_allow_html=True)

    u1, u2 = st.columns(2)
    u3, u4 = st.columns(2)

    with u1:
        st.markdown("**📋 Conversion List** (`.xlsx`)")
        f_conv = st.file_uploader("Conversion List", type=["xlsx","xls"], key="f_conv", label_visibility="collapsed")
    with u2:
        st.markdown("**📝 Seminar Updated Sheet** (`.csv`)")
        f_sem  = st.file_uploader("Seminar CSV", type=["csv"], key="f_sem", label_visibility="collapsed")
    with u3:
        st.markdown("**📂 Offline Indepth Attendees** (`.xlsx`)")
        f_ind  = st.file_uploader("Indepth Attendees", type=["xlsx","xls"], key="f_ind", label_visibility="collapsed")
    with u4:
        st.markdown("**📊 Offline Seminar Report** (`.xlsx`)")
        f_rep  = st.file_uploader("Seminar Report", type=["xlsx","xls"], key="f_rep", label_visibility="collapsed")

    all_uploaded = all([f_conv, f_sem, f_ind, f_rep])

    if all_uploaded:
        with st.spinner("Processing all files…"):
            data = process_all(
                f_conv.read(), f_sem.read(), f_ind.read(), f_rep.read()
            )
        st.session_state.conv    = data['conv']
        st.session_state.seminar = data['seminar']
        st.session_state.indepth = data['indepth']
        st.session_state.report  = data['report']
        st.session_state.loaded  = True
        st.success("✅ All 4 files loaded! Navigate to any tab to explore your dashboard.")

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Conversion records", f"{len(data['conv']):,}")
        c2.metric("Seminar records",     f"{len(data['seminar']):,}")
        c3.metric("Attendee records",    f"{len(data['indepth']):,}")
        c4.metric("Seminar venues",      f"{len(data['report']):,}")
    else:
        missing = []
        if not f_conv: missing.append("Conversion List")
        if not f_sem:  missing.append("Seminar Updated CSV")
        if not f_ind:  missing.append("Offline Indepth Attendees")
        if not f_rep:  missing.append("Seminar Report")
        st.info(f"⏳ Waiting for: **{', '.join(missing)}**")

# ── Guard — show placeholder if not loaded ────────────────────────────────────
def not_loaded_msg():
    st.warning("⚠️ Please upload all 4 files in the **Upload Files** tab first.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
with tab_overview:
    if not st.session_state.loaded:
        not_loaded_msg()
    else:
        df_conv = st.session_state.conv
        df_sem  = st.session_state.seminar
        indepth = st.session_state.indepth
        df_rep  = st.session_state.report

        # Derived KPIs
        attended     = df_sem[df_sem['Is Attended ?'] == 'YES']
        converted_s  = attended[attended['Is Converted ?'].isin(['CONVERTED','YES'])]
        conv_rate    = len(converted_s)/len(attended)*100 if len(attended) else 0
        total_rev    = df_conv['total_amount'].sum()
        total_due    = df_conv['total_due'].sum()
        total_rcvd   = df_conv['payment_received'].sum()
        active_stu   = (indepth['status'] == 'Active').sum()
        total_venues = len(df_rep)
        total_exp    = df_rep['Actual Expenses'].sum()
        total_surplus= df_rep['Surplus or Deficit'].sum()
        roi_pct      = total_surplus/total_exp*100 if total_exp else 0
        avg_attend   = df_rep['Targeted to Attended (%)'].mean()*100

        st.markdown("##### Top-line KPIs")
        r1 = st.columns(4)
        r2 = st.columns(4)

        with r1[0]: st.markdown(kpi_card("Total Conversions",    f"{len(df_conv):,}",           "All-time records",                "neu"), unsafe_allow_html=True)
        with r1[1]: st.markdown(kpi_card("Seminar Attendees",    f"{len(attended):,}",           f"{len(df_sem):,} total registered","neu"), unsafe_allow_html=True)
        with r1[2]: st.markdown(kpi_card("Conversion Rate",      f"{conv_rate:.1f}%",            f"{len(converted_s):,} converted",  "up" if conv_rate>15 else "warn"), unsafe_allow_html=True)
        with r1[3]: st.markdown(kpi_card("Total Revenue",        fmt_cr(total_rev),              "Gross order value",               "up"), unsafe_allow_html=True)
        with r2[0]: st.markdown(kpi_card("Payment Received",     fmt_cr(total_rcvd),             f"Due: {fmt_cr(total_due)}",        "up"), unsafe_allow_html=True)
        with r2[1]: st.markdown(kpi_card("Active Students",      f"{active_stu:,}",              f"Offline batch enrolled",         "up"), unsafe_allow_html=True)
        with r2[2]: st.markdown(kpi_card("Venues Covered",       f"{total_venues}",              "Offline seminar locations",       "neu"), unsafe_allow_html=True)
        with r2[3]: st.markdown(kpi_card("Seminar ROI",          f"{roi_pct:.1f}%",              f"Surplus {fmt_cr(total_surplus)}", "up" if roi_pct>50 else "warn"), unsafe_allow_html=True)

        st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

        # Row: Monthly conversions + Conversion funnel
        oc1, oc2 = st.columns([3, 2])
        with oc1:
            st.markdown('<div class="section-head">Monthly conversion volume & revenue</div>', unsafe_allow_html=True)
            mo = df_conv.groupby('month').agg(count=('_id','count'), revenue=('total_amount','sum')).reset_index().tail(6)
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            fig.add_trace(go.Bar(x=mo['month'], y=mo['count'], name="Orders",
                                 marker_color=BLUE, marker_line_width=0, opacity=0.85), secondary_y=False)
            fig.add_trace(go.Scatter(x=mo['month'], y=mo['revenue'], name="Revenue",
                                     line=dict(color=GRN, width=2.5), mode="lines+markers",
                                     marker=dict(size=7, color=GRN)), secondary_y=True)
            fig.update_layout(**CHART_LAYOUT, height=250, bargap=0.3, legend=dict(orientation="h",x=0,y=1.15,font=dict(size=11)))
            fig.update_yaxes(gridcolor=GRID, secondary_y=False)
            fig.update_yaxes(gridcolor=GRID, tickformat=",.0f", secondary_y=True)
            fig.update_xaxes(showgrid=False)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

        with oc2:
            st.markdown('<div class="section-head">Seminar conversion funnel</div>', unsafe_allow_html=True)
            funnel_vals = [len(df_sem), len(attended), len(converted_s)]
            funnel_labs = ["Registered", "Attended", "Converted"]
            fig2 = go.Figure(go.Funnel(
                y=funnel_labs, x=funnel_vals,
                textinfo="value+percent initial",
                marker=dict(color=[BLUE, PURP, GRN]),
                connector=dict(line=dict(color=BG, width=3)),
            ))
            fig2.update_layout(**CHART_LAYOUT, height=250)
            st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar":False})

        # Row: Status breakdown + Top trainers
        oc3, oc4 = st.columns(2)
        with oc3:
            st.markdown('<div class="section-head">Order status breakdown</div>', unsafe_allow_html=True)
            status_counts = df_conv['status'].value_counts()
            fig3 = go.Figure(go.Pie(
                labels=status_counts.index, values=status_counts.values, hole=0.62,
                marker=dict(colors=[GRN, RED, AMB, GRAY], line=dict(color=BG, width=2)),
                textfont=dict(size=12),
                hovertemplate="%{label}: %{value}<extra></extra>",
            ))
            fig3.update_layout(**CHART_LAYOUT, height=230,
                legend=dict(orientation="h", x=0.5, xanchor="center", y=-0.15, font=dict(size=11)),
                annotations=[dict(text="Status", x=0.5, y=0.5, showarrow=False, font=dict(size=12, color=TEXT))])
            st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar":False})

        with oc4:
            st.markdown('<div class="section-head">Top 8 trainers by conversions</div>', unsafe_allow_html=True)
            top_tr = df_conv['trainer'].value_counts().dropna().head(8)
            top_tr_names = [str(x).split(' - ')[-1] if ' - ' in str(x) else str(x) for x in top_tr.index]
            fig4 = go.Figure(go.Bar(
                x=top_tr.values, y=top_tr_names, orientation='h',
                marker_color=PURP, marker_line_width=0,
                text=top_tr.values, textposition='outside', textfont=dict(size=11, color=TEXT),
            ))
            fig4.update_layout(**CHART_LAYOUT, height=230,
                xaxis=dict(showgrid=False, showticklabels=False), yaxis=dict(showgrid=False), bargap=0.35)
            st.plotly_chart(fig4, use_container_width=True, config={"displayModeBar":False})

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — SEMINAR ANALYTICS
# ══════════════════════════════════════════════════════════════════════════════
with tab_seminar:
    if not st.session_state.loaded:
        not_loaded_msg()
    else:
        df_sem  = st.session_state.seminar
        df_rep  = st.session_state.report
        attended = df_sem[df_sem['Is Attended ?'] == 'YES']
        converted_s = attended[attended['Is Converted ?'].isin(['CONVERTED','YES'])]

        st.markdown("##### Seminar Performance KPIs")
        sr1 = st.columns(4)
        with sr1[0]: st.markdown(kpi_card("Total Venues", f"{len(df_rep)}", "Offline seminar locations", "neu"), unsafe_allow_html=True)
        with sr1[1]: st.markdown(kpi_card("Total Attended", f"{int(df_rep['Total\nAttended'].sum()):,}", "Across all venues", "up"), unsafe_allow_html=True)
        with sr1[2]:
            avg_rate = df_rep['Targeted to Attended (%)'].mean()*100
            st.markdown(kpi_card("Avg Attendance Rate", f"{avg_rate:.1f}%", "Target vs actual", "up" if avg_rate>30 else "warn"), unsafe_allow_html=True)
        with sr1[3]:
            seats = int(df_rep['Total\nSeat\nBooked\n(in Seminar)'].sum())
            st.markdown(kpi_card("Seats Booked", f"{seats:,}", "Via seminar/webinar", "neu"), unsafe_allow_html=True)

        st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

        sc1, sc2 = st.columns(2)
        with sc1:
            st.markdown('<div class="section-head">Top locations by attendance</div>', unsafe_allow_html=True)
            top_loc = df_rep.nlargest(10, 'Total\nAttended')[['Location','Total\nAttended']].copy()
            fig = go.Figure(go.Bar(
                x=top_loc['Total\nAttended'], y=top_loc['Location'], orientation='h',
                marker_color=BLUE, marker_line_width=0,
                text=top_loc['Total\nAttended'].astype(int), textposition='outside', textfont=dict(size=11,color=TEXT),
            ))
            fig.update_layout(**CHART_LAYOUT, height=320,
                xaxis=dict(showgrid=False,showticklabels=False), yaxis=dict(showgrid=False), bargap=0.3)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

        with sc2:
            st.markdown('<div class="section-head">Morning vs Evening conversion split</div>', unsafe_allow_html=True)
            sess_conv = attended.copy()
            sess_conv['Converted'] = sess_conv['Is Converted ?'].isin(['CONVERTED','YES'])
            sess_grp = sess_conv.groupby('Session').agg(Attended=('Converted','count'), Converted=('Converted','sum')).reset_index()
            sess_grp = sess_grp[sess_grp['Session'].isin(['MORNING','EVENING'])]
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(name='Attended',  x=sess_grp['Session'], y=sess_grp['Attended'],
                                  marker_color=BLUE, opacity=0.7))
            fig2.add_trace(go.Bar(name='Converted', x=sess_grp['Session'], y=sess_grp['Converted'],
                                  marker_color=GRN))
            fig2.update_layout(**CHART_LAYOUT, height=320, barmode='group', bargap=0.35,
                legend=dict(orientation="h", x=0, y=1.12, font=dict(size=11)),
                xaxis=dict(showgrid=False), yaxis=dict(gridcolor=GRID))
            st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar":False})

        sc3, sc4 = st.columns(2)
        with sc3:
            st.markdown('<div class="section-head">Top locations by seminar conversion (CSV)</div>', unsafe_allow_html=True)
            loc_conv = converted_s['Place'].value_counts().head(10)
            fig3 = go.Figure(go.Bar(
                x=loc_conv.values, y=loc_conv.index, orientation='h',
                marker_color=GRN, marker_line_width=0,
                text=loc_conv.values, textposition='outside', textfont=dict(size=11, color=TEXT),
            ))
            fig3.update_layout(**CHART_LAYOUT, height=300,
                xaxis=dict(showgrid=False, showticklabels=False), yaxis=dict(showgrid=False), bargap=0.3)
            st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar":False})

        with sc4:
            st.markdown('<div class="section-head">Expense vs Revenue by location (top 10)</div>', unsafe_allow_html=True)
            top10 = df_rep.nlargest(10, 'Actual Revenue(W/O GST)\nAttendees')
            fig4 = go.Figure()
            fig4.add_trace(go.Bar(name='Expenses', x=top10['Location'], y=top10['Actual Expenses'],
                                  marker_color=RED, opacity=0.8))
            fig4.add_trace(go.Bar(name='Revenue',  x=top10['Location'], y=top10['Actual Revenue(W/O GST)\nAttendees'],
                                  marker_color=GRN, opacity=0.8))
            fig4.update_layout(**CHART_LAYOUT, height=300, barmode='group', bargap=0.25,
                legend=dict(orientation="h", x=0, y=1.12, font=dict(size=11)),
                xaxis=dict(showgrid=False, tickangle=-35), yaxis=dict(gridcolor=GRID))
            st.plotly_chart(fig4, use_container_width=True, config={"displayModeBar":False})

        st.markdown('<div class="section-head" style="margin-top:8px">Full seminar report</div>', unsafe_allow_html=True)
        st.dataframe(df_rep[['Location','Seminar Date','Total\nAttended','Actual Expenses',
                               'Expected Revenue','Actual Revenue(W/O GST)\nAttendees',
                               'Surplus or Deficit','Targeted to Attended (%)']].rename(columns={
            'Total\nAttended':'Attended','Actual Revenue(W/O GST)\nAttendees':'Actual Rev',
            'Targeted to Attended (%)':'Attend Rate'}),
            use_container_width=True, height=280)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — CONVERSION INSIGHTS
# ══════════════════════════════════════════════════════════════════════════════
with tab_conversion:
    if not st.session_state.loaded:
        not_loaded_msg()
    else:
        df_conv = st.session_state.conv
        df_sem  = st.session_state.seminar

        st.markdown("##### Conversion KPIs")
        attended    = df_sem[df_sem['Is Attended ?'] == 'YES']
        converted_s = attended[attended['Is Converted ?'].isin(['CONVERTED','YES'])]
        sem_conv_r  = len(converted_s)/len(attended)*100 if len(attended) else 0

        cr1 = st.columns(4)
        active_orders  = (df_conv['status'] == 'Active').sum()
        inactive_orders= (df_conv['status'] == 'Inactive').sum()
        closed_orders  = (df_conv['status'] == 'Closed').sum()
        with cr1[0]: st.markdown(kpi_card("Total Orders",    f"{len(df_conv):,}",    "All time",            "neu"), unsafe_allow_html=True)
        with cr1[1]: st.markdown(kpi_card("Active Orders",   f"{active_orders:,}",   f"Inactive: {inactive_orders:,}", "up"), unsafe_allow_html=True)
        with cr1[2]: st.markdown(kpi_card("Seminar Conv Rate",f"{sem_conv_r:.1f}%",  f"{len(converted_s):,} of {len(attended):,}", "up" if sem_conv_r>15 else "warn"), unsafe_allow_html=True)
        with cr1[3]: st.markdown(kpi_card("Avg Order Value", f"₹{df_conv[df_conv['total_amount']>0]['total_amount'].mean():,.0f}", "Non-zero orders", "up"), unsafe_allow_html=True)

        st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

        cc1, cc2 = st.columns(2)
        with cc1:
            st.markdown('<div class="section-head">Payment mode distribution</div>', unsafe_allow_html=True)
            pm = df_conv['payment_mode'].value_counts()
            pm.index = pm.index.map(lambda x: str(x).replace('mode','Mode '))
            fig = go.Figure(go.Pie(
                labels=pm.index, values=pm.values, hole=0.58,
                marker=dict(colors=[BLUE,PURP,GRN,AMB,RED,TEAL,GRAY], line=dict(color=BG,width=2)),
                textfont=dict(size=11),
            ))
            fig.update_layout(**CHART_LAYOUT, height=260,
                legend=dict(orientation="h",x=0.5,xanchor="center",y=-0.18,font=dict(size=11)),
                annotations=[dict(text="Mode",x=0.5,y=0.5,showarrow=False,font=dict(size=12,color=TEXT))])
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

        with cc2:
            st.markdown('<div class="section-head">Top 8 sales reps by conversions</div>', unsafe_allow_html=True)
            top_sr = df_conv['sales_rep_name'].value_counts().dropna().head(8)
            fig2 = go.Figure(go.Bar(
                x=top_sr.values, y=top_sr.index, orientation='h',
                marker_color=TEAL, marker_line_width=0,
                text=top_sr.values, textposition='outside', textfont=dict(size=11, color=TEXT),
            ))
            fig2.update_layout(**CHART_LAYOUT, height=260,
                xaxis=dict(showgrid=False,showticklabels=False), yaxis=dict(showgrid=False), bargap=0.35)
            st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar":False})

        cc3, cc4 = st.columns(2)
        with cc3:
            st.markdown('<div class="section-head">Top 8 courses by order volume</div>', unsafe_allow_html=True)
            top_courses = df_conv['service_name'].value_counts().head(8)
            short_names = [s[:35]+'…' if len(s)>35 else s for s in top_courses.index]
            fig3 = go.Figure(go.Bar(
                x=top_courses.values, y=short_names, orientation='h',
                marker_color=AMB, marker_line_width=0,
                text=top_courses.values, textposition='outside', textfont=dict(size=11, color=TEXT),
            ))
            fig3.update_layout(**CHART_LAYOUT, height=280,
                xaxis=dict(showgrid=False,showticklabels=False), yaxis=dict(showgrid=False), bargap=0.3)
            st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar":False})

        with cc4:
            st.markdown('<div class="section-head">Seminar amount paid distribution (converted)</div>', unsafe_allow_html=True)
            amt_nonzero = converted_s[converted_s['Amount Paid']>0]['Amount Paid']
            fig4 = go.Figure(go.Histogram(x=amt_nonzero, nbinsx=20,
                marker_color=PURP, marker_line_width=0, opacity=0.85))
            fig4.update_layout(**CHART_LAYOUT, height=280,
                xaxis=dict(title="Amount Paid (₹)", showgrid=False),
                yaxis=dict(gridcolor=GRID, title="Count"))
            st.plotly_chart(fig4, use_container_width=True, config={"displayModeBar":False})

        st.markdown('<div class="section-head" style="margin-top:8px">Recent conversions</div>', unsafe_allow_html=True)
        disp = df_conv[['order_date','student_name','service_name','total_amount','total_due','status','sales_rep_name']]\
                .sort_values('order_date', ascending=False).head(50)
        disp['total_amount'] = disp['total_amount'].apply(lambda x: f"₹{x:,.0f}")
        disp['total_due']    = disp['total_due'].apply(lambda x: f"₹{x:,.0f}")
        st.dataframe(disp, use_container_width=True, height=300, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — REVENUE & FINANCE
# ══════════════════════════════════════════════════════════════════════════════
with tab_revenue:
    if not st.session_state.loaded:
        not_loaded_msg()
    else:
        df_conv = st.session_state.conv
        df_rep  = st.session_state.report
        indepth = st.session_state.indepth

        total_rev    = df_conv['total_amount'].sum()
        total_rcvd   = df_conv['payment_received'].sum()
        total_due    = df_conv['total_due'].sum()
        total_exp    = df_rep['Actual Expenses'].sum()
        total_surplus= df_rep['Surplus or Deficit'].sum()
        offline_rev  = indepth[~indepth['location'].str.upper().str.contains('CON')]['payment_received'].sum()

        st.markdown("##### Revenue KPIs")
        rv1 = st.columns(4)
        with rv1[0]: st.markdown(kpi_card("Gross Revenue",       fmt_cr(total_rev),           "Total order value",           "up"), unsafe_allow_html=True)
        with rv1[1]: st.markdown(kpi_card("Payment Collected",   fmt_cr(total_rcvd),          f"Due: {fmt_cr(total_due)}",   "up"), unsafe_allow_html=True)
        with rv1[2]: st.markdown(kpi_card("Offline Revenue",     fmt_cr(offline_rev),         "From indepth attendees",      "up"), unsafe_allow_html=True)
        with rv1[3]: st.markdown(kpi_card("Seminar Surplus",     fmt_cr(total_surplus),       f"Expenses: {fmt_cr(total_exp)}", "up" if total_surplus>0 else "down"), unsafe_allow_html=True)

        st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

        rc1, rc2 = st.columns(2)
        with rc1:
            st.markdown('<div class="section-head">Monthly revenue trend</div>', unsafe_allow_html=True)
            mo_rev = df_conv.groupby('month')['total_amount'].sum().reset_index().tail(6)
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=mo_rev['month'], y=mo_rev['total_amount'],
                mode='lines+markers', line=dict(color=GRN, width=2.5, shape='spline'),
                marker=dict(size=7, color=GRN, line=dict(color=BG,width=2)),
                fill='tozeroy', fillcolor='rgba(52,211,153,0.08)'))
            fig.update_layout(**CHART_LAYOUT, height=250,
                xaxis=dict(showgrid=False), yaxis=dict(gridcolor=GRID, tickformat=",.0f"))
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

        with rc2:
            st.markdown('<div class="section-head">Payment collected vs due (monthly)</div>', unsafe_allow_html=True)
            mo_pay = df_conv.groupby('month').agg(received=('payment_received','sum'), due=('total_due','sum')).reset_index().tail(6)
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(name='Received', x=mo_pay['month'], y=mo_pay['received'], marker_color=GRN, opacity=0.85))
            fig2.add_trace(go.Bar(name='Due',      x=mo_pay['month'], y=mo_pay['due'],      marker_color=RED, opacity=0.85))
            fig2.update_layout(**CHART_LAYOUT, height=250, barmode='group', bargap=0.3,
                legend=dict(orientation="h",x=0,y=1.12,font=dict(size=11)),
                xaxis=dict(showgrid=False), yaxis=dict(gridcolor=GRID))
            st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar":False})

        rc3, rc4 = st.columns(2)
        with rc3:
            st.markdown('<div class="section-head">Offline location revenue (top 10 — indepth)</div>', unsafe_allow_html=True)
            off_only = indepth[~indepth['location'].str.upper().str.contains('CON')]
            loc_rev  = off_only.groupby('location')['payment_received'].sum().nlargest(10)
            fig3 = go.Figure(go.Bar(
                x=loc_rev.values, y=loc_rev.index, orientation='h',
                marker_color=TEAL, marker_line_width=0,
                text=[fmt_cr(v) for v in loc_rev.values], textposition='outside', textfont=dict(size=11,color=TEXT),
            ))
            fig3.update_layout(**CHART_LAYOUT, height=300,
                xaxis=dict(showgrid=False,showticklabels=False), yaxis=dict(showgrid=False), bargap=0.3)
            st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar":False})

        with rc4:
            st.markdown('<div class="section-head">Seminar surplus/deficit by location</div>', unsafe_allow_html=True)
            surp = df_rep[['Location','Surplus or Deficit']].copy().sort_values('Surplus or Deficit', ascending=True)
            colors_s = [GRN if v >= 0 else RED for v in surp['Surplus or Deficit']]
            fig4 = go.Figure(go.Bar(
                x=surp['Surplus or Deficit'], y=surp['Location'], orientation='h',
                marker_color=colors_s, marker_line_width=0,
            ))
            fig4.update_layout(**CHART_LAYOUT, height=300,
                xaxis=dict(gridcolor=GRID, zeroline=True, zerolinecolor=GRAY),
                yaxis=dict(showgrid=False), bargap=0.25)
            st.plotly_chart(fig4, use_container_width=True, config={"displayModeBar":False})

# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — STUDENT PAYMENT MANAGER
# ══════════════════════════════════════════════════════════════════════════════
with tab_students:
    if not st.session_state.loaded:
        not_loaded_msg()
    else:
        indepth = st.session_state.indepth.copy()
        off_only = indepth[~indepth['location'].str.upper().str.contains('CON')].copy().reset_index(drop=True)

        st.markdown('<div class="tab-head">Offline Attendee — Payment Manager</div>', unsafe_allow_html=True)
        st.markdown('<div class="tab-sub">Only offline seminar attendees shown. Search, filter and update payment status.</div>', unsafe_allow_html=True)

        total_s  = len(off_only)
        active_s = (off_only['status'] == 'Active').sum()
        inactive_s= (off_only['status'] == 'Inactive').sum()
        due_total = off_only['total_due'].sum()
        rcvd_total= off_only['payment_received'].sum()

        pm1 = st.columns(4)
        with pm1[0]: st.markdown(kpi_card("Offline Students",   f"{total_s:,}",           "Excl. conversion sheets",     "neu"), unsafe_allow_html=True)
        with pm1[1]: st.markdown(kpi_card("Active",             f"{active_s:,}",           f"Inactive: {inactive_s:,}",   "up"), unsafe_allow_html=True)
        with pm1[2]: st.markdown(kpi_card("Total Received",     fmt_cr(rcvd_total),        "Payment collected",           "up"), unsafe_allow_html=True)
        with pm1[3]: st.markdown(kpi_card("Total Due",          fmt_cr(due_total),         "Pending recovery",            "down" if due_total>0 else "up"), unsafe_allow_html=True)

        st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

        # Filters
        f1, f2, f3 = st.columns([3,2,2])
        search_s  = f1.text_input("🔍 Search", placeholder="Name, phone, email, location…")
        filt_stat = f2.selectbox("Filter by status", ["All","Active","Inactive","Closed","Refunded"])
        filt_loc  = f3.selectbox("Filter by location", ["All"] + sorted(off_only['location'].unique().tolist()))

        df_view = off_only.copy()
        if search_s:
            mask = pd.Series([False]*len(df_view))
            for c in df_view.columns:
                mask = mask | df_view[c].astype(str).str.lower().str.contains(search_s.lower(), na=False)
            df_view = df_view[mask]
        if filt_stat != "All":
            df_view = df_view[df_view['status'] == filt_stat]
        if filt_loc != "All":
            df_view = df_view[df_view['location'] == filt_loc]

        st.markdown(f"Showing **{len(df_view):,}** of **{total_s:,}** records")

        edited = st.data_editor(
            df_view[['student_name','student_invid','phone','email','location','service_name',
                     'total_amount','payment_received','total_due','status']],
            use_container_width=True, height=400,
            column_config={
                "status": st.column_config.SelectboxColumn("Status",
                    options=["Active","Inactive","Closed","Refunded","Relationship Buildup"], required=True),
                "total_amount":      st.column_config.NumberColumn("Total Amt (₹)", format="₹%.0f"),
                "payment_received":  st.column_config.NumberColumn("Received (₹)",  format="₹%.0f"),
                "total_due":         st.column_config.NumberColumn("Due (₹)",        format="₹%.0f"),
            },
            key="student_editor", num_rows="fixed",
        )

        st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
        st.markdown("**Export**")
        def to_xl(df):
            buf = io.BytesIO(); df.to_excel(buf, index=False); return buf.getvalue()

        ts = datetime.now().strftime("%Y%m%d_%H%M")
        ex1,ex2,ex3 = st.columns(3)
        ex1.download_button("⬇️ Export All Offline Students", to_xl(off_only),
            file_name=f"offline_students_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
        ex2.download_button("⬇️ Export Due > 0",
            to_xl(off_only[off_only['total_due']>0]),
            file_name=f"students_due_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
        ex3.download_button("⬇️ Export Active Only",
            to_xl(off_only[off_only['status']=='Active']),
            file_name=f"students_active_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="border-top:1px solid rgba(255,255,255,0.06);margin-top:2rem;padding-top:1rem;
     display:flex;justify-content:space-between">
  <span style="font-size:12px;color:#374151;font-family:'DM Mono',monospace">seminar intelligence dashboard · {datetime.now().strftime("%B %Y").lower()}</span>
  <span style="font-size:12px;color:#374151">upload fresh files anytime to refresh</span>
</div>""", unsafe_allow_html=True)
