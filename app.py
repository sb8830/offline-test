import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
from datetime import datetime

st.set_page_config(
    page_title="Automation KPI Dashboard",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.main { background-color: #0d0f14; }
.block-container { padding: 1.5rem 2rem 2rem 2rem; max-width: 1400px; }

.dash-header { display:flex; align-items:flex-start; justify-content:space-between; margin-bottom:1.5rem; padding-bottom:1rem; border-bottom:1px solid rgba(255,255,255,0.07); }
.dash-title  { font-size:22px; font-weight:600; color:#f0f2f7; letter-spacing:-0.3px; }
.dash-sub    { font-size:13px; color:#6b7280; margin-top:3px; font-family:'DM Mono',monospace; }
.live-pill   { background:rgba(29,158,117,0.15); color:#34d399; border:1px solid rgba(52,211,153,0.25); padding:4px 12px; border-radius:99px; font-size:11px; font-weight:500; letter-spacing:.05em; }

.metric-card { background:#161922; border:1px solid rgba(255,255,255,0.07); border-radius:12px; padding:1.1rem 1.2rem; position:relative; overflow:hidden; margin-bottom:14px; }
.metric-card::before { content:''; position:absolute; top:0; left:0; right:0; height:2px; background:linear-gradient(90deg,#4f8ef7,#7c5af7); opacity:.6; }
.m-label  { font-size:11px; color:#6b7280; text-transform:uppercase; letter-spacing:.06em; margin-bottom:8px; }
.m-value  { font-size:28px; font-weight:600; color:#f0f2f7; line-height:1; font-family:'DM Mono',monospace; }
.m-badge  { display:inline-block; font-size:11px; padding:3px 8px; border-radius:99px; margin-top:8px; font-weight:500; }
.badge-up  { background:rgba(29,158,117,0.15);  color:#34d399; }
.badge-down{ background:rgba(226,75,74,0.15);   color:#f87171; }
.badge-neu { background:rgba(107,114,128,0.15); color:#9ca3af; }
.section-head { font-size:11px; font-weight:500; color:#6b7280; text-transform:uppercase; letter-spacing:.07em; margin-bottom:14px; }
.section-divider { border:none; border-top:1px solid rgba(255,255,255,0.07); margin:1.5rem 0; }
.tab-head { font-size:16px; font-weight:500; color:#f0f2f7; margin-bottom:4px; }
.tab-sub  { font-size:13px; color:#6b7280; margin-bottom:1rem; }
.stat-inline { display:inline-block; background:#1e2330; border:1px solid rgba(255,255,255,0.08); border-radius:8px; padding:6px 14px; font-size:13px; margin-right:8px; margin-bottom:8px; color:#d1d5db; }
.stat-inline span { font-weight:600; color:#f0f2f7; font-family:'DM Mono',monospace; }
</style>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
for key in ["student_df", "col_map"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="dash-header">
  <div>
    <div class="dash-title">⚡ Automation KPI Dashboard</div>
    <div class="dash-sub">Seminar attendance &amp; payment management · {datetime.now().strftime("%B %Y")}</div>
  </div>
  <div class="live-pill">LIVE OVERVIEW</div>
</div>
""", unsafe_allow_html=True)

# ── Chart palette ─────────────────────────────────────────────────────────────
BG   = "#161922"
GRID = "rgba(255,255,255,0.05)"
TEXT = "#9ca3af"
BLUE = "#4f8ef7"
PURP = "#7c5af7"
GRN  = "#34d399"
AMB  = "#fbbf24"
RED  = "#f87171"
GRAY = "#6b7280"

# ═══════════════════════════════════════════════════════════════════════════════
tab1, tab2, tab3 = st.tabs(["📁  Upload & Filter", "💳  Payment Manager", "📊  KPI Dashboard"])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — UPLOAD & FILTER
# ═══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="tab-head">Upload Student Data</div>', unsafe_allow_html=True)
    st.markdown('<div class="tab-sub">Upload your Excel or CSV file. Map your columns, then filter only offline seminar attendees.</div>', unsafe_allow_html=True)

    uploaded = st.file_uploader("Drop your file here", type=["xlsx", "xls", "csv"],
                                 help="Supports .xlsx, .xls, .csv")

    if uploaded:
        try:
            df_raw = pd.read_csv(uploaded) if uploaded.name.endswith(".csv") else pd.read_excel(uploaded)
            st.success(f"✅ Loaded **{uploaded.name}** — {len(df_raw):,} rows × {len(df_raw.columns)} columns")

            st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
            st.markdown("**Map your columns:**")

            opts = ["— select —"] + list(df_raw.columns)
            c1, c2, c3, c4, c5 = st.columns(5)
            col_name         = c1.selectbox("Student Name",      opts, key="cn")
            col_contact      = c2.selectbox("Phone / Email",     opts, key="cc")
            col_seminar_type = c3.selectbox("Seminar Type",      opts, key="cs")
            col_attendance   = c4.selectbox("Attendance Status", opts, key="ca")
            col_payment      = c5.selectbox("Payment Status",    opts, key="cp")

            st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
            kc1, kc2, kc3 = st.columns(3)
            offline_kw  = kc1.text_input("Keyword for 'Offline'",  value="offline",  help="Case-insensitive")
            attended_kw = kc2.text_input("Keyword for 'Attended'", value="attended", help="Case-insensitive")
            kc3.markdown("<br>", unsafe_allow_html=True)
            apply_btn = kc3.button("🔍  Filter Offline Attendees", use_container_width=True)

            if apply_btn:
                missing = [n for n, v in [("Seminar Type", col_seminar_type), ("Attendance Status", col_attendance)] if v == "— select —"]
                if missing:
                    st.error(f"Please map: {', '.join(missing)}")
                else:
                    df = df_raw.copy()
                    mask = (
                        df[col_seminar_type].astype(str).str.lower().str.contains(offline_kw.lower(), na=False) &
                        df[col_attendance].astype(str).str.lower().str.contains(attended_kw.lower(), na=False)
                    )
                    df_filtered = df[mask].copy().reset_index(drop=True)

                    if col_payment == "— select —" or col_payment not in df_filtered.columns:
                        df_filtered["Payment Status"] = "Unpaid"
                        pay_final = "Payment Status"
                    else:
                        pay_final = col_payment

                    st.session_state.student_df = df_filtered
                    st.session_state.col_map = {
                        "name":    col_name    if col_name    != "— select —" else None,
                        "contact": col_contact if col_contact != "— select —" else None,
                        "seminar": col_seminar_type,
                        "attend":  col_attendance,
                        "payment": pay_final,
                    }

                    total_filt  = len(df_filtered)
                    paid_count  = df_filtered[pay_final].astype(str).str.lower().str.contains("paid", na=False).sum()
                    unpaid_count= total_filt - paid_count

                    st.markdown(f"""
                    <div style="margin:1rem 0">
                      <span class="stat-inline">Total records <span>{len(df_raw):,}</span></span>
                      <span class="stat-inline">Offline attendees <span>{total_filt:,}</span></span>
                      <span class="stat-inline" style="border-color:rgba(52,211,153,0.3)">Paid <span style="color:#34d399">{paid_count:,}</span></span>
                      <span class="stat-inline" style="border-color:rgba(248,113,113,0.3)">Unpaid <span style="color:#f87171">{unpaid_count:,}</span></span>
                    </div>
                    """, unsafe_allow_html=True)

                    st.dataframe(df_filtered, use_container_width=True, height=380)
                    st.info("✅ Filtered data saved. Go to **Payment Manager** tab to update statuses.")

        except Exception as e:
            st.error(f"Could not read file: {e}")

    else:
        st.markdown("""
        <div style="background:#161922;border:1.5px dashed rgba(79,142,247,0.35);border-radius:14px;padding:2rem;text-align:center;margin-bottom:1.5rem">
          <div style="font-size:16px;font-weight:500;color:#f0f2f7;margin-bottom:6px">📂 No file uploaded yet</div>
          <div style="font-size:13px;color:#6b7280">Use the uploader above · supports .xlsx, .xls, .csv</div>
        </div>
        """, unsafe_allow_html=True)

        sample = pd.DataFrame({
            "Student Name":     ["Aarav Shah","Priya Nair","Rohit Das","Sneha Iyer","Karan Mehta"],
            "Phone":            ["9800000001","9800000002","9800000003","9800000004","9800000005"],
            "Email":            ["aarav@x.com","priya@x.com","rohit@x.com","sneha@x.com","karan@x.com"],
            "Seminar Type":     ["Offline","Online","Offline","Offline","Online"],
            "Attendance Status":["Attended","Attended","Absent","Attended","Attended"],
            "Payment Status":   ["Unpaid","Paid","Unpaid","Unpaid","Paid"],
            "Course":           ["Python","ML","Python","Data Science","ML"],
        })
        buf = io.BytesIO()
        sample.to_excel(buf, index=False)
        st.download_button("⬇️  Download sample template", buf.getvalue(),
                           file_name="sample_student_data.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — PAYMENT MANAGER
# ═══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="tab-head">Payment Status Manager</div>', unsafe_allow_html=True)
    st.markdown('<div class="tab-sub">Update payment status for offline seminar attendees. Export updated lists instantly.</div>', unsafe_allow_html=True)

    if st.session_state.student_df is None:
        st.warning("⚠️ No data loaded yet. Please upload and filter in the **Upload & Filter** tab first.")
    else:
        df      = st.session_state.student_df.copy()
        pay_col = st.session_state.col_map["payment"]
        pay_options = ["Paid", "Unpaid", "Pending", "Waived", "Partial"]

        # ── Controls ───────────────────────────────────────────────────────
        sc1, sc2, sc3 = st.columns([3, 2, 2])
        search     = sc1.text_input("🔍 Search student", placeholder="Name, phone, email…")
        filter_pay = sc2.selectbox("Filter by payment", ["All"] + pay_options)
        sc3.markdown("<br>", unsafe_allow_html=True)
        bc1, bc2 = sc3.columns(2)
        if bc1.button("✅ All Paid",   use_container_width=True):
            st.session_state.student_df[pay_col] = "Paid"
            df = st.session_state.student_df.copy()
            st.success("All records marked as Paid.")
        if bc2.button("❌ All Unpaid", use_container_width=True):
            st.session_state.student_df[pay_col] = "Unpaid"
            df = st.session_state.student_df.copy()
            st.success("All records marked as Unpaid.")

        # ── View filter ────────────────────────────────────────────────────
        df_view = df.copy()
        if search:
            mask = pd.Series([False] * len(df_view))
            for c in df_view.columns:
                mask = mask | df_view[c].astype(str).str.lower().str.contains(search.lower(), na=False)
            df_view = df_view[mask]
        if filter_pay != "All":
            df_view = df_view[df_view[pay_col].astype(str).str.lower().str.contains(filter_pay.lower(), na=False)]

        st.markdown(f"Showing **{len(df_view):,}** of **{len(df):,}** records")

        edited = st.data_editor(
            df_view,
            use_container_width=True,
            height=420,
            column_config={
                pay_col: st.column_config.SelectboxColumn(
                    label="💳 Payment Status",
                    options=pay_options,
                    required=True,
                ),
            },
            key="payment_editor",
            num_rows="fixed",
        )

        if edited is not None and not edited.equals(df_view):
            st.session_state.student_df.update(edited)
            st.success("✅ Changes saved.")

        # ── Summary ────────────────────────────────────────────────────────
        st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
        master     = st.session_state.student_df
        pay_counts = master[pay_col].value_counts()
        total      = len(master)

        mc = st.columns(len(pay_options) + 1)
        mc[0].metric("Total Students", f"{total:,}")
        for i, opt in enumerate(pay_options):
            cnt = pay_counts.get(opt, 0)
            mc[i+1].metric(opt, f"{cnt:,}", f"{cnt/total*100:.1f}%" if total else "0%")

        # Donut chart
        fig_pay = go.Figure(go.Pie(
            labels=list(pay_counts.index), values=list(pay_counts.values),
            hole=0.62,
            marker=dict(colors=[GRN, RED, AMB, "#818cf8", "#60a5fa"], line=dict(color=BG, width=2)),
            textfont=dict(family="DM Sans", size=12),
            hovertemplate="%{label}: %{value} students<extra></extra>",
        ))
        fig_pay.update_layout(
            paper_bgcolor=BG, plot_bgcolor=BG, height=220,
            margin=dict(l=0,r=0,t=10,b=0),
            font=dict(family="DM Sans", color=TEXT),
            legend=dict(orientation="h", x=0.5, xanchor="center", y=-0.1,
                        font=dict(size=12, color=TEXT), bgcolor="rgba(0,0,0,0)"),
            annotations=[dict(text="Payment<br>split", x=0.5, y=0.5, showarrow=False,
                              font=dict(size=12, color=TEXT, family="DM Sans"))],
        )
        st.plotly_chart(fig_pay, use_container_width=True, config={"displayModeBar": False})

        # ── Export ─────────────────────────────────────────────────────────
        st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
        st.markdown("**Export updated data**")

        def to_excel_bytes(dataframe):
            buf = io.BytesIO()
            dataframe.to_excel(buf, index=False)
            return buf.getvalue()

        ts = datetime.now().strftime("%Y%m%d_%H%M")
        ec1, ec2, ec3 = st.columns(3)

        ec1.download_button("⬇️ Export all (Excel)",
            to_excel_bytes(st.session_state.student_df),
            file_name=f"students_all_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

        df_paid = st.session_state.student_df[
            st.session_state.student_df[pay_col].astype(str).str.lower() == "paid"]
        ec2.download_button("⬇️ Export Paid only",
            to_excel_bytes(df_paid),
            file_name=f"students_paid_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

        df_unpaid = st.session_state.student_df[
            st.session_state.student_df[pay_col].astype(str).str.lower().isin(["unpaid","pending"])]
        ec3.download_button("⬇️ Export Unpaid / Pending",
            to_excel_bytes(df_unpaid),
            file_name=f"students_unpaid_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — KPI DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="tab-head">KPI Dashboard</div>', unsafe_allow_html=True)

    # Live metrics from uploaded data
    if st.session_state.student_df is not None:
        df_s    = st.session_state.student_df
        pay_col = st.session_state.col_map["payment"]
        total_s = len(df_s)
        paid_s  = (df_s[pay_col].astype(str).str.lower() == "paid").sum()
        unpaid_s= total_s - paid_s
        conv_r  = f"{paid_s/total_s*100:.1f}%" if total_s else "0%"

        st.markdown("##### Live metrics from your data")
        lc = st.columns(4)
        lc[0].metric("Offline Attendees",  f"{total_s:,}")
        lc[1].metric("Paid",               f"{paid_s:,}",   f"{paid_s/total_s*100:.1f}%" if total_s else "")
        lc[2].metric("Unpaid",             f"{unpaid_s:,}", f"-{unpaid_s/total_s*100:.1f}%" if total_s else "")
        lc[3].metric("Conversion Rate",    conv_r)
        st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # Static automation KPIs
    st.markdown("##### Automation performance KPIs")
    kpis = [
        ("Tasks Automated",    "14,820",  "+18% MoM",          "up"),
        ("Success Rate",       "97.3%",   "+1.2% vs last mo",  "up"),
        ("Avg Execution Time", "1.4s",    "-0.3s faster",      "up"),
        ("Hours Saved",        "3,210 h", "+24% YoY",          "up"),
        ("Error Rate",         "2.7%",    "Down from 4.1%",    "up"),
        ("Cost per Task",      "$0.012",  "-8% MoM",           "up"),
        ("Active Workflows",   "84",      "+6 this month",     "neu"),
        ("ROI",                "312%",    "vs 260% last qtr",  "up"),
    ]
    kcols = st.columns(4)
    for i, (label, value, badge, kind) in enumerate(kpis):
        bc = {"up":"badge-up","down":"badge-down","neu":"badge-neu"}[kind]
        with kcols[i % 4]:
            st.markdown(f"""
            <div class="metric-card">
              <div class="m-label">{label}</div>
              <div class="m-value">{value}</div>
              <div><span class="m-badge {bc}">{badge}</span></div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<div style='margin-bottom:20px'></div>", unsafe_allow_html=True)

    months = ["Oct","Nov","Dec","Jan","Feb","Mar"]

    dc1, dc2 = st.columns([3, 2])
    with dc1:
        st.markdown('<div class="section-head">Monthly task volume</div>', unsafe_allow_html=True)
        fig = go.Figure(go.Bar(
            x=months, y=[8200,9100,10400,11200,12600,14820],
            marker_color=[BLUE]*5+[PURP], marker_line_width=0,
        ))
        fig.update_layout(paper_bgcolor=BG, plot_bgcolor=BG, height=230,
            margin=dict(l=0,r=0,t=10,b=0), font=dict(family="DM Sans",color=TEXT,size=12),
            xaxis=dict(showgrid=False), yaxis=dict(gridcolor=GRID,tickformat=",.0f"), bargap=0.3)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

    with dc2:
        st.markdown('<div class="section-head">Failure reasons</div>', unsafe_allow_html=True)
        fig2 = go.Figure(go.Pie(
            labels=["Timeout","Auth error","Bad input","Other"],
            values=[38,27,19,16], hole=0.65,
            marker=dict(colors=[RED,AMB,BLUE,GRAY], line=dict(color=BG, width=2)),
        ))
        fig2.update_layout(paper_bgcolor=BG, plot_bgcolor=BG, height=230,
            margin=dict(l=0,r=0,t=10,b=0), font=dict(family="DM Sans",color=TEXT),
            legend=dict(orientation="v",x=1.02,y=0.5,font=dict(size=11,color=TEXT),bgcolor="rgba(0,0,0,0)"),
            annotations=[dict(text="Errors",x=0.5,y=0.5,showarrow=False,font=dict(size=12,color=TEXT))])
        st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar":False})

    dc3, dc4 = st.columns(2)
    with dc3:
        st.markdown('<div class="section-head">Workflow health by category</div>', unsafe_allow_html=True)
        cats   = ["Data sync","Notifications","File processing","API integrations","Report gen","ML pipelines"]
        scores = [94,98,91,88,96,79]
        colors = [GRN if s>=90 else (AMB if s>=85 else RED) for s in scores]
        fig3 = go.Figure(go.Bar(x=scores, y=cats, orientation="h",
            marker_color=colors, marker_line_width=0,
            text=[f"{s}%" for s in scores], textposition="outside",
            textfont=dict(size=11, color=TEXT)))
        fig3.update_layout(paper_bgcolor=BG, plot_bgcolor=BG, height=260,
            margin=dict(l=0,r=50,t=10,b=0), font=dict(family="DM Sans",color=TEXT,size=12),
            xaxis=dict(range=[60,108],showgrid=False,showticklabels=False),
            yaxis=dict(showgrid=False), bargap=0.35)
        st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar":False})

    with dc4:
        st.markdown('<div class="section-head">Avg execution time trend (s)</div>', unsafe_allow_html=True)
        fig5 = go.Figure()
        fig5.add_trace(go.Scatter(x=months, y=[1.9,1.8,1.7,1.6,1.5,1.4],
            mode="lines+markers",
            line=dict(color=PURP, width=2.5, shape="spline"),
            marker=dict(size=7, color=PURP, line=dict(color=BG, width=2)),
            fill="tozeroy", fillcolor="rgba(124,90,247,0.1)"))
        fig5.update_layout(paper_bgcolor=BG, plot_bgcolor=BG, height=260,
            margin=dict(l=0,r=0,t=10,b=0), font=dict(family="DM Sans",color=TEXT,size=12),
            xaxis=dict(showgrid=False),
            yaxis=dict(gridcolor=GRID, range=[1.0,2.2], ticksuffix="s"), showlegend=False)
        st.plotly_chart(fig5, use_container_width=True, config={"displayModeBar":False})

    st.markdown('<div class="section-head" style="margin-top:8px">Top workflows by volume</div>', unsafe_allow_html=True)
    df_wf = pd.DataFrame({
        "Workflow":     ["Invoice processor","CRM data sync","Email classifier","Report scheduler","ML batch pipeline","Legacy API bridge"],
        "Runs":         [3210,2840,2100,1760,980,640],
        "Success Rate": ["99.1%","97.4%","95.6%","98.2%","79.3%","71.0%"],
        "Status":       ["🟢 Healthy","🟢 Healthy","🟢 Healthy","🟢 Healthy","🟡 Warning","🔴 Critical"],
    })
    st.dataframe(df_wf, use_container_width=True, hide_index=True, height=250)

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="border-top:1px solid rgba(255,255,255,0.06);margin-top:2rem;padding-top:1rem;
     display:flex;justify-content:space-between;align-items:center">
  <span style="font-size:12px;color:#374151;font-family:'DM Mono',monospace">
    automation-kpi-dashboard · {datetime.now().strftime("%B %Y").lower()}
  </span>
  <span style="font-size:12px;color:#374151">data updates on interaction</span>
</div>
""", unsafe_allow_html=True)
