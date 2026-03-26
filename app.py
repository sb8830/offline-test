import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

st.set_page_config(
    page_title="Automation KPI Dashboard",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

.main { background-color: #0d0f14; }
.block-container { padding: 1.5rem 2rem 2rem 2rem; max-width: 1400px; }

.dash-header {
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    margin-bottom: 1.5rem;
    padding-bottom: 1rem;
    border-bottom: 1px solid rgba(255,255,255,0.07);
}
.dash-title {
    font-size: 22px;
    font-weight: 600;
    color: #f0f2f7;
    letter-spacing: -0.3px;
}
.dash-sub {
    font-size: 13px;
    color: #6b7280;
    margin-top: 3px;
    font-family: 'DM Mono', monospace;
}
.live-pill {
    background: rgba(29,158,117,0.15);
    color: #34d399;
    border: 1px solid rgba(52,211,153,0.25);
    padding: 4px 12px;
    border-radius: 99px;
    font-size: 11px;
    font-weight: 500;
    letter-spacing: 0.05em;
}

.metric-card {
    background: #161922;
    border: 1px solid rgba(255,255,255,0.07);
    border-radius: 12px;
    padding: 1.1rem 1.2rem;
    position: relative;
    overflow: hidden;
}
.metric-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, #4f8ef7, #7c5af7);
    opacity: 0.6;
}
.m-label {
    font-size: 11px;
    color: #6b7280;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin-bottom: 8px;
}
.m-value {
    font-size: 28px;
    font-weight: 600;
    color: #f0f2f7;
    line-height: 1;
    font-family: 'DM Mono', monospace;
}
.m-badge {
    display: inline-block;
    font-size: 11px;
    padding: 3px 8px;
    border-radius: 99px;
    margin-top: 8px;
    font-weight: 500;
}
.badge-up { background: rgba(29,158,117,0.15); color: #34d399; }
.badge-down { background: rgba(226,75,74,0.15); color: #f87171; }
.badge-neu { background: rgba(107,114,128,0.15); color: #9ca3af; }

.section-head {
    font-size: 11px;
    font-weight: 500;
    color: #6b7280;
    text-transform: uppercase;
    letter-spacing: 0.07em;
    margin-bottom: 14px;
}

div[data-testid="stMetricValue"] { font-family: 'DM Mono', monospace; }
</style>
""", unsafe_allow_html=True)

# ── Header ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="dash-header">
    <div>
        <div class="dash-title">⚡ Automation KPI Dashboard</div>
        <div class="dash-sub">Deep insights — March 2026</div>
    </div>
    <div class="live-pill">LIVE OVERVIEW</div>
</div>
""", unsafe_allow_html=True)

# ── KPI Data ─────────────────────────────────────────────────────────────────
kpis = [
    ("Tasks Automated",     "14,820",  "+18% MoM",  "up"),
    ("Success Rate",        "97.3%",   "+1.2% vs last mo", "up"),
    ("Avg Execution Time",  "1.4s",    "-0.3s faster", "up"),
    ("Hours Saved",         "3,210 h", "+24% YoY",  "up"),
    ("Error Rate",          "2.7%",    "Down from 4.1%", "up"),
    ("Cost per Task",       "$0.012",  "-8% MoM",   "up"),
    ("Active Workflows",    "84",      "+6 new this mo", "neu"),
    ("ROI",                 "312%",    "vs 260% last qtr", "up"),
]

cols = st.columns(4)
for i, (label, value, badge, kind) in enumerate(kpis):
    badge_class = {"up": "badge-up", "down": "badge-down", "neu": "badge-neu"}[kind]
    with cols[i % 4]:
        st.markdown(f"""
        <div class="metric-card" style="margin-bottom:14px">
            <div class="m-label">{label}</div>
            <div class="m-value">{value}</div>
            <div><span class="m-badge {badge_class}">{badge}</span></div>
        </div>
        """, unsafe_allow_html=True)

st.markdown("<div style='margin-bottom:20px'></div>", unsafe_allow_html=True)

# ── Chart colours ─────────────────────────────────────────────────────────────
BG   = "#161922"
GRID = "rgba(255,255,255,0.05)"
TEXT = "#9ca3af"
BLUE = "#4f8ef7"
PURP = "#7c5af7"
GRN  = "#34d399"
AMB  = "#fbbf24"
RED  = "#f87171"
GRAY = "#6b7280"

months = ["Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]

# ── Row 1: Volume + Donut ─────────────────────────────────────────────────────
c1, c2 = st.columns([3, 2])

with c1:
    st.markdown('<div class="section-head">Monthly task volume</div>', unsafe_allow_html=True)
    fig = go.Figure(go.Bar(
        x=months,
        y=[8200, 9100, 10400, 11200, 12600, 14820],
        marker_color=[BLUE]*5 + [PURP],
        marker_line_width=0,
    ))
    fig.update_layout(
        paper_bgcolor=BG, plot_bgcolor=BG,
        height=240, margin=dict(l=0, r=0, t=10, b=0),
        font=dict(family="DM Sans", color=TEXT, size=12),
        xaxis=dict(showgrid=False, tickcolor=GRID, linecolor=GRID),
        yaxis=dict(gridcolor=GRID, tickformat=",.0f", ticksuffix=""),
        bargap=0.3,
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

with c2:
    st.markdown('<div class="section-head">Failure reasons breakdown</div>', unsafe_allow_html=True)
    fig2 = go.Figure(go.Pie(
        labels=["Timeout", "Auth error", "Bad input", "Other"],
        values=[38, 27, 19, 16],
        hole=0.68,
        marker=dict(colors=[RED, AMB, BLUE, GRAY], line=dict(color=BG, width=2)),
        textfont=dict(family="DM Sans", size=12),
        hovertemplate="%{label}: %{value}%<extra></extra>",
    ))
    fig2.update_layout(
        paper_bgcolor=BG, plot_bgcolor=BG,
        height=240, margin=dict(l=0, r=0, t=10, b=0),
        font=dict(family="DM Sans", color=TEXT),
        legend=dict(
            orientation="v", x=1.02, y=0.5,
            font=dict(size=12, color=TEXT),
            bgcolor="rgba(0,0,0,0)",
        ),
        annotations=[dict(text="Failures", x=0.5, y=0.5, showarrow=False,
                          font=dict(size=13, color=TEXT, family="DM Sans"))],
    )
    st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})

# ── Row 2: Bar health + Table ─────────────────────────────────────────────────
c3, c4 = st.columns([2, 3])

with c3:
    st.markdown('<div class="section-head">Workflow category health</div>', unsafe_allow_html=True)
    cats   = ["Data sync", "Notifications", "File processing", "API integrations", "Report gen", "ML pipelines"]
    scores = [94, 98, 91, 88, 96, 79]
    colors = [GRN if s >= 90 else (AMB if s >= 85 else RED) for s in scores]
    fig3 = go.Figure(go.Bar(
        x=scores, y=cats, orientation="h",
        marker_color=colors, marker_line_width=0,
        text=[f"{s}%" for s in scores],
        textposition="outside",
        textfont=dict(size=12, color=TEXT, family="DM Mono"),
    ))
    fig3.update_layout(
        paper_bgcolor=BG, plot_bgcolor=BG,
        height=280, margin=dict(l=0, r=50, t=10, b=0),
        font=dict(family="DM Sans", color=TEXT, size=12),
        xaxis=dict(range=[60, 105], showgrid=False, showticklabels=False),
        yaxis=dict(showgrid=False, tickcolor=GRID),
        bargap=0.35,
    )
    st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar": False})

with c4:
    st.markdown('<div class="section-head">Top workflows by volume</div>', unsafe_allow_html=True)
    df = pd.DataFrame({
        "Workflow":       ["Invoice processor", "CRM data sync", "Email classifier",
                           "Report scheduler", "ML batch pipeline", "Legacy API bridge"],
        "Runs":           [3210, 2840, 2100, 1760, 980, 640],
        "Success Rate":   ["99.1%", "97.4%", "95.6%", "98.2%", "79.3%", "71.0%"],
        "Status":         ["🟢 Healthy", "🟢 Healthy", "🟢 Healthy",
                           "🟢 Healthy", "🟡 Warning", "🔴 Critical"],
    })
    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        height=260,
    )

# ── Row 3: Line + SLA bar ─────────────────────────────────────────────────────
c5, c6 = st.columns(2)

with c5:
    st.markdown('<div class="section-head">Avg execution time (s) — last 6 months</div>', unsafe_allow_html=True)
    fig5 = go.Figure()
    fig5.add_trace(go.Scatter(
        x=months, y=[1.9, 1.8, 1.7, 1.6, 1.5, 1.4],
        mode="lines+markers",
        line=dict(color=PURP, width=2.5, shape="spline"),
        marker=dict(size=7, color=PURP, line=dict(color=BG, width=2)),
        fill="tozeroy",
        fillcolor="rgba(124,90,247,0.1)",
        name="Avg time",
    ))
    fig5.update_layout(
        paper_bgcolor=BG, plot_bgcolor=BG,
        height=220, margin=dict(l=0, r=0, t=10, b=0),
        font=dict(family="DM Sans", color=TEXT, size=12),
        xaxis=dict(showgrid=False, tickcolor=GRID),
        yaxis=dict(gridcolor=GRID, range=[1.0, 2.2], ticksuffix="s"),
        showlegend=False,
    )
    st.plotly_chart(fig5, use_container_width=True, config={"displayModeBar": False})

with c6:
    st.markdown('<div class="section-head">SLA compliance by category</div>', unsafe_allow_html=True)
    sla_cats   = ["Data sync", "Notif", "Files", "API", "Reports", "ML"]
    sla_scores = [99, 100, 97, 89, 98, 76]
    sla_colors = [GRN if s >= 90 else (AMB if s >= 85 else RED) for s in sla_scores]
    fig6 = go.Figure(go.Bar(
        x=sla_cats, y=sla_scores,
        marker_color=sla_colors, marker_line_width=0,
        text=[f"{s}%" for s in sla_scores],
        textposition="outside",
        textfont=dict(size=12, color=TEXT, family="DM Mono"),
    ))
    fig6.add_hline(y=95, line_dash="dot", line_color="rgba(255,255,255,0.2)",
                   annotation_text="Target 95%", annotation_font_color=TEXT,
                   annotation_font_size=11)
    fig6.update_layout(
        paper_bgcolor=BG, plot_bgcolor=BG,
        height=220, margin=dict(l=0, r=0, t=10, b=30),
        font=dict(family="DM Sans", color=TEXT, size=12),
        xaxis=dict(showgrid=False, tickcolor=GRID),
        yaxis=dict(gridcolor=GRID, range=[60, 108], ticksuffix="%"),
        showlegend=False,
        bargap=0.35,
    )
    st.plotly_chart(fig6, use_container_width=True, config={"displayModeBar": False})

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="border-top:1px solid rgba(255,255,255,0.06);margin-top:1rem;padding-top:1rem;
     display:flex;justify-content:space-between;align-items:center">
    <span style="font-size:12px;color:#374151;font-family:'DM Mono',monospace">
        automation-kpi-dashboard · march 2026
    </span>
    <span style="font-size:12px;color:#374151">
        data refreshes on page reload
    </span>
</div>
""", unsafe_allow_html=True)
