import io
import warnings
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st

warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="Seminar Intelligence Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Syne:wght@700;800&display=swap" rel="stylesheet">
    <style>
      :root {
        --bg: #08111f;
        --panel: #0f1728;
        --panel-2: #111c31;
        --line: rgba(255,255,255,0.08);
        --text: #f2f7ff;
        --muted: #8ea1c2;
        --blue: #60a5fa;
        --green: #4fce8f;
        --red: #fb7185;
        --amber: #fbbf24;
        --purple: #a78bfa;
        --teal: #22d3ee;
      }
      html, body, [class*="css"] {font-family:'Inter', sans-serif;}
      .stApp {
        color: var(--text);
        background:
          radial-gradient(circle at top left, rgba(34,211,238,0.10), transparent 22%),
          radial-gradient(circle at top right, rgba(167,139,250,0.12), transparent 18%),
          linear-gradient(180deg, #08111f 0%, #091220 48%, #060c16 100%);
      }
      #MainMenu, header, footer {visibility:hidden;}
      .block-container {padding-top: 1rem; max-width: 100% !important;}
      div[data-testid="stToolbar"], div[data-testid="stDecoration"], div[data-testid="stStatusWidget"] {display:none !important;}
      .hero {
        background: linear-gradient(135deg, rgba(17,28,49,.98), rgba(9,16,29,.98));
        border:1px solid var(--line);
        border-radius:24px;
        padding:22px 24px;
        margin-bottom:14px;
        box-shadow: 0 18px 46px rgba(0,0,0,.28);
      }
      .hero h1 {
        font-family:'Syne', sans-serif;
        margin:0;
        font-size:34px;
        line-height:1.08;
        letter-spacing:-0.04em;
      }
      .hero p {
        margin:10px 0 0 0;
        color: var(--muted);
        font-size:14px;
        line-height:1.6;
      }
      .chip-row {display:flex; gap:10px; flex-wrap:wrap; margin-top:14px;}
      .chip {
        padding:8px 12px;
        border-radius:999px;
        background:rgba(255,255,255,.05);
        border:1px solid var(--line);
        color:#dbe7fb;
        font-size:12px;
      }
      .panel {
        background: linear-gradient(180deg, rgba(15,23,40,.96), rgba(10,17,30,.98));
        border:1px solid var(--line);
        border-radius:22px;
        padding:16px;
        margin-bottom:14px;
        box-shadow:0 12px 32px rgba(0,0,0,.20);
      }
      .panel-title {
        display:flex; justify-content:space-between; align-items:center; gap:10px; margin-bottom:14px;
      }
      .panel-title h3 {
        margin:0; font-size:14px; text-transform:uppercase; letter-spacing:.12em; color:#f7d168;
      }
      .panel-title p {
        margin:0; color:var(--muted); font-size:12px;
      }
      .kpi-grid {display:grid; gap:12px;}
      .kpi {
        position:relative; overflow:hidden;
        background: linear-gradient(180deg, rgba(18,29,50,.98), rgba(11,19,34,.98));
        border:1px solid rgba(255,255,255,.08);
        border-radius:20px;
        padding:16px;
        min-height:118px;
      }
      .kpi::before {
        content:""; position:absolute; left:0; right:0; top:0; height:3px;
        background: linear-gradient(90deg, var(--accent), transparent);
      }
      .kpi::after {
        content:""; position:absolute; width:120px; height:120px; right:-44px; top:-44px;
        background: radial-gradient(circle, color-mix(in srgb, var(--accent) 22%, transparent), transparent 65%);
      }
      .kpi-label {font-size:11px; text-transform:uppercase; letter-spacing:.12em; color:var(--muted);}
      .kpi-value {font-family:'Syne', sans-serif; font-size:28px; margin-top:10px; line-height:1;}
      .kpi-sub {margin-top:10px; font-size:12px; color:#d6e2f4;}
      .info {--accent: var(--blue);} .success {--accent: var(--green);} .danger {--accent: var(--red);} .warning {--accent: var(--amber);} .purple {--accent: var(--purple);} .teal {--accent: var(--teal);}
      .login-card {
        max-width: 460px;
        margin: 8vh auto 0 auto;
        background: linear-gradient(180deg, rgba(15,23,40,.98), rgba(10,17,30,.98));
        border:1px solid var(--line);
        border-radius:24px;
        padding:26px;
        box-shadow:0 20px 50px rgba(0,0,0,.30);
      }
      .login-card h2 {font-family:'Syne', sans-serif; margin:0; font-size:30px;}
      .login-card p {color:var(--muted); font-size:14px; line-height:1.6;}
      .helper {color: var(--muted); font-size:12px;}
      div[data-baseweb="select"] > div,
      div[data-baseweb="input"] > div,
      .stDateInput > div > div,
      .stNumberInput > div > div {
        background:#111c31 !important;
        border-color:rgba(255,255,255,.08) !important;
        border-radius:12px !important;
      }
      .stMultiSelect [data-baseweb="tag"] {
        background:rgba(96,165,250,.15) !important;
        border:1px solid rgba(96,165,250,.25) !important;
      }
      .stTabs [data-baseweb="tab-list"] {gap:8px; border-bottom:none;}
      .stTabs [data-baseweb="tab"] {
        border:1px solid rgba(255,255,255,.08) !important;
        background:#111c31 !important;
        border-radius:12px !important;
        color:#a9bbd8 !important;
        padding:.5rem 1rem !important;
      }
      .stTabs [aria-selected="true"] {
        color:white !important;
        border-color:rgba(96,165,250,.40) !important;
        box-shadow: inset 0 0 0 1px rgba(96,165,250,.15);
      }
      button[data-testid="stBaseButton-primary"] {
        background:linear-gradient(135deg,#60a5fa,#22d3ee) !important;
        color:#07111f !important;
        border:none !important;
        border-radius:12px !important;
        font-weight:800 !important;
      }
      button[data-testid="stBaseButton-secondary"] {
        background:#111c31 !important;
        color:#dbe7fb !important;
        border:1px solid rgba(255,255,255,.10) !important;
        border-radius:12px !important;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

for key, default in {
    "logged_in": False,
    "master": None,
}.items():
    st.session_state.setdefault(key, default)


def fmt_currency(v):
    if pd.isna(v):
        v = 0
    v = float(v)
    if v >= 1e7:
        return f"₹{v/1e7:.2f}Cr"
    if v >= 1e5:
        return f"₹{v/1e5:.2f}L"
    return f"₹{v:,.0f}"


def safe_col(df, col, default="Unknown"):
    if col not in df.columns:
        df[col] = default
    return df


def clean_phone(series):
    return series.astype(str).str.replace(r"\D", "", regex=True).str[-10:].replace("", pd.NA)


def bool_from_text(series):
    s = series.astype(str).str.strip().str.lower()
    return s.isin(["yes", "true", "converted", "1"])


def draw_kpis(cards, cols=6):
    html = [f'<div class="kpi-grid" style="grid-template-columns:repeat({cols}, minmax(0,1fr))">']
    for c in cards:
        html.append(
            f'''<div class="kpi {c[3]}">
                <div class="kpi-label">{c[0]}</div>
                <div class="kpi-value">{c[1]}</div>
                <div class="kpi-sub">{c[2]}</div>
            </div>'''
        )
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def safe_plot_bar(df, x, y, color=None, height=320, horizontal=False, title=""):
    required = [x, y]
    if any(col not in df.columns for col in required) or df.empty:
        st.info("No chart data available for the current filters.")
        return
    chart_df = df.copy()
    if color not in chart_df.columns:
        color = None
    chart_df[y] = pd.to_numeric(chart_df[y], errors="coerce").fillna(0)
    fig = px.bar(
        chart_df,
        x=x if not horizontal else y,
        y=y if not horizontal else x,
        color=color,
        orientation="h" if horizontal else "v",
        title=title,
    )
    fig.update_layout(
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#dce6f8"),
        margin=dict(l=10, r=10, t=40 if title else 10, b=10),
        legend_title_text="",
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="rgba(255,255,255,0.08)")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def apply_text_filters(df, columns, key_prefix="tbl"):
    if df.empty:
        return df
    st.markdown('<div class="panel-title"><h3>Table Filters</h3><p>Filter any displayed column using contains-match search.</p></div>', unsafe_allow_html=True)
    filtered = df.copy()
    rows = [columns[i:i + 4] for i in range(0, len(columns), 4)]
    for ridx, row_cols in enumerate(rows):
        ui_cols = st.columns(len(row_cols))
        for cidx, col in enumerate(row_cols):
            q = ui_cols[cidx].text_input(col, key=f"{key_prefix}_{ridx}_{cidx}", placeholder=f"Filter {col}")
            if q:
                filtered = filtered[filtered[col].astype(str).str.contains(q, case=False, na=False)]
    return filtered


def render_login():
    st.markdown(
        """
        <div class="login-card">
          <h2>Seminar Intelligence Pro</h2>
          <p>Login to access the unified post-seminar conversion dashboard. This version is built around one core workflow: <b>offline seminar attendees → conversion after seminar → course, payment, due, lead-source intelligence, course share, and combo-course cross-sell</b>.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    left, center, right = st.columns([1.2, 1.5, 1.2])
    with center:
        username = st.text_input("Username", placeholder="Enter username")
        password = st.text_input("Password", type="password", placeholder="Enter password")
        login_clicked = st.button("Login", type="primary", use_container_width=True)
        st.caption("Demo credentials: username `admin` and password `invesmate123`")
        if login_clicked:
            if username == "admin" and password == "invesmate123":
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("Invalid username or password.")


if not st.session_state.logged_in:
    render_login()
    st.stop()


@st.cache_data(show_spinner=False)
def process_all(sem_bytes, conv_bytes, lead_bytes):
    sem = pd.read_csv(io.BytesIO(sem_bytes))
    sem.columns = sem.columns.str.strip()
    for col in [
        "NAME", "Mobile", "Place", "Trainer / Presenter", "Session",
        "Is Attended ?", "Is Converted ?", "TRADER", "Is our Student ?",
        "Mode of Payment", "Remarks"
    ]:
        safe_col(sem, col, "")
    safe_col(sem, "Seminar Date", pd.NaT)
    safe_col(sem, "Amount Paid", 0)
    sem["Seminar Date"] = pd.to_datetime(sem["Seminar Date"], errors="coerce", dayfirst=True)
    sem["Amount Paid"] = pd.to_numeric(sem["Amount Paid"], errors="coerce").fillna(0)
    for col in [
        "NAME", "Place", "Trainer / Presenter", "Session", "Is Attended ?",
        "Is Converted ?", "TRADER", "Is our Student ?", "Mode of Payment", "Remarks"
    ]:
        sem[col] = sem[col].astype(str).str.strip()
    sem["mobile_clean"] = clean_phone(sem["Mobile"])
    sem["Trainer Norm"] = sem["Trainer / Presenter"].str.upper()
    sem["Session"] = sem["Session"].str.upper()
    sem["attended_flag"] = bool_from_text(sem["Is Attended ?"])
    attended = sem[sem["attended_flag"]].copy()

    conv = pd.read_excel(io.BytesIO(conv_bytes))
    conv.columns = conv.columns.str.strip()
    for col in [
        "phone", "service_name", "payment_mode", "status", "sales_rep_name",
        "trainer", "student_name", "email"
    ]:
        safe_col(conv, col, "")
    for col in ["payment_received", "total_amount", "total_due", "total_gst"]:
        safe_col(conv, col, 0)
    safe_col(conv, "order_date", pd.NaT)

    conv["order_date"] = pd.to_datetime(conv["order_date"], errors="coerce", utc=True).dt.tz_localize(None)
    conv["payment_received"] = pd.to_numeric(conv["payment_received"], errors="coerce").fillna(0)
    conv["total_amount"] = pd.to_numeric(conv["total_amount"], errors="coerce").fillna(0)
    conv["total_due"] = pd.to_numeric(conv["total_due"], errors="coerce").fillna(0)
    conv["phone_clean"] = clean_phone(conv["phone"])
    conv["service_name"] = conv["service_name"].astype(str).str.strip()
    conv["status"] = conv["status"].astype(str).str.strip()
    conv["sales_rep_name"] = conv["sales_rep_name"].astype(str).str.strip()
    conv["payment_mode"] = conv["payment_mode"].astype(str).str.strip()
    conv["trainer_name"] = conv["trainer"].astype(str).str.split(" - ").str[-1].str.strip()
    conv["month"] = conv["order_date"].dt.to_period("M").astype(str)
    conv["due_zero"] = conv["total_due"].fillna(0).le(0)
    conv["service_name_norm"] = conv["service_name"].astype(str).str.lower().str.replace(r"\s+", " ", regex=True).str.strip()

    try:
        leads = pd.read_excel(io.BytesIO(lead_bytes), sheet_name="Sheet 1")
    except Exception:
        leads = pd.read_excel(io.BytesIO(lead_bytes))
    leads.columns = leads.columns.str.strip()
    for col in [
        "name", "phone", "email", "converted_from", "leadsource", "campaign_name",
        "leadstatus", "stage_name", "leadownername", "state", "Attempted/Unattempted",
        "servicename", "remarks"
    ]:
        safe_col(leads, col, "Unknown")
    safe_col(leads, "leaddate", pd.NaT)

    leads["leaddate"] = pd.to_datetime(leads["leaddate"], errors="coerce")
    for col in [
        "converted_from", "leadsource", "campaign_name", "leadstatus", "stage_name",
        "leadownername", "state", "Attempted/Unattempted", "servicename", "remarks"
    ]:
        leads[col] = leads[col].astype(str).str.strip().replace({"nan": "Unknown", "": "Unknown"})
    leads["phone_clean"] = clean_phone(leads["phone"])
    leads["leadstatus_lower"] = leads["leadstatus"].astype(str).str.lower()
    leads["is_converted"] = leads["leadstatus_lower"].str.contains("converted", na=False)

    leads_one = (
        leads.sort_values("leaddate", na_position="last")
        .drop_duplicates("phone_clean", keep="last")
        [[
            "phone_clean", "converted_from", "leadsource", "campaign_name", "leadstatus",
            "stage_name", "leadownername", "state", "Attempted/Unattempted", "servicename",
            "is_converted", "leaddate", "name", "email", "remarks"
        ]]
    )

    merged = attended.merge(conv, left_on="mobile_clean", right_on="phone_clean", how="left", suffixes=("", "_conv"))
    merged["after_seminar"] = (
        merged["order_date"].ge(merged["Seminar Date"]) &
        merged["order_date"].notna() &
        merged["Seminar Date"].notna()
    )

    post = merged[merged["after_seminar"]].copy()
    if not post.empty:
        post = post.sort_values(["mobile_clean", "Seminar Date", "order_date", "total_amount"], ascending=[True, True, True, False])
        first_post = post.drop_duplicates(["mobile_clean", "Seminar Date"], keep="first")
    else:
        first_post = attended.copy()
        for col in conv.columns:
            if col not in first_post.columns:
                first_post[col] = pd.NA
        first_post["after_seminar"] = False

    selected_cols = [
        "mobile_clean", "Seminar Date", "order_date", "service_name", "payment_received",
        "total_amount", "total_due", "status", "sales_rep_name", "payment_mode",
        "trainer_name", "month", "due_zero", "after_seminar", "student_name", "email"
    ]
    base = attended.merge(first_post[[c for c in selected_cols if c in first_post.columns]], on=["mobile_clean", "Seminar Date"], how="left")
    base = base.merge(leads_one, left_on="mobile_clean", right_on="phone_clean", how="left")
    base["conversion_status"] = base["after_seminar"].map({True: "Converted After Seminar", False: "Not Converted After Seminar"}).fillna("Not Converted After Seminar")
    base["lead_origin"] = base["converted_from"].fillna("Unknown")
    base["lead_source_name"] = base["leadsource"].fillna("Unknown")
    base["due_bucket"] = "No Order"
    base.loc[base["after_seminar"].fillna(False) & base["total_due"].fillna(0).le(0), "due_bucket"] = "Due 0"
    base.loc[base["after_seminar"].fillna(False) & base["total_due"].fillna(0).gt(0), "due_bucket"] = "Has Due"

    # Combo-course cross sell analysis using full conversion list.
    combo_pattern = "power of trading & investing combo course"
    combo_orders = conv[conv["service_name_norm"].str.contains(combo_pattern, na=False)].copy()
    if not combo_orders.empty:
        combo_first = combo_orders.sort_values("order_date").drop_duplicates("phone_clean", keep="first")[["phone_clean", "order_date"]].rename(columns={"order_date": "combo_order_date"})
        later_orders = conv.merge(combo_first, on="phone_clean", how="inner")
        later_orders = later_orders[later_orders["order_date"] > later_orders["combo_order_date"]]
        later_orders = later_orders[~later_orders["service_name_norm"].str.contains(combo_pattern, na=False)]
        later_orders = later_orders.sort_values(["phone_clean", "order_date"]).drop_duplicates(["phone_clean", "service_name_norm"], keep="first")
    else:
        combo_first = pd.DataFrame(columns=["phone_clean", "combo_order_date"])
        later_orders = pd.DataFrame(columns=list(conv.columns) + ["combo_order_date"])

    return {
        "seminar": sem,
        "attended": attended,
        "conversion": conv,
        "leads": leads,
        "master": base,
        "combo_base": combo_first,
        "combo_other_orders": later_orders,
    }


st.markdown(
    f"""
    <div class="hero">
      <h1>Unified Offline Seminar Conversion Intelligence</h1>
      <p>This dashboard is designed around one journey: <b>who attended the offline seminar</b>, <b>who converted after attending</b>, <b>when the order was created</b>, <b>which course and payment amount were mapped</b>, <b>whether due is zero</b>, and <b>whether the student originated from webinar or non-webinar along with lead source</b>.</p>
      <div class="chip-row">
        <div class="chip">All filters on one page</div>
        <div class="chip">Course share per seminar</div>
        <div class="chip">Combo-course cross-sell</div>
        <div class="chip">Lead table all-column filters</div>
        <div class="chip">{datetime.now().strftime('%d %b %Y %H:%M')}</div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

up1, up2, up3, up4 = st.columns([1, 1, 1, 0.65])
f_sem = up1.file_uploader("Seminar CSV", type=["csv"], key="sem")
f_conv = up2.file_uploader("Conversion XLSX", type=["xlsx", "xls"], key="conv")
f_lead = up3.file_uploader("Leads XLSX", type=["xlsx", "xls"], key="lead")
load_btn = up4.button("Load Data", type="primary", use_container_width=True)

if load_btn:
    if not (f_sem and f_conv and f_lead):
        st.error("Upload all 3 files first.")
    else:
        try:
            st.session_state.master = process_all(f_sem.read(), f_conv.read(), f_lead.read())
            st.success("Files loaded successfully.")
        except Exception as e:
            st.exception(e)
            st.stop()

if not st.session_state.master:
    st.info("Upload the 3 files and click Load Data to continue.")
    st.stop()

D = st.session_state.master
master = D["master"].copy()

st.markdown(
    """
    <div class="panel-title"><h3>Unified Filters</h3><p>Apply once and update every KPI, chart, course-share view, combo analysis, and lead table together.</p></div>
    """,
    unsafe_allow_html=True,
)

c1, c2, c3, c4, c5 = st.columns(5)
seminar_dates = sorted([d for d in master["Seminar Date"].dropna().dt.date.unique()])
seminar_date_sel = c1.multiselect("Seminar Date", seminar_dates, format_func=lambda x: x.strftime("%d %b %Y"))
place_sel = c2.multiselect("Location", sorted(master["Place"].dropna().astype(str).unique()))
trainer_sel = c3.multiselect("Trainer", sorted(master["Trainer Norm"].dropna().astype(str).unique()))
session_sel = c4.multiselect("Session", sorted(master["Session"].dropna().astype(str).unique()))
conv_status_sel = c5.multiselect("Conversion Status", ["Converted After Seminar", "Not Converted After Seminar"])

c6, c7, c8, c9, c10 = st.columns(5)
course_sel = c6.multiselect("Course", sorted([x for x in master["service_name"].dropna().astype(str).unique() if x and x != "nan"]))
due_sel = c7.multiselect("Due Filter", ["Due 0", "Has Due", "No Order"])
lead_origin_sel = c8.multiselect("Lead Origin", sorted([x for x in master["lead_origin"].dropna().astype(str).unique() if x and x != "nan"]))
lead_source_sel = c9.multiselect("Lead Source Name", sorted([x for x in master["lead_source_name"].dropna().astype(str).unique() if x and x != "nan"]))
payment_mode_sel = c10.multiselect("Payment Mode", sorted([x for x in master["payment_mode"].dropna().astype(str).unique() if x and x != "nan"]))

c11, c12, c13, c14, c15 = st.columns(5)
sales_sel = c11.multiselect("Sales Rep", sorted([x for x in master["sales_rep_name"].dropna().astype(str).unique() if x and x != "nan"]))
status_sel = c12.multiselect("Order Status", sorted([x for x in master["status"].dropna().astype(str).unique() if x and x != "nan"]))
only_after = c13.selectbox("Order Timing", ["Only orders after seminar", "Show all attendee matches"], index=0)
payment_min = c14.number_input("Min Paid", min_value=0.0, value=0.0, step=1000.0)
payment_max_default = float(max(master["payment_received"].fillna(0).max(), 0))
payment_max = c15.number_input("Max Paid", min_value=0.0, value=payment_max_default, step=1000.0)

c16, c17, c18, c19 = st.columns([1.2, 1.2, 1.2, 2.4])
search_text = c16.text_input("Student Search", placeholder="Name / mobile / course")
due_zero_only = c17.checkbox("Only Due = 0", value=False)
seminar_share_basis = c18.selectbox("Seminar Share Basis", ["Paid Amount", "Student Count"])
combo_scope = c19.selectbox("Combo Cross-Sell Scope", ["All students in conversion list", "Only filtered attendee phones"])

filtered = master.copy()
if seminar_date_sel:
    filtered = filtered[filtered["Seminar Date"].dt.date.isin(seminar_date_sel)]
if place_sel:
    filtered = filtered[filtered["Place"].isin(place_sel)]
if trainer_sel:
    filtered = filtered[filtered["Trainer Norm"].isin(trainer_sel)]
if session_sel:
    filtered = filtered[filtered["Session"].isin(session_sel)]
if conv_status_sel:
    filtered = filtered[filtered["conversion_status"].isin(conv_status_sel)]
if course_sel:
    filtered = filtered[filtered["service_name"].isin(course_sel)]
if lead_origin_sel:
    filtered = filtered[filtered["lead_origin"].isin(lead_origin_sel)]
if lead_source_sel:
    filtered = filtered[filtered["lead_source_name"].isin(lead_source_sel)]
if payment_mode_sel:
    filtered = filtered[filtered["payment_mode"].isin(payment_mode_sel)]
if sales_sel:
    filtered = filtered[filtered["sales_rep_name"].isin(sales_sel)]
if status_sel:
    filtered = filtered[filtered["status"].isin(status_sel)]
if only_after == "Only orders after seminar":
    filtered = filtered[filtered["after_seminar"].fillna(False) | filtered["order_date"].isna()]
filtered = filtered[filtered["payment_received"].fillna(0).between(payment_min, payment_max)]
if due_zero_only:
    filtered = filtered[filtered["total_due"].fillna(0).le(0)]
if due_sel:
    due_mask = pd.Series(False, index=filtered.index)
    if "Due 0" in due_sel:
        due_mask = due_mask | (filtered["total_due"].fillna(0).le(0) & filtered["after_seminar"].fillna(False))
    if "Has Due" in due_sel:
        due_mask = due_mask | (filtered["total_due"].fillna(0).gt(0) & filtered["after_seminar"].fillna(False))
    if "No Order" in due_sel:
        due_mask = due_mask | filtered["order_date"].isna()
    filtered = filtered[due_mask]
if search_text:
    q = search_text.strip().lower()
    mask = (
        filtered["NAME"].astype(str).str.lower().str.contains(q, na=False)
        | filtered["mobile_clean"].astype(str).str.contains(q, na=False)
        | filtered["service_name"].astype(str).str.lower().str.contains(q, na=False)
        | filtered["lead_source_name"].astype(str).str.lower().str.contains(q, na=False)
    )
    filtered = filtered[mask]

attendee_count = len(filtered)
converted_after = int(filtered["after_seminar"].fillna(False).sum())
not_converted = attendee_count - converted_after
paid = filtered.loc[filtered["after_seminar"].fillna(False), "payment_received"].fillna(0).sum()
due = filtered.loc[filtered["after_seminar"].fillna(False), "total_due"].fillna(0).sum()
due_zero = int((filtered["after_seminar"].fillna(False) & filtered["total_due"].fillna(0).le(0)).sum())
webinar = int((filtered["lead_origin"].astype(str).str.lower() == "webinar").sum())
non_webinar = int((filtered["lead_origin"].astype(str).str.lower() == "non webinar").sum())
conv_rate = (converted_after / attendee_count * 100) if attendee_count else 0

st.markdown('<div class="panel">', unsafe_allow_html=True)
draw_kpis(
    [
        ("Offline Attendees", f"{attendee_count:,}", "Filtered attendee rows", "info"),
        ("Converted After Seminar", f"{converted_after:,}", f"{conv_rate:.1f}% conversion rate", "success"),
        ("Not Converted", f"{not_converted:,}", "Attended but no post-seminar order", "danger"),
        ("Paid Amount", fmt_currency(paid), "Collected from converted students", "teal"),
        ("Total Due", fmt_currency(due), "Outstanding after conversion", "warning"),
        ("Due 0 Students", f"{due_zero:,}", "Converted students with zero due", "purple"),
        ("Webinar Origin", f"{webinar:,}", "Matched in leads as webinar", "teal"),
        ("Non-Webinar Origin", f"{non_webinar:,}", "Matched in leads as non-webinar", "purple"),
    ],
    cols=4,
)
st.markdown('</div>', unsafe_allow_html=True)

t1, t2, t3 = st.tabs(["Overview", "Course & Lead Intelligence", "Student Records"])

with t1:
    a1, a2 = st.columns(2)
    with a1:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="panel-title"><h3>Course-wise Paid Amount</h3><p>Which course collected how much after seminar attendance.</p></div>', unsafe_allow_html=True)
        course_df = (
            filtered[filtered["after_seminar"].fillna(False)]
            .groupby("service_name", dropna=False)
            .agg(Students=("mobile_clean", "count"), Paid=("payment_received", "sum"), Due=("total_due", "sum"), Revenue=("total_amount", "sum"))
            .reset_index()
            .sort_values("Paid", ascending=False)
        )
        course_df["service_name"] = course_df["service_name"].replace("", "Unknown")
        safe_plot_bar(course_df.head(12), "service_name", "Paid", height=360, horizontal=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with a2:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="panel-title"><h3>Seminar Date vs Conversion</h3><p>How each offline seminar date performed post-event.</p></div>', unsafe_allow_html=True)
        trend = (
            filtered.assign(SeminarLabel=filtered["Seminar Date"].dt.strftime("%d %b %Y"))
            .groupby("SeminarLabel", dropna=False)
            .agg(Attendees=("mobile_clean", "count"), Converted=("after_seminar", "sum"))
            .reset_index()
            .rename(columns={"SeminarLabel": "Seminar"})
        )
        safe_plot_bar(trend, "Seminar", "Converted", height=360)
        st.markdown('</div>', unsafe_allow_html=True)

    b1, b2 = st.columns(2)
    with b1:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="panel-title"><h3>Course Share for Selected Seminar</h3><p>Share by course for the selected seminar date(s), based on paid amount or student count.</p></div>', unsafe_allow_html=True)
        seminar_share = filtered[filtered["after_seminar"].fillna(False)].copy()
        if not seminar_share.empty:
            seminar_share["SeminarLabel"] = seminar_share["Seminar Date"].dt.strftime("%d %b %Y")
            if seminar_share_basis == "Paid Amount":
                share_df = seminar_share.groupby(["SeminarLabel", "service_name"], dropna=False)["payment_received"].sum().reset_index(name="Value")
            else:
                share_df = seminar_share.groupby(["SeminarLabel", "service_name"], dropna=False).size().reset_index(name="Value")
            share_df["service_name"] = share_df["service_name"].replace("", "Unknown")
            share_df["Share %"] = share_df.groupby("SeminarLabel")["Value"].transform(lambda s: (s / s.sum() * 100).round(1) if s.sum() else 0)
            safe_plot_bar(share_df.sort_values("Value", ascending=False).head(20), "service_name", "Share %", color="SeminarLabel", height=360, horizontal=True)
            st.dataframe(share_df.sort_values(["SeminarLabel", "Value"], ascending=[True, False]), use_container_width=True, hide_index=True, height=220)
        else:
            st.info("No converted post-seminar rows available for course-share analysis.")
        st.markdown('</div>', unsafe_allow_html=True)

    with b2:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="panel-title"><h3>Lead Source Name</h3><p>Source mapped for filtered attendee records.</p></div>', unsafe_allow_html=True)
        source_df = filtered["lead_source_name"].fillna("Unknown").replace("", "Unknown").value_counts().rename_axis("Lead Source").reset_index(name="Count").head(15)
        safe_plot_bar(source_df, "Lead Source", "Count", height=360, horizontal=True)
        st.markdown('</div>', unsafe_allow_html=True)

with t2:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="panel-title"><h3>Webinar vs Non-Webinar</h3><p>Lead origin for the filtered attendee base.</p></div>', unsafe_allow_html=True)
        origin_df = filtered["lead_origin"].fillna("Unknown").replace("", "Unknown").value_counts().rename_axis("Lead Origin").reset_index(name="Count")
        safe_plot_bar(origin_df, "Lead Origin", "Count", height=340)
        st.markdown('</div>', unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="panel-title"><h3>Course-wise Due 0</h3><p>Which course has students fully paid after seminar conversion.</p></div>', unsafe_allow_html=True)
        due0_df = (
            filtered[filtered["after_seminar"].fillna(False) & filtered["total_due"].fillna(0).le(0)]
            .groupby("service_name", dropna=False)
            .size().rename_axis("Course").reset_index(name="Students")
            .sort_values("Students", ascending=False)
        )
        due0_df["Course"] = due0_df["Course"].replace("", "Unknown")
        safe_plot_bar(due0_df.head(12), "Course", "Students", height=340, horizontal=True)
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="panel-title"><h3>Power Of Trading & Investing Combo Course → Other Course Buyers</h3><p>Students who first took the combo course and later purchased some other course.</p></div>', unsafe_allow_html=True)
    combo_base = D["combo_base"].copy()
    combo_other = D["combo_other_orders"].copy()
    if combo_scope == "Only filtered attendee phones":
        allowed_phones = set(filtered["mobile_clean"].dropna().astype(str))
        combo_base = combo_base[combo_base["phone_clean"].astype(str).isin(allowed_phones)]
        combo_other = combo_other[combo_other["phone_clean"].astype(str).isin(allowed_phones)]
    combo_students = int(combo_base["phone_clean"].nunique())
    combo_cross_sell_students = int(combo_other["phone_clean"].nunique())
    combo_cross_sell_orders = int(len(combo_other))
    cross_rate = (combo_cross_sell_students / combo_students * 100) if combo_students else 0
    draw_kpis(
        [
            ("Combo Buyers", f"{combo_students:,}", "Students with combo course", "info"),
            ("Bought Other Course", f"{combo_cross_sell_students:,}", f"{cross_rate:.1f}% of combo buyers", "success"),
            ("Other Course Orders", f"{combo_cross_sell_orders:,}", "Distinct phone+course rows", "purple"),
            ("Extra Revenue", fmt_currency(combo_other["total_amount"].sum() if not combo_other.empty else 0), "Revenue from later non-combo courses", "teal"),
        ],
        cols=4,
    )
    if not combo_other.empty:
        other_course_df = combo_other.groupby("service_name", dropna=False).agg(Students=("phone_clean", "nunique"), Orders=("phone_clean", "count"), Paid=("payment_received", "sum")).reset_index().sort_values("Students", ascending=False)
        other_course_df["service_name"] = other_course_df["service_name"].replace("", "Unknown")
        x1, x2 = st.columns([1.1, 1.2])
        with x1:
            safe_plot_bar(other_course_df.head(12), "service_name", "Students", height=340, horizontal=True)
        with x2:
            combo_other_show = combo_other[["phone_clean", "combo_order_date", "order_date", "service_name", "payment_received", "total_due", "status", "sales_rep_name"]].copy()
            combo_other_show = combo_other_show.rename(columns={"phone_clean": "Mobile", "combo_order_date": "Combo Order Date", "order_date": "Other Course Order Date", "service_name": "Other Course", "payment_received": "Paid", "total_due": "Due", "status": "Order Status", "sales_rep_name": "Sales Rep"})
            combo_other_show["Combo Order Date"] = pd.to_datetime(combo_other_show["Combo Order Date"], errors="coerce").dt.strftime("%d %b %Y")
            combo_other_show["Other Course Order Date"] = pd.to_datetime(combo_other_show["Other Course Order Date"], errors="coerce").dt.strftime("%d %b %Y")
            st.dataframe(combo_other_show, use_container_width=True, hide_index=True, height=340)
    else:
        st.info("No later non-combo course purchase found for the selected scope.")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="panel-title"><h3>Lead Intelligence Table</h3><p>Cross-check attendee → conversion → lead source path. Every displayed column now has its own filter box.</p></div>', unsafe_allow_html=True)
    lead_columns = [
        "NAME", "mobile_clean", "Seminar Date", "Place", "Session", "Trainer Norm", "service_name",
        "payment_received", "total_due", "order_date", "conversion_status", "lead_origin",
        "lead_source_name", "campaign_name", "leadstatus", "stage_name", "leadownername",
        "state", "Attempted/Unattempted", "servicename", "name", "email", "remarks"
    ]
    lead_table = filtered.copy()
    for col in lead_columns:
        if col not in lead_table.columns:
            lead_table[col] = ""
    lead_table = lead_table[lead_columns].copy()
    lead_table = lead_table.rename(columns={
        "NAME": "Student Name",
        "mobile_clean": "Mobile",
        "Trainer Norm": "Trainer",
        "service_name": "Course",
        "payment_received": "Paid",
        "total_due": "Due",
        "order_date": "Order Date",
        "conversion_status": "Conversion Status",
        "lead_origin": "Webinar/Non-Webinar",
        "lead_source_name": "Lead Source Name",
        "campaign_name": "Campaign Name",
        "leadstatus": "Lead Status",
        "stage_name": "Stage",
        "leadownername": "Lead Owner",
        "state": "State",
        "Attempted/Unattempted": "Attempted/Unattempted",
        "servicename": "Lead Service Name",
        "name": "Lead Name",
        "email": "Lead Email",
        "remarks": "Lead Remarks",
    })
    lead_table["Seminar Date"] = pd.to_datetime(lead_table["Seminar Date"], errors="coerce").dt.strftime("%d %b %Y")
    lead_table["Order Date"] = pd.to_datetime(lead_table["Order Date"], errors="coerce").dt.strftime("%d %b %Y")
    filtered_lead_table = apply_text_filters(lead_table, list(lead_table.columns), key_prefix="leadtable")
    st.dataframe(filtered_lead_table, use_container_width=True, hide_index=True, height=460)
    st.markdown('</div>', unsafe_allow_html=True)

with t3:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="panel-title"><h3>Student-level Records</h3><p>Review each attendee row with seminar, conversion, course, paid, due, and lead mapping.</p></div>', unsafe_allow_html=True)
    rec = filtered.copy()
    rec["Seminar Date"] = pd.to_datetime(rec["Seminar Date"], errors="coerce").dt.strftime("%d %b %Y")
    rec["order_date"] = pd.to_datetime(rec["order_date"], errors="coerce").dt.strftime("%d %b %Y")
    show_cols = [
        "NAME", "mobile_clean", "Place", "Seminar Date", "Session", "Trainer Norm",
        "conversion_status", "order_date", "service_name", "payment_received", "total_due",
        "status", "sales_rep_name", "payment_mode", "lead_origin", "lead_source_name",
        "campaign_name", "leadstatus", "stage_name", "leadownername", "state"
    ]
    show_cols = [c for c in show_cols if c in rec.columns]
    display_df = rec[show_cols].rename(columns={
        "NAME": "Student Name",
        "mobile_clean": "Mobile",
        "Trainer Norm": "Trainer",
        "order_date": "Order Date",
        "service_name": "Course",
        "payment_received": "Paid",
        "total_due": "Due",
        "status": "Order Status",
        "sales_rep_name": "Sales Rep",
        "payment_mode": "Payment Mode",
        "lead_origin": "Webinar/Non-Webinar",
        "lead_source_name": "Lead Source Name",
        "campaign_name": "Campaign Name",
        "leadstatus": "Lead Status",
        "stage_name": "Lead Stage",
        "leadownername": "Lead Owner",
    })
    st.dataframe(display_df, use_container_width=True, hide_index=True, height=520)

    export_bytes = io.BytesIO()
    rec.to_excel(export_bytes, index=False)
    st.download_button(
        "Download filtered records",
        data=export_bytes.getvalue(),
        file_name=f"offline_seminar_filtered_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
        type="primary",
    )
    st.markdown('</div>', unsafe_allow_html=True)
