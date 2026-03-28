import io
import warnings
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

warnings.filterwarnings('ignore')

st.set_page_config(
    page_title='Seminar Intelligence',
    page_icon='📊',
    layout='wide',
    initial_sidebar_state='collapsed',
)

st.markdown(
    """
    <style>
      #MainMenu, footer, header {visibility:hidden}
      .block-container {padding-top:0.8rem; padding-bottom:2rem; max-width:100% !important}
      .stApp {background:#060910; color:#eceef5}
      .card {
        background:#0c1018;
        border:1px solid rgba(255,255,255,.07);
        border-radius:14px;
        padding:16px;
        margin-bottom:12px;
      }
      .kpi {
        background:#111520;
        border:1px solid rgba(255,255,255,.07);
        border-radius:12px;
        padding:14px;
      }
      .muted {color:#8a90aa; font-size:12px}
      .title {font-size:28px; font-weight:800; margin-bottom:4px}
      .section-title {font-size:13px; font-weight:700; color:#f7c948; letter-spacing:.08em; text-transform:uppercase; margin-bottom:12px}
      div[data-testid="metric-container"] {
        background:#111520 !important;
        border:1px solid rgba(255,255,255,.07) !important;
        border-radius:12px !important;
      }
      .stTabs [data-baseweb="tab-list"] {gap:0.5rem}
      .stTabs [data-baseweb="tab"] {background:transparent; border-radius:8px; color:#8a90aa}
      .stTabs [aria-selected="true"] {color:#4fce8f}
    </style>
    """,
    unsafe_allow_html=True,
)

BG = '#0c1018'
GRID = 'rgba(255,255,255,0.07)'
TXT = '#8a90aa'
GREEN = '#4fce8f'
BLUE = '#4f8ef7'
RED = '#f76f4f'
AMBER = '#f7c948'
PURPLE = '#b44fe7'
TEAL = '#4fd8f7'
GRAY = '#596275'
PLOT_BASE = dict(
    paper_bgcolor=BG,
    plot_bgcolor=BG,
    font=dict(color=TXT, size=11),
    margin=dict(l=20, r=20, t=30, b=20),
)
PM = {
    'mode1': 'Full Payment',
    'mode2': 'Instalment',
    'mode3': 'EMI',
    'mode4': 'Partial',
    'mode13': 'Scholarship',
    'mode5': 'Other',
}


def fmt_inr(value):
    if pd.isna(value):
        return '₹0'
    value = float(value)
    if abs(value) >= 1e7:
        return f'₹{value/1e7:.2f}Cr'
    if abs(value) >= 1e5:
        return f'₹{value/1e5:.2f}L'
    return f'₹{value:,.0f}'


def ensure_columns(df, defaults):
    for col, default in defaults.items():
        if col not in df.columns:
            df[col] = default
    return df


def clean_phone(series):
    return series.astype(str).str.replace(r'\D', '', regex=True).str[-10:]


def normalize_bool_text(series, true_values=None):
    true_values = set(true_values or ['YES', 'TRUE', '1', 'CONVERTED'])
    s = series.astype(str).str.strip().str.upper()
    return s.isin(true_values)


@st.cache_data(show_spinner=False)
def load_all(seminar_bytes, conversion_bytes, leads_bytes):
    # Seminar
    sem = pd.read_csv(io.BytesIO(seminar_bytes))
    sem.columns = sem.columns.str.strip()
    sem = ensure_columns(
        sem,
        {
            'NAME': '',
            'Mobile': '',
            'Place': 'Unknown',
            'Trainer / Presenter': 'Unknown',
            'Seminar Date': pd.NaT,
            'Session': 'Unknown',
            'Is Attended ?': 'NO',
            'Is Converted ?': 'NO',
            'Amount Paid': 0,
            'Mode of Payment': 'Unknown',
            'TRADER': 'NO',
            'Is our Student ?': 'NO',
            'Remarks': '',
        },
    )
    for col in ['Place', 'Trainer / Presenter', 'Session', 'TRADER', 'Is our Student ?', 'Mode of Payment']:
        sem[col] = sem[col].astype(str).str.strip()
    sem['Seminar Date'] = pd.to_datetime(sem['Seminar Date'], errors='coerce', dayfirst=True)
    sem['Amount Paid'] = pd.to_numeric(sem['Amount Paid'], errors='coerce').fillna(0)
    sem['mob'] = clean_phone(sem['Mobile'])
    trainer_map = {
        'HIRAMNOY LAHERI/PRATIM CHAKRABORTY': 'HIRANMOY LAHIRI & PRATIM CHAKRABORTY',
        'PRATIM KUMAR CHAKRABORTY/MIHIR KANTI CHAKRABORTY': 'PRATIM CHAKRABORTY & MIHIR KANTI CHAKRABORTY',
        'PRATIM KUMAR CHAKRABORTY & AKASH MISHRA': 'PRATIM CHAKRABORTY & AKASH MISHRA',
        'PRFATIM CHAKRABORTY & AKASH MISHRA': 'PRATIM CHAKRABORTY & AKASH MISHRA',
        'PRATIM CHAKRABORTY & MIHIR KANTI CHAKRABORTYPARIJAT': 'PRATIM CHAKRABORTY & MIHIR KANTI CHAKRABORTY',
    }
    sem['Trainer Norm'] = sem['Trainer / Presenter'].astype(str).str.upper().replace(trainer_map)
    sem['Attended'] = normalize_bool_text(sem['Is Attended ?'])
    sem['Converted Flag'] = normalize_bool_text(sem['Is Converted ?'])
    attended = sem[sem['Attended']].copy().reset_index(drop=True)
    attended['Conv Status'] = attended['Converted Flag'].map({True: 'Converted', False: 'Not Converted'})

    # Conversion
    conv = pd.read_excel(io.BytesIO(conversion_bytes))
    conv.columns = conv.columns.str.strip()
    conv = ensure_columns(
        conv,
        {
            '_id': range(1, len(conv) + 1),
            'phone': '',
            'payment_received': 0,
            'total_amount': 0,
            'total_due': 0,
            'total_gst': 0,
            'order_date': pd.NaT,
            'payment_mode': 'Unknown',
            'service_name': 'Unknown',
            'sales_rep_name': 'Unknown',
            'trainer': 'Unknown',
            'status': 'Unknown',
            'is_refunded': False,
            'is_shortClosed': False,
            'student_name': '',
            'email': '',
            'batch_date': pd.NaT,
            'orderID': '',
            'service_code': '',
            'student_invid': '',
        },
    )
    for col in ['payment_received', 'total_amount', 'total_due', 'total_gst']:
        conv[col] = pd.to_numeric(conv[col], errors='coerce').fillna(0)
    conv['order_date'] = pd.to_datetime(conv['order_date'], errors='coerce', utc=True).dt.tz_localize(None)
    conv['batch_date'] = pd.to_datetime(conv['batch_date'], errors='coerce', utc=True).dt.tz_localize(None)
    conv['phone_clean'] = clean_phone(conv['phone'])
    conv['PM Label'] = conv['payment_mode'].map(PM).fillna(conv['payment_mode'].astype(str))
    conv['service_name'] = conv['service_name'].astype(str).str.strip()
    conv['sales_rep_name'] = conv['sales_rep_name'].astype(str).str.strip()
    conv['trainer_name'] = conv['trainer'].astype(str).str.split(' - ').str[-1].str.strip()
    conv['month'] = conv['order_date'].dt.to_period('M').astype(str)
    conv['is_refunded'] = conv['is_refunded'].fillna(False).astype(bool)
    conv['is_shortClosed'] = conv['is_shortClosed'].fillna(False).astype(bool)

    # Leads
    leads_xf = pd.ExcelFile(io.BytesIO(leads_bytes))
    sheet = 'Sheet 1' if 'Sheet 1' in leads_xf.sheet_names else leads_xf.sheet_names[0]
    leads = pd.read_excel(leads_xf, sheet_name=sheet)
    leads.columns = leads.columns.str.strip()
    leads = ensure_columns(
        leads,
        {
            '_id': range(1, len(leads) + 1),
            'name': '',
            'phone': '',
            'email': '',
            'leaddate': pd.NaT,
            'converted_from': 'Unknown',
            'leadsource': 'Unknown',
            'campaign_name': 'Unknown',
            'leadstatus': 'Unknown',
            'stage_name': 'Unknown',
            'leadownername': 'Unknown',
            'state': 'Unknown',
            'Attempted/Unattempted': 'Unknown',
            'servicename': 'Unknown',
            'remarks': '',
        },
    )
    leads['leaddate'] = pd.to_datetime(leads['leaddate'], errors='coerce')
    for col in ['converted_from', 'leadsource', 'campaign_name', 'leadstatus', 'stage_name', 'leadownername', 'state', 'Attempted/Unattempted', 'servicename']:
        leads[col] = leads[col].astype(str).str.strip().replace({'nan': 'Unknown', 'None': 'Unknown', '': 'Unknown'})
    leads['Attempted/Unattempted'] = leads['Attempted/Unattempted'].replace({
        'attempted': 'Attempted',
        'unattempted': 'Unattempted',
        'ATTEMPTED': 'Attempted',
        'UNATTEMPTED': 'Unattempted',
    })
    leads['phone_clean'] = clean_phone(leads['phone'])
    leads['Lead Month'] = leads['leaddate'].dt.to_period('M').astype(str)
    leads['Is Converted'] = leads['leadstatus'].astype(str).str.lower().str.contains('converted|seat booked|sales closed', na=False)

    # Merge
    order_cols = [
        'phone_clean', 'orderID', 'order_date', 'service_code', 'service_name', 'payment_received',
        'total_amount', 'total_due', 'total_gst', 'payment_mode', 'PM Label', 'status',
        'sales_rep_name', 'trainer', 'trainer_name', 'student_invid', 'batch_date', 'month',
        'is_refunded', 'is_shortClosed'
    ]
    att = attended.merge(conv[order_cols], left_on='mob', right_on='phone_clean', how='left')

    def due_status(row):
        if pd.isna(row.get('total_due')):
            return 'No Order'
        if float(row.get('total_due', 0)) <= 0:
            return 'Fully Paid'
        if float(row.get('total_amount', 0)) > float(row.get('total_due', 0)):
            return 'Partially Paid'
        return 'Fully Due'

    att['Due Status'] = att.apply(due_status, axis=1)
    lead_cols = ['phone_clean', 'converted_from', 'leadsource', 'campaign_name', 'leadstatus', 'stage_name', 'leadownername', 'state', 'Attempted/Unattempted']
    att = att.merge(leads[lead_cols].drop_duplicates('phone_clean'), left_on='mob', right_on='phone_clean', how='left', suffixes=('', '_lead'))
    for col in ['converted_from', 'leadsource', 'campaign_name', 'leadstatus', 'stage_name', 'leadownername', 'state', 'Attempted/Unattempted']:
        att[col] = att[col].fillna('Unknown')

    return {'sem': sem, 'att': att, 'conv': conv, 'leads': leads}


def section(title):
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)


def no_data(msg='Upload all 3 files to continue.'):
    st.markdown(f'<div class="card"><div class="title">No data loaded</div><div class="muted">{msg}</div></div>', unsafe_allow_html=True)


def render_kpis(items, cols=4):
    row = st.columns(cols)
    for i, item in enumerate(items):
        with row[i % cols]:
            st.metric(item['label'], item['value'], item.get('delta'))


def apply_att_filters(df, prefix='att'):
    with st.expander('Filters', expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        c5, c6, c7, c8 = st.columns(4)
        places = c1.multiselect('Location', sorted(df['Place'].dropna().astype(str).unique()), key=f'{prefix}_p')
        trainers = c2.multiselect('Trainer', sorted(df['Trainer Norm'].dropna().astype(str).unique()), key=f'{prefix}_t')
        sessions = c3.multiselect('Session', sorted(df['Session'].dropna().astype(str).unique()), key=f'{prefix}_s')
        conv_status = c4.multiselect('Conversion', sorted(df['Conv Status'].dropna().astype(str).unique()), key=f'{prefix}_c')
        due_status = c5.multiselect('Due Status', sorted(df['Due Status'].dropna().astype(str).unique()), key=f'{prefix}_d')
        courses = c6.multiselect('Course', sorted(df['service_name'].dropna().astype(str).unique()), key=f'{prefix}_co')
        reps = c7.multiselect('Sales Rep', sorted(df['sales_rep_name'].dropna().astype(str).unique()), key=f'{prefix}_r')
        trader = c8.selectbox('Trader?', ['All', 'Yes', 'No'], key=f'{prefix}_tr')
    out = df.copy()
    if places:
        out = out[out['Place'].isin(places)]
    if trainers:
        out = out[out['Trainer Norm'].isin(trainers)]
    if sessions:
        out = out[out['Session'].isin(sessions)]
    if conv_status:
        out = out[out['Conv Status'].isin(conv_status)]
    if due_status:
        out = out[out['Due Status'].isin(due_status)]
    if courses:
        out = out[out['service_name'].isin(courses)]
    if reps:
        out = out[out['sales_rep_name'].isin(reps)]
    if trader == 'Yes':
        out = out[out['TRADER'].astype(str).str.upper().isin(['YES', 'TYES'])]
    elif trader == 'No':
        out = out[~out['TRADER'].astype(str).str.upper().isin(['YES', 'TYES'])]
    return out


def plot_bar(df, x, y, color=None, horizontal=False, height=320, title=None):
    if df is None or df.empty:
        st.info('No data for this chart.')
        return
    data = df.copy()
    if x not in data.columns or y not in data.columns:
        st.info(f'Chart skipped because required columns are missing: {x}, {y}')
        return
    if color and color not in data.columns:
        color = None
    x_col = y if horizontal else x
    y_col = x if horizontal else y
    if x_col not in data.columns or y_col not in data.columns:
        st.info('Chart skipped because the selected axes are unavailable.')
        return
    data = data[[c for c in [x_col, y_col, color] if c and c in data.columns]].copy()
    if horizontal:
        data[x_col] = pd.to_numeric(data[x_col], errors='coerce')
        data = data.dropna(subset=[x_col])
    else:
        data[y_col] = pd.to_numeric(data[y_col], errors='coerce')
        data = data.dropna(subset=[y_col])
    if data.empty:
        st.info('No valid numeric values available for this chart.')
        return
    fig = px.bar(
        data,
        x=x_col,
        y=y_col,
        color=color,
        orientation='h' if horizontal else 'v',
        title=title,
    )
    fig.update_layout(**PLOT_BASE, height=height)
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor=GRID)
    st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})


def plot_donut(labels, values, height=280):
    fig = go.Figure(go.Pie(labels=labels, values=values, hole=0.6, textinfo='label+percent'))
    fig.update_layout(**PLOT_BASE, height=height, showlegend=True)
    st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})


st.markdown(
    f"""
    <div class="card" style="padding:18px 20px; margin-bottom:14px;">
        <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;">
            <div>
                <div class="title">Invesmate Seminar Intelligence Hub</div>
                <div class="muted">Upload 3 files, auto-merge records, and explore seminar, leads, and conversion analytics.</div>
            </div>
            <div class="muted">{datetime.now().strftime('%d %b %Y %H:%M')}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

for key, default in [('loaded', False), ('data', None)]:
    if key not in st.session_state:
        st.session_state[key] = default


tab_upload, tab_overview, tab_seminar, tab_leads, tab_conversion, tab_records = st.tabs(
    ['📁 Upload', '🏠 Overview', '🎯 Seminar', '📣 Leads', '💰 Conversion', '📋 Records']
)

with tab_upload:
    c1, c2, c3 = st.columns(3)
    with c1:
        seminar_file = st.file_uploader('Seminar CSV', type=['csv'])
    with c2:
        conversion_file = st.file_uploader('Conversion XLSX', type=['xlsx', 'xls'])
    with c3:
        leads_file = st.file_uploader('Leads XLSX', type=['xlsx', 'xls'])

    if seminar_file and conversion_file and leads_file:
        if st.button('Generate dashboard', type='primary', use_container_width=True):
            with st.spinner('Processing files...'):
                st.session_state.data = load_all(seminar_file.read(), conversion_file.read(), leads_file.read())
                st.session_state.loaded = True
            st.success('Dashboard ready.')
            att = st.session_state.data['att']
            conv = st.session_state.data['conv']
            leads = st.session_state.data['leads']
            render_kpis([
                {'label': 'Attendees', 'value': f"{len(att):,}"},
                {'label': 'Converted', 'value': f"{(att['Conv Status'] == 'Converted').sum():,}"},
                {'label': 'Orders', 'value': f"{len(conv):,}"},
                {'label': 'Leads', 'value': f"{len(leads):,}"},
            ], cols=4)
    else:
        no_data('Please upload Seminar CSV, Conversion XLSX, and Leads XLSX.')


if not st.session_state.loaded:
    with tab_overview:
        no_data()
    with tab_seminar:
        no_data()
    with tab_leads:
        no_data()
    with tab_conversion:
        no_data()
    with tab_records:
        no_data()
else:
    D = st.session_state.data
    sem = D['sem']
    att = D['att']
    conv = D['conv']
    leads = D['leads']

    with tab_overview:
        conv_count = (att['Conv Status'] == 'Converted').sum()
        conv_rate = (conv_count / len(att) * 100) if len(att) else 0
        render_kpis([
            {'label': 'Seminar Attendees', 'value': f'{len(att):,}'},
            {'label': 'Converted', 'value': f'{conv_count:,}', 'delta': f'{conv_rate:.1f}% rate'},
            {'label': 'Gross Revenue', 'value': fmt_inr(conv['total_amount'].sum())},
            {'label': 'Collected', 'value': fmt_inr(conv['payment_received'].sum())},
            {'label': 'Due', 'value': fmt_inr(conv['total_due'].sum())},
            {'label': 'Total Leads', 'value': f'{len(leads):,}'},
            {'label': 'Webinar Leads', 'value': f"{(leads['converted_from'] == 'Webinar').sum():,}"},
            {'label': 'Locations', 'value': f"{att['Place'].nunique():,}"},
        ], cols=4)

        col1, col2 = st.columns(2)
        with col1:
            section('Seminar funnel')
            fig = go.Figure(go.Funnel(y=['Registered', 'Attended', 'Converted'], x=[len(sem), len(att), conv_count]))
            fig.update_layout(**PLOT_BASE, height=320)
            st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
        with col2:
            section('Due status')
            due_counts = att['Due Status'].value_counts()
            plot_donut(due_counts.index, due_counts.values, height=320)

        c3, c4 = st.columns(2)
        with c3:
            section('Monthly revenue')
            monthly = conv.groupby('month', dropna=False).agg(Orders=('_id', 'count'), Revenue=('total_amount', 'sum')).reset_index()
            monthly = monthly[monthly['month'].notna() & (monthly['month'] != 'NaT')]
            if not monthly.empty:
                fig = make_subplots(specs=[[{'secondary_y': True}]])
                fig.add_trace(go.Bar(x=monthly['month'], y=monthly['Orders'], name='Orders', marker_color=BLUE), secondary_y=False)
                fig.add_trace(go.Scatter(x=monthly['month'], y=monthly['Revenue'], name='Revenue', mode='lines+markers', line=dict(color=GREEN, width=3)), secondary_y=True)
                fig.update_layout(**PLOT_BASE, height=320)
                fig.update_yaxes(gridcolor=GRID, secondary_y=False)
                fig.update_yaxes(showgrid=False, secondary_y=True)
                st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
            else:
                st.info('No monthly revenue data available.')
        with c4:
            section('Lead source mix')
            src = leads['converted_from'].value_counts()
            plot_donut(src.index, src.values, height=320)

    with tab_seminar:
        filt = apply_att_filters(att, 'seminar')
        converted = (filt['Conv Status'] == 'Converted').sum()
        rate = (converted / len(filt) * 100) if len(filt) else 0
        render_kpis([
            {'label': 'Attended', 'value': f'{len(filt):,}'},
            {'label': 'Converted', 'value': f'{converted:,}', 'delta': f'{rate:.1f}% rate'},
            {'label': 'Order Revenue', 'value': fmt_inr(filt['total_amount'].sum())},
            {'label': 'Collected', 'value': fmt_inr(filt['payment_received'].sum())},
        ], cols=4)

        c1, c2 = st.columns(2)
        with c1:
            section('Location performance')
            loc = filt.groupby('Place').agg(Attended=('NAME', 'count'), Converted=('Conv Status', lambda s: (s == 'Converted').sum())).reset_index()
            plot_bar(loc.sort_values('Attended', ascending=False).head(15), 'Place', 'Attended', height=340)
        with c2:
            section('Trainer conversion')
            tr = filt.groupby('Trainer Norm').agg(Attended=('NAME', 'count'), Converted=('Conv Status', lambda s: (s == 'Converted').sum())).reset_index()
            tr['Rate'] = (tr['Converted'] / tr['Attended'] * 100).fillna(0)
            plot_bar(tr.sort_values('Rate', ascending=True).tail(15), 'Rate', 'Trainer Norm', horizontal=True, height=340)

        c3, c4 = st.columns(2)
        with c3:
            section('Session split')
            ses = filt['Session'].value_counts()
            plot_donut(ses.index, ses.values, height=300)
        with c4:
            section('Payment mode at seminar')
            pm = filt['Mode of Payment'].fillna('Unknown').astype(str).value_counts()
            plot_donut(pm.index, pm.values, height=300)

    with tab_leads:
        with st.expander('Filters', expanded=True):
            l1, l2, l3, l4 = st.columns(4)
            src_sel = l1.multiselect('Converted from', sorted(leads['converted_from'].unique()), key='leads_src')
            ls_sel = l2.multiselect('Lead source', sorted(leads['leadsource'].unique()), key='leads_ls')
            owner_sel = l3.multiselect('Lead owner', sorted(leads['leadownername'].unique()), key='leads_owner')
            attempt_sel = l4.selectbox('Attempted?', ['All', 'Attempted', 'Unattempted'], key='leads_attempt')
        ld = leads.copy()
        if src_sel:
            ld = ld[ld['converted_from'].isin(src_sel)]
        if ls_sel:
            ld = ld[ld['leadsource'].isin(ls_sel)]
        if owner_sel:
            ld = ld[ld['leadownername'].isin(owner_sel)]
        if attempt_sel != 'All':
            ld = ld[ld['Attempted/Unattempted'] == attempt_sel]

        lead_conv = ld['Is Converted'].sum()
        lead_rate = (lead_conv / len(ld) * 100) if len(ld) else 0
        render_kpis([
            {'label': 'Total Leads', 'value': f'{len(ld):,}'},
            {'label': 'Converted', 'value': f'{lead_conv:,}', 'delta': f'{lead_rate:.1f}% rate'},
            {'label': 'Campaigns', 'value': f"{ld['campaign_name'].nunique():,}"},
            {'label': 'Owners', 'value': f"{ld['leadownername'].nunique():,}"},
        ], cols=4)

        c1, c2 = st.columns(2)
        with c1:
            section('Lead stage pipeline')
            stage = ld['stage_name'].value_counts().head(12)
            plot_bar(stage.rename_axis('Stage').reset_index(name='Count'), 'Stage', 'Count', height=340)
        with c2:
            section('Attempted vs unattempted')
            attm = ld['Attempted/Unattempted'].value_counts()
            plot_donut(attm.index, attm.values, height=340)

        c3, c4 = st.columns(2)
        with c3:
            section('Top campaigns')
            camp = ld['campaign_name'].value_counts().head(15)
            plot_bar(camp.rename_axis('Campaign').reset_index(name='Count'), 'Count', 'Campaign', horizontal=True, height=360)
        with c4:
            section('Lead owners')
            own = ld['leadownername'].value_counts().head(15)
            plot_bar(own.rename_axis('Owner').reset_index(name='Count'), 'Count', 'Owner', horizontal=True, height=360)

    with tab_conversion:
        with st.expander('Filters', expanded=True):
            c1, c2, c3, c4 = st.columns(4)
            st_sel = c1.multiselect('Status', sorted(conv['status'].astype(str).unique()), key='conv_st')
            sv_sel = c2.multiselect('Service', sorted(conv['service_name'].astype(str).unique()), key='conv_sv')
            rep_sel = c3.multiselect('Sales rep', sorted(conv['sales_rep_name'].astype(str).unique()), key='conv_rep')
            pm_sel = c4.multiselect('Payment mode', sorted(conv['PM Label'].astype(str).unique()), key='conv_pm')
        cv = conv.copy()
        if st_sel:
            cv = cv[cv['status'].isin(st_sel)]
        if sv_sel:
            cv = cv[cv['service_name'].isin(sv_sel)]
        if rep_sel:
            cv = cv[cv['sales_rep_name'].isin(rep_sel)]
        if pm_sel:
            cv = cv[cv['PM Label'].isin(pm_sel)]

        render_kpis([
            {'label': 'Orders', 'value': f'{len(cv):,}'},
            {'label': 'Revenue', 'value': fmt_inr(cv['total_amount'].sum())},
            {'label': 'Collected', 'value': fmt_inr(cv['payment_received'].sum())},
            {'label': 'Due', 'value': fmt_inr(cv['total_due'].sum())},
            {'label': 'Refunded', 'value': f"{int(cv['is_refunded'].sum()):,}"},
            {'label': 'Short Closed', 'value': f"{int(cv['is_shortClosed'].sum()):,}"},
            {'label': 'Services', 'value': f"{cv['service_name'].nunique():,}"},
            {'label': 'Sales Reps', 'value': f"{cv['sales_rep_name'].nunique():,}"},
        ], cols=4)

        c1, c2 = st.columns(2)
        with c1:
            section('Revenue by month')
            month_df = cv.groupby('month', dropna=False).agg(Revenue=('total_amount', 'sum')).reset_index()
            month_df = month_df[month_df['month'].notna() & (month_df['month'] != 'NaT')]
            plot_bar(month_df, 'month', 'Revenue', height=320)
        with c2:
            section('Order status')
            st_counts = cv['status'].value_counts()
            plot_donut(st_counts.index, st_counts.values, height=320)

        c3, c4 = st.columns(2)
        with c3:
            section('Top services')
            svc = cv.groupby('service_name').agg(Orders=('_id', 'count'), Revenue=('total_amount', 'sum')).reset_index().sort_values('Orders', ascending=False).head(15)
            plot_bar(svc, 'Orders', 'service_name', horizontal=True, height=360)
        with c4:
            section('Top sales reps')
            reps = cv.groupby('sales_rep_name').agg(Orders=('_id', 'count'), Revenue=('total_amount', 'sum')).reset_index().sort_values('Revenue', ascending=False).head(15)
            plot_bar(reps, 'Revenue', 'sales_rep_name', horizontal=True, height=360)

    with tab_records:
        source = st.radio('View records from', ['Seminar Attendance', 'Leads', 'Conversion Orders'], horizontal=True)

        def to_excel_bytes(df):
            buf = io.BytesIO()
            df.to_excel(buf, index=False)
            return buf.getvalue()

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

        if source == 'Seminar Attendance':
            df = att.copy()
            show_cols = [c for c in [
                'NAME', 'Mobile', 'Place', 'Seminar Date', 'Session', 'Trainer Norm', 'Conv Status',
                'Amount Paid', 'Mode of Payment', 'service_name', 'total_amount', 'payment_received',
                'total_due', 'Due Status', 'status', 'sales_rep_name', 'trainer_name', 'converted_from',
                'leadsource', 'Remarks'
            ] if c in df.columns]
            st.dataframe(df[show_cols], use_container_width=True, height=480, hide_index=True)
            st.download_button('Download seminar records', data=to_excel_bytes(df[show_cols]), file_name=f'seminar_records_{timestamp}.xlsx')
        elif source == 'Leads':
            df = leads.copy()
            show_cols = [c for c in [
                'name', 'phone', 'email', 'leaddate', 'converted_from', 'leadsource', 'campaign_name',
                'leadstatus', 'stage_name', 'leadownername', 'state', 'Attempted/Unattempted', 'servicename', 'remarks'
            ] if c in df.columns]
            st.dataframe(df[show_cols], use_container_width=True, height=480, hide_index=True)
            st.download_button('Download lead records', data=to_excel_bytes(df[show_cols]), file_name=f'lead_records_{timestamp}.xlsx')
        else:
            df = conv.copy()
            show_cols = [c for c in [
                'student_name', 'phone', 'email', 'order_date', 'service_name', 'total_amount',
                'payment_received', 'total_due', 'status', 'PM Label', 'sales_rep_name', 'trainer_name', 'batch_date'
            ] if c in df.columns]
            st.dataframe(df[show_cols], use_container_width=True, height=480, hide_index=True)
            st.download_button('Download conversion records', data=to_excel_bytes(df[show_cols]), file_name=f'conversion_records_{timestamp}.xlsx')
