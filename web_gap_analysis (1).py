import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

st.set_page_config(
    page_title="Stackly Analytics Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Custom CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Main background dark */
    .main { background-color: #0f1117; }
    .block-container { padding-top: 1rem; }

    /* ── WHITE SIDEBAR ── */
    [data-testid="stSidebar"] {
        background-color: #ffffff !important;
    }
    [data-testid="stSidebar"] * {
        color: #1a1a2e !important;
    }
    [data-testid="stSidebar"] .stRadio label {
        color: #1a1a2e !important;
        font-weight: 500;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {
        background-color: #f0f2f6;
        border-radius: 8px;
        padding: 6px 12px;
        margin-bottom: 4px;
        border: 1px solid #dde1ea;
        color: #1a1a2e !important;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:hover {
        background-color: #e2e6f0;
    }
    /* Active radio option highlight */
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label[data-checked="true"],
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] [aria-checked="true"] ~ label {
        background-color: #5b6af0 !important;
        color: #ffffff !important;
    }
    [data-testid="stSidebar"] hr {
        border-color: #dde1ea !important;
    }
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {
        color: #1a1a2e !important;
    }
    [data-testid="stSidebar"] .stCaption,
    [data-testid="stSidebar"] small,
    [data-testid="stSidebar"] .caption {
        color: #555577 !important;
    }
    /* Sidebar scrollbar */
    [data-testid="stSidebar"]::-webkit-scrollbar { width: 6px; }
    [data-testid="stSidebar"]::-webkit-scrollbar-track { background: #f0f2f6; }
    [data-testid="stSidebar"]::-webkit-scrollbar-thumb { background: #c0c4d6; border-radius: 3px; }

    /* ── METRIC CARDS (main area) ── */
    .metric-card {
        background: linear-gradient(135deg, #1e2130, #252a3a);
        border: 1px solid #2d3550;
        border-radius: 12px;
        padding: 18px 22px;
        margin-bottom: 10px;
    }
    .metric-title { color: #8b9dc3; font-size: 13px; font-weight: 500; margin-bottom: 4px; }
    .metric-value { color: #e8ecf4; font-size: 28px; font-weight: 700; }
    .metric-delta { font-size: 12px; margin-top: 4px; }
    .delta-pos { color: #4ade80; }
    .delta-neg { color: #f87171; }

    /* ── SECTION HEADERS ── */
    .section-header {
        color: #c5d0e8;
        font-size: 18px;
        font-weight: 600;
        border-left: 4px solid #5b6af0;
        padding-left: 10px;
        margin: 18px 0 12px 0;
    }

    /* Native streamlit metric widget */
    div[data-testid="metric-container"] {
        background: #1e2130;
        border-radius: 10px;
        padding: 10px;
    }

    /* Dataframe dark */
    .stDataFrame { background-color: #1e2130; }
</style>
""", unsafe_allow_html=True)

# ── Resolve File Path ─────────────────────────────────────────────────────────
FILE_NAME = "Stackly_PowerBI_Dataset.xlsx"

def find_file(filename):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidate = os.path.join(script_dir, filename)
    if os.path.exists(candidate):
        return candidate
    candidate = os.path.join(os.getcwd(), filename)
    if os.path.exists(candidate):
        return candidate
    home = os.path.expanduser("~")
    for folder in ["Downloads", "Desktop", "Documents", ""]:
        candidate = os.path.join(home, folder, filename) if folder else os.path.join(home, filename)
        if os.path.exists(candidate):
            return candidate
    return None

FILE_PATH = find_file(FILE_NAME)

# ── File Upload Fallback ──────────────────────────────────────────────────────
if FILE_PATH is None:
    st.warning("⚠️ **Stackly_PowerBI_Dataset.xlsx** was not found automatically.")
    st.info("👇 Please upload the file to continue.")
    uploaded = st.file_uploader("Upload Stackly_PowerBI_Dataset.xlsx", type=["xlsx"])
    if uploaded is None:
        st.stop()
    FILE_PATH = uploaded

# ── Load Data ─────────────────────────────────────────────────────────────────
@st.cache_data
def load_data(path):
    xl = pd.ExcelFile(path)
    return {s: xl.parse(s) for s in xl.sheet_names}

data       = load_data(FILE_PATH)
jobs       = data["Job_Listings"]
candidates = data["Candidates"]
incidents  = data["Cyber_Security_Incidents"]
clients    = data["Security_Clients"]
accounts   = data["Bank_Accounts"]
loans      = data["Loan_Applications"]
txns       = data["Transactions"]
kpi        = data["Monthly_KPI_Summary"]

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🔷 Stackly Analytics")
    st.markdown("---")
    domain = st.radio("Select Domain", ["🏠 Overview", "💼 Recruitment", "🛡️ Cyber Security", "🏦 Banking"])
    st.markdown("---")
    st.caption("Data last refreshed: May 2025")

# ═══════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════
CARD_BG   = "#1e2130"
PLOT_FONT = dict(family="Inter, sans-serif", color="#c5d0e8")
COLORS    = ["#5b6af0", "#22d3ee", "#a78bfa", "#34d399", "#fb923c", "#f472b6"]

def dark_fig(fig, height=350):
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor=CARD_BG,
        font=PLOT_FONT,
        height=height,
        margin=dict(l=10, r=10, t=30, b=10),
        legend=dict(bgcolor="rgba(0,0,0,0)"),
    )
    fig.update_xaxes(gridcolor="#2d3550", linecolor="#2d3550")
    fig.update_yaxes(gridcolor="#2d3550", linecolor="#2d3550")
    return fig

def metric_card(title, value, delta=None, pos=True):
    delta_html = ""
    if delta:
        cls   = "delta-pos" if pos else "delta-neg"
        arrow = "▲" if pos else "▼"
        delta_html = f'<div class="metric-delta {cls}">{arrow} {delta}</div>'
    return f"""
    <div class="metric-card">
        <div class="metric-title">{title}</div>
        <div class="metric-value">{value}</div>
        {delta_html}
    </div>"""

# ═══════════════════════════════════════════════════════════════════════════
# OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════
if domain == "🏠 Overview":
    st.markdown("# 📊 Stackly Platform — Executive Overview")
    st.caption("Consolidated KPIs across Recruitment · Cyber Security · Banking")

    col1, col2, col3, col4, col5 = st.columns(5)
    total_placements = kpi[kpi.Domain == "Recruitment"]["Placements_Made"].sum()
    total_incidents  = incidents.shape[0]
    active_clients   = clients[clients.Active == "Yes"].shape[0]
    total_loans      = f"₹{loans['Amount_INR'].sum()/1e7:.1f} Cr"
    total_revenue    = f"${kpi['Revenue_USD'].sum()/1e6:.2f}M"

    col1.markdown(metric_card("Total Placements",   int(total_placements), "+12% MoM"),                 unsafe_allow_html=True)
    col2.markdown(metric_card("Security Incidents", total_incidents,       "50 total tracked"),         unsafe_allow_html=True)
    col3.markdown(metric_card("Active Clients",     active_clients,        f"of {len(clients)} total"), unsafe_allow_html=True)
    col4.markdown(metric_card("Loans Portfolio",    total_loans,           "40 applications"),          unsafe_allow_html=True)
    col5.markdown(metric_card("Total Revenue",      total_revenue,         "+8% YoY"),                  unsafe_allow_html=True)

    st.markdown("---")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-header">Monthly Revenue by Domain</div>', unsafe_allow_html=True)
        rev = kpi.groupby(["Month", "Domain"])["Revenue_USD"].sum().reset_index()
        month_order = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
        rev["Month"] = pd.Categorical(rev["Month"], categories=month_order, ordered=True)
        rev = rev.sort_values("Month")
        fig = px.area(rev, x="Month", y="Revenue_USD", color="Domain", color_discrete_sequence=COLORS)
        st.plotly_chart(dark_fig(fig), use_container_width=True)

    with c2:
        st.markdown('<div class="section-header">Customer Satisfaction by Domain</div>', unsafe_allow_html=True)
        sat = kpi.groupby("Domain")["Customer_Satisfaction"].mean().reset_index()
        fig = px.bar(sat, x="Domain", y="Customer_Satisfaction", color="Domain",
                     color_discrete_sequence=COLORS, text_auto=".2f")
        fig.update_layout(showlegend=False, yaxis_range=[0, 5.5])
        st.plotly_chart(dark_fig(fig), use_container_width=True)

    st.markdown('<div class="section-header">Cross-Domain KPI Trends (2024–2025)</div>', unsafe_allow_html=True)
    kpi_yr = kpi[kpi.Year.isin([2024, 2025])].copy()
    kpi_yr["Period"] = kpi_yr["Month"].astype(str) + " " + kpi_yr["Year"].astype(str)

    rec  = kpi_yr[kpi_yr.Domain == "Recruitment"].groupby("Period")["Placements_Made"].sum().reset_index()
    cyb  = kpi_yr[kpi_yr.Domain == "CyberSecurity"].groupby("Period")["Incidents_Detected"].sum().reset_index()
    bank = kpi_yr[kpi_yr.Domain == "Banking"].groupby("Period")["Loan_Count"].sum().reset_index()

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=rec["Period"],  y=rec["Placements_Made"],    name="Placements", line=dict(color=COLORS[0])))
    fig.add_trace(go.Scatter(x=cyb["Period"],  y=cyb["Incidents_Detected"], name="Incidents",  line=dict(color=COLORS[2])))
    fig.add_trace(go.Scatter(x=bank["Period"], y=bank["Loan_Count"],        name="Loans",      line=dict(color=COLORS[3])))
    st.plotly_chart(dark_fig(fig, height=300), use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════
# RECRUITMENT
# ═══════════════════════════════════════════════════════════════════════════
elif domain == "💼 Recruitment":
    st.markdown("# 💼 Recruitment Analytics")

    c1, c2, c3, c4 = st.columns(4)
    open_jobs   = jobs[jobs.Status == "Open"].shape[0]
    total_apps  = candidates.shape[0]
    offers_made = candidates[candidates.Offer_Status == "Accepted"].shape[0]
    conv_rate   = f"{offers_made/total_apps*100:.1f}%"

    c1.markdown(metric_card("Open Positions",   open_jobs),                       unsafe_allow_html=True)
    c2.markdown(metric_card("Total Candidates", total_apps),                      unsafe_allow_html=True)
    c3.markdown(metric_card("Offers Accepted",  offers_made),                     unsafe_allow_html=True)
    c4.markdown(metric_card("Conversion Rate",  conv_rate, "Accepted / Applied"), unsafe_allow_html=True)

    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="section-header">Applications by Job Category</div>', unsafe_allow_html=True)
        cat_apps = (jobs.groupby("Category")["Applications_Count"]
                       .sum().reset_index()
                       .sort_values("Applications_Count", ascending=True))
        fig = px.bar(cat_apps, x="Applications_Count", y="Category", orientation="h",
                     color="Applications_Count", color_continuous_scale="Blues")
        fig.update_layout(coloraxis_showscale=False)
        st.plotly_chart(dark_fig(fig), use_container_width=True)

    with col2:
        st.markdown('<div class="section-header">Candidate Interview Pipeline</div>', unsafe_allow_html=True)
        pipe = candidates["Interview_Status"].value_counts().reset_index()
        pipe.columns = ["Stage", "Count"]
        fig = px.funnel(pipe, x="Count", y="Stage", color_discrete_sequence=COLORS)
        st.plotly_chart(dark_fig(fig), use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        st.markdown('<div class="section-header">Hiring Source Channels</div>', unsafe_allow_html=True)
        src = candidates["Source_Channel"].value_counts().reset_index()
        src.columns = ["Channel", "Count"]
        fig = px.pie(src, names="Channel", values="Count", hole=0.45, color_discrete_sequence=COLORS)
        fig.update_traces(textposition="outside", textinfo="percent+label")
        st.plotly_chart(dark_fig(fig), use_container_width=True)

    with col4:
        st.markdown('<div class="section-header">Salary Range by Job Type</div>', unsafe_allow_html=True)
        sal = jobs.copy()
        sal["Avg_Salary"] = (sal["Min_Salary_LPA"] + sal["Max_Salary_LPA"]) / 2
        fig = px.box(sal, x="Job_Type", y="Avg_Salary", color="Job_Type",
                     color_discrete_sequence=COLORS, points="all")
        fig.update_layout(showlegend=False)
        st.plotly_chart(dark_fig(fig), use_container_width=True)

    st.markdown('<div class="section-header">Top Jobs by Application Count</div>', unsafe_allow_html=True)
    top = jobs.nlargest(10, "Applications_Count")[
        ["Job_Title","Company","Category","Applications_Count","Status","Remote_OK"]
    ]
    st.dataframe(top.reset_index(drop=True), use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════
# CYBER SECURITY
# ═══════════════════════════════════════════════════════════════════════════
elif domain == "🛡️ Cyber Security":
    st.markdown("# 🛡️ Cyber Security Operations")

    c1, c2, c3, c4 = st.columns(4)
    total_inc     = len(incidents)
    critical_inc  = len(incidents[incidents.Severity == "Critical"])
    total_blocked = f"{incidents['Threats_Blocked'].sum()/1e6:.2f}M"
    fin_impact    = f"${incidents['Financial_Impact_USD'].sum()/1e6:.1f}M"

    c1.markdown(metric_card("Total Incidents",    total_inc),                                     unsafe_allow_html=True)
    c2.markdown(metric_card("Critical Incidents", critical_inc, "Needs urgent attention", False), unsafe_allow_html=True)
    c3.markdown(metric_card("Threats Blocked",    total_blocked),                                 unsafe_allow_html=True)
    c4.markdown(metric_card("Financial Impact",   fin_impact,   "Across all incidents",  False),  unsafe_allow_html=True)

    st.markdown("---")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-header">Incidents by Severity</div>', unsafe_allow_html=True)
        sev = incidents["Severity"].value_counts().reset_index()
        sev.columns = ["Severity", "Count"]
        fig = px.pie(sev, names="Severity", values="Count", hole=0.4,
                     color="Severity",
                     color_discrete_map={"Critical":"#f87171","High":"#fb923c","Medium":"#facc15","Low":"#4ade80"})
        fig.update_traces(textposition="outside", textinfo="percent+label")
        st.plotly_chart(dark_fig(fig), use_container_width=True)

    with c2:
        st.markdown('<div class="section-header">Threat Types Distribution</div>', unsafe_allow_html=True)
        thr = incidents["Threat_Type"].value_counts().reset_index()
        thr.columns = ["Threat", "Count"]
        fig = px.bar(thr, x="Count", y="Threat", orientation="h",
                     color="Count", color_continuous_scale="Reds")
        fig.update_layout(coloraxis_showscale=False)
        st.plotly_chart(dark_fig(fig), use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        st.markdown('<div class="section-header">Financial Impact by Industry</div>', unsafe_allow_html=True)
        fin = (incidents.groupby("Industry_Sector")["Financial_Impact_USD"]
                        .sum().reset_index()
                        .sort_values("Financial_Impact_USD", ascending=False))
        fig = px.bar(fin, x="Industry_Sector", y="Financial_Impact_USD",
                     color="Industry_Sector", color_discrete_sequence=COLORS)
        fig.update_layout(showlegend=False)
        st.plotly_chart(dark_fig(fig), use_container_width=True)

    with c4:
        st.markdown('<div class="section-header">Resolution Status</div>', unsafe_allow_html=True)
        res = incidents["Resolution_Status"].value_counts().reset_index()
        res.columns = ["Status", "Count"]
        fig = px.bar(res, x="Status", y="Count", color="Status", color_discrete_sequence=COLORS)
        fig.update_layout(showlegend=False)
        st.plotly_chart(dark_fig(fig), use_container_width=True)

    st.markdown('<div class="section-header">Response Time vs Financial Impact (by Severity)</div>', unsafe_allow_html=True)
    fig = px.scatter(incidents, x="Response_Time_Hours", y="Financial_Impact_USD",
                     color="Severity", size="Threats_Blocked",
                     color_discrete_map={"Critical":"#f87171","High":"#fb923c","Medium":"#facc15","Low":"#4ade80"},
                     hover_data=["Client_Company","Threat_Type"])
    st.plotly_chart(dark_fig(fig, height=400), use_container_width=True)

    st.markdown('<div class="section-header">Client Portfolio</div>', unsafe_allow_html=True)
    cl = clients[["Company_Name","Industry","Contract_Type","Annual_Value_USD","Satisfaction_Score","Active"]].copy()
    cl["Annual_Value_USD"] = cl["Annual_Value_USD"].apply(lambda x: f"${x:,.0f}")
    st.dataframe(cl.reset_index(drop=True), use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════
# BANKING
# ═══════════════════════════════════════════════════════════════════════════
elif domain == "🏦 Banking":
    st.markdown("# 🏦 Banking & Financial Analytics")

    c1, c2, c3, c4 = st.columns(4)
    active_acc     = accounts[accounts.Account_Status == "Active"].shape[0]
    total_bal      = f"₹{accounts['Balance_INR'].sum()/1e7:.1f} Cr"
    approved_loans = loans[loans.Status == "Approved"].shape[0]
    avg_credit     = f"{loans['Credit_Score'].mean():.0f}"

    c1.markdown(metric_card("Active Accounts",  active_acc,     f"of {len(accounts)} total"), unsafe_allow_html=True)
    c2.markdown(metric_card("Total AUM",        total_bal),                                   unsafe_allow_html=True)
    c3.markdown(metric_card("Approved Loans",   approved_loans, f"of {len(loans)} apps"),     unsafe_allow_html=True)
    c4.markdown(metric_card("Avg Credit Score", avg_credit),                                  unsafe_allow_html=True)

    st.markdown("---")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-header">Loan Application Status</div>', unsafe_allow_html=True)
        ls = loans["Status"].value_counts().reset_index()
        ls.columns = ["Status", "Count"]
        color_map = {"Approved":"#4ade80","Rejected":"#f87171","Pending":"#facc15","Under Review":"#60a5fa"}
        fig = px.pie(ls, names="Status", values="Count", hole=0.45,
                     color="Status", color_discrete_map=color_map)
        fig.update_traces(textposition="outside", textinfo="percent+label")
        st.plotly_chart(dark_fig(fig), use_container_width=True)

    with c2:
        st.markdown('<div class="section-header">Loan Amount by Type</div>', unsafe_allow_html=True)
        lt = loans.groupby("Loan_Type")["Amount_INR"].sum().reset_index().sort_values("Amount_INR")
        fig = px.bar(lt, x="Amount_INR", y="Loan_Type", orientation="h",
                     color="Amount_INR", color_continuous_scale="Blues", text_auto=".2s")
        fig.update_layout(coloraxis_showscale=False)
        st.plotly_chart(dark_fig(fig), use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        st.markdown('<div class="section-header">Account Balance by Segment</div>', unsafe_allow_html=True)
        seg = accounts.groupby("Segment")["Balance_INR"].sum().reset_index()
        fig = px.bar(seg, x="Segment", y="Balance_INR", color="Segment",
                     color_discrete_sequence=COLORS, text_auto=".2s")
        fig.update_layout(showlegend=False)
        st.plotly_chart(dark_fig(fig), use_container_width=True)

    with c4:
        st.markdown('<div class="section-header">Transaction Channels Distribution</div>', unsafe_allow_html=True)
        ch = txns["Channel"].value_counts().reset_index()
        ch.columns = ["Channel", "Count"]
        fig = px.pie(ch, names="Channel", values="Count", hole=0.45, color_discrete_sequence=COLORS)
        fig.update_traces(textposition="outside", textinfo="percent+label")
        st.plotly_chart(dark_fig(fig), use_container_width=True)

    c5, c6 = st.columns(2)
    with c5:
        st.markdown('<div class="section-header">Credit Score vs Loan Amount</div>', unsafe_allow_html=True)
        fig = px.scatter(loans, x="Credit_Score", y="Amount_INR", color="Status",
                         size="EMI_INR", hover_data=["Customer_Name","Loan_Type"],
                         color_discrete_map={"Approved":"#4ade80","Rejected":"#f87171",
                                             "Pending":"#facc15","Under Review":"#60a5fa"})
        st.plotly_chart(dark_fig(fig, height=320), use_container_width=True)

    with c6:
        st.markdown('<div class="section-header">Transactions by Category</div>', unsafe_allow_html=True)
        tc = (txns.groupby("Category")["Amount_INR"]
                  .sum().reset_index()
                  .sort_values("Amount_INR", ascending=False))
        fig = px.bar(tc, x="Category", y="Amount_INR", color="Category",
                     color_discrete_sequence=COLORS)
        fig.update_layout(showlegend=False)
        st.plotly_chart(dark_fig(fig, height=320), use_container_width=True)

    st.markdown('<div class="section-header">KYC Status & Account Health</div>', unsafe_allow_html=True)
    kyc = accounts.groupby(["KYC_Status","Account_Status"]).size().reset_index(name="Count")
    fig = px.bar(kyc, x="KYC_Status", y="Count", color="Account_Status",
                 barmode="group", color_discrete_sequence=COLORS)
    st.plotly_chart(dark_fig(fig, height=300), use_container_width=True)

    st.markdown('<div class="section-header">Recent Transactions</div>', unsafe_allow_html=True)
    txn_view = txns[["Txn_ID","Account_ID","Date","Type","Channel","Amount_INR",
                      "Category","Merchant","Status"]].head(15)
    st.dataframe(txn_view.reset_index(drop=True), use_container_width=True)
