import base64
import hashlib
import json
import os
import pymssql
import warnings
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import pymssql
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Suppress pandas UserWarning regarding raw DBAPI2/pyodbc connections
warnings.filterwarnings(
    "ignore",
    category=UserWarning,
    message=".*pandas only supports SQLAlchemy connectable.*"
)

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="Haldiram's D365 Sales Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# CACHE & LOCAL CONFIGURATION PATHS
# ============================================================
CACHE_DIR = Path(r"C:\DailySalesAutomation\Cache\DailySales").resolve()
CACHE_DIR.mkdir(parents=True, exist_ok=True)

PERMISSIONS_FILE = Path(r"C:\DailySalesAutomation\Cache\user_permissions.json").resolve()
SCRIPT_DIR = Path(r"C:\DailySalesAutomation\Reports\Dashboard\KPI").resolve()

BOOTSTRAP_ADMINS = ["admin", "vijay", "teddiwar"]
DEFAULT_ADMIN_PASSWORD = "admin"

# ============================================================
# LOGO CONFIGURATION & LOADER
# ============================================================
LOGO_FILE_NAMES = [
    "haldiram_logo.png",
    "haldiram_logo.jpg",
    "haldiram_logo.jpeg",
    "haldiram_logo.webp",
    "haldiram_logo.svg",
]

LOGO_MIME_TYPES = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
    ".svg": "image/svg+xml",
}

def load_logo_html():
    for file_name in LOGO_FILE_NAMES:
        candidate = SCRIPT_DIR / file_name
        if candidate.is_file():
            mime_type = LOGO_MIME_TYPES.get(candidate.suffix.lower(), "image/png")
            encoded = base64.b64encode(candidate.read_bytes()).decode()
            return (
                f'<img class="animated-brand-logo" '
                f'src="data:{mime_type};base64,{encoded}" alt="Haldiram\'s Logo" '
                f'style="max-width: 150px; max-height: 80px; width: auto; height: auto; '
                f'display: block; margin: 0 auto; animation: logoFloat 3.5s ease-in-out infinite;" />'
            )
    return (
        '<span class="animated-brand-logo" '
        'style="font-size:1.8rem; font-weight:800; color:#DC2626; letter-spacing:0.03em; '
        'display:inline-block; animation: logoFloat 3.5s ease-in-out infinite;">Haldiram\'s</span>'
    )


# ============================================================
# CORPORATE EXCEL STYLING HELPER
# ============================================================
def style_excel_worksheet(ws, df, header_color="C5221F"):
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    data_font = Font(name="Segoe UI", size=9, color="1F2937")
    zebra_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="9CA3AF")
    medium_border_side = Side(border_style="medium", color="4B5563")

    cell_border = Border(
        left=thin_border_side, right=thin_border_side,
        top=thin_border_side, bottom=thin_border_side
    )
    header_border = Border(
        left=thin_border_side, right=thin_border_side,
        top=thin_border_side, bottom=medium_border_side
    )

    for col_idx in range(1, len(df.columns) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = header_border
    ws.row_dimensions[1].height = 26

    for row_idx in range(2, len(df) + 2):
        is_even = (row_idx % 2 == 0)
        row_fill = zebra_fill if is_even else white_fill
        ws.row_dimensions[row_idx].height = 19

        for col_idx in range(1, len(df.columns) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.fill = row_fill
            cell.border = cell_border

            col_name = str(df.columns[col_idx - 1]).upper()
            if any(term in col_name for term in ["QTY", "QUANTITY", "TOTAL_BILLS", "BILLS"]):
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif any(term in col_name for term in ["AMOUNT", "AMT", "SALES", "TAX", "GST", "DISCOUNT", "PRICE", "NET"]):
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif any(term in col_name for term in ["DATE"]):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    for col_idx, col_name in enumerate(df.columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = max([len(str(val or "")) for val in df[col_name].head(500)] + [len(str(col_name))])
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 42)

    ws.freeze_panes = "A2"


def dataframe_to_colored_excel_bytes(df, sheet_name="Report", header_color="C5221F"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]
        style_excel_worksheet(ws, df, header_color=header_color)
    output.seek(0)
    return output.getvalue()


# ============================================================
# CSS THEME & CUSTOM STYLING
# ============================================================
st.markdown("""
<style>
.stApp {
    background-color: #0B0E14 !important;
    color: #F8FAFC !important;
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
}
header[data-testid="stHeader"] {
    background-color: rgba(11, 14, 20, 0.85) !important;
    backdrop-filter: blur(12px) !important;
}
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1C0A0A 0%, #0F0404 100%) !important;
    border-right: 1px solid rgba(220, 38, 38, 0.35) !important;
}
div[data-testid="stSidebarUserContent"] {
    padding-top: 0.75rem !important;
}
div[data-testid="stSidebar"] div[data-testid="stForm"] {
    background: linear-gradient(145deg, rgba(38, 14, 14, 0.85) 0%, rgba(18, 5, 5, 0.95) 100%) !important;
    border: 1.5px solid #DC2626 !important;
    border-radius: 14px !important;
    box-shadow: 0 4px 20px rgba(220, 38, 38, 0.22), inset 0 0 12px rgba(220, 38, 38, 0.12) !important;
    padding: 16px !important;
}
div[data-testid="stSidebar"] div[data-testid="stForm"] label {
    color: #FCA5A5 !important;
    font-weight: 700 !important;
    font-size: 0.84rem !important;
    letter-spacing: 0.02em !important;
}
div[data-testid="stSidebar"] div[data-baseweb="select"] > div,
div[data-testid="stSidebar"] div[data-baseweb="input"] > div {
    background-color: #1F0A0A !important;
    border: 1.2px solid #7F1D1D !important;
    border-radius: 8px !important;
    color: #F8FAFC !important;
    transition: all 0.3s ease !important;
}
div[data-testid="stSidebar"] div[data-baseweb="select"] > div:hover,
div[data-testid="stSidebar"] div[data-baseweb="input"] > div:hover {
    border-color: #EF4444 !important;
    box-shadow: 0 0 12px rgba(239, 68, 68, 0.45) !important;
}
div[data-testid="stSidebar"] div[data-testid="stForm"] button[kind="secondaryFormSubmit"] {
    background: linear-gradient(135deg, #DC2626 0%, #991B1B 100%) !important;
    color: #FFFFFF !important;
    border: 1px solid #F87171 !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
    font-size: 0.95rem !important;
    box-shadow: 0 4px 15px rgba(220, 38, 38, 0.4) !important;
    transition: all 0.25s ease-in-out !important;
}
@keyframes logoFloat {
    0% { transform: translateY(0px) scale(1); filter: drop-shadow(0 2px 6px rgba(0, 0, 0, 0.5)); }
    50% { transform: translateY(-5px) scale(1.03); filter: drop-shadow(0 6px 14px rgba(220, 38, 38, 0.55)); }
    100% { transform: translateY(0px) scale(1); filter: drop-shadow(0 2px 6px rgba(0, 0, 0, 0.5)); }
}
@keyframes logoGlowHalo {
    0%, 100% { box-shadow: 0 0 0 0 rgba(220, 38, 38, 0.2); }
    50% { box-shadow: 0 0 22px 6px rgba(245, 158, 11, 0.35); }
}
.sidebar-logo-card {
    display: flex !important;
    justify-content: center !important;
    align-items: center !important;
    background: radial-gradient(circle, rgba(220, 38, 38, 0.25) 0%, rgba(15, 4, 4, 0) 80%) !important;
    padding: 12px 14px !important;
    margin-bottom: 0.75rem !important;
    border-radius: 14px !important;
    animation: logoGlowHalo 4s ease-in-out infinite !important;
}
img.animated-brand-logo, span.animated-brand-logo {
    display: block !important;
    animation: logoFloat 3.5s ease-in-out infinite !important;
}
div[data-testid="stMetric"] {
    background: linear-gradient(145deg, #161A24 0%, #10131B 100%) !important;
    border: 1px solid rgba(220, 38, 38, 0.3) !important;
    border-top: 4px solid #DC2626 !important;
    border-radius: 14px !important;
    padding: 18px 20px !important;
    box-shadow: 0 6px 16px rgba(0, 0, 0, 0.5) !important;
    transition: all 0.3s cubic-bezier(0.34, 1.56, 0.64, 1) !important;
}
div[data-testid="stMetric"]:hover {
    transform: translateY(-4px) scale(1.01) !important;
    border-top-color: #FBBF24 !important;
}
div[data-testid="stMetric"] label {
    color: #CBD5E1 !important;
    font-size: 0.82rem !important;
    font-weight: 700 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.07em !important;
}
div[data-testid="stMetric"] [data-testid="stMetricValue"] {
    color: #FFFFFF !important;
    font-size: 1.85rem !important;
    font-weight: 800 !important;
}
.brand-badge {
    display: inline-flex; align-items: center;
    background: linear-gradient(135deg, #DC2626 0%, #991B1B 100%);
    color: #FFFFFF; font-size: 0.75rem; font-weight: 700;
    padding: 4px 10px; border-radius: 6px; text-transform: uppercase;
}
.cache-badge {
    display: inline-flex; align-items: center;
    background: linear-gradient(135deg, #10B981 0%, #059669 100%);
    color: #FFFFFF; font-size: 0.75rem; font-weight: 700;
    padding: 4px 10px; border-radius: 6px; text-transform: uppercase;
}
.context-chip {
    background: #171111 !important;
    border: 1px solid rgba(220, 38, 38, 0.35) !important;
    padding: 4px 12px; border-radius: 8px; font-size: 0.83rem; color: #CBD5E1;
}
.context-chip strong { color: #F59E0B; }
div[data-testid="stPopover"] > button {
    background: linear-gradient(135deg, #DC2626 0%, #B91C1C 100%) !important;
    color: #FFFFFF !important; border: 1px solid #F59E0B !important;
    border-radius: 10px !important; padding: 0.65rem 1.4rem !important;
    font-weight: 700 !important; font-size: 0.92rem !important;
}
div[data-testid="stPopoverBody"] {
    background-color: #12161F !important;
    border: 1px solid rgba(220, 38, 38, 0.45) !important;
    border-radius: 14px !important; padding: 1.25rem !important;
    min-width: 440px !important;
}
.report-card-heading {
    color: #F59E0B; font-size: 0.85rem; font-weight: 800;
    text-transform: uppercase margin-bottom: 0.75rem;
    border-bottom: 1px solid rgba(245, 158, 11, 0.25); padding-bottom: 4px;
}
.analytics-box {
    background: linear-gradient(145deg, #141824 0%, #0D111A 100%);
    border: 1px solid rgba(245, 158, 11, 0.25);
    border-radius: 12px;
    padding: 16px;
    box-shadow: 0 4px 14px rgba(0, 0, 0, 0.4);
}
.analytics-box-title {
    font-size: 0.86rem;
    font-weight: 800;
    color: #FBBF24;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-bottom: 8px;
}

div[data-baseweb="tab-list"] {
    gap: 8px;
    background-color: transparent !important;
    border-bottom: 1px solid rgba(220, 38, 38, 0.3) !important;
    margin-bottom: 16px !important;
}
div[data-baseweb="tab"] {
    background: #171111 !important;
    border: 1px solid rgba(220, 38, 38, 0.35) !important;
    border-radius: 8px 8px 0px 0px !important;
    color: #CBD5E1 !important;
    font-weight: 700 !important;
    padding: 8px 20px !important;
}
div[data-baseweb="tab"][aria-selected="true"] {
    background: linear-gradient(135deg, #DC2626 0%, #991B1B 100%) !important;
    color: #FFFFFF !important;
    border-bottom: 2px solid #F59E0B !important;
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# DATABASE CONFIGURATION
# ============================================================
SERVER = "reportmartserver.database.windows.net"
DATABASE = "Reportmart_DB"
DRIVER = "ODBC Driver 18 for SQL Server"

DB_USERNAME = ""
DB_PASSWORD = ""

def get_connection():
    try:
        username = st.secrets["database"]["username"]
        password = st.secrets["database"]["password"]
    except Exception:
        username = os.environ.get("SQL_USERNAME") or DB_USERNAME
        password = os.environ.get("SQL_PASSWORD") or DB_PASSWORD

    if not username or not password:
        st.error("Database credentials are not configured. Please configure via st.secrets or script variables.")
        st.stop()

    return pymssql.connect(
        server=SERVER,
        user=username,
        password=password,
        database=DATABASE,
        port=1433,
        as_dict=False
    )

    # Automatically select the right driver for Windows or Linux
    if "ODBC Driver 18 for SQL Server" in available_drivers:
        driver = "ODBC Driver 18 for SQL Server"
        extra_args = "Encrypt=yes;TrustServerCertificate=yes;"
    elif "ODBC Driver 17 for SQL Server" in available_drivers:
        driver = "ODBC Driver 17 for SQL Server"
        extra_args = "Encrypt=yes;TrustServerCertificate=yes;"
    elif "FreeTDS" in available_drivers:
        driver = "FreeTDS"
        extra_args = "Port=1433;TDS_Version=7.4;"
    else:
        driver = available_drivers[0] if available_drivers else "ODBC Driver 18 for SQL Server"
        extra_args = "Port=1433;"

    connection_string = (
        f"DRIVER={{{driver}}};"
        f"SERVER={SERVER};"
        f"DATABASE={DATABASE};"
        f"UID={username};"
        f"PWD={password};"
        f"{extra_args}"
        "Connection Timeout=30;"
    )
    return pyodbc.connect(connection_string)

# ============================================================
# RBAC MODULE (LOCAL JSON-BASED ACCESS CONTROL WITH VERTICALS)
# ============================================================
ALL_KPIS = [
    ("tab_v1", "📌 Dashboard V1 View Tab"),
    ("tab_v2", "⚡ Dashboard V2 Advanced View Tab"),
    ("tab_verticals", "🏷️ Vertical-Wise Breakdown Tab"),
    ("total_sales", "💰 Total Sales"),
    ("total_qty", "📦 Total Qty"),
    ("total_bills", "🧾 Total Bills"),
    ("avg_bill", "💵 Avg Bill"),
    ("total_discount", "🏷️ Total Discount"),
    ("comparison_sales", "📅 Last Month & Last Year Comparisons"),
    ("sales_share_growth", "📊 Sales Share & Growth Analytics Section"),
    ("retail_performance", "🛍️ Retail Performance Section"),
    ("restaurant_performance", "🍽️ Restaurant Performance Section")
]

ALL_REPORTS = [
    ("discount_report", "🏷️ Discount Report"),
    ("itemwise_report", "📦 Item-wise Report"),
    ("full_workbook_export", "📥 Export Complete Multi-Sheet Workbook")
]

ALL_SYSTEM_PERMISSIONS = [
    ("admin_portal", "🛠️ Admin Access Portal (Manage Users & Role Permissions)"),
    ("can_force_refresh", "🔄 Allow Force Azure Refresh (Overwrites Cache)")
]

def hash_password(password: str) -> str:
    return hashlib.sha256(password.strip().encode("utf-8")).hexdigest()

def load_user_db():
    all_kpi_keys = [k[0] for k in ALL_KPIS]
    all_rep_keys = [r[0] for r in ALL_REPORTS]

    if not PERMISSIONS_FILE.exists():
        initial_db = {
            "default": {
                "kpis": all_kpi_keys,
                "reports": all_rep_keys,
                "system": ["can_force_refresh"],
                "verticals": []  # Empty means all verticals allowed
            },
            "admin": {
                "password_hash": hash_password(DEFAULT_ADMIN_PASSWORD),
                "system": ["admin_portal", "can_force_refresh"],
                "kpis": all_kpi_keys,
                "reports": all_rep_keys,
                "verticals": []
            }
        }
        with open(PERMISSIONS_FILE, "w", encoding="utf-8") as f:
            json.dump(initial_db, f, indent=4)
        return initial_db

    try:
        with open(PERMISSIONS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            if "admin" not in data or "password_hash" not in data.get("admin", {}):
                data["admin"] = {
                    "password_hash": hash_password(DEFAULT_ADMIN_PASSWORD),
                    "system": ["admin_portal", "can_force_refresh"],
                    "kpis": all_kpi_keys,
                    "reports": all_rep_keys,
                    "verticals": []
                }
                save_user_db(data)
            return data
    except Exception:
        return {}

def save_user_db(data):
    try:
        with open(PERMISSIONS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
        return True, "User data saved successfully!"
    except Exception as e:
        return False, str(e)

def authenticate_user(username, password):
    username_clean = str(username).strip().lower()
    if not username_clean or not password:
        return False, "Please enter both username and password."

    user_db = load_user_db()
    if username_clean not in user_db:
        return False, "User does not exist. Please contact your administrator."

    user_info = user_db[username_clean]
    expected_hash = user_info.get("password_hash")
    if not expected_hash:
        return False, "User account is missing a password configuration."

    if hash_password(password) == expected_hash:
        return True, ""
    return False, "Incorrect password."

def get_user_access_profile(username):
    username_clean = str(username).strip().lower()
    all_kpi_keys = [k[0] for k in ALL_KPIS]
    all_rep_keys = [r[0] for r in ALL_REPORTS]

    user_db = load_user_db()
    user_entry = user_db.get(username_clean, user_db.get("default", {}))

    saved_kpis = user_entry.get("kpis")
    saved_reports = user_entry.get("reports")
    saved_system = user_entry.get("system", [])
    saved_verticals = user_entry.get("verticals", [])  # list of allowed vertical names

    active_kpis = saved_kpis if saved_kpis is not None else all_kpi_keys
    active_reports = saved_reports if saved_reports is not None else all_rep_keys

    is_admin_user = (username_clean in [u.lower() for u in BOOTSTRAP_ADMINS])
    has_admin_portal = ("admin_portal" in saved_system) or is_admin_user
    can_force_refresh = ("can_force_refresh" in saved_system) or is_admin_user

    return {
        "kpis": active_kpis,
        "reports": active_reports,
        "system": saved_system,
        "verticals": saved_verticals,
        "is_admin": has_admin_portal,
        "can_force_refresh": can_force_refresh
    }


# ============================================================
# MASTER DATA LOADER
# ============================================================
@st.cache_data(ttl=3600)
def get_store_list():
    try:
        with get_connection() as conn:
            query = """
                SELECT DISTINCT
                    a.STORENUMBER,
                    a.NAME AS STORENAME,
                    b.VARTICAL,
                    b.CITY
                FROM RetailChannelStaging AS a
                INNER JOIN PwC_RetailStoreExtraFieldsStaging AS b
                    ON a.STORENUMBER = b.STORENUMBER
                WHERE a.STORENUMBER IS NOT NULL
                  AND a.NAME IS NOT NULL
                ORDER BY STORENAME
            """
            stores = pd.read_sql(query, conn)
        stores["STORENAME"] = stores["STORENAME"].astype(str).str.strip()
        stores["VARTICAL"] = stores["VARTICAL"].fillna("Unassigned").astype(str).str.strip().replace("", "Unassigned")
        stores["CITY"] = stores["CITY"].fillna("Unassigned").astype(str).str.strip().replace("", "Unassigned")
        return stores.drop_duplicates(subset=["STORENAME", "VARTICAL", "CITY"])[["STORENAME", "VARTICAL", "CITY"]]
    except Exception as e:
        st.error(f"Unable to load store list: {e}")
        return pd.DataFrame(columns=["STORENAME", "VARTICAL", "CITY"])


store_master = get_store_list()
master_all_verticals = sorted([v for v in store_master["VARTICAL"].unique().tolist() if v])


# ============================================================
# LOGIN PAGE
# ============================================================
def show_login_page():
    st.markdown("<br><br>", unsafe_allow_html=True)
    _, centre, _ = st.columns([1, 1.2, 1])
    with centre:
        st.markdown(f'<div class="sidebar-logo-card">{load_logo_html()}</div>', unsafe_allow_html=True)
        st.markdown("""
        <h2 style='text-align: center; color: #FFFFFF; font-weight: 800; margin-bottom: 4px;'>
            Haldiram\'s D365 Sales Command
        </h2>
        <p style='text-align: center; color: #94A3B8; font-size: 0.9rem; margin-bottom: 24px;'>
            Sign in with your configured dashboard credentials
        </p>
        """, unsafe_allow_html=True)

        with st.form("login_form"):
            username = st.text_input("Username / Login ID", placeholder="e.g. store_user01")
            password = st.text_input("Password", type="password", placeholder="Enter your password")
            submitted = st.form_submit_button("Sign In 🚀", width="stretch")

        if submitted:
            ok, message = authenticate_user(username, password)
            if ok:
                st.session_state["authenticated"] = True
                st.session_state["auth_user"] = username.strip().lower()
                st.rerun()
            else:
                st.error(message)

if not st.session_state.get("authenticated", False):
    show_login_page()
    st.stop()

current_user = st.session_state.get("auth_user", "")
user_profile = get_user_access_profile(current_user)
allowed_kpis = user_profile["kpis"]
allowed_reports = user_profile["reports"]
allowed_verticals = user_profile.get("verticals", [])  # Empty means all
can_access_admin_portal = user_profile["is_admin"]
user_can_force_refresh = user_profile["can_force_refresh"]


# ============================================================
# SIDEBAR NAVIGATION
# ============================================================
st.sidebar.markdown(f'<div class="sidebar-logo-card">{load_logo_html()}</div>', unsafe_allow_html=True)
st.sidebar.success(f"👤 Signed in as **{current_user}**")
if allowed_verticals:
    st.sidebar.info(f"🏷️ Verticals: **{', '.join(allowed_verticals)}**")
if st.sidebar.button("Log out", width="stretch"):
    st.session_state.clear()
    st.rerun()

st.sidebar.markdown("<hr style='margin: 0.75rem 0; border-color: rgba(220,38,38,0.45);'>", unsafe_allow_html=True)

navigation_options = ["📊 Sales Command"]
if can_access_admin_portal:
    navigation_options.append("🛠️ Admin Access Portal")

current_page = st.sidebar.radio("📍 Navigation", navigation_options)
st.sidebar.markdown("<hr style='margin: 0.75rem 0; border-color: rgba(220,38,38,0.45);'>", unsafe_allow_html=True)


# ============================================================
# PAGE 1: DEDICATED ADMIN ACCESS PORTAL (GUARDED WITH st.stop)
# ============================================================
if current_page == "🛠️ Admin Access Portal":
    if not can_access_admin_portal:
        st.error("🚫 Access Denied: You do not have permission to access the Admin Portal.")
        st.stop()

    st.title("🛠️ Admin Portal: User Accounts & Role Permissions")
    st.caption("Manage user credentials, assign vertical-wise permissions (V1, V2, etc.), grant/disable Admin Portal access, configure Tab, KPI & Report permissions.")

    user_db = load_user_db()
    managed_users = [u for u in sorted(user_db.keys()) if u != "default"]

    col_choice, _ = st.columns([1.5, 2])
    with col_choice:
        action_type = st.radio(
            "Account Operation:",
            ["Edit Existing User", "Create New User Account"],
            horizontal=True
        )

    if action_type == "Edit Existing User" and managed_users:
        target_user = st.selectbox("Select User ID to Manage:", managed_users)
        target_info = user_db.get(target_user, {})
        default_kpis = target_info.get("kpis", [k[0] for k in ALL_KPIS])
        default_reps = target_info.get("reports", [r[0] for r in ALL_REPORTS])
        default_system = target_info.get("system", [])
        default_verticals = target_info.get("verticals", [])
        is_creating = False
    else:
        target_user = st.text_input("Enter New Login ID:", placeholder="e.g. vijay or rupesh").strip().lower()
        target_info = {}
        default_kpis = [k[0] for k in ALL_KPIS]
        default_reps = [r[0] for r in ALL_REPORTS]
        default_system = []
        default_verticals = []
        is_creating = True

    if target_user:
        with st.form("admin_user_management_form"):
            st.markdown(f"### 👤 Configure Account: `{target_user}`")

            col_pwd1, col_pwd2 = st.columns(2)
            with col_pwd1:
                new_pwd = st.text_input(
                    "Set / Update Password:" if not is_creating else "Assign Password (Required):",
                    type="password",
                    placeholder="Leave blank to keep existing password" if not is_creating else "Enter password"
                )
            with col_pwd2:
                confirm_pwd = st.text_input("Confirm Password:", type="password", placeholder="Re-type password")

            # --------------------------------------------------------
            # VERTICAL-WISE PERMISSION SECTION
            # --------------------------------------------------------
            st.markdown("---")
            st.markdown("#### 🏷️ 1. Vertical-Wise Access Control")
            st.caption("Select allowed vertical(s) for this user (e.g. select V1 for Vijay, V2 for Rupesh). Leave unselected to grant unrestricted access to ALL verticals.")
            
            chosen_verticals = st.multiselect(
                "Assigned Verticals:",
                options=master_all_verticals,
                default=[v for v in default_verticals if v in master_all_verticals],
                help="If selected, this user can only view sales data and stores associated with these verticals."
            )

            st.markdown("---")
            st.markdown("#### 🔑 2. System Role & Administration")
            is_admin_flag = ("admin_portal" in default_system) or (target_user in BOOTSTRAP_ADMINS)
            admin_portal_enabled = st.checkbox(
                "🛠️ Enable Admin Access Portal (Allow user to create users and manage role permissions)",
                value=is_admin_flag,
                key=f"adm_sys_{target_user}"
            )
            is_refresh_allowed = ("can_force_refresh" in default_system) or (target_user in BOOTSTRAP_ADMINS)
            force_refresh_enabled = st.checkbox(
                "🔄 Allow 'Force Azure Refresh' (Allows user to bypass local cache and query Azure SQL directly)",
                value=is_refresh_allowed,
                key=f"adm_refresh_{target_user}"
            )

            st.markdown("---")
            st.markdown("#### 📊 3. Permitted Tabs, Dashboard KPIs & Sections")
            chosen_kpis = []
            col_k1, col_k2 = st.columns(2)
            for idx, (kpi_key, kpi_name) in enumerate(ALL_KPIS):
                t_col = col_k1 if idx % 2 == 0 else col_k2
                checked = kpi_key in default_kpis
                if t_col.checkbox(kpi_name, value=checked, key=f"adm_kpi_{target_user}_{kpi_key}"):
                    chosen_kpis.append(kpi_key)

            st.markdown("---")
            st.markdown("#### 📥 4. Permitted Reports & Exports")
            chosen_reps = []
            col_r1, col_r2 = st.columns(2)
            for idx, (rep_key, rep_name) in enumerate(ALL_REPORTS):
                t_col = col_r1 if idx % 2 == 0 else col_r2
                checked = rep_key in default_reps
                if t_col.checkbox(rep_name, value=checked, key=f"adm_rep_{target_user}_{rep_key}"):
                    chosen_reps.append(rep_key)

            col_sub1, col_sub2 = st.columns([1.5, 1])
            with col_sub1:
                save_submitted = st.form_submit_button("💾 Save User & Permissions", width="stretch")
            with col_sub2:
                delete_submitted = st.form_submit_button("🗑️ Delete User Account", width="stretch")

            if save_submitted:
                if is_creating and not new_pwd:
                    st.error("A password is required when creating a new user account.")
                elif new_pwd and new_pwd != confirm_pwd:
                    st.error("Passwords do not match.")
                else:
                    if target_user not in user_db:
                        user_db[target_user] = {}

                    if new_pwd:
                        user_db[target_user]["password_hash"] = hash_password(new_pwd)

                    updated_system = []
                    if admin_portal_enabled:
                        updated_system.append("admin_portal")
                    if force_refresh_enabled:
                        updated_system.append("can_force_refresh")

                    user_db[target_user]["system"] = updated_system
                    user_db[target_user]["verticals"] = chosen_verticals
                    user_db[target_user]["kpis"] = chosen_kpis
                    user_db[target_user]["reports"] = chosen_reps

                    ok, msg = save_user_db(user_db)
                    if ok:
                        st.success(f"✅ User '{target_user}' saved with vertical permissions: {chosen_verticals if chosen_verticals else 'ALL'}!")
                        st.rerun()
                    else:
                        st.error(f"Failed to save user: {msg}")

            if delete_submitted:
                if target_user in BOOTSTRAP_ADMINS:
                    st.error("Primary bootstrap admin accounts cannot be deleted.")
                elif target_user in user_db:
                    del user_db[target_user]
                    save_user_db(user_db)
                    st.warning(f"User account '{target_user}' has been deleted.")
                    st.rerun()

    st.markdown("---")
    st.subheader("📋 Configured Users Directory")
    summary_data = []
    for uid, udata in user_db.items():
        if uid == "default":
            continue
        has_adm = ("admin_portal" in udata.get("system", [])) or (uid in BOOTSTRAP_ADMINS)
        has_refr = ("can_force_refresh" in udata.get("system", [])) or (uid in BOOTSTRAP_ADMINS)
        u_verts = udata.get("verticals", [])
        vert_display = ", ".join(u_verts) if u_verts else "All Verticals"
        summary_data.append({
            "User ID": uid,
            "Vertical Access": vert_display,
            "Admin Portal Access": "✅ Enabled" if has_adm else "❌ Disabled",
            "Force Azure Refresh": "✅ Enabled" if has_refr else "❌ Disabled",
            "Allowed KPIs/Tabs": len(udata.get("kpis", [])),
            "Allowed Reports": len(udata.get("reports", []))
        })
    if summary_data:
        st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)

    # Halt execution so dashboard metrics do not run on the Admin Portal screen
    st.stop()


# ============================================================
# PAGE 2: SALES COMMAND (FILTERS & DATA QUERIES)
# ============================================================
def query_azure_for_dates(start_d, end_d):
    query = """
    SELECT
        CAST(a.TRANSACTIONDATE AS DATE) AS SALEDATE,
        a.WAREHOUSE AS StoreCode,
        e.NAME AS STORENAME,
        ISNULL(LTRIM(RTRIM(h.CITY)), 'Unassigned') AS CITY,
        ISNULL(LTRIM(RTRIM(h.VARTICAL)), 'Unassigned') AS Vertical,
        UPPER(ISNULL(g.ORDERTYPE, '')) AS Ordertype,

        SUM(ISNULL(a.NUMBEROFITEMS, 0)) AS QTY,
        SUM(-a.NETAMOUNT) AS NETAMOUNT,
        COUNT(DISTINCT a.RRECEIPTID) AS TOTAL_BILLS,
        SUM(ISNULL(j.CUSTOMERINVOICEDISCOUNTAMOUNT, 0))
          + SUM(ISNULL(j.CASHDISCOUNTAMOUNT, 0))
          + SUM(ISNULL(j.DISCOUNTAMOUNTWITHOUTTAX, 0)) AS Total_Discount_Amt,

        ISNULL(SUM(CASE
            WHEN UPPER(ISNULL(g.ORDERTYPE, '')) NOT IN ('RESTAURANT', 'FINEDINE') THEN -a.NETAMOUNT
            ELSE 0 END), 0) AS Retail_Sale,

        COUNT(DISTINCT CASE
            WHEN UPPER(ISNULL(g.ORDERTYPE, '')) NOT IN ('RESTAURANT', 'FINEDINE') THEN a.RRECEIPTID
            END) AS Retail_Bill_Count,

        SUM(CASE WHEN UPPER(ISNULL(g.ORDERTYPE, '')) IN ('RESTAURANT', 'FINEDINE')
         THEN ISNULL(j.CUSTOMERINVOICEDISCOUNTAMOUNT, 0)
            + ISNULL(j.CASHDISCOUNTAMOUNT, 0)
            + ISNULL(j.DISCOUNTAMOUNTWITHOUTTAX, 0) ELSE 0 END) AS Restaurant_Discount_Amt,

        SUM(CASE WHEN UPPER(ISNULL(g.ORDERTYPE, '')) NOT IN ('RESTAURANT', 'FINEDINE')
         THEN ISNULL(j.CUSTOMERINVOICEDISCOUNTAMOUNT, 0)
            + ISNULL(j.CASHDISCOUNTAMOUNT, 0)
            + ISNULL(j.DISCOUNTAMOUNTWITHOUTTAX, 0)
          ELSE 0 END) AS Retail_Discount_Amt,

        ISNULL(SUM(CASE
            WHEN UPPER(g.ORDERTYPE) IN ('RESTAURANT', 'FINEDINE') THEN -a.NETAMOUNT
            ELSE 0 END), 0) AS Restaurant_Sale,

        COUNT(DISTINCT CASE
            WHEN UPPER(g.ORDERTYPE) IN ('RESTAURANT', 'FINEDINE') THEN a.RRECEIPTID
           END) AS Restaurant_Bill_Count

    FROM RetailTransactionStaging AS a
    INNER JOIN (SELECT DISTINCT TRANSACTIONNUMBER
               FROM RetailTransactionPaymentLineV2Staging) 
               AS b ON a.TRANSACTIONNUMBER = b.TRANSACTIONNUMBER
    INNER JOIN RetailChannelStaging AS e ON a.WAREHOUSE = e.STORENUMBER
    INNER JOIN ( SELECT TRANSACTIONID, MAX(ORDERTYPE) AS ORDERTYPE
                FROM RETAIL_TRANSACTIONADDITIONALDETAILSStaging GROUP BY TRANSACTIONID) 
                AS g ON a.TRANSACTIONNUMBER = g.TRANSACTIONID
    INNER JOIN (SELECT STORENUMBER, MAX(CITY) AS CITY, MAX(VARTICAL) AS VARTICAL
               FROM PwC_RetailStoreExtraFieldsStaging
               GROUP BY STORENUMBER) AS h ON a.WAREHOUSE = h.STORENUMBER
    INNER JOIN (
        SELECT 
            TRANSACTIONNUMBER, 
            WAREHOUSE,
            SUM(ISNULL(CUSTOMERINVOICEDISCOUNTAMOUNT, 0)) AS CUSTOMERINVOICEDISCOUNTAMOUNT,
            SUM(ISNULL(CASHDISCOUNTAMOUNT, 0)) AS CASHDISCOUNTAMOUNT,
            SUM(ISNULL(DISCOUNTAMOUNTWITHOUTTAX, 0)) AS DISCOUNTAMOUNTWITHOUTTAX
        FROM RetailTransactionSalesLineV2Staging
        WHERE TRANSACTIONDATE >= ? AND TRANSACTIONDATE < DATEADD(DAY, 1, ?)
        GROUP BY TRANSACTIONNUMBER, WAREHOUSE
    ) AS j  
        ON a.TRANSACTIONNUMBER = j.TRANSACTIONNUMBER  
        AND a.WAREHOUSE = j.WAREHOUSE  
    WHERE
        a.TRANSACTIONSTATUS IN ('0','2')
        AND a.TRANSACTIONTYPE IN ('2','19')
        AND (
            (a.QUOTATIONEXPIRYDATE = '1900-01-01 00:00:00.000' AND a.TRANSACTIONTYPE = '2')
            OR
            (a.QUOTATIONEXPIRYDATE <> '1900-01-01 00:00:00.000' AND a.TRANSACTIONTYPE IN ('2','19'))
        )
        AND a.TRANSACTIONDATE >= ?
        AND a.TRANSACTIONDATE < DATEADD(DAY, 1, ?)
        AND a.TRANSACTIONORDERTYPE = '0'
    GROUP BY
        CAST(a.TRANSACTIONDATE AS DATE),
        a.WAREHOUSE,
        e.NAME,
        h.CITY,
        h.VARTICAL,
        UPPER(ISNULL(g.ORDERTYPE, ''));
    """
    params = [start_d, end_d, start_d, end_d]
    with get_connection() as conn:
        df = pd.read_sql(query, conn, params=params)
    return df


def get_dashboard_data_with_cache(start_date, end_date, selected_store, selected_vertical, selected_order_type, selected_city, force_refresh=False):
    date_list = [d.date() for d in pd.date_range(start_date, end_date, freq="D")]
    today_dt = datetime.today().date()

    collected_frames = []
    missing_dates = []

    for single_date in date_list:
        file_path = CACHE_DIR / f"sales_{single_date}.xlsx"
        is_today = (single_date == today_dt)

        if file_path.exists() and not force_refresh and not is_today:
            try:
                day_df = pd.read_excel(file_path, engine="openpyxl")
                collected_frames.append(day_df)
            except Exception:
                missing_dates.append(single_date)
        else:
            missing_dates.append(single_date)

    if missing_dates:
        min_date = min(missing_dates)
        max_date = max(missing_dates)
        azure_df = query_azure_for_dates(min_date, max_date)

        if not azure_df.empty:
            azure_df["SALEDATE"] = pd.to_datetime(azure_df["SALEDATE"]).dt.date
            for m_date in missing_dates:
                part = azure_df[azure_df["SALEDATE"] == m_date].copy()
                if not part.empty:
                    try:
                        part.to_excel(CACHE_DIR / f"sales_{m_date}.xlsx", index=False, engine="openpyxl")
                    except Exception:
                        pass
                    collected_frames.append(part)

    if not collected_frames:
        return pd.DataFrame()

    full_df = pd.concat(collected_frames, ignore_index=True)

    # Standardize column naming across cached files and live query
    if "StoreName" in full_df.columns and "STORENAME" not in full_df.columns:
        full_df.rename(columns={"StoreName": "STORENAME"}, inplace=True)
    if "Vertical" in full_df.columns and "VARTICAL" not in full_df.columns:
        full_df.rename(columns={"Vertical": "VARTICAL"}, inplace=True)
    if "Ordertype" in full_df.columns and "ORDERTYPE" not in full_df.columns:
        full_df.rename(columns={"Ordertype": "ORDERTYPE"}, inplace=True)

    full_df["SALEDATE"] = pd.to_datetime(full_df["SALEDATE"])
    full_df["STORENAME"] = full_df["STORENAME"].astype(str).str.strip()
    full_df["VARTICAL"] = full_df["VARTICAL"].fillna("Unassigned").astype(str).str.strip().replace("", "Unassigned")
    full_df["CITY"] = full_df["CITY"].fillna("Unassigned").astype(str).str.strip().replace("", "Unassigned")
    full_df["ORDERTYPE"] = full_df["ORDERTYPE"].fillna("").astype(str).str.strip()

    filtered = full_df.copy()

    # Enforce user vertical permission at data layer
    if allowed_verticals:
        filtered = filtered[filtered["VARTICAL"].isin(allowed_verticals)]

    if selected_vertical != "ALL":
        filtered = filtered[filtered["VARTICAL"] == selected_vertical]
    if selected_city != "ALL":
        filtered = filtered[filtered["CITY"] == selected_city]
    if selected_store != "ALL":
        filtered = filtered[filtered["STORENAME"] == selected_store]
    if selected_order_type != "ALL":
        if selected_order_type == "Restaurant":
            filtered = filtered[filtered["ORDERTYPE"].isin(["RESTAURANT", "FINEDINE"])]
        else:
            filtered = filtered[~filtered["ORDERTYPE"].isin(["RESTAURANT", "FINEDINE"])]

    numeric_columns = [
        "QTY", "NETAMOUNT", "TOTAL_BILLS", "Total_Discount_Amt",
        "Retail_Sale", "Retail_Bill_Count", "Retail_Discount_Amt",
        "Restaurant_Sale", "Restaurant_Bill_Count", "Restaurant_Discount_Amt"
    ]
    for c in numeric_columns:
        if c in filtered.columns:
            filtered[c] = pd.to_numeric(filtered[c], errors="coerce").fillna(0)

    agg_df = filtered.groupby(
        ["SALEDATE", "StoreCode", "STORENAME", "CITY", "VARTICAL"],
        as_index=False
    )[numeric_columns].sum()

    return agg_df


def execute_discount_report(start_date, end_date, selected_store, selected_vertical):
    try:
        where_conditions = ["1=1"]
        parameters = [start_date, end_date]

        if selected_store != "ALL":
            where_conditions.append("e.NAME = ?")
            parameters.append(selected_store)

        if selected_vertical != "ALL":
            where_conditions.append("f.VARTICAL = ?")
            parameters.append(selected_vertical)
        elif allowed_verticals:
            placeholders = ",".join(["?"] * len(allowed_verticals))
            where_conditions.append(f"f.VARTICAL IN ({placeholders})")
            parameters.extend(allowed_verticals)

        extra_filter_sql = "\n    AND ".join(where_conditions)

        query = f"""
        WITH sales AS (
            SELECT
                a.TRANSACTIONNUMBER,
                MAX(a.TRANSACTIONDATE) AS TRANSACTIONDATE,
                MAX(a.WAREHOUSE) AS WAREHOUSE,
                SUM(TRY_CAST(a.CUSTOMERINVOICEDISCOUNTAMOUNT AS DECIMAL(18,2)) + 
                    TRY_CAST(a.CASHDISCOUNTAMOUNT AS DECIMAL(18,2)) + 
                    TRY_CAST(a.DISCOUNTAMOUNTWITHOUTTAX AS DECIMAL(18,2))) AS Discount_Amt,
                SUM(TRY_CAST(a.SALESTAXAMOUNT AS DECIMAL(18,2))) AS Tax,
                MAX(a.PERIODICDISCOUNTPERCENTAGE) AS PERIODICDISCOUNTPERCENTAGE
            FROM RetailTransactionSalesLineV2Staging a
            WHERE a.TRANSACTIONDATE >= ? AND a.TRANSACTIONDATE < DATEADD(DAY, 1, ?)
                AND a.CATEGORYNAME <> ''
            GROUP BY a.TRANSACTIONNUMBER
        ),
        disc_single AS (
            SELECT
                x.TRANSACTIONNUMBER,
                x.DISCOUNTPERCENTAGE,
                x.DISCOUNTOFFERID,
                x.DISCOUNTORIGINTYPE
            FROM (
                SELECT
                    d.TRANSACTIONNUMBER,
                    d.DISCOUNTPERCENTAGE,
                    d.DISCOUNTOFFERID,
                    d.DISCOUNTORIGINTYPE,
                    ROW_NUMBER() OVER (
                        PARTITION BY d.TRANSACTIONNUMBER
                        ORDER BY d.LINENUMBER ASC
                    ) AS rn
                FROM RetailTransactionDiscountLineStaging d
            ) x
            WHERE x.rn = 1
        ),
        txn AS (
            SELECT
                t.TRANSACTIONNUMBER,
                MAX(t.RRECEIPTID) AS RRECEIPTID,
                MAX(t.CUSTOMERACCOUNT) AS CUSTOMERACCOUNT,
                MAX(t.CUSTOMERNAME) AS CUSTOMERNAME,
                MAX(t.PAYMENTAMOUNT) AS PAYMENTAMOUNT,
                MAX(t.GROSSAMOUNT) AS GROSSAMOUNT,
                MAX(t.NETAMOUNT) AS NETAMOUNT
            FROM RetailTransactionStaging t
            WHERE (
                (t.QUOTATIONEXPIRYDATE = '1900-01-01 00:00:00.000' AND t.TRANSACTIONTYPE = '2')
                OR (t.QUOTATIONEXPIRYDATE <> '1900-01-01 00:00:00.000' AND t.TRANSACTIONTYPE IN ('2','19'))
            ) AND t.TRANSACTIONORDERTYPE = '0'
            GROUP BY t.TRANSACTIONNUMBER
        )
        SELECT
            s.TRANSACTIONDATE,
            s.WAREHOUSE,
            e.NAME AS StoreName,
            s.TRANSACTIONNUMBER,
            tx.RRECEIPTID,
            tx.CUSTOMERACCOUNT,
            tx.CUSTOMERNAME,
            f.VARTICAL,
            df.DISCOUNTOFFERID AS DISCOUNTOFFERID,
            CASE
                WHEN df.DISCOUNTORIGINTYPE = 2 THEN 'Periodic Discount'
                WHEN df.DISCOUNTORIGINTYPE = 3 THEN 'Manual Discount'
                ELSE 'Other'
            END AS DiscountType,
            s.Discount_Amt,
            s.PERIODICDISCOUNTPERCENTAGE,
            df.DISCOUNTPERCENTAGE AS DISCOUNTPERCENTAGE,
            tx.PAYMENTAMOUNT,
            tx.GROSSAMOUNT,
            s.Tax,
            tx.NETAMOUNT,
            c.CUSTOMERPHONE
        FROM sales s
        INNER JOIN disc_single df ON s.TRANSACTIONNUMBER = df.TRANSACTIONNUMBER
        INNER JOIN RetailChannelStaging e ON s.WAREHOUSE = e.STORENUMBER
        INNER JOIN PwC_RetailStoreExtraFieldsStaging f ON s.WAREHOUSE = f.STORENUMBER
        LEFT JOIN txn tx ON s.TRANSACTIONNUMBER = tx.TRANSACTIONNUMBER
        LEFT JOIN cust c ON s.TRANSACTIONNUMBER = c.TRANSACTIONNUMBER
        WHERE {extra_filter_sql}
        ORDER BY s.TRANSACTIONDATE DESC;
        """
        with get_connection() as conn:
            df_disc = pd.read_sql(query, conn, params=parameters)
        if not df_disc.empty and "TRANSACTIONDATE" in df_disc.columns:
            df_disc["TRANSACTIONDATE"] = pd.to_datetime(df_disc["TRANSACTIONDATE"]).dt.strftime("%d-%m-%Y %H:%M")
        return df_disc
    except Exception as e:
        st.error(f"Discount report query error: {e}")
        return pd.DataFrame()


def execute_itemwise_report(start_date, end_date, selected_store, selected_vertical, item_search=""):
    try:
        where_conditions = [
            "A.TRANSACTIONSTATUS IN ('0','2')",
            "B.TRANSACTIONSTATUS IN ('0','2')",
            "A.TRANSACTIONDATE >= ?",
            "A.TRANSACTIONDATE < DATEADD(DAY, 1, ?)",
            "A.TRANSACTIONORDERTYPE <> '1'",
            """
            (
                 (A.QUOTATIONEXPIRYDATE = '1900-01-01 00:00:00.000' AND A.TRANSACTIONTYPE = '2')
              OR (A.QUOTATIONEXPIRYDATE <> '1900-01-01 00:00:00.000' AND A.TRANSACTIONTYPE IN ('2','19'))
            )
            """
        ]
        parameters = [start_date, end_date]

        if selected_store != "ALL":
            where_conditions.append("H.NAME = ?")
            parameters.append(selected_store)

        if selected_vertical != "ALL":
            if selected_vertical == "Unassigned":
                where_conditions.append("ISNULL(LTRIM(RTRIM(C.VARTICAL)), '') = ''")
            else:
                where_conditions.append("LTRIM(RTRIM(C.VARTICAL)) = ?")
                parameters.append(selected_vertical)
        elif allowed_verticals:
            placeholders = ",".join(["?"] * len(allowed_verticals))
            where_conditions.append(f"LTRIM(RTRIM(C.VARTICAL)) IN ({placeholders})")
            parameters.extend(allowed_verticals)

        if item_search.strip():
            where_conditions.append("(B.ITEMID LIKE ? OR F.PRODUCTSEARCHNAME LIKE ? OR F.CATEGORYNAME LIKE ?)")
            parameters.extend([f"%{item_search.strip()}%", f"%{item_search.strip()}%", f"%{item_search.strip()}%"])

        where_clause = "\n    AND ".join(where_conditions)

        query = f"""
        SELECT 
            A.WAREHOUSE as Store,
            H.NAME AS Store_Name,
            C.VARTICAL as Vertical,
            B.ITEMID as Item_Code,
            F.PRODUCTSEARCHNAME as Product_Name,
            B.UNIT,
            ISNULL(F.CATEGORYNAME, 'Undefined') AS Category,
            SUM(-B.QUANTITY) AS Qty,
            SUM(-B.NETAMOUNT) AS NetAmount,
            SUM(B.CUSTOMERINVOICEDISCOUNTAMOUNT)
              + SUM(B.CASHDISCOUNTAMOUNT)
              + SUM(B.DISCOUNTAMOUNTWITHOUTTAX) AS Discount_Amt,
            SUM(B.SALESTAXAMOUNT) AS GST_Amt,
            SUM(-B.NETAMOUNTINCLUSIVETAX) AS Gross_Amount
        FROM RetailTransactionStaging AS A
        INNER JOIN RetailTransactionSalesLineV2Staging AS B  
            ON A.TRANSACTIONNUMBER = B.TRANSACTIONNUMBER  
            AND A.WAREHOUSE = B.WAREHOUSE  
        INNER JOIN (
            SELECT STORENUMBER, MAX(VARTICAL) AS VARTICAL 
            FROM PwC_RetailStoreExtraFieldsStaging 
            GROUP BY STORENUMBER
        ) AS C ON A.WAREHOUSE = C.STORENUMBER
        LEFT JOIN PRODUCT_MASTER AS F ON B.ITEMID = F.PRODUCTNUMBER
        INNER JOIN (
            SELECT STORENUMBER, MAX(NAME) AS NAME 
            FROM RetailChannelStaging 
            GROUP BY STORENUMBER
        ) AS H ON A.WAREHOUSE = H.STORENUMBER
        INNER JOIN (
            SELECT TRANSACTIONNUMBER 
            FROM RetailTransactionPaymentLineV2Staging 
            GROUP BY TRANSACTIONNUMBER
        ) AS G ON A.TRANSACTIONNUMBER = G.TRANSACTIONNUMBER
        WHERE {where_clause}
        GROUP BY 
            A.WAREHOUSE, H.NAME, C.VARTICAL, F.CATEGORYNAME,
            B.ITEMID, F.PRODUCTSEARCHNAME, B.UNIT, A.TRANSACTIONTYPE
        ORDER BY A.WAREHOUSE, B.ITEMID;
        """
        with get_connection() as conn:
            df_item = pd.read_sql(query, conn, params=parameters)
        if not df_item.empty and "TRANSACTIONDATE" in df_item.columns:
            df_item["TRANSACTIONDATE"] = pd.to_datetime(df_item["TRANSACTIONDATE"]).dt.strftime("%d-%m-%Y %H:%M")
        return df_item
    except Exception as e:
        st.error(f"Item-wise query error: {e}")
        return pd.DataFrame()


# ============================================================
# DASHBOARD FILTERS & PRESETS (VERTICAL-AWARE)
# ============================================================
if allowed_verticals:
    user_store_master = store_master[store_master["VARTICAL"].isin(allowed_verticals)]
else:
    user_store_master = store_master.copy()

today = datetime.today().date()
first_day_of_month = today.replace(day=1)
yesterday = today - timedelta(days=1)

for key in ["filter_vertical", "filter_city", "filter_store", "filter_order_type", "active_start_date", "active_end_date"]:
    if key not in st.session_state:
        if "date" in key:
            st.session_state[key] = yesterday
        else:
            st.session_state[key] = "ALL"

date_presets = ["Yesterday", "Month to Date", "Last 7 Days", "Last 30 Days", "Custom"]
selected_preset = st.sidebar.selectbox("⚡ Quick Date Presets", date_presets, index=0)

if selected_preset == "Yesterday":
    preset_start, preset_end = yesterday, yesterday
elif selected_preset == "Month to Date":
    preset_start, preset_end = first_day_of_month, yesterday
elif selected_preset == "Last 7 Days":
    preset_start, preset_end = today - timedelta(days=7), yesterday
elif selected_preset == "Last 30 Days":
    preset_start, preset_end = today - timedelta(days=30), yesterday
else:
    preset_start, preset_end = st.session_state["active_start_date"], st.session_state["active_end_date"]

with st.sidebar.form("filter_engine_form"):
    st.markdown("""
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
        <span style="font-weight:800; color:#EF4444; font-size:1.05rem;">🔎 FILTER ENGINE</span>
        <span style="font-size:0.75rem; background:rgba(220,38,38,0.3); color:#FCA5A5; padding:2px 8px; border-radius:4px;">AUTO CACHE</span>
    </div>
    """, unsafe_allow_html=True)

    col_d1, col_d2 = st.columns(2)
    start_date_input = col_d1.date_input("Start Date", value=preset_start)
    end_date_input = col_d2.date_input("End Date", value=preset_end)

    if allowed_verticals:
        if len(allowed_verticals) == 1:
            vertical_options = allowed_verticals
        else:
            vertical_options = ["ALL"] + sorted(allowed_verticals)
    else:
        vertical_options = ["ALL"] + sorted(user_store_master["VARTICAL"].unique().tolist())

    curr_vert = st.session_state["filter_vertical"]
    vert_idx = vertical_options.index(curr_vert) if curr_vert in vertical_options else 0

    selected_vertical_input = st.selectbox(
        "Vertical",
        vertical_options,
        index=vert_idx
    )

    vertical_stores = user_store_master if selected_vertical_input == "ALL" else user_store_master[user_store_master["VARTICAL"] == selected_vertical_input]
    city_options = ["ALL"] + sorted(vertical_stores["CITY"].unique().tolist())
    city_idx = city_options.index(st.session_state["filter_city"]) if st.session_state["filter_city"] in city_options else 0
    selected_city_input = st.selectbox("City", city_options, index=city_idx)

    filtered_stores = vertical_stores if selected_city_input == "ALL" else vertical_stores[vertical_stores["CITY"] == selected_city_input]
    store_options = ["ALL"] + sorted(filtered_stores["STORENAME"].unique().tolist())
    store_idx = store_options.index(st.session_state["filter_store"]) if st.session_state["filter_store"] in store_options else 0
    selected_store_input = st.selectbox("Store", store_options, index=store_idx)

    order_types = ["ALL", "Retail", "Restaurant"]
    order_idx = order_types.index(st.session_state["filter_order_type"]) if st.session_state["filter_order_type"] in order_types else 0
    selected_order_type_input = st.selectbox("Order Type", order_types, index=order_idx)

    if user_can_force_refresh:
        force_refresh_checkbox = st.checkbox("🔄 Force Azure Refresh (Overwrites Cache)", value=False)
    else:
        force_refresh_checkbox = False

    submit_filters = st.form_submit_button("🚀 Apply Filters", width="stretch")

if submit_filters:
    if start_date_input > end_date_input:
        st.sidebar.error("Start Date cannot be greater than End Date.")
        st.stop()
    st.session_state["active_start_date"] = start_date_input
    st.session_state["active_end_date"] = end_date_input
    st.session_state["filter_vertical"] = selected_vertical_input
    st.session_state["filter_city"] = selected_city_input
    st.session_state["filter_store"] = selected_store_input
    st.session_state["filter_order_type"] = selected_order_type_input
    st.session_state["force_refresh"] = force_refresh_checkbox
    st.rerun()

start_date = st.session_state["active_start_date"]
end_date = st.session_state["active_end_date"]
selected_vertical = st.session_state["filter_vertical"]
selected_city = st.session_state["filter_city"]
selected_store = st.session_state["filter_store"]
selected_order_type = st.session_state["filter_order_type"]
force_refresh = st.session_state.get("force_refresh", False) and user_can_force_refresh


# EXPORT POPUP
def reset_export_state():
    st.session_state["generated_report_bytes"] = None
    st.session_state["generated_report_name"] = None
    st.session_state["generated_report_rows"] = 0
    st.session_state["rep_item_search"] = ""

@st.fragment
def render_export_menu(start_date, end_date, selected_store, selected_vertical, user_reports):
    available_choices = []
    if "discount_report" in user_reports:
        available_choices.append("🏷️ Discount Report")
    if "itemwise_report" in user_reports:
        available_choices.append("📦 Item-wise Report")

    if not available_choices:
        return

    if "rep_item_search" not in st.session_state:
        st.session_state["rep_item_search"] = ""

    with st.popover("⚡ Export Reports & Data", width="stretch"):
        st.markdown("<div class='report-card-heading'>📋 On-Demand Report Generator</div>", unsafe_allow_html=True)
        selected_report_choice = st.selectbox("Select Report Type", available_choices, key="rep_choice")

        item_search_input = ""
        if "Item-wise" in selected_report_choice:
            item_search_input = st.text_input("Filter Item / Product / Category (Optional)", placeholder="e.g. NAMKEEN", key="rep_item_search")

        col_run, col_clear = st.columns([2, 1])
        with col_run:
            run_clicked = st.button("🚀 Run & Generate", width="stretch", key="btn_run_report")
        with col_clear:
            st.button("✕ Reset", width="stretch", key="btn_reset_report", on_click=reset_export_state)

        if run_clicked:
            with st.status("⚡ Preparing Report Export...", expanded=True) as status:
                st.write("🔍 **Step 1/3:** Querying Azure SQL staging records...")
                if "Discount" in selected_report_choice:
                    rep_df = execute_discount_report(start_date, end_date, selected_store, selected_vertical)
                    rep_name = "Discount_Report"
                    sheet_name = "Discount Summary"
                    header_hex = "C5221F"
                else:
                    rep_df = execute_itemwise_report(start_date, end_date, selected_store, selected_vertical, item_search_input)
                    rep_name = "Itemwise_Report"
                    sheet_name = "Itemwise Sales"
                    header_hex = "991B1B"

                st.write(f"📊 **Step 2/3:** Loaded {len(rep_df):,} lines. Applying styling...")
                excel_bytes = dataframe_to_colored_excel_bytes(rep_df, sheet_name=sheet_name, header_color=header_hex)

                st.write("✨ **Step 3/3:** Finalizing workbook (.xlsx)...")
                status.update(label="✅ Report Generated Successfully!", state="complete", expanded=False)

                st.session_state["generated_report_bytes"] = excel_bytes
                st.session_state["generated_report_name"] = rep_name
                st.session_state["generated_report_rows"] = len(rep_df)

        if st.session_state.get("generated_report_bytes") is not None:
            active_bytes = st.session_state["generated_report_bytes"]
            active_name = st.session_state["generated_report_name"]
            rows_count = st.session_state["generated_report_rows"]

            if rows_count > 0:
                st.download_button(
                    label=f"📥 Download {active_name}.xlsx ({rows_count:,} Rows)",
                    data=active_bytes,
                    file_name=f"{active_name}_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                    key="btn_download_file"
                )
            else:
                st.info("Query returned 0 records for the selected filters.")


# TOP COMMAND BAR
col_dash_title, col_dash_menu = st.columns([2.8, 1.2])
with col_dash_title:
    st.markdown("""
    <div style="display:flex; align-items:center; gap:12px; margin-bottom: 0.35rem;">
        <span class="brand-badge">Haldiram Edition</span>
        <span class="cache-badge">⚡ Auto Disk Cache Active</span>
    </div>
    <h1 style="margin: 0; font-size: 2.1rem; font-weight: 800; letter-spacing: -0.02em;">
        Haldiram's D365 Sales Command
    </h1>
    """, unsafe_allow_html=True)

with col_dash_menu:
    st.write("")
    if any(r in allowed_reports for r in ["discount_report", "itemwise_report"]):
        render_export_menu(start_date, end_date, selected_store, selected_vertical, allowed_reports)

st.markdown(f"""
<div style="display:flex; flex-wrap:wrap; gap:8px; margin-top: 0.75rem; margin-bottom:1.25rem;">
    <span class="context-chip">📅 <strong>Period:</strong> {start_date.strftime('%d %b %Y')} → {end_date.strftime('%d %b %Y')}</span>
    <span class="context-chip">🏷️ <strong>Vertical:</strong> {selected_vertical}</span>
    <span class="context-chip">📍 <strong>City:</strong> {selected_city}</span>
    <span class="context-chip">🏪 <strong>Store:</strong> {selected_store}</span>
    <span class="context-chip">🛍️ <strong>Order Type:</strong> {selected_order_type}</span>
</div>
""", unsafe_allow_html=True)


# ============================================================
# DATA FETCHING (EXECUTES BEFORE ANY METRICS CALCULATION)
# ============================================================
progress_container = st.empty()
progress_bar = progress_container.progress(0, text="[0%] Checking local cache & Azure SQL...")

# 1. Current Period Sales Data
progress_bar.progress(25, text="[25%] Loading current period sales data...")
df = get_dashboard_data_with_cache(
    start_date, end_date,
    selected_store, selected_vertical, selected_order_type, selected_city,
    force_refresh=force_refresh
)

if df is None or df.empty:
    progress_container.empty()
    st.warning("No sales data found matching the selected filter combination.")
    st.stop()

def shifted(day, **offset):
    return (pd.Timestamp(day) - pd.DateOffset(**offset)).date()

def daily_totals(frame):
    if frame is None or frame.empty or "NETAMOUNT" not in frame.columns:
        return pd.Series(dtype=float)
    return frame.groupby(frame["SALEDATE"].dt.date)["NETAMOUNT"].sum()

# 2. Comparison Period Sales Data (Last Month & Last Year)
progress_bar.progress(60, text="[60%] Loading Last Month (LM) comparison data...")
last_month_df = get_dashboard_data_with_cache(
    shifted(start_date, months=1), shifted(end_date, months=1),
    selected_store, selected_vertical, selected_order_type, selected_city,
    force_refresh=False
)

progress_bar.progress(85, text="[85%] Loading Last Year (LY) comparison data...")
last_year_df = get_dashboard_data_with_cache(
    shifted(start_date, years=1), shifted(end_date, years=1),
    selected_store, selected_vertical, selected_order_type, selected_city,
    force_refresh=False
)

progress_bar.progress(100, text="[100%] Rendering Dashboard...")
progress_container.empty()


# ============================================================
# METRICS & AGGREGATIONS
# ============================================================
total_sales = df["NETAMOUNT"].sum()
total_qty = df["QTY"].sum()
total_bills = df["TOTAL_BILLS"].sum()
avg_bill = (total_sales / total_bills) if total_bills > 0 else 0

total_discount = df["Total_Discount_Amt"].sum() if "Total_Discount_Amt" in df.columns else 0
total_discount_k = total_discount / 1000
overall_disc_pct = (total_discount / total_sales * 100) if total_sales > 0 else 0

current_by_day = daily_totals(df)
last_month_by_day = daily_totals(last_month_df)
last_year_by_day = daily_totals(last_year_df)

trend_dates = pd.date_range(start_date, end_date, freq="D")
trend_df = pd.DataFrame({"SALEDATE": trend_dates})
trend_df["Current"] = [current_by_day.get(d.date(), 0) for d in trend_dates]
trend_df["Last Month"] = [last_month_by_day.get(shifted(d, months=1), 0) for d in trend_dates]
trend_df["Last Year"] = [last_year_by_day.get(shifted(d, years=1), 0) for d in trend_dates]

current_total = trend_df["Current"].sum()
last_month_total = trend_df["Last Month"].sum()
last_year_total = trend_df["Last Year"].sum()

def growth_pct(current, previous):
    if previous and previous > 0:
        return f"{((current - previous) / previous) * 100:+.1f}% vs Prior"
    return None

sales_growth_pct = growth_pct(total_sales, last_month_total)
total_sales_k = total_sales / 1000

retail_sales = df["Retail_Sale"].sum()
retail_bills = df["Retail_Bill_Count"].sum()
retail_avg_bill = (retail_sales / retail_bills) if retail_bills > 0 else 0
retail_discount = df["Retail_Discount_Amt"].sum() if "Retail_Discount_Amt" in df.columns else 0

restaurant_sales = df["Restaurant_Sale"].sum()
restaurant_bills = df["Restaurant_Bill_Count"].sum()
restaurant_avg_bill = (restaurant_sales / restaurant_bills) if restaurant_bills > 0 else 0
restaurant_discount = df["Restaurant_Discount_Amt"].sum() if "Restaurant_Discount_Amt" in df.columns else 0

def share_of_total(value):
    return f"{value / total_sales * 100:.1f}% of total" if total_sales else None

daily_sales = (
    df.groupby("SALEDATE", as_index=False)
    .agg(NETAMOUNT=("NETAMOUNT", "sum"), QTY=("QTY", "sum"), TOTAL_BILLS=("TOTAL_BILLS", "sum"))
    .sort_values("SALEDATE")
)

vertical_sales = (
    df.groupby("VARTICAL", as_index=False)
    .agg(
        NETAMOUNT=("NETAMOUNT", "sum"), Retail_Sale=("Retail_Sale", "sum"),
        Restaurant_Sale=("Restaurant_Sale", "sum"), Retail_Discount_Amt=("Retail_Discount_Amt", "sum"),
        Restaurant_Discount_Amt=("Restaurant_Discount_Amt", "sum"), Total_Discount_Amt=("Total_Discount_Amt", "sum"),
        QTY=("QTY", "sum"), TOTAL_BILLS=("TOTAL_BILLS", "sum")
    )
    .sort_values("NETAMOUNT", ascending=False)
)

store_sales = (
    df.groupby(["STORENAME", "VARTICAL"], as_index=False)
    .agg(NETAMOUNT=("NETAMOUNT", "sum"), QTY=("QTY", "sum"), TOTAL_BILLS=("TOTAL_BILLS", "sum"))
    .sort_values("NETAMOUNT", ascending=False)
)


# ============================================================
# DYNAMIC TABS: V1, V2 & VERTICAL BREAKDOWN
# ============================================================
available_tabs = []
if "tab_v1" in allowed_kpis:
    available_tabs.append("📌 V1 View")
if "tab_v2" in allowed_kpis:
    available_tabs.append("⚡ V2 Advanced View")
if "tab_verticals" in allowed_kpis:
    available_tabs.append("🏷️ Vertical Breakdown")

if not available_tabs:
    st.info("No dashboard views are permitted for your user ID. Please contact your administrator.")
    st.stop()

rendered_tabs = st.tabs(available_tabs)
tab_lookup = dict(zip(available_tabs, rendered_tabs))


# ------------------------------------------------------------
# 1. TAB: V1 VIEW
# ------------------------------------------------------------
if "📌 V1 View" in tab_lookup:
    with tab_lookup["📌 V1 View"]:
        st.markdown("#### 📌 Overview & Core Operations")

        kpi_cards = []
        if "total_sales" in allowed_kpis:
            kpi_cards.append(("💰 Total Sales", f"₹ {total_sales_k:,.2f} K", sales_growth_pct, "normal"))
        if "total_qty" in allowed_kpis:
            kpi_cards.append(("📦 Total Qty", f"{total_qty:,.0f}", None, "normal"))
        if "total_bills" in allowed_kpis:
            kpi_cards.append(("🧾 Total Bills", f"{total_bills:,.0f}", None, "normal"))
        if "avg_bill" in allowed_kpis:
            kpi_cards.append(("💵 Avg Bill", f"₹ {avg_bill:,.2f}", None, "normal"))
        if "total_discount" in allowed_kpis:
            kpi_cards.append(("🏷️ Total Discount", f"₹ {total_discount_k:,.2f} K", f"{overall_disc_pct:.1f}% of Sales", "off"))

        if kpi_cards:
            cols = st.columns(len(kpi_cards))
            for col, (label, val, delta, d_color) in zip(cols, kpi_cards):
                if delta:
                    col.metric(label, val, delta=delta, delta_color=d_color)
                else:
                    col.metric(label, val)

        if "comparison_sales" in allowed_kpis:
            cmp1, cmp2 = st.columns(2)
            cmp1.metric("📅 Last Month Sales (Same Dates)", f"₹ {last_month_total:,.2f}", delta=growth_pct(current_total, last_month_total))
            cmp2.metric("🗓️ Last Year Sales (Same Dates)", f"₹ {last_year_total:,.2f}", delta=growth_pct(current_total, last_year_total))

        if "retail_performance" in allowed_kpis and selected_order_type in ("ALL", "Retail"):
            st.markdown("#### 🛍️ Retail Performance")
            r1, r2, r3, r4 = st.columns(4)
            r1.metric("Retail Sales", f"₹ {retail_sales/1000:,.2f} K", share_of_total(retail_sales), delta_color="off")
            r2.metric("Retail Bills", f"{retail_bills:,.0f}")
            r3.metric("Retail Avg Bill", f"₹ {retail_avg_bill:,.2f}")
            r4.metric("🏷️ Retail Discount", f"₹ {retail_discount/1000:,.2f} K", f"{(retail_discount/retail_sales*100):.1f}% of Retail" if retail_sales > 0 else "0.0%", delta_color="off")

        if "restaurant_performance" in allowed_kpis and selected_order_type in ("ALL", "Restaurant"):
            st.markdown("#### 🍽️ Restaurant Performance")
            q1, q2, q3, q4 = st.columns(4)
            q1.metric("Restaurant Sales", f"₹ {restaurant_sales/1000:,.2f} K", share_of_total(restaurant_sales), delta_color="off")
            q2.metric("Restaurant Bills", f"{restaurant_bills:,.0f}")
            q3.metric("Restaurant Avg Bill", f"₹ {restaurant_avg_bill:,.2f}")
            q4.metric("🏷️ Restaurant Discount", f"₹ {restaurant_discount/1000:,.2f} K", f"{(restaurant_discount/restaurant_sales*100):.1f}% of Rest." if restaurant_sales > 0 else "0.0%", delta_color="off")

        st.markdown("---")

        st.subheader("📈 Daily Sales Trend")
        trend_long = trend_df.melt(id_vars="SALEDATE", var_name="Period", value_name="Sales")
        fig_daily = px.line(
            trend_long, x="SALEDATE", y="Sales", color="Period", markers=True,
            title="Daily Sales: Current vs Last Month vs Last Year (Same Dates)",
            color_discrete_map={"Current": "#DC2626", "Last Month": "#F59E0B", "Last Year": "#64748B"}
        )
        fig_daily.update_traces(selector=dict(name="Last Year"), line_dash="dot")
        fig_daily.update_layout(
            xaxis_title="Sale Date (Current Period)", yaxis_title="Sales Amount (₹)",
            hovermode="x unified", legend_title_text="",
            plot_bgcolor="#141822", paper_bgcolor="#141822", font=dict(color="#F8FAFC"),
            xaxis=dict(gridcolor="#281515"), yaxis=dict(gridcolor="#281515")
        )
        st.plotly_chart(fig_daily, use_container_width=True)

        with st.expander("📅 View Daily Sales Summary Table"):
            daily_display = daily_sales.copy()
            daily_display["NETAMOUNT"] = daily_display["NETAMOUNT"].round(2)
            daily_display.columns = ["Sale Date", "Sales", "Qty", "Bills"]
            st.dataframe(daily_display, use_container_width=True, hide_index=True)

        st.subheader("🏷️ Vertical Wise Sales (Retail vs Restaurant)")
        vertical_bifurcation = vertical_sales.melt(
            id_vars=["VARTICAL"], value_vars=["Retail_Sale", "Restaurant_Sale"],
            var_name="Channel", value_name="Sales"
        )
        vertical_bifurcation["Channel"] = vertical_bifurcation["Channel"].replace({"Retail_Sale": "Retail", "Restaurant_Sale": "Restaurant"})
        fig_vertical = px.bar(
            vertical_bifurcation, x="VARTICAL", y="Sales", color="Channel", barmode="group", text_auto=".2s",
            title="Vertical Wise Sales: Retail vs Restaurant", color_discrete_map={"Retail": "#DC2626", "Restaurant": "#F59E0B"}
        )
        fig_vertical.update_layout(
            xaxis_title="Vertical", yaxis_title="Sales Amount (₹)", hovermode="x unified", legend_title_text="Channel",
            plot_bgcolor="#141822", paper_bgcolor="#141822", font=dict(color="#F8FAFC"),
            xaxis=dict(gridcolor="#281515"), yaxis=dict(gridcolor="#281515")
        )
        st.plotly_chart(fig_vertical, use_container_width=True)

        st.subheader("🏪 Store Wise Sales")
        fig_store = px.bar(
            store_sales, x="STORENAME", y="NETAMOUNT", text_auto=".2s",
            title="Store Wise Sales", color_discrete_sequence=["#DC2626"]
        )
        fig_store.update_layout(
            xaxis_title="Store", yaxis_title="Sales Amount (₹)", xaxis_tickangle=-45,
            plot_bgcolor="#141822", paper_bgcolor="#141822", font=dict(color="#F8FAFC"),
            xaxis=dict(gridcolor="#281515"), yaxis=dict(gridcolor="#281515")
        )
        st.plotly_chart(fig_store, use_container_width=True)


# ------------------------------------------------------------
# 2. TAB: V2 ADVANCED VIEW
# ------------------------------------------------------------
if "⚡ V2 Advanced View" in tab_lookup:
    with tab_lookup["⚡ V2 Advanced View"]:
        st.markdown("#### ⚡ Advanced Sales Share, Momentum & Outlets Breakdown")

        if "sales_share_growth" in allowed_kpis:
            retail_rev = df["Retail_Sale"].sum() if "Retail_Sale" in df.columns else 0
            restaurant_rev = df["Restaurant_Sale"].sum() if "Restaurant_Sale" in df.columns else 0

            lm_diff_val = current_total - last_month_total
            lm_diff_pct = ((lm_diff_val) / last_month_total * 100) if last_month_total > 0 else 0

            ly_diff_val = current_total - last_year_total
            ly_diff_pct = ((ly_diff_val) / last_year_total * 100) if last_year_total > 0 else 0

            col_sh1, col_sh2, col_sh3 = st.columns([1.1, 1.2, 1.7])

            with col_sh1:
                st.markdown("""
                <div class="analytics-box">
                    <div class="analytics-box-title">📈 Growth Variance Overview</div>
                """, unsafe_allow_html=True)
                st.metric(
                    label="Month-over-Month (vs LM)",
                    value=f"₹ {lm_diff_val/1000:+,.2f} K",
                    delta=f"{lm_diff_pct:+.1f}% vs Same Dates LM"
                )
                st.markdown("<hr style='margin: 8px 0; border-color: rgba(245,158,11,0.2);'>", unsafe_allow_html=True)
                st.metric(
                    label="Year-over-Year (vs LY)",
                    value=f"₹ {ly_diff_val/1000:+,.2f} K",
                    delta=f"{ly_diff_pct:+.1f}% vs Same Dates LY"
                )
                st.markdown("</div>", unsafe_allow_html=True)

            with col_sh2:
                donut_fig = go.Figure(data=[go.Pie(
                    labels=["Retail", "Restaurant"],
                    values=[retail_rev, restaurant_rev],
                    hole=0.62,
                    marker=dict(colors=["#DC2626", "#F59E0B"]),
                    textinfo="percent+label",
                    hoverinfo="label+value+percent",
                    insidetextorientation="radial"
                )])
                donut_fig.update_layout(
                    title=dict(text="Revenue Share by Format", font=dict(color="#F8FAFC", size=13)),
                    showlegend=False,
                    height=200,
                    margin=dict(l=10, r=10, t=30, b=10),
                    paper_bgcolor="#141822",
                    plot_bgcolor="#141822"
                )
                st.plotly_chart(donut_fig, use_container_width=True)

            with col_sh3:
                v_summary = df.groupby("VARTICAL", as_index=False).agg(
                    NETAMOUNT=("NETAMOUNT", "sum"),
                    TOTAL_BILLS=("TOTAL_BILLS", "sum")
                )
                v_summary["Share %"] = (v_summary["NETAMOUNT"] / total_sales * 100).round(1).astype(str) + "%"
                v_summary["ABV (₹)"] = (v_summary["NETAMOUNT"] / v_summary["TOTAL_BILLS"]).fillna(0).round(2)
                v_summary["Revenue (₹ K)"] = (v_summary["NETAMOUNT"] / 1000).round(2)
                v_display = v_summary[["VARTICAL", "Revenue (₹ K)", "Share %", "ABV (₹)"]].sort_values("Revenue (₹ K)", ascending=False)
                v_display.columns = ["Vertical", "Sales (₹ K)", "Contr. %", "Avg Bill"]

                st.markdown('<div class="analytics-box-title" style="margin-bottom: 4px;">🎯 Vertical Contribution & ABV</div>', unsafe_allow_html=True)
                st.dataframe(v_display, use_container_width=True, hide_index=True, height=160)

        st.markdown("#### 🏆 Top 10 Revenue Outlets")
        top_outlets = store_sales.head(10).copy()
        top_outlets["Revenue (₹ K)"] = (top_outlets["NETAMOUNT"] / 1000).round(2)
        top_outlets["Contribution %"] = (top_outlets["NETAMOUNT"] / total_sales * 100).round(1).astype(str) + "%"

        fig_top = px.bar(
            top_outlets,
            x="Revenue (₹ K)",
            y="STORENAME",
            color="VARTICAL",
            orientation="h",
            text="Revenue (₹ K)",
            title="Top 10 Stores by Revenue (₹ in Thousands)",
            color_discrete_sequence=["#DC2626", "#F59E0B", "#10B981", "#6366F1"]
        )
        fig_top.update_layout(
            yaxis=dict(autorange="reversed"),
            plot_bgcolor="#141822",
            paper_bgcolor="#141822",
            font=dict(color="#F8FAFC"),
            xaxis=dict(gridcolor="#281515"),
            yaxis_title="Store Name",
            xaxis_title="Sales (₹ K)"
        )
        st.plotly_chart(fig_top, use_container_width=True)


# ------------------------------------------------------------
# 3. TAB: VERTICAL-WISE SEPARATE BREAKDOWN
# ------------------------------------------------------------
if "🏷️ Vertical Breakdown" in tab_lookup:
    with tab_lookup["🏷️ Vertical Breakdown"]:
        st.markdown("#### 🏷️ Vertical-Wise Performance Analytics")
        st.caption("Detailed independent view of sales, bill counts, averages, and store rankings for each vertical.")

        if allowed_verticals:
            available_verticals = sorted([v for v in df["VARTICAL"].unique().tolist() if v in allowed_verticals])
        else:
            available_verticals = sorted(df["VARTICAL"].unique().tolist())

        if not available_verticals:
            st.info("No verticals found for current filters.")
        else:
            v_sub_tabs = st.tabs([f"📁 {v}" for v in available_verticals])

            for v_tab, v_name in zip(v_sub_tabs, available_verticals):
                with v_tab:
                    v_df = df[df["VARTICAL"] == v_name]

                    v_sales = v_df["NETAMOUNT"].sum()
                    v_bills = v_df["TOTAL_BILLS"].sum()
                    v_qty = v_df["QTY"].sum()
                    v_avg_bill = (v_sales / v_bills) if v_bills > 0 else 0
                    v_disc = v_df["Total_Discount_Amt"].sum() if "Total_Discount_Amt" in v_df.columns else 0
                    v_share = (v_sales / total_sales * 100) if total_sales > 0 else 0

                    v_ret_sales = v_df["Retail_Sale"].sum() if "Retail_Sale" in v_df.columns else 0
                    v_ret_bills = v_df["Retail_Bill_Count"].sum() if "Retail_Bill_Count" in v_df.columns else 0
                    v_rest_sales = v_df["Restaurant_Sale"].sum() if "Restaurant_Sale" in v_df.columns else 0
                    v_rest_bills = v_df["Restaurant_Bill_Count"].sum() if "Restaurant_Bill_Count" in v_df.columns else 0

                    vk1, vk2, vk3, vk4, vk5 = st.columns(5)
                    vk1.metric(f"💰 {v_name} Sales", f"₹ {v_sales/1000:,.2f} K", f"{v_share:.1f}% of Total Sales")
                    vk2.metric("🧾 Total Bills", f"{v_bills:,.0f}")
                    vk3.metric("📦 Total Qty", f"{v_qty:,.0f}")
                    vk4.metric("💵 Avg Bill Value", f"₹ {v_avg_bill:,.2f}")
                    vk5.metric("🏷️ Discount Given", f"₹ {v_disc/1000:,.2f} K")

                    col_v_ret, col_v_rest = st.columns(2)
                    with col_v_ret:
                        st.markdown(f"""
                        <div class="analytics-box">
                            <div class="analytics-box-title">🛍️ Retail Format ({v_name})</div>
                            <div style="font-size:1.4rem; font-weight:800; color:#FFFFFF;">₹ {v_ret_sales/1000:,.2f} K</div>
                            <div style="font-size:0.85rem; color:#94A3B8; margin-top:4px;">
                                Bills: <strong>{v_ret_bills:,.0f}</strong> | Avg: <strong>₹ {(v_ret_sales/v_ret_bills if v_ret_bills > 0 else 0):,.2f}</strong>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                    with col_v_rest:
                        st.markdown(f"""
                        <div class="analytics-box">
                            <div class="analytics-box-title">🍽️ Restaurant Format ({v_name})</div>
                            <div style="font-size:1.4rem; font-weight:800; color:#FFFFFF;">₹ {v_rest_sales/1000:,.2f} K</div>
                            <div style="font-size:0.85rem; color:#94A3B8; margin-top:4px;">
                                Bills: <strong>{v_rest_bills:,.0f}</strong> | Avg: <strong>₹ {(v_rest_sales/v_rest_bills if v_rest_bills > 0 else 0):,.2f}</strong>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                    st.markdown("<br>", unsafe_allow_html=True)

                    v_col_c1, v_col_c2 = st.columns([1.2, 1.8])

                    with v_col_c1:
                        v_daily = v_df.groupby("SALEDATE", as_index=False)["NETAMOUNT"].sum().sort_values("SALEDATE")
                        v_fig_daily = px.area(
                            v_daily,
                            x="SALEDATE",
                            y="NETAMOUNT",
                            title=f"Daily Revenue Trend: {v_name}",
                            color_discrete_sequence=["#DC2626"]
                        )
                        v_fig_daily.update_layout(
                            xaxis_title="Sale Date",
                            yaxis_title="Amount (₹)",
                            plot_bgcolor="#141822",
                            paper_bgcolor="#141822",
                            font=dict(color="#F8FAFC"),
                            xaxis=dict(gridcolor="#281515"),
                            yaxis=dict(gridcolor="#281515"),
                            margin=dict(l=10, r=10, t=30, b=10)
                        )
                        st.plotly_chart(v_fig_daily, use_container_width=True)

                    with v_col_c2:
                        v_stores = v_df.groupby("STORENAME", as_index=False).agg(
                            Sales=("NETAMOUNT", "sum"),
                            Bills=("TOTAL_BILLS", "sum"),
                            Qty=("QTY", "sum")
                        ).sort_values("Sales", ascending=False)
                        v_stores["Avg Bill"] = (v_stores["Sales"] / v_stores["Bills"]).round(2)
                        v_stores["Sales (₹)"] = v_stores["Sales"].round(2)

                        st.markdown(f"**Top Performing Outlets in {v_name}**")
                        st.dataframe(
                            v_stores[["STORENAME", "Sales (₹)", "Bills", "Avg Bill", "Qty"]],
                            use_container_width=True,
                            hide_index=True,
                            height=250
                        )


# ============================================================
# WORKBOOK EXPORT & FOOTER
# ============================================================
st.subheader("📋 Sales Details")
detail_df = df.copy().drop(
    columns=["Retail_Sale", "Retail_Bill_Count", "Retail_Discount_Amt", "Restaurant_Sale", "Restaurant_Bill_Count", "Restaurant_Discount_Amt"],
    errors="ignore"
)
detail_df["SALEDATE"] = detail_df["SALEDATE"].dt.strftime("%d-%m-%Y")
detail_df["NETAMOUNT"] = detail_df["NETAMOUNT"].round(2)
st.dataframe(detail_df, use_container_width=True, height=350, hide_index=True)

if "full_workbook_export" in allowed_reports:
    st.subheader("📥 Export Complete Multi-Sheet Workbook")
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        sheets_data = {
            "Sales Data": df,
            "Daily Summary": daily_sales,
            "Vertical Summary": vertical_sales,
            "Store Summary": store_sales,
            "Trend Comparison": trend_df
        }
        for sheet_name, sheet_df in sheets_data.items():
            sheet_df.to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]
            style_excel_worksheet(ws, sheet_df, header_color="C5221F")

    excel_buffer.seek(0)
    file_name = f"Sales_Report_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"
    st.download_button(
        label="📥 Download Complete Excel Workbook (.xlsx)",
        data=excel_buffer.getvalue(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("---")
st.caption("D365 Sales Dashboard | Reportmart_DB | Auto Excel Disk Cache | Local RBAC Permissions Active | Haldiram's IT |")
