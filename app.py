"""app.py - Streamlit Web Application for Hotel.

This module is intentionally thin: it only handles app-level configuration,
theme setup, sidebar navigation, and view dispatching. All page-specific
logic lives in the `views/` package; shared data helpers live in `data/`;
shared UI helpers live in `ui/`.
"""

import os

import streamlit as st

import ui_components
from data.months import get_default_data_dir, discover_month_folders
from ui.auth import require_auth, logout_button
from ui.security import validate_data_dir
from views.monthly_tracker import render_monthly_tracker
from views.sr_separator import render_sr_separator


def _get_logo_path() -> str:
    """Return the path to the local hotel logo, or None if it does not exist."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    local_logo = os.path.join(base_dir, "ui", "hotel.png")
    if os.path.exists(local_logo):
        return local_logo
    return None


# ---------------------------------------------------------------------------
# Page configuration
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Hotel Spending Tracker & Separator",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------------------------------------------------------
# Authentication gate
# ---------------------------------------------------------------------------
if not require_auth():
    st.stop()

# ---------------------------------------------------------------------------
# Theme state
# ---------------------------------------------------------------------------
if "theme_mode" not in st.session_state:
    st.session_state["theme_mode"] = "dark"

# ---------------------------------------------------------------------------
# Sidebar branding & theme mode choice
# ---------------------------------------------------------------------------
logo_path = _get_logo_path()
if logo_path:
    st.sidebar.image(logo_path, width=64)
st.sidebar.title("Hotel")
st.sidebar.markdown("**Room Division & Housekeeping**")

theme_choice = st.sidebar.radio(
    "Theme Appearance",
    ["🌙 Dark Mode", "☀️ Light Mode"],
    index=0 if st.session_state["theme_mode"] == "dark" else 1,
    horizontal=True,
    key="theme_mode_selector"
)
st.session_state["theme_mode"] = "dark" if "Dark" in theme_choice else "light"
is_dark = st.session_state["theme_mode"] == "dark"

# Apply theme CSS variables based on session state
theme_mode = st.session_state.get("theme_mode", "dark")
st.markdown(f"""
<script>
    document.documentElement.setAttribute('data-theme', '{theme_mode}');
</script>
""", unsafe_allow_html=True)

# Theme CSS variables (dark defaults + light overrides)
st.markdown("""
<style>
    :root {
        --bg-main: #0B0F19;
        --bg-card: #151C2C;
        --bg-card-hover: #1A2338;
        --bg-sidebar: #080D1A;
        --text-main: #F8FAFC;
        --text-muted: #94A3B8;
        --border-subtle: rgba(255, 255, 255, 0.08);
        --border-hover: rgba(59, 130, 246, 0.5);
        --accent: #3B82F6;
        --delta-pos: #34D399;
        --delta-neg: #F87171;
        --card-shadow: 0 4px 16px rgba(0, 0, 0, 0.35);
        --card-shadow-hover: 0 6px 22px rgba(0, 0, 0, 0.5);
    }

    [data-theme="light"] {
        --bg-main: #F8FAFC;
        --bg-card: #FFFFFF;
        --bg-card-hover: #F1F5F9;
        --bg-sidebar: #FFFFFF;
        --text-main: #0F172A;
        --text-muted: #475569;
        --border-subtle: #E2E8F0;
        --border-hover: rgba(37, 99, 235, 0.45);
        --accent: #2563EB;
        --delta-pos: #15803D;
        --delta-neg: #DC2626;
        --card-shadow: 0 1px 4px rgba(0, 0, 0, 0.05);
        --card-shadow-hover: 0 4px 14px rgba(0, 0, 0, 0.08);
    }
</style>
""", unsafe_allow_html=True)

# Inject shared metric-card and UI CSS from ui_components
st.markdown(ui_components.get_metric_card_css(), unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Data directory setup
# ---------------------------------------------------------------------------
if "selected_data_dir" not in st.session_state:
    st.session_state.selected_data_dir = get_default_data_dir()

BASE_REVIEW_DIR = st.session_state.selected_data_dir
st.session_state.selected_data_dir = os.path.normpath(BASE_REVIEW_DIR)

if os.path.exists(BASE_REVIEW_DIR):
    data_dir_error = validate_data_dir(BASE_REVIEW_DIR)
    if data_dir_error:
        st.sidebar.error(data_dir_error)

MONTH_FOLDERS = discover_month_folders(BASE_REVIEW_DIR)
if not MONTH_FOLDERS:
    MONTH_FOLDERS = {
        "August 2026 (Agustus)": os.path.join(BASE_REVIEW_DIR, "8.AGUSTUS"),
        "July 2026 (Juli)": os.path.join(BASE_REVIEW_DIR, "7.JULY"),
    }
has_local_folder = os.path.exists(BASE_REVIEW_DIR)

# ---------------------------------------------------------------------------
# Sidebar navigation
# ---------------------------------------------------------------------------
app_mode = st.sidebar.radio(
    "Application Mode",
    ["💰 Monthly Expense Tracker", "📦 Stock Request (SR) Separator"],
    index=0
)
st.sidebar.divider()

# ---------------------------------------------------------------------------
# Logout
# ---------------------------------------------------------------------------
logout_button()

# ---------------------------------------------------------------------------
# Dispatch
# ---------------------------------------------------------------------------
if app_mode == "📦 Stock Request (SR) Separator":
    render_sr_separator()
else:
    render_monthly_tracker()
