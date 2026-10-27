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
    st.sidebar.image(logo_path, width=44)
st.sidebar.markdown(
    """
    <div class="sidebar-brand">
        <div class="brand-name">Hotel</div>
        <div class="brand-role">Room Division &amp; Housekeeping</div>
    </div>
    """,
    unsafe_allow_html=True,
)

theme_choice = st.sidebar.radio(
    "Theme Appearance",
    ["Dark Mode", "Light Mode"],
    index=0 if st.session_state["theme_mode"] == "dark" else 1,
    horizontal=True,
    key="theme_mode_selector"
)
st.session_state["theme_mode"] = "dark" if "Dark" in theme_choice else "light"
is_dark = st.session_state["theme_mode"] == "dark"
st.session_state["plotly_text_color"] = "#E8EEE8" if is_dark else "#202820"

# Theme CSS variables. Tokens live here only; components consume var() so both
# modes stay in sync. The active token set is selected at render time because
# a script-injected html[data-theme] attribute is not reliable in Streamlit.
if is_dark:
    theme_tokens = """
        --bg-main: #111713;
        --bg-card: #171F19;
        --bg-card-hover: #202A22;
        --bg-sidebar: #0D130F;
        --text-main: #E8EEE8;
        --text-muted: #A3ADA4;
        --border-subtle: rgba(225, 238, 226, 0.12);
        --border-hover: rgba(139, 174, 114, 0.65);
        --accent: #8BAE72;
        --text-on-accent: #142018;
        --delta-pos: #8FC18B;
        --delta-neg: #E28C78;
        --card-shadow: none;
        --card-shadow-hover: none;
    """
else:
    theme_tokens = """
        --bg-main: #F3F5F0;
        --bg-card: #FFFFFF;
        --bg-card-hover: #E9EEE7;
        --bg-sidebar: #F8F9F6;
        --text-main: #202820;
        --text-muted: #5F6B60;
        --border-subtle: #D8DED6;
        --border-hover: rgba(85, 122, 62, 0.55);
        --accent: #557A3E;
        --text-on-accent: #FFFFFF;
        --delta-pos: #39723C;
        --delta-neg: #B44735;
        --card-shadow: none;
        --card-shadow-hover: none;
    """

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    :root {{
{theme_tokens}
        --font-sans: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif;
        --font-mono: 'JetBrains Mono', 'Fira Code', ui-monospace, monospace;
    }}
</style>
""", unsafe_allow_html=True)

if is_dark:
    st.markdown("""
    <style>
        [data-testid="stSidebar"] [data-testid="stImage"] img {
            filter: invert(1);
        }
    </style>
    """, unsafe_allow_html=True)

# Inject shared metric-card and UI CSS from ui_components
st.markdown(ui_components.get_metric_card_css(), unsafe_allow_html=True)

# Streamlit's base theme is fixed in .streamlit/config.toml, so light mode
# needs explicit overrides to keep native widgets legible.
if not is_dark:
    st.markdown(ui_components.get_light_mode_css(), unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Data directory setup
# ---------------------------------------------------------------------------
if "selected_data_dir" not in st.session_state:
    try:
        st.session_state.selected_data_dir = get_default_data_dir()
    except Exception:
        st.session_state.selected_data_dir = os.path.join(os.path.expanduser("~"), "Documents", "Business Review")

BASE_REVIEW_DIR = st.session_state.selected_data_dir
st.session_state.selected_data_dir = os.path.normpath(BASE_REVIEW_DIR)

if os.path.exists(BASE_REVIEW_DIR):
    data_dir_error = validate_data_dir(BASE_REVIEW_DIR)
    if data_dir_error:
        st.sidebar.error(data_dir_error)

MONTH_FOLDERS = discover_month_folders(BASE_REVIEW_DIR)
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
