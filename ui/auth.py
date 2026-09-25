"""Simple password gate for the Streamlit app.

The password is resolved in this order:
  1. `.streamlit/secrets.toml` -> `[auth] password = "..."`
  2. Environment variable `HOTEL_APP_PASSWORD`
  3. The built-in default below

Change `DEFAULT_PASSWORD` (or set one of the options above) before
sharing the app with anyone.
"""
import os

import streamlit as st

# Built-in fallback password. Override it in secrets.toml or via env var.
DEFAULT_PASSWORD = "123456"


def _load_password() -> str:
    """Resolve the active password from secrets, environment, or the default."""
    try:
        auth_section = st.secrets.get("auth")
        if auth_section and auth_section.get("password"):
            return str(auth_section["password"])
    except Exception:
        pass

    return os.environ.get("HOTEL_APP_PASSWORD") or DEFAULT_PASSWORD


def require_auth() -> bool:
    """Show a password prompt if the user is not authenticated. Returns True when authenticated."""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    # Hide sidebar so the login screen is focused.
    st.markdown("""
    <style>
        [data-testid="stSidebar"] { display: none; }
        .login-container {
            max-width: 420px;
            margin: 10vh auto 0;
            padding: 2rem;
            border-radius: 16px;
            background: var(--bg-card, #151C2C);
            box-shadow: var(--card-shadow, 0 4px 16px rgba(0,0,0,0.35));
            border: 1px solid var(--border-subtle, rgba(255,255,255,0.08));
            text-align: center;
        }
        .login-title {
            font-size: 1.75rem;
            font-weight: 700;
            margin-bottom: 0.25rem;
            color: var(--text-main, #F8FAFC);
        }
        .login-subtitle {
            font-size: 0.95rem;
            color: var(--text-muted, #94A3B8);
            margin-bottom: 2rem;
        }
        .login-icon {
            font-size: 3rem;
            margin-bottom: 0.75rem;
        }
    </style>
    """, unsafe_allow_html=True)

    _, center_col, _ = st.columns([1, 2, 1])
    with center_col:
        st.markdown("""
        <div class="login-container">
            <div class="login-icon">🔒</div>
            <div class="login-title">Hotel Spending Tracker</div>
            <div class="login-subtitle">Restricted Access</div>
        </div>
        """, unsafe_allow_html=True)

        with st.form("login_form", clear_on_submit=True):
            password = st.text_input("Password", type="password", placeholder="Enter password", label_visibility="collapsed")
            submitted = st.form_submit_button("Log in", use_container_width=True, type="primary")
            if submitted:
                if password and password == _load_password():
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Incorrect password. Please try again.", icon="🚫")
        return False


def logout_button() -> None:
    """Render a logout button in the sidebar."""
    if st.sidebar.button("🔒 Log out", key="logout_button"):
        st.session_state.authenticated = False
        st.rerun()
