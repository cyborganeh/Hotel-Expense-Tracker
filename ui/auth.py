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
            max-width: 430px;
            margin: 13vh auto 0;
            padding: 0 0 1rem;
            border-bottom: 1px solid rgba(225, 238, 226, 0.16);
        }
        .login-eyebrow {
            font-size: 0.75rem;
            font-weight: 600;
            letter-spacing: 0;
            text-transform: uppercase;
            color: #A3ADA4;
            margin-bottom: 0.5rem;
        }
        .login-title {
            font-size: 1.75rem;
            font-weight: 650;
            letter-spacing: 0;
            line-height: 1.3;
            color: #E8EEE8;
            margin-bottom: 0;
        }
        [data-testid="stForm"] {
            max-width: 430px;
            margin: 0 auto;
            padding-top: 1rem;
        }
        @media (max-width: 600px) {
            .login-container { margin-top: 8vh; }
        }
    </style>
    """, unsafe_allow_html=True)

    _, center_col, _ = st.columns([1, 2, 1])
    with center_col:
        st.markdown("""
        <div class="login-container">
            <div class="login-eyebrow">Restricted access</div>
            <div class="login-title">Hotel Spending Tracker</div>
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
