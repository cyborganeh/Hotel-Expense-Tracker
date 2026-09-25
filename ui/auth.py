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

    st.markdown("## 🔒 Hotel Spending Tracker")
    st.markdown("Enter the password to continue.")

    with st.form("login_form", clear_on_submit=True):
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Log in")
        if submitted:
            if password and password == _load_password():
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect password.")
        return False


def logout_button() -> None:
    """Render a logout button in the sidebar."""
    if st.sidebar.button("🔒 Log out", key="logout_button"):
        st.session_state.authenticated = False
        st.rerun()
