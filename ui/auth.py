"""Simple session-based authentication for the Streamlit app.

Credentials can be provided via:
  1. `.streamlit/secrets.toml` under [credentials] or [auth.credentials]
  2. Environment variables `HOTEL_APP_USERNAME` and `HOTEL_APP_PASSWORD`

Example secrets.toml:
    [credentials]
    admin = { password = "your-secure-password" }
"""
import os
from typing import Dict, Optional

import streamlit as st


def _load_credentials() -> Optional[Dict[str, str]]:
    """Load username -> password mapping from secrets or environment."""
    try:
        secrets = st.secrets
        creds = secrets.get("credentials") or secrets.get("auth", {}).get("credentials")
        if creds:
            return {user: info.get("password", info) for user, info in creds.items()}
    except Exception:
        pass

    env_user = os.environ.get("HOTEL_APP_USERNAME", "admin")
    env_pass = os.environ.get("HOTEL_APP_PASSWORD")
    if env_pass:
        return {env_user: env_pass}
    return None


def require_auth() -> bool:
    """Show a login form if the user is not authenticated. Returns True when authenticated."""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    creds = _load_credentials()

    st.markdown("## 🔒 Hotel Spending Tracker")
    st.markdown("Please sign in to continue.")

    if creds is None:
        st.warning(
            "Authentication is not configured. Set credentials in `.streamlit/secrets.toml` "
            "or environment variables `HOTEL_APP_USERNAME` / `HOTEL_APP_PASSWORD`."
        )
        return False

    with st.form("login_form", clear_on_submit=True):
        username = st.text_input("Username", value="")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Log in")
        if submitted:
            stored = creds.get(username)
            if stored and stored == password:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Invalid username or password.")
        return False


def logout_button() -> None:
    """Render a logout button in the sidebar."""
    if st.sidebar.button("🔒 Log out", key="logout_button"):
        st.session_state.authenticated = False
        st.rerun()
