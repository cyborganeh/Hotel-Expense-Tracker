"""Shared formatting helpers and validation banners."""
from typing import List
import pandas as pd
import streamlit as st


def format_idr(val) -> str:
    """Format a numeric value as Indonesian Rupiah."""
    if pd.isna(val):
        return "Rp 0"
    return f"Rp {val:,.0f}"


def format_pct(val) -> str:
    """Format a numeric value as a signed percentage."""
    if pd.isna(val):
        return "0.0%"
    return f"{val:+.1f}%"


def show_validation_banner(messages: List[str], level: str = "warning") -> None:
    """Render a validation banner with bullet captions."""
    if not messages:
        return
    if level == "error":
        st.error("### Data validation issues")
    else:
        st.warning("### Data validation issues")
    for message in messages:
        st.caption(f"• {message}")
