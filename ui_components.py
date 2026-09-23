"""
ui_components.py - Shared Streamlit rendering helpers for Hotel Santika Depok.

Single source of truth for KPI cards and chart styling, used by both the
SR Separator and the Monthly Expense Tracker so the two pages cannot drift
apart again.

All charts are built WITHOUT explicit hex colors: traces carry Streamlit's
themed sentinel palette, which the frontend swaps for theme-appropriate
colors at render time (adapts automatically to light/dark mode).
"""

from typing import Any, Dict, List, Optional, Sequence

import pandas as pd
import streamlit as st
import plotly.express as px


# Default chart margins (title padding handled via margin_top)
CHART_MARGINS = dict(t=20, b=20, l=20, r=20)


# ======================================================================
# KPI cards
# ======================================================================

def render_kpi_row(items: Sequence[Dict[str, Any]], num_columns: Optional[int] = None) -> None:
    """
    Render a row of KPI cards with the app's standard metric-card styling.

    items: list of dicts, one per card:
        label (str)      - small caps label, e.g. 'TOTAL SPEND (SR)'
        value (str)      - pre-formatted value, e.g. 'Rp 1,234,567'
        delta (str, opt) - caption line under the value
        delta_tone (str) - 'neutral' | 'pos' (green) | 'neg' (red); default 'neutral'
    """
    cols = st.columns(num_columns or len(items))
    for col, item in zip(cols, items):
        delta = item.get("delta")
        delta_html = ""
        if delta:
            tone = item.get("delta_tone", "neutral")
            prefix = ""
            if tone == "pos" and not str(delta).startswith(("+", "▲", "↑")):
                prefix = "↑ "
            elif tone == "neg" and not str(delta).startswith(("-", "▼", "↓")):
                prefix = "↓ "
            delta_html = f'<div class="metric-delta delta-{tone}">{prefix}{delta}</div>'
        with col:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-label">{item['label']}</div>
                    <div class="metric-value">{item['value']}</div>
                    {delta_html}
                </div>
                """,
                unsafe_allow_html=True,
            )


# ======================================================================
# Charts (theme-adaptive palette; never pass explicit hex colors here)
# ======================================================================

def get_plotly_template() -> str:
    mode = st.session_state.get("theme_mode", "dark")
    return "plotly_dark" if mode == "dark" else "plotly_white"


def themed_donut(
    df: pd.DataFrame,
    *,
    names: str,
    values: str,
    hole: float = 0.45,
    height: int = 320,
    showlegend: bool = True,
) -> "plotly.graph_objects.Figure":
    """Category donut chart using the themed palette."""
    fig = px.pie(df, names=names, values=values, hole=hole, color=names)
    fig.update_traces(textposition="auto", textinfo="percent+label")
    fig.update_layout(
        template=get_plotly_template(),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=CHART_MARGINS,
        height=height,
        showlegend=showlegend,
    )
    return fig


def themed_bar(
    df: pd.DataFrame,
    *,
    x: str,
    y: str,
    color: Optional[str] = None,
    orientation: str = "v",
    title: Optional[str] = None,
    labels: Optional[Dict[str, str]] = None,
    hover_data: Optional[List[str]] = None,
    height: int = 320,
    value_axis_title: Optional[str] = None,
    value_tickformat: Optional[str] = None,
    margin_top: int = 20,
    showlegend: Optional[bool] = None,
):
    """
    Single-axes bar chart (vertical or horizontal) using the themed palette.
    """
    fig = px.bar(
        df,
        x=x,
        y=y,
        color=color,
        orientation=orientation,
        title=title,
        labels=labels,
        hover_data=hover_data,
    )
    value_axis = "xaxis" if orientation == "h" else "yaxis"
    axis_kwargs: Dict[str, Any] = {}
    if value_axis_title is not None:
        axis_kwargs["title"] = value_axis_title
    if value_tickformat is not None:
        axis_kwargs["tickformat"] = value_tickformat
    if axis_kwargs:
        fig.update_layout(**{value_axis: axis_kwargs})
    if showlegend is not None:
        fig.update_layout(showlegend=showlegend)
    fig.update_layout(
        template=get_plotly_template(),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(CHART_MARGINS, t=margin_top),
        height=height,
    )
    return fig


def themed_comparison_bar(
    categories: Sequence[Any],
    series: Sequence[Dict[str, Any]],
    *,
    orientation: str = "v",
    height: int = 380,
    value_tickformat: Optional[str] = ",",
    legend_horizontal: bool = False,
):
    """
    Grouped comparison bars (e.g. Budget vs Actual, Month vs Month).
    """
    cat_col, val_col, key_col = "Category", "Value", "Series"
    long_df = pd.DataFrame(
        [
            {cat_col: cat, val_col: val, key_col: s["name"]}
            for s in series
            for cat, val in zip(categories, s["values"])
        ]
    )
    fig = px.bar(
        long_df,
        x=cat_col if orientation == "v" else val_col,
        y=val_col if orientation == "v" else cat_col,
        color=key_col,
        barmode="group",
        orientation=orientation,
    )
    value_axis = "yaxis" if orientation == "v" else "xaxis"
    fig.update_layout(
        template=get_plotly_template(),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=CHART_MARGINS,
        height=height,
        **({value_axis: {"tickformat": value_tickformat}} if value_tickformat else {}),
    )
    if legend_horizontal:
        fig.update_layout(
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
    return fig
