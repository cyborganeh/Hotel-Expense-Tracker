"""ui_components.py - Shared Streamlit rendering helpers for Hotel.

Single source of truth for KPI cards, shared CSS, and chart styling, used by
both the SR Separator and the Monthly Expense Tracker so the two pages cannot
drift apart again.

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
        show_arrow (bool) - if True, prepend ↑/↓ arrow based on delta_tone;
                           default False (callers decide whether arrow is needed)
    """
    cols = st.columns(num_columns or len(items))
    for col, item in zip(cols, items):
        delta = item.get("delta")
        delta_html = ""
        if delta:
            tone = item.get("delta_tone", "neutral")
            prefix = ""
            if item.get("show_arrow", False):
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


def get_metric_card_css() -> str:
    """
    Return the CSS for metric cards and shared UI components.
    This should be injected once by the main app via st.markdown(unsafe_allow_html=True).
    """
    return """
    <style>
        /* Metric card styles - uses CSS variables defined by the app's theme */
        .metric-card {
            background-color: var(--bg-card) !important;
            border: 1px solid var(--border-subtle) !important;
            border-radius: 12px;
            padding: 1.1rem 1.25rem;
            box-shadow: var(--card-shadow) !important;
            color: var(--text-main) !important;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            min-height: 100px;
            transition: border-color 0.2s ease, box-shadow 0.2s ease, transform 0.15s ease;
        }
        .metric-card:hover {
            border-color: var(--border-hover) !important;
            box-shadow: var(--card-shadow-hover) !important;
            transform: translateY(-2px);
        }
        .metric-card .metric-label {
            font-size: 0.76rem;
            font-weight: 650;
            letter-spacing: 0.05em;
            text-transform: uppercase;
            color: var(--text-muted) !important;
            margin-bottom: 0.35rem;
        }
        .metric-card .metric-value {
            font-size: 1.65rem;
            font-weight: 750;
            color: var(--text-main) !important;
            letter-spacing: -0.02em;
            line-height: 1.2;
        }
        .metric-card .metric-delta {
            font-size: 0.82rem;
            font-weight: 550;
            margin-top: 0.4rem;
            display: flex;
            align-items: center;
            gap: 4px;
        }
        .metric-card .metric-delta.delta-pos {
            color: var(--delta-pos) !important;
        }
        .metric-card .metric-delta.delta-neg {
            color: var(--delta-neg) !important;
        }
        .metric-card .metric-delta.delta-neutral {
            color: var(--text-muted) !important;
        }

        /* Shared tab styles */
        .stTabs [data-baseweb="tab-list"] {
            gap: 6px;
            border-bottom: 1px solid var(--border-subtle) !important;
        }
        .stTabs [data-baseweb="tab"] {
            padding: 8px 16px;
            border-radius: 8px 8px 0 0;
            font-weight: 550;
            font-size: 0.92rem;
            color: var(--text-muted) !important;
            transition: all 0.2s ease;
        }
        .stTabs [data-baseweb="tab"]:hover {
            color: var(--text-main) !important;
            background-color: var(--bg-card-hover) !important;
        }
        .stTabs [aria-selected="true"] {
            color: var(--accent) !important;
            border-bottom: 2px solid var(--accent) !important;
            font-weight: 650 !important;
        }

        /* DataFrame border */
        [data-testid="stDataFrame"] {
            border: 1px solid var(--border-subtle) !important;
            border-radius: 8px;
        }

        /* Header styles */
        .main-header {
            font-size: 2.1rem;
            font-weight: 750;
            letter-spacing: -0.025em;
            color: var(--text-main) !important;
            margin-bottom: 0.25rem;
            line-height: 1.25;
        }
        .sub-header {
            font-size: 0.96rem;
            font-weight: 450;
            color: var(--text-muted) !important;
            margin-bottom: 1.5rem;
        }
    </style>
    """


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
