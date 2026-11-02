"""ui_components.py - Shared Streamlit rendering helpers for Hotel.

Single source of truth for KPI cards, shared CSS, and chart styling, used by
both the SR Separator and the Monthly Expense Tracker so the two pages cannot
drift apart again.

All charts are built WITHOUT explicit hex colors: traces carry Streamlit's
themed sentinel palette, which the frontend swaps for theme-appropriate
colors at render time (adapts automatically to light/dark mode).
"""

from typing import Any, Dict, List, Optional, Sequence

from functools import lru_cache
from pathlib import Path

import pandas as pd
import streamlit as st
import plotly.express as px


CSS_DIR = Path(__file__).resolve().parent / "ui" / "styles"


@lru_cache(maxsize=None)
def _load_css(filename: str) -> str:
    """Read a shared stylesheet once and cache it."""
    return (CSS_DIR / filename).read_text(encoding="utf-8")


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
    cards = []
    for item in items:
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
        card_class = "metric-card metric-card-primary" if not cards else "metric-card"
        cards.append(
            f'<div class="{card_class}">'
            f'<div class="metric-label">{item["label"]}</div>'
            f'<div class="metric-value">{item["value"]}</div>'
            f"{delta_html}"
            f"</div>"
        )

    # A grid lets cards reflow on narrow screens; st.columns would squeeze them.
    st.markdown(f'<div class="metric-grid">{"".join(cards)}</div>', unsafe_allow_html=True)


def render_empty_state(title: str, detail: str = "") -> None:
    """Render a bordered empty state with a clear next action.

    Presentation only: callers pass the copy they already show to users.
    """
    detail_html = f"<div>{detail}</div>" if detail else ""
    st.markdown(
        f'<div class="empty-state" role="status">'
        f"<strong>{title}</strong>{detail_html}"
        f"</div>",
        unsafe_allow_html=True,
    )


def get_metric_card_css() -> str:
    """
    Return the CSS for metric cards and shared UI components.
    This should be injected once by the main app via st.markdown(unsafe_allow_html=True).
    """
    return _load_css("metric_cards.css")


# ======================================================================
# Light-mode overrides for Streamlit's native widgets
# ======================================================================

def get_light_mode_css() -> str:
    """Return overrides that make native Streamlit widgets legible in light mode.

    Streamlit's base theme is fixed by `.streamlit/config.toml`, so the in-app
    theme toggle has to re-surface native widgets. Selectors prefer `data-testid`
    (stable across releases) and only apply while light mode is active.
    """
    return _load_css("light_mode.css")


# ======================================================================
# Charts (theme-adaptive palette; never pass explicit hex colors here)
# ======================================================================

def get_plotly_template() -> str:
    mode = st.session_state.get("theme_mode", "dark")
    return "plotly_dark" if mode == "dark" else "plotly_white"


def get_plotly_text_color() -> str:
    mode = st.session_state.get("theme_mode", "dark")
    return "#E8EEE8" if mode == "dark" else "#202820"


def _plotly_axis_style() -> Dict[str, Any]:
    text_color = get_plotly_text_color()
    mode = st.session_state.get("theme_mode", "dark")
    grid_color = "rgba(220,235,220,0.12)" if mode == "dark" else "rgba(32,40,32,0.12)"
    return {
        "automargin": True,
        "tickfont": {"color": text_color, "size": 11},
        "title": {"font": {"color": text_color, "size": 12}},
        "showline": False,
        "zeroline": False,
        "gridcolor": grid_color,
    }


def _plotly_layout_style() -> Dict[str, Any]:
    text_color = get_plotly_text_color()
    return {
        "template": get_plotly_template(),
        "paper_bgcolor": "rgba(0,0,0,0)",
        "plot_bgcolor": "rgba(0,0,0,0)",
        "font": {"size": 12, "color": text_color},
        "title": {"font": {"size": 14, "color": text_color}},
        "legend": {
            "font": {"size": 12, "color": text_color},
            "title": {"font": {"size": 12, "color": text_color}},
        },
        "hoverlabel": {
            "font": {"size": 12, "color": text_color},
            "bgcolor": "#172019" if st.session_state.get("theme_mode", "dark") == "dark" else "#FFFFFF",
            "bordercolor": "rgba(128,145,128,0.35)",
        },
    }


def themed_donut(
    df: pd.DataFrame,
    *,
    names: str,
    values: str,
    hole: float = 0.45,
    height: int = 320,
    showlegend: bool = True,
    text_size: int = 15,
    textinfo: str = "percent+label",
    legend_size: int = 13,
    margin: Optional[Dict[str, int]] = None,
) -> "plotly.graph_objects.Figure":
    """Category donut chart using the themed palette.

    text_size   - font size of the on-slice labels (px)
    textinfo    - plotly textinfo mask, e.g. 'percent+label', 'value', 'none'
    legend_size - font size of the legend entries (px)
    """
    fig = px.pie(df, names=names, values=values, hole=hole, color=names)
    text_color = get_plotly_text_color()
    fig.update_traces(
        textposition="auto",
        textinfo=textinfo,
        textfont=dict(size=text_size, color=text_color),
    )
    layout_style = _plotly_layout_style()
    layout_style["legend"] = {
        "font": {"size": legend_size, "color": text_color},
        "itemsizing": "constant",
    }
    fig.update_layout(
        **layout_style,
        margin=margin or CHART_MARGINS,
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
    axis_updates: Dict[str, Dict[str, Any]] = {
        "xaxis": _plotly_axis_style(),
        "yaxis": _plotly_axis_style(),
    }
    if value_axis_title is not None:
        axis_updates[value_axis]["title"] = {
            "text": value_axis_title,
            "font": {"color": get_plotly_text_color(), "size": 12},
        }
    if value_tickformat is not None:
        axis_updates[value_axis]["tickformat"] = value_tickformat
    if showlegend is not None:
        fig.update_layout(showlegend=showlegend)
    fig.update_layout(
        **_plotly_layout_style(),
        margin=dict(CHART_MARGINS, t=margin_top),
        height=height,
        bargap=0.3,
        **axis_updates,
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
    axis_updates: Dict[str, Dict[str, Any]] = {
        "xaxis": _plotly_axis_style(),
        "yaxis": _plotly_axis_style(),
    }
    if value_tickformat:
        axis_updates[value_axis] = {**axis_updates[value_axis], "tickformat": value_tickformat}
    fig.update_layout(
        **_plotly_layout_style(),
        margin=CHART_MARGINS,
        height=height,
        bargap=0.25,
        bargroupgap=0.08,
        **axis_updates,
    )
    if legend_horizontal:
        fig.update_layout(
            legend=dict(
                orientation="h", yanchor="bottom", y=1.02,
                xanchor="right", x=1,
                font=dict(size=12, color=get_plotly_text_color()),
            )
        )
    return fig
