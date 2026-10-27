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
    return """
    <style>
        /* Shared operational type and spacing scale; colors stay in app.py. */
        :root {
            /* Fluid type scale — clamp(min, preferred, max) */
            --type-caption: clamp(0.65rem, 0.6rem + 0.25vw, 0.75rem);
            --type-small: clamp(0.8rem, 0.75rem + 0.25vw, 0.85rem);
            --type-body: clamp(0.9rem, 0.85rem + 0.25vw, 0.95rem);
            --type-section: clamp(1.05rem, 0.95rem + 0.5vw, 1.15rem);
            --type-metric: clamp(1.35rem, 1.1rem + 1vw, 1.65rem);
            --type-title: clamp(1.6rem, 1.2rem + 2vw, 2.4rem);
            --space-1: clamp(4px, 0.25rem, 4px);
            --space-2: clamp(8px, 0.5rem, 8px);
            --space-3: clamp(12px, 0.75rem, 16px);
            --space-4: clamp(20px, 1.25rem, 24px);
            --surface-radius: clamp(6px, 0.4rem, 8px);

            /* Smooth scroll */
            scroll-behavior: smooth;

            /* Subtle gradient glow for accent elements */
            --accent-glow: rgba(139, 174, 114, 0.35);
            --accent-glow-strong: rgba(139, 174, 114, 0.6);
        }

        html, body {
            font-family: var(--font-sans);
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
            text-rendering: optimizeLegibility;
        }

        .stApp {
            background-color: var(--bg-main);
            color: var(--text-main);
            font-size: var(--type-body);
            background-image: radial-gradient(ellipse at top left, var(--accent-glow) 0%, transparent 50%);
            background-size: 100% 100%;
        }
        [data-testid="stAppViewContainer"] > .main .block-container {
            max-width: 1480px;
            padding-top: 2rem;
            padding-bottom: 3rem;
        }
        [data-testid="stSidebar"] {
            background-color: var(--bg-sidebar);
            color: var(--text-main);
            border-right: 1px solid var(--border-subtle);
        }
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 {
            margin: 20px 0 8px;
            padding-bottom: 7px;
            border-bottom: 1px solid var(--border-subtle);
            font-size: var(--type-small);
            text-transform: uppercase;
        }
        [data-testid="stSidebar"] [data-testid="stRadio"] label {
            font-size: var(--type-small);
        }
        [data-testid="stSidebar"] [data-testid="stRadio"] [role="radio"] {
            border-radius: 4px;
        }
        [data-testid="stSidebar"] [data-testid="stRadio"] [role="radio"]:focus-visible {
            outline: 2px solid var(--accent);
            outline-offset: 2px;
        }
        h2 {
            color: var(--text-main);
            font-size: var(--type-metric);
            font-weight: 650;
            line-height: 1.25;
            letter-spacing: -0.02em;
        }
        h3 {
            color: var(--text-main);
            font-size: var(--type-section);
            font-weight: 600;
            line-height: 1.35;
            margin-top: var(--space-3);
            letter-spacing: -0.01em;
        }
        h4 {
            color: var(--text-main);
            font-size: var(--type-body);
            font-weight: 600;
            line-height: 1.4;
            margin-top: var(--space-3);
        }
        [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] p {
            color: var(--text-muted) !important;
        }
        button:focus-visible, input:focus-visible,
        .stTabs [role="tab"]:focus-visible,
        .stTabs [data-baseweb="tab"]:focus-visible {
            outline: 2px solid var(--accent);
            outline-offset: var(--space-1);
        }
        button:disabled { opacity: 0.6; cursor: not-allowed; }
        [data-testid="stBaseButton-primary"] {
            background-color: var(--accent) !important;
            border: 1px solid var(--accent) !important;
            color: var(--text-on-accent) !important;
            font-weight: 600;
            transition: all 200ms ease;
        }
        [data-testid="stBaseButton-primary"]:hover {
            box-shadow: 0 0 20px var(--accent-glow);
            transform: translateY(-1px);
        }
        [data-testid="stBaseButton-secondary"] {
            border-color: var(--border-subtle) !important;
            border-radius: 5px !important;
        }
        /* Keep form controls on the same token surfaces in both themes. */
        [data-baseweb="input"], [data-baseweb="select"] > div,
        [data-baseweb="textarea"] {
            background-color: var(--bg-card) !important;
            border-color: var(--border-subtle) !important;
            color: var(--text-main) !important;
        }
        .stApp input, .stApp textarea {
            color: var(--text-main) !important;
        }
        [role="alert"], [role="status"] {
            border-radius: var(--surface-radius);
            border: 1px solid var(--border-subtle);
            padding: var(--space-3);
            font-size: var(--type-body);
            line-height: 1.5;
        }
        [role="status"] {
            color: var(--text-muted);
            background-color: var(--bg-card);
        }
        /* Metric cards - one flat surface each, reflowing as a grid. */
        .metric-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
            gap: 0;
            margin: 0 0 var(--space-4);
            border-top: 1px solid var(--border-subtle);
            border-bottom: 1px solid var(--border-subtle);
            background: var(--bg-card);
        }
        .metric-card {
            background-color: transparent !important;
            border: 0 !important;
            border-right: 1px solid var(--border-subtle) !important;
            border-radius: 0;
            padding: 14px 18px;
            box-shadow: none !important;
            color: var(--text-main) !important;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            min-width: 0;
            min-height: 84px;
            transition: background-color 200ms ease;
        }
        .metric-card:hover {
            background-color: var(--bg-card-hover) !important;
        }
        .metric-card:last-child { border-right: 0 !important; }
        .metric-card-primary .metric-value {
            font-size: 1.75rem !important;
            background: linear-gradient(135deg, var(--text-main) 0%, var(--accent) 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }
        .metric-card .metric-label {
            font-size: var(--type-caption);
            font-weight: 600;
            letter-spacing: 0;
            text-transform: uppercase;
            color: var(--text-muted) !important;
            margin-bottom: var(--space-1);
        }
        .metric-card .metric-value {
            font-size: var(--type-metric);
            font-weight: 700;
            color: var(--text-main) !important;
            letter-spacing: -0.03em;
            line-height: 1.1;
            font-variant-numeric: tabular-nums;
        }
        .metric-card .metric-delta {
            font-size: var(--type-small);
            font-weight: 500;
            margin-top: var(--space-1);
            display: flex;
            align-items: center;
            gap: var(--space-1);
            line-height: 1.4;
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

        /* Shared tab styles. Role-based selectors survive Streamlit DOM changes;
           data-baseweb is kept for older versions. Tabs wrap instead of hiding
           behind a horizontal scrollbar when there are many categories. */
        .stTabs [role="tablist"],
        .stTabs [data-baseweb="tab-list"] {
            gap: var(--space-1);
            flex-wrap: wrap;
            border-bottom: 1px solid var(--border-subtle) !important;
        }
        .stTabs [role="tab"],
        .stTabs [data-baseweb="tab"] {
            padding: var(--space-2) 10px;
            border-radius: 6px 6px 0 0;
            font-weight: 500;
            font-size: var(--type-small);
            white-space: nowrap;
            color: var(--text-muted) !important;
            transition: color 200ms ease, background-color 200ms ease, border-bottom-color 200ms ease;
        }
        .stTabs [role="tab"]:hover,
        .stTabs [data-baseweb="tab"]:hover {
            color: var(--text-main) !important;
            background-color: var(--bg-card-hover) !important;
        }
        .stTabs [role="tab"][aria-selected="true"],
        .stTabs [aria-selected="true"] {
            color: var(--accent) !important;
            background-color: transparent !important;
            border-bottom: 2px solid var(--accent) !important;
            font-weight: 700 !important;
            letter-spacing: -0.01em;
        }
        .stTabs [role="tab"]:focus-visible {
            outline-offset: -2px;
        }

        /* Data surfaces */
        [data-testid="stDataFrame"] {
            border: 1px solid var(--border-subtle) !important;
            border-radius: var(--surface-radius);
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }
        [data-testid="stDataFrame"] [role="columnheader"] {
            font-weight: 600;
            font-size: var(--type-small);
        }
        [data-testid="stDataFrame"] {
            --gdg-bg-cell: var(--bg-card);
            --gdg-bg-cell-medium: var(--bg-card-hover);
            --gdg-text-dark: var(--text-main);
            --gdg-text-medium: var(--text-muted);
            --gdg-border-color: var(--border-subtle);
        }

        /* Sidebar brand block */
        .sidebar-brand {
            padding: var(--space-1) 0 var(--space-3);
            border-bottom: 1px solid var(--border-subtle);
            margin-bottom: var(--space-3);
        }
        .sidebar-brand .brand-name {
            font-size: var(--type-section);
            font-weight: 700;
            letter-spacing: -0.02em;
            color: var(--text-main);
            line-height: 1.3;
        }
        .sidebar-brand .brand-role {
            font-size: var(--type-small);
            color: var(--text-muted);
            line-height: 1.4;
        }
        .sidebar-brand img { display: block; }

        /* Header styles */
        .main-header {
            font-size: var(--type-title);
            font-weight: 700;
            letter-spacing: -0.03em;
            color: var(--text-main) !important;
            margin-bottom: var(--space-1);
            line-height: 1.15;
        }
        .sub-header {
            font-size: var(--type-body);
            font-weight: 400;
            color: var(--text-muted) !important;
            margin-bottom: var(--space-4);
            line-height: 1.5;
        }

        /* Empty state: explains the next action and keeps tables from looking broken. */
        .empty-state {
            border: 1px dashed var(--border-subtle);
            border-radius: var(--surface-radius);
            background-color: var(--bg-card);
            color: var(--text-muted);
            padding: var(--space-3) var(--space-4);
            margin: var(--space-2) 0 var(--space-3);
            font-size: var(--type-body);
            line-height: 1.5;
        }
        .empty-state strong {
            display: block;
            color: var(--text-main);
            font-weight: 600;
            margin-bottom: var(--space-1);
        }

        @media (max-width: 768px) {
            .main-header { font-size: 1.6rem; }
            .sub-header { margin-bottom: var(--space-3); }
            [data-testid="stAppViewContainer"] > .main .block-container {
                padding: 1.25rem 1rem 2rem;
            }
            .metric-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
            .metric-card { min-height: 76px; padding: 10px 12px; }
            .metric-card:nth-child(2n) { border-right: 0 !important; }
            .metric-card:nth-child(n+3) { border-top: 1px solid var(--border-subtle) !important; }
            .metric-card .metric-value { font-size: 1.35rem; }
            /* One swipeable row beats five stacked rows on a phone. */
            .stTabs [role="tablist"],
            .stTabs [data-baseweb="tab-list"] {
                flex-wrap: nowrap;
                overflow-x: auto;
                scrollbar-width: none;
            }
            .stTabs [role="tablist"]::-webkit-scrollbar,
            .stTabs [data-baseweb="tab-list"]::-webkit-scrollbar {
                display: none;
            }
            .stTabs [role="tab"],
            .stTabs [data-baseweb="tab"] {
                flex: 0 0 auto;
            }
        }
        @media (prefers-reduced-motion: reduce) {
            .stTabs [data-baseweb="tab"] { transition: none; }
            .metric-card { transition: none; }
            [data-testid="stBaseButton-primary"] { transition: none; }
        }
    </style>
    """


# ======================================================================
# Light-mode overrides for Streamlit's native widgets
# ======================================================================

def get_light_mode_css() -> str:
    """Return overrides that make native Streamlit widgets legible in light mode.

    Streamlit's base theme is fixed by `.streamlit/config.toml`, so the in-app
    theme toggle has to re-surface native widgets. Selectors prefer `data-testid`
    (stable across releases) and only apply while light mode is active.
    """
    return """
    <style>
        [data-testid="stWidgetLabel"] p,
        [data-testid="stWidgetLabel"] label,
        [data-testid="stRadioOption"] p,
        [data-testid="stRadioOption"] span,
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] span,
        [data-testid="stMarkdownContainer"],
        [data-testid="stMarkdownContainer"] p,
        [data-testid="stMarkdownContainer"] li,
        [data-testid="stMarkdownContainer"] code {
            color: var(--text-main) !important;
        }

        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] li {
            color: var(--text-main) !important;
        }

        [data-testid="stTextInputRootElement"],
        [data-testid="stNumberInputContainer"],
        [data-testid="stMultiSelect"] [role="group"],
        [data-testid="stSelectbox"] [role="group"],
        [data-testid="stFileUploaderDropzone"],
        [data-testid="stFileUploader"] section,
        [data-testid="stExpander"] details {
            background-color: var(--bg-card) !important;
            border-color: var(--border-subtle) !important;
            color: var(--text-main) !important;
        }
        [data-testid="stTextInputField"],
        [data-testid="stNumberInputField"],
        [data-testid="stTextArea"] textarea,
        [data-testid="stSelectbox"] input,
        [data-testid="stMultiSelect"] input,
        [data-testid="stFileUploaderDropzone"] * {
            color: var(--text-main) !important;
        }

        [data-testid="stBaseButton-secondary"],
        [data-testid="stBaseButton-elementToolbar"],
        [data-testid="stBaseButton-headerNoPadding"] {
            background-color: var(--bg-card) !important;
            border-color: var(--border-subtle) !important;
            color: var(--text-main) !important;
        }
        [data-testid="stBaseButton-headerNoPadding"],
        [data-testid="stBaseButton-elementToolbar"] {
            background-color: transparent !important;
        }
        [data-testid="stBaseButton-primary"] {
            background-color: var(--accent) !important;
            border-color: var(--accent) !important;
            color: var(--text-on-accent) !important;
        }
        [data-testid="stBaseButton-secondary"] p,
        [data-testid="stBaseButton-primary"] p { color: inherit !important; }

        [data-testid="stAlertContainer"] {
            background-color: var(--bg-card) !important;
            border: 1px solid var(--border-subtle);
            color: var(--text-main) !important;
        }
        [data-testid="stAlertContentWarning"] { color: #92400E !important; }
        [data-testid="stAlertContentError"] { color: #B91C1C !important; }
        [data-testid="stAlertContentInfo"] { color: #1D4ED8 !important; }

        [data-testid="stDataFrame"] *, [data-testid="stTable"] * {
            color: var(--text-main);
        }

        [data-baseweb="popover"] > div,
        [data-baseweb="popover"] ul,
        [role="listbox"] {
            background-color: var(--bg-card) !important;
            color: var(--text-main) !important;
        }
    </style>
    """


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
