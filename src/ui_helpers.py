"""
UI/UX Enhancements, Skeleton Loaders, and Table Virtualization Helpers for ReconcileAI.
Provides high-performance shimmering skeleton placeholders,
progressive loading indicators, and fast pagination for mega-importations (Parts A through F+).
"""
import streamlit as st
import pandas as pd
from typing import Tuple, List, Any

# CSS Skeleton & Shimmer Keyframes
SKELETON_CSS = """
<style>
@keyframes skeleton-shimmer {
    0% {
        background-position: -200% 0;
    }
    100% {
        background-position: 200% 0;
    }
}

.skeleton-shimmer {
    background: linear-gradient(90deg, #1E293B 25%, #334155 37%, #1E293B 63%) !important;
    background-size: 300% 100% !important;
    animation: skeleton-shimmer 1.4s ease-in-out infinite !important;
    border-radius: 6px !important;
}

.skeleton-kpi-card {
    height: 82px;
    background-color: #1E293B;
    border: 1px solid rgba(148, 163, 184, 0.2);
    border-radius: 8px;
    padding: 12px 16px;
    margin-bottom: 12px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
}

.skeleton-kpi-title {
    width: 60%;
    height: 12px;
    margin-bottom: 8px;
}

.skeleton-kpi-value {
    width: 85%;
    height: 24px;
    margin-bottom: 6px;
}

.skeleton-kpi-sub {
    width: 45%;
    height: 10px;
}

.skeleton-table-container {
    background-color: #1E293B;
    border: 1px solid rgba(148, 163, 184, 0.2);
    border-radius: 8px;
    padding: 16px;
    margin-top: 14px;
}

.skeleton-table-header {
    height: 34px;
    width: 100%;
    margin-bottom: 12px;
    border-radius: 4px;
}

.skeleton-table-row {
    height: 28px;
    margin-bottom: 8px;
    border-radius: 4px;
}

.skeleton-detail-card {
    background-color: #1E293B;
    border: 1px solid rgba(148, 163, 184, 0.2);
    border-top: 3px solid #0284C7;
    border-radius: 8px;
    padding: 16px;
    height: 180px;
    margin-bottom: 14px;
}
</style>
"""


def inject_skeleton_css() -> None:
    """Injects high-contrast dark mode skeleton shimmer styles into the page."""
    st.markdown(SKELETON_CSS, unsafe_allow_html=True)


def render_skeleton_kpis() -> None:
    """Renders a 6-metric shimmering skeleton row mimicking top KPI cards."""
    cols = st.columns(6)
    for col in cols:
        with col:
            st.markdown("""
            <div class="skeleton-kpi-card">
                <div class="skeleton-box skeleton-shimmer skeleton-kpi-title"></div>
                <div class="skeleton-box skeleton-shimmer skeleton-kpi-value"></div>
                <div class="skeleton-box skeleton-shimmer skeleton-kpi-sub"></div>
            </div>
            """, unsafe_allow_html=True)


def render_skeleton_table(rows: int = 6) -> None:
    """Renders a shimmering placeholder table to indicate data loading."""
    row_html = "".join([
        f'<div class="skeleton-shimmer skeleton-table-row" style="width: {95 - (i % 3) * 8}%;"></div>'
        for i in range(rows)
    ])
    st.markdown(f"""
    <div class="skeleton-table-container">
        <div class="skeleton-shimmer skeleton-table-header"></div>
        {row_html}
    </div>
    """, unsafe_allow_html=True)


def render_skeleton_dashboard() -> None:
    """Renders full dashboard skeleton (KPIs + Tabs + Table)."""
    render_skeleton_kpis()
    st.markdown("""
    <div style="display: flex; gap: 8px; margin: 18px 0;">
        <div class="skeleton-shimmer" style="height: 36px; width: 140px;"></div>
        <div class="skeleton-shimmer" style="height: 36px; width: 160px;"></div>
        <div class="skeleton-shimmer" style="height: 36px; width: 140px;"></div>
        <div class="skeleton-shimmer" style="height: 36px; width: 150px;"></div>
    </div>
    """, unsafe_allow_html=True)
    render_skeleton_table(rows=7)


def paginate_dataframe(
    df: pd.DataFrame,
    page_size_options: List[int] = [50, 100, 250, 500],
    default_page_size: int = 100,
    key_prefix: str = "table"
) -> pd.DataFrame:
    """
    High-performance pagination utility for huge tables (e.g. 5,000+ items across Parts A through F).
    Prevents browser UI freezing and maintains responsive rendering speeds (<50ms).
    Returns sliced_dataframe (pd.DataFrame).
    """
    total_rows = len(df)
    if total_rows == 0 or total_rows <= min(page_size_options):
        return df

    col_size, col_page, col_info = st.columns([1.5, 2, 2.5])
    
    with col_size:
        size_choices = page_size_options + [total_rows]
        page_size = st.selectbox(
            "Filas por página / Rows:",
            options=size_choices,
            index=1 if len(size_choices) > 1 else 0,
            format_func=lambda s: "Todas / All" if s == total_rows else f"{s} filas",
            key=f"{key_prefix}_page_size"
        )

    total_pages = max(1, (total_rows + page_size - 1) // page_size)
    
    with col_page:
        if total_pages > 1:
            current_page = st.number_input(
                "Página / Page:",
                min_value=1,
                max_value=total_pages,
                value=1,
                step=1,
                key=f"{key_prefix}_current_page"
            )
        else:
            current_page = 1

    start_idx = (current_page - 1) * page_size
    end_idx = min(start_idx + page_size, total_rows)
    
    with col_info:
        st.caption(f"Mostrando {start_idx + 1} - {end_idx} de {total_rows} materiales")

    return df.iloc[start_idx:end_idx]
