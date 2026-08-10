"""
╔═══════════════════════════════════════════════════════════════════════════════
║  BI PLATFORM — Монолитное Streamlit Приложение
║  Один файл • Google Sheets • Excel • Дашборды • Тренды • Аномалии
║  Вдохновлено: github.com/mckinsey/vizro/tree/main/vizro-ai
╚═══════════════════════════════════════════════════════════════════════════════

Запуск:
    pip install streamlit pandas plotly openpyxl gspread
    streamlit run app.py

Функционал:
    ✓ Синхронизация Google Sheets (автообновление)
    ✓ Загрузка Excel/CSV файлов
    ✓ Динамические фильтры по всем колонкам
    ✓ KPI карточки с трендами и sparkline
    ✓ Конструктор графиков (5 типов, агрегации, группировки)
    ✓ Сводные таблицы (Pivot)
    ✓ Анализ трендов (скользящее среднее, рост %)
    ✓ Детекция аномалий (>2.5σ)
    ✓ Таблица с сортировкой и экспортом
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import re, io, json, hashlib
from typing import Dict, List, Optional, Tuple, Any

# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIG & THEME
# ═══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="BI Platform",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Vizro-inspired dark theme
VIZRO_COLORS = ['#6366f1', '#22c55e', '#f59e0b', '#ef4444', '#06b6d4', '#a855f7', '#ec4899', '#84cc16', '#f97316', '#14b8a6']

def apply_theme():
    st.markdown("""
    <style>
    .main { background: #f8fafc; }
    .stMetric { background: white; border-radius: 12px; padding: 16px; border: 1px solid #e2e8f0; }
    [data-testid="stMetricValue"] { font-size: 28px; font-weight: 700; color: #1e293b; }
    [data-testid="stMetricDelta"] { font-size: 12px; }
    .card { background: white; border-radius: 16px; padding: 20px; border: 1px solid #e2e8f0; margin-bottom: 16px; }
    .header { background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%); color: white; padding: 20px; border-radius: 16px; margin-bottom: 20px; }
    .stTabs [data-baseweb="tab-list"] { gap: 4px; }
    .stTabs [data-baseweb="tab"] { border-radius: 8px; padding: 8px 16px; }
    .stTabs [aria-selected="true"] { background: #4f46e5; color: white; }
    </style>
    """, unsafe_allow_html=True)

apply_theme()

# ═══════════════════════════════════════════════════════════════════════════════
#  SAMPLE DATA GENERATOR
# ═══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def generate_sample_data(n: int = 1400) -> pd.DataFrame:
    """Генерация демо-данных продаж"""
    np.random.seed(20240715)
    dates = pd.date_range('2024-01-01', periods=730, freq='D')
    regions = ['Москва', 'Санкт-Петербург', 'Урал', 'Сибирь', 'Юг', 'Дальний Восток']
    categories = ['Электроника', 'Одежда', 'Продукты', 'Мебель', 'Спорт', 'Красота', 'Книги']
    channels = ['Онлайн', 'Розница', 'Маркетплейс', 'Опт']
    managers = ['Иванов', 'Петрова', 'Сидоров', 'Кузнецова', 'Смирнов', 'Волкова', 'Козлов', 'Морозова']
    cat_base = {'Электроника': 95000, 'Одежда': 34000, 'Продукты': 18000, 'Мебель': 61000, 
                'Спорт': 27000, 'Красота': 22000, 'Книги': 8500}
    
    data = []
    for _ in range(n):
        d = dates[np.random.randint(0, len(dates))]
        cat = np.random.choice(categories)
        month = d.month
        seasonal = 1 + 0.4 * np.sin((month - 1) / 12 * 2 * np.pi) + (0.5 if month in [11, 12] else 0)
        trend = 1 + (dates.get_loc(d) / len(dates)) * 0.5
        qty = np.random.randint(1, 15)
        revenue = int(cat_base[cat] * seasonal * trend * (0.55 + np.random.random() * 0.9) * qty / 6)
        margin = 0.12 + np.random.random() * 0.26
        profit = int(revenue * margin)
        data.append({
            'Дата': d.strftime('%Y-%m-%d'),
            'Регион': np.random.choice(regions),
            'Категория': cat,
            'Канал': np.random.choice(channels),
            'Менеджер': np.random.choice(managers),
            'Количество': qty,
            'Выручка,₽': revenue,
            'Прибыль,₽': profit,
            'Расходы,₽': revenue - profit
        })
    return pd.DataFrame(data).sort_values('Дата').reset_index(drop=True)

# ═══════════════════════════════════════════════════════════════════════════════
#  DATA LOADING FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

def parse_google_sheet_url(url: str) -> Optional[str]:
    """Извлекает ID таблицы из URL Google Sheets"""
    match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', url)
    if match:
        sheet_id = match.group(1)
        gid_match = re.search(r'gid=(\d+)', url)
        gid = gid_match.group(1) if gid_match else '0'
        return f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}'
    return None

@st.cache_data(ttl=60)
def load_google_sheet(url: str) -> pd.DataFrame:
    """Загрузка данных из Google Sheets"""
    csv_url = parse_google_sheet_url(url)
    if not csv_url:
        raise ValueError("Неверный URL Google Sheets. Пример: https://docs.google.com/spreadsheets/d/ID/edit")
    try:
        df = pd.read_csv(csv_url)
        if df.empty:
            raise ValueError("Таблица пуста")
        return df
    except Exception as e:
        # Fallback через gviz
        sheet_id = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
        if sheet_id:
            gviz_url = f"https://docs.google.com/spreadsheets/d/{sheet_id.group(1)}/gviz/tq?tqx=out:csv"
            df = pd.read_csv(gviz_url)
            return df
        raise e

def load_file(uploaded_file) -> pd.DataFrame:
    """Загрузка Excel/CSV файла"""
    filename = uploaded_file.name.lower()
    if filename.endswith('.csv'):
        return pd.read_csv(uploaded_file)
    elif filename.endswith(('.xlsx', '.xls')):
        return pd.read_excel(uploaded_file)
    elif filename.endswith('.json'):
        return pd.read_json(uploaded_file)
    else:
        # Попытка CSV
        return pd.read_csv(uploaded_file, encoding='utf-8')

# ═══════════════════════════════════════════════════════════════════════════════
#  ANALYTICS FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

def calc_trend(values: pd.Series) -> Tuple[float, float]:
    """Расчёт тренда: slope и % изменение"""
    if len(values) < 2:
        return 0.0, 0.0
    x = np.arange(len(values))
    y = values.values
    n = len(x)
    sum_x, sum_y = x.sum(), y.sum()
    sum_xy = (x * y).sum()
    sum_x2 = (x ** 2).sum()
    denom = n * sum_x2 - sum_x ** 2
    slope = (n * sum_xy - sum_x * sum_y) / denom if denom != 0 else 0
    pct = ((y[-1] - y[0]) / (abs(y[0]) + 1e-9)) * 100
    return slope, pct

def detect_anomalies(df: pd.DataFrame, col: str, threshold: float = 2.5) -> pd.DataFrame:
    """Поиск аномалий по правилу > threshold * σ"""
    values = df[col].dropna()
    if len(values) < 20:
        return pd.DataFrame()
    mean = values.mean()
    std = values.std()
    if std == 0:
        return pd.DataFrame()
    anomalies = df[np.abs(df[col] - mean) > threshold * std]
    return anomalies.head(50)

def aggregate_data(df: pd.DataFrame, x: str, y: str, agg: str, color: Optional[str] = None, 
                   bucket: str = 'month') -> pd.DataFrame:
    """Агрегация данных для графиков"""
    result = df.copy()
    
    # Bucketing для дат
    if pd.api.types.is_datetime64_any_dtype(result[x]) or x.lower() in ['дата', 'date']:
        result[x] = pd.to_datetime(result[x], errors='coerce')
        if bucket == 'day':
            result[x] = result[x].dt.strftime('%Y-%m-%d')
        elif bucket == 'week':
            result[x] = result[x].dt.to_period('W').dt.start_time.dt.strftime('%Y-%m-%d')
        elif bucket == 'month':
            result[x] = result[x].dt.to_period('M').dt.start_time.dt.strftime('%Y-%m')
        elif bucket == 'quarter':
            result[x] = result[x].dt.to_period('Q').dt.start_time.dt.strftime('%Y-Q%q')
        elif bucket == 'year':
            result[x] = result[x].dt.to_period('Y').dt.start_time.dt.strftime('%Y')
    
    # Группировка
    group_cols = [x]
    if color and color in result.columns:
        group_cols.append(color)
    
    if agg == 'count':
        grouped = result.groupby(group_cols).size().reset_index(name=y)
    elif agg == 'avg':
        grouped = result.groupby(group_cols)[y].mean().reset_index()
    elif agg == 'min':
        grouped = result.groupby(group_cols)[y].min().reset_index()
    elif agg == 'max':
        grouped = result.groupby(group_cols)[y].max().reset_index()
    else:  # sum
        grouped = result.groupby(group_cols)[y].sum().reset_index()
    
    return grouped

def create_pivot_table(df: pd.DataFrame, index: str, columns: Optional[str], 
                       values: str, agg: str) -> pd.DataFrame:
    """Создание сводной таблицы"""
    return pd.pivot_table(df, index=index, columns=columns, values=values, 
                         aggfunc=agg, fill_value=0, margins=True, margins_name='Итого')

# ═══════════════════════════════════════════════════════════════════════════════
#  CHART BUILDERS
# ═══════════════════════════════════════════════════════════════════════════════

def create_chart(df: pd.DataFrame, chart_type: str, x: str, y: str, 
                 color: Optional[str] = None, agg: str = 'sum', 
                 bucket: str = 'month', title: str = '') -> go.Figure:
    """Создание графика Plotly"""
    aggregated = aggregate_data(df, x, y, agg, color, bucket)
    
    if chart_type == 'line':
        fig = px.line(aggregated, x=x, y=y, color=color, title=title,
                     color_discrete_sequence=VIZRO_COLORS, markers=True)
    elif chart_type == 'bar':
        fig = px.bar(aggregated, x=x, y=y, color=color, title=title,
                    color_discrete_sequence=VIZRO_COLORS, barmode='group')
    elif chart_type == 'area':
        fig = px.area(aggregated, x=x, y=y, color=color, title=title,
                     color_discrete_sequence=VIZRO_COLORS)
    elif chart_type == 'pie':
        fig = px.pie(aggregated, names=x, values=y, title=title,
                    color_discrete_sequence=VIZRO_COLORS, hole=0.5)
    elif chart_type == 'scatter':
        fig = px.scatter(df, x=x, y=y, color=color, title=title,
                        color_discrete_sequence=VIZRO_COLORS, trendline='ols')
    else:
        fig = px.line(aggregated, x=x, y=y, title=title, color_discrete_sequence=VIZRO_COLORS)
    
    fig.update_layout(
        template='plotly_white',
        height=350,
        margin=dict(l=40, r=20, t=40, b=40),
        showlegend=True if color else False,
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
    )
    fig.update_xaxes(tickfont=dict(size=10))
    fig.update_yaxes(tickfont=dict(size=10), tickformat=',.0f')
    
    return fig

def create_kpi_sparkline(values: pd.Series, color: str = '#4f46e5') -> go.Figure:
    """Создание sparkline для KPI"""
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        y=values.values,
        mode='lines',
        line=dict(color=color, width=2),
        fill='tozeroy',
        fillcolor=color.replace(')', ', 0.2)').replace('rgb', 'rgba'),
        hoverinfo='skip'
    ))
    fig.update_layout(
        height=60,
        margin=dict(l=0, r=0, t=0, b=0),
        xaxis=dict(showgrid=False, showticklabels=False),
        yaxis=dict(showgrid=False, showticklabels=False),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)'
    )
    return fig

# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN APP
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    # ─── HEADER ──────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="header">
        <div style="display: flex; align-items: center; gap: 12px;">
            <div style="font-size: 32px;">📊</div>
            <div>
                <h1 style="margin: 0; font-size: 24px; font-weight: 700;">BI Platform</h1>
                <p style="margin: 4px 0 0 0; opacity: 0.9; font-size: 13px;">
                    Монолитное приложение • Google Sheets • Excel • Дашборды • Тренды • Аномалии
                </p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # ─── SIDEBAR: DATA SOURCE ────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### 📦 Источник данных")
        
        source = st.radio(
            "Выберите источник:",
            ['🎲 Демо-данные', '🟢 Google Sheets', '📁 Загрузить файл'],
            index=0,
            label_visibility='collapsed'
        )
        
        df = pd.DataFrame()
        
        if source == '🎲 Демо-данные':
            df = generate_sample_data(1400)
            st.success(f"**{len(df):,}** записей загружено")
            st.caption("730 дней • 7 регионов • 7 категорий • 4 канала")
            
        elif source == '🟢 Google Sheets':
            st.info("📋 Откройте доступ: **Файл → Поделиться → Все у кого есть ссылка → Читатель**")
            gs_url = st.text_input(
                "URL Google Sheets:",
                placeholder="https://docs.google.com/spreadsheets/d/.../edit",
                help="Вставьте ссылку на Google таблицу"
            )
            auto_sync = st.checkbox("🔄 Автообновление (60 сек)", value=False)
            
            if gs_url:
                if st.button("Синхронизировать", type="primary", use_container_width=True):
                    with st.spinner("Загрузка данных..."):
                        try:
                            df = load_google_sheet(gs_url)
                            st.success(f"✅ Синхронизировано: **{len(df):,}** строк")
                            st.session_state['gs_url'] = gs_url
                        except Exception as e:
                            st.error(f"❌ Ошибка: {str(e)}")
            
            if auto_sync and 'gs_url' in st.session_state:
                st.caption("🔄 Автообновление активно")
                st.rerun()
                
        else:  # File upload
            uploaded = st.file_uploader(
                "Загрузите файл:",
                type=['csv', 'xlsx', 'xls', 'json'],
                help="Поддерживаются: CSV, Excel, JSON"
            )
            if uploaded:
                with st.spinner("Обработка файла..."):
                    try:
                        df = load_file(uploaded)
                        st.success(f"✅ Загружено: **{len(df):,}** строк")
                    except Exception as e:
                        st.error(f"❌ Ошибка чтения: {str(e)}")
        
        if not df.empty:
            st.divider()
            st.markdown("### 📊 Информация о данных")
            col_meta = st.columns(2)
            with col_meta[0]:
                st.metric("Строк", f"{len(df):,}")
            with col_meta[1]:
                st.metric("Колонок", len(df.columns))
            
            # Column types
            st.caption("**Типы колонок:**")
            type_info = df.dtypes.astype(str).value_counts()
            for t, c in type_info.items():
                st.caption(f"• {t}: {c}")
    
    # ─── MAIN CONTENT ────────────────────────────────────────────────────────
    if df.empty:
        st.info("👈 Выберите источник данных в панели слева")
        st.stop()
    
    # ─── TABS ────────────────────────────────────────────────────────────────
    tabs = st.tabs([
        "🏠 Обзор",
        "🛠 Конструктор графиков",
        "🧮 Сводная таблица",
        "📈 Тренды",
        "⚠ Аномалии",
        "📋 Данные"
    ])
    
    # ─── FILTERS (common for all tabs) ─────────────────────────────────────
    with st.expander("🔍 Фильтры и срезы", expanded=False):
        filter_cols = st.columns(min(4, len(df.columns)))
        filters = {}
        
        num_cols = df.select_dtypes(include=['number']).columns.tolist()
        date_cols = [c for c in df.columns if 'дата' in c.lower() or 'date' in c.lower() or pd.api.types.is_datetime64_dtype(df[c])]
        cat_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()
        
        for i, col in enumerate(df.columns[:4]):
            with filter_cols[i % len(filter_cols)]:
                if col in date_cols:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
                    min_date = df[col].min()
                    max_date = df[col].max()
                    date_range = st.date_input(
                        f"📅 {col}",
                        value=(min_date, max_date),
                        key=f"filter_{col}"
                    )
                    if isinstance(date_range, tuple) and len(date_range) == 2:
                        filters[col] = ('date', date_range)
                elif col in num_cols:
                    min_val = float(df[col].min())
                    max_val = float(df[col].max())
                    range_val = st.slider(
                        f"📊 {col}",
                        min_value=min_val,
                        max_value=max_val,
                        value=(min_val, max_val),
                        key=f"filter_{col}"
                    )
                    filters[col] = ('range', range_val)
                elif col in cat_cols and df[col].nunique() <= 30:
                    options = df[col].dropna().unique().tolist()
                    selected = st.multiselect(
                        f"🏷 {col}",
                        options,
                        default=options,
                        key=f"filter_{col}"
                    )
                    filters[col] = ('select', selected)
        
        # Search
        search = st.text_input("🔎 Поиск по всем колонкам:", placeholder="Введите текст для поиска...")
        
        # Apply filters
        filtered_df = df.copy()
        for col, (ftype, fval) in filters.items():
            if ftype == 'date' and isinstance(fval, tuple):
                filtered_df = filtered_df[
                    (pd.to_datetime(filtered_df[col], errors='coerce') >= pd.Timestamp(fval[0])) &
                    (pd.to_datetime(filtered_df[col], errors='coerce') <= pd.Timestamp(fval[1]) + timedelta(days=1))
                ]
            elif ftype == 'range':
                filtered_df = filtered_df[(filtered_df[col] >= fval[0]) & (filtered_df[col] <= fval[1])]
            elif ftype == 'select':
                filtered_df = filtered_df[filtered_df[col].isin(fval)]
        
        if search:
            for col in filtered_df.columns:
                if filtered_df[col].dtype == 'object':
                    filtered_df = filtered_df[filtered_df[col].astype(str).str.contains(search, case=False, na=False)]
        
        st.caption(f"📊 Показано: **{len(filtered_df):,}** из **{len(df):,}** записей")
    
    # ═══════════════════════════════════════════════════════════════════════════
    #  TAB 1: OVERVIEW (KPI + Auto Charts)
    # ═══════════════════════════════════════════════════════════════════════════
    with tabs[0]:
        st.markdown("### 🏠 Обзор показателей")
        
        # KPI Cards
        kpi_cols = st.columns(4)
        numeric_metrics = num_cols[:4] if num_cols else []
        
        for i, metric in enumerate(numeric_metrics):
            with kpi_cols[i]:
                total = filtered_df[metric].sum()
                avg = filtered_df[metric].mean()
                count = len(filtered_df)
                
                # Trend calculation
                if date_cols:
                    date_col = date_cols[0]
                    sorted_df = filtered_df.sort_values(date_col)
                    half = len(sorted_df) // 2
                    if half > 0:
                        prev_half = sorted_df.iloc[:half][metric].sum()
                        curr_half = sorted_df.iloc[half:][metric].sum()
                        delta = ((curr_half - prev_half) / (abs(prev_half) + 1e-9)) * 100
                    else:
                        delta = 0
                else:
                    delta = 0
                
                st.metric(
                    label=metric,
                    value=f"{total:,.0f}",
                    delta=f"{delta:+.1f}%" if delta != 0 else None,
                    delta_color="normal"
                )
                
                # Sparkline
                if date_cols and len(filtered_df) > 5:
                    date_col = date_cols[0]
                    spark_data = filtered_df.groupby(pd.to_datetime(filtered_df[date_col]).dt.to_period('M'))[metric].sum()
                    spark_fig = create_kpi_sparkline(spark_data, VIZRO_COLORS[i % len(VIZRO_COLORS)])
                    st.plotly_chart(spark_fig, use_container_width=True, key=f"spark_{i}")
        
        st.divider()
        
        # Auto-generated charts
        chart_row = st.columns(2)
        
        with chart_row[0]:
            if date_cols and num_cols:
                date_col = date_cols[0]
                metric = num_cols[0]
                monthly = filtered_df.groupby(pd.to_datetime(filtered_df[date_col]).dt.to_period('M').dt.strftime('%Y-%m'))[metric].sum().reset_index()
                fig = px.area(monthly, x=date_col, y=metric, title=f"📈 Динамика {metric} по месяцам",
                             color_discrete_sequence=[VIZRO_COLORS[0]])
                fig.update_layout(height=350, template='plotly_white')
                st.plotly_chart(fig, use_container_width=True)
        
        with chart_row[1]:
            if cat_cols and num_cols:
                cat_col = cat_cols[0]
                metric = num_cols[0]
                cat_data = filtered_df.groupby(cat_col)[metric].sum().nlargest(8).reset_index()
                fig = px.pie(cat_data, names=cat_col, values=metric, title=f"🥧 Структура по {cat_col}",
                            color_discrete_sequence=VIZRO_COLORS, hole=0.5)
                fig.update_layout(height=350, template='plotly_white')
                st.plotly_chart(fig, use_container_width=True)
        
        # Additional charts
        if len(cat_cols) >= 2 and num_cols:
            chart_row2 = st.columns(2)
            with chart_row2[0]:
                bar_data = filtered_df.groupby(cat_cols[1])[num_cols[0]].sum().nlargest(10).reset_index()
                fig = px.bar(bar_data, x=cat_cols[1], y=num_cols[0], title=f"📊 Топ по {cat_cols[1]}",
                            color_discrete_sequence=[VIZRO_COLORS[2]])
                fig.update_layout(height=300, template='plotly_white', xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
            
            with chart_row2[1]:
                metric2 = num_cols[1] if len(num_cols) > 1 else num_cols[0]
                bar2_data = filtered_df.groupby(cat_cols[0])[metric2].sum().nlargest(10).reset_index()
                fig = px.bar(bar2_data, x=cat_cols[0], y=metric2, title=f"{metric2} по {cat_cols[0]}",
                            color_discrete_sequence=[VIZRO_COLORS[3]])
                fig.update_layout(height=300, template='plotly_white', xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
    
    # ═══════════════════════════════════════════════════════════════════════════
    #  TAB 2: CHART BUILDER
    # ═══════════════════════════════════════════════════════════════════════════
    with tabs[1]:
        st.markdown("### 🛠 Конструктор графиков")
        st.caption("Создавайте любые срезы данных: выбирайте оси, агрегации, группировки")
        
        # Initialize session state for charts
        if 'charts' not in st.session_state:
            st.session_state.charts = []
        
        # Add new chart button
        if st.button("➕ Добавить график", type="primary"):
            st.session_state.charts.append({
                'id': len(st.session_state.charts),
                'type': 'line',
                'x': date_cols[0] if date_cols else df.columns[0],
                'y': num_cols[0] if num_cols else df.columns[1],
                'color': cat_cols[0] if len(cat_cols) > 1 else None,
                'agg': 'sum',
                'bucket': 'month'
            })
        
        # Render charts
        for i, chart in enumerate(st.session_state.charts):
            with st.expander(f"📊 График #{i+1}", expanded=True):
                config_col = st.columns(6)
                
                with config_col[0]:
                    chart_type = st.selectbox(
                        "Тип",
                        ['line', 'bar', 'area', 'pie', 'scatter'],
                        index=['line', 'bar', 'area', 'pie', 'scatter'].index(chart['type']),
                        key=f"chart_{i}_type"
                    )
                    chart['type'] = chart_type
                
                with config_col[1]:
                    x_col = st.selectbox("Ось X", df.columns.tolist(), index=df.columns.tolist().index(chart['x']) if chart['x'] in df.columns else 0, key=f"chart_{i}_x")
                    chart['x'] = x_col
                
                with config_col[2]:
                    y_col = st.selectbox("Метрика Y", num_cols, index=num_cols.index(chart['y']) if chart['y'] in num_cols else 0, key=f"chart_{i}_y")
                    chart['y'] = y_col
                
                with config_col[3]:
                    agg_func = st.selectbox("Агрегация", ['sum', 'avg', 'count', 'min', 'max'], index=['sum', 'avg', 'count', 'min', 'max'].index(chart['agg']), key=f"chart_{i}_agg")
                    chart['agg'] = agg_func
                
                with config_col[4]:
                    color_col = st.selectbox("Разбивка", [None] + cat_cols, index=([None] + cat_cols).index(chart['color']) if chart['color'] in [None] + cat_cols else 0, key=f"chart_{i}_color")
                    chart['color'] = color_col if color_col else None
                
                with config_col[5]:
                    bucket = st.selectbox("Период", ['day', 'week', 'month', 'quarter', 'year'], index=['day', 'week', 'month', 'quarter', 'year'].index(chart['bucket']), key=f"chart_{i}_bucket")
                    chart['bucket'] = bucket
                
                # Create chart
                title = f"{chart['y']} по {chart['x']}"
                if chart['color']:
                    title += f" (по {chart['color']})"
                
                fig = create_chart(filtered_df, chart['type'], chart['x'], chart['y'], 
                                  chart['color'], chart['agg'], chart['bucket'], title)
                st.plotly_chart(fig, use_container_width=True)
                
                # Delete button
                if st.button("🗑 Удалить", key=f"del_chart_{i}"):
                    st.session_state.charts.pop(i)
                    st.rerun()
        
        if not st.session_state.charts:
            st.info("Нажмите **➕ Добавить график** для создания визуализации")
    
    # ═══════════════════════════════════════════════════════════════════════════
    #  TAB 3: PIVOT TABLE
    # ═══════════════════════════════════════════════════════════════════════════
    with tabs[2]:
        st.markdown("### 🧮 Сводная таблица")
        
        pivot_cols = st.columns(4)
        
        with pivot_cols[0]:
            pivot_index = st.selectbox("📍 Строки", cat_cols, key="pivot_index")
        
        with pivot_cols[1]:
            pivot_columns = st.selectbox("📍 Колонки", [None] + [c for c in cat_cols if c != pivot_index], key="pivot_columns")
        
        with pivot_cols[2]:
            pivot_values = st.selectbox("📍 Значения", num_cols, key="pivot_values")
        
        with pivot_cols[3]:
            pivot_agg = st.selectbox("📍 Агрегация", ['sum', 'mean', 'count', 'min', 'max'], key="pivot_agg")
        
        if pivot_index and pivot_values:
            pivot_df = create_pivot_table(filtered_df, pivot_index, pivot_columns, pivot_values, pivot_agg)
            st.dataframe(pivot_df.style.format('{:,.0f}'), use_container_width=True)
            
            # Heatmap visualization
            if pivot_columns:
                fig = px.imshow(pivot_df.iloc[:-1, :-1].T, 
                               labels=dict(x=pivot_index, y=pivot_columns, color=pivot_values),
                               color_continuous_scale='Blues',
                               title=f"🔥 Тепловая карта: {pivot_values}")
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)
    
    # ═══════════════════════════════════════════════════════════════════════════
    #  TAB 4: TRENDS
    # ═══════════════════════════════════════════════════════════════════════════
    with tabs[3]:
        st.markdown("### 📈 Анализ трендов")
        
        if not date_cols:
            st.warning("⚠ В данных не найдены колонки с датами")
        else:
            date_col = date_cols[0]
            filtered_df[date_col] = pd.to_datetime(filtered_df[date_col], errors='coerce')
            
            # Select metric
            trend_metric = st.selectbox("Выберите метрику для анализа:", num_cols, index=0 if num_cols else 0)
            
            # Monthly aggregation
            monthly = filtered_df.groupby(filtered_df[date_col].dt.to_period('M').dt.strftime('%Y-%m'))[trend_metric].agg(['sum', 'mean', 'count']).reset_index()
            monthly.columns = ['Период', 'Сумма', 'Среднее', 'Количество']
            
            # Calculate growth
            monthly['Рост %'] = monthly['Сумма'].pct_change() * 100
            monthly['SMA 3'] = monthly['Сумма'].rolling(window=3).mean()
            
            # Main trend chart
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            fig.add_trace(go.Scatter(x=monthly['Период'], y=monthly['Сумма'], 
                                    name='Сумма', line=dict(color=VIZRO_COLORS[0], width=3)),
                         secondary_y=False)
            fig.add_trace(go.Scatter(x=monthly['Период'], y=monthly['SMA 3'], 
                                    name='Скользящее среднее (3)', line=dict(color=VIZRO_COLORS[1], width=2, dash='dash')),
                         secondary_y=False)
            fig.add_trace(go.Bar(x=monthly['Период'], y=monthly['Рост %'], 
                                name='Рост %', marker_color=monthly['Рост %'].apply(lambda x: VIZRO_COLORS[2] if x >= 0 else VIZRO_COLORS[3])),
                         secondary_y=True)
            
            fig.update_layout(
                title=f"📈 Тренд {trend_metric} по месяцам",
                height=450,
                template='plotly_white',
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
            )
            fig.update_xaxes(title_text="Период")
            fig.update_yaxes(title_text="Сумма", secondary_y=False)
            fig.update_yaxes(title_text="Рост %", secondary_y=True)
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Stats
            st.divider()
            stats_cols = st.columns(4)
            with stats_cols[0]:
                st.metric("Всего периодов", len(monthly))
            with stats_cols[1]:
                st.metric("Средний рост", f"{monthly['Рост %'].mean():+.1f}%")
            with stats_cols[2]:
                st.metric("Макс. рост", f"{monthly['Рост %'].max():+.1f}%")
            with stats_cols[3]:
                st.metric("Мин. рост", f"{monthly['Рост %'].min():+.1f}%")
    
    # ═══════════════════════════════════════════════════════════════════════════
    #  TAB 5: ANOMALIES
    # ═══════════════════════════════════════════════════════════════════════════
    with tabs[4]:
        st.markdown("### ⚠ Детекция аномалий")
        st.caption("Поиск выбросов по правилу: отклонение > 2.5σ от среднего")
        
        if num_cols:
            anomaly_col = st.selectbox("Выберите метрику для анализа:", num_cols)
            threshold = st.slider("Порог (σ):", min_value=1.5, max_value=4.0, value=2.5, step=0.1)
            
            anomalies = detect_anomalies(filtered_df, anomaly_col, threshold)
            
            if len(anomalies) > 0:
                st.success(f"✅ Найдено **{len(anomalies)}** аномальных записей")
                
                # Stats
                mean_val = filtered_df[anomaly_col].mean()
                std_val = filtered_df[anomaly_col].std()
                st.caption(f"Среднее: {mean_val:,.0f} • Стд. отклонение: {std_val:,.0f} • Порог: {threshold * std_val:,.0f}")
                
                st.dataframe(anomalies.style.format('{:,.0f}').background_gradient(subset=[anomaly_col], cmap='Reds'), 
                           use_container_width=True)
                
                # Distribution chart
                fig = make_subplots(rows=1, cols=2, 
                                   subplot_titles=('Распределение', 'Аномалии во времени'))
                
                fig.add_trace(go.Histogram(x=filtered_df[anomaly_col], name='Все данные', 
                                          marker_color=VIZRO_COLORS[0], opacity=0.7),
                            row=1, col=1)
                fig.add_trace(go.Histogram(x=anomalies[anomaly_col], name='Аномалии', 
                                          marker_color=VIZRO_COLORS[3], opacity=0.7),
                            row=1, col=1)
                
                if date_cols:
                    date_col = date_cols[0]
                    fig.add_trace(go.Scatter(x=filtered_df[date_col], y=filtered_df[anomaly_col],
                                            mode='markers', name='Все', marker=dict(size=4, color=VIZRO_COLORS[0])),
                                 row=1, col=2)
                    fig.add_trace(go.Scatter(x=anomalies[date_col], y=anomalies[anomaly_col],
                                            mode='markers', name='Аномалии', marker=dict(size=8, color=VIZRO_COLORS[3])),
                                 row=1, col=2)
                
                fig.update_layout(height=400, template='plotly_white', showlegend=True)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("✅ Аномалий не обнаружено при заданном пороге")
    
    # ═══════════════════════════════════════════════════════════════════════════
    #  TAB 6: DATA TABLE
    # ═══════════════════════════════════════════════════════════════════════════
    with tabs[5]:
        st.markdown("### 📋 Таблица данных")
        
        # Sorting
        sort_col = st.selectbox("Сортировать по:", ['—'] + df.columns.tolist())
        sort_order = st.radio("Порядок:", ['⬆ По возрастанию', '⬇ По убыванию'], index=1)
        
        if sort_col != '—':
            ascending = sort_order == '⬆ По возрастанию'
            display_df = filtered_df.sort_values(sort_col, ascending=ascending)
        else:
            display_df = filtered_df
        
        # Show dataframe
        st.dataframe(display_df.style.format('{:,.2f}'), use_container_width=True, height=500)
        
        # Export
        st.divider()
        st.markdown("#### 📥 Экспорт данных")
        export_cols = st.columns(3)
        
        with export_cols[0]:
            csv = display_df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📄 Скачать CSV",
                data=csv,
                file_name=f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        with export_cols[1]:
            json_str = display_df.to_json(orient='records', force_ascii=False, indent=2)
            st.download_button(
                label="📄 Скачать JSON",
                data=json_str,
                file_name=f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                mime="application/json",
                use_container_width=True
            )
        
        with export_cols[2]:
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                display_df.to_excel(writer, index=False, sheet_name='Data')
            st.download_button(
                label="📄 Скачать Excel",
                data=excel_buffer.getvalue(),
                file_name=f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    # ─── FOOTER ──────────────────────────────────────────────────────────────
    st.divider()
    st.markdown("""
    <div style="text-align: center; color: #64748b; font-size: 12px; padding: 20px 0;">
        <strong>BI Platform</strong> • 
        {rows:,} записей • {cols} колонок • {num} числовых • {dates} дат • {cats} категорий •
        Вдохновлено <a href="https://github.com/mckinsey/vizro" target="_blank">Vizro-AI</a>
    </div>
    """.format(
        rows=len(df), 
        cols=len(df.columns), 
        num=len(num_cols), 
        dates=len(date_cols), 
        cats=len(cat_cols)
    ), unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
#  RUN APP
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    main()
