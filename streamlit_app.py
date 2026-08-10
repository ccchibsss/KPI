"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  BI PLATFORM — монолитное Streamlit-приложение (один файл)                   ║
║  Google Sheets · Excel/CSV · KPI · Фильтры · Срезы · Тренды · Аномалии       ║
╚══════════════════════════════════════════════════════════════════════════════╝

Запуск:
    pip install -r requirements.txt
    streamlit run streamlit_app.py
"""

from __future__ import annotations

import io
import re
from datetime import datetime, timedelta, date
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

# ══════════════════════════════════════════════════════════════════════════════
#  1. КОНФИГУРАЦИЯ И ТЕМА
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(page_title="BI Platform", page_icon="📊", layout="wide",
                   initial_sidebar_state="expanded")

COLORS = ['#6366f1', '#22c55e', '#f59e0b', '#ef4444', '#06b6d4',
          '#a855f7', '#ec4899', '#84cc16', '#f97316', '#14b8a6']

AGG_LABELS = {'sum': 'Сумма', 'mean': 'Среднее', 'count': 'Количество',
              'min': 'Минимум', 'max': 'Максимум', 'median': 'Медиана',
              'nunique': 'Уникальных'}
BUCKET_LABELS = {'День': 'D', 'Неделя': 'W', 'Месяц': 'M', 'Квартал': 'Q', 'Год': 'Y'}
CHART_LABELS = {'Линия': 'line', 'Столбцы': 'bar', 'Область': 'area',
                'Круговая': 'pie', 'Точки': 'scatter', 'Гориз. столбцы': 'hbar',
                'Тепловая карта': 'heatmap', 'Ящик с усами': 'box'}

st.markdown("""
<style>
  .block-container {padding-top: 1.6rem; padding-bottom: 2rem; max-width: 1500px;}
  [data-testid="stMetric"] {background:#fff;border:1px solid #e2e8f0;border-radius:14px;
                            padding:14px 16px;box-shadow:0 1px 2px rgba(15,23,42,.04);}
  [data-testid="stMetricValue"] {font-size:26px;font-weight:700;color:#0f172a;}
  [data-testid="stMetricLabel"] {font-size:11px;font-weight:700;letter-spacing:.06em;
                                 text-transform:uppercase;color:#64748b;}
  .bi-header {background:linear-gradient(120deg,#4338ca 0%,#6d28d9 55%,#7c3aed 100%);
              border-radius:18px;padding:20px 26px;color:#fff;margin-bottom:18px;}
  .bi-header h1 {margin:0;font-size:25px;font-weight:800;letter-spacing:-.02em;}
  .bi-header p  {margin:5px 0 0;font-size:13px;opacity:.85;}
  .stTabs [data-baseweb="tab-list"] {gap:6px;}
  .stTabs [data-baseweb="tab"] {border-radius:10px;padding:8px 16px;background:#f1f5f9;font-weight:600;}
  .stTabs [aria-selected="true"] {background:#4f46e5 !important;color:#fff !important;}
  .pill {display:inline-block;background:#eef2ff;color:#4338ca;border-radius:999px;
         padding:3px 11px;font-size:11px;font-weight:700;margin-right:6px;}
</style>
""", unsafe_allow_html=True)


def hex_to_rgba(hex_color: str, alpha: float = .18) -> str:
    h = hex_color.lstrip('#')
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f'rgba({r},{g},{b},{alpha})'


# ══════════════════════════════════════════════════════════════════════════════
#  2. НОРМАЛИЗАЦИЯ ДАННЫХ  (главный фикс AttributeError)
# ══════════════════════════════════════════════════════════════════════════════

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Приводит имена колонок к уникальным непустым СТРОКАМ."""
    cols, seen = [], {}
    for i, raw in enumerate(df.columns):
        name = '' if raw is None else str(raw).strip()
        if not name or name.lower().startswith('unnamed'):
            name = f'Колонка {i + 1}'
        name = re.sub(r'\s+', ' ', name)
        if name in seen:
            seen[name] += 1
            name = f'{name} ({seen[name]})'
        else:
            seen[name] = 0
        cols.append(name)
    out = df.copy()
    out.columns = cols
    return out


def coerce_numeric(s: pd.Series) -> Optional[pd.Series]:
    """Пытается превратить текстовую колонку в число ('1 234,56', '12%', '1 000 ₽')."""
    if pd.api.types.is_numeric_dtype(s):
        return s
    if not pd.api.types.is_object_dtype(s):
        return None
    txt = (s.astype(str)
             .str.replace('\u00a0', '', regex=False)
             .str.replace(' ', '', regex=False)
             .str.replace('₽', '', regex=False)
             .str.replace('%', '', regex=False)
             .str.replace('$', '', regex=False)
             .str.replace('€', '', regex=False)
             .str.replace(',', '.', regex=False)
             .str.strip())
    txt = txt.replace({'': None, 'nan': None, 'None': None, '-': None, '—': None})
    conv = pd.to_numeric(txt, errors='coerce')
    non_null = s.notna().sum()
    if non_null and conv.notna().sum() / non_null >= 0.85:
        return conv
    return None


def coerce_datetime(s: pd.Series) -> Optional[pd.Series]:
    """Пытается распознать дату; числовые колонки не трогаем."""
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    if pd.api.types.is_numeric_dtype(s) or not pd.api.types.is_object_dtype(s):
        return None
    sample = s.dropna().astype(str).head(200)
    if sample.empty:
        return None
    # эвристика: должно быть похоже на дату
    pattern = re.compile(r'\d{1,4}[-./]\d{1,2}[-./]\d{1,4}')
    if (sample.str.contains(pattern).mean()) < 0.7:
        return None
    for dayfirst in (False, True):
        conv = pd.to_datetime(s, errors='coerce', dayfirst=dayfirst, format='mixed')
        non_null = s.notna().sum()
        if non_null and conv.notna().sum() / non_null >= 0.85:
            return conv
    return None


def prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Нормализует имена, типы и удаляет полностью пустые строки/колонки."""
    if df is None or df.empty:
        return pd.DataFrame()
    df = normalize_columns(df)
    df = df.dropna(axis=1, how='all').dropna(axis=0, how='all')
    for col in df.columns:
        dt = coerce_datetime(df[col])
        if dt is not None:
            df[col] = dt
            continue
        num = coerce_numeric(df[col])
        if num is not None:
            df[col] = num
            continue
        if pd.api.types.is_object_dtype(df[col]):
            df[col] = df[col].astype(str).replace({'nan': None, 'None': None, 'NaT': None})
    return df.reset_index(drop=True)


def classify_columns(df: pd.DataFrame) -> Tuple[List[str], List[str], List[str]]:
    """Возвращает (числовые, даты, категориальные) — все имена гарантированно str."""
    numeric, dates, cats = [], [], []
    for col in df.columns:
        s = df[col]
        if pd.api.types.is_datetime64_any_dtype(s):
            dates.append(col)
        elif pd.api.types.is_numeric_dtype(s):
            numeric.append(col)
        else:
            cats.append(col)
    return numeric, dates, cats


# ══════════════════════════════════════════════════════════════════════════════
#  3. ЗАГРУЗКА ДАННЫХ
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def generate_demo(n: int = 1600, seed: int = 20240715) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    days = pd.date_range('2024-01-01', periods=730, freq='D')
    regions = ['Москва', 'Санкт-Петербург', 'Урал', 'Сибирь', 'Юг', 'Дальний Восток']
    cats = ['Электроника', 'Одежда', 'Продукты', 'Мебель', 'Спорт', 'Красота', 'Книги']
    channels = ['Онлайн', 'Розница', 'Маркетплейс', 'Опт']
    managers = ['Иванов', 'Петрова', 'Сидоров', 'Кузнецова', 'Смирнов',
                'Волкова', 'Козлов', 'Морозова']
    base = {'Электроника': 95000, 'Одежда': 34000, 'Продукты': 18000, 'Мебель': 61000,
            'Спорт': 27000, 'Красота': 22000, 'Книги': 8500}

    idx = rng.integers(0, len(days), n)
    d = days[idx]
    cat = rng.choice(cats, n)
    month = d.month.values
    seasonal = 1 + .35 * np.sin((month - 1) / 12 * 2 * np.pi) + np.where(np.isin(month, [11, 12]), .45, 0)
    trend = 1 + (idx / len(days)) * .5
    qty = rng.integers(1, 15, n)
    base_arr = np.array([base[c] for c in cat])
    revenue = (base_arr * seasonal * trend * (.55 + rng.random(n) * .9) * qty / 6).round().astype(int)
    margin = .12 + rng.random(n) * .26
    profit = (revenue * margin).round().astype(int)

    out = pd.DataFrame({
        'Дата': d,
        'Регион': rng.choice(regions, n),
        'Категория': cat,
        'Канал': rng.choice(channels, n),
        'Менеджер': rng.choice(managers, n),
        'Количество': qty,
        'Выручка': revenue,
        'Прибыль': profit,
        'Расходы': revenue - profit,
    })
    # аномалии для демонстрации детектора
    spike = rng.choice(n, 8, replace=False)
    out.loc[spike, 'Выручка'] = (out.loc[spike, 'Выручка'] * rng.uniform(4, 7, 8)).astype(int)
    return out.sort_values('Дата').reset_index(drop=True)


def gsheet_csv_url(url: str) -> Optional[str]:
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9\-_]+)', url)
    if not m:
        return None
    gid = re.search(r'[#&?]gid=(\d+)', url)
    base = f'https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=csv'
    return base + (f'&gid={gid.group(1)}' if gid else '')


@st.cache_data(ttl=60, show_spinner=False)
def load_gsheet(url: str) -> pd.DataFrame:
    csv_url = gsheet_csv_url(url)
    if not csv_url:
        raise ValueError('Не удалось распознать ссылку. Нужен адрес вида '
                         'https://docs.google.com/spreadsheets/d/<ID>/edit')
    try:
        df = pd.read_csv(csv_url)
    except Exception:
        sid = re.search(r'/spreadsheets/d/([a-zA-Z0-9\-_]+)', url).group(1)
        df = pd.read_csv(f'https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?tqx=out:csv')
    if df.empty:
        raise ValueError('Таблица пуста')
    return df


@st.cache_data(show_spinner=False)
def load_upload(raw: bytes, filename: str, sheet: Optional[str] = None) -> pd.DataFrame:
    name = filename.lower()
    buf = io.BytesIO(raw)
    if name.endswith('.csv') or name.endswith('.txt') or name.endswith('.tsv'):
        for enc in ('utf-8-sig', 'utf-8', 'cp1251'):
            for sep in (None, ';', ',', '\t'):
                try:
                    buf.seek(0)
                    df = pd.read_csv(buf, encoding=enc, sep=sep, engine='python')
                    if df.shape[1] > 0:
                        return df
                except Exception:
                    continue
        raise ValueError('Не удалось прочитать CSV')
    if name.endswith('.json'):
        buf.seek(0)
        return pd.read_json(buf)
    if name.endswith('.parquet'):
        buf.seek(0)
        return pd.read_parquet(buf)
    buf.seek(0)
    return pd.read_excel(buf, sheet_name=sheet or 0)


@st.cache_data(show_spinner=False)
def excel_sheet_names(raw: bytes) -> List[str]:
    try:
        return pd.ExcelFile(io.BytesIO(raw)).sheet_names
    except Exception:
        return []


# ══════════════════════════════════════════════════════════════════════════════
#  4. АНАЛИТИКА
# ══════════════════════════════════════════════════════════════════════════════

def human(n: float) -> str:
    if n is None or (isinstance(n, float) and (np.isnan(n) or np.isinf(n))):
        return '—'
    a = abs(n)
    if a >= 1e9:
        return f'{n / 1e9:.2f} млрд'
    if a >= 1e6:
        return f'{n / 1e6:.2f} млн'
    if a >= 1e4:
        return f'{n / 1e3:.1f} тыс'
    return f'{n:,.0f}'.replace(',', ' ') if a >= 100 else f'{n:,.2f}'.replace(',', ' ')


def bucket_series(s: pd.Series, freq: str) -> pd.Series:
    """Приводит даты к периодам с корректной строковой меткой (в т.ч. кварталы)."""
    per = s.dt.to_period(freq)
    if freq == 'Q':
        return per.astype(str)                       # 2024Q1
    if freq == 'W':
        return per.dt.start_time.dt.strftime('%Y-%m-%d')
    if freq == 'M':
        return per.dt.strftime('%Y-%m')
    if freq == 'Y':
        return per.dt.strftime('%Y')
    return s.dt.strftime('%Y-%m-%d')


def aggregate(df: pd.DataFrame, x: str, y: str, agg: str,
              color: Optional[str] = None, freq: Optional[str] = None,
              top_n: int = 25) -> pd.DataFrame:
    if df.empty or x not in df.columns:
        return pd.DataFrame()
    work = df.copy()
    if freq and pd.api.types.is_datetime64_any_dtype(work[x]):
        work[x] = bucket_series(work[x], freq)
    keys = [x] + ([color] if color and color in work.columns and color != x else [])

    if agg == 'count' or y not in work.columns:
        out = work.groupby(keys, dropna=False).size().reset_index(name='Значение')
    else:
        if not pd.api.types.is_numeric_dtype(work[y]):
            work[y] = pd.to_numeric(work[y], errors='coerce')
        out = work.groupby(keys, dropna=False)[y].agg(agg).reset_index()
        out = out.rename(columns={y: 'Значение'})

    out[x] = out[x].astype(str)
    is_time = bool(freq) and pd.api.types.is_datetime64_any_dtype(df[x])
    if is_time:
        out = out.sort_values(x)
    else:
        order = (out.groupby(x)['Значение'].sum().abs()
                    .sort_values(ascending=False).head(top_n).index)
        out = out[out[x].isin(order)]
        out[x] = pd.Categorical(out[x], categories=list(order), ordered=True)
        out = out.sort_values(x)
    return out


def linear_trend(y: np.ndarray) -> Tuple[float, np.ndarray]:
    n = len(y)
    if n < 2:
        return 0.0, y
    x = np.arange(n)
    slope, intercept = np.polyfit(x, y, 1)
    return float(slope), slope * x + intercept


def period_delta(df: pd.DataFrame, date_col: Optional[str], metric: str) -> Optional[float]:
    """% изменения второй половины периода к первой."""
    if metric not in df.columns or df[metric].dropna().empty:
        return None
    if date_col and date_col in df.columns:
        s = df[[date_col, metric]].dropna().sort_values(date_col)
    else:
        s = df[[metric]].dropna()
    if len(s) < 6:
        return None
    half = len(s) // 2
    prev = s.iloc[:half][metric].sum()
    curr = s.iloc[half:][metric].sum()
    if prev == 0:
        return None
    return (curr - prev) / abs(prev) * 100


def detect_anomalies(df: pd.DataFrame, col: str, method: str, k: float) -> pd.DataFrame:
    s = pd.to_numeric(df[col], errors='coerce')
    valid = s.dropna()
    if len(valid) < 12:
        return df.iloc[0:0].copy()
    if method == 'IQR':
        q1, q3 = valid.quantile(.25), valid.quantile(.75)
        iqr = q3 - q1
        if iqr == 0:
            return df.iloc[0:0].copy()
        lo, hi = q1 - k * iqr, q3 + k * iqr
        mask = (s < lo) | (s > hi)
    elif method == 'Медиана (MAD)':
        med = valid.median()
        mad = (valid - med).abs().median()
        if mad == 0:
            return df.iloc[0:0].copy()
        mask = ((s - med).abs() / (1.4826 * mad)) > k
    else:  # Z-score
        mu, sd = valid.mean(), valid.std()
        if not sd:
            return df.iloc[0:0].copy()
        mask = ((s - mu).abs() / sd) > k
    return df[mask.fillna(False)].copy()


def style_fig(fig: go.Figure, height: int = 360, legend: bool = True) -> go.Figure:
    fig.update_layout(template='plotly_white', height=height,
                      margin=dict(l=10, r=10, t=48, b=10),
                      title_font=dict(size=15),
                      hoverlabel=dict(bgcolor='white', font_size=12),
                      legend=dict(orientation='h', yanchor='bottom', y=1.02,
                                  xanchor='right', x=1, title_text=''),
                      showlegend=legend)
    fig.update_xaxes(showgrid=False, tickfont=dict(size=11))
    fig.update_yaxes(gridcolor='#eef2f7', tickfont=dict(size=11))
    return fig


def build_chart(df: pd.DataFrame, kind: str, x: str, y: str, agg: str,
                color: Optional[str], freq: Optional[str], title: str) -> Optional[go.Figure]:
    if df.empty or x not in df.columns:
        return None

    if kind == 'scatter':
        if not (pd.api.types.is_numeric_dtype(df[x]) and pd.api.types.is_numeric_dtype(df[y])):
            return None
        d = df[[x, y] + ([color] if color and color in df.columns else [])].dropna().head(4000)
        fig = px.scatter(d, x=x, y=y, color=color, title=title,
                         color_discrete_sequence=COLORS, opacity=.65,
                         trendline='ols' if len(d) > 3 else None,
                         trendline_color_override='#ef4444')
        return style_fig(fig, legend=bool(color))

    if kind == 'box':
        if not pd.api.types.is_numeric_dtype(df[y]):
            return None
        fig = px.box(df, x=x, y=y, color=color, title=title, color_discrete_sequence=COLORS)
        return style_fig(fig, legend=bool(color))

    if kind == 'heatmap':
        if not color or color not in df.columns:
            return None
        work = df.copy()
        if freq and pd.api.types.is_datetime64_any_dtype(work[x]):
            work[x] = bucket_series(work[x], freq)
        pt = pd.pivot_table(work, index=color, columns=x, values=y,
                            aggfunc=agg if agg != 'count' else 'size', fill_value=0)
        fig = px.imshow(pt, color_continuous_scale='Indigo', aspect='auto', title=title,
                        labels=dict(color=y))
        return style_fig(fig, height=420, legend=False)

    data = aggregate(df, x, y, agg, color, freq)
    if data.empty:
        return None
    cname = color if color and color in data.columns and color != x else None

    if kind == 'pie':
        agg_pie = data.groupby(x, observed=True)['Значение'].sum().reset_index()
        fig = px.pie(agg_pie.head(10), names=x, values='Значение', hole=.55, title=title,
                     color_discrete_sequence=COLORS)
        fig.update_traces(textposition='inside', textinfo='percent+label')
        return style_fig(fig)

    if kind == 'bar':
        fig = px.bar(data, x=x, y='Значение', color=cname, title=title,
                     color_discrete_sequence=COLORS, barmode='group')
    elif kind == 'hbar':
        fig = px.bar(data, y=x, x='Значение', color=cname, title=title, orientation='h',
                     color_discrete_sequence=COLORS, barmode='group')
    elif kind == 'area':
        fig = px.area(data, x=x, y='Значение', color=cname, title=title,
                      color_discrete_sequence=COLORS)
    else:
        fig = px.line(data, x=x, y='Значение', color=cname, title=title, markers=len(data) < 60,
                      color_discrete_sequence=COLORS)
    return style_fig(fig, legend=bool(cname))


# ══════════════════════════════════════════════════════════════════════════════
#  5. БОКОВАЯ ПАНЕЛЬ — ИСТОЧНИК ДАННЫХ
# ══════════════════════════════════════════════════════════════════════════════

def sidebar_source() -> Tuple[pd.DataFrame, str]:
    st.sidebar.markdown('### 📦 Источник данных')
    src = st.sidebar.radio('Источник', ['🎲 Демо-данные', '🟢 Google Sheets', '📁 Excel / CSV'],
                           label_visibility='collapsed')
    raw, label = pd.DataFrame(), ''

    if src == '🎲 Демо-данные':
        rows = st.sidebar.slider('Объём демо-выборки', 300, 5000, 1600, 100)
        raw = generate_demo(rows)
        label = f'Демо-данные · {rows:,} строк'.replace(',', ' ')

    elif src == '🟢 Google Sheets':
        st.sidebar.caption('Файл → Поделиться → «Все, у кого есть ссылка» → Читатель')
        url = st.sidebar.text_input('Ссылка на таблицу',
                                    value=st.session_state.get('gs_url', ''),
                                    placeholder='https://docs.google.com/spreadsheets/d/…')
        c1, c2 = st.sidebar.columns([1, 1])
        do_load = c1.button('🔄 Загрузить', use_container_width=True, type='primary')
        if c2.button('♻️ Сброс кэша', use_container_width=True):
            load_gsheet.clear()
            do_load = True
        auto = st.sidebar.checkbox('Автообновление', value=False)
        if auto:
            interval = st.sidebar.select_slider('Интервал', [30, 60, 120, 300, 600], value=60,
                                                format_func=lambda v: f'{v} сек')
            try:
                st.autorefresh(interval=interval * 1000, key='auto_refresh')
            except Exception:
                st.sidebar.caption('⏱ Обновление при следующем действии')
        if url and (do_load or auto or st.session_state.get('gs_loaded')):
            try:
                raw = load_gsheet(url)
                st.session_state['gs_url'] = url
                st.session_state['gs_loaded'] = True
                label = f'Google Sheets · {len(raw):,} строк'.replace(',', ' ')
                st.sidebar.success(f'Загружено строк: {len(raw):,}'.replace(',', ' '))
            except Exception as exc:
                st.sidebar.error(f'Ошибка: {exc}')

    else:
        up = st.sidebar.file_uploader('Файл', type=['xlsx', 'xls', 'csv', 'txt', 'tsv', 'json', 'parquet'],
                                      label_visibility='collapsed')
        if up is not None:
            data = up.getvalue()
            sheet = None
            if up.name.lower().endswith(('.xlsx', '.xls')):
                names = excel_sheet_names(data)
                if len(names) > 1:
                    sheet = st.sidebar.selectbox('Лист', names)
            try:
                raw = load_upload(data, up.name, sheet)
                label = f'{up.name} · {len(raw):,} строк'.replace(',', ' ')
                st.sidebar.success(f'Загружено строк: {len(raw):,}'.replace(',', ' '))
            except Exception as exc:
                st.sidebar.error(f'Ошибка чтения: {exc}')

    return raw, label


# ══════════════════════════════════════════════════════════════════════════════
#  6. ФИЛЬТРЫ
# ══════════════════════════════════════════════════════════════════════════════

def build_filters(df: pd.DataFrame, numeric: List[str], dates: List[str],
                  cats: List[str]) -> pd.DataFrame:
    st.markdown('#### 🔍 Фильтры и срезы')
    out = df

    with st.container(border=True):
        search = st.text_input('Поиск по всем колонкам', placeholder='введите текст…',
                               label_visibility='collapsed')
        if search:
            mask = pd.Series(False, index=out.index)
            for col in out.columns:
                mask |= out[col].astype(str).str.contains(search, case=False, na=False, regex=False)
            out = out[mask]

        if dates:
            cols = st.columns(min(3, len(dates)))
            for i, col in enumerate(dates):
                with cols[i % len(cols)]:
                    series = pd.to_datetime(df[col], errors='coerce').dropna()
                    if series.empty:
                        continue
                    lo, hi = series.min().date(), series.max().date()
                    if lo == hi:
                        continue
                    picked = st.date_input(f'📅 {col}', value=(lo, hi), min_value=lo,
                                           max_value=hi, key=f'flt_date_{col}')
                    if isinstance(picked, (tuple, list)) and len(picked) == 2:
                        s = pd.to_datetime(out[col], errors='coerce')
                        out = out[(s >= pd.Timestamp(picked[0])) &
                                  (s < pd.Timestamp(picked[1]) + timedelta(days=1))]

        selectable = [c for c in cats if 1 < df[c].nunique(dropna=True) <= 200]
        if selectable:
            cols = st.columns(min(4, len(selectable)))
            for i, col in enumerate(selectable):
                with cols[i % len(cols)]:
                    opts = sorted(df[col].dropna().astype(str).unique().tolist())
                    picked = st.multiselect(f'🏷 {col}', opts, default=[], key=f'flt_cat_{col}',
                                            placeholder='все значения')
                    if picked:
                        out = out[out[col].astype(str).isin(picked)]

        if numeric:
            with st.expander('📊 Числовые диапазоны'):
                cols = st.columns(min(3, len(numeric)))
                for i, col in enumerate(numeric):
                    series = pd.to_numeric(df[col], errors='coerce').dropna()
                    if series.empty or series.min() == series.max():
                        continue
                    lo, hi = float(series.min()), float(series.max())
                    with cols[i % len(cols)]:
                        rng = st.slider(col, lo, hi, (lo, hi), key=f'flt_num_{col}')
                        if rng != (lo, hi):
                            s = pd.to_numeric(out[col], errors='coerce')
                            out = out[(s >= rng[0]) & (s <= rng[1])]

        left, right = st.columns([3, 1])
        left.caption(f'Отобрано **{len(out):,}** из **{len(df):,}** строк'.replace(',', ' '))
        if right.button('Сбросить фильтры', use_container_width=True):
            for key in [k for k in st.session_state if k.startswith('flt_')]:
                del st.session_state[key]
            st.rerun()
    return out


# ══════════════════════════════════════════════════════════════════════════════
#  7. ВКЛАДКИ
# ══════════════════════════════════════════════════════════════════════════════

def tab_overview(fdf: pd.DataFrame, numeric: List[str], dates: List[str], cats: List[str]):
    date_col = dates[0] if dates else None
    metrics = numeric[:4]
    if not metrics:
        st.info('В данных нет числовых колонок для расчёта показателей.')
    else:
        cols = st.columns(len(metrics))
        for i, m in enumerate(metrics):
            s = pd.to_numeric(fdf[m], errors='coerce')
            delta = period_delta(fdf, date_col, m)
            with cols[i]:
                st.metric(m, human(s.sum()),
                          f'{delta:+.1f}%' if delta is not None else None)
                if date_col:
                    spark = (fdf[[date_col, m]].dropna()
                             .assign(_p=lambda d: bucket_series(d[date_col], 'M'))
                             .groupby('_p')[m].sum())
                    if len(spark) > 2:
                        fig = go.Figure(go.Scatter(
                            y=spark.values, mode='lines', line=dict(color=COLORS[i % len(COLORS)], width=2),
                            fill='tozeroy', fillcolor=hex_to_rgba(COLORS[i % len(COLORS)]),
                            hovertemplate='%{y:,.0f}<extra></extra>'))
                        fig.update_layout(height=64, margin=dict(l=0, r=0, t=0, b=0),
                                          xaxis=dict(visible=False), yaxis=dict(visible=False),
                                          paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                                          showlegend=False)
                        st.plotly_chart(fig, use_container_width=True,
                                        config={'displayModeBar': False}, key=f'spark_{i}')

    st.divider()
    y = numeric[0] if numeric else None
    if not y:
        return
    row1 = st.columns(2)
    with row1[0]:
        if date_col:
            fig = build_chart(fdf, 'area', date_col, y, 'sum', None, 'M', f'Динамика · {y} по месяцам')
            if fig:
                st.plotly_chart(fig, use_container_width=True)
    with row1[1]:
        if cats:
            fig = build_chart(fdf, 'pie', cats[0], y, 'sum', None, None, f'Структура · {y} по «{cats[0]}»')
            if fig:
                st.plotly_chart(fig, use_container_width=True)

    row2 = st.columns(2)
    with row2[0]:
        if len(cats) > 1:
            fig = build_chart(fdf, 'bar', cats[1], y, 'sum', None, None, f'{y} по «{cats[1]}»')
            if fig:
                st.plotly_chart(fig, use_container_width=True)
    with row2[1]:
        if cats and len(numeric) > 1:
            fig = build_chart(fdf, 'hbar', cats[0], numeric[1], 'sum', None, None,
                              f'{numeric[1]} по «{cats[0]}»')
            if fig:
                st.plotly_chart(fig, use_container_width=True)


def tab_builder(fdf: pd.DataFrame, all_cols: List[str], numeric: List[str],
                dates: List[str], cats: List[str]):
    st.markdown('#### 🛠 Конструктор графиков')
    st.caption('Любые срезы: ось, метрика, агрегация, разбивка по цвету, гранулярность времени.')

    if 'charts' not in st.session_state:
        default_x = dates[0] if dates else (cats[0] if cats else all_cols[0])
        st.session_state.charts = [{
            'kind': 'Линия' if dates else 'Столбцы', 'x': default_x,
            'y': numeric[0] if numeric else all_cols[0], 'agg': 'sum',
            'color': '— нет —', 'bucket': 'Месяц'}]

    if st.button('➕ Добавить график', type='primary'):
        st.session_state.charts.append({
            'kind': 'Столбцы', 'x': cats[0] if cats else all_cols[0],
            'y': numeric[0] if numeric else all_cols[0], 'agg': 'sum',
            'color': '— нет —', 'bucket': 'Месяц'})

    remove_idx = None
    for i, cfg in enumerate(st.session_state.charts):
        with st.container(border=True):
            c = st.columns([1.1, 1.2, 1.2, 1, 1.2, 1, .5])
            cfg['kind'] = c[0].selectbox('Тип', list(CHART_LABELS), key=f'k{i}',
                                         index=list(CHART_LABELS).index(cfg['kind']))
            cfg['x'] = c[1].selectbox('Ось X', all_cols, key=f'x{i}',
                                      index=all_cols.index(cfg['x']) if cfg['x'] in all_cols else 0)
            y_opts = numeric or all_cols
            cfg['y'] = c[2].selectbox('Метрика Y', y_opts, key=f'y{i}',
                                      index=y_opts.index(cfg['y']) if cfg['y'] in y_opts else 0)
            cfg['agg'] = c[3].selectbox('Агрегация', list(AGG_LABELS), key=f'a{i}',
                                        format_func=lambda v: AGG_LABELS[v],
                                        index=list(AGG_LABELS).index(cfg['agg']))
            color_opts = ['— нет —'] + [c_ for c_ in cats if c_ != cfg['x']]
            cfg['color'] = c[4].selectbox('Разбивка', color_opts, key=f'c{i}',
                                          index=color_opts.index(cfg['color'])
                                          if cfg['color'] in color_opts else 0)
            is_date_x = cfg['x'] in dates
            cfg['bucket'] = c[5].selectbox('Период', list(BUCKET_LABELS), key=f'b{i}',
                                           index=list(BUCKET_LABELS).index(cfg['bucket']),
                                           disabled=not is_date_x)
            c[6].write('')
            if c[6].button('🗑', key=f'del{i}', help='Удалить график'):
                remove_idx = i

            color = None if cfg['color'] == '— нет —' else cfg['color']
            freq = BUCKET_LABELS[cfg['bucket']] if is_date_x else None
            title = f"{AGG_LABELS[cfg['agg']]} «{cfg['y']}» по «{cfg['x']}»"
            fig = build_chart(fdf, CHART_LABELS[cfg['kind']], cfg['x'], cfg['y'],
                              cfg['agg'], color, freq, title)
            if fig:
                st.plotly_chart(fig, use_container_width=True, key=f'chart{i}')
            else:
                st.warning('Недостаточно данных для такой комбинации. Измените параметры.')

    if remove_idx is not None:
        st.session_state.charts.pop(remove_idx)
        st.rerun()


def tab_pivot(fdf: pd.DataFrame, numeric: List[str], dates: List[str], cats: List[str]):
    st.markdown('#### 🧮 Сводная таблица')
    dims = cats + dates
    if not dims or not numeric:
        st.info('Нужны хотя бы одна категориальная и одна числовая колонка.')
        return

    c = st.columns(5)
    row = c[0].selectbox('Строки', dims)
    col = c[1].selectbox('Колонки', ['— нет —'] + [d for d in dims if d != row])
    val = c[2].selectbox('Значения', numeric)
    agg = c[3].selectbox('Агрегация', list(AGG_LABELS), format_func=lambda v: AGG_LABELS[v])
    freq_label = c[4].selectbox('Период дат', list(BUCKET_LABELS), index=2,
                                disabled=(row not in dates and col not in dates))

    work = fdf.copy()
    for dim in (row, col):
        if dim in dates:
            work[dim] = bucket_series(pd.to_datetime(work[dim], errors='coerce'),
                                      BUCKET_LABELS[freq_label])

    try:
        pt = pd.pivot_table(work, index=row, columns=None if col == '— нет —' else col,
                            values=val, aggfunc=agg, fill_value=0,
                            margins=True, margins_name='Итого')
        st.dataframe(pt.style.format('{:,.0f}'), use_container_width=True, height=460)

        body = pt.drop(index='Итого', errors='ignore')
        if 'Итого' in body.columns:
            body = body.drop(columns='Итого')
        if col != '— нет —' and not body.empty:
            fig = px.imshow(body, color_continuous_scale='Indigo', aspect='auto',
                            title=f'Тепловая карта · {AGG_LABELS[agg]} «{val}»',
                            labels=dict(color=val))
            st.plotly_chart(style_fig(fig, height=430, legend=False), use_container_width=True)

        st.download_button('⬇️ Скачать сводную (CSV)',
                           pt.to_csv().encode('utf-8-sig'),
                           file_name='pivot.csv', mime='text/csv')
    except Exception as exc:
        st.error(f'Не удалось построить сводную: {exc}')


def tab_trends(fdf: pd.DataFrame, numeric: List[str], dates: List[str], cats: List[str]):
    st.markdown('#### 📈 Тренды и динамика')
    if not dates or not numeric:
        st.info('Для анализа трендов нужны колонка с датой и числовая метрика.')
        return

    c = st.columns(4)
    date_col = c[0].selectbox('Колонка даты', dates)
    metric = c[1].selectbox('Метрика', numeric)
    freq_label = c[2].selectbox('Гранулярность', list(BUCKET_LABELS), index=2)
    window = c[3].slider('Окно скольз. среднего', 2, 12, 3)

    work = fdf[[date_col, metric]].dropna().copy()
    work[date_col] = pd.to_datetime(work[date_col], errors='coerce')
    work = work.dropna()
    if work.empty:
        st.info('Нет данных после фильтрации.')
        return

    work['Период'] = bucket_series(work[date_col], BUCKET_LABELS[freq_label])
    ts = work.groupby('Период')[metric].sum().reset_index().sort_values('Период')
    if len(ts) < 3:
        st.info('Слишком мало периодов для анализа тренда.')
        return

    ts['SMA'] = ts[metric].rolling(window, min_periods=1).mean()
    ts['Рост %'] = ts[metric].pct_change() * 100
    slope, fit = linear_trend(ts[metric].to_numpy(dtype=float))

    k = st.columns(4)
    total_growth = (ts[metric].iloc[-1] - ts[metric].iloc[0]) / abs(ts[metric].iloc[0]) * 100 \
        if ts[metric].iloc[0] else 0
    k[0].metric('Периодов', len(ts))
    k[1].metric('Итоговый рост', f'{total_growth:+.1f}%')
    k[2].metric('Средний рост за период', f'{ts["Рост %"].mean(skipna=True):+.1f}%')
    k[3].metric('Наклон тренда', ('▲ ' if slope >= 0 else '▼ ') + human(abs(slope)))

    fig = make_subplots(specs=[[{'secondary_y': True}]])
    fig.add_trace(go.Bar(x=ts['Период'], y=ts['Рост %'], name='Рост, %',
                         marker_color=np.where(ts['Рост %'] >= 0, '#bbf7d0', '#fecaca'),
                         hovertemplate='%{y:.1f}%<extra></extra>'), secondary_y=True)
    fig.add_trace(go.Scatter(x=ts['Период'], y=ts[metric], name=metric, mode='lines+markers',
                             line=dict(color=COLORS[0], width=3)), secondary_y=False)
    fig.add_trace(go.Scatter(x=ts['Период'], y=ts['SMA'], name=f'SMA({window})',
                             line=dict(color=COLORS[5], width=2, dash='dash')), secondary_y=False)
    fig.add_trace(go.Scatter(x=ts['Период'], y=fit, name='Линейный тренд',
                             line=dict(color='#94a3b8', width=2, dash='dot')), secondary_y=False)
    fig.update_yaxes(title_text=metric, secondary_y=False)
    fig.update_yaxes(title_text='Рост, %', secondary_y=True, showgrid=False)
    st.plotly_chart(style_fig(fig, height=470), use_container_width=True)

    if cats:
        st.markdown('##### Сравнение динамики по срезу')
        dim = st.selectbox('Разрез', cats, key='trend_dim')
        cmp = build_chart(fdf, 'line', date_col, metric, 'sum', dim,
                          BUCKET_LABELS[freq_label], f'{metric} по «{dim}»')
        if cmp:
            st.plotly_chart(cmp, use_container_width=True)

    with st.expander('📄 Таблица по периодам'):
        st.dataframe(ts, use_container_width=True, height=320)


def tab_anomalies(fdf: pd.DataFrame, numeric: List[str], dates: List[str]):
    st.markdown('#### ⚠️ Поиск аномалий')
    if not numeric:
        st.info('Нет числовых колонок.')
        return

    c = st.columns(3)
    col = c[0].selectbox('Метрика', numeric, key='anom_col')
    method = c[1].selectbox('Метод', ['Z-score', 'IQR', 'Медиана (MAD)'])
    k = c[2].slider('Чувствительность (порог)', 1.0, 5.0, 2.5, .1)

    anom = detect_anomalies(fdf, col, method, k)
    s = pd.to_numeric(fdf[col], errors='coerce').dropna()

    m = st.columns(4)
    m[0].metric('Найдено аномалий', f'{len(anom):,}'.replace(',', ' '))
    m[1].metric('Доля', f'{(len(anom) / max(len(fdf), 1) * 100):.2f}%')
    m[2].metric('Среднее', human(s.mean()) if not s.empty else '—')
    m[3].metric('Ст. отклонение', human(s.std()) if not s.empty else '—')

    if anom.empty:
        st.success('Аномалий не обнаружено при текущем пороге.')
        return

    date_col = dates[0] if dates else None
    if date_col:
        base = fdf[[date_col, col]].dropna()
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=base[date_col], y=base[col], mode='markers', name='Норма',
                                 marker=dict(size=5, color='#c7d2fe')))
        fig.add_trace(go.Scatter(x=anom[date_col], y=anom[col], mode='markers', name='Аномалия',
                                 marker=dict(size=10, color='#ef4444',
                                             line=dict(width=1, color='#7f1d1d'))))
        st.plotly_chart(style_fig(fig, height=380), use_container_width=True)
    else:
        hist = px.histogram(fdf, x=col, nbins=50, color_discrete_sequence=[COLORS[0]])
        st.plotly_chart(style_fig(hist, height=340, legend=False), use_container_width=True)

    st.dataframe(anom, use_container_width=True, height=380)
    st.download_button('⬇️ Скачать аномалии (CSV)',
                       anom.to_csv(index=False).encode('utf-8-sig'),
                       file_name='anomalies.csv', mime='text/csv')


def tab_data(fdf: pd.DataFrame, all_cols: List[str]):
    st.markdown('#### 📋 Данные')
    c = st.columns([2, 1, 1])
    shown = c[0].multiselect('Колонки', all_cols, default=all_cols)
    sort_col = c[1].selectbox('Сортировка', ['— без сортировки —'] + all_cols)
    order = c[2].radio('Порядок', ['По убыванию', 'По возрастанию'], horizontal=True)

    view = fdf[shown] if shown else fdf
    if sort_col != '— без сортировки —' and sort_col in view.columns:
        view = view.sort_values(sort_col, ascending=(order == 'По возрастанию'))

    st.dataframe(view, use_container_width=True, height=520)

    st.markdown('##### 📥 Экспорт')
    e = st.columns(3)
    e[0].download_button('CSV', view.to_csv(index=False).encode('utf-8-sig'),
                         file_name='export.csv', mime='text/csv', use_container_width=True)
    e[1].download_button('JSON', view.to_json(orient='records', force_ascii=False, indent=2,
                                              date_format='iso').encode('utf-8'),
                         file_name='export.json', mime='application/json', use_container_width=True)
    buf = io.BytesIO()
    try:
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            view.to_excel(writer, index=False, sheet_name='Данные')
        e[2].download_button('Excel', buf.getvalue(), file_name='export.xlsx',
                             mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                             use_container_width=True)
    except Exception:
        e[2].caption('Для экспорта в Excel установите openpyxl')

    with st.expander('🔬 Профиль данных'):
        prof = pd.DataFrame({
            'Колонка': fdf.columns,
            'Тип': [str(fdf[c].dtype) for c in fdf.columns],
            'Заполнено': [int(fdf[c].notna().sum()) for c in fdf.columns],
            'Пропуски': [int(fdf[c].isna().sum()) for c in fdf.columns],
            'Уникальных': [int(fdf[c].nunique(dropna=True)) for c in fdf.columns],
            'Пример': [str(fdf[c].dropna().iloc[0]) if fdf[c].notna().any() else '—'
                       for c in fdf.columns],
        })
        st.dataframe(prof, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
#  8. ТОЧКА ВХОДА
# ══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    st.markdown("""
    <div class="bi-header">
      <h1>📊 BI Platform</h1>
      <p>Монолитное приложение · Google Sheets · Excel/CSV · KPI · срезы · тренды · аномалии</p>
    </div>""", unsafe_allow_html=True)

    raw, label = sidebar_source()
    if raw is None or raw.empty:
        st.info('👈 Выберите источник данных на боковой панели: демо-набор, '
                'ссылка на Google Sheets или файл Excel/CSV.')
        st.stop()

    df = prepare_dataframe(raw)
    if df.empty:
        st.warning('Данные загружены, но после очистки таблица пуста.')
        st.stop()

    numeric, dates, cats = classify_columns(df)
    all_cols = list(df.columns)

    st.sidebar.divider()
    st.sidebar.markdown('### 📑 Структура')
    st.sidebar.caption(label)
    s1, s2 = st.sidebar.columns(2)
    s1.metric('Строк', f'{len(df):,}'.replace(',', ' '))
    s2.metric('Колонок', len(all_cols))
    st.sidebar.markdown(
        f'<span class="pill">🔢 числовых: {len(numeric)}</span>'
        f'<span class="pill">📅 дат: {len(dates)}</span>'
        f'<span class="pill">🏷 категорий: {len(cats)}</span>',
        unsafe_allow_html=True)

    fdf = build_filters(df, numeric, dates, cats)
    if fdf.empty:
        st.warning('Под текущие фильтры не подходит ни одна строка. Ослабьте условия.')
        st.stop()

    tabs = st.tabs(['🏠 Обзор', '🛠 Конструктор', '🧮 Сводная',
                    '📈 Тренды', '⚠️ Аномалии', '📋 Данные'])
    with tabs[0]:
        tab_overview(fdf, numeric, dates, cats)
    with tabs[1]:
        tab_builder(fdf, all_cols, numeric, dates, cats)
    with tabs[2]:
        tab_pivot(fdf, numeric, dates, cats)
    with tabs[3]:
        tab_trends(fdf, numeric, dates, cats)
    with tabs[4]:
        tab_anomalies(fdf, numeric, dates)
    with tabs[5]:
        tab_data(fdf, all_cols)

    st.divider()
    st.caption(f'BI Platform · источник: {label} · обновлено '
               f'{datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')


if __name__ == '__main__':
    main()
