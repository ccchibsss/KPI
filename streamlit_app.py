# -*- coding: utf-8 -*-
"""
ADAPTIVE BI — монолитное Streamlit-приложение.
Любые данные (Excel / CSV / JSON / Google Sheets / ссылка / буфер) —
авто-типы столбцов, умные фильтры, любые графики как у сводных таблиц.

Запуск:  streamlit run streamlit_app.py
"""

from __future__ import annotations

import io
import json
import re
import warnings
from datetime import date

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

warnings.filterwarnings("ignore")

# =============================================================================
# CONFIG & STYLES
# =============================================================================

st.set_page_config(
    page_title="Adaptive BI",
    page_icon="M",
    layout="wide",
    initial_sidebar_state="expanded",
)

PALETTE = [
    "#a855f7", "#06b6d4", "#10b981", "#f59e0b", "#ec4899", "#6366f1",
    "#84cc16", "#f43f5e", "#14b8a6", "#eab308", "#8b5cf6", "#22d3ee",
]

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Unbounded:wght@500;700;900&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');

/* === базовый цвет текста — более светлый, чем был === */
html, body, [class*="css"], p, span, label, div, small, code { color: #e8eaf6; }
body { font-family: 'Plus Jakarta Sans', sans-serif; font-size: 14px; }

[data-testid="stAppViewContainer"] {
  background:
    radial-gradient(1100px 600px at 88% -12%, rgba(124,58,237,.22), transparent 62%),
    radial-gradient(820px 520px at -12% 112%, rgba(6,182,212,.18), transparent 60%),
    #050816;
  color: #f1f5f9;
}
[data-testid="stHeader"] { background: rgba(5,8,22,.85); backdrop-filter: blur(12px); }
[data-testid="stMain"] { color: #f1f5f9; }
[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #0a0d22 0%, #07091a 100%);
  border-right: 1px solid rgba(168,85,247,.18);
}
[data-testid="stSidebar"] *,
[data-testid="stSidebar"] p, [data-testid="stSidebar"] span,
[data-testid="stSidebar"] label, [data-testid="stSidebar"] small {
  color: #d8dceb !important;
}

/* === заголовки === */
h1, h2, h3, h4, h5, h6 {
  font-family: 'Unbounded', sans-serif !important;
  letter-spacing: -.01em; color: #ffffff !important;
}
h1 { font-size: 1.75rem !important; font-weight: 800 !important; }
h2 { font-size: 1.3rem  !important; font-weight: 800 !important; }
h3 { font-size: 1.05rem !important; font-weight: 700 !important; }
.subtitle { color: #c4c8e0; font-size: .9rem; margin-top: -8px; margin-bottom: 18px; }

/* === KPI === */
[data-testid="stMetric"] {
  background: linear-gradient(135deg, rgba(28,34,72,.95), rgba(15,18,38,.95));
  border: 1px solid rgba(168,85,247,.25);
  border-radius: 18px; padding: 18px 20px;
  box-shadow: 0 16px 36px -16px rgba(0,0,0,.9), inset 0 1px 0 rgba(255,255,255,.07);
}
[data-testid="stMetricLabel"] {
  color: #c4c8e0 !important; font-size: .75rem !important;
  text-transform: uppercase; letter-spacing: .1em; font-weight: 700 !important;
}
[data-testid="stMetricValue"] {
  font-family: 'JetBrains Mono', monospace !important;
  font-size: 1.55rem !important; color: #ffffff !important; font-weight: 700 !important;
}
[data-testid="stMetricDelta"] { color: #6ee7b7 !important; font-weight: 600 !important; }
[data-testid="stCaptionContainer"] { color: #c4c8e0 !important; }
.stCaption, [data-testid="stCaptionContainer"] p { color: #c4c8e0 !important; }

/* === табы === */
.stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 1px solid rgba(255,255,255,.12); }
.stTabs [data-baseweb="tab"] {
  background: rgba(255,255,255,.03); border-radius: 12px 12px 0 0;
  padding: 9px 18px; font-size: .85rem; font-weight: 700; color: #a0a6c4 !important;
  border: 1px solid transparent;
}
.stTabs [aria-selected="true"] {
  background: rgba(168,85,247,.22) !important; color: #ffffff !important;
  border: 1px solid rgba(168,85,247,.4);
  border-bottom: 1px solid transparent;
  box-shadow: 0 -3px 0 #a855f7 inset;
}

/* === кнопки === */
.stButton > button, .stDownloadButton > button {
  border-radius: 12px !important; border: 1px solid rgba(168,85,247,.35) !important;
  background: rgba(255,255,255,.06) !important; color: #f1f5f9 !important;
  font-weight: 700 !important; font-size: .82rem !important; padding: .5rem 1rem !important;
}
.stButton > button:hover, .stDownloadButton > button:hover {
  border-color: #c084fc !important; background: rgba(168,85,247,.25) !important;
  color: #ffffff !important; transform: translateY(-1px);
}
.stButton > button[kind="primary"] {
  background: linear-gradient(135deg,#7e22ce,#ec4899) !important; border: none !important;
  color: #ffffff !important; box-shadow: 0 8px 22px -8px rgba(168,85,247,.6);
}
.stButton > button[kind="primary"]:hover {
  background: linear-gradient(135deg,#9333ea,#f472b6) !important;
}

/* === поля ввода === */
[data-baseweb="input"], [data-baseweb="select"] > div, [data-baseweb="textarea"] {
  background-color: #1a1f44 !important; color: #f1f5f9 !important;
  border: 1px solid rgba(168,85,247,.25) !important; border-radius: 12px !important;
}
[data-baseweb="input"] input, [data-baseweb="select"] input,
[data-baseweb="select"] [data-testid="stMarkdownContainer"] p,
.stTextArea textarea, .stTextInput input, .stNumberInput input {
  color: #f1f5f9 !important; -webkit-text-fill-color: #f1f5f9 !important;
}
[data-baseweb="select"] [data-baseweb="tag"] {
  background: rgba(168,85,247,.25) !important; color: #f1f5f9 !important;
}
[data-baseweb="select"] svg { color: #c4c8e0 !important; }
[data-testid="stFileUploaderDropzone"] {
  background: #11163a !important; border: 2px dashed rgba(168,85,247,.4) !important;
  border-radius: 14px !important;
}
[data-testid="stFileUploaderDropzone"] * { color: #d8dceb !important; }

/* === radio / checkbox / slider === */
.stRadio label, .stCheckbox label, .stToggle label { color: #f1f5f9 !important; font-weight: 600; }
[data-testid="stSlider"] [data-baseweb="slider"] div[role="slider"] {
  background: #c084fc !important; border-color: #f1f5f9 !important;
}
[data-testid="stDateInput"] input { color: #f1f5f9 !important; -webkit-text-fill-color: #f1f5f9 !important; }

/* === expander === */
[data-testid="stExpander"] {
  border: 1px solid rgba(168,85,247,.2) !important; border-radius: 14px !important;
  background: rgba(15,18,38,.85) !important;
}
[data-testid="stExpander"] summary, [data-testid="stExpander"] summary p {
  color: #f1f5f9 !important; font-weight: 700 !important;
}
[data-testid="stExpander"] svg { color: #c084fc !important; }

/* === data editor и таблицы === */
[data-testid="stDataFrame"], .stDataFrame { border-radius: 12px; overflow: hidden; }
[data-testid="stTable"] { color: #f1f5f9 !important; }
[data-testid="stTable"] td, [data-testid="stTable"] th { color: #f1f5f9 !important; }

/* === markdown === */
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li,
[data-testid="stMarkdownContainer"] strong { color: #f1f5f9; }
[data-testid="stMarkdownContainer"] small, .small { color: #a8aecf; }

/* === alerts === */
[data-testid="stAlert"] {
  background: rgba(168,85,247,.12) !important; border: 1px solid rgba(168,85,247,.3) !important;
  border-radius: 12px !important; color: #f1f5f9 !important;
}
[data-testid="stAlert"] p { color: #f1f5f9 !important; }
.stInfo, .stSuccess, .stWarning, .stError { color: #f1f5f9 !important; }

/* === чипы и карточки === */
.chip {
  display:inline-flex; align-items:center; gap:6px; padding:5px 12px;
  margin:2px 3px 2px 0; border-radius:999px; font-size:.72rem;
  font-weight:600; border:1px solid; color: #f1f5f9 !important;
}
.chip-num  { color:#d1fae5 !important; border-color:rgba(16,185,129,.5);  background:rgba(16,185,129,.18); }
.chip-date { color:#cffafe !important; border-color:rgba(6,182,212,.5);   background:rgba(6,182,212,.18); }
.chip-str  { color:#fef3c7 !important; border-color:rgba(251,146,60,.5);  background:rgba(251,146,60,.18); }
.card {
  background: linear-gradient(180deg, rgba(20,24,55,.95), rgba(12,15,38,.95)) !important;
  border:1px solid rgba(168,85,247,.2) !important;
  border-radius:16px !important; padding:16px 18px !important;
  box-shadow: 0 12px 30px -14px rgba(0,0,0,.85);
  color: #f1f5f9 !important;
}
.card b, .card strong { color: #ffffff !important; }
.hint { font-size:.78rem; color: #b8bdd9 !important; }
.section-title {
  font-family: 'Unbounded', sans-serif; font-weight: 800; font-size: 1.15rem;
  color: #ffffff !important; margin: 4px 0 14px; letter-spacing: -.01em;
  display: flex; align-items: center; gap: 10px;
}
.section-title::before {
  content: ''; display: block; width: 4px; height: 18px;
  background: linear-gradient(180deg,#a855f7,#22d3ee); border-radius: 4px;
}
.brand {
  font-family:'Unbounded',sans-serif; font-weight:900; font-size:1.1rem;
  background:linear-gradient(135deg,#c084fc,#22d3ee);
  -webkit-background-clip:text; -webkit-text-fill-color:transparent;
}
.kpi-label { color: #a8aecf !important; font-size: .72rem; font-weight: 700;
  text-transform: uppercase; letter-spacing: .09em; }
.muted { color: #b8bdd9 !important; }
.dl-ok { color: #6ee7b7 !important; }
.dl-bad { color: #fda4af !important; }
#MainMenu, footer { visibility: hidden; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

NUM, DATE, TXT = "число", "дата", "текст"
TYPE_CHIP = {NUM: "chip-num", DATE: "chip-date", TXT: "chip-str"}
TYPE_ICON = {NUM: "#", DATE: "D", TXT: "T"}

AGGS = {
    "Сумма": "sum",
    "Среднее": "mean",
    "Количество": "count",
    "Уникальных": "nunique",
    "Минимум": "min",
    "Максимум": "max",
    "Медиана": "median",
}
NUMERIC_AGGS = {"sum", "mean", "min", "max", "median"}

CHARTS = [
    "Столбцы", "Столбцы горизонтальные", "Столбцы с накоплением", "Сгруппированные столбцы",
    "Линия", "Область", "Круговая", "Кольцевая", "Treemap", "Sunburst",
    "Радар", "Полярная", "Точечная", "Пузырьковая", "Воронка", "Водопад",
    "Тепловая карта", "Ящик с усами", "Скрипка", "Гистограмма",
    "Индикатор KPI", "Сводная таблица",
]
RAW_CHARTS = {"Ящик с усами", "Скрипка", "Гистограмма"}

# NBSP / replacement chars via chr() to keep source free of unicode escapes
NBSP = chr(0xA0)
REPLACEMENT_CHAR = chr(0xFFFD)


# =============================================================================
# UTILS
# =============================================================================

def to_num(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return s
    cleaned = (
        s.astype(str)
        .str.replace(" ", "", regex=False)
        .str.replace(NBSP, "", regex=False)
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^0-9.\-eE]", "", regex=True)
        .replace({"": None, "-": None, ".": None})
    )
    return pd.to_numeric(cleaned, errors="coerce")


def to_date(s: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    return pd.to_datetime(s, errors="coerce", dayfirst=True)


def detect_type(s: pd.Series) -> str:
    sample = s.dropna()
    sample = sample[sample.astype(str).str.strip() != ""]
    if sample.empty:
        return TXT
    sample = sample.head(300)

    if pd.api.types.is_numeric_dtype(sample):
        return NUM
    if pd.api.types.is_datetime64_any_dtype(sample):
        return DATE

    as_str = sample.astype(str)
    date_pat = r"\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}"
    looks_date = as_str.str.contains(date_pat, regex=True, na=False)
    if looks_date.mean() >= 0.55 and to_date(sample).notna().mean() >= 0.55:
        return DATE

    # numeric-like strings without unicode escapes in the pattern
    money_chars = "0123456789 .," + NBSP + "+-%₽$€"
    def is_num_like(x: str) -> bool:
        x = str(x).strip()
        if not x or len(x) > 20:
            return False
        return all(ch in money_chars for ch in x)

    ratio = as_str.map(is_num_like).mean()
    if ratio >= 0.55 and to_num(sample).notna().mean() >= 0.55:
        return NUM
    return TXT


def detect_schema(df: pd.DataFrame) -> dict:
    return {c: detect_type(df[c]) for c in df.columns}


def coerce_types(df: pd.DataFrame, types: dict) -> pd.DataFrame:
    out = df.copy()
    for c, t in types.items():
        if c not in out.columns:
            continue
        if t == NUM:
            out[c] = to_num(out[c])
        elif t == DATE:
            out[c] = to_date(out[c])
        else:
            out[c] = out[c].astype(str).replace({"nan": "", "NaT": "", "None": ""})
    return out


def guess_primary_measure(types: dict):
    nums = [c for c, t in types.items() if t == NUM]
    if not nums:
        return None
    keys = ["выруч", "revenue", "сумм", "total", "прибыл", "profit", "amount", "продаж", "цена", "price"]
    for kw in keys:
        for c in nums:
            if kw in c.lower():
                return c
    return nums[0]


@st.cache_data(show_spinner=False)
def make_demo(n: int = 400) -> pd.DataFrame:
    rng = np.random.default_rng(7)
    cities = ["Москва", "Санкт-Петербург", "Казань", "Екатеринбург", "Новосибирск", "Краснодар", "Алматы", "Минск"]
    regions = {
        "Москва": "Центр", "Санкт-Петербург": "Северо-Запад", "Казань": "Поволжье",
        "Екатеринбург": "Урал", "Новосибирск": "Сибирь", "Краснодар": "Юг",
        "Алматы": "Казахстан", "Минск": "Беларусь",
    }
    clients = [
        "ООО Вектор", "АО ТехноПром", "ИП Иванов", "Альфа Логистик", "Урал Маш",
        "ТОО Логистик", "СтройГрупп", "ИП Петрова", "ЮгТрейд", "Дельта Систем",
    ]
    cats = ["Оборудование", "Софт", "Сервис", "Комплектующие", "Консалтинг", "Логистика"]
    managers = ["Смирнов А.", "Ковалева Е.", "Соколов Д.", "Морозова О.", "Волков М."]
    channels = ["Прямые", "Дистрибьютор", "Маркетплейс", "Тендер"]

    city = rng.choice(cities, n)
    revenue = rng.integers(25000, 520000, n)
    margin = rng.uniform(0.16, 0.46, n)
    days = rng.integers(0, 730, n)

    return pd.DataFrame({
        "Дата": pd.to_datetime("2023-01-01") + pd.to_timedelta(days, unit="D"),
        "Клиент": rng.choice(clients, n),
        "Город": city,
        "Регион": [regions[c] for c in city],
        "Менеджер": rng.choice(managers, n),
        "Категория": rng.choice(cats, n),
        "Канал продаж": rng.choice(channels, n),
        "Выручка": revenue,
        "Прибыль": (revenue * margin).round(0).astype(int),
        "Количество": rng.integers(1, 60, n),
        "Скидка %": rng.integers(0, 25, n),
        "Статус": rng.choice(["Оплачено", "В работе", "Просрочен"], n, p=[0.78, 0.16, 0.06]),
    })


def to_excel_bytes(df: pd.DataFrame, sheet: str = "Данные") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet[:31])
    return buf.getvalue()


def fmt_num(v) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return str(v)
    if abs(v) >= 1_000_000_000:
        return f"{v / 1_000_000_000:,.2f} млрд".replace(",", " ")
    if abs(v) >= 1_000_000:
        return f"{v / 1_000_000:,.2f} млн".replace(",", " ")
    return f"{v:,.0f}".replace(",", " ")


def nfmt(n: int) -> str:
    return f"{n:,}".replace(",", " ")


# =============================================================================
# STATE
# =============================================================================

def init_state():
    if "df" not in st.session_state:
        demo = make_demo()
        st.session_state.df = demo
        st.session_state.types = detect_schema(demo)
        st.session_state.source = "Демо-набор (12 столбцов)"
    st.session_state.setdefault("reports", [])
    st.session_state.setdefault("pending", None)
    st.session_state.setdefault("pending_name", "")


init_state()
DF: pd.DataFrame = st.session_state.df
TYPES: dict = st.session_state.types

cols_num = [c for c in DF.columns if TYPES.get(c) == NUM]
cols_date = [c for c in DF.columns if TYPES.get(c) == DATE]
cols_txt = [c for c in DF.columns if TYPES.get(c) == TXT]
cols_dim = cols_txt + cols_date
PRIMARY = guess_primary_measure(TYPES)


# =============================================================================
# SMART FILTERS
# =============================================================================

def render_smart_filters(df: pd.DataFrame, types: dict):
    mask = pd.Series(True, index=df.index)
    active = []

    search = st.sidebar.text_input(
        "Поиск по всем столбцам",
        key="flt__search",
        placeholder="ищем во всех полях...",
    )
    if search:
        joined = df.astype(str).agg(" ".join, axis=1).str.lower()
        mask &= joined.str.contains(re.escape(search.lower()), na=False)
        active.append(f'поиск: "{search}"')

    st.sidebar.caption("Фильтры сформированы по вашим столбцам")

    for col in df.columns:
        t = types.get(col, TXT)
        s = df[col]
        key = f"flt_{col}"

        if t == NUM:
            vals = to_num(s).dropna()
            if vals.empty:
                continue
            lo, hi = float(vals.min()), float(vals.max())
            if lo >= hi:
                continue
            step = max((hi - lo) / 100.0, 0.01)
            sel = st.sidebar.slider(
                f"{TYPE_ICON[NUM]} {col}", lo, hi, (lo, hi), step=step, key=key
            )
            if sel != (lo, hi):
                mask &= to_num(s).between(sel[0], sel[1])
                active.append(f"{col}: {fmt_num(sel[0])}–{fmt_num(sel[1])}")

        elif t == DATE:
            d = to_date(s).dropna()
            if d.empty:
                continue
            mn, mx = d.min().date(), d.max().date()
            if mn == mx:
                continue
            sel = st.sidebar.date_input(
                f"{TYPE_ICON[DATE]} {col}", (mn, mx), min_value=mn, max_value=mx, key=key
            )
            if isinstance(sel, (tuple, list)) and len(sel) == 2 and tuple(sel) != (mn, mx):
                dd = to_date(s).dt.date
                mask &= dd.between(sel[0], sel[1])
                active.append(f"{col}: {sel[0]} -> {sel[1]}")

        else:
            uniq = s.dropna().astype(str)
            uniq = uniq[uniq.str.strip() != ""].unique()
            if len(uniq) <= 1:
                continue
            if len(uniq) <= 60:
                sel = st.sidebar.multiselect(
                    f"{TYPE_ICON[TXT]} {col}",
                    sorted(uniq.tolist()),
                    key=key,
                    placeholder="все значения",
                )
                if sel:
                    mask &= s.astype(str).isin(sel)
                    active.append(f"{col}: {len(sel)} знач.")
            else:
                q = st.sidebar.text_input(
                    f"{TYPE_ICON[TXT]} {col} содержит",
                    key=key,
                    placeholder=f"{len(uniq)} уникальных",
                )
                if q:
                    mask &= s.astype(str).str.lower().str.contains(re.escape(q.lower()), na=False)
                    active.append(f'{col} ~ "{q}"')

    return df[mask], active


# =============================================================================
# PIVOT ENGINE
# =============================================================================

def dim_series(df: pd.DataFrame, col: str, gran: str = "Месяц") -> pd.Series:
    if TYPES.get(col) == DATE:
        d = to_date(df[col])
        if gran == "День":
            return d.dt.strftime("%Y-%m-%d").fillna("н/д")
        if gran == "Неделя":
            return d.dt.strftime("%Y-W%V").fillna("н/д")
        if gran == "Месяц":
            return d.dt.strftime("%Y-%m").fillna("н/д")
        if gran == "Квартал":
            y = d.dt.year.astype("Int64").astype(str)
            q = d.dt.quarter.astype("Int64").astype(str)
            return (y + "-Q" + q).fillna("н/д")
        return d.dt.strftime("%Y").fillna("н/д")
    return df[col].astype(str).replace({"": "н/д", "nan": "н/д"})


def build_pivot(df, row, col, measure, agg_label, gran="Месяц") -> pd.DataFrame:
    if df.empty or not row or not measure:
        return pd.DataFrame()
    agg = AGGS[agg_label]
    tmp = pd.DataFrame({"__row": dim_series(df, row, gran)})
    values = to_num(df[measure]) if agg in NUMERIC_AGGS else df[measure]
    tmp["__val"] = values.values

    if col and col != "— нет —":
        tmp["__col"] = dim_series(df, col, gran).values
        out = tmp.groupby(["__row", "__col"], dropna=False)["__val"].agg(agg).unstack(fill_value=0)
        out.columns = [str(c) for c in out.columns]
    else:
        out = tmp.groupby("__row", dropna=False)["__val"].agg(agg).to_frame(f"{agg_label} · {measure}")

    out.index.name = row
    return out.sort_index()


def shape_pivot(pt: pd.DataFrame, top_n=None, sort_desc=False) -> pd.DataFrame:
    if pt.empty:
        return pt
    out = pt.copy()
    if sort_desc:
        out = out.loc[out.sum(axis=1).sort_values(ascending=False).index]
    if top_n:
        out = out.head(int(top_n))
    return out


# =============================================================================
# ANALYTICS HELPERS
# =============================================================================

def numeric_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Возвращает только числовые колонки с приведённым к float типом."""
    out = pd.DataFrame()
    for c in df.columns:
        if pd.api.types.is_numeric_dtype(df[c]):
            out[c] = df[c]
        else:
            out[c] = to_num(df[c])
    return out


def profile_dataframe(df: pd.DataFrame, types: dict) -> pd.DataFrame:
    """Детальный профиль по каждой колонке (как в pandas-profiling, но встроенный)."""
    rows = []
    for c in df.columns:
        s = df[c]
        non_null = int(s.notna().sum())
        nulls = int(s.isna().sum())
        uniq = int(s.nunique(dropna=True))
        rows.append({
            "Столбец": c,
            "Тип": types.get(c, TXT),
            "Заполнено": non_null,
            "Пусто": nulls,
            "Пусто %": round(nulls / max(1, len(df)) * 100, 1),
            "Уникальных": uniq,
            "Пример": (str(s.dropna().iloc[0])[:48] if s.notna().any() else ""),
        })
        if types.get(c) == NUM:
            v = to_num(s).dropna()
            if not v.empty:
                rows[-1].update({
                    "Мин": float(v.min()),
                    "Макс": float(v.max()),
                    "Среднее": round(float(v.mean()), 2),
                    "Медиана": round(float(v.median()), 2),
                    "Ст. откл.": round(float(v.std()) if len(v) > 1 else 0.0, 2),
                    "Сумма": round(float(v.sum()), 2),
                })
        elif types.get(c) == DATE:
            d = to_date(s).dropna()
            if not d.empty:
                rows[-1].update({"Мин": str(d.min().date()), "Макс": str(d.max().date()),
                                 "Уникальных дат": int(d.nunique())})
    return pd.DataFrame(rows)


def correlation_matrix(df: pd.DataFrame) -> pd.DataFrame:
    nums = numeric_frame(df)
    if nums.shape[1] < 2:
        return pd.DataFrame()
    return nums.corr(numeric_only=True).round(2)


def abc_analysis(df: pd.DataFrame, dim: str, measure: str) -> pd.DataFrame:
    """ABC-анализ: ранжируем по вкладу, относим к классам A/B/C (70/20/10)."""
    if dim not in df.columns or measure not in df.columns:
        return pd.DataFrame()
    tmp = pd.DataFrame({
        dim: df[dim].astype(str),
        "val": to_num(df[measure]).values,
    }).groupby(dim, dropna=False)["val"].sum().sort_values(ascending=False)
    if tmp.empty or tmp.sum() == 0:
        return pd.DataFrame()
    total = tmp.sum()
    tmp = tmp.to_frame("Сумма")
    tmp["Доля %"] = (tmp["Сумма"] / total * 100).round(2)
    tmp["Накопленная %"] = tmp["Доля %"].cumsum()
    def cls(x):
        if x <= 70: return "A"
        if x <= 90: return "B"
        return "C"
    tmp["Класс"] = tmp["Накопленная %"].apply(cls)
    return tmp.reset_index()


def xyz_analysis(df: pd.DataFrame, dim: str, measure: str) -> pd.DataFrame:
    """XYZ: классификация по стабильности (коэффициент вариации)."""
    if dim not in df.columns or measure not in df.columns or df.empty:
        return pd.DataFrame()
    g = df.groupby(dim, dropna=False)[measure].apply(lambda v: to_num(v).dropna())
    if g.empty:
        return pd.DataFrame()
    out = g.reset_index()
    out.columns = [dim, "values"]
    stats = out["values"].apply(lambda v: pd.Series({
        "Периодов": len(v), "Среднее": v.mean() if len(v) else 0,
        "Ст. откл.": v.std() if len(v) > 1 else 0,
    }))
    stats["CV %"] = np.where(stats["Среднее"].abs() > 0,
                              stats["Ст. откл."] / stats["Среднее"].abs() * 100, 0)
    def cls(x):
        if x < 10: return "X"
        if x < 25: return "Y"
        return "Z"
    stats["Группа"] = stats["CV %"].apply(cls)
    return stats.round(2).reset_index().rename(columns={"index": dim})


def forecast_linear(series: pd.Series, periods: int = 3) -> pd.DataFrame:
    """Простой линейный прогноз по y = a*x + b."""
    s = pd.to_numeric(series, errors="coerce").dropna()
    if len(s) < 2:
        return pd.DataFrame()
    x = np.arange(len(s))
    a, b = np.polyfit(x, s.values, 1)
    future_x = np.arange(len(s), len(s) + periods)
    future_y = a * future_x + b
    return pd.DataFrame({
        "idx": np.concatenate([x, future_x]),
        "value": np.concatenate([s.values, future_y]),
        "type": ["Факт"] * len(s) + ["Прогноз"] * periods,
    })


def detect_outliers(s: pd.Series) -> int:
    v = to_num(s).dropna()
    if len(v) < 4:
        return 0
    q1, q3 = v.quantile(0.25), v.quantile(0.75)
    iqr = q3 - q1
    return int(((v < q1 - 1.5 * iqr) | (v > q3 + 1.5 * iqr)).sum())


def quick_stats(df: pd.DataFrame, types: dict) -> dict:
    return {
        "rows": int(len(df)),
        "cols": int(df.shape[1]),
        "num_cols": int(sum(1 for t in types.values() if t == NUM)),
        "date_cols": int(sum(1 for t in types.values() if t == DATE)),
        "text_cols": int(sum(1 for t in types.values() if t == TXT)),
        "duplicates": int(df.duplicated().sum()),
        "missing": int(df.isna().sum().sum()),
        "missing_pct": round(df.isna().sum().sum() / max(1, df.size) * 100, 2),
        "memory_kb": round(df.memory_usage(deep=True).sum() / 1024, 1),
    }


# =============================================================================
# CHART ENGINE
# =============================================================================

def style_fig(fig, height: int = 420):
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Plus Jakarta Sans", size=12, color="#cbd5e1"),
        margin=dict(l=10, r=10, t=36, b=10),
        height=height,
        colorway=PALETTE,
        legend=dict(orientation="h", y=-0.18, font=dict(size=10)),
        hoverlabel=dict(bgcolor="#141938", bordercolor="#a855f7", font_size=11),
    )
    fig.update_xaxes(gridcolor="rgba(255,255,255,.06)", zerolinecolor="rgba(255,255,255,.08)")
    fig.update_yaxes(gridcolor="rgba(255,255,255,.06)", zerolinecolor="rgba(255,255,255,.08)")
    return fig


def render_chart(kind, pt, raw, row, measure, agg_label, height=420):
    if kind in RAW_CHARTS:
        if raw.empty or measure not in raw.columns or row not in raw.columns:
            st.info("Нет данных для построения.")
            return
        y = to_num(raw[measure])
        plot_df = pd.DataFrame({row: raw[row].astype(str), measure: y}).dropna()
        if plot_df.empty:
            st.info("Нет данных для построения.")
            return
        if kind == "Ящик с усами":
            fig = px.box(plot_df, x=row, y=measure, color=row, points="outliers")
        elif kind == "Скрипка":
            fig = px.violin(plot_df, x=row, y=measure, color=row, box=True, points=False)
        else:
            fig = px.histogram(plot_df, x=measure, color=row, nbins=40, barmode="overlay", opacity=0.75)
        fig.update_layout(showlegend=False)
        st.plotly_chart(style_fig(fig, height), use_container_width=True)
        return

    if pt is None or pt.empty:
        st.info("Нет данных под текущими фильтрами.")
        return

    idx_name = pt.index.name or "Категория"
    long = pt.reset_index().melt(id_vars=pt.index.name, var_name="Серия", value_name="Значение")
    long = long.rename(columns={pt.index.name: idx_name})
    multi = pt.shape[1] > 1
    color = "Серия" if multi else None
    totals = pt.sum(axis=1)

    if kind == "Столбцы":
        fig = px.bar(long, x=idx_name, y="Значение", color=color, barmode="group")
    elif kind == "Столбцы горизонтальные":
        fig = px.bar(long, y=idx_name, x="Значение", color=color, orientation="h", barmode="group")
        fig.update_layout(yaxis=dict(categoryorder="total ascending"))
    elif kind == "Столбцы с накоплением":
        fig = px.bar(long, x=idx_name, y="Значение", color=color, barmode="stack")
    elif kind == "Сгруппированные столбцы":
        fig = px.bar(long, x=idx_name, y="Значение", color=color, barmode="group", text_auto=".2s")
    elif kind == "Линия":
        fig = px.line(long, x=idx_name, y="Значение", color=color, markers=True)
    elif kind == "Область":
        fig = px.area(long, x=idx_name, y="Значение", color=color)
    elif kind in ("Круговая", "Кольцевая"):
        d = totals.reset_index()
        d.columns = [idx_name, "Значение"]
        fig = px.pie(d, names=idx_name, values="Значение", hole=0.58 if kind == "Кольцевая" else 0)
        fig.update_traces(textposition="inside", textinfo="percent+label")
    elif kind == "Treemap":
        d = totals.reset_index()
        d.columns = [idx_name, "Значение"]
        fig = px.treemap(d, path=[idx_name], values="Значение", color="Значение", color_continuous_scale="Purples")
    elif kind == "Sunburst":
        if multi:
            fig = px.sunburst(
                long, path=[idx_name, "Серия"], values="Значение",
                color="Значение", color_continuous_scale="Purples",
            )
        else:
            d = totals.reset_index()
            d.columns = [idx_name, "Значение"]
            fig = px.sunburst(d, path=[idx_name], values="Значение")
    elif kind == "Радар":
        fig = go.Figure()
        for c in pt.columns:
            fig.add_trace(go.Scatterpolar(
                r=pt[c].values,
                theta=[str(i) for i in pt.index],
                fill="toself",
                name=str(c),
                opacity=0.65,
            ))
        fig.update_layout(polar=dict(bgcolor="rgba(255,255,255,.02)"))
    elif kind == "Полярная":
        d = totals.reset_index()
        d.columns = [idx_name, "Значение"]
        fig = px.bar_polar(d, r="Значение", theta=idx_name, color="Значение", color_continuous_scale="Plasma")
    elif kind == "Точечная":
        d = totals.reset_index()
        d.columns = [idx_name, "Значение"]
        fig = px.scatter(
            d, x=idx_name, y="Значение", color="Значение",
            size="Значение", color_continuous_scale="Plasma",
        )
    elif kind == "Пузырьковая":
        d = totals.reset_index()
        d.columns = [idx_name, "Значение"]
        d["Ранг"] = d["Значение"].rank()
        fig = px.scatter(
            d, x="Ранг", y="Значение", size="Значение", color=idx_name,
            hover_name=idx_name, size_max=55,
        )
    elif kind == "Воронка":
        d = totals.sort_values(ascending=False).reset_index()
        d.columns = [idx_name, "Значение"]
        fig = px.funnel(d, y=idx_name, x="Значение")
    elif kind == "Водопад":
        fig = go.Figure(go.Waterfall(
            x=[str(i) for i in pt.index],
            y=totals.values,
            connector=dict(line=dict(color="rgba(168,85,247,.4)")),
            increasing=dict(marker_color="#10b981"),
            decreasing=dict(marker_color="#f43f5e"),
            totals=dict(marker_color="#a855f7"),
        ))
    elif kind == "Тепловая карта":
        fig = px.imshow(
            pt.values,
            x=[str(c) for c in pt.columns],
            y=[str(i) for i in pt.index],
            color_continuous_scale="Purples",
            aspect="auto",
            text_auto=".2s",
        )
    elif kind == "Индикатор KPI":
        total = float(totals.sum())
        fig = go.Figure(go.Indicator(
            mode="number+gauge",
            value=total,
            number={"font": {"size": 44, "color": "#fff"}},
            gauge={
                "axis": {"range": [0, max(total, 1) * 1.15]},
                "bar": {"color": "#a855f7"},
                "bgcolor": "rgba(255,255,255,.05)",
            },
            title={"text": f"{agg_label} · {measure}", "font": {"size": 13, "color": "#94a3b8"}},
        ))
    else:
        show = pt.copy()
        show["ИТОГО"] = show.sum(axis=1)
        st.dataframe(
            show.style.format(lambda v: fmt_num(v)),
            use_container_width=True,
            height=height,
        )
        return

    st.plotly_chart(style_fig(fig, height), use_container_width=True)


# =============================================================================
# SIDEBAR
# =============================================================================

with st.sidebar:
    st.markdown('<div class="brand">ADAPTIVE BI</div>', unsafe_allow_html=True)
    st.caption("Любые таблицы · Любые отчёты")

    st.markdown(
        f'<div class="card" style="padding:12px 14px">'
        f'<div style="font-size:.72rem;color:#94a3b8">ИСТОЧНИК</div>'
        f'<div style="font-size:.8rem;font-weight:700;color:#fff">{st.session_state.source}</div>'
        f'<div style="font-size:.7rem;color:#94a3b8;margin-top:4px">'
        f'Строк: <b style="color:#c084fc">{nfmt(len(DF))}</b> · '
        f'Столбцов: <b style="color:#22d3ee">{DF.shape[1]}</b></div></div>',
        unsafe_allow_html=True,
    )

    chips = "".join(
        f'<span class="chip {TYPE_CHIP[TYPES.get(c, TXT)]}">{TYPE_ICON[TYPES.get(c, TXT)]} {c}</span>'
        for c in list(DF.columns)[:10]
    )
    st.markdown(f'<div style="margin:10px 0">{chips}</div>', unsafe_allow_html=True)

    st.divider()
    if st.button("Сбросить все фильтры", use_container_width=True):
        for k in [k for k in list(st.session_state.keys()) if str(k).startswith("flt_")]:
            del st.session_state[k]
        st.rerun()

    FDF, ACTIVE = render_smart_filters(DF, TYPES)


# =============================================================================
# HEADER
# =============================================================================

head_l, head_r = st.columns([3, 1])
with head_l:
    st.markdown("# Аналитика любых данных")
    extra = f' · Активно фильтров: <b style="color:#22d3ee">{len(ACTIVE)}</b>' if ACTIVE else ""
    st.markdown(
        f'<div class="hint">Отфильтровано <b style="color:#c084fc">{nfmt(len(FDF))}</b> '
        f'из <b>{nfmt(len(DF))}</b> строк{extra}</div>',
        unsafe_allow_html=True,
    )
with head_r:
    st.download_button(
        "Выгрузить в Excel",
        to_excel_bytes(FDF),
        file_name=f"export_{date.today()}.xlsx",
        use_container_width=True,
    )

if ACTIVE:
    st.markdown(
        "".join(f'<span class="chip chip-num">{a}</span>' for a in ACTIVE),
        unsafe_allow_html=True,
    )

tab_over, tab_profile, tab_correl, tab_abc, tab_build, tab_dash, tab_pivot, tab_data, tab_import = st.tabs([
    "Обзор", "Профиль данных", "Корреляции",
    "ABC / XYZ", "Конструктор", "Дашборды",
    "Сводные срезы", "Данные", "Импорт",
])


# =============================================================================
# TAB 1 — OVERVIEW (обогащённый)
# =============================================================================

with tab_over:
    st.markdown('<div class="section-title">Сводка по набору данных</div>', unsafe_allow_html=True)

    stats = quick_stats(FDF, TYPES)
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Строк", nfmt(stats["rows"]))
    c2.metric("Столбцов", str(stats["cols"]))
    c3.metric("Числовых", str(stats["num_cols"]))
    c4.metric("Дат", str(stats["date_cols"]))
    c5.metric("Дубликаты", nfmt(stats["duplicates"]),
              delta=None if stats["duplicates"] == 0 else "проверьте")
    c6.metric("Пропуски", f"{stats['missing_pct']}%",
              delta=f"{nfmt(stats['missing'])} ячеек")

    st.markdown("---")
    if not cols_num:
    if not cols_num:
        st.warning("В данных нет числовых столбцов — загрузите таблицу с числами для KPI.")
    else:
        st.subheader("Автоматические показатели")
        kpi_cols = st.columns(min(4, len(cols_num)))
        for i, c in enumerate(cols_num[:4]):
            v = to_num(FDF[c])
            total = float(v.sum()) if len(v) else 0.0
            avg = float(v.mean()) if len(v) else 0.0
            base = float(to_num(DF[c]).sum())
            delta = (total - base) / base * 100 if base else 0.0
            with kpi_cols[i]:
                st.metric(
                    c,
                    fmt_num(total),
                    delta=f"{delta:+.1f}% от всего" if abs(delta) > 0.01 else "весь объём",
                )
                st.caption(f"Среднее: **{fmt_num(avg)}** · строк: {nfmt(len(v))}")

        st.divider()
        c1, c2 = st.columns([2, 1])

        with c1:
            st.markdown("##### Динамика")
            o1, o2, o3 = st.columns(3)
            dim_options = cols_dim or list(DF.columns)
            default_dim = 0
            if cols_date:
                try:
                    default_dim = dim_options.index(cols_date[0])
                except ValueError:
                    default_dim = 0
            dim = o1.selectbox("Разрез", dim_options, key="ov_dim", index=default_dim)
            meas_idx = cols_num.index(PRIMARY) if PRIMARY in cols_num else 0
            meas = o2.selectbox("Показатель", cols_num, key="ov_meas_pick", index=meas_idx)
            gran = o3.selectbox(
                "Гранулярность",
                ["День", "Неделя", "Месяц", "Квартал", "Год"],
                index=2,
                key="ov_gran",
            )
            pt = build_pivot(FDF, dim, None, meas, "Сумма", gran)
            kind = "Область" if TYPES.get(dim) == DATE else "Столбцы"
            render_chart(kind, pt, FDF, dim, meas, "Сумма", height=340)

        with c2:
            st.markdown("##### Структура")
            pdim_options = cols_txt or cols_dim or list(DF.columns)
            pdim = st.selectbox("Измерение", pdim_options, key="ov_pdim_pick")
            ppt = shape_pivot(build_pivot(FDF, pdim, None, meas, "Сумма"), 10, True)
            render_chart("Кольцевая", ppt, FDF, pdim, meas, "Сумма", height=300)

        st.divider()
        st.markdown("##### Авто-инсайты по вашим данным")
        i1, i2, i3 = st.columns(3)
        if not ppt.empty:
            top_name = str(ppt.index[0])
            top_val = float(ppt.iloc[0].sum())
            share = top_val / float(ppt.values.sum()) * 100 if ppt.values.sum() else 0
            i1.markdown(
                f'<div class="card"><b style="color:#c084fc">Лидер по «{pdim}»</b><br>'
                f"{top_name} — {fmt_num(top_val)} ({share:.1f}% от топ-10)</div>",
                unsafe_allow_html=True,
            )
        if not pt.empty and len(pt) >= 2:
            last = float(pt.iloc[-1].sum())
            prev = float(pt.iloc[-2].sum())
            growth = (last - prev) / prev * 100 if prev else 0
            arrow = "▲" if growth >= 0 else "▼"
            color = "#6ee7b7" if growth >= 0 else "#fda4af"
            i2.markdown(
                f'<div class="card"><b style="color:#22d3ee">Последний период</b><br>'
                f"{pt.index[-1]}: {fmt_num(last)} "
                f'<span style="color:{color}">{arrow} {abs(growth):.1f}%</span></div>',
                unsafe_allow_html=True,
            )
        i3.markdown(
            f'<div class="card"><b style="color:#6ee7b7">Структура таблицы</b><br>'
            f"{len(cols_num)} числовых · {len(cols_date)} дат · {len(cols_txt)} текстовых</div>",
            unsafe_allow_html=True,
        )


# =============================================================================
# TAB 2 — DATA PROFILE
# =============================================================================

with tab_profile:
    st.markdown('<div class="section-title">Профиль данных по каждому столбцу</div>', unsafe_allow_html=True)
    st.caption("Автоматический отчёт о качестве: типы, пропуски, статистика, уникальные значения.")
    profile_df = profile_dataframe(FDF, TYPES)
    st.dataframe(profile_df.style.format(na_rep="—"), use_container_width=True, height=420)

    if cols_txt:
        st.markdown("##### Топ значений по текстовым колонкам")
        sel = st.selectbox("Столбец", cols_txt, key="prof_str")
        vc = FDF[sel].astype(str).replace("", np.nan).dropna().value_counts().head(15)
        d = vc.reset_index(); d.columns = [sel, "Кол-во"]
        fig = px.bar(d, x="Кол-во", y=sel, orientation="h", color="Кол-во",
                     color_continuous_scale="Purples")
        st.plotly_chart(style_fig(fig, 340), use_container_width=True)


# =============================================================================
# TAB 3 — CORRELATIONS + FORECAST
# =============================================================================

with tab_correl:
    st.markdown('<div class="section-title">Корреляции, выбросы и прогноз</div>', unsafe_allow_html=True)

    if not cols_num:
        st.info("Нет числовых колонок.")
    else:
        st.markdown("##### Корреляционная матрица")
        corr = correlation_matrix(FDF)
        if corr.empty or corr.shape[0] < 2:
            st.info("Нужно минимум 2 числовые колонки.")
        else:
            fig = px.imshow(corr, text_auto=".2f", color_continuous_scale="RdBu",
                            zmin=-1, zmax=1, aspect="auto")
            st.plotly_chart(style_fig(fig, 480), use_container_width=True)

            pairs = []
            cl = list(corr.columns)
            for i, a in enumerate(cl):
                for b in cl[i + 1:]:
                    pairs.append((a, b, corr.loc[a, b]))
            pairs.sort(key=lambda x: abs(x[2]), reverse=True)
            if pairs:
                pp = pd.DataFrame(pairs[:8], columns=["Столбец A", "Столбец B", "Корреляция"])
                st.dataframe(pp.style.background_gradient(cmap="RdBu", subset=["Корреляция"],
                                                          vmin=-1, vmax=1),
                             use_container_width=True, height=300)

        st.markdown("---")
        st.markdown("##### Линейный прогноз")
        if cols_date:
            f1, f2, f3 = st.columns(3)
            date_col = f1.selectbox("Столбец дат", cols_date, key="fc_date")
            meas_col = f2.selectbox("Что прогнозируем", cols_num, key="fc_meas_x",
                                    index=cols_num.index(PRIMARY) if PRIMARY in cols_num else 0)
            horizon = f3.slider("Горизонт (периодов)", 1, 12, 3, key="fc_h")
            gran = st.selectbox("Гранулярность", ["День", "Неделя", "Месяц", "Квартал", "Год"],
                                index=2, key="fc_gran")
            series = (build_pivot(FDF, date_col, None, meas_col, "Сумма", gran)
                      .iloc[:, 0].reset_index(drop=True))
            forecast = forecast_linear(series, horizon)
            if not forecast.empty:
                fig = go.Figure()
                fact = forecast[forecast["type"] == "Факт"]
                pred = forecast[forecast["type"] == "Прогноз"]
                fig.add_trace(go.Scatter(x=fact["idx"], y=fact["value"], name="Факт",
                                         mode="lines+markers", line=dict(color="#22d3ee", width=3)))
                fig.add_trace(go.Scatter(x=pred["idx"], y=pred["value"], name="Прогноз",
                                         mode="lines+markers",
                                         line=dict(color="#a855f7", width=3, dash="dash")))
                fig.update_layout(xaxis_title="Период", yaxis_title=meas_col)
                st.plotly_chart(style_fig(fig, 360), use_container_width=True)

                last_fact = float(fact["value"].iloc[-1])
                last_pred = float(pred["value"].iloc[-1])
                delta = (last_pred - last_fact) / last_fact * 100 if last_fact else 0
                arrow_cls = "dl-ok" if delta >= 0 else "dl-bad"
                st.markdown(
                    f'<div class="card"><b>Прогноз на {horizon} период(а):</b> '
                    f'факт {fmt_num(last_fact)} → прогноз {fmt_num(last_pred)} '
                    f'<span class="{arrow_cls}">({delta:+.1f}%)</span></div>',
                    unsafe_allow_html=True)
        else:
            st.info("Прогноз требует колонку с датами.")


# =============================================================================
# TAB 4 — ABC / XYZ
# =============================================================================

with tab_abc:
    st.markdown('<div class="section-title">ABC и XYZ анализ</div>', unsafe_allow_html=True)
    if not cols_dim or not cols_num:
        st.info("Нужны измерения и числовые колонки.")
    else:
        a1, a2 = st.columns(2)
        dim = a1.selectbox("Измерение (ABC)", cols_dim, key="abc_dim")
            meas_abc = a2.selectbox("Мера (ABC)", cols_num, key="abc_meas_x",
                                    index=cols_num.index(PRIMARY) if PRIMARY in cols_num else 0)
        abc = abc_analysis(FDF, dim, meas_abc)
        if abc.empty:
            st.info("Нет данных.")
        else:
            st.markdown("##### ABC — вклад в общий результат (A=70%, B=20%, C=10%)")
            st.dataframe(
                abc.style.background_gradient(cmap="Purples", subset=["Сумма"])
                .format({"Сумма": lambda v: fmt_num(v),
                         "Доля %": "{:.2f}%".format,
                         "Накопленная %": "{:.2f}%".format}),
                use_container_width=True, height=380,
            )
            fig = px.bar(abc, x=dim, y="Сумма", color="Класс",
                         color_discrete_map={"A": "#10b981", "B": "#f59e0b", "C": "#f43f5e"})
            st.plotly_chart(style_fig(fig, 340), use_container_width=True)

        if cols_date:
            st.markdown("---")
            st.markdown("##### XYZ — стабильность (CV %)")
            d1, d2 = st.columns(2)
            date_col = d1.selectbox("Период", cols_date, key="xyz_date")
            meas_xyz = d2.selectbox("Мера (XYZ)", cols_num, key="xyz_meas_x",
                                    index=cols_num.index(PRIMARY) if PRIMARY in cols_num else 0)
            gran = st.selectbox("Гранулярность", ["День", "Неделя", "Месяц", "Квартал", "Год"],
                                index=2, key="xyz_gran")
            grouped = FDF.copy()
            grouped["__p"] = dim_series(grouped, date_col, gran)
            grouped["__v"] = to_num(grouped[meas_xyz])
            xyz = (grouped.groupby("__p", dropna=False)["__v"]
                   .agg(["mean", "std", "count"]).reset_index().dropna())
            if not xyz.empty:
                xyz["CV %"] = np.where(xyz["mean"].abs() > 0,
                                       xyz["std"] / xyz["mean"].abs() * 100, 0)
                xyz["Группа"] = xyz["CV %"].apply(
                    lambda x: "X" if x < 10 else ("Y" if x < 25 else "Z"))
                fig = px.scatter(xyz, x="__p", y="mean", color="Группа",
                                 size=xyz["std"].abs() + 1,
                                 color_discrete_map={"X": "#10b981", "Y": "#f59e0b", "Z": "#f43f5e"})
                st.plotly_chart(style_fig(fig, 340), use_container_width=True)
                st.dataframe(
                    xyz.rename(columns={"__p": date_col, "mean": "Среднее",
                                        "std": "Ст. откл.", "count": "Кол-во"})
                    [[date_col, "Среднее", "Ст. откл.", "CV %", "Группа"]]
                    .style.background_gradient(cmap="RdYlGn_r", subset=["CV %"]),
                    use_container_width=True, height=320,
                )


# =============================================================================
# TAB 5 — BUILDER
# =============================================================================

with tab_build:
    st.subheader("Конструктор отчётов — как сводная таблица, но с любым графиком")
    st.caption("Выберите столбцы для группировки и показатель. Приложение построит любую визуализацию.")

    left, right = st.columns([1, 2])

    with left:
        title = st.text_input(
            "Название отчёта",
            key="b_title",
            placeholder="например: Выручка по городам",
        )
        row_options = cols_dim or list(DF.columns)
        row = st.selectbox("Группировать по (строки)", row_options, key="b_row")
        col2_options = ["— нет —"] + [c for c in cols_dim if c != row]
        col2 = st.selectbox("Разбить по (колонки)", col2_options, key="b_col")
        agg_label = st.selectbox("Агрегация", list(AGGS.keys()), key="b_agg")
        meas_pool = cols_num if AGGS[agg_label] in NUMERIC_AGGS else list(DF.columns)
        if not meas_pool:
            meas_pool = list(DF.columns)
        meas_idx = meas_pool.index(PRIMARY) if PRIMARY in meas_pool else 0
        meas = st.selectbox("Показатель (мера)", meas_pool, key="b_meas_pick", index=meas_idx)
        kind = st.selectbox("Тип визуализации", CHARTS, key="b_kind")

        gran = "Месяц"
        if TYPES.get(row) == DATE or TYPES.get(col2) == DATE:
            gran = st.selectbox(
                "Гранулярность дат",
                ["День", "Неделя", "Месяц", "Квартал", "Год"],
                index=2,
                key="b_gran",
            )
        cc1, cc2 = st.columns(2)
        sort_desc = cc1.checkbox("Сортировать", True, key="b_sort")
        use_top = cc2.checkbox("Только топ-15", False, key="b_top")

        if st.button("Сохранить на дашборд", type="primary", use_container_width=True):
            st.session_state.reports.append({
                "title": title or f"{agg_label} {meas} по {row}",
                "row": row,
                "col": col2,
                "meas": meas,
                "agg": agg_label,
                "kind": kind,
                "gran": gran,
                "sort": sort_desc,
                "top": use_top,
            })
            st.success("Отчёт добавлен во вкладку «Мои дашборды»")

    with right:
        pt = shape_pivot(
            build_pivot(FDF, row, col2, meas, agg_label, gran),
            15 if use_top else None,
            sort_desc,
        )
        st.caption(
            f"Категорий: **{len(pt)}** · Серий: **{pt.shape[1] if not pt.empty else 0}** · "
            f"{agg_label} по «{meas}»"
        )
        render_chart(kind, pt, FDF, row, meas, agg_label, height=430)
        if not pt.empty and kind != "Сводная таблица":
            with st.expander("Показать таблицу с числами"):
                show = pt.copy()
                show["ИТОГО"] = show.sum(axis=1)
                st.dataframe(show.style.format(lambda v: fmt_num(v)), use_container_width=True)

    st.divider()
    st.markdown("##### Идеи отчётов по вашим столбцам")
    ideas = []
    if cols_date and cols_num:
        ideas.append((f"Динамика «{cols_num[0]}»", cols_date[0], cols_num[0], "Область"))
    for d in cols_txt[:3]:
        m = PRIMARY or (cols_num[0] if cols_num else list(DF.columns)[0])
        ideas.append((f"Рейтинг: {d}", d, m, "Столбцы горизонтальные"))
    if len(cols_txt) >= 2 and cols_num:
        ideas.append((f"{cols_txt[0]} × {cols_txt[1]}", cols_txt[0], cols_num[0], "Тепловая карта"))
    if cols_txt and cols_num:
        ideas.append((f"Структура «{cols_txt[0]}»", cols_txt[0], cols_num[0], "Treemap"))

    if ideas:
        icols = st.columns(min(4, len(ideas)))
        for i, (t, d, m, k) in enumerate(ideas[:4]):
            with icols[i]:
                st.markdown(
                    f'<div class="card"><b>{t}</b><br><span class="hint">{k} · {m}</span></div>',
                    unsafe_allow_html=True,
                )
                if st.button("Построить", key=f"idea_{i}", use_container_width=True):
                    st.session_state.reports.append({
                        "title": t,
                        "row": d,
                        "col": "— нет —",
                        "meas": m,
                        "agg": "Сумма",
                        "kind": k,
                        "gran": "Месяц",
                        "sort": True,
                        "top": True,
                    })
                    st.success("Добавлено в «Мои дашборды»")


# =============================================================================
# TAB 3 — DASHBOARDS
# =============================================================================

with tab_dash:
    st.subheader("Мои дашборды")
    reports = st.session_state.reports
    if not reports:
        st.info(
            "Пока пусто. Соберите отчёт в «Конструкторе» и сохраните его сюда — "
            "все карточки реагируют на фильтры слева."
        )
    else:
        c1, c2 = st.columns([3, 1])
        c1.caption(f"Сохранено отчётов: **{len(reports)}** · все реагируют на текущие фильтры")
        if c2.button("Очистить всё", use_container_width=True):
            st.session_state.reports = []
            st.rerun()

        for start in range(0, len(reports), 2):
            row_cols = st.columns(2)
            for j, rep in enumerate(reports[start:start + 2]):
                idx = start + j
                with row_cols[j]:
                    with st.container(border=True):
                        h1, h2 = st.columns([5, 1])
                        h1.markdown(f"**{rep['title']}**")
                        h1.caption(
                            f"{rep['row']} × {rep['col']} · {rep['agg']} {rep['meas']} · {rep['kind']}"
                        )
                        if h2.button("X", key=f"del_{idx}"):
                            st.session_state.reports.pop(idx)
                            st.rerun()
                        p = shape_pivot(
                            build_pivot(
                                FDF, rep["row"], rep["col"], rep["meas"], rep["agg"], rep["gran"]
                            ),
                            15 if rep["top"] else None,
                            rep["sort"],
                        )
                        render_chart(
                            rep["kind"], p, FDF, rep["row"], rep["meas"], rep["agg"], height=300
                        )


# =============================================================================
# TAB 4 — PIVOT
# =============================================================================

with tab_pivot:
    st.subheader("Сводные срезы по любым полям")
    p1, p2, p3, p4, p5 = st.columns(5)
    prow_opts = cols_dim or list(DF.columns)
    prow = p1.selectbox("Строки", prow_opts, key="p_row")
    pcol = p2.selectbox(
        "Колонки",
        ["— нет —"] + [c for c in cols_dim if c != prow],
        key="p_col",
    )
    pagg = p3.selectbox("Агрегация", list(AGGS.keys()), key="p_agg")
    pmeas_pool = cols_num if AGGS[pagg] in NUMERIC_AGGS else list(DF.columns)
    if not pmeas_pool:
        pmeas_pool = list(DF.columns)
    pmeas = p4.selectbox("Мера", pmeas_pool, key="p_meas_pick")
    pgran = p5.selectbox(
        "Даты", ["День", "Неделя", "Месяц", "Квартал", "Год"], index=2, key="p_gran"
    )

    o1, o2, o3 = st.columns([1, 1, 2])
    transpose = o1.checkbox("Транспонировать", key="p_tr")
    heat = o2.checkbox("Тепловая заливка", True, key="p_heat")

    pv = build_pivot(FDF, prow, pcol, pmeas, pagg, pgran)
    if transpose and not pv.empty:
        pv = pv.T
        pv.index.name = pcol if pcol != "— нет —" else "Показатель"

    if pv.empty:
        st.info("Нет данных под текущими фильтрами.")
    else:
        table = pv.copy()
        table["ИТОГО"] = table.sum(axis=1)
        total_row = pd.DataFrame(table.sum(axis=0)).T
        total_row.index = ["ИТОГО"]
        table = pd.concat([table, total_row])

        styler = table.style.format(lambda v: fmt_num(v))
        if heat:
            data_cols = [c for c in table.columns if c != "ИТОГО"]
            try:
                styler = styler.background_gradient(
                    cmap="Purples",
                    axis=None,
                    subset=(table.index[:-1], data_cols),
                )
            except Exception:
                pass
        st.dataframe(styler, use_container_width=True, height=460)

        o3.download_button(
            "Скачать срез в Excel",
            to_excel_bytes(table.reset_index(), "Срез"),
            file_name=f"pivot_{date.today()}.xlsx",
            use_container_width=True,
        )


# =============================================================================
# TAB 5 — DATA
# =============================================================================

with tab_data:
    st.subheader("Данные и управление столбцами")

    with st.expander("Свои названия столбцов и типы", expanded=False):
        st.caption("Переименуйте столбцы — изменения применятся ко всем отчётам.")
        cfg = pd.DataFrame({
            "Столбец": list(DF.columns),
            "Новое название": list(DF.columns),
            "Тип": [TYPES.get(c, TXT) for c in DF.columns],
            "Оставить": [True] * DF.shape[1],
            "Пример": [
                str(DF[c].dropna().iloc[0])[:40] if DF[c].notna().any() else ""
                for c in DF.columns
            ],
            "Пусто %": [round(float(DF[c].isna().mean() * 100), 1) for c in DF.columns],
        })
        edited = st.data_editor(
            cfg,
            hide_index=True,
            use_container_width=True,
            key="col_editor",
            column_config={
                "Столбец": st.column_config.TextColumn(disabled=True),
                "Тип": st.column_config.SelectboxColumn(options=[NUM, DATE, TXT], required=True),
                "Пример": st.column_config.TextColumn(disabled=True),
                "Пусто %": st.column_config.NumberColumn(disabled=True, format="%.1f%%"),
            },
        )

        b1, b2 = st.columns(2)
        if b1.button("Применить изменения", type="primary", use_container_width=True):
            keep = edited[edited["Оставить"]]
            if keep.empty:
                st.error("Оставьте хотя бы один столбец")
            else:
                new_df = DF[keep["Столбец"].tolist()].copy()
                rename = dict(zip(keep["Столбец"], keep["Новое название"]))
                new_df = new_df.rename(columns=rename)
                new_types = {rename[r["Столбец"]]: r["Тип"] for _, r in keep.iterrows()}
                st.session_state.df = coerce_types(new_df, new_types)
                st.session_state.types = new_types
                st.success("Структура обновлена")
                st.rerun()
        if b2.button("Определить типы заново", use_container_width=True):
            st.session_state.types = detect_schema(DF)
            st.rerun()

    with st.expander("Добавить новый столбец"):
        n1, n2, n3 = st.columns([2, 1, 1])
        new_name = n1.text_input("Название", key="new_col_name")
        new_type = n2.selectbox("Тип", [TXT, NUM, DATE], key="new_col_type")
        if n3.button("Добавить", use_container_width=True) and new_name:
            if new_name in DF.columns:
                st.error("Столбец с таким именем уже есть")
            else:
                df2 = DF.copy()
                df2[new_name] = np.nan if new_type == NUM else ""
                st.session_state.df = df2
                st.session_state.types = {**TYPES, new_name: new_type}
                st.rerun()

    st.markdown("##### Таблица (редактируемая)")
    st.caption("Меняйте значения прямо в ячейках — всё пересчитается.")
    edited_df = st.data_editor(
        FDF, use_container_width=True, num_rows="dynamic", height=460, key="data_editor"
    )
    d1, d2 = st.columns(2)
    if d1.button("Сохранить правки данных", type="primary", use_container_width=True):
        base = DF.copy()
        common = edited_df.index.intersection(base.index)
        for c in edited_df.columns:
            if c in base.columns:
                base.loc[common, c] = edited_df.loc[common, c]
        extra = edited_df.loc[~edited_df.index.isin(base.index)]
        if not extra.empty:
            base = pd.concat([base, extra], ignore_index=False)
        st.session_state.df = coerce_types(base, TYPES)
        st.success("Данные обновлены")
        st.rerun()
    d2.download_button(
        "Скачать текущую выборку",
        to_excel_bytes(FDF),
        file_name=f"data_{date.today()}.xlsx",
        use_container_width=True,
    )


# =============================================================================
# TAB 6 — IMPORT
# =============================================================================

def stage(df: pd.DataFrame, name: str):
    df = df.dropna(axis=1, how="all")
    df.columns = [str(c).strip() or f"Колонка {i + 1}" for i, c in enumerate(df.columns)]
    # drop fully empty rows
    df = df.dropna(how="all").reset_index(drop=True)
    st.session_state.pending = df
    st.session_state.pending_name = name


with tab_import:
    st.subheader("Импорт данных — любой источник")

    src = st.radio(
        "Источник",
        ["Файл", "Google Sheets", "Ссылка", "Вставить текст", "Демо"],
        horizontal=True,
        key="imp_src",
    )

    o1, o2, o3 = st.columns(3)
    has_header = o1.checkbox("Первая строка — заголовки", True, key="imp_head")
    delim = o2.selectbox("Разделитель", ["Авто", ";", ",", "Tab", "|"], key="imp_delim")
    enc = o3.selectbox("Кодировка", ["Авто", "utf-8", "windows-1251"], key="imp_enc")
    delim_map = {"Авто": None, ";": ";", ",": ",", "Tab": "\t", "|": "|"}

    def read_text(raw: bytes) -> str:
        if enc != "Авто":
            return raw.decode(enc, errors="replace")
        for e in ("utf-8", "windows-1251", "cp1252"):
            try:
                txt = raw.decode(e)
                if txt.count(REPLACEMENT_CHAR) < 3:
                    return txt
            except Exception:
                continue
        return raw.decode("utf-8", errors="replace")

    def parse_text(txt: str, name: str):
        stripped = txt.strip()
        if stripped.startswith(("[", "{")):
            data = json.loads(stripped)
            if isinstance(data, dict):
                arr = next((v for v in data.values() if isinstance(v, list)), [data])
                data = arr
            stage(pd.json_normalize(data), name)
            return
        stage(
            pd.read_csv(
                io.StringIO(txt),
                sep=delim_map[delim],
                engine="python",
                header=0 if has_header else None,
            ),
            name,
        )

    if src == "Файл":
        up = st.file_uploader(
            "Excel / CSV / TSV / JSON",
            type=["xlsx", "xls", "csv", "tsv", "txt", "json"],
        )
        if up is not None:
            try:
                name = up.name.lower()
                if name.endswith((".xlsx", ".xls")):
                    xls = pd.ExcelFile(up)
                    sheet = st.selectbox("Лист книги", xls.sheet_names, key="imp_sheet")
                    if st.button("Загрузить лист", type="primary"):
                        stage(
                            pd.read_excel(
                                xls,
                                sheet_name=sheet,
                                header=0 if has_header else None,
                            ),
                            f"{up.name} · {sheet}",
                        )
                else:
                    parse_text(read_text(up.getvalue()), up.name)
            except Exception as e:
                st.error(f"Ошибка чтения: {e}")

    elif src == "Google Sheets":
        url = st.text_input(
            "Ссылка на таблицу с доступом «по ссылке»",
            placeholder="https://docs.google.com/spreadsheets/d/...",
        )
        if st.button("Загрузить из Google Sheets", type="primary") and url:
            m = re.search(r"/d/([a-zA-Z0-9\-_]+)", url)
            if not m:
                st.error("Не удалось распознать ID таблицы")
            else:
                gid = re.search(r"[#&?]gid=(\d+)", url)
                csv_url = f"https://docs.google.com/spreadsheets/d/{m.group(1)}/gviz/tq?tqx=out:csv"
                if gid:
                    csv_url += f"&gid={gid.group(1)}"
                try:
                    stage(pd.read_csv(csv_url), "Google Sheets")
                except Exception as e:
                    st.error(f"Не удалось загрузить: {e}")

    elif src == "Ссылка":
        url = st.text_input("Прямая ссылка на CSV / JSON")
        if st.button("Скачать и распознать", type="primary") and url:
            try:
                if url.lower().endswith(".json"):
                    stage(pd.read_json(url), f"URL · {url.split('/')[-1]}")
                else:
                    stage(pd.read_csv(url), f"URL · {url.split('/')[-1]}")
            except Exception as e:
                st.error(f"Ошибка загрузки: {e}")

    elif src == "Вставить текст":
        txt = st.text_area(
            "Скопируйте диапазон из Excel (Ctrl+C) и вставьте сюда",
            height=200,
            placeholder="Дата\tГород\tВыручка\n2024-01-01\tМосква\t150000",
        )
        if st.button("Распознать вставленное", type="primary") and txt.strip():
            try:
                sep = delim_map[delim]
                if sep is None:
                    sep = "\t" if "\t" in txt else None
                stage(
                    pd.read_csv(
                        io.StringIO(txt),
                        sep=sep,
                        engine="python",
                        header=0 if has_header else None,
                    ),
                    "Буфер обмена",
                )
            except Exception as e:
                st.error(f"Не удалось распознать: {e}")

    else:
        n = st.slider("Строк в демо-наборе", 100, 2000, 400, 100)
        if st.button("Загрузить демо-данные", type="primary"):
            stage(make_demo(n), f"Демо-набор ({n} строк)")

    # Preview / column pick
    pend = st.session_state.pending
    if pend is not None:
        st.divider()
        st.markdown(f"#### Предпросмотр · {st.session_state.pending_name}")
        st.caption(
            f"Найдено {nfmt(len(pend))} строк и {pend.shape[1]} столбцов. "
            "Отметьте нужные, задайте свои названия и типы."
        )

        auto = detect_schema(pend)
        cfg = pd.DataFrame({
            "Использовать": [True] * pend.shape[1],
            "Столбец": list(pend.columns),
            "Новое название": list(pend.columns),
            "Тип": [auto[c] for c in pend.columns],
            "Пример": [
                str(pend[c].dropna().iloc[0])[:40] if pend[c].notna().any() else ""
                for c in pend.columns
            ],
            "Пусто %": [round(float(pend[c].isna().mean() * 100), 1) for c in pend.columns],
            "Уникальных": [int(pend[c].nunique()) for c in pend.columns],
        })
        cfg_edit = st.data_editor(
            cfg,
            hide_index=True,
            use_container_width=True,
            key="prev_cfg",
            column_config={
                "Столбец": st.column_config.TextColumn(disabled=True),
                "Тип": st.column_config.SelectboxColumn(options=[NUM, DATE, TXT], required=True),
                "Пример": st.column_config.TextColumn(disabled=True),
                "Пусто %": st.column_config.NumberColumn(disabled=True, format="%.1f%%"),
                "Уникальных": st.column_config.NumberColumn(disabled=True),
            },
        )

        sel = cfg_edit[cfg_edit["Использовать"]]
        if not sel.empty:
            st.dataframe(
                pend[sel["Столбец"].tolist()].head(20),
                use_container_width=True,
                height=280,
            )

        a1, a2 = st.columns(2)
        if a1.button("Применить и перестроить аналитику", type="primary", use_container_width=True):
            if sel.empty:
                st.error("Выберите хотя бы один столбец")
            else:
                new_df = pend[sel["Столбец"].tolist()].copy()
                rename = dict(zip(sel["Столбец"], sel["Новое название"]))
                new_df = new_df.rename(columns=rename)
                new_types = {rename[r["Столбец"]]: r["Тип"] for _, r in sel.iterrows()}
                st.session_state.df = coerce_types(new_df, new_types)
                st.session_state.types = new_types
                st.session_state.source = st.session_state.pending_name
                st.session_state.pending = None
                for k in [k for k in list(st.session_state.keys()) if str(k).startswith("flt_")]:
                    del st.session_state[k]
                st.success("Данные загружены — фильтры и отчёты перестроены")
                st.rerun()
        if a2.button("Отмена", use_container_width=True):
            st.session_state.pending = None
            st.rerun()
