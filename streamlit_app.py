# -*- coding: utf-8 -*-
"""
ADAPTIVE BI — монолитное Streamlit-приложение.
Загрузите ЛЮБУЮ таблицу (Excel / CSV / JSON / Google Sheets / ссылка / буфер) —
приложение само определит типы столбцов, построит умные фильтры по вашим данным
и позволит собрать любые графики и сводные таблицы.

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

# ──────────────────────────────────────────────────────────────────────────────
#  КОНФИГ
# ──────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Adaptive BI — любые данные, любые отчёты",
    page_icon="📊",
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
html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif; }
[data-testid="stAppViewContainer"] {
  background:
    radial-gradient(1000px 540px at 88% -12%, rgba(124,58,237,.16), transparent 62%),
    radial-gradient(760px 460px at -12% 112%, rgba(6,182,212,.11), transparent 60%),
    #070a18;
  color: #e2e8f0;
}
[data-testid="stHeader"] { background: rgba(7,10,24,.75); backdrop-filter: blur(10px); }
[data-testid="stSidebar"] { background: #0a0d1f; border-right: 1px solid rgba(255,255,255,.07); }
[data-testid="stSidebar"] * { color: #cbd5e1; }
h1, h2, h3 { font-family: 'Unbounded', sans-serif !important; letter-spacing: -.01em; color: #fff; }
h1 { font-size: 1.7rem !important; }
h2 { font-size: 1.25rem !important; }
h3 { font-size: 1.02rem !important; }
[data-testid="stMetric"] {
  background: linear-gradient(135deg, rgba(25,31,58,.9), rgba(15,18,35,.95));
  border: 1px solid rgba(255,255,255,.09);
  border-radius: 16px; padding: 16px 18px;
  box-shadow: 0 12px 30px -14px rgba(0,0,0,.8);
}
[data-testid="stMetricLabel"] {
  color: #94a3b8 !important; font-size: .72rem !important;
  text-transform: uppercase; letter-spacing: .09em; font-weight: 700;
}
[data-testid="stMetricValue"] {
  font-family: 'JetBrains Mono', monospace !important;
  font-size: 1.5rem !important; color: #fff !important;
}
.stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 1px solid rgba(255,255,255,.08); }
.stTabs [data-baseweb="tab"] {
  background: transparent; border-radius: 10px 10px 0 0;
  padding: 8px 16px; font-size: .82rem; font-weight: 700; color: #94a3b8;
}
.stTabs [aria-selected="true"] {
  background: rgba(168,85,247,.16) !important; color: #fff !important;
  box-shadow: inset 0 -2px 0 #a855f7;
}
.stButton > button, .stDownloadButton > button {
  border-radius: 12px; border: 1px solid rgba(255,255,255,.12);
  background: rgba(255,255,255,.04); color: #e2e8f0;
  font-weight: 700; font-size: .8rem;
}
.stButton > button:hover, .stDownloadButton > button:hover {
  border-color: #a855f7; background: rgba(168,85,247,.16); color: #fff;
}
.stButton > button[kind="primary"] {
  background: linear-gradient(135deg,#7e22ce,#ec4899); border: none; color: #fff;
}
[data-baseweb="input"], [data-baseweb="select"] > div, .stTextArea textarea {
  background: #141835 !important; border-color: rgba(255,255,255,.1) !important;
  border-radius: 12px !important; color: #e2e8f0 !important;
}
[data-testid="stExpander"] {
  border: 1px solid rgba(255,255,255,.08); border-radius: 14px;
  background: rgba(18,22,48,.7);
}
.chip {
  display:inline-flex; align-items:center; gap:6px; padding:4px 11px;
  margin:2px 3px 2px 0; border-radius:999px; font-size:.7rem; font-weight:600; border:1px solid;
}
.chip-num  { color:#6ee7b7; border-color:rgba(16,185,129,.35); background:rgba(16,185,129,.1); }
.chip-date { color:#67e8f9; border-color:rgba(6,182,212,.35);  background:rgba(6,182,212,.1); }
.chip-str  { color:#fcd9a8; border-color:rgba(251,146,60,.35); background:rgba(251,146,60,.1); }
.card {
  background: rgba(18,22,48,.86); border:1px solid rgba(255,255,255,.08);
  border-radius:16px; padding:16px 18px; box-shadow:0 12px 30px -14px rgba(0,0,0,.75);
}
.hint { font-size:.75rem; color:#94a3b8; }
.brand {
  font-family:'Unbounded',sans-serif; font-weight:900; font-size:1.05rem;
  background:linear-gradient(135deg,#c084fc,#22d3ee); -webkit-background-clip:text;
  -webkit-text-fill-color:transparent;
}
#MainMenu, footer { visibility: hidden; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

NUM, DATE, TXT = "число", "дата", "текст"
TYPE_CHIP = {NUM: "chip-num", DATE: "chip-date", TXT: "chip-str"}
TYPE_ICON = {NUM: "#", DATE: "📅", TXT: "T"}

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
    "Тепловая карта", "Ящик с усами", "Скрипка", "Гистограмма", "Индикатор KPI", "Сводная таблица",
]
RAW_CHARTS = {"Ящик с усами", "Скрипка", "Гистограмма"}


# ──────────────────────────────────────────────────────────────────────────────
#  УТИЛИТЫ
# ──────────────────────────────────────────────────────────────────────────────

def plot_kwargs() -> dict:
    """Совместимость Streamlit width / use_container_width."""
    try:
        return {"width": "stretch"}
    except Exception:
        return {"use_container_width": True}


def show_df(df: pd.DataFrame, **kwargs):
    """Безопасный показ dataframe (Arrow-compatible)."""
    st.dataframe(make_arrow_safe(df), **{**plot_kwargs(), **kwargs})


def make_arrow_safe(df: pd.DataFrame) -> pd.DataFrame:
    """Приводит dataframe к типам, которые Arrow сериализует без ошибок."""
    out = df.copy()
    for c in out.columns:
        s = out[c]
        if pd.api.types.is_datetime64_any_dtype(s):
            out[c] = pd.to_datetime(s, errors="coerce")
        elif pd.api.types.is_bool_dtype(s):
            continue
        elif pd.api.types.is_numeric_dtype(s):
            out[c] = pd.to_numeric(s, errors="coerce")
        else:
            # object / string / mixed -> plain string, без None/NaT
            out[c] = s.map(lambda x: "" if pd.isna(x) else str(x)).astype(str)
    # index тоже может ломать Arrow
    if out.index.name is not None or not isinstance(out.index, pd.RangeIndex):
        out = out.reset_index(drop=False) if out.index.name else out.reset_index(drop=True)
    return out


def to_num(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")
    cleaned = (
        s.astype(str)
        .str.replace("\u00a0", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^0-9.\-eE]", "", regex=True)
        .replace({"": np.nan, "-": np.nan, ".": np.nan, "None": np.nan, "nan": np.nan})
    )
    return pd.to_numeric(cleaned, errors="coerce")


def to_date(s: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce")
    return pd.to_datetime(s, errors="coerce", dayfirst=True)


def looks_like_date_str(val: str) -> bool:
    return bool(re.search(r"\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}", val))


def looks_like_number_str(val: str) -> bool:
    """Без \u escape-последовательностей — совместимо с Arrow regex."""
    v = val.replace("\u00a0", " ").strip()
    if not v or len(v) > 24:
        return False
    # цифры, пробелы, точки/запятые, знаки, валюты
    return bool(re.fullmatch(r"[0-9\s.,+\-%₽$€]+", v))


def detect_type(s: pd.Series) -> str:
    sample = s.dropna()
    if sample.empty:
        return TXT

    # отсечь пустые строки
    if sample.dtype == object or str(sample.dtype).startswith("string"):
        sample = sample[sample.astype(str).str.strip() != ""]
    if sample.empty:
        return TXT
    sample = sample.head(250)

    if pd.api.types.is_numeric_dtype(sample):
        return NUM
    if pd.api.types.is_datetime64_any_dtype(sample):
        return DATE

    as_str = sample.astype(str)
    # даты
    date_hits = as_str.map(looks_like_date_str)
    if float(date_hits.mean()) >= 0.55 and float(to_date(sample).notna().mean()) >= 0.55:
        return DATE
    # числа
    num_hits = as_str.map(looks_like_number_str)
    if float(num_hits.mean()) >= 0.55 and float(to_num(sample).notna().mean()) >= 0.55:
        return NUM
    return TXT


def detect_schema(df: pd.DataFrame) -> dict:
    return {str(c): detect_type(df[c]) for c in df.columns}


def coerce_types(df: pd.DataFrame, types: dict) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c) for c in out.columns]
    for c, t in types.items():
        if c not in out.columns:
            continue
        if t == NUM:
            out[c] = to_num(out[c])
        elif t == DATE:
            out[c] = to_date(out[c])
        else:
            out[c] = out[c].map(lambda x: "" if pd.isna(x) else str(x))
    return out


def guess_primary_measure(types: dict):
    nums = [c for c, t in types.items() if t == NUM]
    if not nums:
        return None
    for kw in ["выруч", "revenue", "сумм", "total", "прибыл", "profit", "amount", "продаж", "цена", "price"]:
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
    revenue = rng.integers(25_000, 520_000, n)
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
        "Выручка": revenue.astype(float),
        "Прибыль": (revenue * margin).round(0).astype(float),
        "Количество": rng.integers(1, 60, n).astype(float),
        "Скидка %": rng.integers(0, 25, n).astype(float),
        "Статус": rng.choice(["Оплачено", "В работе", "Просрочен"], n, p=[0.78, 0.16, 0.06]),
    })


def to_excel_bytes(df: pd.DataFrame, sheet: str = "Данные") -> bytes:
    buf = io.BytesIO()
    safe = make_arrow_safe(df)
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        safe.to_excel(w, index=False, sheet_name=sheet[:31])
    return buf.getvalue()


def fmt_num(v) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return str(v)
    if abs(v) >= 1_000_000_000:
        return f"{v/1_000_000_000:,.2f} млрд".replace(",", " ")
    if abs(v) >= 1_000_000:
        return f"{v/1_000_000:,.2f} млн".replace(",", " ")
    return f"{v:,.0f}".replace(",", " ")


def safe_colname(c) -> str:
    return str(c)


# ──────────────────────────────────────────────────────────────────────────────
#  STATE
# ──────────────────────────────────────────────────────────────────────────────

def init_state():
    if "df" not in st.session_state:
        demo = make_demo()
        types = detect_schema(demo)
        st.session_state.df = coerce_types(demo, types)
        st.session_state.types = types
        st.session_state.source = "Демо-набор (12 столбцов)"
    st.session_state.setdefault("reports", [])
    st.session_state.setdefault("pending", None)
    st.session_state.setdefault("pending_name", "")


init_state()

DF: pd.DataFrame = st.session_state.df
TYPES: dict = st.session_state.types

# нормализуем имена столбцов
if any(not isinstance(c, str) for c in DF.columns):
    DF = DF.copy()
    DF.columns = [str(c) for c in DF.columns]
    st.session_state.df = DF
    TYPES = {str(k): v for k, v in TYPES.items()}
    st.session_state.types = TYPES

cols_num = [c for c in DF.columns if TYPES.get(c) == NUM]
cols_date = [c for c in DF.columns if TYPES.get(c) == DATE]
cols_txt = [c for c in DF.columns if TYPES.get(c) == TXT]
cols_dim = cols_txt + cols_date
PRIMARY = guess_primary_measure(TYPES)


# ──────────────────────────────────────────────────────────────────────────────
#  УМНЫЕ ФИЛЬТРЫ
# ──────────────────────────────────────────────────────────────────────────────

def render_smart_filters(df: pd.DataFrame, types: dict):
    mask = pd.Series(True, index=df.index)
    active = []

    search = st.sidebar.text_input(
        "🔎 Поиск по всем столбцам",
        key="flt__search",
        placeholder="ищем во всех полях...",
    )
    if search:
        joined = df.astype(str).agg(" ".join, axis=1).str.lower()
        mask &= joined.str.contains(re.escape(search.lower()), na=False)
        active.append(f'поиск: "{search}"')

    st.sidebar.caption("Фильтры сформированы по вашим столбцам")

    for col in df.columns:
        col = str(col)
        t = types.get(col, TXT)
        s = df[col]
        key = f"flt_{col}"

        if t == NUM:
            vals = to_num(s).dropna()
            if vals.empty:
                continue
            lo, hi = float(vals.min()), float(vals.max())
            if not np.isfinite(lo) or not np.isfinite(hi) or lo >= hi:
                continue
            step = max((hi - lo) / 100.0, 0.01)
            sel = st.sidebar.slider(f"{TYPE_ICON[NUM]} {col}", lo, hi, (lo, hi), step=step, key=key)
            if sel != (lo, hi):
                mask &= to_num(s).between(sel[0], sel[1])
                active.append(f"{col}: {fmt_num(sel[0])}–{fmt_num(sel[1])}")

        elif t == DATE:
            d = to_date(s).dropna()
            if d.empty:
                continue
            mn, mx = d.min().date(), d.max().date()
            if mn >= mx:
                continue
            sel = st.sidebar.date_input(
                f"{TYPE_ICON[DATE]} {col}",
                value=(mn, mx),
                min_value=mn,
                max_value=mx,
                key=key,
            )
            if isinstance(sel, (tuple, list)) and len(sel) == 2 and tuple(sel) != (mn, mx):
                dd = to_date(s).dt.date
                mask &= dd.between(sel[0], sel[1])
                active.append(f"{col}: {sel[0]} → {sel[1]}")

        else:
            uniq = s.dropna().astype(str)
            uniq = uniq[uniq.str.strip() != ""].unique().tolist()
            if len(uniq) <= 1:
                continue
            if len(uniq) <= 60:
                sel = st.sidebar.multiselect(
                    f"{TYPE_ICON[TXT]} {col}",
                    sorted(uniq),
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

    return df.loc[mask].copy(), active


# ──────────────────────────────────────────────────────────────────────────────
#  АГРЕГАЦИЯ
# ──────────────────────────────────────────────────────────────────────────────

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
    return df[col].map(lambda x: "н/д" if pd.isna(x) or str(x).strip() == "" else str(x))


def build_pivot(df, row, col, measure, agg_label, gran="Месяц") -> pd.DataFrame:
    if df is None or df.empty or not row or not measure or row not in df.columns or measure not in df.columns:
        return pd.DataFrame()

    agg = AGGS.get(agg_label, "sum")
    tmp = pd.DataFrame({"__row": dim_series(df, row, gran)})
    values = to_num(df[measure]) if agg in NUMERIC_AGGS else df[measure]
    tmp["__val"] = values.values

    if col and col != "— нет —" and col in df.columns:
        tmp["__col"] = dim_series(df, col, gran).values
        out = tmp.groupby(["__row", "__col"], dropna=False)["__val"].agg(agg).unstack(fill_value=0)
        out.columns = [str(c) for c in out.columns]
    else:
        out = tmp.groupby("__row", dropna=False)["__val"].agg(agg).to_frame(f"{agg_label} · {measure}")

    out.index = out.index.map(str)
    out.index.name = str(row)
    out = out.apply(pd.to_numeric, errors="coerce").fillna(0)
    return out.sort_index()


def shape_pivot(pt: pd.DataFrame, top_n=None, sort_desc=True) -> pd.DataFrame:
    if pt is None or pt.empty:
        return pd.DataFrame()
    out = pt.copy()
    if sort_desc:
        order = out.sum(axis=1).sort_values(ascending=False).index
        out = out.loc[order]
    if top_n:
        out = out.head(int(top_n))
    return out


# ──────────────────────────────────────────────────────────────────────────────
#  ГРАФИКИ
# ──────────────────────────────────────────────────────────────────────────────

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


def show_chart(fig, height=420):
    st.plotly_chart(style_fig(fig, height), **plot_kwargs())


def render_chart(kind, pt, raw, row, measure, agg_label, height=420):
    if kind in RAW_CHARTS:
        if raw is None or raw.empty or row not in raw.columns or measure not in raw.columns:
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
        show_chart(fig, height)
        return

    if pt is None or pt.empty:
        st.info("Нет данных под текущими фильтрами.")
        return

    idx_name = str(pt.index.name or "Категория")
    long = pt.reset_index().melt(id_vars=pt.index.name, var_name="Серия", value_name="Значение")
    long = long.rename(columns={pt.index.name: idx_name})
    long[idx_name] = long[idx_name].astype(str)
    long["Серия"] = long["Серия"].astype(str)
    long["Значение"] = pd.to_numeric(long["Значение"], errors="coerce").fillna(0)

    multi = pt.shape[1] > 1
    color = "Серия" if multi else None
    totals = pt.sum(axis=1)
    totals.index = totals.index.map(str)

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
            fig = px.sunburst(long, path=[idx_name, "Серия"], values="Значение", color="Значение",
                              color_continuous_scale="Purples")
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
        fig = px.scatter(d, x=idx_name, y="Значение", color="Значение", size="Значение",
                         color_continuous_scale="Plasma")
    elif kind == "Пузырьковая":
        d = totals.reset_index()
        d.columns = [idx_name, "Значение"]
        d["Ранг"] = d["Значение"].rank()
        fig = px.scatter(d, x="Ранг", y="Значение", size="Значение", color=idx_name,
                         hover_name=idx_name, size_max=55)
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
        show_df(show, height=height)
        return

    show_chart(fig, height)


# ──────────────────────────────────────────────────────────────────────────────
#  SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown('<div class="brand">◆ ADAPTIVE BI</div>', unsafe_allow_html=True)
    st.caption("Любые таблицы · Любые отчёты")

    st.markdown(
        f'<div class="card" style="padding:12px 14px">'
        f'<div style="font-size:.72rem;color:#94a3b8">ИСТОЧНИК</div>'
        f'<div style="font-size:.8rem;font-weight:700;color:#fff">{st.session_state.source}</div>'
        f'<div style="font-size:.7rem;color:#94a3b8;margin-top:4px">'
        f'Строк: <b style="color:#c084fc">{len(DF):,}</b> · '
        f'Столбцов: <b style="color:#22d3ee">{DF.shape[1]}</b></div></div>'.replace(",", " "),
        unsafe_allow_html=True,
    )

    chips = "".join(
        f'<span class="chip {TYPE_CHIP.get(TYPES.get(c, TXT), "chip-str")}">'
        f'{TYPE_ICON.get(TYPES.get(c, TXT), "T")} {c}</span>'
        for c in list(DF.columns)[:10]
    )
    st.markdown(f'<div style="margin:10px 0">{chips}</div>', unsafe_allow_html=True)

    st.divider()
    if st.button("🔄 Сбросить все фильтры", use_container_width=True, key="btn_reset_filters"):
        for k in [k for k in list(st.session_state.keys()) if str(k).startswith("flt_")]:
            del st.session_state[k]
        st.rerun()

    FDF, ACTIVE = render_smart_filters(DF, TYPES)


# ──────────────────────────────────────────────────────────────────────────────
#  HEADER
# ──────────────────────────────────────────────────────────────────────────────

head_l, head_r = st.columns([3, 1])
with head_l:
    st.markdown("# Аналитика любых данных")
    st.markdown(
        f'<div class="hint">Отфильтровано <b style="color:#c084fc">{len(FDF):,}</b> из '
        f'<b>{len(DF):,}</b> строк'.replace(",", " ")
        + (f' · Активно фильтров: <b style="color:#22d3ee">{len(ACTIVE)}</b>' if ACTIVE else "")
        + "</div>",
        unsafe_allow_html=True,
    )
with head_r:
    st.download_button(
        "⬇️ Выгрузить в Excel",
        to_excel_bytes(FDF),
        file_name=f"export_{date.today()}.xlsx",
        use_container_width=True,
        key="btn_export_main",
    )

if ACTIVE:
    st.markdown(
        "".join(f'<span class="chip chip-num">🔎 {a}</span>' for a in ACTIVE),
        unsafe_allow_html=True,
    )

tab_over, tab_build, tab_dash, tab_pivot, tab_data, tab_import = st.tabs(
    ["📈 Обзор", "🪄 Конструктор отчётов", "🗂 Мои дашборды", "🧮 Сводные срезы", "📋 Данные", "⬆️ Импорт"]
)


# ──────────────────────────────────────────────────────────────────────────────
#  ОБЗОР
# ──────────────────────────────────────────────────────────────────────────────

with tab_over:
    if not cols_num:
        st.warning("В данных нет числовых столбцов — загрузите таблицу с числами для KPI.")
    else:
        st.subheader("Автоматические показатели")
        kpi_cols = st.columns(min(4, len(cols_num)))
        for i, c in enumerate(cols_num[:4]):
            v = to_num(FDF[c]) if c in FDF.columns else pd.Series(dtype=float)
            total = float(v.sum()) if len(v) else 0.0
            avg = float(v.mean()) if len(v) else 0.0
            base = float(to_num(DF[c]).sum()) if c in DF.columns else 0.0
            delta = (total - base) / base * 100 if base else 0.0
            with kpi_cols[i]:
                st.metric(
                    c,
                    fmt_num(total),
                    delta=f"{delta:+.1f}% от всего" if abs(delta) > 0.01 else "весь объём",
                )
                st.caption(f"Среднее: **{fmt_num(avg)}** · строк: {len(v):,}".replace(",", " "))

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
            meas = o2.selectbox("Показатель", cols_num, key="ov_meas", index=meas_idx)
            gran = o3.selectbox("Гранулярность", ["День", "Неделя", "Месяц", "Квартал", "Год"], index=2, key="ov_gran")
            pt = build_pivot(FDF, dim, None, meas, "Сумма", gran)
            kind = "Область" if TYPES.get(dim) == DATE else "Столбцы"
            render_chart(kind, pt, FDF, dim, meas, "Сумма", height=340)

        with c2:
            st.markdown("##### Структура")
            pdim_opts = cols_txt or cols_dim or list(DF.columns)
            pdim = st.selectbox("Измерение", pdim_opts, key="ov_pdim")
            ppt = shape_pivot(build_pivot(FDF, pdim, None, meas, "Сумма"), 10, True)
            render_chart("Кольцевая", ppt, FDF, pdim, meas, "Сумма", height=300)

        st.divider()
        st.markdown("##### 💡 Авто-инсайты по вашим данным")
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
            last, prev = float(pt.iloc[-1].sum()), float(pt.iloc[-2].sum())
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


# ──────────────────────────────────────────────────────────────────────────────
#  КОНСТРУКТОР
# ──────────────────────────────────────────────────────────────────────────────

with tab_build:
    st.subheader("Конструктор отчётов — как сводная таблица, но с любым графиком")
    st.caption("Выберите столбцы для группировки и показатель. Приложение построит любую визуализацию.")

    left, right = st.columns([1, 2])

    with left:
        title = st.text_input("Название отчёта", key="b_title", placeholder="например: Выручка по городам")
        row_opts = cols_dim or list(DF.columns)
        row = st.selectbox("Группировать по (строки)", row_opts, key="b_row")
        col2_opts = ["— нет —"] + [c for c in cols_dim if c != row]
        col2 = st.selectbox("Разбить по (колонки)", col2_opts, key="b_col")
        agg_label = st.selectbox("Агрегация", list(AGGS.keys()), key="b_agg")
        meas_pool = cols_num if AGGS[agg_label] in NUMERIC_AGGS else list(DF.columns)
        if not meas_pool:
            meas_pool = list(DF.columns)
        meas_idx = meas_pool.index(PRIMARY) if PRIMARY in meas_pool else 0
        meas = st.selectbox("Показатель (мера)", meas_pool, key="b_meas", index=meas_idx)
        kind = st.selectbox("Тип визуализации", CHARTS, key="b_kind")

        gran = "Месяц"
        if TYPES.get(row) == DATE or TYPES.get(col2) == DATE:
            gran = st.selectbox("Гранулярность дат", ["День", "Неделя", "Месяц", "Квартал", "Год"], index=2, key="b_gran")

        cc1, cc2 = st.columns(2)
        sort_desc = cc1.checkbox("Сортировать", True, key="b_sort")
        use_top = cc2.checkbox("Только топ-15", False, key="b_top")

        if st.button("➕ Сохранить на дашборд", type="primary", use_container_width=True, key="b_save"):
            st.session_state.reports.append(dict(
                title=title or f"{agg_label} {meas} по {row}",
                row=row, col=col2, meas=meas, agg=agg_label, kind=kind,
                gran=gran, sort=sort_desc, top=use_top,
            ))
            st.success("Отчёт добавлен во вкладку «Мои дашборды»")

    with right:
        pt = shape_pivot(build_pivot(FDF, row, col2, meas, agg_label, gran), 15 if use_top else None, sort_desc)
        st.caption(
            f"Категорий: **{len(pt)}** · Серий: **{pt.shape[1] if not pt.empty else 0}** · "
            f"{agg_label} по «{meas}»"
        )
        render_chart(kind, pt, FDF, row, meas, agg_label, height=430)
        if not pt.empty and kind != "Сводная таблица":
            with st.expander("Показать таблицу с числами"):
                show = pt.copy()
                show["ИТОГО"] = show.sum(axis=1)
                show_df(show)

    st.divider()
    st.markdown("##### ⚡ Идеи отчётов по вашим столбцам")
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
                    st.session_state.reports.append(dict(
                        title=t, row=d, col="— нет —", meas=m, agg="Сумма", kind=k,
                        gran="Месяц", sort=True, top=True,
                    ))
                    st.success("Добавлено в «Мои дашборды»")


# ──────────────────────────────────────────────────────────────────────────────
#  ДАШБОРДЫ
# ──────────────────────────────────────────────────────────────────────────────

with tab_dash:
    st.subheader("Мои дашборды")
    reports = st.session_state.reports
    if not reports:
        st.info("Пока пусто. Соберите отчёт в «Конструкторе» и сохраните его сюда.")
    else:
        c1, c2 = st.columns([3, 1])
        c1.caption(f"Сохранено отчётов: **{len(reports)}** · все реагируют на текущие фильтры")
        if c2.button("🗑 Очистить всё", use_container_width=True, key="dash_clear"):
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
                        if h2.button("✕", key=f"del_{idx}"):
                            st.session_state.reports.pop(idx)
                            st.rerun()
                        p = shape_pivot(
                            build_pivot(FDF, rep["row"], rep["col"], rep["meas"], rep["agg"], rep.get("gran", "Месяц")),
                            15 if rep.get("top") else None,
                            rep.get("sort", True),
                        )
                        render_chart(rep["kind"], p, FDF, rep["row"], rep["meas"], rep["agg"], height=300)


# ──────────────────────────────────────────────────────────────────────────────
#  СВОДНЫЕ СРЕЗЫ
# ──────────────────────────────────────────────────────────────────────────────

with tab_pivot:
    st.subheader("Сводные срезы по любым полям")
    p1, p2, p3, p4, p5 = st.columns(5)
    prow_opts = cols_dim or list(DF.columns)
    prow = p1.selectbox("Строки", prow_opts, key="p_row")
    pcol = p2.selectbox("Колонки", ["— нет —"] + [c for c in cols_dim if c != prow], key="p_col")
    pagg = p3.selectbox("Агрегация", list(AGGS.keys()), key="p_agg")
    pmeas_pool = cols_num if AGGS[pagg] in NUMERIC_AGGS else list(DF.columns)
    if not pmeas_pool:
        pmeas_pool = list(DF.columns)
    pmeas = p4.selectbox("Мера", pmeas_pool, key="p_meas")
    pgran = p5.selectbox("Даты", ["День", "Неделя", "Месяц", "Квартал", "Год"], index=2, key="p_gran")

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
        total_row = pd.DataFrame([table.sum(axis=0)], index=["ИТОГО"])
        table = pd.concat([table, total_row])
        show_df(table.reset_index(), height=460)
        o3.download_button(
            "⬇️ Скачать срез в Excel",
            to_excel_bytes(table.reset_index(), "Срез"),
            file_name=f"pivot_{date.today()}.xlsx",
            use_container_width=True,
            key="p_dl",
        )


# ──────────────────────────────────────────────────────────────────────────────
#  ДАННЫЕ
# ──────────────────────────────────────────────────────────────────────────────

with tab_data:
    st.subheader("Данные и управление столбцами")

    with st.expander("✏️ Свои названия столбцов и типы", expanded=False):
        st.caption("Переименуйте столбцы — изменения применятся ко всем отчётам.")
        cfg = pd.DataFrame({
            "Столбец": [str(c) for c in DF.columns],
            "Новое название": [str(c) for c in DF.columns],
            "Тип": [TYPES.get(c, TXT) for c in DF.columns],
            "Оставить": [True] * DF.shape[1],
            "Пример": [
                ("" if DF[c].dropna().empty else str(DF[c].dropna().iloc[0])[:40])
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
                "Пусто %": st.column_config.NumberColumn(disabled=True, format="%.1f"),
            },
        )

        b1, b2 = st.columns(2)
        if b1.button("💾 Применить изменения", type="primary", use_container_width=True, key="apply_cols"):
            keep = edited[edited["Оставить"]]
            if keep.empty:
                st.error("Оставьте хотя бы один столбец")
            else:
                new_df = DF[keep["Столбец"].tolist()].copy()
                rename = dict(zip(keep["Столбец"], keep["Новое название"].map(str)))
                new_df = new_df.rename(columns=rename)
                new_types = {str(rename[r["Столбец"]]): r["Тип"] for _, r in keep.iterrows()}
                st.session_state.df = coerce_types(new_df, new_types)
                st.session_state.types = new_types
                st.success("Структура обновлена")
                st.rerun()
        if b2.button("🔍 Определить типы заново", use_container_width=True, key="redetect"):
            st.session_state.types = detect_schema(DF)
            st.session_state.df = coerce_types(DF, st.session_state.types)
            st.rerun()

    with st.expander("➕ Добавить новый столбец"):
        n1, n2, n3 = st.columns([2, 1, 1])
        new_name = n1.text_input("Название", key="new_col_name")
        new_type = n2.selectbox("Тип", [TXT, NUM, DATE], key="new_col_type")
        if n3.button("Добавить", use_container_width=True, key="add_col_btn") and new_name:
            name = str(new_name).strip()
            if not name:
                st.error("Пустое имя")
            elif name in DF.columns:
                st.error("Столбец с таким именем уже есть")
            else:
                df2 = DF.copy()
                df2[name] = np.nan if new_type == NUM else ""
                st.session_state.df = df2
                st.session_state.types = {**TYPES, name: new_type}
                st.rerun()

    st.markdown("##### Таблица")
    st.caption("Просмотр текущей отфильтрованной выборки.")
    show_df(FDF.head(500), height=460)
    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇️ Скачать текущую выборку",
        to_excel_bytes(FDF),
        file_name=f"data_{date.today()}.xlsx",
        use_container_width=True,
        key="dl_fdf",
    )
    d2.caption(f"Показаны первые {min(500, len(FDF))} из {len(FDF)} строк")


# ──────────────────────────────────────────────────────────────────────────────
#  ИМПОРТ
# ──────────────────────────────────────────────────────────────────────────────

def stage(df: pd.DataFrame, name: str):
    if df is None or df.empty:
        st.error("Пустая таблица")
        return
    df = df.dropna(axis=1, how="all").copy()
    # уникальные строковые имена
    new_cols = []
    seen = {}
    for i, c in enumerate(df.columns):
        base = str(c).strip() or f"Колонка {i+1}"
        if base in seen:
            seen[base] += 1
            base = f"{base}_{seen[base]}"
        else:
            seen[base] = 1
        new_cols.append(base)
    df.columns = new_cols
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
        for e in ("utf-8-sig", "utf-8", "windows-1251", "cp1252"):
            try:
                txt = raw.decode(e)
                if txt.count("\ufffd") < 3:
                    return txt
            except Exception:
                continue
        return raw.decode("utf-8", errors="replace")

    def parse_text(txt: str, name: str):
        stripped = txt.strip()
        if not stripped:
            st.error("Пустой текст")
            return
        if stripped.startswith(("[", "{")):
            data = json.loads(stripped)
            if isinstance(data, dict):
                arr = next((v for v in data.values() if isinstance(v, list)), [data])
                data = arr
            stage(pd.json_normalize(data), name)
            return
        sep = delim_map[delim]
        if sep is None:
            if "\t" in stripped.splitlines()[0]:
                sep = "\t"
            elif ";" in stripped.splitlines()[0] and stripped.splitlines()[0].count(";") >= stripped.splitlines()[0].count(","):
                sep = ";"
            else:
                sep = ","
        stage(
            pd.read_csv(
                io.StringIO(txt),
                sep=sep,
                engine="python",
                header=0 if has_header else None,
                on_bad_lines="skip",
            ),
            name,
        )

    if src == "Файл":
        up = st.file_uploader(
            "Excel / CSV / TSV / JSON",
            type=["xlsx", "xls", "csv", "tsv", "txt", "json"],
            key="imp_file",
        )
        if up is not None:
            try:
                name_l = up.name.lower()
                if name_l.endswith((".xlsx", ".xls")):
                    xls = pd.ExcelFile(up)
                    sheet = st.selectbox("Лист книги", xls.sheet_names, key="imp_sheet")
                    if st.button("Загрузить лист", type="primary", key="imp_xlsx_btn"):
                        stage(
                            pd.read_excel(xls, sheet_name=sheet, header=0 if has_header else None),
                            f"{up.name} · {sheet}",
                        )
                else:
                    if st.button("Загрузить файл", type="primary", key="imp_file_btn"):
                        parse_text(read_text(up.getvalue()), up.name)
            except Exception as e:
                st.error(f"Ошибка чтения: {e}")

    elif src == "Google Sheets":
        url = st.text_input(
            "Ссылка на таблицу с доступом «по ссылке»",
            placeholder="https://docs.google.com/spreadsheets/d/...",
            key="imp_gs_url",
        )
        if st.button("Загрузить из Google Sheets", type="primary", key="imp_gs_btn") and url:
            m = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
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
        url = st.text_input("Прямая ссылка на CSV / JSON", key="imp_url")
        if st.button("Скачать и распознать", type="primary", key="imp_url_btn") and url:
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
            key="imp_paste",
        )
        if st.button("Распознать вставленное", type="primary", key="imp_paste_btn") and txt.strip():
            try:
                parse_text(txt, "Буфер обмена")
            except Exception as e:
                st.error(f"Не удалось распознать: {e}")

    else:
        n = st.slider("Строк в демо-наборе", 100, 2000, 400, 100, key="imp_demo_n")
        if st.button("Загрузить демо-данные", type="primary", key="imp_demo_btn"):
            stage(make_demo(n), f"Демо-набор ({n} строк)")

    pend = st.session_state.pending
    if pend is not None:
        st.divider()
        st.markdown(f"#### Предпросмотр · {st.session_state.pending_name}")
        st.caption(
            f"Найдено {len(pend):,} строк и {pend.shape[1]} столбцов. "
            "Отметьте нужные, задайте свои названия и типы.".replace(",", " ")
        )

        auto = detect_schema(pend)
        cfg = pd.DataFrame({
            "Использовать": [True] * pend.shape[1],
            "Столбец": [str(c) for c in pend.columns],
            "Новое название": [str(c) for c in pend.columns],
            "Тип": [auto.get(str(c), TXT) for c in pend.columns],
            "Пример": [
                ("" if pend[c].dropna().empty else str(pend[c].dropna().iloc[0])[:40])
                for c in pend.columns
            ],
            "Пусто %": [round(float(pend[c].isna().mean() * 100), 1) for c in pend.columns],
            "Уникальных": [int(pend[c].nunique(dropna=True)) for c in pend.columns],
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
                "Пусто %": st.column_config.NumberColumn(disabled=True, format="%.1f"),
                "Уникальных": st.column_config.NumberColumn(disabled=True),
            },
        )

        sel = cfg_edit[cfg_edit["Использовать"]]
        if not sel.empty:
            preview_cols = [c for c in sel["Столбец"].tolist() if c in pend.columns]
            show_df(pend[preview_cols].head(20), height=280)

        a1, a2 = st.columns(2)
        if a1.button("🚀 Применить и перестроить аналитику", type="primary", use_container_width=True, key="imp_apply"):
            if sel.empty:
                st.error("Выберите хотя бы один столбец")
            else:
                cols = [c for c in sel["Столбец"].tolist() if c in pend.columns]
                new_df = pend[cols].copy()
                rename = {
                    str(r["Столбец"]): str(r["Новое название"]).strip() or str(r["Столбец"])
                    for _, r in sel.iterrows()
                    if str(r["Столбец"]) in new_df.columns
                }
                # уникальные имена после rename
                final_names = {}
                used = set()
                for old, new in rename.items():
                    base = new
                    k = 1
                    while new in used:
                        k += 1
                        new = f"{base}_{k}"
                    used.add(new)
                    final_names[old] = new
                new_df = new_df.rename(columns=final_names)
                new_types = {
                    final_names[str(r["Столбец"])]: r["Тип"]
                    for _, r in sel.iterrows()
                    if str(r["Столбец"]) in final_names
                }
                st.session_state.df = coerce_types(new_df, new_types)
                st.session_state.types = new_types
                st.session_state.source = st.session_state.pending_name
                st.session_state.pending = None
                for k in [k for k in list(st.session_state.keys()) if str(k).startswith("flt_")]:
                    del st.session_state[k]
                st.success("Данные загружены — фильтры и отчёты перестроены")
                st.rerun()
        if a2.button("Отмена", use_container_width=True, key="imp_cancel"):
            st.session_state.pending = None
            st.rerun()
