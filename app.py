"""
Дашборд сборки заказов: Google Таблица -> Streamlit -> Telegram WebApp.
Один файл. Запуск локально:  streamlit run app.py
Секреты (Streamlit Cloud -> Settings -> Secrets): блок [gcp_service_account].
"""
import html
import re
from datetime import datetime, date
from io import BytesIO
from string import Template
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

# ───────────────────────── Настройки ─────────────────────────
SHEET_ID = "12CxJMCBUHgkaj-_KbOs1aK7hx_jEYMyS3q4Hh0bHTGw"
GID = "1369918403"
CACHE_SECONDS = 120          # как часто перечитывать таблицу
TZ = ZoneInfo("Asia/Tashkent")
HIDE_FROM_DAYS_PCT = 95      # в сводке по умолчанию прячем дни с готовностью от этого процента

# Магазины, которых НЕ должно быть в сводной «Маршрутизация»
EXCLUDE_STORES = [
    "Узуммаркет",
    "Яндекс Маркет МСА",
    "Е-ком МСА",
    "Импортные поставки",
    "Вайлдберриз МСА",
]

ALL = "all"   # «все даты» в выпадающем списке
# Архив отгрузок: по нему считаем, сколько штук в среднем уезжало на каждой машине
ARCHIVE_SHEET_ID = "1hKZ8ggNLW-OY1bV8xAW7PKl50Fof2co86oxGK92YPAA"
ARCHIVE_GID = "2106470004"
ARCHIVE_TAB = "Архив"
ARCHIVE_FROM = date(2026, 10, 1)
MODES = ["Все", "Не собрано", "Собрано", "Маршрутизация", "Рейсы"]
ORDINALS = ["Первая", "Вторая", "Третья", "Четвёртая", "Пятая", "Шестая", "Седьмая", "Восьмая", "Девятая", "Десятая"]

PAGES = ["Сводка по дням", "Заказы", "Поиск заказа"]
WEEKDAYS = ["пн", "вт", "ср", "чт", "пт", "сб", "вс"]

LIGHT = dict(scheme="light", ink="#14232E", paper="#EEF2F5", card="#FFFFFF", line="#D9E1E7",
             mute="#667985", done="#0F8B8D", left="#E8A33D", track="#E3EAEE", head="#DCE8F2",
             ok_bg="#D6F0EF", ok_fg="#09595A", no_bg="#FBEBCF", no_fg="#7A4E05",
             car_bg="#FFE94A", car_fg="#14232E", band="#1565C0", hot_bg="#FBEFD2", note_bg="#FFF4D6")
DARK = dict(scheme="dark", ink="#E6EEF3", paper="#0E1820", card="#16242E", line="#2A3B47",
            mute="#8CA1AF", done="#2FB8BA", left="#F2B65A", track="#243642", head="#1F3342",
            ok_bg="#12393B", ok_fg="#8FE3E4", no_bg="#3D2E12", no_fg="#F5C77E",
            car_bg="#E5C93A", car_fg="#14232E", band="#1F5FA8", hot_bg="#3A2E16", note_bg="#3A2E16")

st.set_page_config(page_title="Сборка заказов", page_icon="📦", layout="wide",
                   initial_sidebar_state="collapsed")

# ───────────────────────── Тема (день / ночь) ─────────────────────────
if "dark" not in st.session_state:
    try:
        st.session_state.dark = (st.context.theme.type == "dark")   # как в браузере / Telegram
    except Exception:
        st.session_state.dark = False
P = DARK if st.session_state.dark else LIGHT

CSS = Template("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap');
html, body, .stApp, .stApp button, .stApp input { font-family: 'Manrope', sans-serif; }
.stApp { background:$paper; color:$ink; color-scheme:$scheme; }
#MainMenu, footer, header[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"] { display:none; }
.block-container { padding: 1.2rem 1rem 3rem; max-width: 1100px; }

.stApp p, .stApp label, .stApp li, .stApp [data-testid="stWidgetLabel"] p,
.stApp [data-testid="stMarkdownContainer"] p { color:$ink; }
.stApp [data-testid="stCaptionContainer"], .stApp [data-testid="stCaptionContainer"] p { color:$mute; }

/* кнопки и поля */
.stApp .stButton > button, .stApp .stDownloadButton > button {
  background:$card; color:$ink; border:1px solid $line; border-radius:10px; font-weight:700; }
.stApp .stButton > button:hover, .stApp .stDownloadButton > button:hover { border-color:$done; color:$ink; }
.stApp .stButton > button p, .stApp .stDownloadButton > button p { color:inherit; }
.stApp [data-testid="stButtonGroup"] button { background:$card; color:$ink; border:1px solid $line; font-weight:700; }
.stApp [data-testid="stButtonGroup"] button p { color:inherit; }
.stApp [data-testid="stButtonGroup"] button[aria-checked="true"] { background:$ink; color:$paper; border-color:$ink; }
.stApp [data-baseweb="select"] > div, .stApp [data-baseweb="input"], .stApp [data-baseweb="base-input"], .stApp input {
  background:$card !important; color:$ink !important; border-color:$line !important; }
.stApp [data-baseweb="select"] * { color:$ink !important; }
.stApp [data-baseweb="select"] svg { fill:$mute !important; }
.stApp [data-baseweb="tag"] { background:$done !important; }
.stApp [data-baseweb="tag"] * { color:#fff !important; }
[data-baseweb="popover"] ul, [data-baseweb="popover"] li, [data-baseweb="menu"] { background:$card !important; color:$ink !important; }
[data-baseweb="popover"] li * { color:$ink !important; }
[data-baseweb="popover"] li:hover, [data-baseweb="popover"] li[aria-selected="true"] { background:$head !important; }
.stApp [data-testid="stExpander"] { background:$card; border:1px solid $line; border-radius:12px; }
.stApp [data-testid="stExpander"] summary p { color:$ink; font-weight:700; }
.stApp [data-testid="stAlert"] { background:$card; border:1px solid $line; }

/* шапка */
.top { display:flex; flex-direction:column; gap:2px; }
.top h1 { font-size:1.7rem; font-weight:800; letter-spacing:-0.02em; margin:0; color:$ink; }
.top .upd { font-size:.8rem; color:$mute; }

/* карточки-кнопки: невидимая кнопка накрывает всю карточку */
[class*="st-key-tap_"] { position:relative; }
[class*="st-key-tap_"] [data-testid="stElementContainer"]:has(button) { position:absolute; inset:0; z-index:5; margin:0; width:100%; height:100%; }
[class*="st-key-tap_"] [data-testid="stButton"] { width:100%; height:100%; display:block; }
[class*="st-key-tap_"] button { display:block; width:100%; height:100%; min-height:0; opacity:0; cursor:pointer; padding:0; }
[class*="st-key-tap_"]:hover .kpi, [class*="st-key-tap_"]:hover .day { border-color:$done; }

.kpi { background:$card; border:1px solid $line; border-radius:14px; padding:14px 16px; min-height:112px; }
.kpi .v { font-size:2.1rem; font-weight:800; letter-spacing:-0.02em; line-height:1.1; color:$ink; }
.kpi .l { font-size:.88rem; font-weight:700; margin-top:4px; color:$ink; }
.kpi .s { font-size:.8rem; color:$mute; margin-top:2px; }
.kpi.left { border-left:5px solid $left; }
.kpi.all { border-left:5px solid $ink; }

.bar { height:10px; border-radius:99px; background:$track; overflow:hidden; }
.bar > i { display:block; height:100%; background:$done; border-radius:99px; }
.day { background:$card; border:1px solid $line; border-radius:12px; padding:12px 14px; margin-bottom:2px;
       display:grid; grid-template-columns: 92px 1fr 90px; gap:14px; align-items:center; }
.day .d { font-weight:800; font-size:1.05rem; line-height:1.15; color:$ink; }
.day .d small { display:block; font-weight:500; color:$mute; font-size:.78rem; }
.day .meta { font-size:.82rem; color:$mute; margin-top:6px; }
.day .meta b { color:$ink; }
.day .pct { text-align:right; font-weight:800; font-size:1.25rem; color:$ink; }
.day .pct small { display:block; font-weight:500; font-size:.76rem; color:$mute; }
.day.full { border-color:$done; }
@media (max-width: 640px) {
  .day { grid-template-columns: 1fr auto; }
  .day .mid { grid-column: 1 / -1; order: 3; }
}

.badge { display:inline-block; padding:2px 10px; border-radius:99px; font-size:.76rem; font-weight:700; white-space:nowrap; }
.badge.ok { background:$ok_bg; color:$ok_fg; }
.badge.no { background:$no_bg; color:$no_fg; }

.tbl-wrap { max-height:68vh; overflow:auto; border:1px solid $line; border-radius:12px; background:$card; }
table.tbl { border-collapse:collapse; width:100%; font-size:.85rem; color:$ink; }
.tbl th { position:sticky; top:0; background:$head; text-align:left; padding:9px 10px; font-weight:700; white-space:nowrap; color:$ink; }
.tbl td { padding:7px 10px; border-top:1px solid $line; vertical-align:top; color:$ink; }
.tbl .num { text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }
.tbl td.city { font-weight:800; }
.tbl td.shop { font-weight:700; }
.tbl tr.total td { font-weight:800; background:$head; border-top:2px solid $line; }

.hit { background:$card; border:1px solid $line; border-radius:12px; padding:12px 14px; margin-bottom:8px; color:$ink; }
.hit .n { font-weight:800; font-size:1.05rem; }
.hit .s { color:$mute; font-size:.85rem; margin:2px 0 8px; }
.note { background:$note_bg; border-left:5px solid $left; border-radius:10px; padding:10px 14px; margin:4px 0 10px; font-size:.9rem; color:$ink; }
.cars-wrap { overflow-x:auto; border:1px solid $line; border-radius:12px; background:$card; }
table.cars { border-collapse:collapse; width:100%; font-size:.88rem; color:$ink; }
.cars td { padding:6px 10px; vertical-align:top; color:$ink; border-top:1px solid $line; }
.cars tr.car td { border-top:none; }
.cars td.carname { background:$car_bg; color:$car_fg; font-weight:800; white-space:nowrap; width:1%; }
.chips { display:flex; flex-wrap:wrap; gap:6px; align-items:center; }
.chip { display:inline-flex; gap:4px; align-items:baseline; background:$track; color:$ink; border-radius:99px;
        padding:3px 10px; font-size:.8rem; font-weight:600; white-space:nowrap; }
.chip b { font-weight:800; }
.chip.ok { background:$ok_bg; color:$ok_fg; }
.chip.wait { background:$no_bg; color:$no_fg; }
.chips.big { margin:2px 0 10px; }
.chips.big .chip { font-size:.9rem; padding:5px 14px; background:$card; border:1px solid $line; }
.load { flex:1; min-width:80px; max-width:220px; height:6px; border-radius:99px; background:$track; overflow:hidden; }
.load i { display:block; height:100%; background:$done; border-radius:99px; }
.load.hi i { background:$left; }
.loadpct { font-size:.78rem; font-weight:700; color:$mute; }
.ship { background:$card; border:1px solid $line; border-left:5px solid $done; border-radius:14px; padding:14px 16px; margin-bottom:12px; }
.ship .sh-top { display:flex; justify-content:space-between; align-items:baseline; gap:10px; }
.ship .sh-t { font-weight:800; font-size:1.05rem; color:$ink; }
.ship .sh-t small { font-weight:500; color:$mute; font-size:.8rem; margin-left:6px; }
.ship .sh-p { font-weight:800; font-size:1.7rem; letter-spacing:-0.02em; color:$ink; }
.ship .bar { height:14px; margin:8px 0; }
.ship .sh-m { font-size:.85rem; color:$mute; margin-bottom:8px; }
.ship .sh-m b { color:$ink; }
.cars td.carmeta { vertical-align:middle; }
.cars tr.band td { background:$band; padding:0; height:12px; border:none; }
.cars tr.band.thin td { height:7px; }
.cars tr.hot td { background:$hot_bg; font-weight:700; }
.cars td.shop { font-weight:800; white-space:nowrap; }
.cars td.num { text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }
</style>
""")
st.markdown(CSS.safe_substitute(P), unsafe_allow_html=True)


# ───────────────────────── Данные ─────────────────────────
def _client():
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    return gspread.authorize(creds)


@st.cache_data(ttl=CACHE_SECONDS, show_spinner="Загружаю таблицу…")
def load() -> pd.DataFrame:
    sheet = _client().open_by_key(SHEET_ID)
    ws = next(w for w in sheet.worksheets() if str(w.id) == GID)
    rows = ws.get_all_values()
    head = [(h.strip() or f"col{i}") for i, h in enumerate(rows[0])]
    raw = pd.DataFrame(rows[1:], columns=head, dtype=str).fillna("")

    def find(*keys):
        for c in raw.columns:
            if all(k in c.lower() for k in keys):
                return c
        return None

    c_order, c_shop, c_addr = find("№"), find("название"), find("адрес")
    c_city, c_qty = find("город"), find("кол-во")
    c_date, c_wms = find("дата"), find("wms")

    # колонка «в сборке» без заголовка: ищем по содержимому, запасной вариант — 9-я колонка (I)
    c_pick = next((c for c in raw.columns
                   if raw[c].str.lower().str.contains("в сборке").any()), raw.columns[8])

    # оставляем только строки с заказами (без разделителей вида «13/10/2026 (заказы на вт)»)
    raw = raw[raw[c_order].str.match(r"^\s*\w{1,3}-\d{4,}")].copy()

    df = pd.DataFrame({
        "Заказ": raw[c_order].str.strip(),
        "Магазин": raw[c_shop].str.strip(),
        "Адрес": raw[c_addr].str.strip() if c_addr else "",
        "Город": raw[c_city].str.strip(),
        "Штук": pd.to_numeric(raw[c_qty].str.replace(r"\s", "", regex=True), errors="coerce").fillna(0).astype(int),
        "Дата": pd.to_datetime(raw[c_date], dayfirst=True, errors="coerce").dt.date,
        "Сборка": raw[c_pick].str.strip(),
        "Статус WMS": raw[c_wms].str.strip(),
    })
    df["Собран"] = df["Сборка"].str.lower().str.contains("в сборке")
    return df.dropna(subset=["Дата"]).reset_index(drop=True)


@st.cache_data(ttl=CACHE_SECONDS, show_spinner="Читаю отгрузки…")
def load_ship_tabs() -> dict:
    """Две вкладки таблицы отгрузок: «Маршруты New» (текущий день) и «Архив» (прошлые дни)."""
    sheet = _client().open_by_key(ARCHIVE_SHEET_ID)
    sheets = sheet.worksheets()

    def read(ws):
        rows = ws.get_all_values()
        head = [(" ".join(h.split()) or f"col{i}") for i, h in enumerate(rows[0])]
        return pd.DataFrame(rows[1:], columns=head, dtype=str).fillna("")

    arch = next((w for w in sheets if str(w.id) == ARCHIVE_GID), None) or sheet.worksheet(ARCHIVE_TAB)
    cur = next((w for w in sheets if "маршрут" in w.title.lower() and "new" in w.title.lower()), None)
    return {"archive": read(arch), "new": read(cur) if cur else None}


def fmt_date(d: date) -> str:
    return d.strftime("%d.%m")


def n(x) -> str:
    return f"{int(x):,}".replace(",", " ")


def esc(x) -> str:
    return html.escape(str(x))


def day_summary(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("Дата").agg(Заказов=("Заказ", "count"), Собрано=("Собран", "sum"), Штук=("Штук", "sum"))
    g["Штук собрано"] = df[df["Собран"]].groupby("Дата")["Штук"].sum().reindex(g.index).fillna(0).astype(int)
    g["Осталось"] = g["Заказов"] - g["Собрано"]
    g["Процент"] = (g["Собрано"] / g["Заказов"] * 100).round().astype(int)
    return g.sort_index()


# ── нормализация названий магазинов (регистр, пробелы, дефисы, латиница-двойники) ──
_LOOKALIKE = str.maketrans("eaopcxykmt", "еаорсхукмт")


def norm_store(x: str) -> str:
    x = str(x).lower().replace("ё", "е").translate(_LOOKALIKE)
    return re.sub(r"[\s\-_.]+", "", x)


EXCLUDED_KEYS = {norm_store(x) for x in EXCLUDE_STORES}


def split_routing(frame: pd.DataFrame):
    """Возвращает (сводная по городу/магазину/адресу, исключённые заказы)."""
    mask = frame["Магазин"].map(lambda s: norm_store(s) in EXCLUDED_KEYS)
    keep, dropped = frame[~mask], frame[mask]
    g = (keep.groupby(["Город", "Магазин", "Адрес"], as_index=False)["Штук"].sum()
         .sort_values(["Город", "Магазин", "Адрес"]).reset_index(drop=True))
    return g, dropped


def archive_stats(raw: pd.DataFrame) -> dict:
    """Средние по архиву: сколько штук / точек / во сколько уезжала машина №1, №2, … и к какой машине
    обычно относится магазин. Машина = рейс внутри дня (по времени фактической отгрузки)."""
    raw = raw.copy()
    raw.columns = [" ".join(str(c).split()) for c in raw.columns]

    def find(*keys):
        return next((c for c in raw.columns if all(k in c.lower() for k in keys)), None)

    c_shop, c_qty, c_trip = find("название"), find("кол-во"), find("рейс")
    c_stat, c_fact = find("статус"), find("факт")
    if not all([c_shop, c_qty, c_trip, c_fact]):
        raise ValueError("В архиве не найдены колонки: название магазина, кол-во штук, рейс, дата отгрузки факт.")

    d = pd.DataFrame({
        "shop": raw[c_shop].str.strip(),
        "qty": pd.to_numeric(raw[c_qty].str.replace(r"\s", "", regex=True), errors="coerce").fillna(0),
        "trip": raw[c_trip].str.strip(),
        "fact": pd.to_datetime(raw[c_fact].str.strip(), format="%d.%m.%Y %H:%M", errors="coerce"),
    })
    if c_stat:
        d["status"] = raw[c_stat].str.strip().str.upper()
    d = d[d["fact"].notna() & (d["trip"] != "")]
    d = d[d["fact"].dt.date >= ARCHIVE_FROM]
    if c_stat:
        shipped = d[d["status"] == "ОТГРУЖЕН"]
        d = shipped if not shipped.empty else d
    if d.empty:
        raise ValueError(f"В архиве нет отгрузок с {ARCHIVE_FROM:%d.%m.%Y}.")

    d["day"] = d["fact"].dt.date
    trips = (d.groupby(["day", "trip"]).agg(qty=("qty", "sum"), stores=("shop", "nunique"), t=("fact", "min"))
             .reset_index())
    trips["tn"] = pd.to_numeric(trips["trip"], errors="coerce").fillna(999)
    trips = trips.sort_values(["day", "t", "tn"])
    trips["k"] = trips.groupby("day").cumcount() + 1          # какая по счёту машина в этот день
    trips["min"] = trips["t"].dt.hour * 60 + trips["t"].dt.minute
    by_k = trips.groupby("k").agg(qty=("qty", "mean"), stores=("stores", "mean"), time=("min", "mean"))

    d = d.merge(trips[["day", "trip", "k"]], on=["day", "trip"])
    d["key"] = d["shop"].map(norm_store)
    return dict(
        qty={int(k): float(v) for k, v in by_k["qty"].items()},
        stores={int(k): float(v) for k, v in by_k["stores"].items()},
        time={int(k): float(v) for k, v in by_k["time"].items()},
        rank={k: float(v) for k, v in d.groupby("key")["k"].mean().items()},
        avg_all=float(trips["qty"].mean()),
        n_days=int(trips["day"].nunique()),
        date_from=trips["day"].min(), date_to=trips["day"].max(),
    )


def get_archive_stats() -> dict:
    """Средние пересчитываются при каждом чтении архива (кэш = CACHE_SECONDS, кнопка «Обновить» сбрасывает)."""
    return archive_stats(load_ship_tabs()["archive"])


_SHIP_COLS = ["Заказ", "Магазин", "Штук", "Дата", "Рейс", "Отгружен", "Факт"]


def parse_ship(raw) -> pd.DataFrame:
    """Строки вкладки отгрузок -> Заказ, Магазин, Штук, Дата (план), Рейс, Отгружен, Факт."""
    if raw is None or raw.empty:
        return pd.DataFrame(columns=_SHIP_COLS)
    raw = raw.copy()
    raw.columns = [" ".join(str(c).split()) for c in raw.columns]

    def find(*keys):
        return next((c for c in raw.columns if all(k in c.lower() for k in keys)), None)

    c_order, c_shop, c_qty = find("№"), find("название"), find("кол-во")
    c_plan, c_trip, c_stat, c_fact = find("дата", "план"), find("рейс"), find("статус"), find("факт")
    if not all([c_order, c_shop, c_qty, c_plan]):
        return pd.DataFrame(columns=_SHIP_COLS)
    d = pd.DataFrame({
        "Заказ": raw[c_order].str.strip(),
        "Магазин": raw[c_shop].str.strip(),
        "Штук": pd.to_numeric(raw[c_qty].str.replace(r"\s", "", regex=True), errors="coerce").fillna(0).astype(int),
        "Дата": pd.to_datetime(raw[c_plan], dayfirst=True, errors="coerce").dt.date,
        "Рейс": raw[c_trip].str.strip() if c_trip else "",
        "Отгружен": (raw[c_stat].str.strip().str.upper() == "ОТГРУЖЕН") if c_stat else False,
        "Факт": (pd.to_datetime(raw[c_fact].str.strip(), format="%d.%m.%Y %H:%M", errors="coerce")
                 if c_fact else pd.NaT),
    })
    return d[d["Заказ"].str.match(r"^\s*\w{1,3}-\d{4,}") & d["Дата"].notna()].reset_index(drop=True)


def build_ship(tabs: dict) -> pd.DataFrame:
    """Все заказы из «Маршруты New» + «Архив» без дублей (приоритет у «Маршруты New»), без исключённых магазинов."""
    d = pd.concat([parse_ship(tabs.get("new")), parse_ship(tabs.get("archive"))], ignore_index=True)
    d = d.drop_duplicates("Заказ", keep="first")
    return d[~d["Магазин"].map(lambda x: norm_store(x) in EXCLUDED_KEYS)].reset_index(drop=True)


def pick_ship_date(ship: pd.DataFrame, today: date):
    ds = sorted(ship["Дата"].unique())
    if not ds:
        return None
    if today in ds:
        return today
    prev = [x for x in ds if x <= today]
    return prev[-1] if prev else ds[0]


def ship_summary(ship: pd.DataFrame, d: date) -> dict:
    s_ = ship[ship["Дата"] == d]
    tr = (s_[s_["Рейс"] != ""].groupby("Рейс")
          .agg(qty=("Штук", "sum"), done=("Отгружен", "all"), t=("Факт", "min")).reset_index())
    tr["tn"] = pd.to_numeric(tr["Рейс"], errors="coerce").fillna(999)
    tr = tr.sort_values(["tn", "Рейс"])
    return dict(total=len(s_), done=int(s_["Отгружен"].sum()),
                pcs=int(s_["Штук"].sum()), pcs_done=int(s_.loc[s_["Отгружен"], "Штук"].sum()),
                trips=tr.to_dict("records"))


def hhmm(minutes: float) -> str:
    m = int(round(minutes))
    return f"{m // 60:02d}:{m % 60:02d}"


def deadline_minutes(addr: str):
    m = re.search(r"\bдо\s*(\d{1,2})[:.](\d{2})", str(addr).lower())
    return int(m.group(1)) * 60 + int(m.group(2)) if m else None


def ordinal(k: int) -> str:
    return f"{ORDINALS[k - 1]} машина" if k <= len(ORDINALS) else f"{k}-я машина"


def plan_trips(pivot: pd.DataFrame, stats: dict) -> list:
    """Пример раскладки точек по машинам: магазины идут в том порядке, как их обычно развозили
    (по архиву), и набираются, пока машина не достигнет среднего объёма для её номера."""
    def target(k):
        return stats["qty"].get(k) or stats["avg_all"]

    rows = []
    for r in pivot.itertuples(index=False):
        dl = deadline_minutes(r.Адрес)
        hot = dl is not None and dl < 12 * 60                     # «ДО 09:45» и т.п. — срочные
        rk = stats["rank"].get(norm_store(r.Магазин), 98.0)       # новых магазинов в архиве нет — в хвост
        rows.append(dict(city=r.Город, shop=r.Магазин, addr=r.Адрес, qty=int(r.Штук),
                         rk=min(rk, 1.0) if hot else rk, dl=dl, hot=hot))

    cities = sorted({x["city"] for x in rows}, key=lambda c: (norm_store(c) != norm_store("Ташкент"), c))
    machines = []
    for city in cities:                                           # другой город — всегда отдельная машина
        items = sorted((x for x in rows if x["city"] == city),
                       key=lambda x: (x["rk"], x["dl"] if x["dl"] is not None else 9999, x["shop"]))
        local, cur, cur_sum = [], [], 0
        for it in items:
            t = target(len(machines) + len(local) + 1)
            if cur and (cur_sum >= t or (cur_sum + it["qty"] > t * 1.15 and cur_sum >= t * 0.6)):
                local.append(cur)
                cur, cur_sum = [], 0
            cur.append(it)
            cur_sum += it["qty"]
        if cur:
            local.append(cur)
        if len(local) >= 2:                                       # маленький «хвост» присоединяем к предыдущей
            last, prev = sum(x["qty"] for x in local[-1]), sum(x["qty"] for x in local[-2])
            k_last = len(machines) + len(local)
            if last < 0.35 * target(k_last) and prev + last <= 1.35 * target(k_last - 1):
                local[-2] = local[-2] + local.pop()
        machines.extend(local)

    return [dict(n=i, stores=st_, qty=sum(x["qty"] for x in st_), target=target(i), time=stats["time"].get(i))
            for i, st_ in enumerate(machines, 1)]


def machine_meta(m: dict) -> str:
    return f"{n(m['qty'])} шт · точек: {len(m['stores'])}"


def machine_chips(m: dict) -> str:
    pct = round(m["qty"] / m["target"] * 100) if m["target"] else 0
    hi = " hi" if pct > 115 else ""
    chips = (f"<span class='chip'><b>{n(m['qty'])}</b> шт</span>"
             f"<span class='chip'>Точек <b>{len(m['stores'])}</b></span>")
    if m["time"] is not None:
        chips += f"<span class='chip'>🕐 {hhmm(m['time'])}</span>"
    return (f"<div class='chips'>{chips}<span class='load{hi}'><i style='width:{min(pct, 100)}%'></i></span>"
            f"<span class='loadpct'>{pct}%</span></div>")


def trips_table(machines: list) -> str:
    parts = []
    for m in machines:
        parts.append(f"<tr class='car'><td class='carname'>{ordinal(m['n'])}</td>"
                     f"<td class='carmeta' colspan='2'>{machine_chips(m)}</td></tr>"
                     "<tr class='band thin'><td colspan='3'></td></tr>")
        for s_ in m["stores"]:
            parts.append(f"<tr class='{'hot' if s_['hot'] else ''}'><td class='shop'>{esc(s_['shop'])}</td>"
                         f"<td>{esc(s_['addr'])}</td><td class='num'>{n(s_['qty'])}</td></tr>")
        parts.append("<tr class='band'><td colspan='3'></td></tr>")
    return f"<div class='cars-wrap'><table class='cars'>{''.join(parts)}</table></div>"


def trips_excel(machines: list, day_label: str) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook()
    ws = wb.active
    ws.title = "Рейсы (пример)"
    ws.append([f"ПРИМЕР маршрутизации на {day_label}. Это расчёт по средним значениям архива, а не факт."])
    ws.merge_cells("A1:C1")
    ws["A1"].font = Font(bold=True, color="C00000")
    ws["A1"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 32
    yellow, blue = PatternFill("solid", fgColor="FFEB3B"), PatternFill("solid", fgColor="1565C0")
    hot_fill = PatternFill("solid", fgColor="FBEFD2")

    def band(h):
        ws.append([None, None, None])
        for c in ws[ws.max_row]:
            c.fill = blue
        ws.row_dimensions[ws.max_row].height = h

    for m in machines:
        ws.append([ordinal(m["n"]), machine_meta(m), None])
        ws.cell(row=ws.max_row, column=1).fill = yellow
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
        band(6)
        for s_ in m["stores"]:
            ws.append([s_["shop"], s_["addr"], s_["qty"]])
            row = ws.max_row
            ws.cell(row=row, column=1).font = Font(bold=True)
            ws.cell(row=row, column=3).number_format = "#,##0"
            if s_["hot"]:
                for c in ws[row]:
                    c.fill = hot_fill
                    c.font = Font(bold=True)
        band(10)
    ws.append(["Итого", None, sum(m["qty"] for m in machines)])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
    ws.cell(row=ws.max_row, column=3).font = Font(bold=True)
    ws.cell(row=ws.max_row, column=3).number_format = "#,##0"
    for col, w in zip("ABC", (24, 80, 12)):
        ws.column_dimensions[col].width = w
    out = BytesIO()
    wb.save(out)
    return out.getvalue()


# ───────────────────────── Excel ─────────────────────────
def backlog_excel(frame: pd.DataFrame) -> bytes:
    """Все несобранные заказы по датам: ДАТА и КОЛ-ВО (шт)."""
    from openpyxl.styles import Font
    b = (frame[~frame["Собран"]].groupby("Дата")["Штук"].sum().reset_index()
         .rename(columns={"Дата": "ДАТА", "Штук": "КОЛ-ВО"}).sort_values("ДАТА"))
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        b.to_excel(xw, sheet_name="Бэклог", index=False)
        ws = xw.sheets["Бэклог"]
        last = len(b) + 1
        ws.cell(row=last + 1, column=1, value="ИТОГО")
        ws.cell(row=last + 1, column=2, value=f"=SUM(B2:B{last})")
        for c in ws[1] + ws[last + 1]:
            c.font = Font(bold=True)
        for r in range(2, last + 1):
            ws.cell(row=r, column=1).number_format = "DD.MM.YYYY"
        for r in range(2, last + 2):
            ws.cell(row=r, column=2).number_format = "#,##0"
        ws.column_dimensions["A"].width = 16
        ws.column_dimensions["B"].width = 14
    return out.getvalue()


def routing_excel(g: pd.DataFrame) -> bytes:
    """Сводная в том же виде, что и в Excel: Город / Название магазина / Адрес / Сумма штук / Общий итог."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook()
    ws = wb.active
    ws.title = "Маршрутизация"
    head = ["Город", "Название магазина", "Адрес магазина", "Сумма по полю кол-во штук в заказе"]
    ws.append(head)
    fill = PatternFill("solid", fgColor="DCE8F2")
    for c in ws[1]:
        c.font, c.fill = Font(bold=True), fill
        c.alignment = Alignment(wrap_text=True, vertical="center")
    prev_city = None
    for r in g.itertuples(index=False):
        ws.append([r.Город if r.Город != prev_city else None, r.Магазин, r.Адрес, int(r.Штук)])
        if r.Город != prev_city:
            ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
        ws.cell(row=ws.max_row, column=2).font = Font(bold=True)
        prev_city = r.Город
    last = ws.max_row
    ws.append(["Общий итог", None, None, f"=SUM(D2:D{last})"])
    for c in ws[ws.max_row]:
        c.font, c.fill = Font(bold=True), fill
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=4).number_format = "#,##0"
    for col, w in zip("ABCD", (22, 26, 70, 22)):
        ws.column_dimensions[col].width = w
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"
    out = BytesIO()
    wb.save(out)
    return out.getvalue()


# ───────────────────────── Мелкие виджеты ─────────────────────────
def tap(key: str, card_html: str, on_click, args=()):
    """Карточка, на которую можно нажать (поверх неё невидимая кнопка)."""
    with st.container(key=f"tap_{key}"):
        st.markdown(card_html, unsafe_allow_html=True)
        st.button("открыть", key=f"btn_{key}", on_click=on_click, args=args)


def go_orders(mode: str, d):
    st.session_state.page = PAGES[1]
    st.session_state.f_mode = mode
    st.session_state.f_date = d


def toggle_theme():
    st.session_state.dark = not st.session_state.dark


def badge(collected: bool) -> str:
    return '<span class="badge ok">собран</span>' if collected else '<span class="badge no">не собран</span>'


def orders_table(frame: pd.DataFrame, ship_map: dict = None, limit: int = 1500) -> str:
    ship_map = ship_map or {}
    rows = frame.head(limit).to_dict("records")
    show_ship = any(r["Заказ"] in ship_map for r in rows)

    def ship_cell(r):
        info = ship_map.get(r["Заказ"])
        if not info:
            return "<td>—</td>"
        done, trip, fact = info
        if done:
            when = f" {fact:%H:%M}" if pd.notna(fact) else ""
            return f"<td><span class='badge ok'>отгружен · рейс {esc(trip)}{when}</span></td>"
        return (f"<td><span class='badge no'>рейс {esc(trip)}</span></td>" if trip else "<td>—</td>")

    head = ("<tr><th>Дата</th><th>Заказ</th><th>Магазин</th><th>Город</th><th class='num'>Штук</th>"
            "<th>Статус</th><th>WMS</th>" + ("<th>Отгрузка</th>" if show_ship else "") + "</tr>")
    body = []
    for r in rows:
        body.append(
            f"<tr><td>{fmt_date(r['Дата'])}</td><td><b>{esc(r['Заказ'])}</b></td><td>{esc(r['Магазин'])}</td>"
            f"<td>{esc(r['Город'])}</td><td class='num'>{n(r['Штук'])}</td><td>{badge(r['Собран'])}</td>"
            f"<td>{esc(r['Статус WMS'] or '—')}</td>" + (ship_cell(r) if show_ship else "") + "</tr>")
    return f"<div class='tbl-wrap'><table class='tbl'>{head}{''.join(body)}</table></div>"


def routing_table(g: pd.DataFrame) -> str:
    head = ("<tr><th>Город</th><th>Название магазина</th><th>Адрес магазина</th>"
            "<th class='num'>Сумма по полю кол-во штук в заказе</th></tr>")
    body, prev = [], None
    for r in g.itertuples(index=False):
        city = (esc(r.Город) or "—") if r.Город != prev else ""
        body.append(f"<tr><td class='city'>{city}</td><td class='shop'>{esc(r.Магазин)}</td>"
                    f"<td>{esc(r.Адрес)}</td><td class='num'>{n(r.Штук)}</td></tr>")
        prev = r.Город
    body.append(f"<tr class='total'><td colspan='3'>Общий итог</td><td class='num'>{n(g['Штук'].sum())}</td></tr>")
    return f"<div class='tbl-wrap'><table class='tbl'>{head}{''.join(body)}</table></div>"


# ───────────────────────── Загрузка и шапка ─────────────────────────
try:
    df = load()
except Exception as e:
    st.error("Не удалось прочитать таблицу. Проверьте, что таблица открыта для почты "
             "сервисного аккаунта (роль «Читатель») и ключ добавлен в Secrets.")
    st.caption(str(e))
    st.stop()

now = datetime.now(TZ)
h1, h2, h3 = st.columns([6, 1.4, 1.4], vertical_alignment="bottom")
with h1:
    st.markdown(f"""<div class="top"><h1>📦 Сборка заказов</h1>
    <span class="upd">Обновлено {now:%H:%M} · данные кэшируются на {CACHE_SECONDS // 60} мин</span></div>""",
                unsafe_allow_html=True)
with h2:
    st.button("☀️ День" if st.session_state.dark else "🌙 Ночь", on_click=toggle_theme, width="stretch")
with h3:
    if st.button("🔄 Обновить", width="stretch"):
        st.cache_data.clear()
        st.rerun()

if df.empty:
    st.info("В таблице нет заказов.")
    st.stop()

summary = day_summary(df)
dates = list(summary.index)

ship, ship_err = None, ""
try:
    ship = build_ship(load_ship_tabs())
except Exception as e:                      # дашборд работает и без таблицы отгрузок
    ship_err = str(e)

# состояние фильтров (до создания виджетов)
if st.session_state.get("f_date") not in [ALL] + dates:
    st.session_state.pop("f_date", None)
if "f_date" not in st.session_state:
    st.session_state.f_date = next((x for x in dates if x >= now.date()), dates[-1])
st.session_state.setdefault("f_mode", "Все")
st.session_state.setdefault("page", PAGES[0])

page = st.segmented_control("Раздел", PAGES, key="page", label_visibility="collapsed") or PAGES[0]
date_label = lambda x: "Все даты" if x == ALL else f"{fmt_date(x)}, {WEEKDAYS[x.weekday()]}"

# ───────────────────────── 1. Сводка по дням ─────────────────────────
if page == PAGES[0]:
    total = len(df)
    left = int((~df["Собран"]).sum())
    pcs_total = int(df["Штук"].sum())
    pcs_left = int(df.loc[~df["Собран"], "Штук"].sum())

    if ship is not None and not ship.empty:
        sd = pick_ship_date(ship, now.date())
        sm = ship_summary(ship, sd)
        pct_ship = round(sm["pcs_done"] / sm["pcs"] * 100) if sm["pcs"] else 0
        trip_chips = "".join(
            (f"<span class='chip ok'>Рейс {esc(t['Рейс'])} · ✓ {t['t']:%H:%M}</span>" if t["done"] and pd.notna(t["t"])
             else f"<span class='chip ok'>Рейс {esc(t['Рейс'])} · ✓</span>" if t["done"]
             else f"<span class='chip wait'>Рейс {esc(t['Рейс'])} · в работе</span>")
            for t in sm["trips"])
        ship_html = f"""<div class="ship"><div class="sh-top">
          <div class="sh-t">{'Отгрузка сегодня' if sd == now.date() else 'Последняя отгрузка'}
          <small>{fmt_date(sd)}, {WEEKDAYS[sd.weekday()]}</small></div><div class="sh-p">{pct_ship}%</div></div>
          <div class="bar"><i style="width:{pct_ship}%"></i></div>
          <div class="sh-m">Отгружено <b>{n(sm['done'])}</b> из {n(sm['total'])} заказов ·
          <b>{n(sm['pcs_done'])}</b> из {n(sm['pcs'])} шт · осталось <b>{n(sm['total'] - sm['done'])}</b> зак.
          ({n(sm['pcs'] - sm['pcs_done'])} шт)</div>
          <div class="chips">{trip_chips}</div></div>"""
        if sd in dates:
            tap("ship", ship_html, go_orders, ("Все", sd))
        else:
            st.markdown(ship_html, unsafe_allow_html=True)
    elif ship_err:
        st.caption("Статус отгрузки недоступен: проверьте доступ к таблице отгрузок.")

    k1, k2, k3 = st.columns(3)
    with k1:
        tap("all", f"""<div class="kpi all" title="Открыть список всех заказов"><div class="v">{n(total)}</div>
        <div class="l">всего заказов</div><div class="s">{n(pcs_total)} шт</div></div>""", go_orders, ("Все", ALL))
    with k2:
        tap("left", f"""<div class="kpi left" title="Открыть список несобранных заказов"><div class="v">{n(left)}</div>
        <div class="l">не собрано заказов</div><div class="s">из {n(total)}</div></div>""", go_orders, ("Не собрано", ALL))
    with k3:
        tap("pcs", f"""<div class="kpi left" title="Открыть список несобранных заказов"><div class="v">{n(pcs_left)}</div>
        <div class="l">бэклог, шт</div><div class="s">из {n(pcs_total)} шт</div></div>""", go_orders, ("Не собрано", ALL))
    st.caption("Нажмите на карточку, чтобы открыть список заказов.")

    st.download_button(
        "⬇️ Выгрузить бэклог в Excel (дата и кол-во)",
        data=backlog_excel(df),
        file_name=f"backlog_{now:%Y-%m-%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )

    show_all = st.toggle(f"Показать все дни (по умолчанию только дни с готовностью ниже {HIDE_FROM_DAYS_PCT}%)",
                         value=False)
    shown = summary if show_all else summary[summary["Процент"] < HIDE_FROM_DAYS_PCT]
    hidden = len(summary) - len(shown)
    if hidden:
        st.caption(f"Скрыто дней с готовностью {HIDE_FROM_DAYS_PCT}% и выше: {hidden}")
    if shown.empty:
        st.success(f"Все дни собраны на {HIDE_FROM_DAYS_PCT}% и выше.")

    for d, r in shown.iterrows():
        full = "full" if r["Осталось"] == 0 else ""
        left_pcs = r["Штук"] - r["Штук собрано"]
        tap(f"day_{d:%Y%m%d}", f"""<div class="day {full}">
          <div class="d">{fmt_date(d)}<small>{WEEKDAYS[d.weekday()]}, {d:%Y}</small></div>
          <div class="mid"><div class="bar"><i style="width:{r['Процент']}%"></i></div>
          <div class="meta">собрано <b>{r['Собрано']}</b> зак. ({n(r['Штук собрано'])} шт) · осталось <b>{r['Осталось']}</b> зак.
          ({n(left_pcs)} шт) · всего {r['Заказов']} зак., {n(r['Штук'])} шт</div></div>
          <div class="pct">{r['Процент']}%<small>собрано</small></div></div>""",
            go_orders, ("Собрано" if r["Осталось"] == 0 else "Не собрано", d))

# ───────────────────────── 2. Заказы ─────────────────────────
if page == PAGES[1]:
    st.selectbox("Дата отгрузки", [ALL] + dates, key="f_date", format_func=date_label)
    st.segmented_control("Показать", MODES, key="f_mode")
    mode = st.session_state.f_mode or "Все"

    sub = df if st.session_state.f_date == ALL else df[df["Дата"] == st.session_state.f_date]
    done_pcs = int(sub.loc[sub["Собран"], "Штук"].sum())
    left_pcs = int(sub.loc[~sub["Собран"], "Штук"].sum())
    st.caption(f"За выбранную дату: собрано {n(done_pcs)} шт · осталось {n(left_pcs)} шт · "
               f"заказов {len(sub)} · всего {n(sub['Штук'].sum())} шт")
    if mode == "Маршрутизация":
        # сводная по выбранной дате: все заказы даты, без исключённых магазинов
        pivot, dropped = split_routing(sub)
        if pivot.empty:
            st.info("Нет заказов для сводной.")
        else:
            st.caption(f"В сводной: {n(len(sub) - len(dropped))} заказов · {n(pivot['Штук'].sum())} шт · "
                       f"точек: {len(pivot)}. Не включены магазины: {', '.join(EXCLUDE_STORES)}"
                       + (f" (исключено {n(len(dropped))} заказов, {n(dropped['Штук'].sum())} шт)." if len(dropped) else "."))
            label = "все_даты" if st.session_state.f_date == ALL else f"{st.session_state.f_date:%Y-%m-%d}"
            st.download_button(
                "⬇️ Скачать сводную в Excel",
                data=routing_excel(pivot),
                file_name=f"marshrutizatsiya_{label}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )
            st.markdown(routing_table(pivot), unsafe_allow_html=True)
    elif mode == "Рейсы":
        st.markdown('<div class="note">⚠️ <b>Это пример маршрутизации, а не факт.</b> Точки разложены по машинам '
                    'по средним показателям прошлых отгрузок (архив), реальные рейсы могут отличаться.</div>',
                    unsafe_allow_html=True)
        if st.session_state.f_date == ALL:
            st.info("Выберите конкретную дату отгрузки: рейсы считаются на один день.")
        else:
            pivot, dropped = split_routing(sub)
            try:
                stats = get_archive_stats()
            except Exception as e:
                stats = None
                mail = ""
                try:
                    mail = st.secrets["gcp_service_account"].get("client_email", "")
                except Exception:
                    pass
                st.warning("Не удалось прочитать архив отгрузок. Откройте таблицу архива для почты сервисного "
                           f"аккаунта {mail} (роль «Читатель»).")
                st.caption(str(e))
            if stats is not None:
                if pivot.empty:
                    st.info("Нет заказов для расчёта.")
                else:
                    machines = plan_trips(pivot, stats)
                    st.markdown(f"<div class='chips big'><span class='chip'>Машин <b>{len(machines)}</b></span>"
                                f"<span class='chip'>Точек <b>{len(pivot)}</b></span>"
                                f"<span class='chip'>Штук <b>{n(pivot['Штук'].sum())}</b></span></div>",
                                unsafe_allow_html=True)
                    st.download_button(
                        "⬇️ Скачать рейсы в Excel (пример)",
                        data=trips_excel(machines, f"{st.session_state.f_date:%d.%m.%Y}"),
                        file_name=f"reisy_primer_{st.session_state.f_date:%Y-%m-%d}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width="stretch",
                    )
                    st.markdown(trips_table(machines), unsafe_allow_html=True)
    else:
        if mode == "Не собрано":
            sub = sub[~sub["Собран"]]
        elif mode == "Собрано":
            sub = sub[sub["Собран"]]

        sub = sub.sort_values(["Дата", "Заказ"])
        st.markdown(f"**{mode}: {len(sub)} заказов · {n(sub['Штук'].sum())} шт**")
        if sub.empty:
            st.info("Заказов нет.")
        else:
            ship_map = ({r.Заказ: (r.Отгружен, r.Рейс, r.Факт) for r in ship.itertuples()} if ship is not None else {})
            st.markdown(orders_table(sub, ship_map), unsafe_allow_html=True)
            if len(sub) > 1500:
                st.caption("Показаны первые 1500 заказов, уточните дату.")

# ───────────────────────── 3. Поиск ─────────────────────────
if page == PAGES[2]:
    q = st.text_input("Номер заказа", placeholder="например, 1M-0024629 или 0024629").strip()
    if q:
        hits = df[df["Заказ"].str.contains(q, case=False, regex=False)]
        if hits.empty:
            st.warning(f"Заказ «{q}» не найден. Можно вводить только последние цифры номера.")
        else:
            st.caption(f"Найдено: {len(hits)}")
            for r in hits.head(30).to_dict("records"):
                st.markdown(f"""<div class="hit"><div class="n">{esc(r['Заказ'])} &nbsp;{badge(r['Собран'])}</div>
                <div class="s">{esc(r['Магазин'])} · {esc(r['Город'])}</div>
                Отгрузка: <b>{fmt_date(r['Дата'])}</b> · {n(r['Штук'])} шт · WMS: <b>{esc(r['Статус WMS'] or '—')}</b></div>""",
                            unsafe_allow_html=True)
    else:
        st.info("Введите номер заказа или его часть.")
