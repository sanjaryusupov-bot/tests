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
PAGES = ["Сводка по дням", "Заказы", "Поиск заказа"]
WEEKDAYS = ["пн", "вт", "ср", "чт", "пт", "сб", "вс"]

LIGHT = dict(scheme="light", ink="#14232E", paper="#EEF2F5", card="#FFFFFF", line="#D9E1E7",
             mute="#667985", done="#0F8B8D", left="#E8A33D", track="#E3EAEE", head="#DCE8F2",
             ok_bg="#D6F0EF", ok_fg="#09595A", no_bg="#FBEBCF", no_fg="#7A4E05")
DARK = dict(scheme="dark", ink="#E6EEF3", paper="#0E1820", card="#16242E", line="#2A3B47",
            mute="#8CA1AF", done="#2FB8BA", left="#F2B65A", track="#243642", head="#1F3342",
            ok_bg="#12393B", ok_fg="#8FE3E4", no_bg="#3D2E12", no_fg="#F5C77E")

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
</style>
""")
st.markdown(CSS.safe_substitute(P), unsafe_allow_html=True)


# ───────────────────────── Данные ─────────────────────────
@st.cache_data(ttl=CACHE_SECONDS, show_spinner="Загружаю таблицу…")
def load() -> pd.DataFrame:
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    sheet = gspread.authorize(creds).open_by_key(SHEET_ID)
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


def orders_table(frame: pd.DataFrame, limit: int = 1500) -> str:
    head = ("<tr><th>Дата</th><th>Заказ</th><th>Магазин</th><th>Город</th>"
            "<th class='num'>Штук</th><th>Статус</th><th>WMS</th></tr>")
    body = []
    for r in frame.head(limit).to_dict("records"):
        body.append(
            f"<tr><td>{fmt_date(r['Дата'])}</td><td><b>{esc(r['Заказ'])}</b></td><td>{esc(r['Магазин'])}</td>"
            f"<td>{esc(r['Город'])}</td><td class='num'>{n(r['Штук'])}</td><td>{badge(r['Собран'])}</td>"
            f"<td>{esc(r['Статус WMS'] or '—')}</td></tr>")
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
    st.segmented_control("Показать", ["Все", "Не собрано", "Собрано", "Маршрутизация"], key="f_mode")
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
            st.markdown(orders_table(sub), unsafe_allow_html=True)
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
