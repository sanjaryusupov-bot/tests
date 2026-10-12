"""
Дашборд сборки заказов: Google Таблица -> Streamlit -> Telegram WebApp.
Один файл. Запуск локально:  streamlit run app.py
"""
from datetime import datetime, date
from zoneinfo import ZoneInfo

import pandas as pd
import gspread
import plotly.graph_objects as go
import streamlit as st
from google.oauth2.service_account import Credentials

# ───────────────────────── Настройки ─────────────────────────
SHEET_ID = "12CxJMCBUHgkaj-_KbOs1aK7hx_jEYMyS3q4Hh0bHTGw"
GID = "1369918403"
CACHE_SECONDS = 120          # как часто перечитывать таблицу
TZ = ZoneInfo("Asia/Tashkent")

# Палитра
INK, PAPER, CARD, LINE, MUTE = "#14232E", "#EEF2F5", "#FFFFFF", "#D9E1E7", "#667985"
DONE, LEFT = "#0F8B8D", "#E8A33D"          # собрано / осталось
WEEKDAYS = ["пн", "вт", "ср", "чт", "пт", "сб", "вс"]

st.set_page_config(page_title="Сборка заказов", page_icon="📦", layout="wide",
                   initial_sidebar_state="collapsed")

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"], .stApp {{ font-family: 'Manrope', sans-serif; color: {INK}; }}
.stApp {{ background: {PAPER}; }}
#MainMenu, footer, header[data-testid="stHeader"] {{ display: none; }}
.block-container {{ padding: 1.2rem 1rem 3rem; max-width: 1100px; }}

.top {{ display:flex; justify-content:space-between; align-items:flex-end; gap:12px; flex-wrap:wrap; margin-bottom:14px; }}
.top h1 {{ font-size: 1.7rem; font-weight: 800; letter-spacing: -0.02em; margin:0; }}
.top .upd {{ font-size: .8rem; color: {MUTE}; }}

.kpis {{ display:grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap:10px; margin-bottom:12px; }}
.kpi {{ background:{CARD}; border:1px solid {LINE}; border-radius:14px; padding:14px 16px; }}
.kpi .v {{ font-size:1.9rem; font-weight:800; letter-spacing:-0.02em; line-height:1.1; }}
.kpi .l {{ font-size:.8rem; color:{MUTE}; margin-top:4px; }}
.kpi.done {{ border-left:5px solid {DONE}; }}
.kpi.left {{ border-left:5px solid {LEFT}; }}

.total {{ background:{INK}; color:#fff; border-radius:14px; padding:16px 18px; margin-bottom:18px; }}
.total .row {{ display:flex; justify-content:space-between; font-weight:600; font-size:.95rem; margin-bottom:10px; gap:10px; flex-wrap:wrap; }}
.bar {{ height:12px; border-radius:99px; background:rgba(255,255,255,.18); overflow:hidden; }}
.bar > i {{ display:block; height:100%; background:{DONE}; border-radius:99px; }}
.total .bar {{ background:rgba(255,255,255,.18); }}

.day {{ background:{CARD}; border:1px solid {LINE}; border-radius:12px; padding:12px 14px; margin-bottom:8px;
        display:grid; grid-template-columns: 92px 1fr 150px; gap:14px; align-items:center; }}
.day .d {{ font-weight:800; font-size:1.05rem; line-height:1.15; }}
.day .d small {{ display:block; font-weight:500; color:{MUTE}; font-size:.78rem; }}
.day .bar {{ background:#E7EDF1; height:10px; }}
.day .meta {{ font-size:.82rem; color:{MUTE}; margin-top:6px; }}
.day .meta b {{ color:{INK}; }}
.day .pct {{ text-align:right; font-weight:800; font-size:1.25rem; }}
.day .pct small {{ display:block; font-weight:500; font-size:.76rem; color:{MUTE}; }}
.day.full {{ border-color:{DONE}; }}
@media (max-width: 640px) {{
  .day {{ grid-template-columns: 1fr auto; }}
  .day .mid {{ grid-column: 1 / -1; order: 3; }}
}}

.badge {{ display:inline-block; padding:3px 10px; border-radius:99px; font-size:.78rem; font-weight:700; }}
.badge.ok {{ background:#D6F0EF; color:#09595A; }}
.badge.no {{ background:#FBEBCF; color:#7A4E05; }}
.hit {{ background:{CARD}; border:1px solid {LINE}; border-radius:12px; padding:12px 14px; margin-bottom:8px; }}
.hit .n {{ font-weight:800; font-size:1.05rem; }}
.hit .s {{ color:{MUTE}; font-size:.85rem; margin:2px 0 8px; }}

.stTabs [data-baseweb="tab-list"] {{ gap:6px; }}
.stTabs [data-baseweb="tab"] {{ background:{CARD}; border:1px solid {LINE}; border-radius:10px; padding:8px 16px; font-weight:700; }}
.stTabs [aria-selected="true"] {{ background:{INK}; color:#fff; border-color:{INK}; }}
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] {{ display:none; }}
div[data-testid="stMetric"] {{ background:{CARD}; border:1px solid {LINE}; border-radius:12px; padding:10px 14px; }}
</style>
""", unsafe_allow_html=True)


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

    c_order, c_shop = find("№"), find("название")
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


def day_summary(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("Дата").agg(
        Заказов=("Заказ", "count"),
        Собрано=("Собран", "sum"),
        Штук=("Штук", "sum"),
    )
    g["Штук собрано"] = df[df["Собран"]].groupby("Дата")["Штук"].sum().reindex(g.index).fillna(0).astype(int)
    g["Осталось"] = g["Заказов"] - g["Собрано"]
    g["Процент"] = (g["Собрано"] / g["Заказов"] * 100).round().astype(int)
    return g.sort_index()


# ───────────────────────── Шапка ─────────────────────────
try:
    data = load()
except Exception as e:
    st.error("Не удалось прочитать таблицу. Проверьте, что таблица открыта для почты "
             "сервисного аккаунта (роль «Читатель») и ключ добавлен в Secrets.")
    st.caption(str(e))
    st.stop()

now = datetime.now(TZ)
col_l, col_r = st.columns([5, 1])
with col_l:
    st.markdown(f"""<div class="top"><h1>📦 Сборка заказов</h1>
    <span class="upd">Обновлено {now:%H:%M}, данные кэшируются на {CACHE_SECONDS // 60} мин</span></div>""",
                unsafe_allow_html=True)
with col_r:
    if st.button("🔄 Обновить", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

cities = sorted(c for c in data["Город"].unique() if c)
with st.expander("Фильтр по городу"):
    chosen = st.multiselect("Город", cities, placeholder="Все города", label_visibility="collapsed")
df = data[data["Город"].isin(chosen)] if chosen else data

if df.empty:
    st.info("Нет заказов по выбранному фильтру.")
    st.stop()

summary = day_summary(df)

# ───────────────────────── Вкладки ─────────────────────────
tab_sum, tab_day, tab_search = st.tabs(["Сводка по дням", "Детали по дате", "Поиск заказа"])

# ---- 1. Сводка ----
with tab_sum:
    total, done = len(df), int(df["Собран"].sum())
    left = total - done
    pct = round(done / total * 100)
    pcs_left = int(df.loc[~df["Собран"], "Штук"].sum())

    st.markdown(f"""
    <div class="kpis">
      <div class="kpi"><div class="v">{total}</div><div class="l">всего заказов</div></div>
      <div class="kpi done"><div class="v">{done}</div><div class="l">собрано</div></div>
      <div class="kpi left"><div class="v">{left}</div><div class="l">осталось собрать</div></div>
      <div class="kpi left"><div class="v">{pcs_left:,}</div><div class="l">штук осталось</div></div>
    </div>
    <div class="total">
      <div class="row"><span>Общий прогресс сборки</span><span>{done} из {total} · {pct}%</span></div>
      <div class="bar"><i style="width:{pct}%"></i></div>
    </div>""".replace(",", " "), unsafe_allow_html=True)

    fig = go.Figure()
    xs = [f"{fmt_date(d)} {WEEKDAYS[d.weekday()]}" for d in summary.index]
    fig.add_bar(x=xs, y=summary["Собрано"], name="Собрано", marker_color=DONE)
    fig.add_bar(x=xs, y=summary["Осталось"], name="Осталось", marker_color=LEFT)
    fig.update_layout(barmode="stack", height=280, margin=dict(l=0, r=0, t=10, b=0),
                      paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      font=dict(family="Manrope", color=INK),
                      legend=dict(orientation="h", y=1.15, x=0),
                      xaxis=dict(type="category"), yaxis=dict(gridcolor=LINE, title=None))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    for d, r in summary.iterrows():
        full = "full" if r["Осталось"] == 0 else ""
        st.markdown(f"""
        <div class="day {full}">
          <div class="d">{fmt_date(d)}<small>{WEEKDAYS[d.weekday()]}, {d:%Y}</small></div>
          <div class="mid">
            <div class="bar"><i style="width:{r['Процент']}%"></i></div>
            <div class="meta">собрано <b>{r['Собрано']}</b> · осталось <b>{r['Осталось']}</b> · всего {r['Заказов']} заказов · {r['Штук']:,} шт</div>
          </div>
          <div class="pct">{r['Процент']}%<small>собрано</small></div>
        </div>""".replace(",", " "), unsafe_allow_html=True)

# ---- 2. По дате ----
with tab_day:
    dates = list(summary.index)
    today = now.date()
    default = next((i for i, d in enumerate(dates) if d >= today), len(dates) - 1)
    d = st.selectbox("Дата отгрузки", dates, index=default,
                     format_func=lambda x: f"{fmt_date(x)}, {WEEKDAYS[x.weekday()]}")
    sub = df[df["Дата"] == d]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Заказов", len(sub))
    c2.metric("Собрано", int(sub["Собран"].sum()))
    c3.metric("Осталось", int((~sub["Собран"]).sum()))
    c4.metric("Штук осталось", f"{int(sub.loc[~sub['Собран'], 'Штук'].sum()):,}".replace(",", " "))

    mode = st.radio("Показать", ["Все", "Не собраны", "Собраны"], horizontal=True)
    if mode == "Не собраны":
        sub = sub[~sub["Собран"]]
    elif mode == "Собраны":
        sub = sub[sub["Собран"]]

    show = sub.assign(Статус=sub["Собран"].map({True: "✅ собран", False: "⏳ осталось"}))
    st.dataframe(show[["Заказ", "Магазин", "Город", "Штук", "Статус", "Статус WMS"]],
                 hide_index=True, use_container_width=True)

    wms = sub["Статус WMS"].replace("", "—").value_counts()
    if len(wms):
        st.caption("Статусы WMS на выбранную дату")
        st.bar_chart(wms, color=DONE, height=200)

# ---- 3. Поиск ----
with tab_search:
    q = st.text_input("Номер заказа", placeholder="например, 1M-0024629 или 0024629").strip()
    if q:
        # поиск по всем данным, без фильтра по городу
        hits = data[data["Заказ"].str.contains(q, case=False, regex=False)]
        if hits.empty:
            st.warning(f"Заказ «{q}» не найден. Проверьте номер: можно вводить только последние цифры.")
        else:
            st.caption(f"Найдено: {len(hits)}")
            for _, r in hits.head(30).iterrows():
                badge = ('<span class="badge ok">собран</span>' if r["Собран"]
                         else '<span class="badge no">ещё не собран</span>')
                st.markdown(f"""
                <div class="hit">
                  <div class="n">{r['Заказ']} &nbsp;{badge}</div>
                  <div class="s">{r['Магазин']} · {r['Город']}</div>
                  Отгрузка: <b>{fmt_date(r['Дата'])}</b> · {r['Штук']} шт · WMS: <b>{r['Статус WMS'] or '—'}</b>
                </div>""", unsafe_allow_html=True)
    else:
        st.info("Введите номер заказа или его часть.")
