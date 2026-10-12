"""
Дашборд сборки заказов: Google Таблица -> Streamlit -> Telegram WebApp.
Один файл. Запуск локально:  streamlit run app.py
"""
from datetime import datetime, date
from io import BytesIO
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
.total .big {{ font-size:2.2rem; font-weight:800; letter-spacing:-0.02em; line-height:1.1; }}
.total .big span {{ font-size:1rem; font-weight:600; opacity:.8; margin-left:6px; }}
.total .hint {{ margin-top:8px; font-size:.9rem; line-height:1.6; opacity:.92; }}

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



def backlog_excel(frame: pd.DataFrame) -> bytes:
    """Excel: все несобранные заказы, сгруппированные по дате: ДАТА и КОЛ-ВО (шт)."""
    b = (frame[~frame["Собран"]].groupby("Дата")["Штук"].sum().reset_index()
         .rename(columns={"Дата": "ДАТА", "Штук": "КОЛ-ВО"}).sort_values("ДАТА"))
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        b.to_excel(xw, sheet_name="Бэклог", index=False)
        ws = xw.sheets["Бэклог"]
        n = len(b) + 1
        ws.cell(row=n + 1, column=1, value="ИТОГО")
        ws.cell(row=n + 1, column=2, value=f"=SUM(B2:B{n})")
        from openpyxl.styles import Font
        for c in ws[1] + ws[n + 1]:
            c.font = Font(bold=True)
        for r in range(2, n + 1):
            ws.cell(row=r, column=1).number_format = "DD.MM.YYYY"
            ws.cell(row=r, column=2).number_format = "#,##0"
        ws.cell(row=n + 1, column=2).number_format = "#,##0"
        ws.column_dimensions["A"].width = 16
        ws.column_dimensions["B"].width = 14
    return out.getvalue()

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
    if st.button("🔄 Обновить", width="stretch"):
        st.cache_data.clear()
        st.rerun()

df = data

if df.empty:
    st.info("Нет заказов по выбранному фильтру.")
    st.stop()

summary = day_summary(df)

# ───────────────────────── Вкладки ─────────────────────────
PAGES = ["Сводка по дням", "Заказы", "Поиск заказа"]
_dates = list(summary.index)
if st.session_state.get("f_date") not in [None] + _dates:
    st.session_state.pop("f_date", None)
if "f_date" not in st.session_state:
    st.session_state.f_date = next((x for x in _dates if x >= now.date()), _dates[-1])
st.session_state.setdefault("f_mode", "Все")
st.session_state.setdefault("page", PAGES[0])


def open_backlog():
    """Клик по бэклогу: переходим в «Заказы», все даты, только несобранные."""
    st.session_state.page = PAGES[1]
    st.session_state.f_date = None
    st.session_state.f_mode = "Не собраны"


page = st.segmented_control("Раздел", PAGES, key="page", label_visibility="collapsed") or PAGES[0]

# ---- 1. Сводка ----
if page == PAGES[0]:
    total, done = len(df), int(df["Собран"].sum())
    left = total - done
    pcs_total = int(df["Штук"].sum())
    pcs_left = int(df.loc[~df["Собран"], "Штук"].sum())
    pct_pcs = round((pcs_total - pcs_left) / pcs_total * 100) if pcs_total else 0

    pcs_done = pcs_total - pcs_left
    hint = (f"Собрано {pcs_done:,} шт · осталось {pcs_left:,} шт · "
            f"не собрано {left} из {total} заказов · всего {pcs_total:,} шт").replace(",", " ")
    st.markdown(f"""
    <div class="kpis">
      <div class="kpi"><div class="v">{total}</div><div class="l">всего заказов</div></div>
      <div class="kpi done"><div class="v">{done}</div><div class="l">собрано заказов</div></div>
      <div class="kpi left"><div class="v">{left}</div><div class="l">осталось заказов</div></div>
    </div>
    <div class="total" title="{hint}">
      <div class="big">{pcs_left:,}<span>шт в бэклоге</span></div>
      <div class="hint">собрано <b>{pcs_done:,} шт</b> · осталось <b>{pcs_left:,} шт</b><br>
      не собрано <b>{left} из {total} заказов</b> · всего {pcs_total:,} шт</div>
    </div>""".replace(",", " "), unsafe_allow_html=True)
    st.button(f"Открыть заказы бэклога ({left}) →", on_click=open_backlog,
              type="primary", width="stretch", help=hint)

    st.download_button(
        "⬇️ Выгрузить бэклог в Excel (дата и кол-во)",
        data=backlog_excel(df),
        file_name=f"backlog_{now:%Y-%m-%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )

    show_all = st.toggle("Показать все дни (по умолчанию только дни с готовностью ниже 95%)", value=False)
    shown = summary if show_all else summary[summary["Процент"] < 95]
    hidden = len(summary) - len(shown)
    if hidden:
        st.caption(f"Скрыто дней с готовностью 95% и выше: {hidden}")
    if shown.empty:
        st.success("Все дни собраны на 95% и выше.")

    fig = go.Figure()
    xs = [f"{fmt_date(d)} {WEEKDAYS[d.weekday()]}" for d in shown.index]
    fig.add_bar(x=xs, y=shown["Собрано"], name="Собрано", marker_color=DONE)
    fig.add_bar(x=xs, y=shown["Осталось"], name="Осталось", marker_color=LEFT)
    fig.update_layout(barmode="stack", height=280, margin=dict(l=0, r=0, t=10, b=0),
                      paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      font=dict(family="Manrope", color=INK),
                      legend=dict(orientation="h", y=1.15, x=0),
                      xaxis=dict(type="category"), yaxis=dict(gridcolor=LINE, title=None))
    st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})

    for d, r in shown.iterrows():
        full = "full" if r["Осталось"] == 0 else ""
        st.markdown(f"""
        <div class="day {full}">
          <div class="d">{fmt_date(d)}<small>{WEEKDAYS[d.weekday()]}, {d:%Y}</small></div>
          <div class="mid">
            <div class="bar"><i style="width:{r['Процент']}%"></i></div>
            <div class="meta">собрано <b>{r['Собрано']}</b> зак. ({r['Штук собрано']:,} шт) · осталось <b>{r['Осталось']}</b> зак. ({r['Штук'] - r['Штук собрано']:,} шт) · всего {r['Заказов']} зак., {r['Штук']:,} шт</div>
          </div>
          <div class="pct">{r['Процент']}%<small>собрано</small></div>
        </div>""".replace(",", " "), unsafe_allow_html=True)

# ---- 2. По дате ----
if page == PAGES[1]:
    dates = list(summary.index)
    opts = [None] + dates
    d = st.selectbox("Дата отгрузки", opts, key="f_date",
                     format_func=lambda x: "Все даты" if x is None
                     else f"{fmt_date(x)}, {WEEKDAYS[x.weekday()]}")
    sub = df if d is None else df[df["Дата"] == d]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Заказов", len(sub))
    c2.metric("Собрано", int(sub["Собран"].sum()))
    c3.metric("Осталось", int((~sub["Собран"]).sum()))
    c4.metric("Штук осталось", f"{int(sub.loc[~sub['Собран'], 'Штук'].sum()):,}".replace(",", " "))

    st.caption(
        f"Собрано {int(sub.loc[sub['Собран'], 'Штук'].sum()):,} шт · "
        f"осталось {int(sub.loc[~sub['Собран'], 'Штук'].sum()):,} шт · "
        f"заказов {len(sub)} · всего {int(sub['Штук'].sum()):,} шт".replace(",", " ")
    )
    mode = st.radio("Показать", ["Все", "Не собраны", "Собраны"], horizontal=True, key="f_mode")
    if mode == "Не собраны":
        sub = sub[~sub["Собран"]]
    elif mode == "Собраны":
        sub = sub[sub["Собран"]]

    show = sub.sort_values(["Дата", "Заказ"]).assign(
        Статус=lambda x: x["Собран"].map({True: "✅ собран", False: "⏳ осталось"}))
    st.dataframe(show[["Дата", "Заказ", "Магазин", "Город", "Штук", "Статус", "Статус WMS"]],
                 hide_index=True, width="stretch",
                 column_config={"Дата": st.column_config.DateColumn(format="DD.MM.YYYY")})

    wms = sub["Статус WMS"].replace("", "—").value_counts()
    if len(wms):
        st.caption("Статусы WMS на выбранную дату")
        st.bar_chart(wms, color=DONE, height=200)

# ---- 3. Поиск ----
if page == PAGES[2]:
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
