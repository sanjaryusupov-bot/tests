# ───────────────────────── Остаток позиции по Magnit Stock ─────────────────────────
LIST_HEADERS = {"внешний id позиции", "товар", "sku wms id", "name", "sku", "название"}


def _mpick(cols: dict, *keys):
    """Первая найденная колонка из списка вариантов названий."""
    for k in keys:
        if k in cols:
            return cols[k]
    return None


def _key(x) -> str:
    """Ключ для сравнения значений: без регистра, лишних пробелов и «.0» на конце."""
    s = str(x).strip().lower().replace("ё", "е")
    s = re.sub(r"\.0$", "", s)
    return re.sub(r"\s+", " ", s)


def magnit_norm(raw: pd.DataFrame) -> pd.DataFrame:
    """Magnit Stock -> sku / ext / wms / name / zone / cell / inv / loc / res / avail / kind (индекс как в raw)."""
    cols = {_nk(c): c for c in raw.columns}
    c_ext = _mpick(cols, "внешнийидпозиции", "внешнийид", "externalid", "externalitemid")
    c_wms = _mpick(cols, "skuwmsid", "skuwms")
    c_name = _mpick(cols, "name", "товар", "наименование")
    c_zone = _mpick(cols, "zoneprefix", "зона")
    c_loc = _mpick(cols, "amountinlocation")
    c_res = _mpick(cols, "reserved", "резерв")
    c_av = _mpick(cols, "amountavailable")
    c_cell = _mpick(cols, "cellbarcode", "ячейка")
    c_inv = _mpick(cols, "originalinvoicenumber", "накладная")
    miss = [t for t, c in (("Внешний ID позиции / Sku Wms ID", c_ext or c_wms), ("name", c_name),
                           ("zone_prefix", c_zone), ("Amount In Location", c_loc)) if c is None]
    if miss:
        raise ValueError("В Magnit Stock нет колонок: " + ", ".join(miss) + ". Есть: " + ", ".join(map(str, raw.columns)))

    def txt(c):
        if c is None:
            return pd.Series("", index=raw.index)
        return raw[c].fillna("").astype(str).str.strip().str.replace(r"\.0$", "", regex=True)

    ext, wms = txt(c_ext), txt(c_wms)
    loc = _num(raw[c_loc])
    avail = _num(raw[c_av]) if c_av else loc
    res = _num(raw[c_res]) if c_res else (loc - avail)
    d = pd.DataFrame({"sku": ext.where(ext != "", wms), "ext": ext, "wms": wms, "name": txt(c_name),
                      "zone": txt(c_zone), "cell": txt(c_cell), "inv": txt(c_inv),
                      "loc": loc, "res": res, "avail": avail}, index=raw.index)
    zk = d["zone"].map(zkey)
    d["kind"] = np.where(zk.isin(RES_KEYS), "yes", np.where(zk.isin(NORES_KEYS), "no", "?"))
    return d[d["sku"] != ""]


def magnit_find(m: pd.DataFrame, items: list):
    """Ищет позиции по Внешнему ID, Sku Wms ID или названию. Возвращает (найденные строки, не найденные значения)."""
    kext, kwms, kname = m["ext"].map(_key), m["wms"].map(_key), m["name"].map(_key)
    total = pd.Series(False, index=m.index)
    missing = []
    for it in items:
        k = _key(it)
        if not k:
            continue
        mask = (kext == k) | (kwms == k) | (kname == k)
        if not mask.any() and len(k) >= 4:                      # запасной вариант: часть названия
            mask = kname.str.contains(k, regex=False)
        if mask.any():
            total |= mask
        else:
            missing.append(it)
    return m[total], missing


def magnit_rep(sub: pd.DataFrame, raw_all: pd.DataFrame, missing: list) -> dict:
    """Собирает отчёт в том же формате, что использует stock_now_tables / stock_now_excel."""
    g = sub.copy()
    zk = g["zone"].map(zkey)
    g["comment"] = g["kind"].map(KIND_TEXT)
    g["o"] = zk.map(ZONE_ORDER).fillna(99)
    g = g.sort_values(["sku", "o", "zone", "cell"]).reset_index(drop=True)
    return dict(date=datetime.now(TZ).date(), rows=g,
                raw=raw_all.loc[sub.index].reset_index(drop=True), missing=missing)


def read_item_list(up) -> list:
    """Список позиций из колонки A загруженного файла (xlsx или csv). Строка-заголовок пропускается."""
    if up.name.lower().endswith(".csv"):
        d = pd.read_csv(up, header=None, dtype=str, sep=None, engine="python")
    else:
        d = pd.read_excel(up, header=None, dtype=str)
    out = [str(x).strip() for x in d.iloc[:, 0].dropna()]
    out = [x for x in out if x]
    if out and _key(out[0]) in LIST_HEADERS:
        out = out[1:]
    return list(dict.fromkeys(out))
