# -*- coding: utf-8 -*-
"""Araç Bakım Takip - tamamen çevrimdışı çalışan masaüstü programı.
Çalıştırınca tarayıcıda açılır. Veriler yanındaki bakim_verileri.db dosyasında durur."""
import os, re, sys, io, socket, sqlite3, threading, webbrowser, datetime as dt
from flask import Flask, request, redirect, render_template, send_file, abort
from jinja2 import DictLoader
import openpyxl

BASE = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(BASE, "bakim_verileri.db")
LOGO = "data:image/png;base64,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"
FAVICON = "data:image/png;base64,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"

SCHEMA = """
CREATE TABLE IF NOT EXISTS km(id INTEGER PRIMARY KEY, date TEXT, km INTEGER, note TEXT);
CREATE TABLE IF NOT EXISTS kontrol(id INTEGER PRIMARY KEY, date TEXT, kind TEXT, result TEXT, note TEXT);
CREATE TABLE IF NOT EXISTS bakim(id INTEGER PRIMARY KEY, name TEXT, grup TEXT, interval_km INTEGER, interval_ay INTEGER, last_date TEXT, last_km INTEGER, note TEXT);
CREATE TABLE IF NOT EXISTS parca(id INTEGER PRIMARY KEY, part TEXT, action TEXT, reason TEXT, date TEXT, brand_code TEXT, mechanic TEXT, cost REAL, note TEXT);
CREATE TABLE IF NOT EXISTS harcama(id INTEGER PRIMARY KEY, date TEXT, category TEXT, descr TEXT, amount REAL, km INTEGER);
CREATE TABLE IF NOT EXISTS evrak(id INTEGER PRIMARY KEY, kind TEXT, expiry TEXT, note TEXT);
CREATE TABLE IF NOT EXISTS rehber(id INTEGER PRIMARY KEY, kind TEXT, name TEXT, phone TEXT, note TEXT);
CREATE TABLE IF NOT EXISTS arac(id INTEGER PRIMARY KEY, name TEXT, plate TEXT, year INTEGER, status TEXT DEFAULT 'Aktif', note TEXT);
CREATE TABLE IF NOT EXISTS ayar(key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS silinen(id INTEGER PRIMARY KEY, uid TEXT, tbl TEXT, upd TEXT);
"""
SCOPED = ["km", "kontrol", "bakim", "parca", "harcama", "evrak"]   # araca bağlı tablolar
SYNC_TABLES = SCOPED + ["rehber", "arac"]                           # senkron için kimlik/zaman damgası taşıyanlar
NOW_SQL = "strftime('%Y-%m-%dT%H:%M:%fZ','now')"
DEFAULT_BAKIM = [("Motor yağı ve filtresi", "Rutin", 10000, 12), ("Hava filtresi", "Rutin", 10000, 12),
                 ("Polen filtresi", "Rutin", 10000, 12), ("Mazot / yakıt filtresi", "Rutin", 20000, None),
                 ("Triger seti, devirdaim, V-kayışı ve rulmanlar", "Ağır", 80000, 48), ("Şanzıman yağı", "Ağır", 100000, 60),
                 ("Soğutma sıvısı (antifriz)", "Ağır", None, 36), ("Fren hidroliği", "Ağır", None, 24)]

KINDS = ["Motor yağı seviyesi", "Soğutma sıvısı seviyesi", "Turbo akustik", "Egzoz emisyon", "Yakıt tüketimi", "Lastik basıncı"]
PERIOD = {"Motor yağı seviyesi": 15, "Soğutma sıvısı seviyesi": 15, "Turbo akustik": 7,
          "Egzoz emisyon": 15, "Yakıt tüketimi": 15, "Lastik basıncı": 15}
CATS = ["Bakım", "Yedek parça", "İşçilik", "Yakıt", "Sigorta", "Muayene", "Vergi", "Lastik", "Diğer"]

# (alan, etiket, tür, seçenekler)  tür: text, area, date, int, num, select
TABLES = {
    "km": dict(title="KM Takibi", order="date DESC, id DESC", cols=[
        ("date", "Tarih", "date", None), ("km", "Güncel KM", "int", None), ("note", "Not", "text", None)]),
    "kontrol": dict(title="Kondisyon Kontrolleri", order="date DESC, id DESC", cols=[
        ("date", "Tarih", "date", None), ("kind", "Kontrol", "select", KINDS),
        ("result", "Durum / Sonuç", "text", None), ("note", "Not", "area", None)]),
    "bakim": dict(title="Periyodik ve Ağır Bakım", order="grup, name", cols=[
        ("name", "Kalem", "text", None), ("grup", "Grup", "select", ["Rutin", "Ağır"]),
        ("interval_km", "Periyot (km)", "int", None), ("interval_ay", "Periyot (ay)", "int", None),
        ("last_date", "Son işlem tarihi", "date", None), ("last_km", "Son işlem km", "int", None),
        ("note", "Not", "area", None)]),
    "parca": dict(title="Parça Değişim / Onarım", order="date DESC, id DESC", cols=[
        ("part", "Parça", "text", None), ("action", "İşlem", "text", None), ("reason", "Sebebi", "text", None),
        ("date", "Tarih", "date", None), ("brand_code", "Marka / Parça kodu", "text", None),
        ("mechanic", "Usta / Servis", "text", None), ("cost", "Tutar (₺)", "num", None), ("note", "Not", "area", None)]),
    "harcama": dict(title="Harcamalar", order="date DESC, id DESC", cols=[
        ("date", "Tarih", "date", None), ("category", "Kategori", "select", CATS),
        ("descr", "Açıklama", "text", None), ("amount", "Tutar (₺)", "num", None), ("km", "KM", "int", None)]),
    "evrak": dict(title="Evrak ve Muayene", order="expiry", cols=[
        ("kind", "Evrak", "select", ["Muayene (TÜVTÜRK)", "Zorunlu Trafik Sigortası", "Kasko", "Diğer"]),
        ("expiry", "Bitiş tarihi", "date", None), ("note", "Not", "text", None)]),
    "arac": dict(title="Araçlar", order="name", cols=[
        ("name", "Araç adı / model", "text", None), ("plate", "Plaka", "text", None), ("year", "Model yılı", "int", None),
        ("status", "Durum", "select", ["Aktif", "Satıldı"]), ("note", "Not", "area", None)]),
    "rehber": dict(title="Usta ve Parça Rehberi", order="kind, name", cols=[
        ("kind", "Tür", "select", ["Usta / Servis", "Parça"]), ("name", "Ad / Marka", "text", None),
        ("phone", "Telefon / Parça kodu", "text", None), ("note", "Not", "area", None)]),
}

# ----------------------------------------------------------------- yardımcılar
def pdate(s):
    try:
        return dt.date.fromisoformat(s) if s else None
    except ValueError:
        return None

def fdate(s):
    d = pdate(s)
    return d.strftime("%d.%m.%Y") if d else ""

def _tr(s):
    return s.replace(",", "X").replace(".", ",").replace("X", ".")

def fint(v):
    return "" if v in (None, "") else _tr(f"{int(round(v)):,}")

def fmoney(v):
    return "" if v in (None, "") else _tr(f"{float(v):,.2f}")

def parse_num(s):
    s = s.replace(" ", "").replace("₺", "").replace("TL", "")
    try:
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        elif re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
            s = s.replace(".", "")
        return float(s)
    except ValueError:
        return None

def parse_val(col, raw):
    raw = (raw or "").strip()
    if raw == "":
        return None
    if col[2] == "int":
        n = parse_num(raw)
        return int(round(n)) if n is not None else None
    if col[2] == "num":
        return parse_num(raw)
    return raw

def add_months(d, n):
    m = d.month - 1 + n
    y, m = d.year + m // 12, m % 12 + 1
    day = min(d.day, [31, 29 if y % 4 == 0 and (y % 100 or y % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1])
    return dt.date(y, m, day)

def get_db():
    db = sqlite3.connect(DB)
    db.row_factory = sqlite3.Row
    return db

def seed_arac(db, aid, bakim=True):
    """Yeni araç için evrak satırlarını (ve istenirse standart bakım kalemlerini) ekler."""
    if not db.execute("SELECT 1 FROM evrak WHERE arac_id=?", (aid,)).fetchone():
        for k in ("Muayene (TÜVTÜRK)", "Zorunlu Trafik Sigortası", "Kasko"):
            db.execute("INSERT INTO evrak(arac_id, kind, expiry, note) VALUES(?,?,?,?)", (aid, k, None, "Bitiş tarihini girin"))
    if bakim and not db.execute("SELECT 1 FROM bakim WHERE arac_id=?", (aid,)).fetchone():
        for n, g, ikm, iay in DEFAULT_BAKIM:
            db.execute("INSERT INTO bakim(arac_id, name, grup, interval_km, interval_ay) VALUES(?,?,?,?,?)", (aid, n, g, ikm, iay))

def migrate(db):
    db.executescript(SCHEMA)
    cols = lambda t: [r[1] for r in db.execute(f"PRAGMA table_info({t})")]
    for t in SYNC_TABLES:
        for c in ("uid", "upd"):
            if c not in cols(t):
                db.execute(f"ALTER TABLE {t} ADD COLUMN {c} TEXT")
        db.execute(f"UPDATE {t} SET uid=lower(hex(randomblob(16))), upd={NOW_SQL} WHERE uid IS NULL")
    for t in SCOPED:
        if "arac_id" not in cols(t):
            db.execute(f"ALTER TABLE {t} ADD COLUMN arac_id INTEGER")
    if not db.execute("SELECT 1 FROM arac").fetchone():
        db.execute("INSERT INTO arac(name, status) VALUES('Aracım', 'Aktif')")
    first = db.execute("SELECT id FROM arac ORDER BY id").fetchone()[0]
    for t in SCOPED:
        db.execute(f"UPDATE {t} SET arac_id=? WHERE arac_id IS NULL", (first,))
    seed_arac(db, first, bakim=not db.execute("SELECT 1 FROM bakim").fetchone())
    for t in SYNC_TABLES:   # kayıt eklenince kimlik, değişince/silinince zaman damgası otomatik
        db.executescript(f"""
        CREATE TRIGGER IF NOT EXISTS {t}_ai AFTER INSERT ON {t} BEGIN
          UPDATE {t} SET uid=COALESCE(uid, lower(hex(randomblob(16)))), upd=COALESCE(upd, {NOW_SQL}) WHERE id=NEW.id; END;
        CREATE TRIGGER IF NOT EXISTS {t}_au AFTER UPDATE ON {t} WHEN NEW.upd IS OLD.upd BEGIN
          UPDATE {t} SET upd={NOW_SQL} WHERE id=NEW.id; END;
        CREATE TRIGGER IF NOT EXISTS {t}_ad AFTER DELETE ON {t} BEGIN
          INSERT INTO silinen(uid, tbl, upd) VALUES(OLD.uid, '{t}', {NOW_SQL}); END;""")

def init_db():
    db = get_db()
    migrate(db)
    db.commit()
    db.close()

def cur_arac(db):
    r = db.execute("SELECT value FROM ayar WHERE key='aktif_arac'").fetchone()
    a = db.execute("SELECT * FROM arac WHERE id=?", (r[0],)).fetchone() if r else None
    return a or db.execute("SELECT * FROM arac ORDER BY (status='Aktif') DESC, id LIMIT 1").fetchone()

# ----------------------------------------------------------------- hesaplar
def km_stats(db, aid):
    rows = db.execute("SELECT date, km FROM km WHERE arac_id=? AND date IS NOT NULL AND km IS NOT NULL ORDER BY date, km", (aid,)).fetchall()
    if not rows:
        return None, None
    last_d = pdate(rows[-1]["date"])
    base = [r for r in rows if (last_d - pdate(r["date"])).days <= 120]
    if len(base) < 2:
        base = rows
    days = (pdate(base[-1]["date"]) - pdate(base[0]["date"])).days
    perday = (base[-1]["km"] - base[0]["km"]) / days if days > 0 else None
    return rows[-1]["km"], perday

LV = {"bad": 0, "warn": 1, "unk": 2, "ok": 3}

def bakim_status(it, cur_km, perday):
    """Bir bakım kaleminin durumunu ve açıklamasını döndürür."""
    level, parts = "ok", []
    def worse(l):
        nonlocal level
        if LV[l] < LV[level]:
            level = l
    known = False
    if it["interval_km"] and it["last_km"] is not None and cur_km is not None:
        known = True
        nxt = it["last_km"] + it["interval_km"]
        rem = nxt - cur_km
        if rem <= 0:
            worse("bad"); parts.append(f"{fint(-rem)} km geçti (hedef {fint(nxt)} km)")
        else:
            if rem <= 1000:
                worse("warn")
            txt = f"{fint(rem)} km kaldı (hedef {fint(nxt)} km)"
            if perday and perday > 0:
                txt += f", yaklaşık {int(rem / perday)} gün"
            parts.append(txt)
    if it["interval_ay"] and pdate(it["last_date"]):
        known = True
        nd = add_months(pdate(it["last_date"]), it["interval_ay"])
        days = (nd - dt.date.today()).days
        if days < 0:
            worse("bad"); parts.append(f"tarih {-days} gün geçti ({nd.strftime('%d.%m.%Y')})")
        else:
            if days <= 30:
                worse("warn")
            parts.append(f"{days} gün kaldı ({nd.strftime('%d.%m.%Y')})")
    if not known:
        return "unk", "Son işlem bilinmiyor, tarih ve km girin"
    return level, "; ".join(parts)

def build_alerts(db, aid):
    cur_km, perday = km_stats(db, aid)
    out, today = [], dt.date.today()
    for it in db.execute("SELECT * FROM bakim WHERE arac_id=? ORDER BY grup, name", (aid,)):
        lv, det = bakim_status(it, cur_km, perday)
        out.append(dict(group="Bakım", title=it["name"], detail=det, level=lv, link="/t/bakim"))
    for k in KINDS:
        r = db.execute("SELECT MAX(date) d FROM kontrol WHERE kind=? AND arac_id=?", (k, aid)).fetchone()
        if not r["d"]:
            out.append(dict(group="Kontrol", title=k, detail="Hiç kayıt yok", level="unk", link="/t/kontrol")); continue
        due = pdate(r["d"]) + dt.timedelta(days=PERIOD[k])
        days = (due - today).days
        lv = "bad" if days < 0 else "warn" if days <= 2 else "ok"
        det = f"{-days} gün gecikti (son: {fdate(r['d'])})" if days < 0 else f"{days} gün sonra ({due.strftime('%d.%m.%Y')})"
        out.append(dict(group="Kontrol", title=k, detail=det, level=lv, link="/t/kontrol"))
    for e in db.execute("SELECT * FROM evrak WHERE arac_id=? ORDER BY kind", (aid,)):
        d = pdate(e["expiry"])
        if not d:
            out.append(dict(group="Evrak", title=e["kind"], detail="Bitiş tarihi girilmemiş", level="unk", link="/t/evrak")); continue
        days = (d - today).days
        lv = "bad" if days < 0 else "warn" if days <= 30 else "ok"
        det = f"Süresi {-days} gün önce doldu" if days < 0 else f"Bitişe {days} gün ({d.strftime('%d.%m.%Y')})"
        out.append(dict(group="Evrak", title=e["kind"], detail=det, level=lv, link="/t/evrak"))
    out.sort(key=lambda a: LV[a["level"]])
    return out, cur_km, perday

def harcama_ozet(db, aid):
    yil = dt.date.today().year
    rows = db.execute("SELECT date, category, amount FROM harcama WHERE arac_id=? AND amount IS NOT NULL", (aid,)).fetchall()
    this = [r for r in rows if (r["date"] or "").startswith(str(yil))]
    cat, mon = {}, [0.0] * 12
    for r in this:
        cat[r["category"] or "Diğer"] = cat.get(r["category"] or "Diğer", 0) + r["amount"]
        mon[int(r["date"][5:7]) - 1] += r["amount"]
    return dict(yil=yil, total=sum(r["amount"] for r in this), all_total=sum(r["amount"] for r in rows),
                cats=sorted(cat.items(), key=lambda x: -x[1]), months=mon)


def svg_charts(rows):
    """KM Takibi için iki grafik: KM-zaman çizgisi ve aralıklara göre günlük ortalama."""
    from markupsafe import Markup
    pts = sorted((pdate(r["date"]), r["km"]) for r in rows if pdate(r["date"]) and r["km"] is not None)
    if len(pts) < 2:
        return ""
    W, H, L, R, T, B = 520, 270, 58, 22, 14, 34
    GRID, TXT, ACC = "#e1e6ec", "#69778a", "#1f6feb"
    d0, span = pts[0][0].toordinal(), max(pts[-1][0].toordinal() - pts[0][0].toordinal(), 1)
    kmin, kmax = min(p[1] for p in pts), max(p[1] for p in pts)
    pad = max((kmax - kmin) * 0.08, 50)
    lo, hi = kmin - pad, kmax + pad
    X = lambda d: L + (d.toordinal() - d0) / span * (W - L - R)
    Y = lambda k: T + (1 - (k - lo) / (hi - lo)) * (H - T - B)
    g = []
    for i in range(5):
        k = lo + (hi - lo) * i / 4
        g.append(f'<line x1="{L}" x2="{W - R}" y1="{Y(k):.1f}" y2="{Y(k):.1f}" stroke="{GRID}"/>'
                 f'<text x="{L - 8}" y="{Y(k) + 4:.1f}" text-anchor="end" font-size="12" fill="{TXT}">{fint(k)}</text>')
        d = dt.date.fromordinal(d0 + round(span * i / 4))
        g.append(f'<text x="{X(d):.1f}" y="{H - 12}" text-anchor="{"end" if i == 4 else "start" if i == 0 else "middle"}" font-size="12" fill="{TXT}">{d.strftime("%d.%m.%y")}</text>')
    line = " ".join(f"{X(d):.1f},{Y(k):.1f}" for d, k in pts)
    dots = "".join(f'<circle cx="{X(d):.1f}" cy="{Y(k):.1f}" r="4" fill="{ACC}"><title>{d.strftime("%d.%m.%Y")}: {fint(k)} km</title></circle>' for d, k in pts)
    svg1 = (f'<svg viewBox="0 0 {W} {H}" width="100%">{"".join(g)}<polyline points="{line}" fill="none" stroke="{ACC}" stroke-width="2.5"/>{dots}</svg>')

    iv = []
    for (da, ka), (db_, kb) in zip(pts, pts[1:]):
        days = (db_ - da).days
        if days > 0:
            iv.append((da, db_, (kb - ka) / days, kb - ka, days))
    svg2 = ""
    if iv:
        mx = max(v[2] for v in iv) or 1
        bw = (W - L - R) / len(iv)
        g2 = []
        for i in range(5):
            v = mx * 1.1 * i / 4
            y = T + (1 - i / 4) * (H - T - B)
            g2.append(f'<line x1="{L}" x2="{W - R}" y1="{y:.1f}" y2="{y:.1f}" stroke="{GRID}"/>'
                      f'<text x="{L - 8}" y="{y + 4:.1f}" text-anchor="end" font-size="12" fill="{TXT}">{v:.0f}</text>')
        bars, step = [], max(1, len(iv) // 6)
        for i, (da, db_, pd_, dk, days) in enumerate(iv):
            h = pd_ / (mx * 1.1) * (H - T - B)
            x = L + i * bw + bw * 0.15
            bars.append(f'<rect x="{x:.1f}" y="{H - B - h:.1f}" width="{bw * 0.7:.1f}" height="{h:.1f}" rx="3" fill="{ACC}" opacity="0.85">'
                        f'<title>{da.strftime("%d.%m.%Y")} - {db_.strftime("%d.%m.%Y")}: {fint(dk)} km / {days} gün = {pd_:.0f} km/gün</title></rect>')
            if i % step == 0 or i == len(iv) - 1:
                bars.append(f'<text x="{x + bw * 0.35:.1f}" y="{H - 12}" text-anchor="middle" font-size="11" fill="{TXT}">{db_.strftime("%d.%m")}</text>')
        svg2 = f'<svg viewBox="0 0 {W} {H}" width="100%">{"".join(g2)}{"".join(bars)}</svg>'
    toplam, gun = pts[-1][1] - pts[0][1], (pts[-1][0] - pts[0][0]).days
    ozet = f'<small>İlk kayıttan bu yana <b>{fint(toplam)} km</b>, {gun} günde (ortalama {toplam / gun:.0f} km/gün)</small>' if gun else ""
    return Markup(f'<div class="cols"><div class="card"><h3 style="margin-top:0">KM gelişimi</h3>{svg1}</div>'
                  f'<div class="card"><h3 style="margin-top:0">Günlük ortalama km (kayıtlar arası)</h3>{svg2}</div></div><p>{ozet}</p>')

# ----------------------------------------------------------------- Excel'den içe aktarma
def _d(v):
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    return None

def _t(v):
    s = "" if v is None else str(v).strip()
    return "" if s in ("-", '""') else s.replace('"" ', "").strip()

def _cap(s):
    s = s.strip()
    if not s:
        return s
    first = {"i": "İ", "ı": "I"}.get(s[0], s[0].upper())
    return first + s[1:]

def _clean(s):
    return re.sub(r"^[-\s]+", "", s).strip()

def _psi(v):
    s = _t(v)
    n = _n(s)
    return (f"{n:g}" if n is not None else s) or "-"

def _n(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None

def _fuel(v):  # Excel "5.5" yazısını tarih sanmış olabilir
    if isinstance(v, (dt.datetime, dt.date)):
        return float(f"{v.day}.{v.month}")
    return _n(v)

def _sheet(wb, key):
    for ws in wb.worksheets:
        if key.lower() in ws.title.lower():
            return ws
    return None

def import_legacy(db, path, aid):
    wb = openpyxl.load_workbook(path)
    for t in ("km", "kontrol", "bakim", "parca"):
        db.execute(f"DELETE FROM {t} WHERE arac_id=?", (aid,))
    ws = _sheet(wb, "KM")
    if ws:
        for r in range(3, 400):
            d, k = _d(ws.cell(r, 2).value), _n(ws.cell(r, 4).value)
            if d and k is not None:
                db.execute("INSERT INTO km(date, km) VALUES(?,?)", (d, int(k)))
    ws = _sheet(wb, "Kondisyon")
    if ws:
        c = lambda r, col: ws.cell(r, col).value
        for r in range(4, 400):
            if _d(c(r, 2)):
                lvl = _n(c(r, 4)); ek = _cap(_t(c(r, 5)))
                res = (f"Seviye %{round(lvl * 100)}" if lvl is not None else "") + (f", ekleme: {ek}" if ek else "")
                note = _t(c(r, 6)); db.execute("INSERT INTO kontrol(date,kind,result,note) VALUES(?,?,?,?)",
                    (_d(c(r, 2)), KINDS[0], res, f"Eklenen: {note}" if note else ""))
            if _d(c(r, 8)):
                lvl = _n(c(r, 10)); ek = _cap(_t(c(r, 11)))
                res = (f"Seviye %{round(lvl * 100)}" if lvl is not None else "") + (f", ekleme: {ek}" if ek else "")
                note = _t(c(r, 12)); db.execute("INSERT INTO kontrol(date,kind,result,note) VALUES(?,?,?,?)",
                    (_d(c(r, 8)), KINDS[1], res, f"Eklenen: {note}" if note else ""))
            if _d(c(r, 14)):
                note = " | ".join(x for x in (_clean(_t(c(r, 17))), _clean(_t(c(r, 18)))) if x)
                db.execute("INSERT INTO kontrol(date,kind,result,note) VALUES(?,?,?,?)",
                    (_d(c(r, 14)), KINDS[2], _cap(_t(c(r, 16))), note))
            if _d(c(r, 20)):
                db.execute("INSERT INTO kontrol(date,kind,result,note) VALUES(?,?,?,?)",
                    (_d(c(r, 20)), KINDS[3], _cap(_t(c(r, 22))), _t(c(r, 23))))
            if _d(c(r, 25)):
                litre, km = _fuel(c(r, 28)), _n(c(r, 29))
                if litre and km:
                    res, note = f"{litre / km * 100:.1f} L/100km", f"{litre:g} L / {km:g} km (son sıfırlamadan sonra)"
                else:
                    res, note = "", ""
                db.execute("INSERT INTO kontrol(date,kind,result,note) VALUES(?,?,?,?)", (_d(c(r, 25)), KINDS[4], res, note))
            if _d(c(r, 31)):
                n = lambda col: _psi(c(r, col))
                res = f"Sağ ön {n(33)}, sol ön {n(34)}, sağ arka {n(35)}, sol arka {n(36)}, stepne {n(37)} (psi)"
                db.execute("INSERT INTO kontrol(date,kind,result,note) VALUES(?,?,?,?)", (_d(c(r, 31)), KINDS[5], res, ""))
    ws = _sheet(wb, "Periyodik")
    items = [("Motor yağı ve filtresi", 2, 10000, 12), ("Hava filtresi", 6, 10000, 12),
             ("Polen filtresi", 10, 10000, 12), ("Mazot filtresi", 14, 20000, None)]
    if ws:
        for name, col, ikm, iay in items:
            kms = [int(x) for x in (_n(ws.cell(r, col + 1).value) for r in range(5, 60)) if x is not None]
            note = ""
            if name.startswith("Motor"):
                note = _t(ws.cell(2, 2).value).replace("NOT :", "").strip()
            if len(kms) > 1:
                note = (note + "\n" if note else "") + "Önceki işlemler (km): " + ", ".join(fint(k) for k in kms[:-1])
            db.execute("INSERT INTO bakim(name,grup,interval_km,interval_ay,last_km,note) VALUES(?,?,?,?,?,?)",
                       (name, "Rutin", ikm, iay, max(kms) if kms else None, note))
    for name, ikm, iay, note in [("Triger seti, devirdaim, V-kayışı ve rulmanlar", 80000, 48, ""),
                                 ("Şanzıman yağı", 100000, 60, ""), ("Soğutma sıvısı (antifriz)", None, 36, ""),
                                 ("Fren hidroliği", None, 24, "Nem oranı %3'e ulaşırsa da değişir")]:
        db.execute("INSERT INTO bakim(name,grup,interval_km,interval_ay,note) VALUES(?,?,?,?,?)", (name, "Ağır", ikm, iay, note))
    ws = _sheet(wb, "Par")
    if ws:
        for r in range(3, 200):
            if _t(ws.cell(r, 2).value):
                db.execute("INSERT INTO parca(part,action,reason,date,note) VALUES(?,?,?,?,?)",
                           (_t(ws.cell(r, 2).value), _t(ws.cell(r, 3).value), _t(ws.cell(r, 4).value),
                            _d(ws.cell(r, 5).value), _t(ws.cell(r, 6).value)))
    for t in ("km", "kontrol", "bakim", "parca"):
        db.execute(f"UPDATE {t} SET arac_id=? WHERE arac_id IS NULL", (aid,))
    db.commit()

# ----------------------------------------------------------------- şablonlar
CSS = """
:root{--bg:#f3f5f8;--card:#fff;--ink:#1c2733;--mut:#69778a;--line:#e1e6ec;--acc:#1f6feb;--bad:#cf222e;--warn:#c77700;--ok:#1a7f37;--th:#eef1f5;--inp:#fff;--sec:#e8ecf1}
@media(prefers-color-scheme:dark){:root{--bg:#0f1720;--card:#18232f;--ink:#e6edf3;--mut:#93a1b1;--line:#2a3846;--acc:#4493f8;--bad:#ff7b72;--warn:#e3a008;--ok:#3fb950;--th:#1f2d3b;--inp:#0f1720;--sec:#2a3846}}
*{box-sizing:border-box}body{margin:0;font:15px/1.45 Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--ink);display:flex;min-height:100vh}
nav{width:225px;background:linear-gradient(170deg,#2f80ff,#0b3a8f);color:#dbe7ff;padding:16px 0;flex-shrink:0}
.vsel select{width:calc(100% - 32px);margin:0 16px 4px;background:#ffffff26;color:#fff;border:1px solid #ffffff55}.vsel option{color:#1c2733}.mut{color:var(--mut);font-weight:400;font-size:.7em}.brand{display:flex;align-items:center;gap:11px;margin:0 16px 18px;color:#fff}.brand img{width:46px;height:46px;border-radius:11px;box-shadow:0 2px 8px #0004}.brand b{font-size:16px;line-height:1.15;display:block}.brand small{color:#bcd3ff}nav a{display:block;padding:9px 18px;color:#dbe7ff;text-decoration:none;border-left:4px solid transparent}
nav a:hover{background:#ffffff1f}nav a.on{background:#ffffff33;color:#fff;font-weight:600;border-left-color:#ffb703}
main{flex:1;padding:24px;max-width:1100px}h2{margin:0 0 16px}h3{margin:22px 0 8px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px;margin-bottom:8px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 16px}
.card b{display:block;font-size:22px}.card span{color:var(--mut);font-size:13px}
.row{display:flex;gap:10px;align-items:flex-start;background:var(--card);border:1px solid var(--line);border-left:5px solid var(--line);border-radius:8px;padding:9px 12px;margin-bottom:6px;text-decoration:none;color:inherit}
.row.bad{border-left-color:var(--bad)}.row.warn{border-left-color:var(--warn)}.row.ok{border-left-color:var(--ok)}.row.unk{border-left-color:#9aa6b4}
.row small{color:var(--mut)}.tag{font-size:11px;text-transform:uppercase;color:var(--mut);min-width:54px;padding-top:3px}
.bad .d{color:var(--bad)}.warn .d{color:var(--warn)}
table{width:100%;border-collapse:collapse;background:var(--card);border:1px solid var(--line);border-radius:8px;overflow:hidden}
th,td{padding:8px 10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}th{background:var(--th);font-size:13px}
td.n,th.n{text-align:right;white-space:nowrap}td.w{max-width:320px;white-space:pre-wrap}
a.btn,button{background:var(--acc);color:#fff;border:0;border-radius:6px;padding:7px 14px;font:inherit;cursor:pointer;text-decoration:none;display:inline-block}
button.sec,a.sec{background:var(--sec);color:var(--ink)}button.del{background:transparent;color:var(--bad);border:1px solid var(--line)}
form.inl{display:inline-flex;gap:6px;align-items:center}input,select,textarea{font:inherit;padding:7px 9px;border:1px solid var(--line);border-radius:6px;width:100%;background:var(--inp);color:var(--ink)}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:10px;margin-bottom:10px}label{font-size:13px;color:var(--mut)}
details{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px 14px;margin-bottom:14px}summary{cursor:pointer;font-weight:600}
.bar{display:flex;align-items:center;gap:8px;margin:4px 0}.bar i{display:block;height:14px;background:var(--acc);border-radius:3px;min-width:2px}
.bar em{font-style:normal;min-width:120px}.bar u{text-decoration:none;color:var(--mut);font-size:13px}
.cols{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:16px}
@media(max-width:700px){body{display:block}nav{width:auto}main{padding:14px}}
"""

BASE_T = """<!doctype html><html lang="tr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Araç Bakım Takip</title><link rel="icon" href="{{ favicon }}"><style>""" + CSS + """</style></head><body>
<nav><div class="brand"><img src="{{ logo }}" alt=""><div><b>Araç Bakım</b><small>Takip</small></div></div>
<a href="/" class="{{ 'on' if active=='' }}">Ana Ekran</a>
<div class="vsel"><form method="post" action="/arac/sec"><select name="id" onchange="this.form.submit()" title="Araç seç">{% for a in araclar %}<option value="{{a.id}}"{{ " selected" if a.id==aktif.id }}>{{a.name}}{{ " · "+a.plate if a.plate }}{{ " (satıldı)" if a.status=="Satıldı" }}</option>{% endfor %}</select></form></div>
<a href="/araclar" class="{{ 'on' if active=='araclar' }}">＋ Araçlarım / Araç ekle</a><hr style="border:0;border-top:1px solid #ffffff40;margin:8px 16px">
{% for k,t in tables.items() if k!='arac' %}<a href="/t/{{k}}" class="{{ 'on' if active==k }}">{{t.title}}</a>{% endfor %}
<a href="/yedek" class="{{ 'on' if active=='yedek' }}">Yedekleme / Excel</a></nav>
<main>{% block body %}{% endblock %}</main></body></html>"""

DASH_T = """{% extends 'base.html' %}{% block body %}
<h2>Ana Ekran <span class="mut">· {{ aktif.name }}{{ " · "+aktif.plate if aktif.plate }}</span></h2>
<div class="cards">
<div class="card"><b>{{ fint(km) if km else '-' }} km</b><span>Güncel KM</span></div>
<div class="card"><b>{{ ('%.0f'|format(perday)) if perday else '-' }} km</b><span>Günlük ortalama (son ~4 ay)</span></div>
<div class="card"><b>{{ fmoney(h.total) }} ₺</b><span>{{ h.yil }} yılı harcaması</span></div>
<div class="card"><b>{{ dikkat|length }}</b><span>Dikkat gerektiren kalem</span></div></div>
<h3>Dikkat gerekenler</h3>
{% for a in dikkat %}<a class="row {{a.level}}" href="{{a.link}}"><span class="tag">{{a.group}}</span><div><b>{{a.title}}</b><br><small class="d">{{a.detail}}</small></div></a>
{% else %}<p>Her şey yolunda 👍</p>{% endfor %}
<h3>Takipte (sorun yok)</h3>
{% for a in tamam %}<a class="row ok" href="{{a.link}}"><span class="tag">{{a.group}}</span><div><b>{{a.title}}</b><br><small>{{a.detail}}</small></div></a>
{% else %}<p><small>Henüz yok.</small></p>{% endfor %}{% endblock %}"""

LIST_T = """{% extends 'base.html' %}{% block body %}
<h2>{{ t.title }}{% if name in scoped_tables %} <span class="mut">· {{ aktif.name }}</span>{% endif %}</h2>
{% if name=='harcama' %}
<div class="cards"><div class="card"><b>{{ fmoney(h.total) }} ₺</b><span>{{ h.yil }} yılı toplam</span></div>
<div class="card"><b>{{ fmoney(h.all_total) }} ₺</b><span>Tüm zamanların toplamı</span></div></div>
<div class="cols"><div><h3>{{ h.yil }} - kategoriye göre</h3>
{% for c,v in h.cats %}<div class="bar"><em>{{c}}</em><i style="width:{{ (v/h.cats[0][1]*180)|int }}px"></i><u>{{ fmoney(v) }} ₺</u></div>{% else %}<small>Bu yıl kayıt yok.</small>{% endfor %}</div>
<div><h3>{{ h.yil }} - aylara göre</h3>
{% set mx = h.months|max %}{% for v in h.months %}{% if mx>0 %}<div class="bar"><em>{{ aylar[loop.index0] }}</em><i style="width:{{ (v/mx*180)|int }}px"></i><u>{{ fmoney(v) if v else '' }}</u></div>{% endif %}{% endfor %}</div></div>
{% endif %}
{% if charts %}{{ charts }}{% endif %}
{% if name=='parca' %}<p><small>İpucu: yapılan işlemin tutarını girerseniz, harcamalara ayrıca "Harcamalar" sayfasından eklemeyi unutmayın.</small></p>{% endif %}
<details {{ 'open' if not rows }}><summary>+ Yeni kayıt ekle</summary><br>
<form method="post" action="/t/{{name}}/add"><div class="grid">
{% for c in t.cols %}<div><label>{{c[1]}}</label>{{ field(c, {}, true) }}</div>{% endfor %}</div>
<button>Kaydet</button></form></details>
<div style="overflow-x:auto"><table><tr>{% for c in t.cols %}<th class="{{ 'n' if c[2] in ('int','num') }}">{{c[1]}}</th>{% endfor %}{% if name=='bakim' %}<th>Durum</th>{% endif %}<th></th></tr>
{% for r in rows %}<tr>
{% for c in t.cols %}<td class="{{ 'n' if c[2] in ('int','num') else 'w' if c[2]=='area' }}">{{ show(c, r[c[0]]) }}</td>{% endfor %}
{% if name=='bakim' %}<td class="w"><b style="color:var(--{{ st[r.id][0] if st[r.id][0] in ('bad','warn','ok') else 'mut' }})">{{ st[r.id][1] }}</b>
<form class="inl" method="post" action="/bakim/{{r.id}}/done" onsubmit="return confirm('Bu bakım şimdi yapıldı olarak işaretlensin mi? (Tarih: bugün, KM: güncel km)')"><input name="amount" placeholder="Tutar ₺ (ops.)" style="width:110px"><button class="sec">Yapıldı</button></form></td>{% endif %}
<td style="white-space:nowrap"><a class="btn sec" href="/t/{{name}}/edit/{{r.id}}">Düzenle</a>
<form class="inl" method="post" action="/t/{{name}}/delete/{{r.id}}" onsubmit="return confirm('Bu kayıt silinsin mi?')"><button class="del">Sil</button></form></td></tr>
{% else %}<tr><td colspan="20"><small>Kayıt yok.</small></td></tr>{% endfor %}</table></div>{% endblock %}"""

EDIT_T = """{% extends 'base.html' %}{% block body %}
<h2>{{ t.title }} - Düzenle</h2>
<form method="post"><div class="grid">{% for c in t.cols %}<div><label>{{c[1]}}</label>{{ field(c, row, false) }}</div>{% endfor %}</div>
<button>Kaydet</button> <a class="btn sec" href="{{ '/araclar' if name=='arac' else '/t/'+name }}">Vazgeç</a></form>{% endblock %}"""

ARACLAR_T = """{% extends 'base.html' %}{% block body %}
<h2>Araçlarım</h2>
<p><small>Verileri girmek için önce soldaki listeden aracı seçin. Aracı satarsanız silmeyin, "Satıldı" yapın, geçmişi kalsın.</small></p>
<div class="cols">{% for a in tum %}<div class="card" style="{{ 'border:2px solid var(--acc)' if a.id==aktif.id }}">
<b>{{ a.name }}</b><span>{{ a.plate or '' }} {{ a.year or '' }}</span><br>
<small>{{ a.status }}{% if km[a.id] %} · {{ fint(km[a.id]) }} km{% endif %}{% if a.id==aktif.id %} · <b style="display:inline;font-size:inherit;color:var(--acc)">şu an seçili</b>{% endif %}</small>
<div style="display:flex;gap:6px;flex-wrap:wrap;margin-top:12px">{% if a.id!=aktif.id %}<form class="inl" method="post" action="/arac/sec/{{a.id}}"><button>Bu aracı seç</button></form> {% endif %}
<a class="btn sec" href="/t/arac/edit/{{a.id}}">Düzenle</a>
{% if tum|length>1 %}<form class="inl" method="post" action="/arac/sil/{{a.id}}" onsubmit="return confirm('DİKKAT: {{ a.name|e }} aracı ve bu araca ait TÜM kayıtlar (km, bakım, harcama...) kalıcı olarak silinecek. Aracı satıyorsanız silmeyin, Düzenle > Durum: Satıldı yapın. Yine de silinsin mi?')"><button class="del">Sil</button></form>{% endif %}</div></div>{% endfor %}</div>
<details open style="margin-top:16px"><summary>+ Yeni araç ekle</summary><br>
<form method="post" action="/arac/ekle"><div class="grid">
<div><label>Araç adı / model *</label><input name="name" required placeholder="örn. Passat 1.6 TDI"></div>
<div><label>Plaka</label><input name="plate"></div>
<div><label>Model yılı</label><input name="year" inputmode="numeric"></div>
<div><label>Not</label><input name="note"></div></div>
<label style="display:block;margin:8px 0"><input type="checkbox" name="std" value="1" checked style="width:auto"> Standart bakım kalemlerini ekle (yağ, filtreler, triger, şanzıman yağı, antifriz, fren hidroliği)</label>
<button>Aracı ekle ve seç</button></form></details>{% endblock %}"""

YEDEK_T = """{% extends 'base.html' %}{% block body %}
<h2>Yedekleme ve Excel</h2>
<div class="cols"><div class="card"><h3 style="margin-top:0">Yedek al</h3>
<p>Tüm verileriniz tek dosyada: <code>bakim_verileri.db</code>. Yedeği güvenli bir yere (USB, bulut) kopyalayın.</p>
<a class="btn" href="/yedek/db">Veritabanı yedeğini indir</a></div>
<div class="card"><h3 style="margin-top:0">Excel'e aktar</h3><p>Tüm sayfalar tek bir Excel dosyasında.</p>
<a class="btn" href="/yedek/excel">Excel olarak indir</a></div>
<div class="card"><h3 style="margin-top:0">Eski Excel'den içe aktar</h3>
<p><b>Dikkat:</b> KM, kondisyon, bakım ve parça kayıtları silinip dosyadan yeniden yüklenir. Harcama, evrak ve rehber korunur.</p>
<form method="post" action="/yedek/import" enctype="multipart/form-data" onsubmit="return confirm('Mevcut KM, kondisyon, bakım ve parça kayıtları silinecek. Devam edilsin mi?')">
<input type="file" name="f" accept=".xlsx" required><br><br><button class="sec">İçe aktar</button></form>{% if msg %}<p><b>{{msg}}</b></p>{% endif %}</div></div>{% endblock %}"""

app = Flask(__name__)
app.jinja_loader = DictLoader({"base.html": BASE_T, "dash.html": DASH_T, "list.html": LIST_T,
                               "edit.html": EDIT_T, "yedek.html": YEDEK_T, "araclar.html": ARACLAR_T})
AYLAR = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]

def show(col, v):
    if v in (None, ""):
        return ""
    return fdate(v) if col[2] == "date" else fint(v) if col[2] == "int" else fmoney(v) if col[2] == "num" else v

def field(col, row, new):
    from markupsafe import Markup, escape
    name, typ, opts = col[0], col[2], col[3]
    v = row.get(name) if hasattr(row, "get") else (row[name] if row else None)
    if v is None:
        v = dt.date.today().isoformat() if (new and typ == "date" and name == "date") else ""
    if typ == "select":
        o = "".join(f'<option{" selected" if x == v else ""}>{escape(x)}</option>' for x in opts)
        return Markup(f'<select name="{name}">{o}</select>')
    if typ == "area":
        return Markup(f'<textarea name="{name}" rows="2">{escape(v)}</textarea>')
    if typ == "date":
        return Markup(f'<input type="date" name="{name}" value="{escape(v)}">')
    if typ in ("int", "num") and v != "":
        v = fint(v) if typ == "int" else fmoney(v)
    return Markup(f'<input name="{name}" value="{escape(v)}"{" inputmode=decimal" if typ in ("int", "num") else ""}>')

@app.context_processor
def inject():
    db = get_db()
    aktif = cur_arac(db)
    araclar = db.execute("SELECT * FROM arac WHERE status='Aktif' OR id=? ORDER BY name", (aktif["id"],)).fetchall()
    db.close()
    return dict(aktif=aktif, araclar=araclar, scoped_tables=SCOPED, logo=LOGO, favicon=FAVICON, tables=TABLES, fint=fint, fmoney=fmoney, show=show, field=field, aylar=AYLAR)

# ----------------------------------------------------------------- sayfalar
@app.route("/")
def dash():
    db = get_db()
    aid = cur_arac(db)["id"]
    alerts, km, perday = build_alerts(db, aid)
    h = harcama_ozet(db, aid)
    db.close()
    return render_template("dash.html", active="", km=km, perday=perday, h=h,
                           dikkat=[a for a in alerts if a["level"] != "ok"], tamam=[a for a in alerts if a["level"] == "ok"])

def _tbl(name):
    if name not in TABLES:
        abort(404)
    return TABLES[name]

@app.route("/t/<name>")
def lst(name):
    if name == "arac":
        return redirect("/araclar")
    t = _tbl(name)
    db = get_db()
    aid = cur_arac(db)["id"]
    where, args = ("WHERE arac_id=?", (aid,)) if name in SCOPED else ("", ())
    rows = db.execute(f"SELECT * FROM {name} {where} ORDER BY {t['order']}", args).fetchall()
    kw = {}
    if name == "km":
        kw["charts"] = svg_charts(rows)
    if name == "harcama":
        kw["h"] = harcama_ozet(db, aid)
    if name == "bakim":
        cur, pd_ = km_stats(db, aid)
        kw["st"] = {r["id"]: bakim_status(r, cur, pd_) for r in rows}
    db.close()
    return render_template("list.html", active=name, name=name, t=t, rows=rows, **kw)

_back = lambda name: "/araclar" if name == "arac" else f"/t/{name}"

def _vals(t):
    return [parse_val(c, request.form.get(c[0])) for c in t["cols"]]

@app.route("/t/<name>/add", methods=["POST"])
def add(name):
    t = _tbl(name)
    cols = [c[0] for c in t["cols"]]
    vals = _vals(t)
    db = get_db()
    if name in SCOPED:
        cols, vals = cols + ["arac_id"], vals + [cur_arac(db)["id"]]
    db.execute(f"INSERT INTO {name}({','.join(cols)}) VALUES({','.join('?' * len(cols))})", vals)
    db.commit(); db.close()
    return redirect(_back(name))

@app.route("/t/<name>/edit/<int:i>", methods=["GET", "POST"])
def edit(name, i):
    t = _tbl(name)
    db = get_db()
    if request.method == "POST":
        cols = [c[0] for c in t["cols"]]
        db.execute(f"UPDATE {name} SET {','.join(c + '=?' for c in cols)} WHERE id=?", _vals(t) + [i])
        db.commit(); db.close()
        return redirect(_back(name))
    row = db.execute(f"SELECT * FROM {name} WHERE id=?", (i,)).fetchone()
    db.close()
    if not row:
        abort(404)
    return render_template("edit.html", active=name, name=name, t=t, row=dict(row))

@app.route("/t/<name>/delete/<int:i>", methods=["POST"])
def delete(name, i):
    _tbl(name)
    db = get_db()
    db.execute(f"DELETE FROM {name} WHERE id=?", (i,))
    db.commit(); db.close()
    return redirect(_back(name))

@app.route("/bakim/<int:i>/done", methods=["POST"])
def done(i):
    db = get_db()
    it = db.execute("SELECT * FROM bakim WHERE id=?", (i,)).fetchone()
    if it:
        cur, _ = km_stats(db, it["arac_id"])
        today = dt.date.today().isoformat()
        db.execute("UPDATE bakim SET last_date=?, last_km=? WHERE id=?", (today, cur, i))
        amt = parse_num(request.form.get("amount", "") or "")
        if amt:
            db.execute("INSERT INTO harcama(arac_id,date,category,descr,amount,km) VALUES(?,?,?,?,?,?)", (it["arac_id"], today, "Bakım", it["name"], amt, cur))
        db.commit()
    db.close()
    return redirect("/t/bakim")

@app.route("/araclar")
def araclar():
    db = get_db()
    tum = db.execute("SELECT * FROM arac ORDER BY (status='Aktif') DESC, name").fetchall()
    km = {a["id"]: km_stats(db, a["id"])[0] for a in tum}
    db.close()
    return render_template("araclar.html", active="araclar", tum=tum, km=km)

def _sec(db, aid):
    db.execute("INSERT OR REPLACE INTO ayar(key, value) VALUES('aktif_arac', ?)", (str(aid),))

@app.route("/arac/sec", methods=["POST"])
@app.route("/arac/sec/<int:aid>", methods=["POST"])
def arac_sec(aid=None):
    aid = aid or request.form.get("id", type=int)
    db = get_db()
    if aid and db.execute("SELECT 1 FROM arac WHERE id=?", (aid,)).fetchone():
        _sec(db, aid); db.commit()
    db.close()
    return redirect("/")

@app.route("/arac/ekle", methods=["POST"])
def arac_ekle():
    name = (request.form.get("name") or "").strip()
    if not name:
        return redirect("/araclar")
    year = parse_num(request.form.get("year", "") or "")
    db = get_db()
    cur = db.execute("INSERT INTO arac(name, plate, year, note, status) VALUES(?,?,?,?, 'Aktif')",
                     (name, (request.form.get("plate") or "").strip() or None, int(year) if year else None,
                      (request.form.get("note") or "").strip() or None))
    seed_arac(db, cur.lastrowid, bakim=bool(request.form.get("std")))
    _sec(db, cur.lastrowid)
    db.commit(); db.close()
    return redirect("/")

@app.route("/arac/sil/<int:aid>", methods=["POST"])
def arac_sil(aid):
    db = get_db()
    if db.execute("SELECT COUNT(*) FROM arac").fetchone()[0] > 1:
        for t in SCOPED:
            db.execute(f"DELETE FROM {t} WHERE arac_id=?", (aid,))
        db.execute("DELETE FROM arac WHERE id=?", (aid,))
        db.commit()
    db.close()
    return redirect("/araclar")

@app.route("/yedek")
def yedek():
    return render_template("yedek.html", active="yedek", msg=request.args.get("msg"))

@app.route("/yedek/db")
def yedek_db():
    buf = io.BytesIO()
    src = get_db(); mem = sqlite3.connect(":memory:")
    src.backup(mem); src.close()
    buf.write(mem.serialize()); mem.close(); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"bakim_verileri_yedek_{dt.date.today()}.db")

@app.route("/yedek/excel")
def yedek_excel():
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    db = get_db()
    for name, t in TABLES.items():
        ws = wb.create_sheet(re.sub(r"[\\/*?:\[\]]", "-", t["title"])[:31])
        sc = name in SCOPED
        ws.append((["Araç"] if sc else []) + [c[1] for c in t["cols"]])
        q = f"SELECT x.*, (SELECT name FROM arac WHERE arac.id=x.arac_id) AS _arac FROM {name} x ORDER BY {t['order']}" if sc else f"SELECT * FROM {name} ORDER BY {t['order']}"
        for r in db.execute(q):
            ws.append(([r["_arac"]] if sc else []) + [(pdate(r[c[0]]) if c[2] == "date" else r[c[0]]) for c in t["cols"]])
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 22
        for cell in ws[1]:
            cell.font = openpyxl.styles.Font(bold=True)
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, dt.date):
                    cell.number_format = "DD.MM.YYYY"
    db.close()
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"arac_bakim_{dt.date.today()}.xlsx")

@app.route("/yedek/import", methods=["POST"])
def yedek_import():
    f = request.files.get("f")
    db = get_db()
    try:
        import_legacy(db, f.stream, cur_arac(db)["id"])
        msg = "Excel verileri içe aktarıldı."
    except Exception as e:
        db.rollback(); msg = f"İçe aktarılamadı: {e}"
    db.close()
    return redirect(f"/yedek?msg={msg}")

def _port_free(port):
    sk = socket.socket()
    try:
        return sk.connect_ex(("127.0.0.1", port)) != 0
    finally:
        sk.close()

def main():
    # Konsolsuz (pencereli) exe'de stdout/stderr olmayabilir
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w")
    init_db()
    port = 5000
    while not _port_free(port):
        port += 1
    url = f"http://127.0.0.1:{port}"
    srv = threading.Thread(target=lambda: app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False), daemon=True)
    srv.start()
    for _ in range(50):  # sunucu hazır olana kadar bekle
        if not _port_free(port):
            break
        threading.Event().wait(0.1)
    print(f"Araç Bakım Takip çalışıyor: {url}")
    try:
        import webview  # kendi penceresinde aç
        try:
            webview.settings["ALLOW_DOWNLOADS"] = True
        except Exception:
            pass
        webview.create_window("Araç Bakım Takip", url, width=1200, height=820, min_size=(800, 560))
        webview.start()  # pencere kapanınca program da kapanır
        return
    except Exception as e:
        print("Pencere modu açılamadı, tarayıcıda açılıyor:", e)
    webbrowser.open(url)
    print("Kapatmak için bu pencereyi kapatın.")
    srv.join()

if __name__ == "__main__":
    main()
