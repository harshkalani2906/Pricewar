"""
Price Parity Tracker - live Nat Habit prices on 10 channels, checked against
the aligned price.

  Quick commerce, pincode by pincode : Blinkit, Zepto, Instamart, BigBasket, FK Minutes
  Marketplaces, one all-India price  : Amazon, Flipkart, Myntra, Nykaa, FirstCry

Inputs (folder 1_Input) - see inputs.py:
    Links.xlsx           Channel Name | Channel ID      which products to check
    Product Master.xlsx  Channel Name | Channel ID | Product Name | NH SKU
    Pincodes.xlsx        Channel Name | Pincode         (quick commerce only)
    Pincode Master.xlsx  Pincode | Tier | City | State | Region
    Aligned price        newest ..\\Price Check\\Aligned Price*.xlsx (or one in 1_Input)

Outputs:
    2_Output\\Price_Parity_DD-MM-YY-HH-MM.xlsx
    2_Output\\history\\qc_price_history.csv     (all runs, used by the dashboard)
    3_Dashboard\\price-pulse.html + data.js

    python qc_price_tracker.py                 every channel
    python qc_price_tracker.py --only qc       only the 5 quick-commerce apps
    python qc_price_tracker.py --only marketplaces
    python qc_price_tracker.py --only zepto,amazon
    python qc_price_tracker.py --pins 2        only the first 2 pincodes (test)
    python qc_price_tracker.py --show          show the browser windows
    python qc_price_tracker.py --login fkminutes   open a window to log in once
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import glob
import json
import os
import shutil
import re
import sys
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import channels as CH          # noqa: E402
import marketplaces as MP      # noqa: E402
import inputs as IN            # noqa: E402
from matcher import Matcher    # noqa: E402

EVERY = CH.ALL + MP.ALL         # quick-commerce (pincode-wise) + marketplaces (national)
NAMES = {c.key: c.name for c in EVERY}
ORDER = {c.key: i for i, c in enumerate(EVERY)}

IN_DIR = os.path.join(ROOT, "1_Input")
OUT_DIR = os.path.join(ROOT, "2_Output")
HIST_DIR = os.path.join(OUT_DIR, "history")
LOG_DIR = os.path.join(OUT_DIR, "_Logs")
DEBUG_DIR = os.path.join(LOG_DIR, "debug")
STATE_DIR = os.path.join(HERE, "state")
HIST_FILE = os.path.join(HIST_DIR, "qc_price_history.csv")
GEO_FILE = os.path.join(STATE_DIR, "geo_cache.json")
LINKS_FILE = os.path.join(IN_DIR, "Links.xlsx")
MASTER_FILE = os.path.join(IN_DIR, "Product Master.xlsx")
PIN_FILE = os.path.join(IN_DIR, "Pincodes.xlsx")
PINMASTER_FILE = os.path.join(IN_DIR, "Pincode Master.xlsx")
PROD_FILE = os.path.join(IN_DIR, "NH Products.xlsx")
PRIORITY_FILE = os.path.join(IN_DIR, "SKU Priority.xlsx")
PRICE_CHECK_DIR = os.path.join(os.path.dirname(ROOT), "Price Check")

HIST_COLS = ["Run Time", "Date", "Channel", "Pincode", "City", "NH SKU", "Channel ID",
             "Channel Product Name", "Pack Size", "MRP", "Selling Price", "Aligned SP",
             "Gap", "Gap %", "In Stock", "Stock Qty", "Status", "Match", "Seller", "URL"]

_log_fh = None


PREFIX = ""          # "[Zepto] " when running as one of the parallel workers


def log(msg: str) -> None:
    line = f"[{dt.datetime.now():%H:%M:%S}] {PREFIX}{msg}"
    try:
        print(line, flush=True)
    except UnicodeEncodeError:
        print(line.encode("ascii", "replace").decode(), flush=True)
    if _log_fh:
        _log_fh.write(line + "\n")
        _log_fh.flush()


def num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.\-]", "", str(v))
    try:
        return float(s) if s not in ("", "-", ".") else None
    except ValueError:
        return None


def clean_pin(v) -> str | None:
    s = re.sub(r"\D", "", str(v or "").split(".")[0])
    return s if len(s) == 6 else None


def head_map(row) -> dict:
    return {re.sub(r"[^a-z]", "", str(h or "").lower()): i for i, h in enumerate(row)}


# --------------------------------------------------------------------------
# Inputs
# --------------------------------------------------------------------------
def newest_aligned_file():
    files = glob.glob(os.path.join(IN_DIR, "Aligned Price*.xlsx")) + \
        glob.glob(os.path.join(PRICE_CHECK_DIR, "Aligned Price*.xlsx"))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    return max(files, key=os.path.getmtime) if files else None


def read_products():
    """{sku: {"name", "names", "aligned_sp", "aligned_mrp"}}, aligned label"""
    import openpyxl
    prods: dict[str, dict] = {}
    if os.path.exists(PROD_FILE):
        ws = openpyxl.load_workbook(PROD_FILE, read_only=True, data_only=True).worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
        h = head_map(rows[0])
        si = h.get("nhsku", 0)
        ni = h.get("productname", h.get("product", 1))
        ai = h.get("alignedsp")
        mi = h.get("alignedmrp")
        oi = h.get("othernames")
        for r in rows[1:]:
            if not r or not r[si]:
                continue
            sku = str(r[si]).strip()
            d = prods.setdefault(sku, {"name": None, "names": [], "aligned_sp": None, "aligned_mrp": None})
            if ni < len(r) and r[ni]:
                d["name"] = str(r[ni]).strip()
            if ai is not None and ai < len(r):
                d["aligned_sp"] = num(r[ai])
            if mi is not None and mi < len(r):
                d["aligned_mrp"] = num(r[mi])
            if oi is not None and oi < len(r) and r[oi]:
                d["names"] += [x.strip() for x in str(r[oi]).split("||") if x.strip()]
    label = "NH Products.xlsx"
    f = newest_aligned_file()
    if f:
        try:
            ws = openpyxl.load_workbook(f, read_only=True, data_only=True).worksheets[0]
            rows = list(ws.iter_rows(values_only=True))
            col = None
            hd = [str(x or "").lower() for x in rows[0]]
            for key in ("event", "aligned sp", "aligned"):
                col = next((i for i, x in enumerate(hd) if key in x and i > 0), None)
                if col is not None:
                    break
            for r in ([] if col is not None else rows[1:6]):
                for i, v in enumerate(r[1:], 1):
                    if num(v) is not None:
                        col = i
                        break
                if col:
                    break
            n = 0
            if col:
                for r in rows[1:]:
                    if r and r[0] and col < len(r) and num(r[col]) is not None:
                        sku = str(r[0]).strip()
                        prods.setdefault(sku, {"name": None, "names": [], "aligned_sp": None,
                                               "aligned_mrp": None})["aligned_sp"] = num(r[col])
                        n += 1
                label = f"{os.path.basename(f)} ({rows[0][col] or 'aligned SP'})"
                log(f"Aligned prices: {n} SKUs from {label}")
        except Exception as e:
            log(f"Could not read {f}: {e}")
    return prods, label


# --------------------------------------------------------------------------
# Browser
# --------------------------------------------------------------------------
# Apps read as a GUEST (never logged in). Instamart prices are cohort based: a logged-in
# account can see its own special price, so we always read the normal guest price.
# A fresh empty browser profile is used on every run - no login, no old cookies.
GUEST = {"instamart"}


def open_context(pw, key, show):
    if key in GUEST:
        prof = os.path.join(STATE_DIR, f"{key}_guest_profile")
        shutil.rmtree(prof, ignore_errors=True)
    else:
        prof = os.path.join(STATE_DIR, f"{key}_profile")
    os.makedirs(prof, exist_ok=True)
    args = ["--disable-blink-features=AutomationControlled", "--lang=en-IN"]
    if not show:
        args.append("--window-position=-32000,-32000")
    kw = dict(user_data_dir=prof, headless=False, args=args, viewport={"width": 1366, "height": 900},
              locale="en-IN", timezone_id="Asia/Kolkata",
              geolocation=None, permissions=[])
    try:
        ctx = pw.chromium.launch_persistent_context(channel="chrome", **kw)
    except Exception:
        ctx = pw.chromium.launch_persistent_context(**kw)
    ctx.add_init_script("Object.defineProperty(navigator,'webdriver',{get:()=>undefined})")
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.set_default_timeout(30000)
    return ctx, page


def page_problem(page):
    """A plain-English reason when the app shows a login wall or blocks the browser."""
    try:
        t = (page.inner_text("body", timeout=3000) or "")[:3000].lower()
    except Exception:
        return ""
    if "access denied" in t or "you don't have permission" in t:
        return "the app blocked the browser (Access Denied) - try again later"
    if ("log in" in t or "login" in t) and ("continue" in t or "phone number" in t) and len(t) < 600:
        return "the app asks to log in - double-click LOG IN TO APPS (once).bat"
    return ""


def shot(page, name):
    try:
        os.makedirs(DEBUG_DIR, exist_ok=True)
        if len(os.listdir(DEBUG_DIR)) < 60:
            page.screenshot(path=os.path.join(DEBUG_DIR, f"{dt.datetime.now():%d%m-%H%M%S}_{name}.png"))
    except Exception:
        pass


def load_geo():
    try:
        with open(GEO_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_geo(g):
    os.makedirs(STATE_DIR, exist_ok=True)
    with open(GEO_FILE, "w", encoding="utf-8") as f:
        json.dump(g, f, indent=1)


def resolve_all_geo(pw, pins, show):
    geo = load_geo()
    todo = [p for p in pins if p not in geo]
    if not todo:
        return geo
    log(f"Looking up map location for {len(todo)} new pincode(s)...")
    ctx, page = open_context(pw, "blinkit", show)
    try:
        page.goto("https://blinkit.com/", wait_until="domcontentloaded", timeout=60000)
        time.sleep(3)
        for p in todo:
            g = CH.resolve_geo(page, p, log)
            if g:
                geo[p] = g
                log(f"  {p} -> {g.get('city')} ({g['lat']:.4f}, {g['lon']:.4f})")
            else:
                log(f"  {p} -> not found")
            time.sleep(0.8)
    finally:
        ctx.close()
    save_geo(geo)
    return geo


# --------------------------------------------------------------------------
# Run one channel
# --------------------------------------------------------------------------
def run_channel(pw, cls, pins, geo, cities, show, results, ch_status):
    ch = cls()
    log(f"=== {ch.name}: {len(pins)} pincode(s)" + (" - guest, no login" if ch.key in GUEST else "") + " ===")
    ctx, page = open_context(pw, ch.key, show)
    ok_pins = 0
    try:
        if getattr(ch, "home", None):
            try:
                page.goto(ch.home, wait_until="domcontentloaded", timeout=60000)
                time.sleep(3)
            except Exception:
                pass
        for i, pin in enumerate(pins, 1):
            g = geo.get(pin)
            city = cities.get(pin) or (g or {}).get("city") or ""
            listings, note = [], ""
            for attempt in (1, 2):
                try:
                    if not ch.set_location(page, pin, g):
                        note = "could not set location"
                        shot(page, f"{ch.key}_{pin}_location")
                        time.sleep(3)
                        continue
                    listings = ch.fetch(page, pin, g, log)
                    if listings:
                        note = ""
                        break
                    note = "no Nat Habit products found (store may not serve this pincode)"
                    why = page_problem(page)
                    if why:
                        note = why
                    shot(page, f"{ch.key}_{pin}_empty")
                except Exception as e:
                    note = f"error: {str(e).splitlines()[0][:150]}"
                    shot(page, f"{ch.key}_{pin}_error")
                time.sleep(4)
            if listings:
                ok_pins += 1
            log(f"  [{i}/{len(pins)}] {pin} {city}: {len(listings)} NH listing(s)"
                + (f" - {note}" if note else ""))
            results.append({"channel": ch.key, "pin": pin, "city": city,
                            "listings": listings, "note": note})
            time.sleep(2.5)
    finally:
        try:
            ctx.close()
        except Exception:
            pass
    ch_status[ch.key] = {"pins": len(pins), "ok": ok_pins}


# --------------------------------------------------------------------------
# Turn raw listings into rows
# --------------------------------------------------------------------------
def _row(run_time, ch, pin, city, **kw):
    r = {c: "" for c in HIST_COLS}
    r.update({"Run Time": run_time.strftime("%Y-%m-%d %H:%M"), "Date": run_time.date().isoformat(),
              "Channel": NAMES[ch], "Pincode": pin, "City": city})
    r.update(kw)
    return r


def _status(sku, L, al, tol):
    sp = L.get("sp")
    gap = round(sp - al, 2) if (sp is not None and al) else None
    if not sku:
        st = "Not mapped"
    elif not L.get("in_stock"):
        st = "Out of stock"
    elif al is None:
        st = "No aligned price"
    elif sp is None:
        st = "No price"
    elif abs(gap) <= tol:
        st = "At par"
    elif gap < 0:
        st = "Below aligned"
    else:
        st = "Above aligned"
    return st, gap


def build_rows(results, links, master, prods, run_time, tol):
    """One row per (channel, pincode, product).
    Products in Links.xlsx get a price + status (or 'Not available here' when the app
    does not show them at that pincode). Nat Habit products the app shows that are
    NOT in Links.xlsx are kept as 'Not in Links' so they can be added later."""
    matcher = Matcher(prods)
    rows, extra = [], {}
    for res in results:
        ch, pin, city = res["channel"], res["pin"], res["city"]
        want = [str(x) for x in links.get(ch, [])]
        low = {w.lower(): w for w in want}
        if res.get("note") and not res["listings"]:
            for cid in want:
                m = master.get((ch, cid)) or {}
                rows.append(_row(run_time, ch, pin, city, **{
                    "NH SKU": m.get("sku", ""), "Channel ID": cid, "Status": "Check failed",
                    "Aligned SP": (prods.get(m.get("sku")) or {}).get("aligned_sp"),
                    "Channel Product Name": res["note"], "Match": "Product Master" if m else ""}))
            continue
        found = set()
        for L in res["listings"]:
            ids = [str(L["id"])] + [str(x) for x in (L.get("alt") or []) if x]
            cid = next((low[i.lower()] for i in ids if i.lower() in low), None)
            base = {"Channel Product Name": L.get("name") or "", "Pack Size": L.get("size") or "",
                    "MRP": L.get("mrp"), "Selling Price": L.get("sp"),
                    "In Stock": "Yes" if L.get("in_stock") else "No", "Stock Qty": L.get("qty"),
                    "Seller": L.get("seller") or "", "URL": L.get("url") or ""}
            if cid is None:                                   # shown by the app, not in Links
                sku, pct = matcher.guess(L.get("name") or "", L.get("size"))
                m = next((master.get((ch, i)) for i in ids if master.get((ch, i))), None)
                guess = (m or {}).get("sku") or sku
                note = (f"In Product Master as {m['sku']}" if m else
                        f"Guess: {sku} ({pct}%)" if sku else "No guess")
                rows.append(_row(run_time, ch, pin, city, **base, **{
                    "Channel ID": ids[0], "Status": "Not in Links", "Match": note}))
                e = extra.setdefault((ch, ids[0]), {"Channel Name": NAMES[ch], "Channel ID": ids[0],
                                                     "Product Name": L.get("name") or "",
                                                     "Pack Size": L.get("size") or "",
                                                     "Suggested NH SKU": guess or "", "Why": note,
                                                     "Seen at pincodes": set(), "URL": L.get("url") or "",
                                                     "Other ids": ", ".join(ids[1:10])})
                e["Seen at pincodes"].add(pin)
                continue
            found.add(cid)
            m = next((master.get((ch, i)) for i in [cid] + ids if master.get((ch, i))), None)
            sku = (m or {}).get("sku")
            al = (prods.get(sku) or {}).get("aligned_sp") if sku else None
            st, gap = _status(sku, L, al, tol)
            rows.append(_row(run_time, ch, pin, city, **base, **{
                "NH SKU": sku or "", "Channel ID": cid, "Aligned SP": al, "Gap": gap,
                "Gap %": round(gap / al * 100, 1) if (gap is not None and al) else None,
                "Status": st, "Match": "Product Master" if m else "Missing in Product Master"}))
        national = ch in IN.MP_KEYS
        for cid in want:
            if cid in found:
                continue
            m = master.get((ch, cid)) or {}
            sku = m.get("sku", "")
            rows.append(_row(run_time, ch, pin, city, **{
                "NH SKU": sku, "Channel ID": cid,
                "Aligned SP": (prods.get(sku) or {}).get("aligned_sp") if sku else None,
                "Status": "Not live / not read" if national else "Not available here",
                "Match": "Product Master" if m else "Missing in Product Master",
                "Channel Product Name": m.get("name") or (prods.get(sku) or {}).get("name") or "",
                "URL": MP.URL[ch].format(cid) if national else ""}))
    return rows, extra

# --------------------------------------------------------------------------
# Excel + history
# --------------------------------------------------------------------------
FILL = {"At par": "D9F2D9", "Below aligned": "F8D0D0", "Above aligned": "FFE9B3",
        "Not in Links": "FFF4CC", "Not live / not read": "F3C6E8",
        "Out of stock": "E5E5E5", "Not available here": "EFEFEF", "Not mapped": "DCE8FA",
        "Check failed": "F3C6E8", "No aligned price": "F3F3F3", "No price": "F3F3F3"}


def write_excel(rows, prods, run_time, aligned_label, ch_status):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    os.makedirs(OUT_DIR, exist_ok=True)
    path = os.path.join(OUT_DIR, f"Price_Parity_{run_time:%d-%m-%y-%H-%M}.xlsx")
    wb = openpyxl.Workbook()
    hdr_font = Font(bold=True, color="FFFFFF")
    hdr_fill = PatternFill("solid", fgColor="1F2937")

    # --- Parity Report: one row per SKU, per channel lowest / highest SP
    ws = wb.active
    ws.title = "Parity Report"
    chans = [c.name for c in EVERY if any(r["Channel"] == c.name for r in rows)]
    heads = ["NH SKU", "Product", "Aligned SP"]
    for c in chans:
        heads += [f"{c} Lowest SP", f"{c} Highest SP", f"{c} Status"]
    heads += ["Lowest anywhere", "Highest anywhere", "Spread", "Worst gap vs aligned"]
    ws.append([f"Aligned price source: {aligned_label}   |   Run: {run_time:%d %b %Y %H:%M}"])
    ws.append(heads)
    by = {}
    for r in rows:
        if r["NH SKU"] and isinstance(r["Selling Price"], (int, float)):
            by.setdefault(r["NH SKU"], {}).setdefault(r["Channel"], []).append(r)
    for sku in sorted(by):
        al = (prods.get(sku) or {}).get("aligned_sp")
        line = [sku, (prods.get(sku) or {}).get("name"), al]
        allsp = []
        for c in chans:
            rs = [r for r in by[sku].get(c, []) if r["In Stock"] == "Yes"] or by[sku].get(c, [])
            sps = [r["Selling Price"] for r in rs]
            if not sps:
                line += [None, None, "Not listed"]
                continue
            lo, hi = min(sps), max(sps)
            allsp += sps
            if al:
                st = ("Below aligned" if lo < al - 1 else "Above aligned" if lo > al + 1 else "At par")
                if hi > al + 1 and st == "At par":
                    st = "Above aligned"
            else:
                st = "No aligned price"
            line += [lo, hi, st]
        lo, hi = (min(allsp), max(allsp)) if allsp else (None, None)
        worst = None
        if al and allsp:
            worst = max((s - al for s in allsp), key=abs)
        line += [lo, hi, (hi - lo) if allsp else None, worst]
        ws.append(line)
    for c in ws[2]:
        c.font, c.fill = hdr_font, hdr_fill
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for row in ws.iter_rows(min_row=3):
        for c in row:
            if isinstance(c.value, str) and c.value in FILL:
                c.fill = PatternFill("solid", fgColor=FILL[c.value])
    ws.freeze_panes = "D3"
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 46

    # --- Detail
    wd = wb.create_sheet("Detail")
    wd.append(HIST_COLS)
    for r in rows:
        wd.append([r.get(c) for c in HIST_COLS])
    for c in wd[1]:
        c.font, c.fill = hdr_font, hdr_fill
    st_col = HIST_COLS.index("Status")
    for row in wd.iter_rows(min_row=2):
        v = row[st_col].value
        if v in FILL:
            row[st_col].fill = PatternFill("solid", fgColor=FILL[v])
    wd.freeze_panes = "A2"
    wd.auto_filter.ref = wd.dimensions
    for col, w in zip("ABCDEFGHIJKLMNOPQRS", (16, 11, 11, 9, 13, 18, 22, 55, 14, 8, 10, 10, 8, 8, 8, 8, 18, 12, 40)):
        wd.column_dimensions[col].width = w

    # --- Summary
    wsu = wb.create_sheet("Summary")
    wsu.append(["Channel", "Pincodes checked", "Pincodes with data", "Prices read", "At par",
                "Below aligned", "Above aligned", "Out of stock", "Not available here", "Not mapped",
                "Not in Links", "Not live / not read", "Check failed"])
    for c in EVERY:
        rs = [r for r in rows if r["Channel"] == c.name]
        if not rs and c.key not in ch_status:
            continue
        cnt = lambda s: sum(1 for r in rs if r["Status"] == s)  # noqa: E731
        st = ch_status.get(c.key, {})
        wsu.append([c.name, st.get("pins"), st.get("ok"),
                    sum(1 for r in rs if r["Selling Price"] not in ("", None)), cnt("At par"), cnt("Below aligned"),
                    cnt("Above aligned"), cnt("Out of stock"), cnt("Not available here"), cnt("Not mapped"),
                    cnt("Not in Links"), cnt("Not live / not read"), cnt("Check failed")])
    for c in wsu[1]:
        c.font, c.fill = hdr_font, hdr_fill
    for col in "ABCDEFGHIJKLM":
        wsu.column_dimensions[col].width = 16
    wb.save(path)
    return path


def append_history(rows):
    os.makedirs(HIST_DIR, exist_ok=True)
    new = not os.path.exists(HIST_FILE)
    with open(HIST_FILE, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=HIST_COLS, extrasaction="ignore")
        if new:
            w.writeheader()
        for r in rows:
            w.writerow(r)


def trim_history(keep_days=7):
    """Keeps only the last 7 days: price history rows, Excel reports, run logs, debug screenshots."""
    old = time.time() - keep_days * 86400
    for pat in (os.path.join(OUT_DIR, "Price_Parity_*.xlsx"), os.path.join(OUT_DIR, "QC_Prices_*.xlsx"),
                os.path.join(LOG_DIR, "run_*.txt"), os.path.join(DEBUG_DIR, "*.png")):
        for f in glob.glob(pat):
            try:
                if os.path.getmtime(f) < old:
                    os.remove(f)
            except OSError:
                pass
    if not os.path.exists(HIST_FILE):
        return
    cut = (dt.date.today() - dt.timedelta(days=keep_days)).isoformat()
    with open(HIST_FILE, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    keep = [r for r in rows if (r.get("Date") or "") >= cut]
    if len(keep) == len(rows):
        return
    with open(HIST_FILE, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=HIST_COLS, extrasaction="ignore")
        w.writeheader()
        w.writerows(keep)


# --------------------------------------------------------------------------
def login_window(key):
    from playwright.sync_api import sync_playwright
    cls = CH.BY_KEY.get(key)
    if not cls:
        raise SystemExit(f"Unknown channel {key}")
    url = {"fkminutes": "https://www.flipkart.com/account/login",
           "zepto": "https://www.zepto.com/", "blinkit": "https://blinkit.com/",
           "instamart": "https://www.swiggy.com/instamart", "bigbasket": "https://www.bigbasket.com/"}[key]
    if key in GUEST:
        print(f"\n  {cls.name} is always read WITHOUT login (guest price) - no need to log in.")
        return
    with sync_playwright() as pw:
        ctx, page = open_context(pw, key, show=True)
        page.goto(url)
        print(f"\n  {cls.name}: a browser window opened.")
        print("  1. Click Login / Sign in and log in with the phone number + OTP.")
        print("  2. Wait until you see you are logged in.")
        input(f"  3. Come back to this black window and press Enter (closes {cls.name})... ")
        ctx.close()


def write_attention(path, problems, links, master, extra):
    """Adds a 'Needs Attention' sheet: what to fix in the 4 input files."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    wb = openpyxl.load_workbook(path)
    ws = wb.create_sheet("Needs Attention", 1)
    bold = Font(bold=True, color="FFFFFF")
    dark = PatternFill("solid", fgColor="1F2937")
    sec = Font(bold=True, size=12)

    def header(cols):
        ws.append(cols)
        for c in ws[ws.max_row]:
            c.font, c.fill = bold, dark

    ws.append(["1. In Links.xlsx but NOT in Product Master.xlsx  ->  add these to Product Master with their NH SKU"])
    ws[ws.max_row][0].font = sec
    header(["Channel Name", "Channel ID"])
    n1 = 0
    for k, ids in links.items():
        for cid in ids:
            if (k, str(cid)) not in master:
                ws.append([IN.NAMES[k], cid])
                n1 += 1
    if not n1:
        ws.append(["(none - all good)"])
    ws.append([])
    ws.append(["2. Nat Habit products the apps show that are NOT in Links.xlsx  ->  add to Links "
               "(and Product Master) if you want them tracked"])
    ws[ws.max_row][0].font = sec
    header(["Channel Name", "Channel ID", "Product Name", "Pack Size", "Suggested NH SKU", "Why",
            "Seen at pincodes", "URL", "Other ids the app sends"])
    for e in sorted(extra.values(), key=lambda e: (e["Channel Name"], e["Product Name"])):
        ws.append([e["Channel Name"], e["Channel ID"], e["Product Name"], e["Pack Size"],
                   e["Suggested NH SKU"], e["Why"], ", ".join(sorted(e["Seen at pincodes"])), e["URL"],
                   e.get("Other ids", "")])
    if not extra:
        ws.append(["(none)"])
    ws.append([])
    ws.append(["3. Rows in your input files that could not be read"])
    ws[ws.max_row][0].font = sec
    header(["File", "Row", "Problem", "Value"])
    for p in problems:
        ws.append(list(p))
    if not problems:
        ws.append(["(none)"])
    for col, w in zip("ABCDEFGH", (16, 40, 60, 14, 20, 30, 30, 50)):
        ws.column_dimensions[col].width = w
    wb.save(path)
    return n1, len(extra), len(problems)


SELECTED = None          # channels the user asked for (Amazon + Amazon Now are read together)


def parse_only(text):
    global SELECTED
    only = set()
    for x in text.split(","):
        s = x.strip().lower().replace(" ", "")
        if not s:
            continue
        if s in ("qc", "quickcommerce"):
            only |= set(IN.QC_KEYS) | set(IN.NOW_KEYS)
        elif s in ("marketplaces", "marketplace", "mp"):
            only |= set(IN.MP_KEYS)
        elif IN.channel_key(s) in IN.NAMES:
            only.add(IN.channel_key(s))
        else:
            log(f"  --only: unknown channel '{x}' ignored")
    SELECTED = set(only) if only else None
    if only and (only & {"amazon", "amazonnow"}):
        only |= {"amazon", "amazonnow"}            # one Amazon read gives both; extra one is dropped later
    return only or None


# ---------------------------------------------------------------------------
# Menu (double-click RUN PRICE TRACKER.bat) - like the SLA tracker
# ---------------------------------------------------------------------------
MENU = [("amazon", "Amazon"), ("amazonnow", "Amazon Now"), ("flipkart", "Flipkart"), ("myntra", "Myntra"),
        ("nykaa", "Nykaa"), ("firstcry", "FirstCry"), ("blinkit", "Blinkit"), ("zepto", "Zepto"),
        ("instamart", "Instamart"), ("bigbasket", "BigBasket"), ("fkminutes", "FK Minutes")]


def menu():
    """Returns the command-line arguments for the choice, or None to exit."""
    while True:
        print("""
  ==========================================================
     PRICE WAR  -  Nat Habit prices on every channel
  ==========================================================
     1  Run ALL channels
""")
        for i, (k, n) in enumerate(MENU, 2):
            print(f"    {i:>2}  {n}")
        print("""
     Q  Only quick commerce (QC apps + Amazon Now)
     M  Only marketplaces (Amazon, Flipkart, Myntra, Nykaa, FirstCry)
     T  Test run - 2 pincodes, browser windows visible
     D  Rebuild and open the dashboard
     0  Exit

     More than one channel: type the numbers with commas, like 2,3,8
""")
        c = input("  Choose: ").strip().lower().replace(" ", "")
        if c in ("", "0"):
            return None
        if c == "d":
            return ["--dashboard-only"]
        args = []
        if c == "t":
            args = ["--pins", "2", "--show"]
        elif c == "q":
            args = ["--only", "qc"]
        elif c == "m":
            args = ["--only", "marketplaces"]
        elif c != "1":
            keys = []
            for x in c.split(","):
                if x.isdigit() and 2 <= int(x) < 2 + len(MENU):
                    keys.append(MENU[int(x) - 2][0])
                else:
                    keys = None
                    break
            if not keys:
                print("\n  Not understood - type a number from the list.")
                continue
            args = ["--only", ",".join(keys)]
        t = input("\n  Which SKUs on the marketplaces?  1 = Top 50   2 = 80% revenue (Enter)   3 = All : ").strip()
        args += ["--tier", {"1": "top50", "3": "all"}.get(t, "80")]
        return args


UNIT_NAME = dict(IN.NAMES, mp="Marketplaces")


def run_parallel(units, a):
    """Runs every quick-commerce app + the marketplaces at the SAME TIME, each in its own
    Python process (own browser), then collects their results."""
    import subprocess
    part_dir = os.path.join(LOG_DIR, "_parts")
    os.makedirs(part_dir, exist_ok=True)
    procs = {}
    for u in units:
        out = os.path.join(part_dir, f"{u}.json")
        if os.path.exists(out):
            os.remove(out)
        cmd = [sys.executable, os.path.abspath(__file__), "--worker", u, "--worker-out", out, "--no-dashboard",
               "--tolerance", str(a.tolerance)]
        if a.pins:
            cmd += ["--pins", str(a.pins)]
        if a.show:
            cmd.append("--show")
        if a.only:
            cmd += ["--only", a.only]
        cmd += ["--tier", a.tier]
        procs[u] = (subprocess.Popen(cmd, cwd=ROOT), out, time.time())
    log("Checking at the same time: " + ", ".join(UNIT_NAME[u] for u in units))
    results, status = [], {}
    for u, (proc, out, t0) in procs.items():
        try:
            proc.wait(timeout=3 * 3600)
        except subprocess.TimeoutExpired:
            proc.kill()
            log(f"!! {UNIT_NAME[u]} took more than 3 hours - stopped")
        if os.path.exists(out):
            with open(out, encoding="utf-8") as fh:
                d = json.load(fh)
            results += d["results"]
            status.update(d["status"])
            log(f"  {UNIT_NAME[u]} finished in {(time.time() - t0) / 60:.1f} min")
        else:
            log(f"!! {UNIT_NAME[u]} stopped without results - see _Logs\\worker_{u}.txt")
        wl = os.path.join(LOG_DIR, f"worker_{u}.txt")
        if _log_fh and os.path.exists(wl):
            with open(wl, encoding="utf-8", errors="replace") as fh:
                _log_fh.write(fh.read())
            _log_fh.flush()
    return results, status


def main():
    global _log_fh, PREFIX
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", default="", help="comma list: blinkit,zepto,instamart,bigbasket,fkminutes,"
                    "amazon,flipkart,myntra,nykaa,firstcry  (also: qc, marketplaces)")
    ap.add_argument("--no-marketplaces", action="store_true", help="skip Amazon/Flipkart/Myntra/Nykaa/FirstCry")
    ap.add_argument("--pins", type=int, default=0, help="only the first N pincodes per app (testing)")
    ap.add_argument("--show", action="store_true", help="show the browser windows")
    ap.add_argument("--tolerance", type=float, default=1.0, help="rupees still counted as At par")
    ap.add_argument("--login", default="", help="open a window to log in to one app")
    ap.add_argument("--no-dashboard", action="store_true")
    ap.add_argument("--serial", action="store_true", help="one app after another (slower, for testing)")
    ap.add_argument("--tier", default="80", choices=["top50", "80", "all"],
                    help="marketplace SKUs to check: top50 = Top 50 revenue SKUs, 80 = SKUs making 80%% of revenue "
                         "(default), all = every id in Links (from 1_Input\\SKU Priority.xlsx)")
    ap.add_argument("--dashboard-only", action="store_true", help="rebuild + open the dashboard")
    ap.add_argument("--worker", default="", help=argparse.SUPPRESS)
    ap.add_argument("--worker-out", default="", help=argparse.SUPPRESS)
    argv = sys.argv[1:]
    if not argv and sys.stdin and sys.stdin.isatty():
        argv = menu()
        if argv is None:
            return
    a = ap.parse_args(argv)
    if a.dashboard_only:
        import build_dashboard as BD
        BD.build(log=print, tolerance=a.tolerance)
        p = os.path.join(ROOT, "3_Dashboard", "price-pulse.html")
        if os.path.exists(p):
            os.startfile(p) if hasattr(os, "startfile") else print(p)
        return

    if a.login:
        login_window(IN.channel_key(a.login))
        return

    os.makedirs(LOG_DIR, exist_ok=True)
    if not a.worker:
        # only one run at a time (the 40-minute and daily runs can meet) - the browsers share profiles
        lock = os.path.join(STATE_DIR, "run.lock")
        os.makedirs(STATE_DIR, exist_ok=True)
        if os.path.exists(lock) and time.time() - os.path.getmtime(lock) < 60 * 60:
            print("Another Price War run is still going - this one is skipped.")
            with open(os.path.join(LOG_DIR, "skipped_runs.txt"), "a", encoding="utf-8") as f:
                f.write(f"{dt.datetime.now():%d-%m-%Y %H:%M}  skipped (--tier {a.tier}) - previous run still going\n")
            return
        with open(lock, "w") as f:
            f.write(str(os.getpid()))
        import atexit
        atexit.register(lambda: os.path.exists(lock) and os.remove(lock))
    run_time = dt.datetime.now().replace(second=0, microsecond=0)
    if a.worker:
        PREFIX = f"[{UNIT_NAME.get(a.worker, a.worker)}] "
        _log_fh = open(os.path.join(LOG_DIR, f"worker_{a.worker}.txt"), "w", encoding="utf-8")
    else:
        _log_fh = open(os.path.join(LOG_DIR, f"run_{run_time:%Y-%m-%d_%H-%M}.txt"), "w", encoding="utf-8")
    log("Price Parity Tracker started")

    only = parse_only(a.only)
    if a.no_marketplaces:
        only = (only or set(IN.ALL_KEYS)) - set(IN.MP_KEYS) - set(IN.NOW_KEYS)
    problems = []
    links = IN.read_links(LINKS_FILE, log, problems)
    master = IN.read_master(MASTER_FILE, log, problems)
    jobs, cities = IN.read_pincodes(PIN_FILE, PINMASTER_FILE, log, problems)
    now_pins = sorted({p for v in jobs.values() for p in v})      # Amazon Now: the quick-commerce pincodes
    if only:
        links = {k: v for k, v in links.items() if k in only}
        jobs = {k: v for k, v in jobs.items() if k in only}
    jobs = {k: v for k, v in jobs.items() if v}
    tiers = IN.read_priority(PRIORITY_FILE)
    if tiers and a.tier != "all":
        keep = {"Top 50"} if a.tier == "top50" else {"Top 50", "80% revenue"}
        cut = 0
        for k in IN.MP_KEYS:
            ids = links.get(k, [])
            links[k] = [i for i in ids if tiers.get((master.get((k, i)) or {}).get("sku")) in keep]
            cut += len(ids) - len(links[k])
        log(f"SKU Priority: checking {'Top 50' if a.tier == 'top50' else '80% revenue'} SKUs on marketplaces "
            f"({cut} lower-revenue id(s) skipped this run)")
    # Amazon Now = the same ASINs as Amazon (unless Links has its own "Amazon Now" rows)
    if links.get("amazon") and not links.get("amazonnow"):
        links["amazonnow"] = list(links["amazon"])
    for (k, cid), m in list(master.items()):
        if k == "amazon":
            master.setdefault(("amazonnow", cid), m)
    if a.pins:
        jobs = {k: v[:a.pins] for k, v in jobs.items()}
        now_pins = now_pins[:a.pins]
    qc_keys = [k for k in IN.QC_KEYS if k in jobs]
    for k in IN.QC_KEYS:
        if links.get(k) and k not in jobs and (only is None or k in only):
            log(f"  {IN.NAMES[k]}: {len(links[k])} id(s) in Links but no pincode in Pincodes.xlsx - skipped")
        if k in jobs and not links.get(k):
            log(f"  {IN.NAMES[k]}: no ids in Links.xlsx yet - every Nat Habit item it shows will be 'Not in Links'")
    mp_keys = MP.available([k for k in IN.MP_KEYS if links.get(k)], log) if a.worker in ("", "mp") else []
    if a.worker:
        qc_keys = [k for k in qc_keys if k == a.worker]
    if not qc_keys and not mp_keys:
        raise SystemExit("Nothing to check - fill 1_Input\\Links.xlsx and 1_Input\\Pincodes.xlsx")
    log("Links.xlsx: " + ", ".join(f"{IN.NAMES[k]} {len(v)}" for k, v in links.items() if v))
    log(f"Product Master: {len(master)} Nat Habit id(s) mapped to NH SKU")
    for k in qc_keys:
        log(f"  {IN.NAMES[k]}: {len(jobs[k])} pincode(s)")
    prods, aligned_label = read_products()
    for (k, cid), m in master.items():
        d = prods.setdefault(m["sku"], {"name": None, "names": [], "aligned_sp": None, "aligned_mrp": None})
        if m.get("name"):
            d["names"].append(m["name"])
            d["name"] = d["name"] or m["name"]

    results, ch_status = [], {}
    mp_pin = next((p for v in jobs.values() for p in v), "110001")
    units = qc_keys + (["mp"] if mp_keys else [])
    if not a.worker and not a.serial and len(units) > 1:
        if qc_keys:                                   # map locations once, before the apps start
            from playwright.sync_api import sync_playwright
            with sync_playwright() as pw:
                resolve_all_geo(pw, sorted({p for k in qc_keys for p in jobs[k]}), a.show)
        results, ch_status = run_parallel(units, a)
        qc_keys, mp_keys = [], []
    if qc_keys:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            all_pins = sorted({p for k in qc_keys for p in jobs[k]})
            geo = resolve_all_geo(pw, all_pins, a.show)
            for cls in CH.ALL:
                if cls.key not in qc_keys:
                    continue
                try:
                    run_channel(pw, cls, jobs[cls.key], geo, cities, a.show, results, ch_status)
                except Exception as e:
                    log(f"!! {cls.name} stopped: {e}")
                    log(traceback.format_exc()[-800:])

    if mp_keys:
        try:
            mres = MP.fetch({k: links[k] for k in mp_keys}, log, _log_fh, pin=mp_pin,
                            now_pins=now_pins or [mp_pin], cities=cities)
            results += mres
            for r in mres:
                s = ch_status.setdefault(r["channel"], {"pins": 0, "ok": 0})
                s["pins"] += 1
                s["ok"] += 1 if r["listings"] else 0
        except Exception as e:
            log(f"!! Marketplaces stopped: {e}")
            log(traceback.format_exc()[-800:])

    if a.worker:
        with open(a.worker_out, "w", encoding="utf-8") as fh:
            json.dump({"results": results, "status": ch_status}, fh, default=list)
        log("finished")
        return

    if SELECTED:                                  # e.g. only Amazon Now asked -> drop the Amazon rows
        results = [r for r in results if r["channel"] in SELECTED]
        links = {k: (v if k in SELECTED else []) for k, v in links.items()}
    rows, extra = build_rows(results, links, master, prods, run_time, a.tolerance)
    path = write_excel(rows, prods, run_time, aligned_label, ch_status)
    n1, n2, n3 = write_attention(path, problems, links, master, extra)
    append_history(rows)
    trim_history()
    log(f"Saved {os.path.relpath(path, ROOT)}  ({len(rows)} rows)")
    for k, s in ch_status.items():
        log(f"  {NAMES[k]}: " + (f"data from {s['ok']} of {s['pins']} pincode(s)" if (k in CH.BY_KEY or k == "amazonnow")
                                 else ("prices read" if s["ok"] else "NO prices read")))
    if n1 or n2 or n3:
        log(f"Needs Attention: {n1} Links id(s) missing in Product Master, {n2} app item(s) not in Links, "
            f"{n3} unreadable input row(s) - see the 'Needs Attention' sheet")
    if not a.no_dashboard:
        try:
            import build_dashboard
            build_dashboard.build(log=log, aligned_label=aligned_label, tolerance=a.tolerance)
        except Exception as e:
            log(f"Dashboard not rebuilt: {e}")
    log("Done.")


if __name__ == "__main__":
    main()
