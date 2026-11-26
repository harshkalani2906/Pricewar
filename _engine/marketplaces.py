"""
Amazon, Flipkart, Myntra, Nykaa and FirstCry prices for the Price Parity Tracker.

These prices are the same for all of India, so each is checked once per run
and shown under City = "All India".

Which products: the Channel IDs in 1_Input\\Links.xlsx (Channel Name = Amazon /
Flipkart / Myntra / Nykaa / FirstCry). Reuses the existing scrapers in the
folders next to the tracker:
    Amazon / Flipkart / Myntra : ..\\Price Check (price_check.py) + ..\\Amazon, ..\\Flipkart, ..\\Myntra
    Nykaa                      : ..\\Nykaa   (brand shelf + product details by id)
    FirstCry                   : ..\\FirstCry (product lookup by FCID, 20 per call)
"""
from __future__ import annotations

import argparse
import datetime as _dt
import glob
import importlib.util
import os
import sys
import threading

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PS = os.path.dirname(ROOT)                     # the folder that holds QC Price Tracker
NATIONAL_PIN = "All India"
NYKAA_BRAND_CATEGORY = "24386"


class _MP:
    national = True


class Amazon(_MP):
    key, name = "amazon", "Amazon"


class Flipkart(_MP):
    key, name = "flipkart", "Flipkart"


class Myntra(_MP):
    key, name = "myntra", "Myntra"


class Nykaa(_MP):
    key, name = "nykaa", "Nykaa"


class FirstCry(_MP):
    key, name = "firstcry", "FirstCry"


ALL = [Amazon, Flipkart, Myntra, Nykaa, FirstCry]
NEEDS = {"amazon": ["Price Check", "Amazon"], "flipkart": ["Price Check", "Flipkart"],
         "myntra": ["Price Check", "Myntra"], "nykaa": ["Nykaa"], "firstcry": ["FirstCry"]}
URL = {"amazon": "https://www.amazon.in/dp/{}", "flipkart": "https://www.flipkart.com/product/p/itme?pid={}",
       "myntra": "https://www.myntra.com/{}", "nykaa": "https://www.nykaa.com/x/p/{}",
       "firstcry": "https://www.firstcry.com/x/x/{}/product-detail"}


def _paths():
    for sub in ("Price Check", "FirstCry", "Nykaa", "Myntra", "Flipkart", "Amazon",
                os.path.join("Amazon", "_engine")):
        p = os.path.join(PS, sub)
        if os.path.isdir(p) and p not in sys.path:
            sys.path.insert(0, p)


def available(keys, log):
    """Keeps only the channels whose scraper folders are next to the tracker."""
    ok = []
    for k in keys:
        miss = [n for n in NEEDS[k] if not os.path.isdir(os.path.join(PS, n))]
        if miss:
            log(f"{ {c.key: c.name for c in ALL}[k]}: skipped - folder(s) missing next to QC Price Tracker: {', '.join(miss)}")
        else:
            ok.append(k)
    return ok


def _args():
    return argparse.Namespace(headful=True, offscreen=True, delay=1.0, tolerance=1.0, only="",
                              pincode="", no_pincode=True, pages=1, limit=0, top=0,
                              no_seller=True, no_sellers=False, no_http2=False, workers=3)


def _f(v):
    try:
        return float(v) if v not in (None, "") else None
    except (TypeError, ValueError):
        return None


# Amazon + Flipkart use the SLA Tracker v2 readers (direct API, many requests at once -
# the fastest versions). Missing ones fall back to the Price Check readers.
SLA_FIRST = {"amazon", "flipkart"}


def sla_engine_dir():
    """..\\SLA TRACKER\\SLA Tracker v2\\_engine (or any 'SLA Tracker v2' next to the tracker)."""
    for pat in ("SLA Tracker v2", os.path.join("*", "SLA Tracker v2"), os.path.join("*", "*", "SLA Tracker v2")):
        for d in glob.glob(os.path.join(PS, pat)):
            e = os.path.join(d, "_engine", "channels")
            if os.path.isfile(os.path.join(e, "amazon.py")):
                return e
    return None


class _SlaCtx:
    def __init__(self, mod, key, log):
        self.rows, self._lock, self._log = [], threading.Lock(), log
        self.settings = dict(getattr(mod, "DEFAULTS", {}))
        self.settings["headless"] = True
        self.headless = True
        self.debug_dir = os.path.join(ROOT, "2_Output", "_Logs", "debug", f"sla_{key}")
        self.state_dir = os.path.join(HERE, "state", f"sla_{key}")
        os.makedirs(self.debug_dir, exist_ok=True)
        os.makedirs(self.state_dir, exist_ok=True)
        self.today = _dt.date.today()

    def emit(self, row):
        with self._lock:
            self.rows.append(row)

    def progress(self, n=1):
        pass

    def log(self, msg):
        with self._lock:
            self._log("  " + str(msg))

    def should_stop(self):
        return False


def _sla(key, ids, pin, log):
    d = sla_engine_dir()
    if not d:
        log(f"  SLA Tracker v2 not found next to the tracker - using the Price Check reader")
        return []
    spec = importlib.util.spec_from_file_location(f"sla_{key}", os.path.join(d, f"{key}.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    ctx = _SlaCtx(mod, key, log)
    jobs = [(mod.parse_id(i) or i, pin) for i in ids]
    log(f"  using the SLA Tracker reader ({len(jobs)} products, pincode {pin})")
    mod.run(jobs, ctx)
    back = {mod.parse_id(i) or i: i for i in ids}
    out = []
    for r in ctx.rows:
        sp = _f(r.get("Price (Rs)"))
        if sp is None:
            continue
        pid = back.get(str(r.get("PID")), str(r.get("PID")))
        out.append({"id": pid, "name": r.get("Product Name") or "", "brand": r.get("Brand") or "",
                    "size": r.get("Variant") or "", "mrp": _f(r.get("MRP (Rs)")), "sp": sp,
                    "in_stock": r.get("Status") != "Out of stock", "qty": None,
                    "seller": r.get("Seller") or "", "url": URL[key].format(pid)})
    return out


def fetch(links, log, log_fh, pin="110001"):
    """links = {channel_key: [ids]}  ->  list of results
    {"channel", "pin", "city", "listings", "note", "checked": True}"""
    _paths()
    results = []
    for key, ids in links.items():
        if not ids:
            continue
        nm = {c.key: c.name for c in ALL}[key]
        log(f"=== {nm}: {len(ids)} product(s) from Links.xlsx, all-India price ===")
        listings, note = [], ""
        try:
            if key == "nykaa":
                listings = _nykaa(ids, log, log_fh)
            elif key == "firstcry":
                listings = _firstcry(ids, log, log_fh)
            elif key in SLA_FIRST:
                try:
                    listings = _sla(key, ids, pin, log)
                except Exception as e:
                    log(f"  SLA reader failed ({str(e).splitlines()[0][:120]}) - using the Price Check reader")
                    listings = []
                got = {x["id"] for x in listings}
                left = [i for i in ids if i not in got]
                if left and len(left) > len(ids) * 0.1:
                    log(f"  {len(left)} without a price - reading them with the Price Check reader")
                    listings += _price_check(key, left, log, log_fh)
            else:
                listings = _price_check(key, ids, log, log_fh)
        except Exception as e:
            note = f"error: {str(e).splitlines()[0][:160]}"
            log(f"  {nm} failed: {note}")
        log(f"  {nm}: price read for {len(listings)} of {len(ids)}")
        if not listings and not note:
            note = "no prices read (see log)"
        results.append({"channel": key, "pin": NATIONAL_PIN, "city": NATIONAL_PIN,
                        "listings": listings, "note": note if not listings else "",
                        "checked": bool(listings)})
    return results


def _price_check(key, ids, log, log_fh):
    import price_check as PC
    PC._log_fh = log_fh
    fn = {"amazon": PC.amazon_prices, "flipkart": PC.flipkart_prices, "myntra": PC.myntra_prices}[key]
    got = fn(list(ids), _args())
    out = []
    for code, d in got.items():
        if d.get("missing") or d.get("sp") is None:
            continue
        out.append({"id": str(code), "name": d.get("name") or "", "brand": "Nat Habit", "size": "",
                    "mrp": d.get("mrp"), "sp": d.get("sp"), "in_stock": not d.get("oos"),
                    "qty": None, "seller": d.get("seller") or "", "url": URL[key].format(code)})
    return out


def _nykaa(ids, log, log_fh=None):
    import nykaa_scraper as NS
    NS._log_fh = log_fh
    import ny_extract as nx
    import ny_codes as NC
    args = _args()
    sess = NS.make_session()
    # Nykaa SKUs (NATHA...) -> product ids
    want, sku_of = [], {}
    skus = [i for i in ids if not str(i).isdigit()]
    if skus:
        res = NC.resolve_skus(skus, NC.Getter(NS), log)
        for s in skus:
            if res.get(s.upper()):
                sku_of[res[s.upper()]] = s
    for i in ids:
        want.append(next((p for p, s in sku_of.items() if s == i), i) if not str(i).isdigit() else str(i))
    rows = {}
    try:
        for _, raw in NS.scrape_listing(sess, NYKAA_BRAND_CATEGORY, args):
            r = nx.to_row(raw)
            if r["product_id"] in want:
                rows[r["product_id"]] = r
    except Exception as e:
        log(f"  Nykaa brand shelf not read: {e}")
    left = [p for p in want if p not in rows]
    if left:
        log(f"  Nykaa: {len(left)} id(s) not on the brand shelf - opening them one by one")
        for p in left:
            try:
                for _, raw in NS.scrape_product(sess, NC.product_url(p), args):
                    r = nx.to_row(raw)
                    rows[p] = r
            except Exception:
                pass
    out = []
    for p, r in rows.items():
        sp, mrp = _f(r.get("selling_price")), _f(r.get("mrp"))
        if sp is None:
            continue
        out.append({"id": sku_of.get(p, p), "alt": [p], "name": r.get("product_name") or "",
                    "brand": r.get("brand"), "size": r.get("pack_size"), "mrp": mrp, "sp": sp,
                    "in_stock": r.get("in_stock") != "FALSE", "qty": r.get("quantity") or None,
                    "seller": "", "url": r.get("product_url") or URL["nykaa"].format(p)})
    return out


def _firstcry(ids, log, log_fh=None):
    import firstcry_scraper as FS
    FS._log_fh = log_fh
    import fc_extract as fx
    args = _args()
    sess = FS.make_session("https://www.firstcry.com/")
    got = FS.fetch_products(sess, [str(i) for i in ids], args)
    out = []
    for pid, raw in got.items():
        r = fx.to_row(raw)
        sp = _f(r.get("selling_price"))
        if sp is None:
            continue
        out.append({"id": str(pid), "name": r.get("product_name") or "", "brand": r.get("brand"),
                    "size": r.get("pack_size"), "mrp": _f(r.get("mrp")), "sp": sp,
                    "in_stock": r.get("in_stock") != "FALSE",
                    "qty": r.get("stock_qty") or None, "seller": "",
                    "url": r.get("product_url") or URL["firstcry"].format(pid)})
    return out
