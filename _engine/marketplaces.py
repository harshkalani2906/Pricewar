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


class AmazonNow:                      # Amazon's quick commerce - read city by city (QC pincodes)
    key, name, national = "amazonnow", "Amazon Now", False


ALL = [Amazon, Flipkart, Myntra, Nykaa, FirstCry, AmazonNow]

# Amazon prices are read ONLY from Nat Habit's own sellers (not the buy box / other sellers):
AMAZON_SELLER = "RK World"            # Amazon        -> price of this seller
AMAZON_NOW_SELLER = "RK World Now"    # Amazon Now    -> price of this seller
NEEDS = {"amazon": ["Price Check", "Amazon"], "flipkart": ["Price Check", "Flipkart"],
         "myntra": ["Price Check", "Myntra"], "nykaa": ["Nykaa"], "firstcry": ["FirstCry"]}
URL = {"amazon": "https://www.amazon.in/dp/{}", "amazonnow": "https://www.amazon.in/dp/{}", "flipkart": "https://www.flipkart.com/product/p/itme?pid={}",
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
SLA_FIRST = {"flipkart"}


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


def fetch(links, log, log_fh, pin="110001", now_pins=None, cities=None):
    """links = {channel_key: [ids]}  ->  list of results
    {"channel", "pin", "city", "listings", "note", "checked": True}
    Amazon also gives Amazon Now (seller RK World Now) at every quick-commerce pincode."""
    _paths()
    results = []
    for key, ids in links.items():
        if not ids or key == "amazonnow":
            continue
        if key == "amazon":
            try:
                results += _amazon(ids, now_pins or [pin], cities or {}, log)
            except Exception as e:
                log(f"  Amazon failed: {str(e).splitlines()[0][:160]}")
                results.append({"channel": "amazon", "pin": NATIONAL_PIN, "city": NATIONAL_PIN, "listings": [],
                                "note": f"error: {str(e).splitlines()[0][:160]}", "checked": False})
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


# ---------------------------------------------------------------------------
# Amazon + Amazon Now: every seller's offer on the product ("All offers" panel),
# keep only RK World (Amazon) and RK World Now (Amazon Now).
# ---------------------------------------------------------------------------
def _norm(s):
    import re
    return re.sub(r"[^a-z]", "", str(s or "").lower())


def _is_now(seller):
    n = _norm(seller)
    return n == _norm(AMAZON_NOW_SELLER) or (n.startswith(_norm(AMAZON_SELLER)) and "now" in n)


def _is_main(seller):
    return _norm(seller) == _norm(AMAZON_SELLER)


def parse_offers(html):
    """-> (title, [{"seller", "price", "pinned"}]) from Amazon's All-offers (AOD) panel.
    pinned = the buy-box offer (what Amazon shows by default)."""
    import re
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html or "", "html.parser")
    t = soup.select_one("#aod-asin-title-text")
    title = " ".join(t.get_text(" ").split()) if t else ""
    out = []
    for blk in soup.select("#aod-pinned-offer, #aod-offer"):
        se = (blk.select_one("#aod-offer-soldBy a") or blk.select_one("#aod-offer-soldBy .a-col-right .a-size-small")
              or blk.select_one("#aod-offer-soldBy .a-col-right"))
        seller = " ".join(se.get_text(" ").split()) if se else ""
        seller = re.sub(r"^Sold by\s*", "", seller, flags=re.I).split(" Seller rating")[0].strip()
        pe = blk.select_one(".a-price:not(.a-text-price) .a-offscreen") or blk.select_one(".a-price .a-offscreen")
        txt = pe.get_text() if pe else ""
        if not txt.strip():
            w = blk.select_one(".a-price-whole")
            txt = w.get_text() if w else ""
        m = re.search(r"[\d,]+(?:\.\d+)?", txt)
        price = float(m.group(0).replace(",", "")) if m else None
        if seller and price:
            out.append({"seller": seller, "price": price, "pinned": blk.get("id") == "aod-pinned-offer"})
    return title, out


# same reading, done inside the Amazon page (used when Amazon refuses the direct method)
JS_OFFERS = r"""
async ([asins, par]) => {
  const one = async (a) => {
    const offers = []; let title = '';
    for (let pg = 1; pg <= 3; pg++) {
      try {
        const r = await fetch('/gp/product/ajax/aodAjaxMain/?asin=' + a + '&pc=dp&experienceId=aodAjaxMain' + (pg > 1 ? '&pageno=' + pg : ''));
        if (!r.ok) break;
        const html = await r.text();
        if (!/aod-offer|aod-pinned-offer/.test(html) && /captcha|characters you see/i.test(html)) return {cap: 1};
        const d = new DOMParser().parseFromString(html, 'text/html');
        const t = d.querySelector('#aod-asin-title-text');
        if (t && !title) title = t.textContent.replace(/\s+/g, ' ').trim();
        let n = 0;
        d.querySelectorAll('#aod-pinned-offer, #aod-offer').forEach(b => {
          const se = b.querySelector('#aod-offer-soldBy a') || b.querySelector('#aod-offer-soldBy .a-col-right .a-size-small')
                     || b.querySelector('#aod-offer-soldBy .a-col-right');
          let s = se ? se.textContent.replace(/\s+/g, ' ').trim() : '';
          s = s.replace(/^Sold by\s*/i, '').split(' Seller rating')[0].trim();
          const pe = b.querySelector('.a-price:not(.a-text-price) .a-offscreen') || b.querySelector('.a-price .a-offscreen');
          let tx = pe ? pe.textContent : '';
          if (!tx.trim()) { const w = b.querySelector('.a-price-whole'); tx = w ? w.textContent : ''; }
          const m = tx.match(/[\d,]+(?:\.\d+)?/);
          const p = m ? parseFloat(m[0].replace(/,/g, '')) : null;
          if (s && p) { offers.push({seller: s, price: p, pinned: b.id === 'aod-pinned-offer'}); n++; }
        });
        if (n < 10) break;
      } catch (e) { break; }
    }
    return {offers, title};
  };
  const out = new Array(asins.length); let i = 0;
  const w = async () => { while (i < asins.length) { const k = i++; out[k] = await one(asins[k]); } };
  await Promise.all(Array.from({length: Math.min(par, asins.length)}, w));
  return out;
}
"""


def _amazon(ids, pins, cities, log):
    """Amazon  : RK World's price. If RK World has no offer on a product -> the default
                 (buy box) price Amazon shows, so the cell is never blank.
       Amazon Now: RK World Now's price at each quick-commerce pincode."""
    d = sla_engine_dir()
    if not d:
        raise RuntimeError("SLA Tracker v2 not found next to the tracker (needed for Amazon)")
    spec = importlib.util.spec_from_file_location("sla_amazon", os.path.join(d, "amazon.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    ctx = _SlaCtx(mod, "amazon", log)
    asins = [mod.parse_id(i) or i for i in ids]
    back = dict(zip(asins, ids))
    pins = list(dict.fromkeys(pins))
    got, titles, lock = {}, {}, threading.Lock()

    def keep(pin, asin, offers, title=""):
        with lock:
            if offers or asin not in got.get(pin, {}):
                got.setdefault(pin, {})[asin] = offers or []
            if title:
                titles.setdefault(asin, title)

    # 1) direct method (fast)
    class Api(mod._ApiRunner):
        def read_one(self, sess, asin, pin):
            offers = []
            try:
                for page in (1, 2, 3):
                    code, txt = self.req(sess, "GET", f"/gp/product/ajax/aodAjaxMain/?asin={asin}&pc=dp"
                                         f"&experienceId=aodAjaxMain" + (f"&pageno={page}" if page > 1 else ""),
                                         headers={"Referer": f"{mod.BASE}/dp/{asin}"})
                    if not 200 <= code < 300:
                        break
                    title, part = parse_offers(txt)
                    keep(pin, asin, [], title)
                    offers += part
                    if len(part) < 10:
                        break
                return {"offers": offers}
            except mod._Blocked:
                return {"cap": 1}
            except Exception as e:
                self.note_error(e)
                return {"offers": offers}

        def emit_result(self, sess, asin, pin, r):
            keep(pin, asin, r.get("offers") or [])

    log(f"=== Amazon + Amazon Now: {len(asins)} product(s) x {len(pins)} pincode(s) - "
        f"sellers '{AMAZON_SELLER}' and '{AMAZON_NOW_SELLER}' ===")
    try:
        Api([(a, p) for p in pins for a in asins], ctx).main()
    except Exception as e:
        log(f"  Amazon direct method failed: {str(e).splitlines()[0][:150]}")

    # 2) whatever the direct method could not read -> inside a real browser
    left = [(a, p) for p in pins for a in asins if not (got.get(p) or {}).get(a)]
    if left:
        log(f"  Amazon: reading {len(left)} product x pincode check(s) in the browser")

        class Browser(mod._Runner):
            async def check_pin(self, page, pin, todo):
                ok = False
                for _ in range(2):
                    try:
                        ok = await self.safe_eval(page, mod.JS_SET_PIN, pin)
                    except Exception as e:
                        self.note_error(e)
                    if ok:
                        break
                    await self.open_home(page)
                if not ok:
                    return {a: {"offers": []} for a in todo}
                res = await self.safe_eval(page, JS_OFFERS, [todo, max(2, self.per_session)])
                return {a: (r or {}) for a, r in zip(todo, res or [])}

            def to_row(self, asin, pin, r):
                keep(pin, asin, r.get("offers") or [], r.get("title") or "")
                return {"PID": asin}

        try:
            mod._drive_browser(Browser(left, ctx, {a: {"name": "-", "brand": "-"} for a in asins}))
        except Exception as e:
            log(f"  Amazon browser reading failed: {str(e).splitlines()[0][:150]}")

    def listing(asin, o, seller, note=""):
        return {"id": back.get(asin, asin), "name": titles.get(asin, ""), "brand": "Nat Habit", "size": "",
                "mrp": None, "sp": o["price"], "in_stock": True, "qty": None, "seller": seller + note,
                "url": URL["amazon"].format(asin)}

    main, default, results, sellers = {}, {}, [], {}
    for p in pins:
        now = []
        for asin, offers in (got.get(p) or {}).items():
            for o in offers:
                sellers[o["seller"]] = sellers.get(o["seller"], 0) + 1
            m = [o for o in offers if _is_main(o["seller"])]
            if m and asin not in main:
                main[asin] = listing(asin, min(m, key=lambda o: o["price"]), AMAZON_SELLER)
            other = [o for o in offers if not _is_now(o["seller"])]
            if other and asin not in default:            # buy box first, else the cheapest offer
                o = next((x for x in other if x.get("pinned")), None) or min(other, key=lambda x: x["price"])
                default[asin] = listing(asin, o, o["seller"], f" (default - no {AMAZON_SELLER} offer)")
            n = [o for o in offers if _is_now(o["seller"])]
            if n:
                now.append(listing(asin, min(n, key=lambda o: o["price"]), AMAZON_NOW_SELLER))
        results.append({"channel": "amazonnow", "pin": p, "city": cities.get(p, ""), "listings": now,
                        "note": "" if got.get(p) else "Amazon did not answer for this pincode", "checked": bool(got.get(p))})
        log(f"  Amazon Now {p} {cities.get(p, '')}: {len(now)} product(s) sold by {AMAZON_NOW_SELLER}")

    amz = dict(main)
    for a, L in default.items():
        amz.setdefault(a, L)
    # 3) still nothing for a product -> the old buy-box reader (never leave Amazon blank)
    rest = [back.get(a, a) for a in asins if a not in amz]
    if rest:
        log(f"  Amazon: {len(rest)} product(s) without any offer read - using the default Amazon price reader")
        try:
            for L in _sla_amazon_default(mod, rest, pins[0], log) or _price_check("amazon", rest, log, None):
                L["seller"] = (L.get("seller") or "Amazon default") + f" (default - no {AMAZON_SELLER} offer read)"
                amz.setdefault(mod.parse_id(L["id"]) or L["id"], L)
        except Exception as e:
            log(f"  default Amazon reader failed: {str(e).splitlines()[0][:150]}")
    log(f"  Amazon: {len(main)} of {len(asins)} product(s) at {AMAZON_SELLER}'s price, "
        f"{len(amz) - len(main)} at the default price, {len(asins) - len(amz)} not read")
    top = sorted(sellers.items(), key=lambda x: -x[1])[:8]
    if top:
        log("  sellers seen on these products: " + ", ".join(f"{k} ({v})" for k, v in top))
    results.insert(0, {"channel": "amazon", "pin": NATIONAL_PIN, "city": NATIONAL_PIN,
                       "listings": list(amz.values()), "note": "" if amz else "no Amazon prices read",
                       "checked": bool(amz)})
    return results


def _sla_amazon_default(mod, ids, pin, log):
    """The SLA Tracker's normal Amazon reading (buy-box price, direct + browser fallback)."""
    ctx = _SlaCtx(mod, "amazon", log)
    jobs = [(mod.parse_id(i) or i, pin) for i in ids]
    mod.run(jobs, ctx)
    back = {mod.parse_id(i) or i: i for i in ids}
    out = []
    for r in ctx.rows:
        sp = _f(r.get("Price (Rs)"))
        if sp is None:
            continue
        pid = back.get(str(r.get("PID")), str(r.get("PID")))
        out.append({"id": pid, "name": r.get("Product Name") or "", "brand": r.get("Brand") or "",
                    "size": "", "mrp": _f(r.get("MRP (Rs)")), "sp": sp,
                    "in_stock": r.get("Status") != "Out of stock", "qty": None,
                    "seller": r.get("Seller") or "", "url": URL["amazon"].format(pid)})
    return out
