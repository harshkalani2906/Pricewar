"""
Quick-commerce channel adapters for the QC Price Tracker.

Every adapter does only two things, inside a real (Playwright) browser page:
  set_location(page, pin, geo)  -> True / False
  fetch(page, pin, geo, log)    -> list of listing dicts (Nat Habit only)

A listing dict:
  {"id", "name", "size", "brand", "mrp", "sp", "in_stock", "qty", "url"}

Prices are read from the same JSON the website itself loads, so they are
exactly what a shopper at that pincode sees (no login, no cart).
"""
from __future__ import annotations

import json
import re
import time

BRAND_WORDS = ("nat habit", "nathabit", "nat-habit")


def is_nh(brand: str | None, name: str | None) -> bool:
    s = f"{brand or ''} {name or ''}".lower()
    return any(w in s for w in BRAND_WORDS)


def _num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.]", "", str(v))
    try:
        return float(s) if s else None
    except ValueError:
        return None


def _walk(o, test, out, depth=0):
    """Collect every dict inside o for which test(dict) is True."""
    if depth > 45:
        return
    if isinstance(o, dict):
        if test(o):
            out.append(o)
            return
        for v in o.values():
            _walk(v, test, out, depth + 1)
    elif isinstance(o, list):
        for v in o:
            _walk(v, test, out, depth + 1)


def all_ids(*objs, depth=2):
    """Every id-like value (keys ending in 'id', sku, slug numbers) inside the given dicts,
    so a product can be matched whichever id the user copied (app link, seller panel, etc.)."""
    out = []

    def walk(o, d):
        if not isinstance(o, dict) or d < 0:
            return
        for k, v in o.items():
            kl = str(k).lower()
            if isinstance(v, (str, int)) and not isinstance(v, bool) and (
                    kl.endswith("id") or kl.endswith("_id") or kl in ("sku", "skuid", "sku_id", "fsn", "pid")):
                sv = str(v).strip()
                if 3 <= len(sv) <= 64 and sv not in out:
                    out.append(sv)
            elif isinstance(v, dict):
                walk(v, d - 1)
    for o in objs:
        walk(o, depth)
    return out


class Capture:
    """Keeps the JSON bodies of responses whose URL contains a pattern."""

    def __init__(self, page, patterns):
        self.patterns = patterns
        self.bodies = []
        page.on("response", self._on)

    def _on(self, resp):
        try:
            u = resp.url
            if any(p in u for p in self.patterns) and resp.status == 200:
                t = resp.text()
                if t and t.lstrip()[:1] in "{[":
                    self.bodies.append(json.loads(t))
        except Exception:
            pass

    def reset(self):
        self.bodies = []


def _scroll(page, times=6, pause=1.2):
    for _ in range(times):
        try:
            page.mouse.wheel(0, 2500)
        except Exception:
            pass
        time.sleep(pause)


def _click_first_visible(page, selector, timeout=8000):
    loc = page.locator(selector)
    end = time.time() + timeout / 1000
    while time.time() < end:
        n = loc.count()
        for i in range(n):
            el = loc.nth(i)
            try:
                if el.is_visible():
                    el.click(timeout=3000)
                    return True
            except Exception:
                pass
        time.sleep(0.4)
    return False


# ---------------------------------------------------------------------------
# Pincode -> latitude / longitude (via Blinkit's own location search)
# ---------------------------------------------------------------------------
GEO_JS = r"""
async (pin) => {
  const s = await (await fetch('/location/autoSuggest?query=' + pin + '&lat=28.61&lng=77.20&session_token=')).json();
  const list = (s.ui_data && s.ui_data.suggestions) || [];
  if (!list.length) return null;
  // prefer a suggestion whose text actually contains the pincode
  let pick = list.find(x => JSON.stringify(x).includes(pin)) || list[0];
  const m = pick.meta || {};
  const t = encodeURIComponent((pick.title && pick.title.text) || pin);
  const d = encodeURIComponent((pick.subtitle && pick.subtitle.text) || pin);
  const info = await (await fetch('/location/info?place_id=' + m.place_id + '&title=' + t + '&description=' + d +
        '&is_pin_moved=false&session_token=' + (m.session_token || ''))).json();
  const c = info.coordinate || {};
  const li = info.location_info || {};
  return {lat: c.lat, lon: c.lon, city: li.city || info.city || '', state: li.state || '',
          area: (pick.subtitle && pick.subtitle.text) || '', serviceable: !!info.is_serviceable};
}
"""


def resolve_geo(page, pin, log):
    """Returns {"lat","lon","city","state","area"} or None. page must be on blinkit.com."""
    try:
        g = page.evaluate(GEO_JS, str(pin))
        if g and g.get("lat"):
            return g
    except Exception as e:
        log(f"  geo lookup failed for {pin}: {e}")
    return None


# ---------------------------------------------------------------------------
# Blinkit  - direct call to its search API with lat/lon headers
# ---------------------------------------------------------------------------
class Blinkit:
    key, name = "blinkit", "Blinkit"
    home = "https://blinkit.com/"

    SEARCH_JS = r"""
    async ([lat, lon, q]) => {
      const H = {'content-type':'application/json','lat':String(lat),'lon':String(lon),'app_client':'consumer_web'};
      let url = '/v1/layout/search?q=' + encodeURIComponent(q) + '&search_type=type_to_search';
      const pages = [];
      for (let i = 0; i < 8 && url; i++) {
        const r = await fetch(url, {method:'POST', headers:H, body:'{}'});
        if (!r.ok) { pages.push({__status: r.status}); break; }
        const j = await r.json();
        pages.push(j);
        const nx = j.response && j.response.pagination && j.response.pagination.next_url;
        url = nx || null;
        await new Promise(r => setTimeout(r, 700));
      }
      return pages;
    }
    """

    def set_location(self, page, pin, geo):
        if not geo:
            return False
        try:
            if "blinkit.com" not in page.url:
                page.goto(self.home, wait_until="domcontentloaded", timeout=45000)
            ctx = page.context
            ctx.add_cookies([
                {"name": "gr_1_lat", "value": str(geo["lat"]), "domain": "blinkit.com", "path": "/"},
                {"name": "gr_1_lon", "value": str(geo["lon"]), "domain": "blinkit.com", "path": "/"},
            ])
            return True
        except Exception:
            return False

    def fetch(self, page, pin, geo, log):
        pages = page.evaluate(self.SEARCH_JS, [geo["lat"], geo["lon"], "nat habit"])
        found = []
        for j in pages:
            if "__status" in j:
                log(f"  Blinkit search returned HTTP {j['__status']}")
                continue
            cards = []
            _walk(j, lambda d: isinstance(d.get("atc_action"), dict) and "product_id" in d, cards)
            for c in cards:
                item = (((c.get("atc_action") or {}).get("add_to_cart") or {}).get("cart_item") or {})
                name = item.get("product_name") or ((c.get("name") or {}).get("text"))
                brand = item.get("brand") or c.get("brand_name")
                if not is_nh(brand, name):
                    continue
                pid = str(c.get("product_id") or item.get("product_id"))
                sold_out = bool(c.get("is_sold_out")) or (c.get("product_state") not in (None, "available"))
                found.append({
                    "id": pid, "name": name, "brand": brand or "Nat Habit",
                    "alt": all_ids(c, item),
                    "size": item.get("unit") or ((c.get("variant") or {}).get("text")),
                    "mrp": _num(item.get("mrp")), "sp": _num(item.get("price")),
                    "in_stock": not sold_out, "qty": item.get("inventory"),
                    "url": f"https://blinkit.com/prn/x/prid/{pid}",
                })
        return found


# ---------------------------------------------------------------------------
# Instamart (Swiggy) - location is a cookie; prices come from the search page
# ---------------------------------------------------------------------------
class Instamart:
    key, name = "instamart", "Instamart"
    home = "https://www.swiggy.com/instamart"
    search = "https://www.swiggy.com/instamart/search?custom_back=true&query=nat%20habit"

    def __init__(self):
        self.cap = None

    def set_location(self, page, pin, geo):
        if not geo:
            return False
        if self.cap is None:
            self.cap = Capture(page, ["/api/instamart/search"])
        loc = {"lat": geo["lat"], "lng": geo["lon"], "address": geo.get("area") or str(pin),
               "area": "", "showUserDefaultAddressHint": False}
        page.context.add_cookies([{
            "name": "userLocation", "value": json.dumps(loc, separators=(",", ":")),
            "domain": ".swiggy.com", "path": "/"}])
        return True

    def fetch(self, page, pin, geo, log):
        self.cap.reset()
        page.goto(self.search, wait_until="domcontentloaded", timeout=60000)
        time.sleep(5)
        _scroll(page, 5, 1.5)
        found = {}
        for j in self.cap.bodies:
            if isinstance(j, dict) and j.get("statusCode") == 429:
                log("  Instamart is rate-limiting (429) - waiting longer next time")
            groups = []
            _walk(j, lambda d: isinstance(d.get("variations"), list) and "displayName" in d, groups)
            for g in groups:
                for v in g["variations"]:
                    name = v.get("displayName") or g.get("displayName")
                    brand = v.get("brandName") or g.get("brand")
                    if not is_nh(brand, name):
                        continue
                    pr = v.get("price") or {}
                    sid = str(v.get("spinId") or v.get("skuId"))
                    inv = v.get("inventory") or {}
                    found[sid] = {
                        "id": sid, "name": name, "brand": brand,
                        "alt": all_ids(v, g, depth=1),
                        "size": v.get("quantityDescription"),
                        "mrp": _num((pr.get("mrp") or {}).get("units")),
                        "sp": _num((pr.get("offerPrice") or {}).get("units")),
                        "in_stock": bool(inv.get("inStock", g.get("inStock", True))),
                        "qty": (v.get("cartAllowedQuantity") or {}).get("allowedQuantity"),
                        "url": f"https://www.swiggy.com/instamart/item/{v.get('productId') or g.get('productId') or ''}",
                    }
        return list(found.values())


# ---------------------------------------------------------------------------
# Zepto - location via its own address picker; prices from the search API
# ---------------------------------------------------------------------------
class Zepto:
    key, name = "zepto", "Zepto"
    home = "https://www.zepto.com/"
    search = "https://www.zepto.com/search?query=nat%20habit"

    def __init__(self):
        self.cap = None

    def set_location(self, page, pin, geo):
        if self.cap is None:
            self.cap = Capture(page, ["search-federator-service/api/v1/search"])
        try:
            if "zepto.com/search" not in page.url:
                page.goto(self.search, wait_until="domcontentloaded", timeout=60000)
                time.sleep(3)
            if not _click_first_visible(page, "button[aria-label='Select Location'], [data-testid='user-address']"):
                return False
            inp = page.locator("input[placeholder*='Search a new address']").first
            inp.wait_for(state="visible", timeout=10000)
            inp.fill("")
            inp.type(str(pin), delay=80)
            time.sleep(2.5)
            if not _click_first_visible(page, "[data-testid='address-search-item']", 10000):
                return False
            time.sleep(4)
            return True
        except Exception:
            return False

    def store_id(self, page):
        try:
            for c in page.context.cookies("https://www.zepto.com"):
                if c["name"] == "serviceability":
                    m = re.search(r'storeId%22%3A%22([0-9a-f-]+)|storeId":"([0-9a-f-]+)', c["value"])
                    if m:
                        return m.group(1) or m.group(2)
        except Exception:
            pass
        return None

    def fetch(self, page, pin, geo, log):
        self.cap.reset()
        page.goto(self.search, wait_until="domcontentloaded", timeout=60000)
        time.sleep(5)
        _scroll(page, 8, 1.4)
        found = {}
        for j in self.cap.bodies:
            items = []
            _walk(j, lambda d: "discountedSellingPrice" in d and isinstance(d.get("product"), dict), items)
            for it in items:
                p = it["product"]
                pv = it.get("productVariant") or {}
                if not is_nh(p.get("brand"), p.get("name")):
                    continue
                pvid = str(pv.get("id") or it.get("id"))
                found[pvid] = {
                    "id": pvid, "name": p.get("name"), "brand": p.get("brand"),
                    "alt": all_ids(it, p, pv, depth=1),
                    "size": pv.get("formattedPacksize"),
                    "mrp": (_num(it.get("mrp")) or 0) / 100 or None,
                    "sp": (_num(it.get("discountedSellingPrice")) or 0) / 100 or None,
                    "in_stock": not it.get("outOfStock", False),
                    "qty": it.get("availableQuantity"),
                    "url": f"https://www.zepto.com/pn/x/pvid/{pvid}",
                }
        return list(found.values())


# ---------------------------------------------------------------------------
# BigBasket - location via its picker; prices from the Nat Habit brand page
# ---------------------------------------------------------------------------
class BigBasket:
    key, name = "bigbasket", "BigBasket"
    home = "https://www.bigbasket.com/"
    brand_page = "https://www.bigbasket.com/pb/nat-habit/"

    def __init__(self):
        self.cap = None

    def set_location(self, page, pin, geo):
        if self.cap is None:
            self.cap = Capture(page, ["/listing-svc/v2/products"])
        try:
            if "bigbasket.com" not in page.url:
                page.goto(self.home, wait_until="domcontentloaded", timeout=60000)
                time.sleep(3)
            opener = "text=/Select Location|Delivery in/i"
            if not _click_first_visible(page, opener):
                return False
            inp = page.locator("input[placeholder*='Search for area']")
            end = time.time() + 8
            target = None
            while time.time() < end and target is None:
                for i in range(inp.count()):
                    if inp.nth(i).is_visible():
                        target = inp.nth(i)
                        break
                time.sleep(0.3)
            if target is None:
                return False
            target.click()
            target.fill("")
            target.type(str(pin), delay=80)
            time.sleep(2.5)
            if not _click_first_visible(page, f"li:has-text('{pin}')", 10000):
                return False
            time.sleep(5)
            for c in page.context.cookies("https://www.bigbasket.com"):
                if c["name"] == "_bb_pin_code":
                    return c["value"].strip('"') == str(pin) or True
            return True
        except Exception:
            return False

    def fetch(self, page, pin, geo, log):
        self.cap.reset()
        page.goto(self.brand_page, wait_until="domcontentloaded", timeout=60000)
        time.sleep(4)
        _scroll(page, 6, 1.3)
        found = {}
        for j in self.cap.bodies:
            prods = []
            _walk(j, lambda d: "pricing" in d and "desc" in d and "id" in d, prods)
            for p in prods:
                brand = (p.get("brand") or {}).get("name") if isinstance(p.get("brand"), dict) else p.get("brand")
                name = f"{brand or 'Nat Habit'} {p.get('desc')}"
                if not is_nh(brand, name):
                    continue
                disc = ((p.get("pricing") or {}).get("discount") or {})
                av = p.get("availability") or {}
                pid = str(p.get("id"))
                found[pid] = {
                    "id": pid, "name": name, "brand": brand, "alt": all_ids(p, depth=1),
                    "size": " ".join(x for x in [p.get("w"), p.get("pack_desc")] if x),
                    "mrp": _num(disc.get("mrp")),
                    "sp": _num((disc.get("prim_price") or {}).get("sp")),
                    "in_stock": str(av.get("avail_status", "001")) == "001" and not av.get("not_for_sale"),
                    "qty": None,
                    "url": "https://www.bigbasket.com" + (p.get("absolute_url") or f"/pd/{pid}/"),
                }
        return list(found.values())


# ---------------------------------------------------------------------------
# Flipkart Minutes (BETA) - location via Flipkart's hyperlocal address sheet
# ---------------------------------------------------------------------------
class FlipkartMinutes:
    key, name = "fkminutes", "FK Minutes"
    search = "https://www.flipkart.com/search?q=nat%20habit&marketplace=HYPERLOCAL"

    def __init__(self):
        self.cap = None

    def set_location(self, page, pin, geo):
        if self.cap is None:
            self.cap = Capture(page, ["rome.api.flipkart.com/api/4/page/fetch", "/api/4/page/fetch"])
        try:
            page.goto(self.search, wait_until="domcontentloaded", timeout=60000)
            time.sleep(5)
            # open the address sheet if it is not already open
            inp = page.locator("input[placeholder*='pin code'], input[placeholder*='pincode']")
            if not (inp.count() and inp.first.is_visible()):
                _click_first_visible(page, "text=/Deliver to|Change|Select delivery address|Delivering to/i", 5000)
                time.sleep(2)
            inp = page.locator("input[placeholder*='pin code'], input[placeholder*='pincode']").first
            inp.wait_for(state="visible", timeout=10000)
            inp.click()
            inp.fill("")
            inp.type(str(pin), delay=90)
            time.sleep(3)
            if not _click_first_visible(page, f"div:has-text('{pin}') >> nth=-1", 8000):
                return False
            time.sleep(5)
            # map page: confirm the pin
            _click_first_visible(page, "text=/Confirm|Save and proceed|Proceed/i", 6000)
            time.sleep(4)
            return "hyperlocal-preview" not in page.url and "address-map" not in page.url
        except Exception:
            return False

    def fetch(self, page, pin, geo, log):
        self.cap.reset()
        page.goto(self.search, wait_until="domcontentloaded", timeout=60000)
        time.sleep(5)
        _scroll(page, 6, 1.4)
        found = {}
        # first, products embedded in the page itself (server-rendered state)
        try:
            st = page.evaluate("() => window.__INITIAL_STATE__ || null")
            if st:
                self.cap.bodies.append(st)
        except Exception:
            pass
        for j in self.cap.bodies:
            vals = []
            _walk(j, lambda d: isinstance(d.get("pricing"), dict) and isinstance(d.get("titles"), dict), vals)
            for v in vals:
                t = v.get("titles") or {}
                name = t.get("title") or t.get("newTitle")
                brand = v.get("productBrand") or (name or "").split(" ")[0]
                if not is_nh(brand, name):
                    continue
                pr = v.get("pricing") or {}
                fsn = str(v.get("id") or v.get("itemId"))
                av = (v.get("availability") or {}).get("displayState", "")
                found[fsn] = {
                    "id": fsn, "name": name, "brand": brand, "alt": all_ids(v, depth=1),
                    "size": t.get("subtitle"),
                    "mrp": _num((pr.get("mrp") or {}).get("value")),
                    "sp": _num((pr.get("finalPrice") or {}).get("value")),
                    "in_stock": str(av).lower() in ("", "in_stock", "available"),
                    "qty": None,
                    "url": "https://www.flipkart.com" + (v.get("baseUrl") or v.get("smartUrl") or ""),
                }
        return list(found.values())


ALL = [Blinkit, Zepto, Instamart, BigBasket, FlipkartMinutes]
BY_KEY = {c.key: c for c in ALL}
BY_NAME = {c.name.lower().replace(" ", ""): c for c in ALL}
ALIASES = {"swiggy": "instamart", "swiggyinstamart": "instamart", "bb": "bigbasket",
           "bbnow": "bigbasket", "fkminutes": "fkminutes", "flipkartminutes": "fkminutes",
           "minutes": "fkminutes", "zeptonow": "zepto"}


def channel_key(text: str | None) -> str | None:
    """'Blinkit' / 'FK Minutes' / 'All' -> internal key ('all' for All)."""
    if not text:
        return "all"
    s = re.sub(r"[^a-z]", "", str(text).lower())
    if s in ("all", "", "allchannels", "every"):
        return "all"
    if s in ("amazon", "amz", "flipkart", "fk", "myntra", "nykaa"):
        return {"amz": "amazon", "fk": "flipkart"}.get(s, s)
    if s in BY_KEY:
        return s
    if s in BY_NAME:
        return BY_NAME[s].key
    return ALIASES.get(s)
