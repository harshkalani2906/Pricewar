"""
The 4 input files of the Price Parity Tracker (same idea as the SLA tracker).

  1_Input\\Links.xlsx           Channel Name | Channel ID
        which products to check. Channel ID can be the id or the full product link.
  1_Input\\Product Master.xlsx  Channel Name | Channel ID | Product Name | NH SKU
        Channel ID -> NH SKU, for every channel.
  1_Input\\Pincodes.xlsx        Channel Name | Pincode
        quick-commerce pincodes. Channel Name = one app, or "All QC" for all 5 apps.
  1_Input\\Pincode Master.xlsx  Pincode | Tier | City | State | Region
        city names for the pincodes.

Channel Name is one of:
  Amazon, Flipkart, Myntra, Nykaa, FirstCry                 (one all-India price)
  Blinkit, Zepto, Instamart, BigBasket, FK Minutes          (pincode by pincode)
"""
from __future__ import annotations

import os
import re

QC_KEYS = ["blinkit", "zepto", "instamart", "bigbasket", "fkminutes"]
MP_KEYS = ["amazon", "flipkart", "myntra", "nykaa", "firstcry"]
ALL_KEYS = QC_KEYS + MP_KEYS
NAMES = {"blinkit": "Blinkit", "zepto": "Zepto", "instamart": "Instamart", "bigbasket": "BigBasket",
         "fkminutes": "FK Minutes", "amazon": "Amazon", "flipkart": "Flipkart", "myntra": "Myntra",
         "nykaa": "Nykaa", "firstcry": "FirstCry"}

_ALIAS = {"amz": "amazon", "amazonin": "amazon", "fk": "flipkart", "flipkartminutes": "fkminutes",
          "fkminutes": "fkminutes", "minutes": "fkminutes", "fkm": "fkminutes",
          "swiggy": "instamart", "swiggyinstamart": "instamart", "im": "instamart",
          "bb": "bigbasket", "bbnow": "bigbasket", "bigbasketnow": "bigbasket",
          "zeptonow": "zepto", "fc": "firstcry", "first": "firstcry", "firstcrycom": "firstcry"}
_ALL_QC = {"all", "allqc", "qc", "quickcommerce", "allapps", "every", "allchannels"}


def channel_key(text) -> str | None:
    """'FK Minutes' -> 'fkminutes', 'All QC' -> 'allqc', unknown -> None"""
    s = re.sub(r"[^a-z]", "", str(text or "").lower())
    if not s:
        return None
    if s in _ALL_QC:
        return "allqc"
    if s in NAMES:
        return s
    for k, n in NAMES.items():
        if s == re.sub(r"[^a-z]", "", n.lower()):
            return k
    return _ALIAS.get(s)


# ---------------------------------------------------------------------------
# Channel ID from an id or a product link
# ---------------------------------------------------------------------------
_PAT = {
    "amazon":    [r"/(?:dp|gp/product|gp/aw/d|product)/([A-Z0-9]{10})", r"^([A-Z0-9]{10})$"],
    "flipkart":  [r"[?&]pid=([A-Z0-9]{16})", r"^([A-Z0-9]{16})$"],
    "fkminutes": [r"[?&]pid=([A-Z0-9]{16})", r"^([A-Z0-9]{16})$"],
    "myntra":    [r"/(\d{5,10})/buy", r"myntra\.com/(?:.*/)?(\d{5,10})/?(?:\?|$)", r"^(\d{5,10})$"],
    "nykaa":     [r"/p/(\d{4,10})", r"[?&]productId=(\d+)", r"^(\d{4,10})$", r"^([A-Z]{3,}[A-Z0-9]*\d{3,})$"],
    "firstcry":  [r"/(\d{5,10})/product-detail", r"[?&]pid=(\d{5,10})", r"^(?:FCID)?\s*(\d{5,10})$"],
    "blinkit":   [r"/prid/(\d+)", r"^(\d{3,10})$"],
    "zepto":     [r"/pvid/([0-9a-fA-F-]{36})", r"^([0-9a-fA-F-]{36})$"],
    "instamart": [r"/item/([A-Za-z0-9]+)", r"^([A-Za-z0-9]{4,})$"],
    "bigbasket": [r"/pd/(\d+)", r"^(\d{3,12})$"],
}
_UPPER = {"amazon", "flipkart", "fkminutes", "nykaa", "instamart"}


def clean_id(key: str, value) -> str | None:
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    v = str(value).strip().strip('"').strip("'")
    if not v or v.lower() in ("x", "-", "na", "n/a", "none", "nil", "not listed"):
        return None
    if v.endswith(".0") and v[:-2].isdigit():
        v = v[:-2]
    for p in _PAT.get(key, []):
        m = re.search(p, v, re.I)
        if m:
            got = m.group(1)
            if key in _UPPER and key != "instamart":
                return got.upper()
            return got.lower() if key == "zepto" else got
    return None


def _rows(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = [r for r in ws.iter_rows(values_only=True)]
    wb.close()
    # header = first row that has a "Channel" or "Pincode" cell
    for i, r in enumerate(rows[:10]):
        if any(re.sub(r"[^a-z]", "", str(c or "").lower()) in ("channelname", "channel", "pincode")
               for c in r):
            return r, rows[i + 1:]
    return (rows[0] if rows else ()), rows[1:]


def _h(head):
    return {re.sub(r"[^a-z]", "", str(c or "").lower()): i for i, c in enumerate(head)}


def _get(r, i):
    return r[i] if (i is not None and i < len(r)) else None


def nh_sku(v) -> str:
    """The team's NH SKU. Quick-commerce listings carry a 'BL-' prefix in the company
    master (same product, separate code) - the tracker always shows the plain NH SKU."""
    s = str(v or "").strip()
    return s[3:] if s.upper().startswith("BL-") else s


def read_priority(path):
    """{NH SKU: tier} from 1_Input/SKU Priority.xlsx  (tier = 'Top 50' / '80% revenue' / 'Rest')."""
    out = {}
    if not os.path.exists(path):
        return out
    head, rows = _rows_any(path)
    h = _h(head)
    si, ti = h.get("nhsku"), h.get("tier")
    if si is None or ti is None:
        return out
    for r in rows:
        if r and _get(r, si):
            out[nh_sku(_get(r, si))] = str(_get(r, ti) or "").strip()
    return out


def _rows_any(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = list(wb.worksheets[0].iter_rows(values_only=True))
    wb.close()
    return (rows[0] if rows else ()), rows[1:]


# ---------------------------------------------------------------------------
def read_links(path, log, problems):
    """{channel_key: [ids in order]}"""
    out = {k: [] for k in ALL_KEYS}
    if not os.path.exists(path):
        log(f"Missing {path}")
        return out
    head, rows = _rows(path)
    h = _h(head)
    ci, ii = h.get("channelname", h.get("channel", 0)), h.get("channelid", h.get("id", h.get("link", 1)))
    for n, r in enumerate(rows, 2):
        if not r or all(c in (None, "") for c in r):
            continue
        key = channel_key(_get(r, ci))
        raw = _get(r, ii)
        if key not in NAMES:
            problems.append(("Links.xlsx", n, f"Unknown Channel Name '{_get(r, ci)}'", raw))
            continue
        cid = clean_id(key, raw)
        if not cid:
            if raw not in (None, ""):
                problems.append(("Links.xlsx", n, f"Could not read a {NAMES[key]} id from this", raw))
            continue
        if cid not in out[key]:
            out[key].append(cid)
    return out


def read_master(path, log, problems):
    """{(channel_key, id): {"sku", "name"}}"""
    m = {}
    if not os.path.exists(path):
        log(f"Missing {path}")
        return m
    head, rows = _rows(path)
    h = _h(head)
    ci, ii = h.get("channelname", h.get("channel", 0)), h.get("channelid", h.get("id", 1))
    si, ni = h.get("nhsku"), h.get("productname")
    bi = h.get("brand")
    for n, r in enumerate(rows, 2):
        if not r or all(c in (None, "") for c in r):
            continue
        key = channel_key(_get(r, ci))
        if key not in NAMES:
            continue
        brand = str(_get(r, bi) or "").strip().lower()
        if brand and "nat" not in brand:
            continue                       # competitor rows (if any) are not priced here
        cid = clean_id(key, _get(r, ii))
        sku = nh_sku(_get(r, si))
        if not cid or not sku:
            continue
        m.setdefault((key, cid), {"sku": sku, "name": str(_get(r, ni) or "").strip()})
    return m


def read_pincodes(path, master_path, log, problems):
    """({channel_key: [pins]}, {pin: city})"""
    jobs = {k: [] for k in QC_KEYS}
    cities = {}
    if os.path.exists(master_path):
        head, rows = _rows(master_path)
        h = _h(head)
        pi, ci = h.get("pincode", 0), h.get("city")
        for r in rows:
            p = re.sub(r"\D", "", str(_get(r, pi) or "").split(".")[0])
            if len(p) == 6 and _get(r, ci):
                cities[p] = str(_get(r, ci)).strip()
    if not os.path.exists(path):
        log(f"Missing {path}")
        return jobs, cities
    head, rows = _rows(path)
    h = _h(head)
    chi, pi, cyi = h.get("channelname", h.get("channel", 0)), h.get("pincode", 1), h.get("city")
    for n, r in enumerate(rows, 2):
        if not r or all(c in (None, "") for c in r):
            continue
        p = re.sub(r"\D", "", str(_get(r, pi) or "").split(".")[0])
        if len(p) != 6:
            problems.append(("Pincodes.xlsx", n, "Pincode is not 6 digits", _get(r, pi)))
            continue
        key = channel_key(_get(r, chi)) or "allqc"
        if key in MP_KEYS:
            continue                       # marketplaces use one all-India price
        if key not in jobs and key != "allqc":
            problems.append(("Pincodes.xlsx", n, f"Unknown Channel Name '{_get(r, chi)}'", p))
            continue
        if _get(r, cyi):
            cities[p] = str(_get(r, cyi)).strip()
        for k in (QC_KEYS if key == "allqc" else [key]):
            if p not in jobs[k]:
                jobs[k].append(p)
    return jobs, cities
