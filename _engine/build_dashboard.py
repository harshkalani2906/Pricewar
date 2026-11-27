"""
Simple price dashboard -> 3_Dashboard/price-pulse.html (one file, data inside).

One row per NH SKU, one column per channel, the price in each cell.
Cells whose price is not the aligned price are highlighted.
SKUs are sorted by how many channels are NOT at the aligned price (most first).
Download button saves the table as CSV (opens in Excel).

Uses the latest run of each channel from 2_Output/history/qc_price_history.csv.
"""
from __future__ import annotations

import base64
import csv
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
HIST_FILE = os.path.join(ROOT, "2_Output", "history", "qc_price_history.csv")
DASH_DIR = os.path.join(ROOT, "3_Dashboard")
TEMPLATE = os.path.join(HERE, "dashboard_template.html")

CHANNELS = ["Amazon", "Flipkart", "Myntra", "Nykaa", "FirstCry",
            "Blinkit", "Zepto", "Instamart", "BigBasket", "FK Minutes"]
LOGO_DIR = os.path.join(HERE, "logos")
# channel -> website whose icon is shown in the column header
LOGO_SITE = {"Amazon": "amazon.in", "Flipkart": "flipkart.com", "Myntra": "myntra.com",
             "Nykaa": "nykaa.com", "FirstCry": "firstcry.com", "Blinkit": "blinkit.com",
             "Zepto": "zepto.com", "Instamart": "swiggy.com", "BigBasket": "bigbasket.com",
             "FK Minutes": "flipkart.com"}
LOGO_URLS = ["https://www.google.com/s2/favicons?domain={d}&sz=128",
             "https://icons.duckduckgo.com/ip3/{d}.ico"]
PRICED = {"At par", "Below aligned", "Above aligned", "No aligned price", "Out of stock"}


def f(v):
    try:
        x = float(v)
        return int(x) if x.is_integer() else round(x, 2)
    except (TypeError, ValueError):
        return None


def logos(log=print):
    """{channel: data-URI}. Downloads each site's icon once into _engine/logos/."""
    os.makedirs(LOGO_DIR, exist_ok=True)
    out = {}
    for ch, site in LOGO_SITE.items():
        own = os.path.join(LOGO_DIR, ch.lower().replace(" ", "") + ".png")   # e.g. logos/fkminutes.png
        path = own if os.path.exists(own) else os.path.join(LOGO_DIR, site.split(".")[0] + ".png")
        if not os.path.exists(path) or os.path.getsize(path) < 200:
            for u in LOGO_URLS:
                try:
                    import requests
                    r = requests.get(u.format(d=site), timeout=15,
                                     headers={"User-Agent": "Mozilla/5.0"})
                    if r.status_code == 200 and len(r.content) > 200:
                        with open(path, "wb") as fh:
                            fh.write(r.content)
                        break
                except Exception:
                    pass
        if os.path.exists(path) and os.path.getsize(path) >= 200:
            with open(path, "rb") as fh:
                raw = fh.read()
            kind = "image/png" if raw[:4] == b"\x89PNG" else "image/x-icon"
            out[ch] = f"data:{kind};base64," + base64.b64encode(raw).decode()
        else:
            log(f"  logo for {ch} not downloaded - showing its name instead")
    return out


def aligned_both():
    """Event + BAU aligned selling price per NH SKU from 1_Input/Aligned Price.xlsx
    (columns: NH SKU | ... | Event Aligned SP ... | BAU Aligned SP ...).
    -> ({sku: event}, {sku: bau}, event header, bau header)"""
    path = os.path.join(ROOT, "1_Input", "Aligned Price.xlsx")
    ev, bau, el, bl = {}, {}, "", ""
    try:
        import openpyxl
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
        rows = ws.iter_rows(values_only=True)
        head = [str(h or "").strip() for h in next(rows)]
        low = [h.lower() for h in head]
        si = next(i for i, h in enumerate(low) if "sku" in h)
        ei = next((i for i, h in enumerate(low) if "event" in h), None)
        bi = next((i for i, h in enumerate(low) if "bau" in h), None)
        el = head[ei] if ei is not None else ""
        bl = head[bi] if bi is not None else ""
        for r in rows:
            if not r or si >= len(r) or not r[si]:
                continue
            sku = str(r[si]).strip()
            for i, dst in ((ei, ev), (bi, bau)):
                if i is not None and i < len(r) and f(r[i]) is not None:
                    dst[sku] = f(r[i])
    except (StopIteration, FileNotFoundError):
        pass
    except Exception:
        pass
    return ev, bau, el, bl


def priority():
    """{NH SKU: (tier, revenue rank)} from 1_Input/SKU Priority.xlsx."""
    out = {}
    try:
        import openpyxl
        ws = openpyxl.load_workbook(os.path.join(ROOT, "1_Input", "SKU Priority.xlsx"),
                                    read_only=True, data_only=True).worksheets[0]
        rows = ws.iter_rows(values_only=True)
        head = [str(h or "").strip().lower() for h in next(rows)]
        ri, si, ti = head.index("rank"), head.index("nh sku"), head.index("tier")
        for r in rows:
            if r and r[si]:
                sku = str(r[si]).strip()
                out[sku[3:] if sku.upper().startswith("BL-") else sku] = (str(r[ti] or ""), r[ri])
    except Exception:
        pass
    return out


def channel_ranks():
    """{NH SKU: {channel: sales rank in that channel}} from SKU Priority.xlsx, sheet 'Channel Rank'.
    Only the rank goes into the dashboard (no revenue figures)."""
    out = {}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(os.path.join(ROOT, "1_Input", "SKU Priority.xlsx"), read_only=True, data_only=True)
        if "Channel Rank" not in wb.sheetnames:
            return out
        rows = wb["Channel Rank"].iter_rows(values_only=True)
        head = [str(h or "").strip().lower() for h in next(rows)]
        ci, si = head.index("channel"), head.index("nh sku")
        ri = next(i for i, h in enumerate(head) if h.startswith("rank"))
        for r in rows:
            if r and r[si] and r[ri]:
                sku = str(r[si]).strip()
                sku = sku[3:] if sku.upper().startswith("BL-") else sku
                out.setdefault(sku, {})[str(r[ci]).strip()] = int(r[ri])
    except Exception:
        pass
    return out


def categories():
    """{NH SKU: Category} from 1_Input/Product Master.xlsx (Category column)."""
    import collections
    path = os.path.join(ROOT, "1_Input", "Product Master.xlsx")
    out = collections.defaultdict(collections.Counter)
    try:
        import openpyxl
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
        rows = ws.iter_rows(values_only=True)
        head = [str(h or "").strip().lower() for h in next(rows)]
        si, ci = head.index("nh sku"), head.index("category")
        for r in rows:
            if r and len(r) > max(si, ci) and r[si] and r[ci]:
                out[str(r[si]).strip()][str(r[ci]).strip()] += 1
    except Exception:
        pass
    return {k: v.most_common(1)[0][0] for k, v in out.items()}


def price_history(rows, cutoff_days=7):
    """{NH SKU: {channel: [[run time, lowest in-stock price, city], ...]}} - one point per run,
    only where the price changed (for the Culprit tab: who cut the price first, who followed)."""
    import datetime as _dt
    times = sorted({r["Run Time"] for r in rows if r["Run Time"]})
    if not times:
        return {}
    try:
        cut = (_dt.datetime.strptime(times[-1], "%Y-%m-%d %H:%M")
               - _dt.timedelta(days=cutoff_days)).strftime("%Y-%m-%d %H:%M")
    except ValueError:
        cut = ""
    best = {}
    for r in rows:
        if r["Run Time"] < cut or r["Status"] not in ("At par", "Below aligned", "Above aligned", "No aligned price"):
            continue
        sku = r["NH SKU"]
        sku = sku[3:] if sku.upper().startswith("BL-") else sku
        sp = f(r["Selling Price"])
        if not sku or sp is None:
            continue
        k = (sku, r["Channel"], r["Run Time"])
        city = "" if r["Pincode"] == "All India" else (r["City"] or r["Pincode"])
        if k not in best or sp < best[k][0]:
            best[k] = (sp, city)
    out = {}
    for (sku, ch, t), (p, city) in sorted(best.items(), key=lambda x: x[0][2]):
        pts = out.setdefault(sku, {}).setdefault(ch, [])
        if not pts or round(pts[-1][1]) != round(p):
            pts.append([t, p, city])
    return out


# product page of each channel, from the id (used when a reading has no link saved)
PAGE = {"Amazon": "https://www.amazon.in/dp/{}", "Flipkart": "https://www.flipkart.com/product/p/itme?pid={}",
        "Myntra": "https://www.myntra.com/{}", "Nykaa": "https://www.nykaa.com/x/p/{}",
        "FirstCry": "https://www.firstcry.com/x/x/{}/product-detail", "Blinkit": "https://blinkit.com/prn/x/prid/{}",
        "Zepto": "https://www.zepto.com/pn/x/pvid/{}", "Instamart": "https://www.swiggy.com/instamart/item/{}",
        "BigBasket": "https://www.bigbasket.com/pd/{}/", "FK Minutes": "https://www.flipkart.com/product/p/itme?pid={}&marketplace=HYPERLOCAL"}


def product_url(channel, cid, url=""):
    url = (url or "").strip()
    if url.startswith("http"):
        return url
    return PAGE[channel].format(cid) if cid and channel in PAGE else ""


def build(log=print, aligned_label="", tolerance=1.0):
    if not os.path.exists(HIST_FILE):
        log("Dashboard: no history yet")
        return
    with open(HIST_FILE, encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    # newest reading of every product on every channel + pincode (Top 50 SKUs are checked
    # every 2 hours, the rest once a day - so each product keeps its own latest price)
    import datetime as _dt
    newest = max((r["Run Time"] for r in rows if r["Run Time"]), default="")
    try:
        cutoff = (_dt.datetime.strptime(newest, "%Y-%m-%d %H:%M") - _dt.timedelta(days=3)).strftime("%Y-%m-%d %H:%M")
    except ValueError:
        cutoff = ""
    latest = {}
    for r in rows:
        if not r["Channel"] or not r["Channel ID"] or r["Run Time"] < cutoff:
            continue
        k = (r["Channel"], r["Channel ID"], r["Pincode"])
        if r["Run Time"] >= latest.get(k, {}).get("Run Time", ""):
            latest[k] = r
    hist = price_history(rows, cutoff_days=7)
    rows = list(latest.values())
    last = {}
    for r in rows:
        last[r["Channel"]] = max(last.get(r["Channel"], ""), r["Run Time"])

    QC = ["Blinkit", "Zepto", "Instamart", "BigBasket", "FK Minutes"]
    pins = {}
    for r in rows:
        if r["Channel"] in QC and r["Pincode"] and r["Pincode"] != "All India":
            pins.setdefault(r["Pincode"], r["City"] or "")
    skus = {}
    for r in rows:
        sku = r["NH SKU"]
        if sku.upper().startswith("BL-"):          # quick-commerce code -> plain NH SKU
            sku = sku[3:]
        sp = f(r["Selling Price"])
        if not sku or sp is None or r["Status"] not in PRICED:
            continue
        d = skus.setdefault(sku, {"sku": sku, "name": "", "al": None, "ch": {}, "pc": {}})
        if r["Channel Product Name"] and not d["name"]:
            d["name"] = r["Channel Product Name"]
        if f(r["Aligned SP"]) is not None:
            d["al"] = f(r["Aligned SP"])
        place = "" if r["Pincode"] == "All India" else f"{r['City']} {r['Pincode']}".strip()
        url = product_url(r["Channel"], r["Channel ID"], r.get("URL"))
        d["ch"].setdefault(r["Channel"], []).append(
            {"p": sp, "oos": r["Status"] == "Out of stock", "at": place, "t": r["Run Time"][-5:], "u": url})
        if place:
            d["pc"].setdefault(r["Channel"], {}).setdefault(r["Pincode"], []).append(
                (sp, r["Status"] == "Out of stock", url))

    ev, bau, ev_label, bau_label = aligned_both()
    out = []
    for d in skus.values():
        cells = {}
        for ch, lst in d["ch"].items():
            live = [x for x in lst if not x["oos"]]
            prices = [x["p"] for x in (live or lst)]
            tip = "\n".join(f"{x['at'] or 'All India'}: Rs {round(x['p'])}  ({x['t']})" + (" (out of stock)" if x["oos"] else "")
                            for x in lst)
            best = min(live or lst, key=lambda x: x["p"])
            cells[ch] = [min(prices), max(prices), not live, tip, best["u"]]
        qc = {}
        for ch, per in d["pc"].items():
            for pin, lst in per.items():
                live = [x[0] for x in lst if not x[1]]
                best = min([x for x in lst if not x[1]] or lst, key=lambda x: x[0])
                qc.setdefault(ch, {})[pin] = [best[0], not live, best[2]]
        sku = d["sku"]
        out.append({"sku": sku, "name": d["name"], "cells": cells, "qc": qc,
                    "ale": ev.get(sku) if ev else d["al"], "alb": bau.get(sku)})
    cats = categories()
    pri = priority()
    crk = channel_ranks()
    for d in out:
        d["cr"] = crk.get(d["sku"], {})
        d["hist"] = hist.get(d["sku"], {})
        d["cat"] = cats.get(d["sku"], "")
        d["tier"], d["rank"] = pri.get(d["sku"], ("", None))

    chans = [c for c in CHANNELS if c in last]
    data = {"updated": max(last.values()) if last else "", "aligned": aligned_label,
            "eventLabel": ev_label or aligned_label, "bauLabel": bau_label, "tol": tolerance,
            "channels": chans, "rows": out, "logos": logos(log),
            "qcApps": [c for c in QC if c in last],
            "pins": [{"pin": p, "city": c} for p, c in sorted(pins.items(), key=lambda x: (x[1], x[0]))]}
    with open(TEMPLATE, encoding="utf-8") as fh:
        page = fh.read()
    page = page.replace("/*__DATA__*/null", json.dumps(data, ensure_ascii=False).replace("</", "<\\/"))
    os.makedirs(DASH_DIR, exist_ok=True)
    with open(os.path.join(DASH_DIR, "price-pulse.html"), "w", encoding="utf-8") as fh:
        fh.write(page)
    # GitHub Pages copy: https://<user>.github.io/<repo>/  (Settings -> Pages -> main, /docs)
    docs = os.path.join(ROOT, "docs")
    os.makedirs(docs, exist_ok=True)
    with open(os.path.join(docs, "index.html"), "w", encoding="utf-8") as fh:
        fh.write(page)
    open(os.path.join(docs, ".nojekyll"), "w").close()
    # same page at the top of the repository, so GitHub Pages works with "/ (root)" too
    with open(os.path.join(ROOT, "index.html"), "w", encoding="utf-8") as fh:
        fh.write(page)
    open(os.path.join(ROOT, ".nojekyll"), "w").close()
    log(f"Dashboard rebuilt: {len(out)} SKUs x {len(chans)} channels -> 3_Dashboard\\price-pulse.html")
    try:
        import push_github
        push_github.auto_push(log)
    except Exception:
        pass


def aligned_label():
    """'Aligned Price.xlsx (8th to 18th October)' - same file the tracker uses."""
    import glob
    files = glob.glob(os.path.join(ROOT, "1_Input", "Aligned Price*.xlsx")) + \
        glob.glob(os.path.join(os.path.dirname(ROOT), "Price Check", "Aligned Price*.xlsx"))
    files = [x for x in files if not os.path.basename(x).startswith("~$")]
    if not files:
        return ""
    f = max(files, key=os.path.getmtime)
    try:
        import openpyxl
        ws = openpyxl.load_workbook(f, read_only=True, data_only=True).worksheets[0]
        head = next(ws.iter_rows(values_only=True))
        return f"{os.path.basename(f)} ({head[1]})" if len(head) > 1 and head[1] else os.path.basename(f)
    except Exception:
        return os.path.basename(f)


if __name__ == "__main__":
    build(aligned_label=aligned_label())
