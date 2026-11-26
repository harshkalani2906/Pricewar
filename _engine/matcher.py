"""
Guess which NH SKU a quick-commerce listing is, from its name + pack size.

Used only when the listing is not already in QC Mapping.xlsx. The guess is
written to the mapping file as Match = "Auto xx%" so a person can confirm or
correct it once; after that the mapping file always wins.
"""
from __future__ import annotations

import re
from difflib import SequenceMatcher

STOP = {"nat", "habit", "for", "and", "with", "the", "of", "women", "men", "&", "in", "to",
        "100", "natural", "pure", "|", "-", "a", "pack", "free", "paraben", "sulphate",
        "all", "type", "types", "skin", "hair", "by", "x", "online", "india", "buy", "best", "price",
        "from", "gm", "each", "100%", "helps", "use", "daily"}
SYN = {"haldi": "turmeric", "chandan": "sandalwood", "mehendi": "henna", "mehndi": "henna",
       "kesar": "saffron", "neem": "neem", "shikakai": "shikakai", "facewash": "face wash",
       "ubtan": "ubtan", "conditioner": "conditioner", "kumkumadi": "kumkumadi"}


def tokens(s: str) -> set[str]:
    s = (s or "").lower().replace("face wash", "facewash")
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    out = set()
    for w in s.split():
        w = SYN.get(w, w)
        if w in STOP or len(w) < 3 or w.isdigit():
            continue
        if w.endswith("s") and len(w) > 4:
            w = w[:-1]
        out.add(w)
    return out


def sizes(s: str) -> set[int]:
    """All pack sizes (g / ml / pc) written in a text."""
    out = set()
    for n, u in re.findall(r"(\d+(?:\.\d+)?)\s*(kg|g|gm|gms|grams?|ml|l|ltr|litre|pcs?|pieces?|n)\b",
                           (s or "").lower()):
        v = float(n)
        if u in ("kg", "l", "ltr", "litre"):
            v *= 1000
        out.add(int(round(v)))
    return out


def sku_size(sku: str) -> int | None:
    m = re.search(r"-(\d{2,4})$", sku or "")
    return int(m.group(1)) if m else None


def sku_pack(sku: str) -> int:
    """HR-SH-HB2-250 -> 2 (pack of 2), CB-...-02 -> 2 (combo), else 1."""
    parts = (sku or "").split("-")
    if len(parts) >= 2:
        m = re.search(r"[A-Z](\d)$", parts[-2])
        if m and int(m.group(1)) >= 2:
            return int(m.group(1))
    m = re.search(r"-0(\d)$", sku or "")
    if m and int(m.group(1)) >= 2:
        return int(m.group(1))
    return 1


def sku_is_combo(sku: str) -> bool:
    return sku_pack(sku) > 1 or (sku or "").startswith("CB-")


def listing_pack(text: str) -> int:
    t = (text or "").lower()
    m = re.search(r"(?:pack|set|combo) of (\d)|(\d)\s*x\s*\d|\((\d) pcs?\)|(\d) pcs?\b", t)
    if m:
        n = int(next(g for g in m.groups() if g))
        return n if n >= 2 else 1
    return 2 if (re.search(r"combo|\+|\bduo\b|\bkit\b", t) or t.count("nat habit") > 1) else 1


class Matcher:
    def __init__(self, products: dict[str, dict]):
        """products: {sku: {"name": str, "names": [str,...], "aligned_sp": float}}"""
        self.p = []
        self.pack = {}
        self.has_price = {k for k, d in products.items() if d.get("aligned_sp")}
        self.dupes = {k for k in products if re.match(r"^[A-Z]{2,3}-", k) and k.split("-", 1)[1] in products}
        for sku, d in products.items():
            names = [n for n in ([d.get("name")] + list(d.get("names") or [])) if n]
            if not names:
                continue
            toks = [tokens(n) for n in names[:12]]
            szs = set()
            for n in names:
                szs |= sizes(n)
            if not szs and sku_size(sku) and not sku_is_combo(sku):
                szs.add(sku_size(sku))
            self.p.append((sku, toks, szs, " ".join(names).lower()))
            packs = [listing_pack(n) for n in names[:6]]
            self.pack[sku] = max(sku_pack(sku), 2 if sum(1 for x in packs if x > 1) * 2 > len(packs) else 1)
        # rare words (turmeric, rosemary...) count more than common ones (soap, shampoo...)
        import math
        df = {}
        for _, tl, _, _ in self.p:
            for t in set().union(*tl):
                df[t] = df.get(t, 0) + 1
        n = max(1, len(self.p))
        self.idf = {t: math.log(1 + n / c) for t, c in df.items()}

    def guess(self, name: str, size: str | None) -> tuple[str | None, int]:
        lt = tokens(name)
        ls = sizes(f"{name} {size or ''}")
        lp = listing_pack(f"{name} {size or ''}")
        best, best_score = None, 0.0
        w = lambda ts: sum(self.idf.get(t, 3.0) for t in ts)  # noqa: E731
        for sku, tl, szs, text in self.p:
            if not lt or sku in self.dupes:
                continue
            s = -1.0
            for toks in tl:
                if not toks:
                    continue
                inter = w(lt & toks)
                cover = inter / max(w(lt), 1e-6)
                jacc = inter / max(w(lt | toks), 1e-6)
                extra = w(toks - lt) / max(w(toks), 1e-6)
                v = 0.55 * cover + 0.30 * jacc + 0.15 * SequenceMatcher(None, " ".join(sorted(lt)), " ".join(sorted(toks))).ratio()
                v -= 0.10 * extra
                s = max(s, v)
            if s < 0:
                continue
            if ls and szs:
                s += 0.15 if (ls & szs) else -0.12
            sp_ = self.pack.get(sku, 1) if not sku.startswith("CB-") else 2
            if (lp > 1) != (sp_ > 1):
                continue            # single pack never matches a multi-pack / combo SKU
            if lp != sp_ and not sku.startswith("CB-"):
                s -= 0.10
            if sku in self.has_price:
                s += 0.04
            if s > best_score:
                best, best_score = sku, s
        pct = int(max(0, min(1, best_score)) * 100)
        return best, pct
