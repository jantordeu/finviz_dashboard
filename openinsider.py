"""
OpenInsider : transactions d'initiés (achats P / ventes S) pour une liste de tickers.

- Levées d'options TOUJOURS exclues (paramètre excludeDerivRelated=1).
- Tickers interrogés par paquets de 100 (S&P 500 complet ≈ 5 requêtes, ≈ 5 s).
- Colonnes lues par leur nom (la page change de colonnes selon le nombre de tickers).
- Diagnostic : chaque requête est journalisée (URL, code HTTP, nb de lignes).
"""
from __future__ import annotations

import re
import time
from datetime import date, timedelta
from typing import Callable, Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup

URLS = ["http://openinsider.com/screener", "https://openinsider.com/screener"]
PAGE_SIZE = 5000
BATCH_SIZE = 100          # 5 requêtes pour tout le S&P 500
CLUSTER_DAYS = 30            # fenêtre pour un « achat groupé »
STRONG_DELTA_OWN = 50.0      # hausse de participation jugée forte (%)
LARGE_TRADE_USD = 50e6       # montant individuel jugé inhabituel

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml",
    "Accept-Language": "en-US,en;q=0.9",
}

CRITERIA = {
    "buyers": "Initiés distincts : acheteurs > vendeurs",
    "tx": "Transactions : nb achats > nb ventes",
    "usd": "Montants : solde $ > 0",
}

ENTITY_RE = re.compile(
    r"\b(inc|llc|l\.?p\.?|ltd|trust|holdings?|fund|capital|partners|corp|corporation|"
    r"management|group|foundation|company|co\.)\b", re.I)
ROLES = [("CEO", "CEO"), ("CFO", "CFO"), ("COO", "COO"), ("COB", "Président"),
         ("Chair", "Président"), ("Pres", "Président"), ("Dir", "Administrateur")]


# --------------------------------------------------------------------------- #
# Lecture des pages
# --------------------------------------------------------------------------- #
def _clean(t: str) -> str:
    return re.sub(r"\s+", " ", (t or "").replace("\xa0", " ")).strip()


def _money(t: str) -> float:
    t = re.sub(r"[^0-9.\-]", "", t or "")
    try:
        return float(t) if t not in ("", "-", ".") else 0.0
    except ValueError:
        return 0.0


def _delta_own(t: str) -> Optional[float]:
    t = _clean(t)
    if not t:
        return None
    if t.lower() == "new":
        return 100.0          # nouvelle position
    if t.startswith(">"):
        return 999.0
    try:
        return float(t.replace("%", "").replace("+", ""))
    except ValueError:
        return None


def has_table(html: str) -> bool:
    return 'class="tinytable"' in html or "class='tinytable'" in html


def parse_screener(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    table = soup.select_one("table.tinytable")
    if table is None or table.tbody is None:
        return []
    headers = [_clean(th.get_text()) for th in table.select("thead th")]
    idx = {h: i for i, h in enumerate(headers)}

    def cell(c, name):
        i = idx.get(name)
        return _clean(c[i].get_text()) if i is not None and i < len(c) else ""

    out = []
    for tr in table.tbody.find_all("tr"):
        c = tr.find_all("td")
        if not c:
            continue
        tt = cell(c, "Trade Type")
        out.append({
            "ticker": cell(c, "Ticker").upper(),
            "filing_date": cell(c, "Filing Date")[:10],
            "trade_date": cell(c, "Trade Date"),
            "insider": cell(c, "Insider Name"),
            "title": cell(c, "Title"),
            "type": tt[:1],
            "trade_type": tt,
            "price": _money(cell(c, "Price")),
            "qty": _money(cell(c, "Qty")),
            "owned": _money(cell(c, "Owned")),
            "delta_own": _delta_own(cell(c, "ΔOwn")),
            "value": abs(_money(cell(c, "Value"))),
        })
    return out


# --------------------------------------------------------------------------- #
# Téléchargement
# --------------------------------------------------------------------------- #
def date_range(days: int, years_back: int = 0, today: Optional[date] = None,
               periods_back: int = 0) -> tuple[date, date]:
    """Période [aujourd'hui − days ; aujourd'hui], décalée de N années
    ou de N périodes consécutives (periods_back=1 : les `days` jours précédents)."""
    end = (today or date.today()) - timedelta(days=days * periods_back)
    if years_back:
        try:
            end = end.replace(year=end.year - years_back)
        except ValueError:                      # 29 février
            end = end.replace(year=end.year - years_back, day=28)
    start = end - timedelta(days=days)
    if periods_back:          # périodes consécutives sans chevauchement
        end -= timedelta(days=1)
    return start, end


def _params(batch, days, page, period: Optional[tuple[date, date]] = None):
    p = {"s": " ".join(batch), "fd": days, "xp": 1, "xs": 1,
         "excludeDerivRelated": 1, "cnt": PAGE_SIZE, "page": page}
    if period:                                  # dates de déclaration explicites
        p["fd"] = -1
        p["fdr"] = f"{period[0]:%m/%d/%Y} - {period[1]:%m/%d/%Y}"
    return p


def fetch_transactions(tickers: list[str], days: int = 90, delay: float = 0.5,
                       session: Optional[requests.Session] = None,
                       progress: Optional[Callable[[int, int, str], None]] = None,
                       period: Optional[tuple[date, date]] = None
                       ) -> tuple[pd.DataFrame, list[dict]]:
    """Retourne (transactions, journal de diagnostic)."""
    s = session or requests.Session()
    s.headers.update(HEADERS)
    batches = [tickers[i:i + BATCH_SIZE] for i in range(0, len(tickers), BATCH_SIZE)]
    rows, log = [], []
    for b, batch in enumerate(batches, start=1):
        if progress:
            progress(b, len(batches), f"{batch[0]} … {batch[-1]}")
        page = 1
        while True:
            part, entry = None, {"lot": b, "page": page, "tickers": len(batch)}
            for url in URLS:                      # http puis https
                try:
                    r = s.get(url, params=_params(batch, days, page, period), timeout=30)
                    entry.update(url=r.url, http=r.status_code, octets=len(r.text))
                    if r.status_code == 200 and has_table(r.text):
                        part = parse_screener(r.text)
                        break
                    entry["erreur"] = "pas de tableau dans la page" if r.status_code == 200 \
                        else f"HTTP {r.status_code}"
                except requests.RequestException as e:
                    entry.update(url=url, erreur=str(e)[:200])
            entry["lignes"] = len(part) if part is not None else 0
            log.append(entry)
            if not part:
                break
            rows.extend(part)
            if len(part) < PAGE_SIZE:
                break
            page += 1
            time.sleep(delay)
        time.sleep(delay)
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df[df["ticker"].isin({t.upper() for t in tickers})].drop_duplicates()
        df = df.reset_index(drop=True)
    return df, log


# --------------------------------------------------------------------------- #
# Indicateurs par valeur
# --------------------------------------------------------------------------- #
def _cluster_size(buys: pd.DataFrame) -> int:
    """Nb max d'acheteurs distincts dans une fenêtre de CLUSTER_DAYS jours."""
    if buys.empty:
        return 0
    d = buys.assign(dt=pd.to_datetime(buys["trade_date"], errors="coerce")).dropna(subset=["dt"])
    best = 0
    for start in d["dt"].unique():
        w = d[(d["dt"] >= start) & (d["dt"] <= start + pd.Timedelta(days=CLUSTER_DAYS))]
        best = max(best, w["insider"].nunique())
    return best


def _roles(titles) -> str:
    found = []
    for t in titles:
        for key, label in ROLES:
            if re.search(rf"\b{key}", t or "") and label not in found:
                found.append(label)
                break
    return ", ".join(found[:3])


def _atypical(buys: pd.DataFrame) -> str:
    notes = []
    for _, r in buys.iterrows():
        if "10%" in (r["title"] or ""):
            notes.append(f"actionnaire >10 % : {r['insider']}")
        elif ENTITY_RE.search(r["insider"] or ""):
            notes.append(f"entité : {r['insider']}")
        if r["value"] >= LARGE_TRADE_USD:
            notes.append(f"montant inhabituel {r['value'] / 1e6:.0f} M$")
    return " ; ".join(dict.fromkeys(notes))


def _m(v: float) -> str:
    return f"{v / 1e6:.1f} M$" if v >= 1e5 else f"{v / 1e3:.0f} k$"


def summarize(tx: pd.DataFrame, tickers: list[str], exclude_10pct: bool = False) -> pd.DataFrame:
    if not tx.empty and exclude_10pct:
        tx = tx[~tx["title"].str.contains("10%", na=False)]
    out = []
    for t in tickers:
        d = tx[tx["ticker"] == t.upper()] if not tx.empty else tx
        buys = d[d["type"] == "P"] if not d.empty else d
        sells = d[d["type"] == "S"] if not d.empty else d
        nb, ns = (buys["insider"].nunique(), sells["insider"].nunique()) if not d.empty else (0, 0)
        bv, sv = (float(buys["value"].sum()), float(sells["value"].sum())) if not d.empty else (0.0, 0.0)
        cl = _cluster_size(buys) if not d.empty else 0
        mdo = buys["delta_own"].max() if not d.empty and not buys.empty else None

        # Pastilles et texte, façon tableau de bord
        if d is None or d.empty or (nb == 0 and ns == 0):
            dots, txt = "⚪", "aucune transaction"
        else:
            if nb > ns and cl >= 2:
                dots = "🟢🟢"
            elif nb > ns:
                dots = "🟢"
            else:
                dots = "🟠"
            parts = []
            if cl >= 2:
                roles = _roles(buys["title"])
                parts.append(f"achats groupés ({cl} en {CLUSTER_DAYS} j)"
                             + (f", dont {roles}" if roles else "") + f" +{_m(bv)}")
            elif nb:
                roles = _roles(buys["title"])
                parts.append(f"{nb} acheteur{'s' if nb > 1 else ''}"
                             + (f" ({roles})" if roles else "") + f" +{_m(bv)}")
            if ns:
                parts.append(f"ventes {_m(sv)} ({ns} vendeur{'s' if ns > 1 else ''})")
            txt = " · ".join(parts)
        out.append({
            "ticker": t,
            "oi_signal": f"{dots} {txt}",
            "buyers": int(nb), "sellers": int(ns),
            "buy_tx": int(len(buys)) if not d.empty else 0,
            "sell_tx": int(len(sells)) if not d.empty else 0,
            "buy_usd": bv, "sell_usd": sv, "balance_usd": bv - sv,
            "max_delta_own": mdo,
            "strong_delta": bool(mdo is not None and mdo >= STRONG_DELTA_OWN),
            "cluster": cl,
            "atypical": _atypical(buys) if not d.empty else "",
            "last_buy": buys["trade_date"].max() if not d.empty and not buys.empty else None,
            "buyer_names": ", ".join(f"{n} ({ti})" for n, ti in
                                     buys.drop_duplicates("insider")[["insider", "title"]].values)
            if not d.empty else "",
        })
    return pd.DataFrame(out)


def criterion_mask(df: pd.DataFrame, criterion: str) -> pd.Series:
    if criterion == "tx":
        return df["buy_tx"] > df["sell_tx"]
    if criterion == "usd":
        return df["balance_usd"] > 0
    return df["buyers"] > df["sellers"]


# --------------------------------------------------------------------------- #
# Sentiment de l'univers
# --------------------------------------------------------------------------- #
def sentiment(tx: pd.DataFrame) -> dict:
    """Indicateurs agrégés achats / ventes sur un ensemble de transactions."""
    if tx is None or tx.empty:
        return {"achats": 0, "ventes": 0, "acheteurs": 0, "vendeurs": 0,
                "achats_usd": 0.0, "ventes_usd": 0.0, "ratio_tx": None, "ratio_ins": None,
                "ratio_usd": None, "valeurs_pos": 0}
    b, s = tx[tx["type"] == "P"], tx[tx["type"] == "S"]
    nb_ins = lambda d: d[["ticker", "insider"]].drop_duplicates().shape[0]
    per = tx.groupby("ticker").apply(
        lambda d: d.loc[d["type"] == "P", "insider"].nunique() > d.loc[d["type"] == "S", "insider"].nunique(),
        include_groups=False)
    out = {"achats": len(b), "ventes": len(s), "acheteurs": nb_ins(b), "vendeurs": nb_ins(s),
           "achats_usd": float(b["value"].sum()), "ventes_usd": float(s["value"].sum()),
           "valeurs_pos": int(per.sum())}
    out["ratio_tx"] = out["achats"] / out["ventes"] if out["ventes"] else None
    out["ratio_ins"] = out["acheteurs"] / out["vendeurs"] if out["vendeurs"] else None
    out["ratio_usd"] = out["achats_usd"] / out["ventes_usd"] if out["ventes_usd"] else None
    return out
