"""SEC EDGAR (sociétés cotées aux États-Unis) — accès direct aux API publiques de la SEC, sans bibliothèque tierce.

- annual_facts(ticker) : séries annuelles des 10-K (XBRL « companyfacts ») sur ~15 ans : CA, résultat
  d'exploitation, amortissements, résultat net, flux d'exploitation, investissements, rémunération en actions,
  rachats d'actions, nombre d'actions dilué.
- form4_rows(ticker, days) : déclarations d'initiés (Form 4) au format des lignes OpenInsider / fiche Finviz
  [nom, fonction, « Sep 10 '26 », Buy/Sale, prix, titres, valeur, détention après].

La SEC exige un en-tête User-Agent « Nom adresse@mail » (pas de clé, pas d'inscription) ; limite 10 requêtes/s.
L'adresse est saisie dans la barre latérale du dashboard (ou variable d'environnement SEC_USER_AGENT), jamais
enregistrée dans un fichier. Réponses mises en cache sur disque (cache_finviz/edgar).
"""
from __future__ import annotations

import gzip
import json
import os
import re
import time
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta

import requests

CONFIG = {"ua": os.environ.get("SEC_USER_AGENT", "").strip(), "cache": os.path.join("cache_finviz", "edgar")}
MONTHS_EN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
_LAST = [0.0]


class EdgarError(RuntimeError):
    pass


def configure(ua=None, cache=None):
    if ua is not None:
        CONFIG["ua"] = ua.strip()
    if cache:
        CONFIG["cache"] = cache


def enabled():
    """Accès EDGAR activé : un User-Agent contenant une adresse e-mail est renseigné."""
    return bool(re.search(r"\S+@\S+\.\S+", CONFIG.get("ua") or ""))


def _path(*parts):
    p = os.path.join(CONFIG["cache"], *parts)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    return p


def _get(url, timeout=30):
    if not enabled():
        raise EdgarError("EDGAR : renseignez « Nom adresse@mail » (exigé par la SEC) dans la barre latérale")
    wait = 0.12 - (time.time() - _LAST[0])          # ≤ ~8 requêtes / s (limite SEC : 10)
    if wait > 0:
        time.sleep(wait)
    for attempt in range(3):
        _LAST[0] = time.time()
        r = requests.get(url, headers={"User-Agent": CONFIG["ua"], "Accept-Encoding": "gzip, deflate"},
                         timeout=timeout)
        if r.status_code in (429, 503):
            time.sleep(min(float(r.headers.get("Retry-After") or 2 * (attempt + 1)), 20))
            continue
        if r.status_code == 403:
            raise EdgarError("EDGAR : accès refusé par la SEC (403) — vérifiez le « Nom adresse@mail » saisi")
        r.raise_for_status()
        return r
    raise EdgarError("EDGAR : la SEC limite temporairement les requêtes, réessayez dans une minute")


def _cached_json(url, name, max_age_h):
    p = _path(name)
    if os.path.exists(p) and time.time() - os.path.getmtime(p) < max_age_h * 3600:
        try:
            with gzip.open(p, "rt", encoding="utf-8") as fh:
                return json.load(fh)
        except (OSError, ValueError, EOFError):
            pass                                       # cache illisible (écriture interrompue) : relu en ligne
    data = _get(url).json()
    with gzip.open(p, "wt", encoding="utf-8") as fh:
        json.dump(data, fh)
    return data


# ------------------------------------------------------------------ ticker -> CIK
def cik_for(ticker):
    m = _cached_json("https://www.sec.gov/files/company_tickers.json", "company_tickers.json.gz", 24 * 7)
    t = ticker.upper().replace(".", "-")
    for row in m.values():
        if str(row.get("ticker", "")).upper() == t:
            return int(row["cik_str"])
    return None


# ------------------------------------------------------------------ séries annuelles (10-K)
TAGS = {   # premier concept disponible (par exercice), en millions sauf actions (millions) et BNA
    "rev": ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax", "SalesRevenueNet",
            "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueGoodsNet"],
    "ebit": ["OperatingIncomeLoss"],
    "da": ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization",
           "DepreciationAmortizationAndAccretionNet", "Depreciation"],
    "ni": ["NetIncomeLoss", "ProfitLoss"],
    "ocf": ["NetCashProvidedByUsedInOperatingActivities",
            "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"],
    "capex": ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets"],
    "sbc": ["ShareBasedCompensation", "AllocatedShareBasedCompensationExpense"],
    "buyback": ["PaymentsForRepurchaseOfCommonStock", "PaymentsForRepurchaseOfEquity"],
    "interest": ["InterestExpense", "InterestExpenseNonoperating", "InterestPaidNet"],
    "shares": ["WeightedAverageNumberOfDilutedSharesOutstanding"],
    "eps": ["EarningsPerShareDiluted"],
}
ANNUAL_FORMS = ("10-K", "10-K/A", "10-KT", "20-F", "20-F/A", "40-F")


def parse_companyfacts(data):
    """companyfacts -> {exercice: {clé: valeur}} ; exercice = année de clôture ; durée ≈ 1 an ; pour chaque
    clôture, valeur du dépôt le plus récent (retraitements compris). Montants en millions."""
    facts = (data or {}).get("facts", {})
    gaap = {**facts.get("ifrs-full", {}), **facts.get("us-gaap", {})}
    out = {}
    for key, tags in TAGS.items():
        best = {}                                         # année -> (priorité du concept, date de dépôt, valeur)
        for prio, tag in enumerate(tags):
            units = (gaap.get(tag) or {}).get("units", {})
            unit = "shares" if key == "shares" else ("USD/shares" if key == "eps" else "USD")
            for e in units.get(unit, []):
                if e.get("form") not in ANNUAL_FORMS or "end" not in e:
                    continue
                if e.get("start"):
                    try:
                        dur = (datetime.strptime(e["end"], "%Y-%m-%d") - datetime.strptime(e["start"], "%Y-%m-%d")).days
                    except ValueError:
                        continue
                    if not 340 <= dur <= 390:
                        continue
                elif key not in ("shares",):              # flux : période obligatoire
                    continue
                y = int(e["end"][:4])                     # année de clôture (comme les libellés Finviz « 2026FY »)
                cand = (prio, e.get("filed", ""), e["val"])
                cur = best.get(y)
                if cur is None or cand[0] < cur[0] or (cand[0] == cur[0] and cand[1] > cur[1]):
                    best[y] = cand
        for y, (_, _, v) in best.items():
            scale = 1 if key == "eps" else 1e6
            out.setdefault(y, {})[key] = v / scale
    return dict(sorted(out.items()))


def annual_facts(ticker):
    """Séries annuelles 10-K d'une société US -> {exercice: {...}} (cache 1 jour) ; {} si inconnue de la SEC."""
    cik = cik_for(ticker)
    if not cik:
        return {}
    data = _cached_json(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json",
                        f"facts_{cik}.json.gz", 24)
    return parse_companyfacts(data)


def split_adjusted_shares(series):
    """Nombre d'actions par exercice corrigé des divisions/regroupements (saut > ×1,6 ou < ×0,62 d'une année sur
    l'autre) : les années antérieures sont ramenées à la base actuelle. series : [(année, actions)] croissant."""
    s = [list(x) for x in series if x[1]]
    for i in range(len(s) - 1, 0, -1):
        r = s[i][1] / s[i - 1][1]
        if r > 1.6 or r < 0.62:
            k = round(r) if r > 1.6 else 1 / round(1 / r)
            for j in range(i):
                s[j][1] *= k
    return [tuple(x) for x in s]


# ------------------------------------------------------------------ Form 4 (initiés)
def _txt(node, path):
    el = node.find(path) if node is not None else None
    return el.text.strip() if el is not None and el.text else None


def _num(x):
    try:
        return float(str(x).replace(",", ""))
    except (TypeError, ValueError):
        return None


def parse_form4(xml_text):
    """Form 4 XML -> lignes agrégées par (déclarant, date, achat/vente) : [nom, fonction, date, Buy/Sale, prix,
    titres, valeur, détention après]. Achats « P » et ventes « S » en bourse seulement ; vente marquée « avec levée
    d'options » si le même dépôt contient un exercice (code M)."""
    try:
        root = ET.fromstring(xml_text.encode("utf-8") if isinstance(xml_text, str) else xml_text)
    except ET.ParseError:
        return []
    own = root.find("reportingOwner")
    name = _txt(own, "reportingOwnerId/rptOwnerName") or "?"
    rel = own.find("reportingOwnerRelationship") if own is not None else None
    yes = lambda tag: (_txt(rel, tag) or "").lower() in ("1", "true")
    title = _txt(rel, "officerTitle")
    role = title or ("Director" if yes("isDirector") else "10% Owner" if yes("isTenPercentOwner") else
                     "Officer" if yes("isOfficer") else "Other")
    if title and yes("isDirector"):
        role = f"{title}, Director"
    codes = [_txt(t, "transactionCoding/transactionCode") for t in root.iter("derivativeTransaction")]
    codes += [_txt(t, "transactionCoding/transactionCode") for t in root.iter("nonDerivativeTransaction")]
    oe = "M" in codes
    groups = {}
    for t in root.iter("nonDerivativeTransaction"):
        code = _txt(t, "transactionCoding/transactionCode")
        if code not in ("P", "S"):
            continue
        d = (_txt(t, "transactionDate/value") or "")[:10]
        sh = _num(_txt(t, "transactionAmounts/transactionShares/value")) or 0.0
        px = _num(_txt(t, "transactionAmounts/transactionPricePerShare/value")) or 0.0
        after = _num(_txt(t, "postTransactionAmounts/sharesOwnedFollowingTransaction/value"))
        g = groups.setdefault((d, code), {"sh": 0.0, "val": 0.0, "after": None})
        g["sh"] += sh
        g["val"] += sh * px
        g["after"] = after if after is not None else g["after"]
    rows = []
    for (d, code), g in sorted(groups.items()):
        try:
            dt = datetime.strptime(d, "%Y-%m-%d")
        except ValueError:
            continue
        if g["sh"] <= 0:
            continue
        typ = "Buy" if code == "P" else "Sale"
        rows.append([name.title(), role + (" · avec levée d'options" if typ == "Sale" and oe else ""),
                     f"{MONTHS_EN[dt.month - 1]} {dt.day:02d} '{dt.year % 100:02d}", typ,
                     f"{g['val'] / g['sh']:.2f}", f"{g['sh']:,.0f}", f"{g['val']:,.0f}",
                     f"{g['after']:,.0f}" if g["after"] is not None else ""])
    return rows


def form4_rows(ticker, days=365, max_filings=150, log=print):
    """Déclarations Form 4 des 12 derniers mois (dépôts de la société) -> lignes au format OpenInsider.
    None si la société est inconnue de la SEC."""
    cik = cik_for(ticker)
    if not cik:
        return None
    sub = _cached_json(f"https://data.sec.gov/submissions/CIK{cik:010d}.json", f"sub_{cik}.json.gz", 6)
    rec = (sub.get("filings") or {}).get("recent") or {}
    since = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
    idx = [i for i, fm in enumerate(rec.get("form", [])) if fm == "4" and rec["filingDate"][i] >= since]
    rows = []
    for i in idx[:max_filings]:
        acc = rec["accessionNumber"][i]
        doc = re.sub(r"^xsl[^/]*/", "", rec["primaryDocument"][i])
        p = _path("form4", f"{acc}.xml")
        if os.path.exists(p):
            with open(p, encoding="utf-8") as fh:
                xml = fh.read()
        else:
            try:
                xml = _get(f"https://www.sec.gov/Archives/edgar/data/{cik}/{acc.replace('-', '')}/{doc}").text
            except Exception as e:  # noqa: BLE001
                log(f"{ticker} : Form 4 {acc} illisible ({e})")
                continue
            with open(p, "w", encoding="utf-8") as fh:
                fh.write(xml)
        rows += parse_form4(xml)
    return rows
