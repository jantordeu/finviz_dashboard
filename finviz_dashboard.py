# ============================================================
#  Finviz Dashboard — portefeuille US et valeurs à étudier (Streamlit)
#
#  Installation : pip install -r requirements.txt
#  Lancement    : streamlit run finviz_dashboard.py
#
#  Sources :
#   - finviz.com (en ligne) : fiche (ratios, consensus, analystes, initiés, actualités), états financiers
#     (API JSON), chaîne d'options, sociétés comparables, multiples d'industrie, screener
#   - openinsider.com : achats / ventes des dirigeants sur 12 mois (formulaires 4 SEC)
#   - dataroma.com : positions et mouvements des « super investors » (13F, trimestriels)
#   - apewisdom.io : actions les plus citées sur Reddit (24 h)
#   - Zonebourse : uniquement des pages enregistrées (Cmd+S) et déposées par l'utilisateur
#     (Prévisions/Finances + Valorisation) : consensus sur 3 ans, PER historique, objectif
#  Cache : ./cache_finviz (2 h) + pages HTML en mémoire ; case « Utiliser le cache » et bouton
#  « Vider le cache ». Le fichier de positions n'est jamais mis en cache.
#
#  Onglets :
#   - Portefeuille : valeurs en USD du fichier de positions (+ tickers ajoutés), combinées avec les
#     pages Zonebourse déposées pour la même valeur (📄)
#   - Valeurs à étudier : pages Finviz / Zonebourse déposées (doublons ignorés) + valeurs envoyées par
#     le Screener JV ; prix d'entrée si la valeur est en portefeuille
#   - Fear & Greed : indice CNN du sentiment de marché (jauge, historique 1 an avec le S&P 500, composantes)
#   - Sentiment screener : actions les plus citées sur Reddit à l'actualité positive
#   - Screener JV : lecture d'un screener Finviz, exclusion d'industries, juste valeur rapide,
#     envoi en analyse complète
#   - Screener investors (NYSE / Nasdaq) : S&P 500 détenu par les superinvestors (Dataroma), signal des
#     dirigeants (OpenInsider, 12 mois) et des fonds ; modules analyse.py, openinsider.py,
#     dataroma_scraper.py, exports.py, signaux.py dans le même dossier
#   - Screener AMF (Euronext Paris) : déclarations des dirigeants (base BDIF de l'AMF) ; amf.py, signaux.py,
#     sbf120.csv
#   - Méthode : explication des calculs
#
#  Juste valeur :
#   - DCF d'entreprise (FCFF sur 10 ans, CMPC avec bêta ajusté, valeur terminale = moyenne croissance
#     perpétuelle / multiple de sortie, 3 scénarios 25/50/25, Monte Carlo P10–P90), consensus
#     Zonebourse pour les premières années si les pages sont déposées
#   - JV = 50 % DCF + 50 % médiane des multiples (PER cible, comparables, PER historique), réglable
#   - Banques / courtiers / assureurs : actif net × P/B justifié + PER bancaire
#   - Valeurs cycliques (⟳) : PER réduit, croissance plafonnée, marge ramenée à la médiane du cycle
#   - Précision JV (écart DCF / multiples), lecture de l'écart, DCF inversé (croissance implicite /
#     attendue sur 5 ans)
#  Aussi : initiés et super investors, analystes, rapport Put/Call, scénarios BNA / EBITDA / capex /
#  dette nette, export Excel (synthèse, scénarios, modèles DCF).
#
#  Ce n'est pas un conseil en investissement.
# ============================================================
import copy
import hashlib
import io
import json
import os
import random
import re
import time
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from bs4 import BeautifulSoup

import edgar_sec as ES

TODAY = datetime.now()
CACHE_DIR = "./cache_finviz"
CACHE_HOURS = 2                       # durée de validité du cache des pages en ligne
UA = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9", "Referer": "https://finviz.com/",
}
DEFAULT_TICKERS = ["UBER", "META", "NVDA", "GOOGL", "AVGO", "MSFT", "AMZN", "AAPL", "ORCL"]


# ============================================================
#  OUTILS
# ============================================================
def num(v):
    """'1.5B' -> 1.5e9 ; '12.3%' -> 12.3 ; '55,227.00' -> 55227 ; '-' -> None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return None if pd.isna(v) else float(v)
    s = str(v).replace(",", "").replace("%", "").replace("$", "").strip().split(" ")[0]
    if s in ("", "-", "—", "N/A"):
        return None
    mult = {"K": 1e3, "M": 1e6, "B": 1e9, "T": 1e12}
    try:
        if s[-1].upper() in mult and len(s) > 1:
            return float(s[:-1]) * mult[s[-1].upper()]
        return float(s)
    except ValueError:
        return None


def clamp(x, lo, hi):
    return max(lo, min(hi, x))


def stmt(data, label, i=0):
    """Valeur i (0 = TTM ou dernier exercice) d'une ligne d'état financier Finviz (en millions)."""
    if not data or label not in data or len(data[label]) <= i:
        return None
    return num(data[label][i])


# ============================================================
#  PORTEFEUILLE
# ============================================================
def read_xlsx_stdlib(file):
    """Lecture directe du xlsx (zip + XML), sans openpyxl : contourne les styles que
    certains courtiers écrivent de façon non standard (erreur « expected Fill »)."""
    import zipfile
    import xml.etree.ElementTree as ET
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    data = file.read() if hasattr(file, "read") else open(file, "rb").read()
    z = zipfile.ZipFile(io.BytesIO(data))
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", ns):
            shared.append("".join(t.text or "" for t in si.iter(f"{{{ns['m']}}}t")))
    sheet = sorted(n for n in z.namelist() if n.startswith("xl/worksheets/sheet"))[0]
    rows = []
    for r in ET.fromstring(z.read(sheet)).iter(f"{{{ns['m']}}}row"):
        vals = {}
        for c in r.findall("m:c", ns):
            col = 0
            for ch in re.match(r"[A-Z]+", c.get("r")).group():
                col = col * 26 + ord(ch) - 64
            t, v = c.get("t"), c.find("m:v", ns)
            if t == "s" and v is not None:
                val = shared[int(v.text)]
            elif t == "inlineStr":
                val = "".join(x.text or "" for x in c.iter(f"{{{ns['m']}}}t"))
            elif v is not None:
                try:
                    val = float(v.text)
                except ValueError:
                    val = v.text
            else:
                val = None
            vals[col - 1] = val
        rows.append([vals.get(i) for i in range(max(vals) + 1)] if vals else [])
    return pd.DataFrame(rows)


def read_excel_any(file):
    for eng in ("openpyxl", "calamine"):
        try:
            if hasattr(file, "seek"):
                file.seek(0)
            return pd.read_excel(file, header=None, engine=eng)
        except Exception:
            continue
    if hasattr(file, "seek"):
        file.seek(0)
    return read_xlsx_stdlib(file)


def load_portfolio(file, only_nyse=False, all_ccy=False):
    """Garde les actions cotées en USD (colonne « Devise »), hors ETF ; option : NYSE seulement.
    all_ccy=True : toutes les actions, toutes devises (prix d'entrée des valeurs à étudier)."""
    raw = read_excel_any(file)
    hdr = next(i for i in range(min(10, len(raw))) if "Symbole" in raw.iloc[i].astype(str).tolist())
    df = raw.iloc[hdr + 1:].copy()
    df.columns = raw.iloc[hdr].astype(str).str.replace("\n", " ").str.strip()
    df = df[df["Symbole"].astype(str).str.contains(":")]
    df["Ticker"] = df["Symbole"].str.split(":").str[0].str.upper()
    df["Bourse"] = df["Symbole"].str.split(":").str[1].str.lower()
    df["Devise"] = df["Devise"].astype(str).str.strip().str.upper()
    if not all_ccy:
        df = df[df["Devise"] == "USD"]
    typ = next((c for c in df.columns if c.lower().startswith("type d")), None)
    if typ:
        df = df[~df[typ].astype(str).str.contains("ETF|Fund|Fonds", case=False)]
    if only_nyse and not all_ccy:
        df = df[df["Bourse"] == "xnys"]
    keep = {"Instruments": "Nom", "Quantité": "Qté", "Prix entrée": "Prix d'entrée", "Devise": "Devise"}
    out = df[["Ticker", "Bourse"] + [c for c in keep if c in df.columns]].rename(columns=keep)
    out["Bourse"] = out["Bourse"].map({"xnys": "NYSE", "xnas": "NASDAQ"}).fillna(out["Bourse"].str.upper())
    return out.reset_index(drop=True)


# ============================================================
#  COLLECTE (en ligne, avec cache disque)
# ============================================================
def _session():
    try:
        import cloudscraper
        return cloudscraper.create_scraper(browser={"browser": "chrome", "platform": "darwin", "mobile": False})
    except ImportError:
        import requests
        return requests.Session()


class FinvizBlocked(RuntimeError):
    """Site qui bloque temporairement les requêtes automatiques (protection anti-robots, trop de requêtes)."""


def _retry_after(r, default=5.0):
    try:
        return float(r.headers.get("Retry-After") or default)
    except ValueError:
        return default


def _get(s, url, pause=(1.0, 2.0)):
    for attempt in range(2):
        r = s.get(url, headers=UA, timeout=25)
        if r.status_code == 429 and attempt == 0:           # trop de requêtes : on attend le délai demandé, une fois
            time.sleep(min(_retry_after(r), 30))
            continue
        break
    if r.status_code in (403, 429, 503):
        challenge = r.headers.get("cf-mitigated") == "challenge" or "Just a moment" in r.text[:4000]
        if challenge or r.status_code == 429:
            host = re.sub(r"^https?://(www\.)?([^/]+).*$", r"\2", url)
            wait = int(_retry_after(r, 60)) if r.status_code == 429 else None
            raise FinvizBlocked(f"{host} bloque temporairement les requêtes automatiques ("
                                + ("protection anti-robots" if challenge else "trop de requêtes")
                                + (f", attendre {wait} s" if wait else "") + ") — réessayez dans quelques minutes")
    r.raise_for_status()
    time.sleep(random.uniform(*pause))
    return r.text


def parse_quote(html):
    soup = BeautifulSoup(html, "html.parser")
    snap = {}
    for tb in soup.select("table.snapshot-table2"):
        cells = [td.get_text(" ", strip=True) for td in tb.find_all("td")]
        for i in range(0, len(cells) - 1, 2):
            k = cells[i]
            snap[k + " %" if k in snap else k] = cells[i + 1]   # 2e « EPS next Y » = croissance en %

    def rows(sel):
        tb = soup.select_one(sel)
        if not tb:
            return []
        return [[re.sub(r"\s+", " ", c.get_text(" ", strip=True)) for c in tr.find_all(["td", "th"])]
                for tr in tb.find_all("tr")][1:]
    return {"snap": snap, "ratings": rows("table.js-table-ratings"), "insiders": rows("table.body-table"),
            "news": parse_news(soup),
            "sector": (lambda a: a.get_text(strip=True) if a else None)(soup.select_one('a[href*="f=sec_"]')),
            "industry": (lambda a: (a.get("title") or a.get_text(strip=True)) if a else None)(
                soup.select_one('a[href*="f=ind_"]')),
            "peers": (lambda a: [x for x in re.sub(r".*[?&]t=([^&]+).*", r"\1", a.get("href", "")).upper().split(",")
                                 if x][1:] if a else [])(
                next((a for a in soup.find_all("a", href=True) if a.get_text(strip=True) == "Peers"), None)),
            "company": re.sub(r"^[A-Z0-9.\-]+ - |\s*Stock Price and Quote.*$", "",
                              soup.title.get_text(" ", strip=True)) if soup.title else ""}


def parse_news(soup, ref=None):
    """Titres d'actualité de la fiche Finviz -> [[date ISO, titre, source, lien], …] (les lignes sans date
    reprennent la date de la ligne précédente ; « Today » = aujourd'hui)."""
    tb = soup.select_one("#news-table") or soup.select_one("table.news-table")
    out, day = [], None
    ref = ref or TODAY
    for tr in tb.find_all("tr") if tb else []:
        tds = tr.find_all("td")
        a = tr.select_one("a.tab-link-news") or tr.find("a")
        if len(tds) < 2 or a is None:
            continue
        stamp = tds[0].get_text(" ", strip=True)
        m = re.match(r"(Today|[A-Z][a-z]{2}-\d\d-\d\d)?\s*(\d\d:\d\d[AP]M)", stamp)
        if not m:
            continue
        if m.group(1) == "Today":
            day = ref.date()
        elif m.group(1):
            day = datetime.strptime(m.group(1), "%b-%d-%y").date()
        if day is None:
            continue
        src = tr.find("span")
        src = src.get_text(" ", strip=True).strip("()") if src else ""
        t = datetime.combine(day, datetime.strptime(m.group(2), "%I:%M%p").time())
        href = a.get("href") or ""
        if href.startswith("/"):
            href = "https://finviz.com" + href
        out.append([t.isoformat(timespec="minutes"), a.get_text(" ", strip=True), src, href])
    return out


def parse_options_page(html):
    m = re.search(r'<script id="route-init-data" type="application/json">(.*?)</script>', html, re.S)
    return json.loads(m.group(1)) if m else None


def parse_dataroma(html):
    soup = BeautifulSoup(html, "html.parser")
    txt = re.sub(r"\s+", " ", soup.get_text(" "))
    out = {"n": None, "hold": None, "rows": []}
    m = re.search(r"Ownership count:\s*(\d+)", txt)
    out["n"] = int(m.group(1)) if m else 0
    m = re.search(r"Hold Price \* :\s*\$([\d.,]+)", txt)
    out["hold"] = num(m.group(1)) if m else None
    grid = soup.select_one("#grid")
    if grid:
        for tr in grid.select("tbody tr"):
            c = [td.get_text(" ", strip=True) for td in tr.find_all("td")]
            if len(c) >= 4:
                out["rows"].append({"Gérant": c[1].split(" - ")[0], "% portefeuille": num(c[2]),
                                    "Mouvement": c[3] or "=", "Actions": num(c[4]) if len(c) > 4 else None})
    return out


def agg_options(d, price):
    """Agrège une échéance (JSON route-init-data de la page Options Finviz)."""
    a = {"Échéance": d.get("currentExpiry"), "OI calls": 0, "OI puts": 0, "Vol calls": 0, "Vol puts": 0,
         "OI calls ±10 %": 0, "OI puts ±10 %": 0}
    for o in d.get("options", []):
        side = "calls" if o.get("type") == "call" else "puts"
        oi, vol = o.get("openInterest") or 0, o.get("lastVolume") or 0
        a[f"OI {side}"] += oi
        a[f"Vol {side}"] += vol
        if price and abs(o["strike"] / price - 1) <= 0.10:
            a[f"OI {side} ±10 %"] += oi
    return a


MONTHS_EN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def parse_openinsider(html):
    """Tableau OpenInsider -> lignes au format de la fiche Finviz
    [nom, fonction, « Sep 10 '26 », Buy/Sale, prix, titres, valeur]. None si pas de tableau."""
    tb = BeautifulSoup(html, "html.parser").select_one("table.tinytable")
    if tb is None:
        return None
    rows = []
    for tr in tb.find_all("tr")[1:]:
        c = [td.get_text(" ", strip=True) for td in tr.find_all("td")]
        if len(c) < 12:
            continue
        typ = "Buy" if c[6].startswith("P") else "Sale" if c[6].startswith("S") else None
        if not typ:
            continue
        try:
            d = datetime.strptime(c[2][:10], "%Y-%m-%d")
        except ValueError:
            continue
        price, qty, val = num(c[7]), abs(num(c[8]) or 0), abs(num(c[11]) or 0)
        rows.append([c[4], c[5] + (" · avec levée d'options" if "OE" in c[6] else ""),
                     f"{MONTHS_EN[d.month - 1]} {d.day:02d} '{d.year % 100:02d}", typ,
                     f"{price or 0:.2f}", f"{qty:,.0f}", f"{val:,.0f}", c[9]])
    return rows


def fetch_openinsider(t, s=None, days=365):
    """OpenInsider, 12 mois : http puis https, délai 40 s, deux essais (site parfois lent juste après un screener)."""
    s = s or _session()
    err = None
    for attempt in range(2):
        for base in ("http", "https"):
            try:
                r = s.get(f"{base}://openinsider.com/screener?s={t}&fd={days}&xp=1&xs=1&cnt=500", headers=UA,
                          timeout=40)
                r.raise_for_status()
                time.sleep(random.uniform(0.8, 1.5))
                return parse_openinsider(r.text)
            except Exception as e:  # noqa: BLE001
                err = e
        time.sleep(3)
    raise err


def parse_screener_rows(html):
    """Lignes d'une page de screener Finviz -> [{colonne: texte}] (attributs data-column-id)."""
    soup = BeautifulSoup(html, "html.parser")
    out = []
    for r in soup.select("div.fv-table_row[data-row-key]"):
        d = {x.get("data-column-id"): x.get_text(" ", strip=True) for x in r.select("[data-column-id]")}
        d["ticker"] = r.get("data-row-key").upper()
        out.append(d)
    return out


def fetch_peers(peers, s=None):
    """Sociétés comparables désignées par Finviz (lien « Peers ») : multiples en une requête (vue Valuation)."""
    if not peers:
        return []
    # vue personnalisée : capitalisation, PER, PER prévisionnel, P/FCF, croissance 5 ans, VE, VE/EBITDA
    rows = parse_screener_rows(_get(s or _session(), "https://finviz.com/screener.ashx?v=152&c=1,6,7,8,13,20,144,145&t="
                                    + ",".join(peers[:15]), pause=(0.5, 1.0)))
    return [{"ticker": r["ticker"], "fpe": num(r.get("forwardPE")), "pfcf": num(r.get("PFCF")),
             "pe": num(r.get("PE")), "g": num(r.get("estLTGrowth")), "mcap": num(r.get("marketCap")),
             "ev_ebitda": num(r.get("evEbitda"))} for r in rows]


def add_peers(raw, t, s=None, log=print):
    if raw.get("peers_data") is None and raw.get("peers"):
        try:
            raw["peers_data"] = fetch_peers(raw["peers"], s)
        except Exception as e:
            raw["peers_data"] = []
            log(f"{t} : sociétés comparables indisponibles ({e})")
    return raw


ONLINE_KEYS = ("IA", "BA", "CA", "options", "dataroma", "insiders", "insider_source", "peers_data", "oi_failed",
               "edgar")


def clear_caches():
    """Vide tous les caches : fichiers du dossier de cache, pages lues en mémoire, multiples d'industrie, Reddit."""
    n = 0
    if os.path.isdir(CACHE_DIR):
        for fn in os.listdir(CACHE_DIR):
            if fn.endswith(".json"):
                try:
                    os.remove(os.path.join(CACHE_DIR, fn))
                    n += 1
                except OSError:
                    pass
    for fn_ in (parse_zonebourse_cached, raws_from_finviz_pages_cached, fetch_apewisdom):
        try:
            fn_.clear()
        except Exception:
            pass
    _IND.update(data=None, t=0.0)
    return n


def complete_online_cached(raw, t, max_days=200, use_cache=True, log=print, openinsider=True):
    """Compléments en ligne d'une fiche Finviz déposée, mis en cache 2 h par valeur : redéposer ou ajouter des
    pages ne relance pas les téléchargements déjà faits (le cours et la fiche restent ceux de la page)."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    path = os.path.join(CACHE_DIR, f"{t}_complements{'_oi' if openinsider else ''}_{max_days}_v{INS_MERGE_VERSION}.json")
    if use_cache and os.path.exists(path) and time.time() - os.path.getmtime(path) < CACHE_HOURS * 3600:
        with open(path, encoding="utf-8") as fh:
            raw.update(json.load(fh))
        dirty = _ensure_edgar(raw, t, log)
        if openinsider and not raw.get("insider_source") and time.time() - raw.get("oi_failed", 0) > 600:
            add_openinsider(raw, t, None, log, openinsider)
            dirty = dirty or bool(raw.get("insider_source"))
        if dirty:
            with open(path, "w", encoding="utf-8") as fh:
                json.dump({k: raw.get(k) for k in ONLINE_KEYS if k in raw}, fh)
        return raw, True
    complete_online(raw, t, max_days, log, None, openinsider)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump({k: raw.get(k) for k in ONLINE_KEYS if k in raw}, fh)
    return raw, False


@st.cache_data(show_spinner=False, max_entries=300)
def parse_zonebourse_cached(htmls, name=""):
    """Lecture des pages Zonebourse, gardée en mémoire : une page déjà lue (même contenu) n'est pas relue."""
    return parse_zonebourse(list(htmls), name)


PAGE_PARSE_VERSION = 3          # à augmenter quand la lecture des pages change (invalide la mémoire)


@st.cache_data(show_spinner=False, max_entries=1000)
def finviz_page_cached(name, html, version=PAGE_PARSE_VERSION):
    """Lecture d'UNE page Finviz gardée en mémoire : ajouter une page ne relit pas les autres."""
    return finviz_page(name, html)


def raws_from_finviz_pages_cached(pages):
    return raws_from_finviz_pages(list(pages), page_fn=finviz_page_cached)


INS_MERGE_VERSION = 3       # change le nom des caches : les anciennes fusions OpenInsider + Finviz sont relues


def _name_key(n):
    return frozenset(w for w in re.findall(r"[a-z]+", (n or "").lower()) if len(w) > 1)


def _oi_covers(oi_rows, r):
    """Une transaction de la fiche Finviz est-elle déjà dans OpenInsider ? Même dirigeant (mêmes mots du nom, ordre
    et casse ignorés), même sens, à 3 jours près : OpenInsider regroupe en une ligne les ventes étalées sur plusieurs
    jours (ex. 25,8 + 25,6 M$ sur Finviz = 51,4 M$ sur OpenInsider), avec l'indication « levée d'options »."""
    def dt(x):
        try:
            return datetime.strptime(x.replace("'", "20"), "%b %d %Y")
        except ValueError:
            return None
    k, d = _name_key(r[0]), dt(r[2])
    for o in oi_rows:
        if o[3] == r[3] and _name_key(o[0]) == k:
            do = dt(o[2])
            if d is None or do is None or abs((do - d).days) <= 3:
                return True
    return False


def complete_online(raw, t, max_days=200, log=print, s=None, openinsider=True):
    """Ajoute en ligne ce qui manque : états financiers, initiés (OpenInsider), options, super investors."""
    s = s or _session()
    add_peers(raw, t, s, log)
    add_openinsider(raw, t, s, log, openinsider)
    _complete_rest(raw, t, max_days, log, s)


def _ensure_edgar(raw, t, log=print):
    """Séries annuelles des 10-K (SEC EDGAR) si l'option est activée et qu'elles manquent. -> True si ajoutées."""
    if not ES.enabled() or raw.get("edgar") is not None:
        return False
    try:
        raw["edgar"] = {str(y): v for y, v in ES.annual_facts(t).items()}
        return True
    except Exception as e:  # noqa: BLE001
        log(f"{t} : SEC EDGAR indisponible ({e})")
        return False


def _edgar_insiders(raw, t, log=print):
    """Secours quand OpenInsider ne répond pas : déclarations Form 4 lues directement à la SEC (option EDGAR)."""
    if not ES.enabled():
        return False
    try:
        rows = ES.form4_rows(t, log=log)
    except Exception as e:  # noqa: BLE001
        log(f"{t} : Form 4 de la SEC indisponibles ({e})")
        return False
    if rows is None:
        return False
    extra = [r for r in raw.get("insiders", []) if len(r) >= 7 and r[3] in ("Buy", "Sale") and not _oi_covers(rows, r)]
    raw["insiders"] = rows + extra
    raw["insider_source"] = f"SEC EDGAR (Form 4) + fiche Finviz, 12 mois ({len(rows) + len(extra)} transactions)"
    raw.pop("oi_failed", None)
    return True


def add_openinsider(raw, t, s=None, log=print, openinsider=True):
    """Initiés OpenInsider (12 mois) fusionnés avec la fiche Finviz ; en cas d'échec, la fiche Finviz reste."""
    if openinsider and not raw.get("insider_source"):
        s = s or _session()
        try:
            rows = fetch_openinsider(t, s)
            if rows:                                     # sinon (société non suivie) : on garde la fiche Finviz
                extra = [r for r in raw.get("insiders", []) if len(r) >= 7 and r[3] in ("Buy", "Sale")
                         and not _oi_covers(rows, r)]   # transactions de la fiche Finviz absentes d'OpenInsider
                raw["insiders"] = rows + extra
                raw["insider_source"] = f"OpenInsider + fiche Finviz, 12 mois ({len(rows) + len(extra)} transactions)"
            elif rows is not None:                       # réponse sans transaction : pas de nouvel essai
                raw["insider_source"] = "fiche Finviz (aucune transaction sur OpenInsider)"
            if rows is None:                             # page sans tableau (erreur du site) : nouvel essai plus tard
                if not _edgar_insiders(raw, t, log):
                    raw["oi_failed"] = time.time()
            else:
                raw.pop("oi_failed", None)
        except Exception as e:
            if _edgar_insiders(raw, t, log):
                log(f"{t} : OpenInsider indisponible, déclarations Form 4 lues à la SEC ({e})")
                return
            raw["oi_failed"] = time.time()               # nouvel essai au plus tôt dans 10 minutes
            log(f"{t} : OpenInsider indisponible, initiés Finviz utilisés ({e}) — nouvel essai au prochain affichage "
                "(après 10 min)")


def _complete_rest(raw, t, max_days, log, s):
    """États financiers, options, super investors… (tout sauf les initiés)."""
    for code in ("IA", "BA", "CA"):                      # compte de résultat, bilan, flux (annuels + TTM)
        if raw.get(code):
            continue
        try:
            raw[code] = json.loads(_get(s, f"https://finviz.com/api/statement?t={t}&so=F&s={code}", (0.5, 1.0)))["data"]
        except Exception as e:
            raw[code] = None
            log(f"{t} : état {code} indisponible ({e})")
    if not raw.get("options"):
        opts = []
        try:
            first = parse_options_page(_get(s, f"https://finviz.com/stock?t={t}&p=d&ty=oc"))
            price = (first or {}).get("lastClose") or num(raw.get("snap", {}).get("Price"))
            exps = [e for e in (first or {}).get("expiries", [])
                    if 0 <= (datetime.strptime(e, "%Y-%m-%d") - TODAY).days <= max_days]
            for e in exps:
                d = parse_options_page(_get(s, f"https://finviz.com/stock?t={t}&p=d&ty=oc&e={e}", (0.8, 1.6)))
                if d and d.get("currentExpiry") == e:
                    opts.append(agg_options(d, price))
        except Exception as e:
            log(f"{t} : options indisponibles ({e})")
        raw["options"] = opts
    if not (raw.get("dataroma") or {}).get("rows"):
        try:
            raw["dataroma"] = parse_dataroma(_get(s, f"https://www.dataroma.com/m/stock.php?sym={t}"))
        except Exception as e:
            raw["dataroma"] = {"n": 0, "hold": None, "rows": []}
            log(f"{t} : Dataroma indisponible ({e})")
    _ensure_edgar(raw, t, log)
    return raw


def _news_old(raw):
    """Fiche mise en cache par une version antérieure : titres d'actualité sans lien -> à relire."""
    n = raw.get("news") or []
    return bool(n) and all(len(x) < 4 for x in n)


def fetch_ticker(t, max_days=200, use_cache=True, log=print, openinsider=True):
    os.makedirs(CACHE_DIR, exist_ok=True)
    path = os.path.join(CACHE_DIR, f"{t}{'_oi' if openinsider else ''}_v{INS_MERGE_VERSION}.json")
    if use_cache and os.path.exists(path) and time.time() - os.path.getmtime(path) < CACHE_HOURS * 3600:
        with open(path, encoding="utf-8") as f:
            cached = json.load(f)
        if not _news_old(cached):
            dirty = _ensure_edgar(cached, t, log)
            if openinsider and not cached.get("insider_source") and time.time() - cached.get("oi_failed", 0) > 600:
                add_openinsider(cached, t, None, log, openinsider)
                dirty = dirty or bool(cached.get("insider_source"))
            if dirty:
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(cached, f)
            return cached
    s = _session()
    raw = {"ticker": t, "fetched": datetime.now().isoformat(timespec="minutes")}
    raw.update(parse_quote(_get(s, f"https://finviz.com/quote.ashx?t={t}&p=d")))
    complete_online(raw, t, max_days, log, s, openinsider)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(raw, f)
    return raw


# ============================================================
#  PAGES ENREGISTRÉES (valeurs à étudier)
# ============================================================
def page_kind(name, html):
    head = (name + html[:6000]).lower()
    if "zonebourse" in head:
        return "zonebourse"
    if "finviz" in head:
        return "finviz"
    return None


def finviz_page(name, html):
    """Une page Finviz -> ("opt", ticker, options) · ("quote", ticker, fiche) · None."""
    m = re.search(r"<title>\s*([A-Z][A-Z0-9.\-]{0,9})\s+-\s", html)
    t = m.group(1) if m else os.path.basename(name).split("_")[0].split(" ")[0].upper()
    d = parse_options_page(html)
    if d and d.get("options"):
        return ("opt", t, d)
    q = parse_quote(html)
    return ("quote", t, q) if q["snap"].get("Price") else None


def raws_from_finviz_pages(pages, page_fn=None):
    """pages : liste de (nom, html) Finviz -> {ticker: raw}. Fiche « quote » et pages « Options ».
    page_fn : lecture d'une page (version mise en cache : chaque page n'est lue qu'une fois)."""
    raws, opts = {}, {}
    for name, html in pages:
        part = (page_fn or finviz_page)(name, html)
        if not part:
            continue
        kind, t, d = part
        if kind == "opt":
            opts.setdefault(t, []).append(copy.deepcopy(d))
            continue
        raws[t] = {"ticker": t, "source": "page Finviz", **copy.deepcopy(d), "options": [], "dataroma": None}
    for t, ds in opts.items():
        if t in raws:
            price = num(raws[t]["snap"].get("Price"))
            seen = {}
            for d in ds:
                seen[d.get("currentExpiry")] = agg_options(d, price or d.get("lastClose"))
            raws[t]["options"] = list(seen.values())
    return raws


# industries Zonebourse cycliques par nature (semi-conducteurs, matériel/mémoire, auto, pétrole, matériaux…) :
# pas d'exception « redressement durable » même si le consensus reste haussier
CYCLICAL_ZB = re.compile(r"materiel-informatique|semi|automobile|combustible-fossile|materiaux|chimi|metau|mine|"
                         r"mineraux|acier|papier|maritime|transport-aerien")


def num_fr(v):
    """Nombre au format français de Zonebourse : '28 905' ; '-612,8' ; '1,19x' ; '12,4 %'."""
    if v is None:
        return None
    s = re.sub(r"[\s\u202f\xa0]", "", str(v)).replace("%", "").rstrip("x").replace(",", ".")
    if s in ("", "-", "--", "-.--"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def zb_company(html):
    """Nom de la société d'après le titre de la page Zonebourse (sert à regrouper Finances + Valorisation)."""
    m = re.search(r"<title>(.*?)</title>", html, re.S)
    return re.split(r"\s*:", m.group(1).strip())[0].strip() if m else ""


def _isin_ok(code):
    """Clé de contrôle d'un code ISIN (algorithme de Luhn sur les lettres converties en chiffres)."""
    digits = "".join(str(int(c, 36)) for c in code)
    tot = 0
    for i, ch in enumerate(reversed(digits)):
        n = int(ch) * (2 if i % 2 else 1)
        tot += n - 9 if n > 9 else n
    return tot % 10 == 0


def zb_isin(html):
    """Code ISIN de la société (le plus fréquent des codes valides de la page Zonebourse), ou "" ."""
    from collections import Counter
    c = Counter(m for m in re.findall(r"\b([A-Z]{2}[A-Z0-9]{9}[0-9])\b", html) if _isin_ok(m))
    return c.most_common(1)[0][0] if c else ""


def page_identity(name, html):
    """Identité d'une page déposée : même société + même onglet = même page, quel que soit le nom du fichier
    (« … (1).html », version plus récente…)."""
    kind = page_kind(name, html)
    m = re.search(r"<title>(.*?)</title>", html, re.S)
    title = re.sub(r"\s+", " ", m.group(1)).strip() if m else name
    if kind == "zonebourse":
        onglet = "valorisation" if "Ratios de Valorisation" in title else "finances" if "Prévisions" in title else title
        return ("zb", zb_company(html), onglet)
    if kind == "finviz":
        return ("fv", title, hashlib.md5(html[:2000].encode("utf-8", "ignore")).hexdigest()
                if "Options" in title else "")
    return ("autre", name)


def dedupe_pages(pages):
    """Garde une seule version de chaque page (la plus récente selon l'heure de cotation Zonebourse, sinon la
    dernière déposée) -> (pages, noms ignorés)."""
    best, order = {}, []
    for i, (name, html) in enumerate(pages):
        k = page_identity(name, html)
        rank = (zb_stamp(html), i)
        if k not in best:
            order.append(k)
            best[k] = (rank, name, html, [])
        elif rank > best[k][0]:
            best[k] = (rank, name, html, best[k][3] + [best[k][1]])
        else:
            best[k][3].append(name)
    kept = [(best[k][1], best[k][2]) for k in order]
    dups = [n for k in order for n in best[k][3]]
    return kept, dups


def zb_stamp(html):
    """Horodatage de la cotation dans l'en-tête Zonebourse (« 16:39:53 27/10/2026 ») -> (aaaa, mm, jj, hh:mm:ss)."""
    m = re.search(r"(\d{2}:\d{2}:\d{2})\s+(\d{2})/(\d{2})/(\d{4})", html)
    return (m.group(4), m.group(3), m.group(2), m.group(1)) if m else ("", "", "", "")


def zb_peers(soups, company=""):
    """Tableau des sociétés comparables de l'onglet Valorisation (colonnes « PER (N) », « VE / CA (N) »,
    « VE / EBITDA (N) », ou « (Y) » sur les pages plus anciennes) -> (pairs, ligne de la société)."""
    def x_(t):
        t = (t or "").replace("\xa0", " ").replace("x", "").replace(" ", "").replace(",", ".")
        try:
            return float(t)
        except ValueError:
            return None
    for sp in soups:
        for tb in sp.find_all("table"):
            trs = tb.find_all("tr")
            if not trs:
                continue
            hdr = [c.get_text(" ", strip=True) for c in trs[0].find_all(["th", "td"])]
            ip = next((i for i, h in enumerate(hdr) if re.match(r"PER \((N|Y)\)", h)), None)
            ie = next((i for i, h in enumerate(hdr) if re.match(r"VE / EBITDA \((N|Y)\)", h)), None)
            ic = next((i for i, h in enumerate(hdr) if re.match(r"VE / CA \((N|Y)\)", h)), None)
            if ip is None or ie is None:
                continue
            rows = []
            for tr in trs[1:]:
                c = [td.get_text(" ", strip=True) for td in tr.find_all(["th", "td"])]
                if len(c) <= max(ip, ie) or not c[1] or c[1].startswith("Moyenne"):
                    continue
                rows.append({"name": c[1], "pe": x_(c[ip]), "ev_ebitda": x_(c[ie]),
                             "ev_sales": x_(c[ic]) if ic is not None and len(c) > ic else None})
            if not rows:
                continue
            key = re.sub(r"[^a-z]", "", company.lower())[:8]
            i_self = next((i for i, r in enumerate(rows) if key and key in re.sub(r"[^a-z]", "", r["name"].lower())), 0)
            return [r for i, r in enumerate(rows) if i != i_self], rows[i_self]
    return [], {}


def parse_zonebourse(htmls, name=""):
    """Onglets « Finances » (et « Valorisation ») d'une valeur Zonebourse -> faits normalisés
    (montants en millions de la devise). htmls : une page ou une liste de pages de la même société."""
    htmls = [htmls] if isinstance(htmls, str) else list(htmls)
    htmls.sort(key=zb_stamp, reverse=True)       # même page déposée deux fois : la plus récente est lue en premier
    soups = [BeautifulSoup(h, "html.parser") for h in htmls]
    titles = [sp.title.get_text(" ", strip=True) if sp.title else "" for sp in soups]
    txt = " ".join(re.sub(r"\s+", " ", sp.get_text(" ")) for sp in soups)
    m = next((m for m in (re.search(r"\|\s*([A-Z0-9.\-]{1,12})\s*\|\s*Zonebourse", t) for t in titles) if m), None)
    title = titles[0]
    tk = m.group(1) if m else (re.split(r"\s*:", title)[0].strip() or os.path.splitext(os.path.basename(name))[0])[:12].upper()
    T = {}
    for tb in (t for sp in soups for t in sp.find_all("table")):
        trs = tb.find_all("tr")
        if not trs:
            continue
        hdr = [c.get_text(" ", strip=True) for c in trs[0].find_all(["th", "td"])]
        yrs = hdr[1:]
        if not hdr or not hdr[0].startswith("Période Fiscale") or not yrs or \
                not all(re.fullmatch(r"\d{4}", y) for y in yrs):
            continue
        for tr in trs[1:]:
            cells = [c.get_text(" ", strip=True) for c in tr.find_all(["th", "td"])]
            lab = re.sub(r"\s+\d$", "", cells[0]).strip() if cells else ""
            if len(cells) >= 2 and lab and lab != "Variation" and lab not in T:
                T[lab] = {int(y): v for y, v in zip(yrs, cells[1:])}

    # devise : Zonebourse publie le consensus dans la devise de publication de la société (argenx : USD) même
    # sur la page de la cotation européenne ; le cours en temps réel de l'en-tête est dans la devise de cotation
    price_m = re.search(r"Dernier Cours de Cl[oô]ture\s*([\d ]+,\d+)\s*([A-Z]{3})", txt)
    live_m = re.search(r"Varia\. 1 janv\.\s*([\d ]+,\d+)\s*([A-Z]{3})", txt)
    cur_cons = price_m.group(2) if price_m else None
    fx, fx_note = 1.0, None
    if price_m and live_m and live_m.group(2) != cur_cons and num_fr(price_m.group(1)) and num_fr(live_m.group(1)):
        fx = num_fr(live_m.group(1)) / num_fr(price_m.group(1))
        fx_note = (f"Consensus Zonebourse publié en {cur_cons}, cotation en {live_m.group(2)} : montants convertis "
                   f"au taux implicite 1 {live_m.group(2)} = {1 / fx:.3f} {cur_cons} (cours {live_m.group(1)} "
                   f"{live_m.group(2)} / {price_m.group(1)} {cur_cons}).")
    NO_FX = ("PER", "PBR", "ROE", "Nbr de Titres")

    def ser(*labels):
        for lab in labels:
            for k, v in T.items():
                if k.lower().startswith(lab.lower()):
                    k_fx = 1.0 if lab in NO_FX else fx
                    return {y: (num_fr(x) * k_fx if num_fr(x) is not None else None) for y, x in v.items()}
        return {}
    pub = T.get("Date de publication", {})
    est = sorted(y for y, v in pub.items() if v.strip() in ("-", ""))
    act = sorted(y for y, v in pub.items() if v.strip() not in ("-", ""))
    bna = ser("BNA")
    if not est or not act or not bna:
        raise ValueError("tableaux annuels introuvables : enregistrez l'onglet « Finances » de la valeur")
    N, L = est[0], act[-1]
    ca, eb, capex, fcf = ser("Chiffre d'affaires"), ser("EBITDA"), ser("CAPEX"), ser("Flux de trésorerie libre")
    nd, cash = ser("Endettement Net"), ser("Trésorerie Nette")
    net_debt = {y: (nd.get(y) if nd.get(y) is not None else (-cash[y] if cash.get(y) is not None else None))
                for y in set(nd) | set(cash)}
    titres, div = ser("Nbr de Titres"), ser("Dividende / Action")
    bvps, roe, rn = ser("Actif net par Action"), ser("ROE"), ser("Résultat net")
    g = lambda d, y: d.get(y)
    # cours : temps réel de l'en-tête (au moment de l'enregistrement), sinon dernière clôture du consensus
    price = num_fr(live_m.group(1)) if live_m and (fx != 1.0 or live_m.group(2) == cur_cons) else \
        num_fr(price_m.group(1)) if price_m else ser("Cours de référence").get(N)
    tgt = re.search(r"Objectif de cours Moyen\s*([\d ]+,\d+)", txt)
    rec = re.search(r"Recommandation moyenne\s*([A-ZÉÈ ]+?)\s+Nombre", txt)
    nba = re.search(r"Nombre d.Analystes\s*(\d+)", txt)
    shares = (g(titres, N) or g(titres, L) or 0) / 1e3          # milliers -> millions
    e0, e1, e2, eL = g(bna, N), g(bna, N + 1), g(bna, N + 2), g(bna, L)
    f = {"ticker": tk, "name": title.split(":")[0].strip(), "currency": (live_m.group(2) if fx != 1.0 else cur_cons) if price_m else "EUR",
         "price": price, "target": num_fr(tgt.group(1)) * fx if tgt else None, "recom": None,
         "consensus_label": rec.group(1).strip().capitalize() if rec else None,
         "n_analysts": int(nba.group(1)) if nba else None,
         "pe": None, "fwd_pe": price / e1 if price and e1 else None, "eps_ttm": eL,
         "eps_this_y": e0, "eps_next_y": e1,
         "eps_next_y_pct": (e1 / e0 - 1) * 100 if e0 and e1 and e0 > 0 else None,
         "eps_this_y_pct": (e0 / eL - 1) * 100 if e0 and eL and eL > 0 else None,
         "eps_5y": (((e2 / eL) ** (1 / 3) - 1) * 100 if eL and e2 and eL > 0 and e2 > 0 else
                    ((e2 / e0) ** 0.5 - 1) * 100 if e0 and e2 and e0 > 0 and e2 > 0 else None),
         "sales_yoy": (g(ca, N) / g(ca, L) - 1) * 100 if g(ca, N) and g(ca, L) else None,
         "rsi": None, "shares": shares, "mcap": price * shares if price else 0, "ev": None,
         "earnings": None, "div_ps": g(div, L) or g(div, N) or 0.0, "short": None,
         "rev": g(ca, L), "ebitda": g(eb, L), "sbc": 0.0, "cfo": None,
         "capex": abs(g(capex, L)) if g(capex, L) is not None else None,
         "fcf_hist": [x for x in (g(fcf, N), g(fcf, L), g(fcf, L - 1), g(fcf, L - 2)) if x is not None],
         "buyback": 0.0, "net_debt": g(net_debt, L) if g(net_debt, L) is not None else g(net_debt, N),
         "insiders": None,
         "dataroma": {"n": None, "rows": []}, "ratings": [], "options": [],
         "zb": {"N": N, "L": L, "bna": bna, "ebitda": eb, "capex": capex, "fcf": fcf, "nd": net_debt,
                "div": div, "titres": titres, "bvps": bvps, "roe": roe, "rn": rn}}
    f["fx_note"] = fx_note
    st_ = zb_stamp(htmls[0])                       # date de la page la plus récente (en-tête Zonebourse)
    f["zb_date"] = f"{st_[2]}/{st_[1]}/{st_[0]}" if st_[0] else None
    # multiples historiques (onglet Valorisation) : PER médian des exercices publiés
    pe_h = sorted(v for y, v in ser("PER").items() if y in act and v and 0 < v < 60)
    f["pe_hist"] = pe_h[len(pe_h) // 2] if len(pe_h) >= 3 else None
    pb_h = sorted(v for y, v in ser("PBR").items() if y in act and v and v > 0)
    f["pb_hist"] = pb_h[len(pb_h) // 2] if len(pb_h) >= 3 else None
    f["name"] = re.split(r"\s*:", title)[0].strip()
    # objectif « façon analystes » : VE/EBIT moyen des 5 derniers exercices publiés × EBIT N+1 (onglet Valorisation)
    _ve = {y: num_fr(v) for y, v in (T.get("Valeur Entreprise / EBIT") or {}).items()}
    _eb = {y: (num_fr(v) * fx if num_fr(v) is not None else None)
           for y, v in (T.get("EBIT") or T.get("Résultat d'exploitation (EBIT)") or {}).items()}
    f["zb"]["ev_ebit_hist"] = [_ve[y] for y in act[-5:] if _ve.get(y) and 0 < _ve[y] < 80]
    # pairs (onglet Valorisation) : PER, VE/CA, VE/EBITDA de l'exercice en cours ; la 1re ligne est la société
    f["zb_peers"], f["zb_self"] = zb_peers(soups, f.get("name") or "")
    f["zb"]["ebit"] = _eb
    # banque / assurance : pas d'EBITDA ni de flux libre, mais un actif net par action et un ROE
    has_eb = any(v is not None for v in eb.values())
    has_fcf = any(v is not None for v in fcf.values())
    # banque / assurance : ni EBITDA ni flux libre publiés (une industrielle sans EBITDA garde son flux libre)
    f["bank"] = not has_eb and not has_fcf
    # banques et assureurs d'après le secteur Zonebourse : leurs flux et leur « dette » viennent du bilan
    # (dépôts, placements, réserves techniques) -> pas de DCF d'entreprise, modèle financier (P/B justifié + PER)
    sec0 = re.search(r"/secteurs/([a-z0-9\-]+/[a-z0-9\-]+)/", "".join(htmls))
    sec0 = sec0.group(1) if sec0 else ""
    f["insurer"] = "assurance" in sec0
    f["bank"] = bool(f["bank"] or f["insurer"] or "services-bancaires" in sec0)
    if not has_eb and has_fcf:                        # EBITDA absent (ex. constructeur avec filiale de crédit, biotech) :
        ebit = ser("Résultat d'exploitation")          # résultat d'exploitation utilisé à la place
        if any(v is not None for v in ebit.values()):
            f["zb"]["ebitda"], f["ebitda"] = ebit, g(ebit, L)
            f["ebitda_proxy"] = True
    f["missing_finances"] = not price_m                   # onglet Finances absent : pas de consensus
    if f["bank"]:
        f["capex"], f["fcf_hist"] = None, []          # « capex » d'une banque = actifs loués : sans objet
    # cyclique : BNA de l'année en hausse > 150 %, ou perte réapparue après un exercice bénéficiaire (ex. automobile) ;
    # une sortie de pertes continue (biotech qui devient rentable) n'est pas cyclique
    hist_eps = [bna[y] for y in act[-6:] if bna.get(y) is not None]
    yrs = [y for y in act[-6:] if bna.get(y) is not None]
    # retour à la perte « cyclique » seulement si l'activité recule aussi (EBITDA en baisse d'au moins 20 % par rapport
    # au dernier exercice bénéficiaire) ; une perte purement comptable (amortissements d'acquisition, charges de
    # fusion, dépréciations) avec un EBITDA stable ou en hausse n'est pas un signe de cycle (ex. Coherent 2023-2025)
    def _relapse(i_, j_):
        e_a, e_b = g(eb, yrs[i_]), g(eb, yrs[j_])
        return e_a is None or e_b is None or e_a <= 0 or e_b < 0.8 * e_a
    relapse = any(bna[yrs[i_]] > 0 and bna[yrs[j_]] < 0 and _relapse(i_, j_)
                  for i_ in range(len(yrs)) for j_ in range(i_ + 1, len(yrs)))
    # scission / cession (ex. Aptiv 2026) : CA prévu N < 80 % du dernier exercice publié alors que la marge d'EBITDA
    # tient et que le CA repart en N+1 -> le dernier exercice n'est pas comparable : base = consensus N (pro forma)
    mL = g(eb, L) / g(ca, L) if g(eb, L) is not None and g(ca, L) else None
    mN = g(eb, N) / g(ca, N) if g(eb, N) is not None and g(ca, N) else None
    f["spin"] = bool(g(ca, L) and g(ca, N) and g(ca, N + 1) and g(ca, N) < 0.8 * g(ca, L) and g(ca, N + 1) > g(ca, N)
                     and mL and mN and mL > 0 and mN >= 0.9 * mL and not f.get("bank"))
    if f["spin"]:
        sp_ = lambda v: f"{v:,.0f}".replace(",", " ")
        f["spin_note"] = (f"⚠️ scission ou cession détectée : CA prévu {N} {sp_(g(ca, N))} contre {sp_(g(ca, L))} en {L} "
                          f"(marge d'EBITDA maintenue) → DCF, croissance et dette nette sur la base pro forma {N} ; "
                          f"bond du BNA {L}→{N} non pris comme signal cyclique")
        f["spin_short"] = f"scission/cession détectée : DCF sur la base pro forma {N}"
        f["eps_this_y_pct"] = None                   # BNA publié et BNA pro forma non comparables
    f["cyclical"] = bool((f["eps_this_y_pct"] or 0) > 150 or relapse) and not f["bank"]
    # exception : redressement / croissance durable — le consensus prévoit encore un BNA en hausse ≥ 10 % par an en
    # N+1 et N+2, une marge d'EBITDA qui progresse jusqu'en N+2 et un CA en croissance ≥ 5 % par an (un pic de cycle
    # se lit au contraire par un BNA qui plafonne ou recule après N)
    mg = [g(eb, y) / g(ca, y) if g(eb, y) is not None and g(ca, y) else None for y in (N, N + 1, N + 2)]
    sec_m = re.search(r"/secteurs/([a-z0-9\-]+/[a-z0-9\-]+)/", "".join(htmls))
    f["zb_sector"] = sec_m.group(1) if sec_m else ""
    f["isin"] = zb_isin("".join(htmls))
    f["turnaround"] = bool(
        f["cyclical"] and not CYCLICAL_ZB.search(f["zb_sector"]) and e0 and e1 and e2 and e0 > 0 and e1 >= 1.10 * e0 and e2 >= 1.10 * e1
        and None not in mg and mg[0] < mg[1] < mg[2]
        and g(ca, N) and g(ca, N + 2) and (g(ca, N + 2) / g(ca, N)) ** 0.5 - 1 >= 0.05)
    # redressement après une crise (pertes ou BNA quasi nul) dans un secteur non cyclique (santé, services…) : le saut
    # du BNA vient de la sortie de crise, pas d'un cycle ; le consensus prévoit un BNA qui continue de monter
    pub = [bna[y] for y in act[-4:] if bna.get(y) is not None]
    f["recovery"] = bool(
        f["cyclical"] and f["zb_sector"] and not CYCLICAL_ZB.search(f["zb_sector"]) and pub and e1 and e1 > 0
        and min(pub[-3:]) <= 0.1 * e1 and (e0 is None or e1 > e0) and (e2 is None or e2 >= 0.95 * e1))
    if f["recovery"] and not f["turnaround"]:
        f["turnaround"] = True
        f["recovery_note"] = (f"redressement après une crise (BNA publiés {', '.join(f'{v:.2f}' for v in pub[-3:])}, "
                              f"consensus N+1 {e1:.2f}) dans un secteur non cyclique : non classé cyclique")
    if f["turnaround"]:
        f["cyclical"] = False
    # croissance du CA (dernier exercice publié -> N+2) : plafonne la croissance du DCF
    c2 = g(ca, N + 2) or g(ca, N + 1)
    f["sales_cagr"] = ((c2 / g(ca, L)) ** (1 / (3 if g(ca, N + 2) else 2)) - 1) * 100 if c2 and g(ca, L) else None
    if f["spin"]:                                    # croissance mesurée depuis la base pro forma N
        f["sales_cagr"] = ((c2 / g(ca, N)) ** (1 / (2 if g(ca, N + 2) else 1)) - 1) * 100
    # flux libre irrégulier (acomptes clients, grands contrats…) : flux normalisé = médiane publiée + prévue
    act_fcf = [v for y, v in fcf.items() if y in act and v is not None][-5:]
    if len(act_fcf) >= 3:
        mean = sum(act_fcf) / len(act_fcf)
        cv = (sum((v - mean) ** 2 for v in act_fcf) / len(act_fcf)) ** 0.5 / abs(mean) if mean else 9
        fut = [v for y, v in fcf.items() if y > act[-1] and v is not None]
        ramp = act_fcf[0] < 0 < act_fcf[-1] and fut and min(fut) > 0     # sortie de pertes (biotech, croissance) :
        if cv > 0.35 and not ramp:                                        # pas de normalisation, le flux prévu sert de base
            allv = sorted(v for y, v in fcf.items() if v is not None and y >= act[-1] - 4)
            m = len(allv) // 2
            norm = allv[m] if len(allv) % 2 else (allv[m - 1] + allv[m]) / 2
            eb1 = g(eb, N + 1) or g(eb, N)
            f["fcf_norm"] = min(norm, 0.8 * eb1) if eb1 else norm
    try:
        a_, e_ = (list(act) + [N], [y for y in est if y != N]) if f["spin"] else (act, est)   # N = base pro forma
        f["dcfx"] = None if f.get("bank") else dcfx_from_zonebourse(
            f["zb"], ca, eb, ser("Résultat d'exploitation"), capex, net_debt, a_, e_, shares, fcf)
        if f["spin"] and f["dcfx"]:
            f["dcfx"]["src"] += f" — base pro forma {N}"
            if g(net_debt, N) is not None:
                f["net_debt"] = g(net_debt, N)
    except Exception:
        f["dcfx"] = None
    return f


# ============================================================
#  FAITS NORMALISÉS (montants en millions de $)
# ============================================================
DEFENSIVE_SECTORS = {"Utilities", "Healthcare", "Consumer Defensive", "Communication Services", "Real Estate"}


def build_facts(raw):
    sn = raw.get("snap", {})
    IA, BA, CA = raw.get("IA"), raw.get("BA"), raw.get("CA")
    g = lambda k: num(sn.get(k))
    f = {"price": g("Price"), "target": g("Target Price"), "recom": g("Recom"),
         "pe": g("P/E"), "fwd_pe": g("Forward P/E"), "eps_ttm": g("EPS (ttm)"),
         "eps_next_y": g("EPS next Y"), "eps_next_y_pct": g("EPS next Y %"), "eps_this_y_pct": g("EPS this Y"),
         "eps_5y": g("EPS next 5Y"), "sales_yoy": g("Sales Y/Y TTM"), "rsi": g("RSI (14)"),
         "mcap": (g("Market Cap") or 0) / 1e6, "ev": (g("Enterprise Value") or 0) / 1e6,
         "shares": (g("Shs Outstand") or 0) / 1e6, "earnings": sn.get("Earnings"),
         "div_ps": g("Dividend TTM") or 0.0, "short": g("Short Float"), "beta": g("Beta")}
    # actions à plusieurs catégories (Alphabet A/C…) : le nombre d'actions cohérent avec la capitalisation
    if f["mcap"] and f["price"] and f["shares"] and abs(f["mcap"] / f["price"] / f["shares"] - 1) > 0.2:
        f["shares"] = f["mcap"] / f["price"]
    f["rev"], f["ebitda"] = stmt(IA, "Total Revenue"), stmt(IA, "EBITDA")
    # BNA publié des exercices clos (comptable), du plus ancien au plus récent
    if IA and IA.get("Period") and IA.get("EPS (Diluted)"):
        f["eps_hist"] = [(str(pp).replace("FY", ""), num(v)) for pp, v in zip(IA["Period"], IA["EPS (Diluted)"])
                         if "FY" in str(pp) and num(v) is not None][::-1]
    try:
        f["dcfx"] = dcfx_from_finviz(IA, BA, CA, sn)
    except Exception:
        f["dcfx"] = None
    if f["dcfx"] and cash_unrepresentative(f, IA, CA, g):
        # flux d'exploitation faussés (créances cédées classées en investissement, appels de marge…) : le DCF
        # repart du résultat d'exploitation (NOPAT + amortissements − investissements), BFR standard
        f["dcfx"] = dict(f["dcfx"], ocf0=None, ocf_hist=[], nwc_k=0.05,
                         src=f["dcfx"]["src"] + ", flux reconstruits depuis le résultat")
    f["sbc"] = stmt(IA, "Stock Option Compensation Expense") or 0.0
    f["cfo"], f["capex"] = stmt(CA, "Cash from Operating Activities"), abs(stmt(CA, "Capital Expenditures") or 0)
    f["fcf_hist"] = [x for x in (stmt(CA, "Free Cash Flow", i) for i in range(4)) if x is not None]
    f["buyback"] = abs(stmt(CA, "Repurchase of Common Pref Stock") or 0)
    f["cap_hist"] = capital_hist_finviz(IA, CA)
    f["edgar"] = {int(k): v for k, v in (raw.get("edgar") or {}).items()} or None
    f["net_debt"] = f["ev"] - f["mcap"] if f["ev"] and f["mcap"] else None
    if not CA:                                             # page enregistrée sans états financiers
        f["capex"] = None
        if not f["fcf_hist"] and g("P/FCF"):
            f["fcf_hist"] = [f["mcap"] / g("P/FCF")]
    f["cash_unrep"] = cash_unrepresentative(f, IA, CA, g)
    lh = long_history(f) if f["edgar"] and f.get("dcfx") else None   # appliqué si la valeur est cyclique
    if lh:
        f["dcfx_long"], f["eps_hist_long_all"], f["long_note"] = lh
    if f["rev"] is None and g("Sales"):
        f["rev"] = g("Sales") / 1e6
    if f["ebitda"] is None and f["ev"] and g("EV/EBITDA"):
        f["ebitda"] = f["ev"] / g("EV/EBITDA")
    # BNA de l'exercice en cours, déduit du BNA N+1 et de sa croissance
    if f["eps_next_y"] and f["eps_next_y_pct"] is not None:
        f["eps_this_y"] = f["eps_next_y"] / (1 + f["eps_next_y_pct"] / 100)
    else:
        f["eps_this_y"] = None
    # valeur cyclique / bénéfices de pic (détection auto, modifiable dans la barre latérale)
    # bond > 150 % du BNA = pic de cycle (semi-conducteurs, matériaux, énergie, auto…) ; dans un secteur défensif
    # (services publics, santé, conso de base…) c'est un rebond depuis une base déprimée, sauf si le BNA
    # plafonne l'an prochain (ex. Vistra : +296 % puis +18 % -> non cyclique)
    f["sector"] = raw.get("sector")
    f["industry"] = raw.get("industry") or ""
    f["peers"] = [x for x in (raw.get("peers_data") or []) if x["ticker"] != raw.get("ticker")]
    # croissance du CA « durable » : médiane (3 ans, 5 ans, 12 derniers mois)
    sp = str(sn.get("Sales past 3/5Y") or "").split()
    sales_g = [x for x in (num(sp[0]) if sp else None, num(sp[1]) if len(sp) > 1 else None, f["sales_yoy"])
               if x is not None]
    f["sales_g"] = sorted(sales_g)[len(sales_g) // 2] if sales_g else None
    f["bvps"], f["roe_ttm"] = g("Book/sh"), g("ROE")
    # banque / courtier / assureur : pas d'EBITDA ni de flux libre significatif -> modèle bancaire (actif net, PER)
    f["bank"] = f["sector"] == "Financial" and bool(
        re.search(r"Bank|Capital Markets|Insurance|Mortgage", f["industry"]) or f["ebitda"] is None)
    f["insurer"] = f["bank"] and "Insurance" in f["industry"]
    jump = (f["eps_this_y_pct"] or 0) > 150
    defensive = f["sector"] in DEFENSIVE_SECTORS
    f["cyclical"] = bool(jump and (not defensive or (f["eps_next_y_pct"] is not None and f["eps_next_y_pct"] < 10))) \
        and not f["bank"]
    # PER prévisionnel < 7 : le marché juge les bénéfices de l'an prochain non durables (pic de cycle, ex. mémoire)
    f["low_fpe"] = bool(f["fwd_pe"] and 0 < f["fwd_pe"] < 7 and not f["bank"])
    if f["low_fpe"]:
        f["cyclical"] = True
    f["insiders"] = insider_stats(raw.get("insiders", []))
    f["currency"] = "USD"
    f["source"] = raw.get("source", "Finviz en ligne")
    f["insider_source"] = raw.get("insider_source", "fiche Finviz")
    f["news"] = raw.get("news", [])
    f["company"] = raw.get("company", "")
    f["dataroma"] = raw.get("dataroma") or {"n": None, "rows": []}
    f["ratings"] = raw.get("ratings", [])
    f["options"] = raw.get("options", [])
    return f


def merge_facts(fv, zb):
    """Action US avec fiche Finviz ET pages Zonebourse : Finviz pour le marché (cours, bêta, initiés, options,
    actualités, analystes) ; Zonebourse pour les prévisions pluriannuelles (BNA, EBITDA, capex, flux, dette),
    le PER historique et la normalisation du flux. BNA homogène avec chaque multiple : BNA Finviz (consensus ajusté,
    même base que les PER prévisionnels) pour le PER cible et les comparables ; BNA Zonebourse pour le PER historique
    (même base que les PER publiés par Zonebourse) ; jamais de moyenne des deux."""
    f = dict(fv)
    notes = []
    f["merged"] = True
    f["bank"] = bool(fv.get("bank") or zb.get("bank"))
    f["insurer"] = bool(fv.get("insurer") or zb.get("insurer"))
    for k in ("eps_this_y", "eps_next_y"):
        a, b = fv.get(k), zb.get(k)
        if not (a and a > 0) and b:
            f[k] = b                                      # pas de consensus Finviz : Zonebourse
    a, b = fv.get("eps_next_y"), zb.get("eps_next_y")
    if b and b > 0:
        f["eps_next_y_zb"] = b
        if a and a > 0 and abs(a / b - 1) > 0.03:
            notes.append(f"BNA N+1 : Finviz {a:.2f} (consensus ajusté) pour le PER cible et les comparables · "
                         f"Zonebourse {b:.2f} (écart {(a / b - 1) * 100:+.0f} %) pour le PER historique")
    if f.get("eps_this_y") and f.get("eps_next_y") and f["eps_this_y"] > 0:
        f["eps_next_y_pct"] = (f["eps_next_y"] / f["eps_this_y"] - 1) * 100
    for k in ("pe_hist", "pb_hist", "sales_cagr", "fcf_norm", "zb", "n_analysts", "zb_date", "zb_peers", "zb_self"):
        if zb.get(k) is not None:
            f[k] = zb[k]
    if f.get("capex") is None and zb.get("capex") is not None:      # page Finviz sans états financiers
        f["capex"], f["fcf_hist"] = zb["capex"], zb["fcf_hist"]
    if not f.get("dataroma"):
        f["dataroma"] = {"n": None, "rows": []}
    f["cyclical"] = bool(fv.get("cyclical") or zb.get("cyclical"))
    if fv.get("dcfx") and (zb.get("dcfx") or {}).get("fc"):          # DCF Finviz + consensus Zonebourse (CA, marges, capex)
        fc = dict(zb["dcfx"]["fc"])
        revs = fc.get("rev") or []
        x0, zx = fv["dcfx"], zb["dcfx"]
        if revs and x0["rev0"] and zx.get("rev0"):
            # calendrier : le CA des 12 derniers mois Finviz se situe entre le dernier exercice clos (L) et
            # l'exercice N ; on interpole le consensus (en log) aux dates « 12 derniers mois + 1, + 2… an(s) »
            import math
            pts = [zx["rev0"]] + revs                               # L, N, N+1, N+2
            q = 0.0
            if pts[1] > pts[0] > 0 and x0["rev0"] > pts[0]:
                q = clamp(math.log(x0["rev0"] / pts[0]) / math.log(pts[1] / pts[0]), 0.0, 1.0)
            ms = [zx["m0"]] + [m if m is not None else zx["m0"] for m in fc["m"]]

            def at(arr, tpos, log=True):                            # tpos : années depuis L
                i = min(int(tpos), len(arr) - 2)
                w_ = tpos - i
                a, b = arr[i], arr[i + 1]
                return math.exp(math.log(a) * (1 - w_) + math.log(b) * w_) if log and a > 0 and b > 0 \
                    else a * (1 - w_) + b * w_
            n = len(revs) - (1 if q > 0 else 0)                     # années complètes couvertes
            # scission / cession : le consensus ne compte plus l'activité sortie, les 12 derniers mois si -> le DCF
            # repart du CA et des marges du consensus de l'exercice N (pro forma)
            spin = (x0.get("disc_ops") and not f["cyclical"] and len(revs) >= 2 and revs[0] < 0.97 * x0["rev0"]
                    and revs[1] > revs[0])
            if spin:
                notes.append(f"⚠️ scission ou cession détectée : CA prévu N {revs[0]:,.0f} < 12 derniers mois "
                             f"{x0['rev0']:,.0f} M → DCF recalé sur le consensus (pro forma) ; dette nette du dernier "
                             f"bilan publié".replace(",", " "))
                f["spin_short"] = "scission/cession détectée : DCF sur la base pro forma du consensus"
                fc["g"] = [revs[i + 1] / revs[i] - 1 for i in range(len(revs) - 1)]
                fc["m"] = ms[2:2 + len(fc["g"])]
                fc["capex"] = (fc.get("capex") or [])[1:1 + len(fc["g"])]
                fc["fcf_m"] = [None] * len(fc["g"])
                f["dcfx"] = dict(x0, fc=fc, rev0=revs[0], m0=ms[1], src=x0["src"] + " + consensus Zonebourse (pro forma)")
            else:
                prev, g, m_ = x0["rev0"], [], []
                # marges : l'EBITDA « ajusté » du consensus peut différer de celui des états Finviz -> on applique
                # l'évolution prévue (rapport à la dernière marge publiée Zonebourse) à la marge Finviz actuelle
                m_q = at(ms, q, log=False)                              # marge Zonebourse à la date des 12 derniers mois
                scale = x0["m0"] / m_q if m_q and m_q > 0 and x0["m0"] > 0 else 1.0
                for k in range(1, n + 1):
                    rv = at(pts, q + k)
                    g.append(rv / prev - 1)
                    prev = rv
                    m_.append(at(ms, q + k, log=False) * scale)
                fc["g"], fc["m"] = g, m_
                fc["capex"] = (fc.get("capex") or [])[:n]
                fc["fcf_m"] = [None] * n                                # flux : marge de trésorerie Finviz × marges
        if "pro forma" not in f["dcfx"].get("src", ""):
            f["dcfx"] = dict(fv["dcfx"], fc=fc, src=fv["dcfx"]["src"] + " + consensus Zonebourse")
    elif not fv.get("dcfx") and zb.get("dcfx"):
        f["dcfx"] = zb["dcfx"]
    if zb.get("target") and fv.get("target"):
        notes.append(f"objectif moyen : Finviz {fv['target']:.0f} · Zonebourse {zb['target']:.0f} "
                     f"({zb.get('n_analysts') or '?'} analystes, « {zb.get('consensus_label') or '?'} »)")
    f["merge_notes"] = notes
    f["isin"] = fv.get("isin") or zb.get("isin") or ""
    f["source"] = (fv.get("source") or "Finviz") + " + " + (zb.get("source") or "Zonebourse")
    return f


TOP_ROLES = re.compile(r"\b(CEO|Chief Executive|CFO|Chief Financial|President|Chairman|COB|Founder)\b", re.I)


def insider_stats(rows, ref=None):
    """Analyse des transactions des dirigeants (achats « P » / ventes « S » en bourse, dédoublonnées) :
    montants, nombre d'acheteurs, achats groupés (≥ 2 dirigeants en 30 j), achats du PDG / directeur financier,
    part de la détention achetée, ventes accompagnées d'une levée d'options ; signal de synthèse."""
    ref = ref or TODAY
    seen, tx = set(), []
    exer = []                                    # levées d'options (fiche Finviz) : la vente associée est repérée
    for r in rows:
        if len(r) >= 4 and r[3] == "Option Exercise":
            try:
                exer.append((r[0].strip().lower(), datetime.strptime(r[2].replace("'", "20"), "%b %d %Y")))
            except ValueError:
                pass
    for r in rows:
        if len(r) < 7:
            continue
        name, rel, date, typ, price, sh, val = r[:7]
        if typ not in ("Buy", "Sale"):          # on ignore « Proposed Sale » (doublon) et « Option Exercise »
            continue
        key = (date, typ, sh)
        if key in seen:                          # déclarants conjoints
            continue
        seen.add(key)
        try:
            d = datetime.strptime(date.replace("'", "20"), "%b %d %Y")
        except ValueError:
            continue
        shares, total = num(sh) or 0, num(r[7]) if len(r) > 7 else None
        pct = None
        if typ == "Buy" and total and total > shares:
            pct = shares / (total - shares) * 100          # hausse de la détention du dirigeant
        pct_sold = shares / (total + shares) * 100 if typ == "Sale" and total is not None and total + shares > 0 \
            else None                                # part de la détention vendue (détention après + titres vendus)
        tx.append({"date": d, "age": (ref - d).days, "name": name.strip(), "rel": rel, "typ": typ, "pct_sold": pct_sold,
                   "val": (num(val) or 0) / 1e6, "pct": pct,
                   "oe": "levée d'options" in rel or (typ == "Sale" and any(
                       n == name.strip().lower() and abs((d - e).days) <= 3 for n, e in exer)),
                   "top": bool(TOP_ROLES.search(rel)), "big": "10%" in rel})
    y = [t for t in tx if t["age"] <= 365]
    buys, sales = [t for t in y if t["typ"] == "Buy"], [t for t in y if t["typ"] == "Sale"]
    out = {"Achats 90j": round(sum(t["val"] for t in buys if t["age"] <= 90), 1),
           "Ventes 90j": round(sum(t["val"] for t in sales if t["age"] <= 90), 1),
           "Achats 365j": round(sum(t["val"] for t in buys), 1),
           "Ventes 365j": round(sum(t["val"] for t in sales), 1)}
    out["Net 365j"] = round(out["Achats 365j"] - out["Ventes 365j"], 1)
    buyers = sorted({t["name"] for t in buys})
    cluster = 0                                       # plus grand nombre d'acheteurs distincts sur 30 jours
    for t in buys:
        cluster = max(cluster, len({u["name"] for u in buys if 0 <= (u["date"] - t["date"]).days <= 30}))
    top_buy = sum(t["val"] for t in buys if t["top"])
    pcts = [t["pct"] for t in buys if t["pct"] is not None]
    oe = sum(t["val"] for t in sales if t["oe"])
    out.update({"Acheteurs": len(buyers), "Achats groupés (max en 30 j)": cluster,
                "Achats PDG / DAF (M$)": round(top_buy, 1),
                "Hausse médiane de détention (%)": round(sorted(pcts)[len(pcts) // 2]) if pcts else None,
                "Ventes avec levée d'options (M$)": round(oe, 1)})
    # signal
    sold = out["Ventes 365j"] - oe                    # ventes hors levées d'options
    if buys and (cluster >= 2 or top_buy >= 0.25):
        sig = "🟢🟢 fort"
    elif sold >= 10 and sold > 5 * out["Achats 365j"]:
        sig = "🟠 ventes"
    elif out["Achats 365j"] >= 0.1:
        sig = "🟢 achats"
    else:
        sig = "⚪ peu de mouvements"
    parts = []
    if buys:
        lab = f"{len(buyers)} acheteur{'s' if len(buyers) > 1 else ''}"
        if cluster >= 2:
            lab = f"achats groupés ({cluster} en 30 j)"
        if top_buy >= 0.1:
            lab += ", dont PDG/DAF"
        parts.append(f"{lab} +{out['Achats 365j']:.1f} M$")
    if sales:
        parts.append(f"ventes {out['Ventes 365j']:.1f} M$" + (f" (dont {oe:.1f} avec levée d'options)" if oe >= 0.05 else ""))
    out["Signal"] = sig
    out["Résumé"] = f"{sig.split()[0]} " + (" · ".join(parts) if parts else "aucun achat ni vente en bourse")
    out["_tx"] = [{k: t[k] for k in ("date", "name", "rel", "typ", "val", "oe", "pct_sold")} for t in y]
    out["Achats notables"] = [f"{t['name']} ({t['rel']}) {t['date']:%d/%m/%y} : {t['val']:.1f} M$"
                              + (f", +{t['pct']:.0f} % de sa détention" if t["pct"] is not None else "")
                              for t in sorted(buys, key=lambda t: -t["val"])[:5]]
    return out


# ============================================================
#  JUSTE VALEUR
# ============================================================
def dcf_value(fcf0, g0, tg, w, years=10, terminal=True):
    tot, f = 0.0, fcf0
    for t in range(1, years + 1):
        f *= 1 + g0 + (tg - g0) * (t - 1) / (years - 1)
        tot += f / (1 + w) ** (t - 0.5)
    return tot + (f * (1 + tg) / (w - tg) / (1 + w) ** (years - 0.5) if terminal else 0.0)


def synthese(out, methods):
    """JV centrale = moyenne (2 méthodes) ou médiane (3 méthodes : une méthode aberrante ne tire plus tout) ;
    fourchette = plus bas / plus haut des méthodes retenues, écartant la plus éloignée de la médiane quand il y en a 3."""
    ms = sorted((out[m], m) for m in methods if out.get(m))
    if not ms:
        return
    if len(ms) >= 3:
        vals = [v for v, _ in ms]
        n = len(vals)
        med = vals[n // 2] if n % 2 else (vals[n // 2 - 1] + vals[n // 2]) / 2
        out["JV centrale"] = med
        keep = [m for v, m in ms if abs(v / med - 1) <= 0.5]
        for v, m in ms:
            if m not in keep:
                out["notes"].append(f"méthode {m} écartée de la fourchette ({v:.0f}, trop éloignée)")
        keep = keep or [m for _, m in ms]
    else:
        out["JV centrale"] = sum(v for v, _ in ms) / len(ms)
        keep = [m for _, m in ms]
    out["_keep"], out["_excl"] = keep, [m for _, m in ms if m not in keep]
    out["JV basse"] = min(out[m + " bas"] for m in keep)
    out["JV haute"] = max(out[m + " haut"] for m in keep)


def add_hist_pe(f, out):
    """3e méthode si l'onglet Valorisation Zonebourse est chargé : BNA N+1 × PER médian de la société."""
    eps, pe = f.get("eps_next_y_zb") or f["eps_next_y"], f.get("pe_hist")   # même base que les PER Zonebourse
    # PER historique mesuré sur des bénéfices bien plus petits (BNA N+1 ≥ 2,5 × le BNA du dernier exercice) et deux
    # fois plus élevé que le PER cible : il reflète une croissance passée exceptionnelle, pas un multiple applicable
    # aux bénéfices futurs (ex. Nvidia : 51,6x historique sur un BNA multiplié par 3) -> non retenu
    eL = f.get("eps_ttm")
    pe_c = out["PER"] / f["eps_next_y"] if out.get("PER") and f.get("eps_next_y") else None
    if (not f.get("cyclical") and eps and pe and pe_c and eL and eL > 0 and eps >= 2.5 * eL
            and pe > 2 * pe_c):
        out["notes"].append(f"PER historique {pe:.1f}x non retenu : mesuré sur des bénéfices {eps / eL:.1f} fois plus "
                            f"petits que le BNA N+1 (croissance passée exceptionnelle) et plus de 2 × le PER cible "
                            f"({pe_c:.1f}x)")
        out.setdefault("_alerts", []).append(f"ℹ️ PER historique ({pe:.0f}x) non retenu : bénéfices passés bien plus "
                                             f"petits")
        return
    if eps and eps > 0 and pe:
        out.update({"PER hist. bas": eps * pe * 0.85, "PER hist.": eps * pe, "PER hist. haut": eps * pe * 1.15})
        out.setdefault("_mult", {})["PER hist."] = f"PER historique {pe:.1f}x × BNA {eps:.2f}"
        out["notes"].append(f"PER historique médian {pe:.1f}x" + (f" (P/B historique médian {f['pb_hist']:.2f})"
                                                                 if f.get("pb_hist") else ""))


def fair_value_bank(f, p):
    """Banque : actif net justifié par la rentabilité (P/B = (ROE − g) / (coût des fonds propres − g))
    et PER bancaire, moyenne des deux."""
    out = {"notes": []}
    g, k = p["tg"], p["coe_bank"]
    z = f.get("zb")
    if z:                                            # pages Zonebourse : actif net et ROE prévus
        N = z["N"]
        bv = z["bvps"].get(N) or z["bvps"].get(z["L"])
        roe = z["roe"].get(N + 1) or z["roe"].get(N)
    else:                                            # fiche Finviz : actif net par action, ROE prévu = BNA N+1 / actif net
        bv = f.get("bvps")
        roe = (f["eps_next_y"] / bv * 100 if bv and f.get("eps_next_y") else f.get("roe_ttm"))
        if f.get("beta") and p.get("capm"):
            k = max(cost_of_equity(f, p)[0], 0.08)
    if f.get("currency", "USD") == "USD" and not z:
        pe_b = p.get("pe_bank_us", 12.0)             # banques US : PER historique ~12x (Europe ~8-9x)
    else:
        pe_b = p["pe_bank"]
    ins = bool(f.get("insurer"))
    if ins:
        # assureur : coût des fonds propres de la valeur (CAPM ou taux unique, ≥ 8 %) plutôt que celui des banques ;
        # PER cible selon la croissance, plafonné à 1,25 × le PER historique, plutôt que le PER bancaire
        k = max(cost_of_equity(f, p)[0], 0.08)
        g5, _ = growth_expected(f)
        pe_b = clamp(p["pe_base"] + 0.4 * (min(g5, 25) - 10), p["pe_min"], p["pe_max"])
        if f.get("pe_hist") and pe_b > 1.25 * f["pe_hist"]:
            pe_b = 1.25 * f["pe_hist"]
        out["notes"].append("assureur : modèle financier (actif net × P/B justifié, PER) — pas de DCF d'entreprise, "
                            "les flux et la dette d'un assureur dépendent de son bilan (placements, réserves)")
    if bv and roe:
        r = roe / 100
        v = lambda kk: bv * max(r - g, 0) / (kk - g)
        out.update({"P/B bas": v(k + 0.01), "P/B": v(k), "P/B haut": v(k - 0.01)})
        out.setdefault("_mult", {})["P/B"] = f"P/B justifié {v(k) / bv:.2f} × actif net {bv:.1f}"
        out["notes"].append(f"Actif net {bv:.1f} × P/B justifié {v(k) / bv:.2f} (ROE {roe:.1f} %, coût des fonds propres "
                            f"{k * 100:.1f} %, g {g * 100:.1f} %)")
    eps = f["eps_next_y"]
    if eps and eps > 0:
        pe = pe_b
        out.update({"PER bas": eps * pe * 0.85, "PER": eps * pe, "PER haut": eps * pe * 1.15})
        lbl_ = "PER cible" if ins else "PER bancaire"
        out.setdefault("_mult", {})["PER"] = f"{lbl_} {pe:.1f}x × BNA N+1 {eps:.2f}"
        out["notes"].append(f"{lbl_} {pe:.1f}x sur BNA N+1 {eps:.2f}")
    add_hist_pe(f, out)
    synthese(out, ("P/B", "PER", "PER hist."))
    if out.get("JV centrale") and f["price"]:
        out["Écart (%)"] = (out["JV centrale"] / f["price"] - 1) * 100
    precision(out, ("P/B", "PER", "PER hist."))
    return out


def cost_of_equity(f, p):
    """Taux d'actualisation propre à la valeur (CAPM) : taux sans risque + bêta × prime de risque.
    Sans bêta (pages Zonebourse) ou si l'option est décochée : taux unique de la barre latérale."""
    b = f.get("beta")
    if p.get("capm") and b:
        bb = clamp(0.67 * b + 0.33, 0.6, 2.0)            # bêta ajusté (Blume) : tend vers 1 à long terme
        return p["rf"] + bb * p["erp"], (f"CAPM {p['rf'] * 100:.1f} % + bêta ajusté {bb:.2f} (brut {b:.2f})"
                                         + f" × {p['erp'] * 100:.1f} %")
    if f.get("cyclical") and p.get("cyc_prem"):     # pas de bêta : prime pour les valeurs cycliques (auto, semi-conducteurs…)
        return p["wacc"] + p["cyc_prem"], f"taux unique {p['wacc'] * 100:.1f} % + prime cyclique {p['cyc_prem'] * 100:.1f} %"
    return p["wacc"], "taux unique" + (" (pas de bêta)" if p.get("capm") else "")


def wacc_info(f, ke, p):
    """WACC indicatif (non utilisé dans le calcul) : poids de la dette nette, coût de la dette après impôt."""
    d, e = max(f.get("net_debt") or 0.0, 0.0), f.get("mcap") or 0.0
    if not e:
        return None
    kd = (p["rf"] + 0.015) * (1 - 0.25)          # taux sans risque + 1,5 pt de marge, impôt 25 %
    return (e * ke + d * kd) / (e + d)


# ============================================================
#  DCF D'ENTREPRISE (flux de trésorerie disponibles pour l'entreprise, FCFF)
#  scénarios pessimiste / central / optimiste, valeur terminale par croissance perpétuelle ET par multiple
#  de sortie, CMPC (WACC), valeur d'entreprise -> valeur des fonds propres, Monte Carlo, sensibilité
# ============================================================
DCF_YEARS = 10
SCEN = (("Pessimiste", 0.25, -0.04, -0.02), ("Central", 0.50, 0.0, 0.0), ("Optimiste", 0.25, 0.04, 0.02))


def _med(xs):
    xs = sorted(x for x in xs if x is not None)
    if not xs:
        return None
    n = len(xs)
    return xs[n // 2] if n % 2 else (xs[n // 2 - 1] + xs[n // 2]) / 2


def cash_unrepresentative(f, IA, CA, g=None):
    """Société rentable dont les flux de trésorerie publiés ne reflètent pas l'activité (ex. Constellation : créances
    cédées encaissées en « investissement », appels de marge) : flux d'exploitation des 3 derniers exercices
    < 40 % de (résultat + amortissements) ; fiche seule : flux libre des 12 derniers mois < 30 % du résultat."""
    if f.get("bank"):
        return False
    if CA and IA:
        cfo = [stmt(CA, "Cash from Operating Activities", i) for i in (1, 2, 3)]
        ni = [stmt(IA, "Net Income", i) for i in (1, 2, 3)]
        da = [stmt(CA, "Depreciation", i) or 0.0 for i in (1, 2, 3)]
        if None in cfo or None in ni or sum(ni) <= 0:
            return False
        return sum(cfo) < 0.4 * (sum(ni) + sum(da))
    earn = (f.get("eps_ttm") or 0) * (f.get("shares") or 0)
    fcf0 = (f.get("fcf_hist") or [None])[0]
    return bool(earn > 0 and fcf0 is not None and fcf0 < 0.3 * earn)


def capital_hist_finviz(IA, CA):
    """{exercice: {shares, sbc, buyback}} (millions) des états Finviz, exercices clos."""
    out = {}
    for d, keys in ((IA, (("shares", "Shares Outstanding"), ("sbc", "Stock Option Compensation Expense"))),
                    (CA, (("buyback", "Repurchase of Common Pref Stock"),))):
        if not d or not d.get("Period"):
            continue
        for i, pp in enumerate(d["Period"]):
            m = re.match(r"^(\d{4})FY$", str(pp))
            if not m:
                continue
            for k, lab in keys:
                v = stmt(d, lab, i)
                if v is not None:
                    out.setdefault(int(m.group(1)), {})[k] = abs(v) if k == "buyback" else v
    return out


def dcfx_from_finviz(IA, BA, CA, sn):
    """Hypothèses du DCF d'entreprise tirées des états financiers Finviz (millions ; indice 0 = 12 derniers mois)."""
    if not IA or not CA:
        return None
    s = lambda d, k, i=0: stmt(d, k, i)
    rev = [s(IA, "Total Revenue", i) for i in range(6)]
    if not rev[0] or rev[0] <= 0:
        return None
    eb = [s(IA, "EBITDA", i) for i in range(6)]
    da = [s(CA, "Depreciation", i) for i in range(6)]
    cap = [abs(s(CA, "Capital Expenditures", i)) if s(CA, "Capital Expenditures", i) is not None else None
           for i in range(6)]
    wc = [s(CA, "Changes in Working Capital", i) for i in range(6)]
    if eb[0] is None:
        return None
    pct = lambda a, b: a / b if a is not None and b else None
    m_hist = [pct(eb[i], rev[i]) for i in range(1, 6)]               # exercices complets
    cap_hist = [pct(cap[i], rev[i]) for i in range(1, 6)]
    # impôt effectif sur 3 exercices, borné (crédits d'impôt / reports déficitaires exclus)
    tx = [s(IA, "Income Taxes", i) for i in range(1, 4)]
    pt = [s(IA, "Net Income Before Taxes", i) for i in range(1, 4)]
    tsum, psum = sum(x for x in tx if x is not None), sum(x for x in pt if x is not None)
    tax = tsum / psum if psum > 0 else None
    tax = clamp(tax, 0.12, 0.30) if tax is not None and tax > 0.05 else 0.21
    # besoin en fonds de roulement : part de la hausse du CA absorbée (médiane historique, bornée)
    k = [-(wc[i]) / (rev[i] - rev[i + 1]) for i in range(0, 5)
         if wc[i] is not None and rev[i] and rev[i + 1] and abs(rev[i] - rev[i + 1]) > 0.02 * rev[i + 1]]
    nwc_k = clamp(_med(k), 0.0, 0.25) if k else 0.05
    sbc = s(IA, "Stock Option Compensation Expense") or s(IA, "Stock Option Compensation Expense", 1) or 0.0
    # marge de trésorerie d'exploitation avant investissements, après impôt, avant intérêts, rémunération en
    # actions déduite : (flux d'exploitation + intérêts × (1 − impôt) − rémunération en actions) / CA
    NH = 9                                       # 12 derniers mois + jusqu'à 8 exercices (cycle complet)
    rv = [s(IA, "Total Revenue", i) for i in range(NH)]
    cfo = [s(CA, "Cash from Operating Activities", i) for i in range(NH)]
    itr = [s(IA, "Interest Expense", i) or 0.0 for i in range(NH)]
    sb = [s(IA, "Stock Option Compensation Expense", i) or (sbc if i == 0 else 0.0) for i in range(NH)]
    ocf_m = [(cfo[i] + itr[i] * (1 - tax) - sb[i]) / rv[i] for i in range(NH) if cfo[i] is not None and rv[i]]
    debt = (s(BA, "Short Term Debt Incl. Current Port. of LT Debt") or 0) + (s(BA, "Long Term Debt") or 0)
    cash = s(BA, "Cash & Short Term Investments") or 0
    return {"src": "états financiers Finviz (12 derniers mois)", "rev0": rev[0], "m0": eb[0] / rev[0],
            "m_hist": [x for x in m_hist if x is not None], "da_pct": pct(da[0], rev[0]) or 0.0,
            "capex0": pct(cap[0], rev[0]) or 0.0, "capex_norm": _med(cap_hist) or pct(cap[0], rev[0]) or 0.0,
            "tax": tax, "nwc_k": nwc_k, "sbc_pct": (sbc or 0) / rev[0], "debt": debt, "cash": cash,
            "net_debt": debt - cash, "minority": s(BA, "Accumulated Minority Interest") or 0.0,
            "interest": s(IA, "Interest Expense") or 0.0, "ebit": s(IA, "Operating Income"),
            "ev_ebitda": num(sn.get("EV/EBITDA")) if sn else None, "fc": None,
            "ocf0": ocf_m[0] if ocf_m else None, "ocf_hist": ocf_m[1:],
            "intang": pct(s(BA, "Intangible Assets"), s(BA, "Total Assets")) or 0.0,
            "disc_ops": disc_ops(IA)}


def disc_ops(IA):
    """Activités abandonnées (scission, cession) : bénéfice net ≠ bénéfice des activités poursuivies (12 derniers mois
    ou dernier exercice)."""
    for i in (0, 1):
        a, b = stmt(IA, "Net Income", i), stmt(IA, "Net Income From Continuing Operations", i)
        if a is not None and b and abs(a - b) > 0.005 * abs(b):
            return True
    return False


def asset_light(x):
    """Société peu consommatrice d'investissements dont les amortissements sont surtout ceux d'incorporels rachetés
    (incorporels ≥ 40 % de l'actif, investissements < 3 % du CA et < la moitié des amortissements)."""
    cap = max(x.get("capex0") or 0.0, x.get("capex_norm") or 0.0)
    return (x.get("intang") or 0) >= 0.4 and cap < 0.03 and (x.get("da_pct") or 0) > 2 * cap


def dcfx_from_zonebourse(z, ca, eb, ebit, capex, nd, act, est, shares, fcf=None):
    """Hypothèses du DCF tirées des pages Zonebourse : historique + prévisions du consensus (CA, EBITDA, capex)."""
    L = act[-1] if act else None
    if L is None or not ca.get(L) or not eb.get(L):
        return None
    rev0 = ca[L]
    pct = lambda a, b: a / b if a is not None and b else None
    m_hist = [pct(eb.get(y), ca.get(y)) for y in act[-5:]]
    cap_hist = [pct(abs(capex[y]) if capex.get(y) is not None else None, ca.get(y)) for y in act[-5:]]
    da_l = [(eb[y] - ebit[y]) / ca[y] for y in act[-5:]
            if eb.get(y) is not None and ebit.get(y) is not None and ca.get(y) and eb[y] - ebit[y] > 0]
    fcf = fcf or {}
    fc = {"g": [], "m": [], "capex": [], "fcf_m": [], "rev": []}
    prev = rev0
    for y in est[:3]:
        if ca.get(y) and prev:
            fc["rev"].append(ca[y])
            fc["g"].append(ca[y] / prev - 1)
            fc["m"].append(pct(eb.get(y), ca[y]))
            fc["capex"].append(pct(abs(capex[y]) if capex.get(y) is not None else None, ca[y]))
            fc["fcf_m"].append(pct(fcf.get(y), ca[y]))
            prev = ca[y]
    # conversion de l'EBITDA en flux libre (consensus + 3 derniers exercices) : reprend ce que prévoient les analystes
    # (intérêts, loyers, R&D immobilisée… inclus), plus prudent qu'un FCFF reconstruit
    conv = [fcf[y] / eb[y] for y in list(est[:3]) + list(act[-3:])
            if fcf.get(y) is not None and eb.get(y) and eb[y] > 0]
    net_debt = nd.get(L) if nd.get(L) is not None else 0.0
    return {"src": "pages Zonebourse (historique + consensus)", "rev0": rev0, "m0": eb[L] / rev0,
            "m_hist": [x for x in m_hist if x is not None], "da_pct": _med(da_l) or 0.04,
            "capex0": pct(abs(capex[L]) if capex.get(L) is not None else None, rev0) or 0.04,
            "capex_norm": _med(cap_hist) or 0.04, "tax": 0.25, "nwc_k": 0.05, "sbc_pct": 0.0,
            "debt": max(net_debt, 0.0), "cash": max(-net_debt, 0.0), "net_debt": net_debt, "minority": 0.0,
            "interest": 0.0, "ebit": ebit.get(L), "ev_ebitda": None, "fc": fc if fc["g"] else None,
            "conv": clamp(_med(conv), -0.2, 0.9) if conv else None}


def wacc_calc(f, x, p):
    """CMPC : coût des fonds propres (CAPM) et coût de la dette après impôt pondérés aux valeurs de marché."""
    ke, how = cost_of_equity(f, p)
    ke = max(ke, p["tg"] + 0.03)
    e, d = f.get("mcap") or 0.0, max(x.get("debt") or 0.0, 0.0)
    cover = x["ebit"] / x["interest"] if x.get("ebit") and x.get("interest") else None
    if cover is not None:
        spread = 0.015 if cover > 6 else 0.025 if cover > 3 else 0.04
    else:                                         # pas de charge d'intérêts (pages Zonebourse) : dette nette / EBITDA
        eb_ = x["rev0"] * x["m0"]
        lev = (x.get("net_debt") or 0.0) / eb_ if eb_ > 0 else 9.0
        spread = 0.015 if lev <= 1 else 0.025 if lev <= 2.5 else 0.04 if lev <= 4 else 0.06
    # taux unique (pas de bêta de marché) : il vaut pour un endettement courant (dette / capitalisation ≈ 30 %) ;
    # au-delà, le risque des actionnaires augmente (bêta réendetté, Hamada) — sinon une dette lourde et bon marché
    # ferait baisser le CMPC (ex. Air France-KLM : dette = 3× la capitalisation)
    if "taux unique" in how and e > 0 and d / e > 0.3:
        de = min(d / e, 4.0)
        bu = max(ke - p["rf"], 0.0) / p["erp"] / (1 + 0.75 * 0.3)
        add = p["erp"] * bu * 0.75 * (de - 0.3)
        ke = min(ke + add, 0.20)
        how += f" + prime d'endettement {add * 100:.1f} pt (dette = {d / e:.1f}× la capitalisation)"
    kd = (p["rf"] + spread) * (1 - x["tax"])
    w = (e * ke + d * kd) / (e + d) if e + d > 0 else ke
    return {"ke": ke, "kd": kd, "spread": spread, "E": e, "D": d, "wacc": max(w, p["tg"] + 0.03), "how": how}


def dcfx_project(x, g0, dm, wacc, tg, exit_mult, cyclical, g_shift=0.0, years=DCF_YEARS, fc_shift=0.5):
    """Projection sur 10 ans -> valeur des fonds propres (croissance perpétuelle / multiple de sortie) et tableau.
    g0 : croissance du CA en année 1 (sans consensus) ; avec le consensus Zonebourse, ses croissances sont
    reprises pour les premières années, décalées de g_shift (scénario) ; dm : décalage de marge EBITDA."""
    fc = x.get("fc") or {}
    fg, fm, fcap = fc.get("g") or [], fc.get("m") or [], fc.get("capex") or []
    # valeur cyclique : marge « de milieu de cycle » = moyenne du cycle (creux et pic compris), pas la médiane, qui
    # retient une année de creux ou de pic quand l'historique est court
    m_norm = ((sum(x["m_hist"]) / len(x["m_hist"]) if cyclical else _med(x["m_hist"])) if x["m_hist"] else x["m0"])
    m_base = m_norm if cyclical else x["m0"]
    capex_term = max(x["da_pct"] * 1.05, min(x["capex0"], x["capex_norm"]))
    if asset_light(x):        # amortissements = incorporels rachetés, pas des investissements à renouveler
        capex_term = max(x["capex0"], x["capex_norm"]) * 1.25
    # marge de trésorerie d'exploitation (états Finviz) : 12 derniers mois et médiane des 3 derniers exercices,
    # médiane sur 5 ans pour une valeur cyclique ; sinon (Zonebourse) FCFF reconstruit depuis l'EBITDA
    ocf = None
    if x.get("ocf0") is not None:
        h = x.get("ocf_hist") or []
        ocf = (sum(h) / len(h) if cyclical and h else (x["ocf0"] + (_med(h[:3]) if h else x["ocf0"])) / 2)
        if x["m0"] and x["m0"] > 0:
            ocf_ratio = ocf / (m_base if m_base else x["m0"])     # part de l'EBITDA convertie en trésorerie
    ronic = max(wacc + 0.02, 0.10)                  # rentabilité des nouveaux investissements (prudente)
    rev, pv, rows = x["rev0"], 0.0, []
    nf = len(fg)
    # après le consensus : la décroissance vers la croissance perpétuelle part au plus de 25 % (même borne que la
    # croissance de l'année 1 sans consensus) — une hyper-croissance prévue sur 3 ans n'est pas prolongée telle quelle
    g_start = min(fg[-1], 0.25) if fg else g0
    for t in range(1, years + 1):
        if t <= nf:                                           # années couvertes par le consensus
            g = fg[t - 1] + g_shift * fc_shift
            m = (fm[t - 1] if fm[t - 1] is not None else m_base) + dm
            cpx = fcap[t - 1] if t - 1 < len(fcap) and fcap[t - 1] is not None else x["capex0"]
        else:
            base = g0 if not fg else g_start + g_shift
            if cyclical and fg:                               # après le consensus, pas de prolongation du rebond
                base = min(base, 0.05 + g_shift)
            g = base + (tg - base) * (t - nf - 1) / max(years - nf - 1, 1)
            m_last = (fm[-1] if fm and fm[-1] is not None else m_base)
            if cyclical:                                      # retour vers la marge médiane du cycle en 4 ans
                tgt = min(m_norm, m_last) if fm else m_norm     # jamais au-dessus du consensus
                m_last = m_last + (tgt - m_last) * min(1.0, (t - nf) / 4)
            m = m_last + dm
            c_last = fcap[-1] if fcap and fcap[-1] is not None else x["capex0"]
            cpx = c_last + (capex_term - c_last) * (t - nf) / (years - nf)
        da = x["da_pct"] + (capex_term / 1.05 - x["da_pct"]) * t / years    # amortissements -> niveau d'entretien
        new = rev * (1 + g)
        ebitda = new * max(m, 0.0)
        ebit = ebitda - new * da
        ffm = (fc.get("fcf_m") or [])
        if ocf is not None:
            # marges prévues par le consensus : la trésorerie d'exploitation suit l'EBITDA (conversion constante)
            ocf_t = ocf * (m - dm) / m_base if fm and m_base and m_base > 0 else ocf
            fcff = new * (ocf_t + dm * (ocf_ratio if x["m0"] and x["m0"] > 0 else 1.0)) - new * cpx
        elif t <= nf and t - 1 < len(ffm) and ffm[t - 1] is not None:      # flux libre du consensus
            fcff = new * (ffm[t - 1] + dm * (x.get("conv") or 0.5))
        elif ffm and ffm[-1] is not None:
            # après le consensus : transition sur 3 ans du dernier flux prévu vers le flux « économique »
            # NOPAT × (1 − g / RONIC) (résultat d'exploitation après impôt, moins l'investissement nécessaire à la
            # croissance) — intègre loyers et amortissements, que l'EBITDA × conversion surestime
            drv = ebit * (1 - x["tax"]) * (1 - max(g, 0.0) / ronic)
            a_ = min(1.0, (t - nf) / 3)
            fcff = (1 - a_) * new * ffm[-1] + a_ * drv
        elif x.get("conv") is not None:                                    # EBITDA × conversion observée / prévue
            fcff = ebitda * x["conv"]
        else:
            fcff = ebit * (1 - x["tax"]) + new * da - new * cpx - x["nwc_k"] * (new - rev) - new * x["sbc_pct"]
        pv += fcff / (1 + wacc) ** (t - 0.5)
        rows.append({"Année": t, "Croissance CA (%)": g * 100, "CA": new, "Marge EBITDA (%)": m * 100,
                     "EBITDA": ebitda, "Capex (% CA)": cpx * 100, "FCFF": fcff, "_nopat": ebit * (1 - x["tax"])})
        rev = new
    last = rows[-1]
    # valeur terminale par croissance perpétuelle : formule des « value drivers » (NOPAT × (1 − g / RONIC)),
    # sans dépasser celle du dernier flux projeté
    tv_drv = last["_nopat"] * (1 + tg) * (1 - tg / ronic) / (wacc - tg)
    tv_fcf = last["FCFF"] * (1 + tg) / (wacc - tg)
    # (flux calés sur la trésorerie réelle des états Finviz : on garde le dernier flux projeté)
    tv_perp = tv_fcf if ocf is not None else (min(tv_drv, tv_fcf) if tv_drv > 0 else tv_fcf)
    # multiple de sortie cohérent avec la trésorerie générée : au plus 1,5 × le multiple implicite de la valeur
    # perpétuelle (une société qui convertit peu d'EBITDA en trésorerie ne vaut pas un multiple normal)
    em = exit_mult
    if last["EBITDA"] > 0 and tv_perp > 0:
        em = min(em, max(2.0, 1.5 * tv_perp / last["EBITDA"]))
    tv_exit = last["EBITDA"] * em
    disc = (1 + wacc) ** (years - 0.5)
    bridge = x["net_debt"] + x.get("minority", 0.0)
    ev_p, ev_x = pv + tv_perp / disc, pv + tv_exit / disc
    return {"eq_perp": ev_p - bridge, "eq_exit": ev_x - bridge, "ev_perp": ev_p, "ev_exit": ev_x, "em": em,
            "pv_fcff": pv, "tv_share": (tv_perp / disc) / ev_p if ev_p > 0 else None, "rows": rows}


def dcf_entreprise(f, p, x, g_rev, mc_n=2000):
    """DCF complet : 3 scénarios pondérés (25/50/25), deux valeurs terminales, Monte Carlo, valeur par action."""
    sh = f.get("shares")
    if not sh or not x:
        return None
    wc = wacc_calc(f, x, p)
    w, tg = wc["wacc"], p["tg"]
    cyc = bool(f.get("cyclical"))
    evm = x.get("ev_ebitda")
    eb0 = x["rev0"] * x["m0"]
    if not evm and f.get("mcap") and eb0 > 0:                 # VE / EBITDA actuel du marché (pages Zonebourse)
        evm = (f["mcap"] + x["net_debt"]) / eb0
    exit_mult = clamp(evm or 8.0, 4.0, 12.0)
    if cyc:
        exit_mult = min(exit_mult, 9.0)   # EBITDA déjà normalisé sur le cycle : 9x (et non 7x, double pénalisation)
    scen, central = [], 0.0
    for name, wgt, dg, dm in SCEN:
        dm_ = dm
        if cyc and name == "Pessimiste" and x["m_hist"]:
            dm_ = min(x["m_hist"]) - (_med(x["m_hist"]) or x["m0"])          # creux de cycle
        r = dcfx_project(x, g_rev + dg, dm_, w, tg, exit_mult, cyc, g_shift=dg)
        v_p, v_x = r["eq_perp"] / sh, r["eq_exit"] / sh
        v = (v_p + v_x) / 2
        scen.append({"Scénario": name, "Poids": wgt, "Croissance CA an 1 (%)": (g_rev + dg) * 100,
                     "Écart de marge (pts)": dm_ * 100, "Valeur / action (croiss. perpétuelle)": v_p,
                     "Valeur / action (multiple de sortie)": v_x, "Valeur / action": v,
                     "Valeur terminale (% VE)": (r["tv_share"] or 0) * 100, "_rows": r["rows"]})
        central += wgt * v
    base = scen[1]
    # Monte Carlo : croissance, marge, CMPC, croissance perpétuelle et multiple de sortie incertains
    rng = np.random.default_rng(42)
    vals = []
    for _ in range(mc_n):
        dg_ = rng.normal(0, 0.03)
        gg = g_rev + dg_
        mm = rng.normal(0, 0.015)
        ww = max(w + rng.normal(0, 0.0075), tg + 0.02)
        tt = rng.uniform(0.015, 0.035)
        tt = min(tt, ww - 0.015)
        em = exit_mult * rng.uniform(0.8, 1.2)
        r = dcfx_project(x, gg, mm, ww, tt, em, cyc, g_shift=dg_)
        vals.append((r["eq_perp"] + r["eq_exit"]) / 2 / sh)
    vals = np.array(vals)
    pr = f.get("price")
    return {"central": central, "p10": float(np.percentile(vals, 10)), "p50": float(np.percentile(vals, 50)),
            "p90": float(np.percentile(vals, 90)),
            "prob_up": float((vals > pr).mean() * 100) if pr else None,
            "scen": scen, "wacc": wc, "tg": tg, "exit_mult": exit_mult, "g_rev": g_rev,
            "gap_tv": abs(base["Valeur / action (croiss. perpétuelle)"] / base["Valeur / action (multiple de sortie)"] - 1)
            if base["Valeur / action (multiple de sortie)"] > 0 else None,
            "src": x["src"], "x": x}


def dcfx_implied_growth(f, p, x, res):
    """DCF inversé sur le modèle d'entreprise : croissance du CA en année 1 que suppose le cours."""
    pr, sh = f.get("price"), f.get("shares")
    if not (pr and sh and res):
        return None
    w, tg, em, cyc = res["wacc"]["wacc"], res["tg"], res["exit_mult"], bool(f.get("cyclical"))
    g_ref = res["g_rev"]
    # même écart de croissance appliqué à toutes les années (consensus compris)
    proj = lambda d: dcfx_project(x, g_ref + d, 0.0, w, tg, em, cyc, g_shift=d, fc_shift=1.0)
    v = lambda d: (lambda r: (r["eq_perp"] + r["eq_exit"]) / 2 / sh)(proj(d))
    avg5 = lambda r: sum(row["Croissance CA (%)"] for row in r["rows"][:5]) / 5
    lo, hi = -0.40, 0.60
    if v(lo) > pr:
        d = lo
    elif v(hi) < pr:
        d = hi
    else:
        for _ in range(40):
            mid = (lo + hi) / 2
            lo, hi = (mid, hi) if v(mid) < pr else (lo, mid)
        d = (lo + hi) / 2
    # croissance annuelle moyenne du CA sur 5 ans : supposée par le cours / scénario central
    return avg5(proj(d)) / 100, avg5(proj(0.0)) / 100


def revenue_growth(f):
    """Croissance du CA retenue pour l'année 1 (hors consensus Zonebourse) : moitié croissance du CA constatée
    (médiane 3 ans / 5 ans / 12 mois), moitié croissance du BNA plafonnée au CA + 5 points ; bornée −5 / +25 %."""
    gs = f.get("sales_g")
    if gs is None:
        gs = f.get("sales_cagr")
    g5, _ = growth_expected(f)
    if gs is None:
        return clamp(g5 / 100 * 0.6, -0.05, 0.25)
    return clamp((0.5 * gs + 0.5 * min(g5, gs + 5)) / 100, -0.05, 0.25)


_IND = {"data": None, "t": 0.0}


def parse_industry_groups(html):
    """Page Finviz « Groups / Industry / Valuation » -> {industrie: {"fpe": PER prévisionnel, "pfcf": P/FCF}}."""
    soup = BeautifulSoup(html, "html.parser")
    tb = soup.select_one("table.groups_table") or next(
        (t for t in soup.find_all("table") if "Fwd P/E" in t.get_text()), None)
    if tb is None:
        return {}
    trs = tb.find_all("tr")
    hdr = [c.get_text(" ", strip=True) for c in trs[0].find_all(["th", "td"])]
    out = {}
    for tr in trs[1:]:
        c = dict(zip(hdr, [x.get_text(" ", strip=True) for x in tr.find_all("td")]))
        if c.get("Name"):
            out[c["Name"]] = {"fpe": num(c.get("Fwd P/E")), "pfcf": num(c.get("P/FCF"))}
    return out


def industry_multiples(use_cache=True):
    """Multiples moyens des ~145 industries Finviz (une page), cache disque 24 h ; {} si indisponible."""
    if _IND["data"] is not None and time.time() - _IND["t"] < (24 * 3600 if _IND["data"] else 600):
        return _IND["data"]
    path = os.path.join(CACHE_DIR, "_industries.json")
    try:
        if use_cache and os.path.exists(path) and time.time() - os.path.getmtime(path) < 24 * 3600:
            with open(path, encoding="utf-8") as fh:
                data = json.load(fh)
        else:
            data = parse_industry_groups(_get(_session(), "https://finviz.com/groups.ashx?g=industry&v=120&o=name",
                                              pause=(0.3, 0.6)))
            if data:
                os.makedirs(CACHE_DIR, exist_ok=True)
                with open(path, "w", encoding="utf-8") as fh:
                    json.dump(data, fh)
    except Exception:
        data = {}
    _IND.update(data=data, t=time.time())
    return data


def growth_expected(f):
    """Croissance du BNA retenue (%) : BNA à 5 ans du consensus, plafonnée à la croissance du CA + 10 points
    (une croissance du BNA calculée depuis une base déprimée ne dure pas)."""
    g5 = f["eps_5y"] if f.get("eps_5y") is not None else 8.0
    gs = f.get("sales_g")
    if gs is None and f.get("sales_cagr") is not None:
        gs = f["sales_cagr"]
    return (min(g5, gs + 10), True) if gs is not None and g5 > gs + 10 else (g5, False)


def implied_growth(fcf, price, shares, tg, ke):
    """DCF inversé : croissance annuelle des 10 prochaines années (dégressive vers g terminal) que suppose le cours."""
    if not (fcf and fcf > 0 and price and shares) or ke <= tg:
        return None
    v = lambda g: dcf_value(fcf, g, tg, ke) / shares
    lo, hi = -0.30, 0.80
    if v(lo) > price:
        return lo
    if v(hi) < price:
        return hi
    for _ in range(60):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if v(mid) < price else (lo, mid)
    return (lo + hi) / 2


def _quart(xs, q):
    """Quantile (interpolation linéaire) d'une liste non vide."""
    xs = sorted(xs)
    i = (len(xs) - 1) * q
    lo, hi = int(i), min(int(i) + 1, len(xs) - 1)
    return xs[lo] + (xs[hi] - xs[lo]) * (i - lo)


def comparables(f, p, eps_mid=None):
    """Méthode des comparables : PER, VE/EBITDA (et P/FCF pour les pairs Finviz) — médiane des pairs pour la valeur
    centrale, 1er et 3e quartiles pour la fourchette ; position de la société par rapport à ses pairs.
    Pairs Finviz : PER prévisionnel × BNA N+1, VE/EBITDA (12 mois) × EBITDA 12 mois.
    Pairs Zonebourse : PER (N) × BNA N, VE/EBITDA (N) × EBITDA N (même exercice que les multiples).
    Valeur cyclique : BNA et EBITDA de milieu de cycle (déjà normalisés : multiples des pairs non plafonnés)."""
    cyc = f.get("cyclical")
    pr = [x for x in (f.get("peers") or []) if x.get("fpe")]
    zp = f.get("zb_peers") or []
    z = f.get("zb") or {}
    N = z.get("N")
    sh = f.get("shares")
    nd = f.get("net_debt") or 0.0
    if len(pr) >= 3:
        base, n = "Finviz", len(pr)
        pe_l = [x["fpe"] for x in pr if 0 < x["fpe"] < 80]
        ev_l = [x["ev_ebitda"] for x in pr if x.get("ev_ebitda") and 0 < x["ev_ebitda"] < 40]
        pf_l = [x["pfcf"] for x in pr if x.get("pfcf") and 0 < x["pfcf"] < 80]
        eps_c, ebitda_c = f.get("eps_next_y"), f.get("ebitda")
        names = ", ".join(x["ticker"] for x in pr[:6]) + ("…" if len(pr) > 6 else "")
        self_pe, self_ev = f.get("fwd_pe"), (f.get("dcfx") or {}).get("ev_ebitda")
        lbl_pe = "PER prév."
    elif len(zp) >= 3:
        base, n = "Zonebourse", len(zp)
        pe_l = [x["pe"] for x in zp if x.get("pe") and 0 < x["pe"] < 80]
        ev_l = [x["ev_ebitda"] for x in zp if x.get("ev_ebitda") and 0 < x["ev_ebitda"] < 40]
        pf_l = []
        eps_c = f.get("eps_this_y")
        ebitda_c = (z.get("ebitda") or {}).get(N) if N else None
        if N and (z.get("nd") or {}).get(N) is not None:
            nd = max(z["nd"][N], nd)                    # prudence : la plus élevée des dettes publiée / prévue
        names = ", ".join(x["name"].title()[:22] for x in zp[:5]) + ("…" if len(zp) > 5 else "")
        sf = f.get("zb_self") or {}
        self_pe, self_ev = sf.get("pe"), sf.get("ev_ebitda")
        # VE de Zonebourse incohérent avec capitalisation + dette nette (loyers IFRS 16, minoritaires…) : le VE/EBITDA
        # des pairs n'est pas transposable -> PER seul
        if self_ev and ebitda_c and f.get("mcap") and abs((f["mcap"] + nd) / ebitda_c / self_ev - 1) > 0.3:
            ev_l = []
        lbl_pe = "PER"
    else:
        im = industry_multiples().get(f.get("industry")) if f.get("industry") else None
        if not im:
            return None
        base, n = f"industrie « {f['industry']} »", 0
        pe_l = [im["fpe"]] if im.get("fpe") else []
        pf_l = [im["pfcf"]] if im.get("pfcf") else []
        ev_l, eps_c, ebitda_c, names, self_pe, self_ev, lbl_pe = [], f.get("eps_next_y"), None, "", None, None, "PER prév."
    eps_peak = eps_c
    if cyc and eps_mid:
        eps_c = eps_mid
    ebitda_peak = ebitda_c
    if cyc and ebitda_c:                                # EBITDA de milieu de cycle (moyenne publiée + prévue)
        if base == "Zonebourse" and z.get("ebitda"):
            L = z.get("L")
            ser_ = [v for y, v in sorted(z["ebitda"].items()) if v is not None and L and L - 4 <= y <= N + 1]
            if len(ser_) >= 4:
                ebitda_c = min(ebitda_c, sum(ser_) / len(ser_))
        elif (f.get("dcfx") or {}).get("m_hist"):
            x_ = f["dcfx"]
            ebitda_c = min(ebitda_c, x_["rev0"] * sum(x_["m_hist"]) / len(x_["m_hist"]))
    parts, labels, pos = [], [], []

    def trio(vals, lo_c, hi_c, cap=None):
        m_, a_, b_ = _med(vals), (_quart(vals, 0.25) if len(vals) >= 4 else None), \
            (_quart(vals, 0.75) if len(vals) >= 4 else None)
        cl = lambda v: min(clamp(v, lo_c, hi_c), cap) if cap else clamp(v, lo_c, hi_c)
        m_ = cl(m_)
        return m_, cl(a_) if a_ else m_ * 0.85, cl(b_) if b_ else m_ * 1.15

    peak = []
    if pe_l and eps_c and eps_c > 0:
        m_, a_, b_ = trio(pe_l, p["pe_min"], p["pe_max"])     # cyclique : BNA de milieu de cycle, pas de plafond
        parts.append((eps_c * m_, eps_c * a_, eps_c * b_))
        peak.append((eps_peak * m_, eps_peak * a_, eps_peak * b_) if eps_peak and eps_peak > 0 else parts[-1])
        labels.append(f"{lbl_pe} {m_:.1f}x".replace(".", ","))
        if self_pe and self_pe > 0 and len(pe_l) >= 4:
            q1, q3 = _quart(pe_l, 0.25), _quart(pe_l, 0.75)
            pos.append(f"{lbl_pe} {self_pe:.1f}x".replace(".", ",") + " pour la société, pairs "
                       + f"{q1:.1f}x–{q3:.1f}x (médiane {_med(pe_l):.1f}x)".replace(".", ",")
                       + (" : prime" if self_pe > q3 else " : décote" if self_pe < q1 else " : dans la norme"))
    lev_c = nd / ebitda_c if ebitda_c and ebitda_c > 0 else None
    ev_note = ""
    if ev_l and lev_c is not None and lev_c > 4:
        # dette nette > 4× l'EBITDA : la valeur des fonds propres déduite d'un multiple de VE devient très sensible au
        # multiple (effet de levier) et aux définitions de la dette (loyers IFRS 16…) -> VE/EBITDA non transposé
        ev_l, ev_note = [], f" ; VE/EBITDA des pairs non transposé (dette nette {lev_c:.1f}× l'EBITDA)"
    if ev_l and ebitda_c and ebitda_c > 0 and sh and not f.get("ebitda_proxy"):
        m_, a_, b_ = trio(ev_l, 4.0, 25.0)
        v = lambda m, e: (m * e - nd - (f.get("minority") or 0.0)) / sh
        if v(m_, ebitda_c) > 0:
            parts.append((v(m_, ebitda_c), v(a_, ebitda_c), v(b_, ebitda_c)))
            peak.append((v(m_, ebitda_peak), v(a_, ebitda_peak), v(b_, ebitda_peak)))
            labels.append(f"VE/EBITDA {m_:.1f}x".replace(".", ","))
        if self_ev and self_ev > 0 and len(ev_l) >= 4:
            q1, q3 = _quart(ev_l, 0.25), _quart(ev_l, 0.75)
            pos.append(f"VE/EBITDA {self_ev:.1f}x".replace(".", ",") + " pour la société, pairs "
                       + f"{q1:.1f}x–{q3:.1f}x (médiane {_med(ev_l):.1f}x)".replace(".", ",")
                       + (" : prime" if self_ev > q3 else " : décote" if self_ev < q1 else " : dans la norme"))
    fcf_ps = ((f["fcf_hist"][0] - (f["sbc"] if p["sbc"] else 0)) / sh if f.get("fcf_hist") and sh else None)
    if pf_l and fcf_ps and fcf_ps > 0 and not cyc and not f.get("cash_unrep"):
        m_, a_, b_ = trio(pf_l, 8.0, 25.0)
        parts.append((fcf_ps * m_, fcf_ps * a_, fcf_ps * b_))
        peak.append(parts[-1])
        labels.append(f"P/FCF {m_:.1f}x".replace(".", ","))
    if not parts:
        return None
    avg = lambda xs, i: sum(x[i] for x in xs) / len(xs)
    out = {"mid": avg(parts, 0), "lo": avg(parts, 1), "hi": avg(parts, 2), "label": " · ".join(labels),
           "note": (f"comparables — {n} pairs {base} ({names}) : " if n else f"comparables — moyennes de l'{base} : ")
                   + ", ".join(labels) + " (médianes bornées ; fourchette 1er–3e quartile)"
                   + (" ; BNA et EBITDA de milieu de cycle" if cyc else "") + ev_note,
           "pos": " · ".join(pos)}
    if cyc and eps_mid:
        out["peak"] = (avg(peak, 1), avg(peak, 0), avg(peak, 2))
    return out


def near_zero_base(f):
    """BNA de départ quasi nul ou négatif (sortie de pertes) : une croissance en % n'a pas de sens."""
    e_last = f.get("eps_ttm")
    e1 = f.get("eps_next_y")
    return bool(e1 and e1 > 0 and (e_last is None or e_last <= 0.1 * e1)) or (f.get("eps_5y") or 0) > 100


def fair_value(f, p):
    if f.get("bank"):
        return fair_value_bank(f, p)
    out = {"notes": []}
    g5, capped = growth_expected(f)
    if capped:
        out["notes"].append(f"croissance du BNA ramenée à {g5:.0f} % (CA + 10 points ; consensus 5 ans "
                            + ("non significatif : BNA de départ quasi nul)" if near_zero_base(f) else
                               f"{f['eps_5y']:.0f} %)" if f.get("eps_5y") is not None else "absent, 8 % par défaut)"))
    out["_alerts"] = []
    if f.get("cyclical") and f.get("dcfx_long") and not f.get("zb"):
        f = dict(f, dcfx=f["dcfx_long"], eps_hist_long=f["eps_hist_long_all"])
        out["notes"].append(f["long_note"])
    if f.get("recovery_note"):
        out["notes"].append(f["recovery_note"])
    if f.get("spin_note"):
        out["notes"].append(f["spin_note"])
    if f.get("spin_short"):
        out["_alerts"].append("⚠️ " + f["spin_short"])
    if f.get("recovery_note") and not f.get("cyclical"):
        out["_alerts"].append("ℹ️ redressement après une crise : non classée cyclique")
    elif f.get("turnaround") and not f.get("cyclical"):
        out["_alerts"].append("ℹ️ redressement durable : non classée cyclique")
        out["notes"].append("redressement durable : bond du BNA mais consensus encore en hausse en N+1 et N+2 (marges et CA "
                            "en progression) → non classée cyclique (correction possible dans « Forcer le classement cyclique »)")
    if f.get("low_fpe"):
        out["notes"].append(f"PER prévisionnel {f['fwd_pe']:.1f} < 7 : bénéfices traités comme un haut de cycle")
    # --- 1. DCF d'entreprise (états financiers disponibles) : méthode centrale ---
    res = None
    if f.get("dcfx") and f.get("shares"):
        try:
            g_r = revenue_growth(f)
            if f.get("cyclical"):                     # valeur cyclique : pas de croissance durable tirée du rebond
                g_r = min(g_r, 0.05)
            res = dcf_entreprise(f, p, f["dcfx"], g_r)
        except Exception as e:
            out["notes"].append(f"DCF d'entreprise impossible ({type(e).__name__})")
            res = None
        x_ = f["dcfx"]
        boom = (res and not x_.get("fc") and not f.get("cyclical") and x_.get("capex0", 0) > 0.25
                and x_.get("capex0", 0) > 2.5 * max(x_.get("da_pct") or 0, 0.01))
        if boom:
            # vague d'investissements (ex. Oracle : capex > CA, 7× les amortissements) sans consensus : le DCF
            # prolonge la dépense sans voir le CA qu'elle doit produire -> non retenu ; JV = multiples
            out["notes"].append(f"⚠️ investissements massifs ({x_['capex0'] * 100:.0f} % du CA, "
                                f"{x_['capex0'] / max(x_['da_pct'], 0.01):.0f}× les amortissements) sans consensus "
                                f"pluriannuel : DCF non significatif ({res['central']:.0f}), non retenu — déposez les "
                                f"pages Zonebourse pour un DCF qui intègre le CA attendu de ces investissements")
            out["_alerts"].append("⚠️ DCF non retenu : investissements massifs sans consensus (JV = multiples)")
            res, f = None, dict(f, fcf_hist=[])
            out["_no_dcf_debt"] = True
        if res and res["central"] <= 0:
            lev_ = (f["net_debt"] / f["ebitda"] if f.get("net_debt") is not None and f.get("ebitda")
                    and f["ebitda"] > 0 else None)
            out["notes"].append("DCF d'entreprise négatif (flux trop faibles ou dette élevée) : non retenu ; pas de DCF "
                                "simplifié de secours (il ignorerait la dette)")
            out["_alerts"].append("⚠️ DCF non retenu : " + (f"dette nette {lev_:.1f}× l'EBITDA, " if lev_ else "")
                                  + "la valeur d'entreprise ne couvre pas la dette dans le DCF (JV = multiples)")
            res, f = None, dict(f, fcf_hist=[])
            out["_no_dcf_debt"] = True
    if res:
        out.update({"DCF bas": res["p10"], "DCF": res["central"], "DCF haut": res["p90"],
                    "CMPC (%)": res["wacc"]["wacc"] * 100, "Coût fonds propres (%)": res["wacc"]["ke"] * 100,
                    "Valeur terminale (% VE)": res["scen"][1]["Valeur terminale (% VE)"]})
        if res.get("prob_up") is not None:
            out["Probabilité JV DCF > cours (%)"] = res["prob_up"]
        out["_dcfx"] = res
        try:
            out["_irr"] = dcfx_implied_return(f, res)
        except Exception:  # noqa: BLE001
            out["_irr"] = None
        gg = dcfx_implied_growth(f, p, f["dcfx"], res)
        if gg is not None:
            out["g implicite (%)"], out["g attendue (%)"] = gg[0] * 100, gg[1] * 100
            out["_g5"] = True                         # moyennes sur 5 ans
        out["notes"].append(f"DCF d'entreprise ({res['src']}) : CMPC {res['wacc']['wacc'] * 100:.1f} %, croissance du CA "
                            f"an 1 {res['g_rev'] * 100:.0f} %, 3 scénarios 25/50/25, valeur terminale moyenne croissance "
                            f"perpétuelle {res['tg'] * 100:.1f} % / multiple de sortie {res['exit_mult']:.1f}x EBITDA ; "
                            f"fourchette P10–P90 de 2 000 simulations")
        if f.get("cash_unrep") and "reconstruits" in res.get("src", ""):
            out["_alerts"].append("⚠️ flux de trésorerie non représentatifs : DCF reconstruit depuis le résultat")
            out["notes"].append("⚠️ flux de trésorerie publiés non représentatifs (flux d'exploitation des 3 derniers "
                                "exercices < 40 % du résultat + amortissements) : flux du DCF reconstruits depuis le "
                                "résultat d'exploitation, P/FCF exclu des comparables")
        if asset_light(f["dcfx"]):
            out["notes"].append(f"amortissements ({f['dcfx']['da_pct'] * 100:.1f} % du CA) surtout d'incorporels "
                                f"rachetés : investissements d'entretien = investissements réels "
                                f"({max(f['dcfx']['capex0'], f['dcfx']['capex_norm']) * 100:.1f} % du CA, +25 %)")
        if res.get("gap_tv") and res["gap_tv"] > 0.3:
            out["notes"].append(f"⚠️ valeurs terminales divergentes ({res['gap_tv'] * 100:.0f} %) : hypothèses à vérifier")
    hist = [] if res else f["fcf_hist"]
    if hist and f.get("cash_unrep"):
        out["notes"].append("⚠️ flux de trésorerie publiés non représentatifs de l'activité (flux libre très inférieur "
                            "au résultat) : pas de DCF sur le flux libre, ni de P/FCF dans les comparables")
        out["_alerts"].append("⚠️ flux de trésorerie non représentatifs : pas de DCF sur le flux libre")
        hist = []
    if hist:
        fcf = min(hist[0], sum(hist) / len(hist)) if f["cyclical"] else hist[0]   # cyclique : jamais au-dessus du flux actuel
        if f["cyclical"]:
            out["notes"].append("cyclique : flux normalisé (moyenne TTM + 3 exercices, plafonnée au flux actuel)")
        elif f.get("fcf_norm") is not None:
            fcf = f["fcf_norm"]
            out["notes"].append("flux libre irrégulier : flux normalisé (médiane publiée + prévue, ≤ 80 % de l'EBITDA)")
        if p["sbc"]:
            fcf -= f["sbc"]
        g0 = clamp(g5 / 100, 0.0, p["gcap"])
        if f.get("sales_cagr") is not None:
            g0 = clamp(min(g0, f["sales_cagr"] / 100), 0.0, p["gcap"])   # le flux suit le CA, pas le rebond du BNA
        if f["cyclical"]:
            g0 = min(g0, 0.05)
        if fcf > 0 and f["shares"]:
            ke, how = cost_of_equity(f, p)
            ke = max(ke, p["tg"] + 0.03)
            v = lambda w: dcf_value(fcf, g0, p["tg"], w) / f["shares"]
            out.update({"DCF bas": v(ke + 0.01), "DCF": v(ke), "DCF haut": v(ke - 0.01),
                        "Coût fonds propres (%)": ke * 100})
            out["_dcf"] = (fcf, g0, ke, f["shares"])
            # part de la valeur terminale (au-delà de 10 ans) dans le DCF : > 80 % = valeur très dépendante du long terme
            tv_share = 1 - dcf_value(fcf, g0, p["tg"], ke, years=10, terminal=False) / dcf_value(fcf, g0, p["tg"], ke)
            out["Valeur terminale (% du DCF)"] = tv_share * 100
            if tv_share > 0.8:
                out["notes"].append(f"⚠️ valeur terminale = {tv_share * 100:.0f} % du DCF : très sensible au long terme")
            wi = wacc_info(f, ke, p)
            if wi:
                out["WACC indicatif (%)"] = wi * 100
            gi = implied_growth(fcf, f["price"], f["shares"], p["tg"], ke)
            if gi is not None:                       # cyclique : croissance attendue = celle du DCF (≤ 5 %), pas le rebond
                out["g implicite (%)"], out["g attendue (%)"] = gi * 100, (g0 * 100 if f["cyclical"] else g5)
            out["notes"].append(f"DCF : flux {fcf / 1e3:.1f} Md {f.get('currency', 'USD')}, g {g0 * 100:.0f} %→{p['tg'] * 100:.1f} %, "
                                f"coût des fonds propres {ke * 100:.1f} % ({how})"
                                + (f" ; WACC indicatif {wi * 100:.1f} %" if wi else ""))
        else:
            out["notes"].append("pas de DCF (flux libre négatif après rémunération en actions)")
    # --- 2. PER cible x BNA N+1 (redressement : BNA N+2 actualisé d'un an) ---
    eps = f["eps_next_y"]
    e_lbl = f"BNA N+1 {eps:.2f}" if eps else ""
    eps_mid = None
    if f["cyclical"]:                                  # milieu de cycle : jamais le BNA d'une année de pic ou de creux
        eps_mid, n_mid, n_pub = cycle_eps(f)
        if eps_mid and eps_mid > 0 and eps:
            out["_eps_mid"], out["_eps_peak"] = eps_mid, eps
            e_lbl = f"BNA de milieu de cycle {eps_mid:.2f} (moyenne de {n_mid} exercices publiés et prévus)"
            eps = eps_mid
            out["notes"].append(f"cyclique : multiples appliqués au BNA de milieu de cycle {eps_mid:.2f} (moyenne de "
                                f"{n_mid} exercices) au lieu du BNA N+1 {out['_eps_peak']:.2f}")
            if n_pub < 5:
                out["_alerts"].append(f"ℹ️ historique court ({n_pub} exercices publiés) : milieu de cycle approximatif")
        else:
            eps_mid = None
            if eps:
                out["notes"].append("cyclique : BNA de milieu de cycle non calculable (historique trop court ou "
                                    "moyenne négative) — multiples sur le BNA N+1")
    rec = recovery_eps(f, p)
    if rec:
        eps, e_lbl = rec[0], rec[1]
        out["_alerts"].append("ℹ️ redressement : PER cible sur le BNA N+2 actualisé")
    if eps and eps > 0:
        if f["cyclical"]:
            # BNA de milieu de cycle (déjà normalisé) : multiple « normalisé » 10–12x ; BNA N+1 (pic possible) : 8x
            pe = p.get("pe_cyc_mid", 11.0) if eps_mid else p["pe_cyc"]
        else:
            pe = clamp(p["pe_base"] + 0.4 * (min(g5, 25) - 10), p["pe_min"], p["pe_max"])
        if f.get("pe_hist") and not f["cyclical"] and pe > 1.25 * f["pe_hist"]:
            pe = 1.25 * f["pe_hist"]
            out["notes"].append(f"PER cible plafonné à 1,25 × le PER historique ({f['pe_hist']:.1f}x)")
        lev_ = (f["net_debt"] / f["ebitda"] if f.get("net_debt") is not None and f.get("ebitda") and f["ebitda"] > 0
                else None)
        if lev_ and lev_ > 3 and not f["cyclical"]:
            # actionnaire très endetté : bénéfice plus risqué, multiple plus bas (−10 % par point de dette nette /
            # EBITDA au-delà de 3, au plus −40 %)
            k_ = max(0.6, 1 - 0.1 * (lev_ - 3))
            pe *= k_
            out["notes"].append(f"PER cible réduit de {(1 - k_) * 100:.0f} % : dette nette {lev_:.1f}× l'EBITDA "
                                f"(au-delà de 3×, −10 % par point)")
        out.update({"PER bas": eps * pe * 0.85, "PER": eps * pe, "PER haut": eps * pe * 1.15})
        out.setdefault("_mult", {})["PER"] = f"PER cible {pe:.1f}x × {e_lbl.split(' (')[0].split(' —')[0]}"
        out["notes"].append(f"PER cible {pe:.1f}x sur {e_lbl}")
    eps = eps_mid or f["eps_next_y"]                    # comparables : même BNA que le PER cible
    # --- 3. comparables : pairs Finviz (lien « Peers ») ou pairs Zonebourse (onglet Valorisation), sinon moyennes de
    # l'industrie ; PER, VE/EBITDA (tient compte de la dette) et P/FCF ; médiane des pairs, fourchette 1er–3e quartile
    comp = comparables(f, p, eps_mid)
    if comp:
        out.update({"Comparables bas": comp["lo"], "Comparables": comp["mid"], "Comparables haut": comp["hi"]})
        out.setdefault("_mult", {})["Comparables"] = comp["label"]
        out["notes"].append(comp["note"])
        if comp.get("pos"):
            out["_peers_txt"] = comp["pos"]
        if comp.get("peak"):
            out["_comp_peak"] = comp["peak"]
    # --- synthèse ---
    add_hist_pe(f, out)
    if out.get("_eps_mid"):                    # même synthèse, multiples sur le BNA N+1 : « si les bénéfices durent »
        k = out["_eps_peak"] / out["_eps_mid"]
        pk = {kk: (list(v) if isinstance(v, list) else v) for kk, v in out.items()}
        for m in ("PER", "Comparables"):
            for sfx in ("", " bas", " haut"):
                if pk.get(m + sfx):
                    pk[m + sfx] *= k
        if out.get("_comp_peak"):                  # comparables recalculés sur le BNA N+1 et l'EBITDA actuel
            pk["Comparables bas"], pk["Comparables"], pk["Comparables haut"] = out["_comp_peak"]
        if out.get("_dcfx"):
            synthese_dcf(pk, p.get("w_dcf", 0.5))
        else:
            synthese(pk, ("DCF", "PER", "PER hist.", "Comparables"))
        if pk.get("JV centrale"):
            out.update({"JV si pic": pk["JV centrale"], "JV si pic basse": pk.get("JV basse"),
                        "JV si pic haute": pk.get("JV haute")})
    if out.get("_dcfx"):
        synthese_dcf(out, p.get("w_dcf", 0.5))
        if out.get("JV centrale") and f["price"]:
            out["Écart (%)"] = (out["JV centrale"] / f["price"] - 1) * 100
        return out
    synthese(out, ("DCF", "PER", "PER hist.", "Comparables"))
    if sum(1 for m in ("DCF", "PER", "PER hist.", "Comparables") if out.get(m)) == 1:
        out["notes"].append("⚠️ une seule méthode disponible : juste valeur fragile " + (
            "(DCF écarté : dette trop élevée)" if out.get("_no_dcf_debt") else
            "(pour une page Zonebourse, ajoutez l'onglet Finances pour le DCF)"))
    if out.get("JV centrale") and f["price"]:
        out["Écart (%)"] = (out["JV centrale"] / f["price"] - 1) * 100
    precision(out, ("DCF", "PER", "PER hist.", "Comparables"))
    if f["cyclical"] and "DCF" in (out.get("_excl") or []):
        out["_peak_only"] = True                     # PER et comparables appliqués au même BNA de pic : pas indépendants
    return out


def cycle_eps(f):
    """BNA de milieu de cycle (valeur cyclique) : moyenne des BNA publiés (jusqu'à 5 exercices) et prévus (N, N+1).
    -> (BNA, nb d'exercices) ou (None, n) si moins de 4 valeurs."""
    ser = [(v, kind) for lbl, v, kind in _eps_series(f)
           if v is not None and (kind == "publié" or lbl in ("N (en cours)", "N+1")
                                 or (f.get("zb") and f["zb"].get("N") and int(lbl) <= f["zb"]["N"] + 1))]
    vals = [v for v, _ in ser]
    n_pub = sum(1 for _, k in ser if k == "publié")
    if len(vals) < 4:
        return None, len(vals), n_pub
    return sum(vals) / len(vals), len(vals), n_pub


def recovery_eps(f, p):
    """Redressement (pages Zonebourse, valeur non cyclique) : BNA N+2 ≥ 2 × BNA N+1 (ou BNA N+1 ≤ 0 et N+2 > 0) —
    le BNA N+1 est celui du creux ; on retient le BNA N+2 du consensus actualisé d'un an au coût des fonds propres.
    -> (BNA retenu, libellé) ou None."""
    z = f.get("zb")
    if not z or f.get("cyclical") or f.get("bank"):
        return None
    N = z.get("N")
    b = z.get("bna") or {}
    e1, e2 = b.get(N + 1) if N else None, b.get(N + 2) if N else None
    if not e2 or e2 <= 0 or e1 is None or (e1 > 0 and e2 < 2 * e1):
        return None
    ke, _ = cost_of_equity(f, p)
    ke = max(ke, 0.08)
    v = e2 / (1 + ke)
    return v, (f"BNA N+2 {e2:.2f} actualisé à {v:.2f} ({ke * 100:.1f} %) — redressement : BNA N+1 "
               f"{e1:.2f} (Zonebourse), creux non représentatif")


def synthese_dcf(out, w_dcf=0.5):
    """DCF d'entreprise au centre : JV = w × DCF + (1 − w) × médiane des multiples (PER, comparables, PER historique) ;
    fourchette = même pondération des bornes (P10–P90 du DCF, bas–haut des multiples) ;
    précision = écart des deux piliers (DCF / multiples) à la JV."""
    mult = [(out[m], m) for m in ("PER", "Comparables", "PER hist.") if out.get(m)]
    dcf = out["DCF"]
    if not mult:
        out.update({"JV centrale": dcf, "JV basse": out["DCF bas"], "JV haute": out["DCF haut"]})
        out["notes"].append("JV = DCF seul (pas de multiple disponible)")
        out["_keep"], out["_excl"] = ["DCF"], []
        return
    mc = _med([v for v, _ in mult])
    keep = [m for v, m in mult if abs(v / mc - 1) <= 0.5] or [m for _, m in mult]
    lo, hi = min(out[m + " bas"] for m in keep), max(out[m + " haut"] for m in keep)
    out["Multiples (médiane)"] = mc
    out["JV centrale"] = w_dcf * dcf + (1 - w_dcf) * mc
    out["JV basse"] = w_dcf * out["DCF bas"] + (1 - w_dcf) * lo
    out["JV haute"] = w_dcf * out["DCF haut"] + (1 - w_dcf) * hi
    jv = out["JV centrale"]
    out["Dispersion (%)"] = max(abs(dcf / jv - 1), abs(mc / jv - 1)) * 100
    out["_keep"], out["_excl"] = ["DCF"] + keep, [m for _, m in mult if m not in keep]
    out["_mode"] = "dcf"
    out["_w_dcf"] = w_dcf
    out["notes"].append(f"JV = {w_dcf * 100:.0f} % DCF ({dcf:.0f}) + {(1 - w_dcf) * 100:.0f} % multiples ({mc:.0f}, "
                        f"médiane de {', '.join(m for _, m in mult)})")


def precision(out, methods):
    """Précision de la JV : écart maximal à la JV centrale des méthodes retenues (celles qui s'écartent de plus de
    50 % de la médiane sont signalées à part, pas comptées)."""
    keep = out.get("_keep") or [m for m in methods if out.get(m)]
    vals = [out[m] for m in keep if out.get(m)]
    jv = out.get("JV centrale")
    if jv and len(vals) >= 2:
        out["Dispersion (%)"] = max(abs(v / jv - 1) for v in vals) * 100


def analyst_style_target(f):
    """Cours visé à 12 mois « façon analystes » (méthode des multiples historiques, ex. BofA) : VE/EBIT moyen des
    5 derniers exercices × EBIT N+1, moins la dette nette N+1, par action. -> dict ou None (onglet Valorisation
    Zonebourse nécessaire)."""
    z = f.get("zb") or {}
    hist, ebit, nd, N = z.get("ev_ebit_hist") or [], z.get("ebit") or {}, z.get("nd") or {}, z.get("N")
    sh = f.get("shares")
    if len(hist) < 3 or not N or not sh or not ebit.get(N + 1) or ebit[N + 1] <= 0:
        return None
    nd1 = nd.get(N + 1) if nd.get(N + 1) is not None else nd.get(N, f.get("net_debt") or 0.0)
    mean = sum(hist) / len(hist)
    med = sorted(hist)[len(hist) // 2] if len(hist) % 2 else sum(sorted(hist)[len(hist) // 2 - 1:len(hist) // 2 + 1]) / 2
    v = lambda m: (m * ebit[N + 1] - nd1) / sh
    out = {"mult": mean, "mult_med": med, "ebit1": ebit[N + 1], "nd1": nd1, "target": v(mean), "target_med": v(med),
           "n": len(hist), "lo": min(hist), "hi": max(hist)}
    if ebit.get(N) and ebit[N] > 0:
        out["g_ebit"] = (ebit[N + 1] / ebit[N] - 1) * 100
        out["now"] = (mean * ebit[N] - (nd.get(N) if nd.get(N) is not None else nd1)) / sh   # même multiple, EBIT N
    # garde-fous : multiples passés mesurés sur des bénéfices bien plus petits (hyper-croissance, ex. Nvidia), ou
    # méthode qui ne reproduit pas l'objectif réel (ex. multiples d'avant une crise) -> explication seulement
    L_ = z.get("L")
    cur_mult = (f["mcap"] + nd1) / ebit[N + 1] if f.get("mcap") else None
    if L_ and ebit.get(L_) and ebit[L_] > 0 and ebit[N + 1] >= 2.5 * ebit[L_] and cur_mult and mean > 2 * cur_mult:
        out["skip"] = (f"multiples passés ({mean:.1f}x".replace(".", ",") + f") mesurés sur un EBIT "
                       f"{ebit[N + 1] / ebit[L_]:.1f}".replace(".", ",") + " fois plus "
                       f"petit que l'EBIT N+1 : la méthode des multiples historiques ne s'applique pas")
    elif f.get("target") and abs(out["target"] / f["target"] - 1) > 0.35:
        out["skip"] = (f"VE/EBIT moyen des {len(hist)} derniers exercices ({mean:.1f}x".replace(".", ",") + ") × EBIT N+1 = "
                       f"{out['target']:.0f}, loin de l'objectif réel ({f['target']:.0f}) : les analystes n'utilisent "
                       f"visiblement pas les multiples passés ici (crise, changement de profil…)")
    return out


def analyst_style_txt(a, f, cur):
    """Phrase d'explication de la ligne « objectif à 12 mois façon analystes »."""
    if a.get("skip"):
        return "Façon analystes : " + a["skip"] + "."
    x_ = lambda v: f"{v:.1f}".replace(".", ",") + "x"
    eb_ = f"{a['ebit1']:,.0f}".replace(",", " ")
    t = (f"Façon analystes : VE/EBIT moyen des {a['n']} derniers exercices ({x_(a['mult'])}, de {x_(a['lo'])} à "
         f"{x_(a['hi'])}) × EBIT N+1 ({eb_} M{cur}"
         + (f", {a['g_ebit']:+.0f} % sur N" if a.get("g_ebit") is not None else "") + ") = "
         f"{a['target']:.0f} {cur} (médiane du multiple : {a['target_med']:.0f}). C'est un cours visé dans 12 mois : il "
         "suppose que l'EBIT prévu se réalise et que le marché garde son multiple moyen.")
    if a.get("now"):
        t += (f" Le même multiple appliqué à l'EBIT de l'exercice en cours donne {a['now']:.0f} {cur}, à comparer à la "
              "juste valeur.")
    if f.get("cyclical"):
        t += " Valeur cyclique : EBIT N+1 possiblement proche d'un pic."
    return t


def target_bridge(f, fv):
    """Pont de l'objectif des analystes : objectif × titres − trésorerie nette N+1 = VE implicite, traduite en
    multiples (VE/EBIT N+1, PER N+1) et comparée à ceux qu'implique la juste valeur. Montre si l'écart JV / objectif
    vient du multiple ou de la trésorerie (déjà comptée des deux côtés). -> dict ou None."""
    tg, jv, sh = f.get("target"), fv.get("JV centrale"), f.get("shares")
    if not tg or not jv or not sh or sh <= 0:
        return None
    z = f.get("zb") or {}
    N, ebit, ndz = z.get("N"), z.get("ebit") or {}, z.get("nd") or {}
    nd = ndz.get(N + 1) if N and ndz.get(N + 1) is not None else (ndz.get(N) if N and ndz.get(N) is not None
                                                                  else f.get("net_debt"))
    if nd is None:
        return None
    eb = ebit.get(N + 1) if N else None
    eps = f.get("eps_next_y")
    side = lambda v: {"eq": v * sh, "ev": v * sh + nd,
                      "ev_ebit": (v * sh + nd) / eb if eb and eb > 0 and (v * sh + nd) / eb >= 2 else None,
                      "pe": v / eps if eps and eps > 0 else None}
    out = {"nd": nd, "ebit": eb, "eps": eps, "tg": side(tg), "jv": side(jv), "gap": tg - jv,
           "cash_ps": -nd / sh}
    return out


def target_bridge_txt(b, f, cur, fv_pk=None):
    """Phrase « pont de l'objectif » pour l'écran et le PDF."""
    if not b:
        return ""
    md = lambda v: f"{v / 1000:,.1f}".replace(",", " ").replace(".", ",") + f" Md{cur}"
    x_ = lambda v: f"{v:.1f}".replace(".", ",") + "x"
    t, j = b["tg"], dict(b["jv"])
    if not t["ev_ebit"]:
        j["ev_ebit"] = None
    tr = (f"− trésorerie nette N+1 {md(-b['nd'])}" if b["nd"] < 0 else f"+ dette nette N+1 {md(b['nd'])}")
    mult = lambda s: " et ".join(m for m in ((x_(s["ev_ebit"]) + " l'EBIT N+1") if s["ev_ebit"] else None,
                                              (x_(s["pe"]) + " le BNA N+1") if s["pe"] else None) if m)
    if not mult(t):
        return ""
    txt = (f"Pont de l'objectif : {f['target']:.0f} {cur} × {f['shares']:.0f} M titres = capitalisation "
           f"{md(t['eq'])} {tr} → VE {md(t['ev'])}, soit {mult(t)}. La juste valeur ({b['jv']['eq'] / f['shares']:.0f} "
           f"{cur}) correspond à {mult(j)}. ")
    if abs(b["gap"]) < 0.05 * f["target"]:
        return txt + "Juste valeur et objectif sont cohérents."
    if f.get("cyclical") and b["gap"] > 0:
        return txt + (f"L'écart ({b['gap']:+.0f} {cur}/action) vient surtout de la base de bénéfices : valeur cyclique, "
                      "la juste valeur applique un multiple normal à un bénéfice de milieu de cycle, l'objectif des "
                      "analystes au BNA N+1 (haut de cycle possible)"
                      + (" — voir la ligne « si les bénéfices actuels durent »." if fv_pk else "."))
    sens = "prime" if b["gap"] > 0 else "décote"
    cash = (f"la trésorerie nette ({b['cash_ps']:.0f} {cur}/action) est déjà comptée des deux côtés" if b["nd"] < 0
            else f"la dette nette ({-b['cash_ps']:.0f} {cur}/action) est déduite des deux côtés")
    return txt + (f"L'écart ({b['gap']:+.0f} {cur}/action) vient du multiple — les analystes appliquent une {sens} "
                  f"de {abs(t['pe'] / j['pe'] - 1) * 100 if t['pe'] and j['pe'] else abs(t['ev'] / j['ev'] - 1) * 100:.0f} % "
                  f"(croissance, qualité ou risque jugés autrement) — pas du passage VE → capitaux propres : {cash}.")


def long_history(f):
    """Valeur cyclique US + option EDGAR : cycle mesuré sur 10 exercices — états Finviz pour les années récentes,
    10-K de la SEC pour les plus anciennes (marge EBITDA, marge de trésorerie d'exploitation, BNA publié)."""
    ed, x = f["edgar"], dict(f["dcfx"])
    fin = sorted(int(l) for l, _ in f.get("eps_hist") or [] if str(l).isdigit())
    if not fin or not x:
        return None
    m_years = fin[-5:]                                    # marge EBITDA Finviz : 5 derniers exercices
    add_m = [(ed[y]["ebit"] + ed[y]["da"]) / ed[y]["rev"] for y in sorted(ed, reverse=True)
             if y < m_years[0] and all(ed[y].get(k) is not None for k in ("ebit", "da")) and ed[y].get("rev")]
    add_m = add_m[:10 - len(x.get("m_hist") or [])]
    tax = x.get("tax") or 0.21
    add_o = [(ed[y]["ocf"] + (ed[y].get("interest") or 0) * (1 - tax) - (ed[y].get("sbc") or 0)) / ed[y]["rev"]
             for y in sorted(ed, reverse=True) if y < fin[0] and ed[y].get("ocf") is not None and ed[y].get("rev")]
    add_o = add_o[:10 - len(x.get("ocf_hist") or [])]
    # BNA des années antérieures : résultat net ÷ actions (corrigées des divisions, ramenées à la base Finviz)
    ch = f.get("cap_hist") or {}
    sh = dict(ES.split_adjusted_shares([(y, ed[y]["shares"]) for y in sorted(ed) if ed[y].get("shares")]))
    ov = [y for y in sh if ch.get(y, {}).get("shares")]
    eps_old = []
    if ov:
        k = ch[max(ov)]["shares"] / sh[max(ov)]
        eps_old = [(str(y), ed[y]["ni"] / (sh[y] * k)) for y in sorted(ed)
                   if y < fin[0] and ed[y].get("ni") is not None and sh.get(y)]
    eps_long = (eps_old + list(f.get("eps_hist") or []))[-10:]
    if not (add_m or add_o or len(eps_long) > 5):
        return None
    x["m_hist"] = list(x.get("m_hist") or []) + add_m
    x["ocf_hist"] = list(x.get("ocf_hist") or []) + add_o
    yrs = [int(l) for l, _ in eps_long]
    note = (f"historique long (10-K de la SEC) : cycle mesuré sur {len(x['m_hist'])} exercices pour la marge, "
            f"{len(x['ocf_hist'])} pour la trésorerie, {len(eps_long)} BNA publiés ({min(yrs)}–{max(yrs)})")
    return x, eps_long, note


def dcfx_implied_return(f, res):
    """Rendement annuel implicite au cours : taux d'actualisation qui ramène la valeur du scénario central du DCF au
    cours (si ces flux se réalisent, l'acheteur au cours obtient ce rendement). -> (taux, borne « ≤ »/« ≥ »/"") ou None."""
    pr, sh = f.get("price"), f.get("shares")
    if not (pr and sh and res):
        return None
    x, tg, em, cyc, g = res["x"], res["tg"], res["exit_mult"], bool(f.get("cyclical")), res["g_rev"]
    v = lambda w: (lambda r: (r["eq_perp"] + r["eq_exit"]) / 2 / sh)(dcfx_project(x, g, 0.0, w, tg, em, cyc))
    lo, hi = tg + 0.01, 0.40
    if v(lo) < pr:
        return lo, "≤"
    if v(hi) > pr:
        return hi, "≥"
    for _ in range(40):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if v(mid) > pr else (lo, mid)
    return (lo + hi) / 2, ""


def epv_value(f, fv):
    """Valeur sans croissance (EPV, Greenwald) : résultat d'exploitation normalisé après impôt ÷ CMPC − dette nette,
    par action ; investissements de maintien supposés égaux aux amortissements. Cyclique : EBIT moyen du cycle.
    -> dict(ps, how, g_price, g_target…) ou None."""
    if f.get("bank") or f.get("insurer"):
        return None
    sh, pr = f.get("shares"), f.get("price")
    w = (fv.get("CMPC (%)") or fv.get("Coût fonds propres (%)") or 0) / 100
    if not sh or not w:
        return None
    z = f.get("zb") or {}
    x = (f.get("dcfx_long") if f.get("cyclical") and f.get("dcfx_long") and not z else f.get("dcfx")) or {}
    N, L = z.get("N"), z.get("L")
    ebit, how = None, ""
    if z.get("ebit"):
        e = z["ebit"]
        if f.get("cyclical"):
            vals = [e[y] for y in sorted(e) if e[y] is not None and (not L or y >= L - 4) and (not N or y <= N + 1)]
            if len(vals) >= 4:
                ebit, how = sum(vals) / len(vals), f"EBIT moyen du cycle ({len(vals)} exercices)"
        if ebit is None:
            yy = N if N and e.get(N) is not None else L
            if yy and e.get(yy) is not None:
                ebit, how = e[yy], f"EBIT {yy}" + (" (exercice en cours)" if yy == N else "")
        ndz = z.get("nd") or {}
        nd = ndz.get(N) if N and ndz.get(N) is not None else ndz.get(L)
        nd = nd if nd is not None else x.get("net_debt")
        tax = 0.25
    elif x:
        if f.get("cyclical") and x.get("m_hist"):
            ebit = x["rev0"] * (sum(x["m_hist"]) / len(x["m_hist"]) - (x.get("da_pct") or 0))
            how = f"EBIT de milieu de cycle (marge moyenne de {len(x['m_hist'])} exercices)"
        else:
            ebit, how = x.get("ebit"), "EBIT des 12 derniers mois"
        nd = (x.get("net_debt") or 0) + (x.get("minority") or 0)
        tax = x.get("tax") or 0.21
    else:
        return None
    if ebit is None or nd is None:
        return None
    ps = (ebit * (1 - tax) / w - nd) / sh
    out = {"ps": ps, "ebit": ebit, "how": how, "tax": tax, "w": w, "nd": nd,
           "cash_heavy": bool(f.get("mcap") and nd < 0 and -nd > 0.5 * f["mcap"])}
    if pr:
        out["g_price"] = (1 - ps / pr) * 100
    if f.get("target"):
        out["g_target"] = (1 - ps / f["target"]) * 100
    return out


def net_dilution(f):
    """Évolution annuelle du nombre d'actions (jusqu'à 5 ans, exercices publiés) avec rachats et rémunération en
    actions des 3 derniers exercices. Source la plus longue : 10-K (SEC), états Finviz ou Zonebourse. -> dict ou None."""
    z, ed, ch = f.get("zb") or {}, f.get("edgar") or {}, f.get("cap_hist") or {}
    ser, src, flows = [], None, {}
    if z.get("titres") and z.get("L"):
        ser, src = [(y, v) for y, v in sorted(z["titres"].items()) if v and y <= z["L"]], "Zonebourse"
    elif ch:
        ser, src, flows = [(y, ch[y]["shares"]) for y in sorted(ch) if ch[y].get("shares")], "états Finviz", ch
    ed_ser = [(y, ed[y]["shares"]) for y in sorted(ed) if ed[y].get("shares")]
    if ed and not z and len(ed_ser) > len(ser):
        ser, src, flows = ed_ser, "10-K de la SEC", ed
    ser = ES.split_adjusted_shares(ser)[-6:]
    if len(ser) < 3:
        return None
    (y0, s0), (y1, s1) = ser[0], ser[-1]
    if y1 <= y0 or s0 <= 0:
        return None
    yrs = [y for y, _ in ser][-3:]
    avg = lambda k: (lambda v: sum(v) / len(v) if v else None)(
        [flows[y][k] for y in yrs if (flows.get(y) or {}).get(k) is not None])
    return {"cagr": ((s1 / s0) ** (1 / (y1 - y0)) - 1) * 100, "y0": y0, "y1": y1, "bb": avg("buyback"),
            "sbc": avg("sbc"), "src": src}


def dilution_txt(d, cur):
    """-> (valeur, lecture, signe ✅ 1 / ➖ 0 / ⚠️ −1)."""
    c = d["cagr"]
    m = lambda v: f"{v:,.0f}".replace(",", " ") + f" M{cur}/an"
    v = f"{c:+.1f} %/an ({d['y0']}–{d['y1']})".replace(".", ",")
    bb, sbc = d.get("bb"), d.get("sbc")
    if c <= -1:
        t = "relution : moins d'actions, chaque action détient une part croissante des bénéfices"
        t += f" (rachats {m(bb)})" if bb else ""
        return v, t, 1
    if c < 1:
        if bb and sbc and sbc >= 0.5 * bb:
            return v, (f"nombre d'actions stable : les rachats ({m(bb)}) servent surtout à compenser la rémunération "
                       f"en actions ({m(sbc)})"), 0
        return v, "nombre d'actions stable", 0
    t = "dilution : les actionnaires actuels voient leur part des bénéfices diminuer"
    t += " (émissions d'actions : augmentation de capital ou acquisition payée en titres ?)" if c >= 4 else ""
    t += f", malgré {m(bb)} de rachats" if bb else ""
    t += f" (rémunération en actions {m(sbc)})" if sbc else ""
    return v, t, -1 if c >= 2 else 0


def value_markers(f, fv, cur):
    """Repères complémentaires de la section 7 : valeur sans croissance, rendement implicite au cours, dilution.
    -> [[repère, valeur, lecture]]."""
    rows = []
    ep = epv_value(f, fv)
    if ep:
        x_ = lambda v: f"{v * 100:.1f}".replace(".", ",") + " %"
        t = (f"{ep['how']} × (1 − impôt {ep['tax'] * 100:.0f} %) ÷ CMPC {x_(ep['w'])} − dette nette, sans croissance. ")
        if ep["ps"] <= 0:
            t += ("Valeur négative : la dette dépasse la valeur des bénéfices actuels, tout le cours repose sur la "
                  "croissance ou le redressement")
        elif ep.get("g_price") is not None:
            t += ("Le cours est inférieur à cette valeur : le marché anticipe un recul des bénéfices"
                  if ep["g_price"] < 0 else f"Le cours paie {ep['g_price']:.0f} % pour la croissance")
            if ep.get("g_target") is not None and ep["g_target"] > 0:
                t += f", l'objectif des analystes {ep['g_target']:.0f} %"
        t += "."
        if ep.get("cash_heavy"):
            t += (" ⚠️ Trésorerie nette supérieure à la moitié de la capitalisation (banque captive, avances clients…) : "
                  "valeur et rendement implicite à prendre avec prudence.")
        rows.append(["Valeur sans croissance (EPV)", f"{ep['ps']:.0f} {cur}", t])
    ir = fv.get("_irr")
    if ir:
        r, bound = ir
        w = (fv.get("CMPC (%)") or 0) / 100
        d = (r - w) * 100
        dot = "🟢" if d >= 1.5 else "🔴" if d <= -1.5 else "⚪"
        rows.append(["Rendement implicite au cours", f"{bound}{r * 100:.1f} %/an".replace(".", ","),
                     f"{dot} si les flux du DCF central se réalisent, contre {w * 100:.1f} % exigé (CMPC)".replace(".", ",")
                     + (f" : marge de sécurité de {d:.1f} pt".replace(".", ",") if d >= 1.5 else
                        " : rendement inférieur au rendement exigé" if d <= -1.5 else " : proche du rendement exigé")])
    dl = net_dilution(f)
    if dl:
        v, t, sg = dilution_txt(dl, cur)
        rows.append(["Nombre d'actions", v, f"{ICON_P[sg]} {t} — {dl['src']}"])
    return rows


def premium_reasons(f, fv, force=False):
    """« Pourquoi cette prime ? » : quand l'objectif des analystes dépasse nettement la JV (PER implicite ≥ +10 %) ou
    que la valeur cote nettement plus cher que ses pairs (PER ≥ +25 %), confronte la prime aux critères chiffrés qui
    la justifient d'ordinaire (croissance du BNA, PEG, rentabilité, bilan, conversion en trésorerie, historique).
    -> dict(prem_tg, prem_peer, pe_tg, pe_jv, pe_self, pe_peers, rows[(critère, valeur, lecture, signe)], verdict) ou None."""
    if f.get("bank") or f.get("insurer"):
        return None
    b = target_bridge(f, fv)
    pe_tg = b["tg"]["pe"] if b else None
    pe_jv = b["jv"]["pe"] if b else None
    prem_tg = (pe_tg / pe_jv - 1) * 100 if pe_tg and pe_jv else None
    peers = [q["pe"] for q in f.get("zb_peers") or [] if q.get("pe") and 0 < q["pe"] < 100]
    pe_self = (f.get("zb_self") or {}).get("pe") if peers else None
    if not peers:
        peers = [q["fpe"] for q in f.get("peers_data") or [] if q.get("fpe") and 0 < q["fpe"] < 100]
        pe_self = f.get("fwd_pe")
    pe_peers = _med(peers) if len(peers) >= 3 else None
    prem_peer = (pe_self / pe_peers - 1) * 100 if pe_self and pe_peers and pe_self > 0 else None
    if not force and not ((prem_tg is not None and prem_tg >= 10) or (prem_peer is not None and prem_peer >= 25)):
        return None
    pc = lambda v: f"{v:.0f} %"
    rows = []
    z = f.get("zb") or {}
    N, L = z.get("N"), z.get("L")
    bna, ebit, rn, fcf = z.get("bna") or {}, z.get("ebit") or {}, z.get("rn") or {}, z.get("fcf") or {}
    # 1. croissance du BNA (dernier exercice publié -> N+1), sinon croissance Finviz
    g = None
    if N and L and bna.get(L) and bna.get(N + 1) and bna[L] > 0.15 * bna[N + 1] > 0 and N + 1 > L:
        g = ((bna[N + 1] / bna[L]) ** (1 / (N + 1 - L)) - 1) * 100
        g_lbl = f"BNA {L} → {N + 1} : {bna[L]:.2f} → {bna[N + 1]:.2f}".replace(".", ",")
    elif N and L and bna.get(N + 1) and bna[N + 1] > 0 and bna.get(L) is not None:
        rows.append(("Croissance du BNA", "n.s.", f"BNA {L} quasi nul ou négatif ({bna[L]:.2f}) : croissance ".replace(".", ",")
                     + "non mesurable, la prime parie sur le redressement", -1))
    elif not N and (f.get("eps_5y") is not None or f.get("eps_next_y_pct") is not None):
        g = f.get("eps_5y") if f.get("eps_5y") is not None else f["eps_next_y_pct"]
        g_lbl = "croissance annuelle attendue du BNA (Finviz)"
    if g is not None:
        rows.append(("Croissance du BNA", f"{g:+.0f} %/an", g_lbl + (" — forte, justifie un multiple élevé" if g >= 20
                     else " — solide" if g >= 10 else " — modeste, ne justifie pas à elle seule un multiple élevé"),
                     1 if g >= 20 else (0 if g >= 10 else -1)))
        pe_ref = pe_tg or pe_self
        if pe_ref and g > 0:
            peg = pe_ref / g
            rows.append(("PEG (PER / croissance)", f"{peg:.2f}".replace(".", ","),
                         f"PER {'de l' + chr(39) + 'objectif' if pe_tg else 'actuel'} {pe_ref:.1f}x ".replace(".", ",")
                         + ("÷ croissance : ≤ 1,2 = multiple payé par la croissance" if peg <= 1.2 else
                            "÷ croissance : > 2 = multiple cher même rapporté à la croissance" if peg > 2 else
                            "÷ croissance : zone neutre"), 1 if peg <= 1.2 else (-1 if peg > 2 else 0)))
    if N and L and ebit.get(L) and ebit.get(N + 1) and ebit[L] > 0 and ebit[N + 1] > 0 and N + 1 > L:
        ge = ((ebit[N + 1] / ebit[L]) ** (1 / (N + 1 - L)) - 1) * 100
        lbl_e = f"EBIT {L} → {N + 1}"
        if g is not None and ge - g >= 15:
            lbl_e += (" — bien supérieure à celle du BNA : levier opérationnel, ou BNA de départ gonflé par des "
                      "éléments exceptionnels")
        rows.append(("Croissance de l'EBIT", f"{ge:+.0f} %/an", lbl_e, 1 if ge >= 20 else (0 if ge >= 8 else -1)))
    else:
        ge = None
    # 2. rentabilité
    roe = (z.get("roe") or {}).get(N + 1) if N else None
    roe = roe if roe is not None else f.get("roe_ttm")
    if roe is not None:
        rows.append(("Rentabilité des fonds propres (ROE)", pc(roe), ("N+1 " if N else "12 derniers mois ") +
                     ("— élevée : entreprise de qualité" if roe >= 20 else "— correcte" if roe >= 10 else "— faible"),
                     1 if roe >= 20 else (0 if roe >= 10 else -1)))
    # 3. bilan
    nd = (z.get("nd") or {}).get(N) if N else None
    nd = nd if nd is not None else f.get("net_debt")
    eb = (z.get("ebitda") or {}).get(N) if N else None
    eb = eb or f.get("ebitda")
    if nd is not None and eb and eb > 0:
        lv = nd / eb
        rows.append(("Dette nette / EBITDA", "trésorerie nette" if nd < 0 else f"{lv:.1f}x".replace(".", ","),
                     "bilan solide : rachats d'actions et dividendes possibles" if lv <= 1.5 else
                     "endettement modéré" if lv <= 3 else "endettement élevé : fragilise la prime",
                     1 if lv <= 1.5 else (0 if lv <= 3 else -1)))
    # 4. conversion du bénéfice en trésorerie
    conv = None
    if N and L:
        yrs = [y for y in range(L - 1, N + 2) if rn.get(y) and rn[y] > 0 and fcf.get(y) is not None]
        if len(yrs) >= 3:                              # médiane annuelle : une année atypique (acquisition…) ne fausse pas
            conv = _med([fcf[y] / rn[y] for y in yrs])
    elif f.get("fcf_norm") and f.get("eps_ttm") and f["eps_ttm"] > 0 and f.get("shares"):
        conv = f["fcf_norm"] / (f["eps_ttm"] * f["shares"])
    if conv is not None and -1 <= conv <= 3:
        rows.append(("Conversion du bénéfice en trésorerie", pc(conv * 100),
                     "flux libre ≥ bénéfice : bénéfices de qualité" if conv >= 0.9 else
                     "flux libre nettement inférieur au bénéfice" if conv < 0.6 else "conversion correcte",
                     1 if conv >= 0.9 else (-1 if conv < 0.6 else 0)))
    # 4 bis. intensité capitalistique (part de l'EBITDA réinvestie)
    cx = z.get("capex") or {}
    if N:
        yc = [y for y in range(N, N + 2) if cx.get(y) and (z.get("ebitda") or {}).get(y, 0) > 0]
        if yc:
            ci = _med([cx[y] / z["ebitda"][y] for y in yc])
            rows.append(("Investissements / EBITDA", pc(ci * 100),
                         "peu d'investissements : la trésorerie revient aux actionnaires" if ci <= 0.3 else
                         "modèle très capitalistique : l'essentiel de l'EBITDA est réinvesti" if ci > 0.6 else
                         "investissements modérés", 1 if ci <= 0.3 else (-1 if ci > 0.6 else 0)))
    # 4 ter. part du prix payée pour la croissance (EPV) et dilution
    ep = epv_value(f, fv)
    gp = (ep.get("g_target") if ep.get("g_target") is not None else ep.get("g_price")) if ep and ep["ps"] > 0 else None
    if gp is not None:
        ref_ = "l'objectif" if ep.get("g_target") is not None else "le cours"
        rows.append(("Part du prix payée pour la croissance", pc(gp),
                     f"{ref_} vs valeur sans croissance (EPV {ep['ps']:.0f}) : "
                     + ("largement couverte par les bénéfices actuels" if gp <= 40 else
                        "repose surtout sur la croissance future" if gp > 65 else "équilibre croissance / bénéfices actuels"),
                     1 if gp <= 40 else (-1 if gp > 65 else 0)))
    dl = net_dilution(f)
    if dl:
        v_, t_, sg_ = dilution_txt(dl, "")
        rows.append(("Nombre d'actions", v_, t_.replace(" M/an", " M/an"), sg_))
    # 5. historique / cycle
    if L:
        losses = [y for y in range(L - 4, L + 1) if bna.get(y) is not None and bna[y] < 0]
        if losses:
            rows.append(("Historique", "pertes en " + ", ".join(map(str, losses)),
                         "redressement récent : la prime parie sur des bénéfices durables, sans long historique", -1))
    if f.get("cyclical"):
        rows.append(("Cycle", "valeur cyclique", "BNA N+1 possiblement proche d'un pic : prime fragile", -1))
    if not rows:
        return None
    pos, neg = sum(1 for r in rows if r[3] > 0), sum(1 for r in rows if r[3] < 0)
    risks = [r[0] if r[0][1:2].isupper() else r[0][0].lower() + r[0][1:] for r in rows if r[3] < 0]
    gv = max(x for x in (g, ge) if x is not None) if (g is not None or ge is not None) else None
    losses_ = any(r[0] == "Historique" for r in rows)
    if f.get("cyclical") and gv is not None and gv >= 15:
        verdict = (f"Valeur cyclique : la croissance de {gv:.0f} %/an reflète la phase haute du cycle. L'objectif "
                   "capitalise ces bénéfices, la juste valeur un bénéfice de milieu de cycle : prime fragile si le cycle "
                   "se retourne.")
    elif gv is not None and gv >= 15 and pos > neg:
        src_ = "du BNA" if g is not None and gv == g else "de l'EBIT"
        g = gv
        verdict = (f"Prime cohérente avec les chiffres : elle est payée par une croissance {src_} de {g:.0f} %/an"
                   + (" et un profil financier solide" if pos >= 3 else "")
                   + ". Elle suppose que les prévisions se réalisent" + (f" ; points de fragilité : {', '.join(risks)}." if risks else "."))
    elif gv is None or gv < 10:
        verdict = ("Prime peu expliquée par la croissance chiffrée jusqu'en N+1 : elle repose plutôt sur "
                   + ("un redressement attendu au-delà de N+1, " if losses_ else "")
                   + "une qualité perçue, un thème porteur ou l'optimisme des analystes — à vérifier.")
    else:
        verdict = ("Prime en partie justifiée : " + (f"{pos} critère(s) favorable(s), " if pos else "")
                   + (f"points faibles : {', '.join(risks)}." if risks else "aucun point faible net."))
    return {"prem_tg": prem_tg, "prem_peer": prem_peer, "pe_tg": pe_tg, "pe_jv": pe_jv, "pe_self": pe_self,
            "pe_peers": pe_peers, "n_peers": len(peers), "rows": rows, "verdict": verdict}


def premium_intro(pr):
    """Phrase d'ouverture du bloc « Pourquoi cette prime ? »."""
    x_ = lambda v: f"{v:.1f}".replace(".", ",") + "x"
    parts = []
    if pr["prem_tg"] is not None and pr["prem_tg"] >= 10:
        parts.append(f"l'objectif des analystes correspond à {x_(pr['pe_tg'])} le BNA N+1 contre {x_(pr['pe_jv'])} pour "
                     f"la juste valeur ({pr['prem_tg']:+.0f} %)")
    if pr["prem_peer"] is not None and pr["prem_peer"] >= 25:
        parts.append(f"la valeur cote {x_(pr['pe_self'])} les bénéfices contre {x_(pr['pe_peers'])} en médiane pour "
                     f"ses {pr['n_peers']} pairs ({pr['prem_peer']:+.0f} %)")
    return ("Prime constatée : " + " ; ".join(parts) + ". Critères qui justifient d'ordinaire un multiple élevé "
            "(✅ justifie · ➖ neutre · ⚠️ fragilise) :")


ICON_P = {1: "✅", 0: "➖", -1: "⚠️"}
PREMIUM_NOTE = ("Non mesuré ici : carnet de commandes, thème porteur (IA, réseaux électriques…), révisions récentes du "
                "consensus et position concurrentielle — à lire dans les publications de la société.")


def gap_causes(f, fv, lev=None):
    """Raisons probables de l'écart juste valeur / objectif des analystes (fiche à l'écran et export PDF / Word)."""
    causes = []
    why = _divergence_causes(f, fv)
    if why:
        causes.append("écart entre méthodes : " + why)
    wv = wacc_effect(f, fv)
    if wv:
        causes.append(wv)
    g5, cap_ = growth_expected(f)
    if cap_:
        causes.append(f"croissance du BNA plafonnée à {g5:.0f} % (" + (
            "consensus en % non significatif : BNA de départ quasi nul" if near_zero_base(f) else
            f"consensus {f.get('eps_5y') or 0:.0f} %") + ") ; les analystes retiennent souvent la croissance du "
            "consensus et le BNA N+2")
    if lev is not None and lev > 3:
        causes.append(f"endettement : dette nette {lev:.1f}× l'EBITDA déduite de la valeur")
    if f.get("cyclical"):
        causes.append(("classement cyclique : PER cible 11x sur le BNA de milieu de cycle" if fv.get("_eps_mid") else
                       "classement cyclique : PER cible 8x sur le BNA N+1") + ", croissance ≤ 5 % (correction possible dans "
                      "« Forcer le classement cyclique »)")
    return causes


def divergence_txt(f, fv):
    """Lecture de l'écart entre méthodes (précision 🟠 / 🔴) : causes probables, chiffrées, les plus fortes d'abord."""
    alerts = " · ".join(fv.get("_alerts") or [])
    why = _divergence_causes(f, fv)
    return " · ".join(x for x in (alerts, why) if x)


def _divergence_causes(f, fv):
    d = fv.get("Dispersion (%)")
    if not fv.get("JV centrale") or (d is not None and d <= 15 and not fv.get("_excl") and not fv.get("_peak_only")):
        return ""
    res = fv.get("_dcfx")
    dcf = fv.get("DCF")
    mult = fv.get("Multiples (médiane)") or _med([fv.get(m) for m in ("PER", "Comparables", "PER hist.") if fv.get(m)])
    causes = []                                        # (poids, texte)
    if dcf and mult:
        low = dcf < mult
        if f.get("cyclical"):
            causes.append((5 if low else 2, "bénéfices de haut de cycle : les multiples s'appliquent à un BNA non durable"
                           if low else ("valeur cyclique : PER cible 11x sur BNA de milieu de cycle" if fv.get("_eps_mid") else "valeur cyclique : PER cible réduit (8x)") + ", le DCF intègre le consensus"))
        if res:
            x, w = res["x"], res["wacc"]
            rows = res["scen"][1]["_rows"]
            earn = (f.get("eps_next_y") or 0) * (f.get("shares") or 0)
            conv = rows[0]["FCFF"] / earn if earn > 0 else None
            if low:
                if x["capex0"] > max(1.3 * x["da_pct"], 0.06):
                    causes.append((4, f"investissements lourds (capex {x['capex0'] * 100:.0f} % du CA, "
                                      f"{x['capex0'] / max(x['da_pct'], 1e-6):.1f}× les amortissements) : la trésorerie "
                                      f"est absorbée"))
                if conv is not None and conv < 0.7:
                    causes.append((3.5, f"faible conversion du bénéfice en trésorerie (flux an 1 = {conv * 100:.0f} % "
                                        f"du résultat attendu)"))
                if w["wacc"] >= 0.105:
                    causes.append((3, f"taux d'actualisation élevé (CMPC {w['wacc'] * 100:.1f} %, "
                                      f"{w['how'].split('(')[0].strip()})"))
                gi, ga = fv.get("g implicite (%)"), fv.get("g attendue (%)")
                if gi is not None and ga is not None and gi - ga >= 8:
                    causes.append((3.2, f"le cours anticipe une croissance du CA de {gi:.0f} %/an sur 5 ans (retenue : {ga:.0f} %)"))
                eb = x["rev0"] * x["m0"]
                if eb > 0 and x["net_debt"] / eb > 3:
                    causes.append((2.5, f"dette nette élevée ({x['net_debt'] / eb:.1f}× l'EBITDA) déduite de la valeur"))
                gs, g5 = f.get("sales_g"), f.get("eps_5y")
                if gs is not None and g5 is not None and g5 > gs + 15:
                    causes.append((2, f"BNA attendu en forte hausse (+{g5:.0f} %/an) sans croissance comparable "
                                      f"des ventes (+{gs:.0f} %)"))
            else:
                mc = f.get("mcap") or 0
                if mc > 0 and x["net_debt"] < 0 and -x["net_debt"] / mc > 0.25:
                    causes.append((4, f"trésorerie nette importante ({-x['net_debt'] / mc * 100:.0f} % de la "
                                      f"capitalisation) ajoutée au DCF — le marché la décote (avances clients, "
                                      f"risque, holding…)"))
                if conv is not None and conv > 1.3:
                    causes.append((3, f"flux de trésorerie supérieurs aux bénéfices ({conv * 100:.0f} % du résultat : "
                                      f"amortissements élevés, BNA déprimé)"))
                if x.get("fc") and x["fc"]["g"] and max(x["fc"]["g"]) > 0.15:
                    causes.append((2.5, f"forte croissance du CA prévue par le consensus "
                                        f"(+{max(x['fc']['g']) * 100:.0f} %) intégrée au DCF"))
                gi, ga = fv.get("g implicite (%)"), fv.get("g attendue (%)")
                if gi is not None and ga is not None and ga - gi >= 8:
                    causes.append((3, f"le cours n'anticipe que {gi:+.0f} %/an de croissance du CA sur 5 ans "
                                      f"(retenue : {ga:+.0f} %)"))
                if res["wacc"]["wacc"] <= 0.085:
                    causes.append((1.5, f"taux d'actualisation bas (CMPC {res['wacc']['wacc'] * 100:.1f} %)"))
        else:                                          # fiche seule : DCF simplifié (flux libre des 12 derniers mois)
            earn = (f.get("eps_ttm") or 0) * (f.get("shares") or 0)
            fcf0 = (f.get("fcf_hist") or [None])[0]
            conv = fcf0 / earn if fcf0 and earn > 0 else None
            gi, ga = fv.get("g implicite (%)"), fv.get("g attendue (%)")
            ke = fv.get("Coût fonds propres (%)")
            if low:
                if conv is not None and conv < 0.7:
                    causes.append((3.5, f"faible conversion du bénéfice en trésorerie (flux libre = {conv * 100:.0f} % "
                                        f"du résultat)"))
                if gi is not None and ga is not None and gi - ga >= 8:
                    causes.append((3.2, f"le cours anticipe une croissance du flux de {gi:.0f} % (attendue : {ga:.0f} %)"))
                if ke and ke >= 11:
                    causes.append((3, f"taux d'actualisation élevé ({ke:.1f} %)"))
            elif conv is not None and conv > 1.3:
                causes.append((3, "flux libre très supérieur au bénéfice, quasi nul" if conv > 10 else
                               f"flux libre supérieur aux bénéfices ({conv * 100:.0f} % du résultat)"))
            causes.append((0.5, "DCF simplifié (fiche seule) : lancez l'analyse complète pour le DCF d'entreprise"))
    if fv.get("_excl") and not causes:
        causes.append((1, "méthode(s) " + ", ".join(fv["_excl"]) + " trop éloignée(s) de la médiane"))
    if not causes:
        return ""
    sens = ("DCF < multiples" if dcf and mult and dcf < mult else "DCF > multiples" if dcf and mult else "écart")
    return f"{sens} : " + " ; ".join(t for _, t in sorted(causes, key=lambda c: -c[0])[:2])


def precision_txt(fv):
    d = fv.get("Dispersion (%)")
    excl = fv.get("_excl") or []
    tail = (" · " + ", ".join(excl) + (" écarté" if len(excl) == 1 else " écartés")) if excl else ""
    if fv.get("_peak_only"):
        return "🔴 multiples du seul BNA de pic" + tail
    if fv.get("_mode") == "dcf" and d is not None:
        return ("🟢" if d <= 15 else "🟠" if d <= 35 else "🔴") + f" ±{d:.0f} % (DCF / multiples)" + tail
    if d is None:
        if excl:                                     # méthodes en désaccord : une seule retenue
            return "🔴 1 seule méthode retenue" + tail
        return ("n.d. (1 méthode)" if fv.get("JV centrale") else "n.d.") + tail
    return ("🟢" if d <= 15 else "🟠" if d <= 35 else "🔴") + f" ±{d:.0f} %" + tail


# ============================================================
#  ANALYSTES, SUPER INVESTORS, OPTIONS
# ============================================================
def analyst_stats(f, days=180):
    rec = []
    for r in f["ratings"]:
        try:
            d = datetime.strptime(r[0], "%b-%d-%y")
        except (ValueError, IndexError):
            continue
        if (TODAY - d).days <= days:
            rec.append({"Date": d.date(), "Action": r[1], "Courtier": r[2], "Recommandation": r[3],
                        "Objectif": r[4] if len(r) > 4 else ""})
    up = sum(1 for r in rec if r["Action"] == "Upgrade")
    down = sum(1 for r in rec if r["Action"] == "Downgrade")
    pot = (f["target"] / f["price"] - 1) * 100 if f["target"] and f["price"] else None
    lab = f.get("consensus_label")
    if f["recom"]:
        lab = ("Achat fort" if f["recom"] < 1.5 else "Achat" if f["recom"] < 2.5 else
               "Conserver" if f["recom"] < 3.5 else "Vente")
    return {"Recom (1=achat fort, 5=vente)": f["recom"], "Consensus": lab, "Objectif moyen": f["target"],
            "Potentiel (%)": pot, "Relèvements 180j": up, "Abaissements 180j": down, "Actions récentes": rec}


SELL_WORDS = ("Reduce", "Sell")


def superinvestor_stats(d, price=None):
    """Super investors (Dataroma, 13F du dernier trimestre) : nouveaux entrants, renforcements, allègements,
    sorties, pondérés par la conviction (part du portefeuille du gérant), et cours vs prix de détention moyen."""
    rows = d.get("rows", [])
    mv = lambda r: r["Mouvement"] or "="
    pct_of = lambda r: num(re.sub(r"^\D+", "", mv(r))) or 0
    new = [r for r in rows if mv(r).startswith("Buy")]
    add = [r for r in rows if mv(r).startswith("Add")]
    red = [r for r in rows if mv(r).startswith("Reduce")]
    out_ = [r for r in rows if mv(r).startswith("Sell")]
    conv = lambda r: (r.get("% portefeuille") or 0) >= 5
    hold = d.get("hold")
    res = {"Super investors": d.get("n") or len(rows), "Nouveaux": len(new), "Renforcent": len(add),
           "Allègent": len(red), "Sortent": len(out_),
           "Achats/renforts": len(new) + len(add), "Allègements": len(red) + len(out_),
           "Nouveaux entrants": ", ".join(r["Gérant"] for r in new) or "—",
           "Prix de détention moyen": hold,
           "Cours vs prix moyen (%)": round((price / hold - 1) * 100) if price and hold else None, "Lignes": rows}
    up = len(new) + len(add)
    down = len(red) + len(out_)
    strong = [r for r in new + add if conv(r)]
    if not rows:
        sig = "⚪"
    elif up > down and (strong or len(new) >= 2):
        sig = "🟢🟢"
    elif up > down:
        sig = "🟢"
    elif down > up:
        sig = "🟠"
    else:
        sig = "⚪"
    key = sorted(new + add, key=lambda r: -(r.get("% portefeuille") or 0))[:2]
    short = lambda n: n.split()[-1] if len(n.split()) == 2 else n
    txt = ", ".join(f"{short(r['Gérant'])} {r.get('% portefeuille') or 0:.1f} %"
                    + (" (nouveau)" if r in new else f" (+{pct_of(r):.0f} %)") for r in key)
    import signaux as sg
    conv = sg.conviction_signal(rows)                 # signal pondéré par la conviction (poids des lignes)
    res["Conviction"] = conv
    sig = conv["dot"] if rows else sig
    res["Clés"] = txt
    res["Signal"] = sig
    res["Résumé"] = (f"{conv['txt']} · {res['Super investors']} fonds · {len(new)} nouv. {len(add)} renf. "
                     f"{len(red) + len(out_)} allèg.") if rows else ""
    return res


def put_call_stats(opts):
    if not opts:
        return {}
    s = lambda k: sum(o[k] for o in opts)
    r = lambda a, b: round(s(a) / s(b), 2) if s(b) else None
    pc, near = r("OI puts", "OI calls"), r("OI puts ±10 %", "OI calls ±10 %")
    sig = "—"
    if pc is not None:
        sig = "prudent / couverture" if pc > 1.0 else "haussier" if pc < 0.7 else "neutre"
    return {"P/C encours": pc, "P/C encours ±10 %": near, "P/C volumes": r("Vol puts", "Vol calls"),
            "Échéances": len(opts), "Lecture P/C": sig}


# ============================================================
#  SCÉNARIOS : BNA, EBITDA, CAPEX, DETTE NETTE (2 exercices)
# ============================================================
def scenarios(f, ov=None):
    """ov (facultatif) : dict avec 'g EBITDA N+1 (%)', 'Capex N+1 (Md$)', 'Capex N+2 (Md$)'."""
    ov = ov or {}
    if f.get("zb"):
        return scenarios_consensus(f)
    if not f["ebitda"] or not f["rev"]:
        return pd.DataFrame()
    cyc = f["cyclical"]
    d = 0.25 if cyc else 0.10                                     # dispersion du BNA autour du consensus
    g5 = f["eps_5y"] if f["eps_5y"] is not None else 8.0
    gy = f["eps_next_y_pct"] if f["eps_next_y_pct"] is not None else g5
    g1 = ov.get("g EBITDA N+1 (%)")
    g1 = g1 if g1 is not None and pd.notna(g1) else clamp(f["sales_yoy"] if f["sales_yoy"] is not None else gy, -20, 50)
    g2 = 0.0 if cyc else clamp(min(g1, g5), -10, 25)              # cyclique : plateau en N+2
    intensity = f["capex"] / f["rev"] if f["capex"] is not None else None
    fmargin = f["fcf_hist"][0] / f["ebitda"] if f["fcf_hist"] else 0.5   # si capex inconnu
    conv = clamp((f["cfo"] or f["ebitda"]) / f["ebitda"], 0.4, 1.1)  # conversion EBITDA -> flux d'exploitation
    returns = f["buyback"] + f["div_ps"] * f["shares"]               # rachats + dividendes (rythme actuel)
    eps1 = f["eps_next_y"]
    eps2 = eps1 * (1 + (0 if cyc else clamp(g5, -10, 30)) / 100) if eps1 else None
    rows = []
    for name, dg, de, kc in (("Pessimiste", -25 if cyc else -10, -1, 1.2),
                             ("Central", 0, 0, 1.0),
                             ("Optimiste", 15 if cyc else 8, 1, 0.9)):
        nd, rev, eb = f["net_debt"] or 0.0, f["rev"], f["ebitda"]
        for i, (g, n) in enumerate(((g1, "N+1"), (g2, "N+2"))):
            gg = (g + dg) / 100
            rev, eb = rev * (1 + gg), eb * (1 + gg)
            cap = ov.get(f"Capex {n} (Md$)")
            cap = cap * 1e3 if cap is not None and pd.notna(cap) else (rev * intensity * kc if intensity is not None else None)
            fcf = eb * conv - cap if cap is not None else eb * fmargin
            nd = nd - fcf + returns
            eps = (eps1 if i == 0 else eps2)
            eps = eps * (1 + de * d * (1 if i == 0 else 1.5)) if eps else None
            rows.append({"Scénario": name, "Exercice": n, "BNA": eps, "EBITDA (Md$)": eb / 1e3,
                         "Capex (Md$)": cap / 1e3 if cap is not None else None, "Flux libre (Md$)": fcf / 1e3, "Dette nette (Md$)": nd / 1e3,
                         "Dette nette / EBITDA": nd / eb if eb > 0 else None})
    return pd.DataFrame(rows)


def scenarios_consensus(f):
    """Zonebourse : scénario central = consensus des analystes (N+1, N+2), pessimiste / optimiste autour."""
    z, N = f["zb"], f["zb"]["N"]
    d = 0.25 if f["cyclical"] else 0.10
    if f.get("bank"):
        return scenarios_bank(f, d)
    nd_c = f["net_debt"] or 0.0
    gap, rows = {"Pessimiste": 0.0, "Central": 0.0, "Optimiste": 0.0}, []
    for i, y in enumerate((N + 1, N + 2)):
        eb, cap, fcf, eps = z["ebitda"].get(y), z["capex"].get(y), z["fcf"].get(y), z["bna"].get(y)
        cap = abs(cap) if cap is not None else None
        if z["nd"].get(y) is not None:
            nd_c = z["nd"][y]
        elif fcf is not None:
            nd_c = nd_c - fcf + (z["div"].get(y) or 0) * f["shares"]
        for name, k in (("Pessimiste", -1), ("Central", 0), ("Optimiste", 1)):
            s_ = 1 + k * d * (1 if i == 0 else 1.5)
            eb_s = eb * s_ if eb is not None else None
            cap_s = cap * (1.1 if k < 0 else 0.95 if k > 0 else 1.0) if cap is not None else None
            fcf_s = None
            if fcf is not None:
                fcf_s = fcf + ((eb_s - eb) * 0.7 if eb is not None else fcf * (s_ - 1)) - ((cap_s or 0) - (cap or 0))
                gap[name] += fcf - fcf_s
            nd_s = nd_c + gap[name]
            rows.append({"Scénario": name, "Exercice": f"N+{i + 1} ({y})", "BNA": eps * s_ if eps else None,
                         "EBITDA (Md)": eb_s / 1e3 if eb_s is not None else None,
                         "Capex (Md)": cap_s / 1e3 if cap_s is not None else None,
                         "Flux libre (Md)": fcf_s / 1e3 if fcf_s is not None else None,
                         "Dette nette (Md)": nd_s / 1e3,
                         "Dette nette / EBITDA": nd_s / eb_s if eb_s and eb_s > 0 else None})
    order = {"Pessimiste": 0, "Central": 1, "Optimiste": 2}
    return pd.DataFrame(rows).sort_values(["Scénario", "Exercice"], key=lambda c: c.map(order) if c.name == "Scénario" else c)


def scenarios_bank(f, d):
    """Banque : BNA, résultat net, dividende, actif net par action et ROE (consensus ± dispersion)."""
    z, N, rows = f["zb"], f["zb"]["N"], []
    for name, k in (("Pessimiste", -1), ("Central", 0), ("Optimiste", 1)):
        bv_gap = 0.0
        for i, y in enumerate((N + 1, N + 2)):
            s_ = 1 + k * d * (1 if i == 0 else 1.5)
            eps, div, bv = z["bna"].get(y), z["div"].get(y), z["bvps"].get(y)
            eps_s = eps * s_ if eps else None
            div_s = div * s_ if div else None                 # taux de distribution constant
            if eps_s is not None and eps:
                bv_gap += (eps_s - eps) - ((div_s or 0) - (div or 0))
            bv_s = bv + bv_gap if bv else None
            rn = z["rn"].get(y)
            rows.append({"Scénario": name, "Exercice": f"N+{i + 1} ({y})", "BNA": eps_s,
                         "Résultat net (Md)": rn * s_ / 1e3 if rn else None, "Dividende / action": div_s,
                         "Rendement (%)": div_s / f["price"] * 100 if div_s and f["price"] else None,
                         "Actif net / action": bv_s,
                         "ROE (%)": eps_s / bv_s * 100 if eps_s and bv_s else None,
                         "Cours / actif net": f["price"] / bv_s if bv_s and f["price"] else None})
    return pd.DataFrame(rows)


# ============================================================
#  SYNTHÈSE PAR VALEUR
# ============================================================
# ============================================================
#  SENTIMENT : titres d'actualité Finviz (VADER) + mentions Reddit (ApeWisdom)
# ============================================================
FIN_LEXICON = {"upgrade": 2.0, "upgraded": 2.0, "upgrades": 2.0, "downgrade": -2.0, "downgraded": -2.0,
               "downgrades": -2.0, "beat": 1.5, "beats": 1.5, "tops": 1.2, "miss": -1.5, "misses": -1.5,
               "missed": -1.5, "surge": 2.0, "surges": 2.0, "soar": 2.0, "soars": 2.0, "jump": 1.5, "jumps": 1.5,
               "rally": 1.5, "rallies": 1.5, "plunge": -2.5, "plunges": -2.5, "tumble": -2.0, "tumbles": -2.0,
               "slump": -2.0, "slumps": -2.0, "sinks": -1.8, "falls": -1.2, "drops": -1.2, "outperform": 1.5,
               "underperform": -1.5, "underperforms": -1.5, "outperforms": 1.5, "slides": -1.0, "gains": 1.0, "overweight": 1.2, "underweight": -1.2, "bullish": 2.0, "bearish": -2.0,
               "raises": 1.0, "raised": 1.0, "lowers": -1.0, "lowered": -1.0, "cut": -1.5, "cuts": -1.5,
               "lawsuit": -2.0, "probe": -1.5, "investigation": -1.5, "layoffs": -1.5, "record": 1.0,
               "buyback": 1.2, "dividend": 0.5, "guidance": 0.0, "volatile": -0.5}


def _vader():
    try:
        from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
    except ImportError:
        return None
    a = SentimentIntensityAnalyzer()
    a.lexicon.update(FIN_LEXICON)
    return a


ALIASES = {"GOOGL": ["Google", "Alphabet"], "GOOG": ["Google", "Alphabet"], "META": ["Meta", "Facebook"],
           "SKHY": ["SK Hynix", "Hynix"], "BRK-B": ["Berkshire"]}


def relevant(news, ticker, company):
    """Garde les titres qui citent la société (Finviz mêle des titres sur les concurrents) ; tous si trop peu."""
    if not ticker and not company:
        return news
    keys = [ticker] + ALIASES.get(ticker, [])
    if company:
        w = re.sub(r"\b(Inc|Corp|Corporation|Co|Ltd|plc|NV|N\.V\.|SE|SA|AG|Holdings|Group|Technologies|The)\b\.?", "",
                   company).split()
        if w:                                           # « NXP », « Uber »… ou « GE HealthCare » si le 1er mot est court
            keys.append(w[0] if len(w[0]) >= 3 or len(w) == 1 else " ".join(w[:2]))
    pat = re.compile(r"\b(" + "|".join(re.escape(k) for k in keys) + r")\b", re.I)
    sel = [n for n in news if pat.search(n[1])]
    return sel if len(sel) >= 5 else news


def news_sentiment(news, ref=None, ticker="", company=""):
    """Score de -100 (très négatif) à +100 (très positif) : moyenne des scores VADER des titres."""
    a = _vader()
    if a is None or not news:
        return None
    ref = ref or TODAY
    news = relevant(news, ticker, company)
    rows = []
    for n_ in news:
        iso, title, src = n_[:3]
        d = datetime.fromisoformat(iso)
        rows.append({"Date": d, "Titre": title, "Source": src,
                     "Score": round(a.polarity_scores(title)["compound"] * 100)})
    df = pd.DataFrame(rows)
    age = (ref - df["Date"]).dt.days
    out = {"détail": df.sort_values("Date", ascending=False)}
    for n in (7, 30):
        sub = df[age <= n]
        out[f"n{n}"] = len(sub)
        out[f"s{n}"] = round(sub["Score"].mean()) if len(sub) else None
        out[f"pos{n}"] = int((sub["Score"] >= 15).sum())
        out[f"neg{n}"] = int((sub["Score"] <= -15).sum())
    return out


@st.cache_data(ttl=CACHE_HOURS * 3600, show_spinner=False)
def fetch_apewisdom(max_pages=10):
    """Classement des actions citées sur Reddit sur 24 h (ApeWisdom, gratuit, sans clé ; ~600 actions)."""
    import requests
    out, pg, pages = {}, 1, 1
    while pg <= min(pages, max_pages):
        r = requests.get(f"https://apewisdom.io/api/v1.0/filter/all-stocks/page/{pg}", headers=UA, timeout=8)
        r.raise_for_status()
        js = r.json()
        pages = js.get("pages") or 1
        for x in js.get("results", []):
            out[str(x.get("ticker", "")).upper()] = x
        pg += 1
        time.sleep(0.4)
    return out


def score_txt(v, n):
    """Score affiché avec une pastille : 🟢 positif (≥ 15), 🔴 négatif (≤ -15), ⚪ neutre ; vide sans titre."""
    if v is None or not n:
        return ""
    dot = "🟢" if v >= 15 else "🔴" if v <= -15 else "⚪"
    return f"{dot} {v:+d}".replace("+0", "0") + (" (1 titre)" if n == 1 else f" ({n} titres)")


def sentiment_table(facts):
    """facts : {ticker: faits} -> (tableau de synthèse, détails par ticker, message Reddit)."""
    try:
        reddit, msg = fetch_apewisdom(), None
    except Exception as e:
        reddit, msg = {}, f"Reddit (ApeWisdom) indisponible : {e}"
    if _vader() is None:
        msg = (msg + " · " if msg else "") + "installez vaderSentiment (pip install vaderSentiment) pour noter les actualités"
    rows, det = [], {}
    for t, f in facts.items():
        ns = news_sentiment(f.get("news") or [], ticker=t, company=f.get("company", ""))
        det[t] = ns
        x = reddit.get(t)
        if ns is None:
            lect, titres, tend = "n.d.", "n.d.", "n.d."
        else:
            n30, n7 = ns["n30"], ns["n7"]
            neu = n30 - ns["pos30"] - ns["neg30"]
            titres = f"{n30} : {ns['pos30']} positif(s), {ns['neg30']} négatif(s), {neu} neutre(s)" if n30 else "aucun titre"
            lect = ("aucun titre" if not n30 else "trop peu de titres" if n30 < 3 else
                    "positif" if ns["s30"] >= 15 else "négatif" if ns["s30"] <= -15 else "neutre")
            if n7 == 0:
                tend = "aucun titre sur 7 j"
            elif n7 < 3 or n30 < 3:
                tend = f"trop peu de titres ({n7} sur 7 j)"
            else:
                diff = ns["s7"] - ns["s30"]
                tend = "▲ s'améliore" if diff >= 10 else "▼ se dégrade" if diff <= -10 else "= stable"
        if x:
            m, m0 = int(x.get("mentions") or 0), int(x.get("mentions_24h_ago") or 0)
            ment = f"{m} (veille : {m0})"
            evol = (f"{(m / m0 - 1) * 100:+.0f} %" if m0 >= 10 else "faible volume") if m0 else "nouveau"
        else:
            ment, evol = "", "hors classement"
        rows.append({"Ticker": t,
                     "Actualités 7 j": score_txt(ns["s7"], ns["n7"]) if ns else "",
                     "Actualités 30 j": score_txt(ns["s30"], ns["n30"]) if ns else "",
                     "Lecture 30 j": lect,
                     "Titres 30 j": titres,
                     "Tendance (7 j vs 30 j)": tend,
                     "Reddit : rang": x.get("rank") if x else None,
                     })
    return pd.DataFrame(rows), det, msg


# ---------- Indice CNN Fear & Greed (onglet « Fear & Greed »)
CNN_FG_URL = "https://production.dataviz.cnn.io/index/fearandgreed/graphdata"
FG_LABELS = {"extreme fear": "Peur extrême", "fear": "Peur", "neutral": "Neutre", "greed": "Avidité",
             "extreme greed": "Avidité extrême"}
FG_COMPONENTS = [
    ("stock_price_strength", "Stock Price Strength", "#8e44ad", "Écart plus hauts / plus bas sur 52 semaines",
     "Compare le nombre d'actions du NYSE à un plus haut sur 52 semaines à celles à un plus bas."),
    ("stock_price_breadth", "Stock Price Breadth", "#16a085", "McClellan Volume Summation Index",
     "Compare le volume des actions en hausse à celui des actions en baisse (largeur du marché)."),
    ("put_call_options", "Put/Call (moyenne 5 jours)", "#d35400", "Ratio put/call",
     "Volume d'options de vente / options d'achat : un ratio élevé traduit la nervosité des investisseurs."),
]


@st.cache_data(ttl=600, show_spinner=False)
def fetch_fear_greed():
    """Indice CNN Fear & Greed : niveau actuel, historique 1 an (avec le S&P 500), composantes."""
    import requests
    r = requests.get(CNN_FG_URL, timeout=15, headers={
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/124.0.0.0 Safari/537.36"})
    r.raise_for_status()
    js = r.json()
    cur = js["fear_and_greed"]
    to_df = lambda rows, col: (pd.DataFrame(rows).assign(
        date=lambda d_: pd.to_datetime(d_["x"], unit="ms", utc=True).dt.normalize()).rename(columns={"y": col}))
    fg = to_df(js["fear_and_greed_historical"]["data"], "score")[["date", "score", "rating"]]
    sp = to_df(js["market_momentum_sp500"]["data"], "sp500")[["date", "sp500"]]
    df = fg.merge(sp, on="date", how="inner").sort_values("date").drop_duplicates("date").reset_index(drop=True)
    comps = {}
    for k, *_ in FG_COMPONENTS:
        c = js.get(k) or {}
        if c.get("data"):
            comps[k] = {"score": c.get("score"), "rating": c.get("rating"),
                        "df": to_df(c["data"], "value")[["date", "value"]].sort_values("date")}
    return cur, df, comps


@st.cache_data(ttl=600, show_spinner=False)
def fear_greed_commentary(api_key, score, rating, prev_close, prev_w, prev_m, prev_y, trend, sp):
    """Commentaire de 3-4 phrases (API Claude, facultatif) ; gardé en mémoire tant que les données ne changent pas."""
    try:
        import anthropic
    except ImportError:
        return "Le paquet « anthropic » n'est pas installé (pip install anthropic)."
    chg = (sp[-1] / sp[0] - 1) * 100 if sp and sp[0] else 0
    prompt = (f"Tu es un analyste de marché concis. Données actuelles du CNN Fear & Greed Index et du S&P 500 :\n"
              f"- score actuel : {score:.1f}/100 ({rating}) ; clôture précédente {prev_close:.1f} ; il y a 1 semaine "
              f"{prev_w:.1f} ; il y a 1 mois {prev_m:.1f} ; il y a 1 an {prev_y:.1f}\n"
              f"- 21 dernières valeurs (jours ouvrés) : {list(trend)}\n"
              f"- S&P 500 il y a ~1 mois : {sp[0]:.0f} ; aujourd'hui : {sp[-1]:.0f} ({chg:+.1f} %)\n"
              "En 3-4 phrases courtes et factuelles en français, décris la tendance récente du sentiment de marché en "
              "la mettant en relation avec l'évolution du S&P 500 sur la période, sans conseil d'investissement.")
    msg = anthropic.Anthropic(api_key=api_key).messages.create(
        model="claude-sonnet-4-5", max_tokens=300, messages=[{"role": "user", "content": prompt}])
    return msg.content[0].text


@st.fragment      # un clic dans cet onglet ne relance que l'onglet, pas le portefeuille
def render_fear_greed():
    st.subheader("Indice CNN Fear & Greed — sentiment du marché américain")
    try:
        import plotly.graph_objects as go
        from plotly.subplots import make_subplots
    except ImportError:
        st.error("Bibliothèque « plotly » manquante : pip install -r requirements.txt")
        return
    c0, c1 = st.columns([1, 3])
    if c0.button("🔄 Rafraîchir", key="fg_refresh"):
        fetch_fear_greed.clear()
        st.session_state.pop("_fail_fg", None)
        st.session_state.pop("_fail_fg_msg", None)
    if time.time() - st.session_state.get("_fail_fg", 0) < 600:          # échec récent : pas de nouvel essai
        st.error(f"Indice CNN indisponible ({st.session_state.get('_fail_fg_msg', '')}) — nouvel essai dans quelques "
                 "minutes ou avec « 🔄 Rafraîchir ».")
        return
    try:
        with st.spinner("Indice CNN Fear & Greed…"):
            cur, df, comps = fetch_fear_greed()
    except Exception as e:
        st.session_state["_fail_fg"] = time.time()
        st.session_state["_fail_fg_msg"] = f"{type(e).__name__}"
        st.error(f"Indice CNN indisponible ({type(e).__name__} : {e}).")
        return
    try:
        ts = datetime.fromisoformat(str(cur["timestamp"]).replace("Z", "+00:00"))
        c1.caption(f"Données CNN Business du {ts:%d/%m/%Y à %H:%M} UTC · mises en cache 10 minutes.")
    except Exception:
        pass
    score, rating = float(cur["score"]), str(cur["rating"])
    g, m = st.columns([1.2, 1])
    with g:
        fig = go.Figure(go.Indicator(
            mode="gauge+number", value=score, number={"suffix": " / 100", "font": {"size": 40}},
            title={"text": FG_LABELS.get(rating, rating).upper(), "font": {"size": 22}},
            gauge={"axis": {"range": [0, 100], "tickwidth": 1}, "bar": {"color": "#262626", "thickness": 0.25},
                   "steps": [{"range": [0, 25], "color": "#c0392b"}, {"range": [25, 45], "color": "#e67e22"},
                             {"range": [45, 55], "color": "#b1b1b1"}, {"range": [55, 75], "color": "#27ae60"},
                             {"range": [75, 100], "color": "#145a32"}]}))
        fig.update_layout(height=320, margin=dict(l=30, r=30, t=60, b=10))
        st.plotly_chart(fig, width="stretch", key="fg_gauge")
    with m:
        st.markdown("**Historique récent**")
        a, b = st.columns(2)
        a.metric("Clôture précédente", f"{cur['previous_close']:.1f}")
        b.metric("Il y a 1 semaine", f"{cur['previous_1_week']:.1f}", f"{score - cur['previous_1_week']:+.1f}")
        a, b = st.columns(2)
        a.metric("Il y a 1 mois", f"{cur['previous_1_month']:.1f}", f"{score - cur['previous_1_month']:+.1f}")
        b.metric("Il y a 1 an", f"{cur['previous_1_year']:.1f}", f"{score - cur['previous_1_year']:+.1f}")
        st.caption("0–25 peur extrême · 25–45 peur · 45–55 neutre · 55–75 avidité · 75–100 avidité extrême. "
                   "Indicateur à contre-courant : une peur extrême accompagne souvent les points bas du marché, une "
                   "avidité extrême les excès — c'est un indicateur d'ambiance, pas un signal d'achat ou de vente.")
    with st.expander("🤖 Commentaire automatique (API Claude, facultatif)"):
        # le champ « mot de passe » n'est créé qu'à la demande : sa seule présence dans la page fait proposer au
        # navigateur ses mots de passe enregistrés (« Manage Passwords ») sur les autres champs, comme le menu « Valeur »
        key = None
        if st.checkbox("Activer le commentaire automatique", key="fg_ai_on"):
            key = st.text_input("Clé API Anthropic", type="password", key="fg_api_key", autocomplete="off",
                                help="Utilisée seulement pour ce commentaire, jamais enregistrée dans le script.")
        if key and len(df) >= 21:
            try:
                with st.spinner("Rédaction du commentaire…"):
                    st.info(fear_greed_commentary(
                        key, score, rating, cur["previous_close"], cur["previous_1_week"], cur["previous_1_month"],
                        cur["previous_1_year"], tuple(round(v, 1) for v in df.tail(21)["score"]),
                        tuple(round(v, 1) for v in df.tail(21)["sp500"])))
            except Exception as e:
                st.warning(f"Commentaire impossible : {e}")
    st.markdown("**S&P 500 et Fear & Greed (12 derniers mois)**")
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Scatter(x=df["date"], y=df["sp500"], mode="lines", line=dict(color="#1f2937", width=2),
                             name="S&P 500"), secondary_y=False)
    fig.add_trace(go.Scatter(x=df["date"], y=df["score"], mode="lines", name="Fear & Greed", opacity=0.85,
                             line=dict(color="#006598", width=1.5, dash="dot")), secondary_y=True)
    fig.update_yaxes(title_text="S&P 500 (points)", secondary_y=False)
    fig.update_yaxes(title_text="Fear & Greed (0–100)", range=[0, 100], showgrid=False, secondary_y=True)
    fig.update_layout(height=440, margin=dict(l=10, r=10, t=10, b=10), hovermode="x unified",
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    st.plotly_chart(fig, width="stretch", key="fg_hist")
    st.caption("Ligne pleine : S&P 500 (échelle de gauche) · pointillés : Fear & Greed (échelle de droite).")
    if comps:
        st.markdown("**Composantes de l'indice**")
        for k, lbl, color, ytitle, desc in FG_COMPONENTS:
            c = comps.get(k)
            if not c:
                continue
            a, b = st.columns([1, 3])
            a.metric(lbl, f"{c['score']:.1f} / 100" if c.get("score") is not None else "n.d.", help=desc)
            a.caption(FG_LABELS.get(c.get("rating"), c.get("rating") or "") + " — " + desc)
            f2 = go.Figure(go.Scatter(x=c["df"]["date"], y=c["df"]["value"], mode="lines",
                                      line=dict(color=color, width=2)))
            f2.update_layout(height=240, margin=dict(l=10, r=10, t=10, b=10), yaxis=dict(title=ytitle),
                             showlegend=False)
            b.plotly_chart(f2, width="stretch", key=f"fg_{k}")
        st.caption("Chaque composante a sa propre échelle (celle de CNN avant normalisation sur 0–100). L'indice en "
                   "compte sept : momentum du S&P 500, force et largeur des cours, put/call, volatilité (VIX), "
                   "demande de valeurs refuges et d'obligations à haut rendement.")
    with st.expander("Données brutes"):
        st.dataframe(df, width="stretch", hide_index=True)
        st.download_button("⬇️ CSV", df.to_csv(index=False).encode("utf-8"),
                           f"fear_greed_{TODAY:%Y%m%d}.csv", "text/csv", key="fg_csv")


def render_sentiment(facts):
    if not facts:
        st.info("Récupérez d'abord les données du portefeuille, ou analysez des pages Finviz dans « Valeurs à étudier ».")
        return
    df, det, msg = sentiment_table(facts)
    if msg:
        st.warning(msg)
    st.dataframe(df, hide_index=True, width="stretch", column_config={
        "Actualités 7 j": st.column_config.TextColumn(help="moyenne des scores des titres de la semaine, de -100 à +100"),
        "Actualités 30 j": st.column_config.TextColumn(help="moyenne des scores des titres du mois, de -100 à +100"),
        "Reddit : rang": st.column_config.NumberColumn(format="%d", help="place parmi les ~600 actions les plus citées sur Reddit sur 24 h (1 = la plus citée)"),
    })
    st.caption("Actualités : titres de la fiche Finviz qui citent la société, notés par VADER + vocabulaire financier, de -100 à +100 ; "
               "🟢 ≥ 15 positif, 🔴 ≤ -15 négatif, ⚪ neutre. Moins de 3 titres : pas de conclusion. "
               "Reddit (ApeWisdom) : rang parmi les actions les plus citées sur 24 h — l'intensité de la discussion, pas son ton. "
               "Valeurs Zonebourse : pas d'actualités en anglais, donc n.d.")
    t = st.selectbox("Titres d'une valeur", [x for x in det if det[x] is not None] or ["—"], key="sent_sel")
    if det.get(t):
        d = det[t]["détail"].copy()
        d["Date"] = d["Date"].dt.strftime("%d/%m %H:%M")
        d["Score"] = [score_txt(v, 1).replace(" (1 titre)", "") for v in d["Score"]]
        st.dataframe(d, hide_index=True, width="stretch")


# ---------- Reddit : actions les plus citées dont l'actualité est positive
NOT_STOCKS = {"SPY", "QQQ", "IWM", "DIA", "VOO", "VTI", "TQQQ", "SQQQ", "SPX", "SPXL", "SPXS", "SOXL", "SOXS", "UVXY",
              "VIX", "VXX", "TLT", "GLD", "SLV", "ARKK", "SMH", "XLF", "XLE", "XLK", "SCHD", "JEPI", "BITO", "IBIT",
              "TSLL", "NVDL", "MSTU", "CONL", "QQQM", "VUG", "VGT", "USO", "HYG", "SPXU", "UPRO", "LABU", "TNA"}


def fetch_quote(t, use_cache=True):
    """Fiche Finviz seule (pas d'états financiers ni d'options), cache 2 h."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    path = os.path.join(CACHE_DIR, f"{t}_quote.json")
    if use_cache and os.path.exists(path) and time.time() - os.path.getmtime(path) < CACHE_HOURS * 3600:
        with open(path, encoding="utf-8") as f:
            cached = json.load(f)
            if not _news_old(cached):
                return cached
    raw = {"ticker": t, "fetched": datetime.now().isoformat(timespec="minutes")}
    s = _session()
    raw.update(parse_quote(_get(s, f"https://finviz.com/quote.ashx?t={t}&p=d", pause=(0.6, 1.2))))
    add_peers(raw, t, s, log=lambda *_: None)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(raw, f)
    return raw


def reddit_scan(p, n_cand=40, n_top=10, use_cache=True, progress=None, log=print, min_s=15):
    """1) les n_cand actions les plus citées sur Reddit (ApeWisdom, 24 h), hors ETF ;
    2) ton des titres d'actualité Finviz de chacune (VADER, 30 j et 7 j) ;
    3) on garde les actualités nettement positives (≥ 15 sur 30 j, ≥ 5 titres, 7 j non négatif)
       et on classe par score = 60 % ton 30 j + 40 % ton 7 j, + bonus d'attention (mentions en hausse)."""
    ape = fetch_apewisdom()
    cand = [x for x in sorted(ape.values(), key=lambda x: int(x.get("rank") or 9999))
            if str(x.get("ticker", "")).upper() not in NOT_STOCKS][:n_cand]
    rows = []
    for i, x in enumerate(cand):
        t = str(x["ticker"]).upper()
        if progress:
            progress((i + 1) / len(cand), t)
        try:
            raw = fetch_quote(t, use_cache)
        except Exception as e:
            log(f"{t} : fiche Finviz indisponible ({type(e).__name__})")
            continue
        snap = raw.get("snap", {})
        if not snap.get("Market Cap") or "Asset Type" in snap or "Category" in snap:      # ETF, fonds, OTC
            continue
        try:
            ns = news_sentiment(raw.get("news") or [], ticker=t, company=raw.get("company", ""))
        except Exception:
            ns = None
        if not ns or ns["n30"] < 5 or ns["s30"] is None or ns["s30"] < min_s:
            continue
        s7 = ns["s7"] if ns["n7"] >= 3 and ns["s7"] is not None else ns["s30"]
        if s7 <= -15:
            continue
        m, m0 = int(x.get("mentions") or 0), int(x.get("mentions_24h_ago") or 0)
        mom = (m / m0 - 1) * 100 if m0 >= 10 else None
        bonus = 5 if mom is not None and mom >= 50 else 0
        score = round(0.6 * ns["s30"] + 0.4 * s7 + bonus)
        try:
            f = build_facts(raw)
            fv = fair_value(f, p)
            an = analyst_stats(f)
        except Exception as e:
            log(f"{t} : juste valeur impossible ({type(e).__name__} : {e})")
            f, fv, an = {"price": None, "target": None}, {}, {"Consensus": None, "Potentiel (%)": None}
        ana = f"{an['Consensus'] or '?'} · {f['target']:.0f}" if f.get("target") else (an.get("Consensus") or "n.d.")
        if an.get("Potentiel (%)") is not None:
            ana += f" ({an['Potentiel (%)']:+.0f} %)"
        rows.append({"Ticker": t, "Société": raw.get("company", "")[:40],
                     "Reddit : rang": int(x.get("rank") or 0),
                     "Mentions 24 h": f"{m}" + (f" ({mom:+.0f} %)" if mom is not None else " (nouveau)" if not m0 else ""),
                     "Actualités 30 j": score_txt(ns["s30"], ns["n30"]),
                     "Actualités 7 j": score_txt(ns["s7"], ns["n7"]),
                     "Score": score,
                     "Cours": f.get("price"),
                     "Juste valeur": (f"{fv['JV centrale']:.0f} ({fv['JV basse']:.0f}–{fv['JV haute']:.0f})"
                                      if fv.get("JV centrale") else "n.d."),
                     "Écart JV (%)": fv.get("Écart (%)"),
                     "Précision JV": precision_txt(fv) if fv.get("JV centrale") else "n.d.",
                     "Analystes · objectif": ana})
    df = pd.DataFrame(rows)
    if not df.empty:                    # toutes les valeurs retenues, triées ; le top n_top est pris à l'affichage
        df = df.sort_values(["Score", "Reddit : rang"], ascending=[False, True])
    return df, len(cand)


def send_to_watchlist(sel, label, max_days, use_cache, use_oi):
    """Envoie des tickers en analyse complète (fiche Finviz en ligne + états financiers…) dans « Valeurs à étudier »,
    gardés à part des pages HTML déposées (jamais écrasées)."""
    wl, errs = dict(st.session_state.get("wl_scr") or {}), list(st.session_state.get("wl_errs_scr") or [])
    bar = st.progress(0.0)
    for i, t in enumerate(sel):
        bar.progress((i + 1) / len(sel), text=t)
        try:
            f = build_facts(fetch_ticker(t, max_days, use_cache, log=errs.append, openinsider=use_oi))
            f["source"] = f"{label} + Finviz en ligne"
            wl[t] = f
        except FinvizBlocked as e:
            errs.append(f"⛔ {e} — envoi interrompu")
            break
        except Exception as e:
            errs.append(f"{t} : {type(e).__name__} — {e}")
    bar.empty()
    st.session_state["wl_scr"], st.session_state["wl_errs_scr"] = wl, errs
    st.success(f"{len(sel)} valeur(s) ajoutée(s) à l'onglet « 🛒 Valeurs à étudier ».")
    st.rerun()


@st.fragment      # un clic dans cet onglet ne relance que l'onglet, pas le portefeuille
def render_reddit_top(p, use_cache, max_days=200, use_oi=True):
    st.subheader("Reddit : 10 actions les plus citées à l'actualité positive")
    c1, c2 = st.columns([1, 3])
    n_cand = c1.slider("Actions examinées (rangs Reddit)", 20, 80, 40, 10, key="rd_n",
                       help="Une fiche Finviz est lue par action (≈ 1 s chacune, cache 2 h).")
    go_ = c2.button("🔎 Scanner Reddit", key="rd_go")
    if go_:
        logs = []
        bar = st.progress(0.0)
        try:
            df, n = reddit_scan(p, n_cand, 10, use_cache, progress=lambda v, t: bar.progress(v, text=t),
                                log=logs.append)
            st.session_state["rd_res"] = (df, n, datetime.now().strftime("%d/%m %H:%M"), logs)
        except Exception as e:
            st.session_state["rd_res"] = None
            st.error(f"Reddit (ApeWisdom) indisponible : {e}")
        bar.empty()
    res = st.session_state.get("rd_res")
    if not res:
        st.caption("Cliquez sur « Scanner Reddit » pour lancer la recherche.")
        return
    df, n, when, logs = res
    only_prec = st.checkbox("Seulement les JV précises (🟢 / 🟠)", value=False, key="rd_prec",
                            help="Méthodes concordantes à ± 35 % près : une forte marge avec des méthodes divergentes "
                                 "est souvent un piège. Le classement reprend les suivantes pour garder 10 valeurs.")
    if only_prec and not df.empty:
        df = df[df["Précision JV"].astype(str).str.startswith(("🟢", "🟠"))]
    df = df.head(10)
    if df.empty:
        st.info(f"Aucune des {n} actions examinées n'a une actualité nettement positive"
                + (" et une JV précise" if only_prec else "") + f" ({when}).")
    else:
        st.dataframe(df, hide_index=True, width="stretch", column_config={
            **FMT,
            "Reddit : rang": st.column_config.NumberColumn(format="%d", help="1 = action la plus citée sur Reddit sur 24 h"),
            "Mentions 24 h": st.column_config.TextColumn(help="nombre de mentions sur 24 h (évolution par rapport à la veille)"),
            "Score": st.column_config.NumberColumn(format="%d", help="60 % ton 30 j + 40 % ton 7 j, +5 si les mentions ont bondi de 50 % ou plus"),
            "Juste valeur": st.column_config.TextColumn(help="calculée sur la fiche Finviz seule (sans états financiers) : indicative"),
            "Précision JV": st.column_config.TextColumn(help="écart maximal des méthodes retenues à la JV : 🟢 ≤ 15 %, "
                                                              "🟠 ≤ 35 %, 🔴 au-delà (même lecture que le Screener JV)"),
        })
    st.caption(f"Scan du {when} sur les {n} actions les plus citées (hors ETF). Reddit mesure l'attention, pas le ton : "
               "le ton vient des titres d'actualité Finviz (≥ 5 titres sur 30 j, score ≥ 15, 7 j non négatif). "
               "Une action très discutée et bien notée peut déjà être chère : vérifiez la juste valeur.")
    if not df.empty:
        sel = st.multiselect("Envoyer en analyse complète (onglet « Valeurs à étudier »)", list(df["Ticker"]),
                             default=[], key="rd_sel")
        if st.button("➜ Analyse complète", key="rd_full", disabled=not sel):
            send_to_watchlist(sel, "Sentiment screener", max_days, use_cache, use_oi)
    if logs:
        with st.expander(f"{len(logs)} valeur(s) non lue(s) ou sans juste valeur"):
            st.write("\n".join(f"- {x}" for x in logs))


# ---------- Screener JV : liste d'un screener Finviz -> juste valeur rapide sur fiche seule
SCREENER_URL = ("https://finviz.com/screener/stocks?v=111&f=an_recom_buybetter%2Ccap_midover%2Cfa_debteq_u1%2C"
                "fa_estltgrowth_o10%2Cfa_fpe_u25%2Cfa_opermargin_o10%2Cfa_pfcf_u25%2Cfa_roi_o10%2Csh_avgvol_o200%2C"
                "ta_sma200_pa&ft=4&preset=s154743054")
CYCLICAL_INDUSTRIES = re.compile(r"Gold|Silver|Precious Metals|Copper|Aluminum|Steel|Coal|Uranium|Industrial Metals")


def parse_screener(html):
    """Page de résultats du screener Finviz (vue Overview) -> ([{ticker, company, sector, industry, country}], total)."""
    soup = BeautifulSoup(html, "html.parser")
    rows = []
    for r in soup.select("div.fv-table_row[data-row-key]"):                     # tableau actuel
        c = {x.get("data-column-id"): x.get_text(" ", strip=True) for x in r.select("[data-column-id]")}
        rows.append({"Ticker": r.get("data-row-key").upper(), "Société": c.get("company", ""),
                     "Secteur": c.get("sector", ""), "Industrie": c.get("industry", ""), "Pays": c.get("country", "")})
    if not rows:                                                                 # ancien tableau HTML
        for tr in soup.select("table tr"):
            tds = [td.get_text(" ", strip=True) for td in tr.find_all("td")]
            a = tr.find("a", href=re.compile(r"quote\.ashx\?t="))
            if a and len(tds) >= 6 and tds[0].isdigit():
                rows.append({"Ticker": tds[1].upper(), "Société": tds[2], "Secteur": tds[3], "Industrie": tds[4],
                             "Pays": tds[5]})
    m = re.search(r"/\s*(\d+)\s*Total", soup.get_text(" "))
    return rows, int(m.group(1)) if m else None


def fetch_screener(url, max_rows=500, log=print):
    """Toutes les pages d'un screener Finviz (20 lignes par page)."""
    base = re.sub(r"([?&])(v|r)=\d+&?", r"\1", url.strip()).rstrip("&?")
    base += ("&" if "?" in base else "?") + "v=111"
    s, out, seen, total, r = _session(), [], set(), None, 1
    while r <= max_rows:
        rows, tot = parse_screener(_get(s, f"{base}&r={r}", pause=(0.8, 1.5)))
        total = total or tot
        new = [x for x in rows if x["Ticker"] not in seen]
        if not new:
            break
        for x in new:
            seen.add(x["Ticker"])
            out.append(x)
        if total and len(out) >= total:
            break
        r += 20
    if total and total > len(out):
        log(f"{total} résultats : seuls les {len(out)} premiers sont lus — affinez le screener.")
    return out, total


def screen_values(tickers, p, use_cache=True, progress=None, log=print):
    rows = []
    for i, t in enumerate(tickers):
        if progress:
            progress((i + 1) / len(tickers), t)
        try:
            f = build_facts(fetch_quote(t, use_cache))
            fv = fair_value(f, p)
            an = analyst_stats(f)
        except Exception as e:
            log(f"{t} : {type(e).__name__} — {e}")
            continue
        ana = f"{an['Consensus'] or '?'} · {f['target']:.0f}" if f.get("target") else (an.get("Consensus") or "n.d.")
        if an.get("Potentiel (%)") is not None:
            ana += f" ({an['Potentiel (%)']:+.0f} %)"
        rows.append({"Ticker": t + (" ⟳" if f.get("cyclical") else "") + (" 🏦" if f.get("bank") else ""),
                     "Cours": f.get("price"),
                     "Juste valeur (fourchette)": (f"{fv['JV centrale']:.0f} ({fv['JV basse']:.0f}–{fv['JV haute']:.0f})"
                                                   if fv.get("JV centrale") else "n.d."),
                     "Écart JV (%)": fv.get("Écart (%)"),
                     "JV si bénéfices actuels durent": peak_txt(fv, f.get("price")),
                     "Précision JV": precision_txt(fv),
                     "Lecture de l'écart": divergence_txt(f, fv),
                     "Croissance implicite / attendue": reverse_txt(fv),
                     "Analystes · objectif": ana,
                     "BNA N+1 (évol.)": bna_evol(f.get("eps_this_y"), f.get("eps_next_y")),
                     "Résultats": f.get("earnings"),
                     "_t": t, "_cyc": bool(f.get("cyclical")), "_disp": fv.get("Dispersion (%)")})
    return pd.DataFrame(rows)


@st.fragment      # un clic dans cet onglet ne relance que l'onglet, pas le portefeuille
def render_screener(p, use_cache, max_days, use_oi):
    st.subheader("Screener : actions dont la juste valeur dépasse le cours")
    url = st.text_input("Adresse du screener Finviz", value=SCREENER_URL, key="scr_url",
                        help="Composez vos filtres sur finviz.com/screener, puis copiez l'adresse de la page de résultats.")
    if st.button("① Lire la liste du screener", key="scr_list"):
        logs = []
        with st.spinner("Lecture du screener Finviz…"):
            try:
                rows, total = fetch_screener(url, log=logs.append)
                st.session_state["scr_rows"] = (rows, total, logs)
                st.session_state.pop("scr_res", None)
            except Exception as e:
                st.error(f"Screener illisible : {type(e).__name__} — {e}")
    lst = st.session_state.get("scr_rows")
    if not lst:
        st.caption("① lit la liste des actions du screener (quelques secondes), ② calcule leur juste valeur "
                   "(≈ 1 s par action la première fois, cache 2 h).")
        return
    rows, total, logs = lst
    for x in logs:
        st.warning(x)
    if not rows:
        st.info("Aucune action dans ce screener.")
        return
    df0 = pd.DataFrame(rows)
    inds = sorted(df0["Industrie"].dropna().unique())
    excl = st.multiselect("Industries à exclure", inds, key=f"scr_excl_{hash(tuple(inds))}",
                          default=[i for i in inds if CYCLICAL_INDUSTRIES.search(i)],
                          help="Par défaut : mines (or, argent, cuivre…), acier, charbon, uranium — bénéfices de haut de cycle.")
    keep = df0[~df0["Industrie"].isin(excl)]
    st.caption(f"{len(df0)} actions lues, {len(keep)} retenues : " + ", ".join(keep["Ticker"]))
    with st.expander("Liste détaillée"):
        st.dataframe(keep, hide_index=True, width="stretch")
    if st.button(f"② Calculer la juste valeur ({len(keep)} actions)", type="primary", key="scr_go",
                 disabled=keep.empty):
        logs2 = []
        bar = st.progress(0.0)
        res = screen_values(list(keep["Ticker"]), p, use_cache,
                            progress=lambda v, t: bar.progress(v, text=t), log=logs2.append)
        bar.empty()
        if not res.empty:
            res = res.merge(keep[["Ticker", "Société", "Industrie"]].rename(columns={"Ticker": "_t"}), on="_t", how="left")
        st.session_state["scr_res"] = (res, logs2, datetime.now().strftime("%d/%m %H:%M"))
    out = st.session_state.get("scr_res")
    if not out:
        return
    res, logs2, when = out
    if res.empty:
        st.info("Aucune juste valeur calculée.")
        return
    c1, c2, c3 = st.columns([2, 1, 1])
    marge = c1.slider("Marge minimale (écart JV, %)", -50, 100, 15, 5, key="scr_marge")
    no_cyc = c2.checkbox("Masquer les valeurs cycliques ⟳", value=True, key="scr_nocyc")
    only_prec = c3.checkbox("Seulement les JV précises (🟢 / 🟠)", value=False, key="scr_prec",
                            help="Méthodes concordantes à ± 35 % près : une forte marge avec des méthodes divergentes "
                                 "est souvent un piège.")
    v = res[res["Écart JV (%)"].fillna(-999) >= marge]
    if no_cyc:
        v = v[~v["_cyc"]]
    if only_prec:
        v = v[v["Précision JV"].astype(str).str.startswith(("🟢", "🟠"))]
    v = v.sort_values("Écart JV (%)", ascending=False)
    cols = ["Ticker", "Société", "Industrie", "Cours", "Juste valeur (fourchette)", "Écart JV (%)",
            "JV si bénéfices actuels durent", "Précision JV",
            "Lecture de l'écart", "Croissance implicite / attendue", "Analystes · objectif", "BNA N+1 (évol.)", "Résultats"]
    st.dataframe(v[cols], hide_index=True, width="stretch", column_config=FMT)
    st.caption(f"Calcul du {when} : {len(v)} action(s) sur {len(res)} avec une marge ≥ {marge} %. Juste valeur rapide "
               "(fiche Finviz seule : flux libre déduit du P/FCF, sans états financiers) — indicative. "
               "⟳ cyclique, 🏦 modèle bancaire. Un écart très élevé est souvent un piège : vérifiez en analyse complète.")
    if logs2:
        with st.expander(f"{len(logs2)} fiche(s) non lue(s)"):
            st.write("\n".join(f"- {x}" for x in logs2))
    sel = st.multiselect("Envoyer en analyse complète (onglet « Valeurs à étudier »)", list(v["_t"]),
                         default=list(v["_t"])[:5], key="scr_sel")
    if st.button("➜ Analyse complète", key="scr_full", disabled=not sel):
        send_to_watchlist(sel, "Screener JV", max_days, use_cache, use_oi)


# ---------- Screeners des initiés : investors (NYSE / Nasdaq : Dataroma × OpenInsider) et AMF (Euronext)
def _ins_modules():
    try:
        import analyse as an_
        import openinsider as oi_
        from dataroma_scraper import GRIDS as grids_
        return an_, oi_, grids_
    except ImportError as e:
        return None, None, str(e)


SIG_FILTERS = ["🚨 / ✅ alerte 7 j", "🟢", "🟠", "🔴", "⚪", "achats groupés", "ventes groupées",
               "achats répétés (≥ 2 mois)"]


def _sig_filter(df, picks, col="signal", dot_col="dot"):
    """Pastilles choisies : OU entre les couleurs, ET pour les autres critères."""
    dots = [p for p in picks if p in ("🟢", "🟠", "🔴", "⚪")]
    m = pd.Series(True, index=df.index)
    if dots:
        m &= df[dot_col].isin(dots)
    if "achats groupés" in picks:
        m &= df[col].str.contains("achats groupés", na=False)
    if "ventes groupées" in picks:
        m &= df[col].str.contains("ventes groupées", na=False)
    if "🚨 / ✅ alerte 7 j" in picks:
        m &= df["alerte"].fillna("") != ""
    if "achats répétés (≥ 2 mois)" in picks:
        m &= (df["months6"].fillna(0) >= 2) | (df["buyers6"].fillna(0) >= 2)
    return df[m]


def _sig_fill(df):
    """Valeurs sans aucune transaction : ⚪ et cases vides lisibles."""
    df["dot"] = df["dot"].fillna("⚪")
    df["signal"] = df["signal"].fillna("⚪ aucune opération")
    df["rank"] = df["rank"].fillna(3)
    for c in ("achats_txt", "ventes_txt", "derniere_txt"):
        df[c] = df[c].fillna("—")
    for c in ("net12", "buy6", "sell3", "months6", "buyers6", "n_tx"):
        df[c] = df[c].fillna(0)
    df["monthly"] = df["monthly"].apply(lambda v: v if isinstance(v, list) else [0.0] * 12)
    df["frise"] = df["frise"].fillna("⬜" * 12)
    df["alerte"] = df["alerte"].fillna("")
    return df


def _sig_columns(cur):
    return {
        "Alerte 7 j": st.column_config.TextColumn(help="transaction d'un dirigeant dans les 7 derniers jours (date de "
                                                      "l'opération) : 🚨 vente · ✅ achat"),
        "Achats 6 mois": st.column_config.TextColumn(help="montant · dirigeants distincts (nb de mois avec au moins un achat)"),
        "Ventes 3 mois": st.column_config.TextColumn(help="montant · dirigeants distincts ; (mandat) = ventes faites "
                                                          "surtout par un gérant sous mandat"),
        "Net 12 mois": st.column_config.NumberColumn(format=f"%,.0f {cur}", help="achats − ventes sur 12 mois"),
        "12 derniers mois": st.column_config.TextColumn(help="un carré par mois, du plus ancien au plus récent : "
                                                            "🟩 achats nets · 🟥 ventes nettes · ⬜ aucune opération"),
    }


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def _inv_load(min_si):
    an_, oi_, _ = _ins_modules()
    bar = st.progress(0.0, text="Grilles Dataroma…")
    uni = an_.investors_universe(min_si, progress=lambda i, n, t: bar.progress(0.04 * i, text=f"Dataroma : {t}"))
    raw_tx, log = an_.fetch_openinsider(uni, 365, include_oe=True, progress=lambda i, n, t: bar.progress(
        0.2 + 0.8 * i / n, text=f"OpenInsider (12 mois) {i}/{n}"))
    bar.empty()
    return uni, raw_tx, log


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def _inv_conviction(tickers):
    """Fiches Dataroma (détenteurs, poids, dernier mouvement) des valeurs de l'univers -> {ticker: lignes}."""
    from concurrent.futures import ThreadPoolExecutor, as_completed
    bar = st.progress(0.0, text="Fiches Dataroma (conviction des fonds)…")

    def one(t):
        return t, parse_dataroma(_get(_session(), f"https://www.dataroma.com/m/stock.php?sym={t}", (0.2, 0.5)))["rows"]
    out = {}
    with ThreadPoolExecutor(4) as ex:
        futs = [ex.submit(one, t) for t in tickers]
        for i, fu in enumerate(as_completed(futs), 1):
            try:
                t, rows = fu.result()
                out[t] = rows
            except Exception:  # noqa: BLE001 — fiche illisible : signal par le nombre de fonds
                pass
            bar.progress(i / len(futs), text=f"Fiches Dataroma {i}/{len(futs)}")
    bar.empty()
    return out


@st.fragment      # un clic dans cet onglet ne relance que l'onglet, pas le portefeuille
def render_investors_screener(max_days, use_cache, use_oi):
    """Onglet « Screener investors » : valeurs du S&P 500 détenues par les superinvestors (grille Dataroma
    « Ownership ») ; signal des dirigeants (OpenInsider, 12 mois) et signal des fonds (activité Dataroma)."""
    import signaux as sg
    an, oi, grids = _ins_modules()
    if an is None:
        st.error(f"Modules du screener introuvables ({grids}) : placez analyse.py, openinsider.py, "
                 "dataroma_scraper.py, exports.py et signaux.py dans le dossier du dashboard.")
        return
    st.caption("Univers : valeurs du S&P 500 détenues par les superinvestors (grille Dataroma « Ownership »). "
               "Deux signaux côte à côte : les dirigeants (OpenInsider, 12 mois, levées d'options exclues) et les "
               "fonds (renforcements / allègements, Dataroma).")
    with st.expander("⚙️ Réglages", expanded="inv_data" not in st.session_state):
        c = st.columns(3)
        min_si = c[0].number_input("Nb minimum de superinvestors détenteurs", 1, 50, 1, key="inv_min_si",
                                   help="1 ≈ 430 valeurs, 3 ≈ 250, 5 ≈ 150 (novembre 2026).")
        min_sig = c[1].select_slider("Montant significatif", [25_000, 50_000, 100_000, 250_000, 500_000, 1_000_000],
                                     100_000, key="inv_min_sig",
                                     format_func=lambda v: f"{v:,.0f} $".replace(",", " "),
                                     help="Sous ce montant, achats (6 mois) ou ventes (3 mois) ne comptent pas "
                                          "dans le signal.")
        min_value = c[1].select_slider("Montant minimum par achat", [0, 25_000, 50_000, 90_000, 250_000],
                                       an.DEFAULT_MIN_BUY, key="inv_min_value",
                                       format_func=lambda v: "aucun" if v == 0 else f"{v:,.0f} $".replace(",", " "),
                                       help="Ignore les petits achats symboliques ; les ventes sont toutes comptées.")
        excl10 = c[2].checkbox("Exclure les actionnaires >10 %", value=False, key="inv_excl10",
                               help="Ex. Berkshire Hathaway, déclaré initié chez Lennar.")
        use_conv = c[2].checkbox("Fonds : pondérer par la conviction", value=True, key="inv_conv",
                                 help="Lit la fiche Dataroma de chaque valeur (poids de la ligne dans le portefeuille de "
                                      "chaque gérant) : ≈ 1 à 2 min la première fois, puis gardée 6 h. Décoché : "
                                      "nombre de fonds qui renforcent / allègent.")
        go = c[2].button("Lancer l'analyse", type="primary", key="inv_go")
    if go:
        try:
            if not use_cache:
                _inv_load.clear()
            st.session_state["inv_data"] = _inv_load(int(min_si))
            st.session_state["inv_conv_rows"] = (_inv_conviction(tuple(st.session_state["inv_data"][0]["ticker"]))
                                                 if use_conv else None)
        except Exception as e:  # noqa: BLE001
            st.error(f"Erreur : {e}")
    if "inv_data" not in st.session_state:
        st.info("Réglez les paramètres puis cliquez sur **Lancer l'analyse** (≈ 15 s).")
        return
    uni, raw_tx, log = st.session_state["inv_data"]
    tx_all = an.filter_tx(raw_tx, excl10, min_value)
    tx, tx_oe = an.split_oe(tx_all)              # ventes liées à une levée d'options : hors signal
    sig = sg.summarize(an.signal_tx(tx), today=TODAY.date(), min_sig=min_sig, cur="$")
    final = _sig_fill(uni.merge(sig, left_on="ticker", right_on="key", how="left"))
    ex = an.exit_rows(tx_oe)
    warn = {k: sg.top_exit_warning(g, today=TODAY.date()) for k, g in ex.groupby("key")} if not ex.empty else {}
    final["sortie"] = final["ticker"].map(warn).fillna("")
    final["fonds"] = [sg.fund_signal(*r) for r in final[["add_q", "red_q", "add_6m", "red_6m"]].itertuples(index=False)]
    conv_rows = st.session_state.get("inv_conv_rows")
    if conv_rows:                          # conviction (poids des lignes) : même calcul que l'analyse des valeurs
        cv = {t: sg.conviction_signal(r)["txt"] for t, r in conv_rows.items() if r}
        final["fonds"] = final["ticker"].map(cv).fillna(final["fonds"])
    final["fonds_dot"] = final["fonds"].str[:1]
    qs = [sg.fund_quarters(*r) for r in final[["add_q", "red_q", "add_6m", "red_6m"]].itertuples(index=False)]
    final["fonds_q"] = [f"▲ {a}  ▼ {r}" for a, r, _, _ in qs]
    final["fonds_p"] = [f"▲ {a}  ▼ {r}" for _, _, a, r in qs]
    final["concord"] = (final["dot"] == "🟢") & (final["fonds_dot"] == "🟢")
    final = final.sort_values(["rank", "concord", "net12"], ascending=[True, False, False]).reset_index(drop=True)
    if raw_tx.empty:
        st.warning("OpenInsider n'a renvoyé aucune transaction : voir l'onglet **Diagnostic**.")

    c = st.columns(6)
    c[5].metric("Alertes 7 j", f"🚨 {final['alerte'].str.contains('🚨').sum()} · ✅ {final['alerte'].str.contains('✅').sum()}")
    c[0].metric("Univers", len(final))
    c[1].metric("🟢 dirigeants", int((final["dot"] == "🟢").sum()))
    c[2].metric("🔴 dirigeants", int((final["dot"] == "🔴").sum()))
    c[3].metric("Achats groupés", int(final["signal"].str.contains("achats groupés").sum()))
    c[4].metric("🟢 dirigeants + 🟢 fonds", int(final["concord"].sum()))

    # sentiment de l'univers : 90 derniers jours comparés aux 90 précédents (dates de déclaration)
    if not tx.empty:
        fd = pd.to_datetime(tx["filing_date"], errors="coerce")
        now_ = oi.sentiment(tx[fd > TODAY - pd.Timedelta(days=90)])
        prev_ = oi.sentiment(tx[(fd <= TODAY - pd.Timedelta(days=90)) & (fd > TODAY - pd.Timedelta(days=180))])
        with st.container(border=True):
            st.markdown("**Sentiment des dirigeants de l'univers** · 90 derniers jours, comparés aux 90 précédents")
            k = st.columns(4)
            for col, lbl, key_, fmt in ((k[0], "Acheteurs / vendeurs", "ratio_ins", "{:.3f}"),
                                        (k[1], "Nb achats / nb ventes", "ratio_tx", "{:.3f}"),
                                        (k[2], "Montants achats / ventes", "ratio_usd", "{:.3f}"),
                                        (k[3], "Valeurs acheteurs > vendeurs", "valeurs_pos", "{:.0f}")):
                a, b = now_.get(key_), prev_.get(key_)
                col.metric(lbl, "–" if a is None else fmt.format(a),
                           f"{(a / b - 1):+.0%} vs 90 j préc." if a is not None and b else None)
            st.caption("Les ventes dépassent toujours largement les achats dans les grandes capitalisations : "
                       "c'est l'évolution qui compte, pas le niveau.")

    f1, f2 = st.columns([3, 1])
    picks = f1.pills("Filtrer (dirigeants)", SIG_FILTERS + ["⚠️ sortie PDG / DAF", "🟢 dirigeants + 🟢 fonds"],
                     selection_mode="multi", key="inv_pills") or []
    sect = f2.multiselect("Secteur", sorted(final["sector"].dropna().unique()), key="inv_sector")
    view = _sig_filter(final, [p for p in picks if p not in ("🟢 dirigeants + 🟢 fonds", "⚠️ sortie PDG / DAF")])
    if "⚠️ sortie PDG / DAF" in picks:
        view = view[view["sortie"] != ""]
    if "🟢 dirigeants + 🟢 fonds" in picks:
        view = view[view["concord"]]
    if sect:
        view = view[view["sector"].isin(sect)]
    show = pd.DataFrame({
        "Alerte 7 j": view["alerte"], "Ticker": view["ticker"], "Société": view["name"], "Secteur": view["sector"],
        "Sortie PDG / DAF": view["sortie"],
        "Dirigeants · signal": view["signal"], "Dirigeants · achats 6 mois": view["achats_txt"],
        "Dirigeants · ventes 3 mois": view["ventes_txt"], "Dirigeants · net 12 mois": view["net12"],
        "Dirigeants · 12 derniers mois": view["frise"], "Dirigeants · dernière opération": view["derniere_txt"],
        "Fonds · signal": view["fonds"], "Fonds · dernier trimestre": view["fonds_q"],
        "Fonds · trimestre précédent": view["fonds_p"],
        "Fonds · détenteurs": view["owners"],
        "Dataroma": dr_link(view["ticker"]), "OpenInsider": "http://openinsider.com/" + view["ticker"]})
    t1, t2, t3 = st.tabs(["Tableau", "Transactions", "Diagnostic"])
    with t1:
        cfg = {(k if k == "Alerte 7 j" else f"Dirigeants · {k[0].lower()}{k[1:]}"): v
               for k, v in _sig_columns("$").items()}
        st.dataframe(show, width="stretch", hide_index=True, height=min(600, 38 + 35 * max(len(show), 1)),
                     column_config=cfg | {
            "Sortie PDG / DAF": st.column_config.TextColumn(help=(
                "⚠️ un PDG, DAF, président ou fondateur a vendu au moins la moitié de sa détention dans les 6 derniers "
                "mois, en lien avec une levée d'options — hors signal 🟢/🟠/🔴/⚪")),
            "Dirigeants · signal": st.column_config.TextColumn(width="medium", help="dirigeants et administrateurs "
                                                               "(OpenInsider, 12 mois)"),
            "Dirigeants · dernière opération": st.column_config.TextColumn(width="medium",
                                                                           help="dernière transaction d'un dirigeant"),
            "Fonds · signal": st.column_config.TextColumn(width="large", help=(
                "Pondéré par la conviction (case cochée) : chaque mouvement compte pour le poids qu'il représente dans "
                "le portefeuille du gérant (points), 5 points au plus par fonds — 🟢 ≥ +0,5 · 🔴 ≤ −0,5 · 🟠 achats et "
                "ventes qui se compensent · ⚪ symbolique. Sinon : "
                "superinvestors (Dataroma), dernier trimestre : 🟢 plus de fonds qui renforcent que de fonds qui "
                "allègent · 🔴 l'inverse · 🟠 autant · ⚪ aucun mouvement ; ↗ / ↘ : solde meilleur / moins bon "
                "qu'au trimestre précédent (écart ≥ 2 fonds)")),
            "Fonds · dernier trimestre": st.column_config.TextColumn(help="▲ nb de superinvestors qui ont acheté ou "
                                                                          "renforcé · ▼ nb qui ont vendu ou allégé"),
            "Fonds · trimestre précédent": st.column_config.TextColumn(help="même lecture, trimestre d'avant (grille "
                                                                            "« 6 mois » moins le dernier trimestre)"),
            "Fonds · détenteurs": st.column_config.NumberColumn(help="nb de superinvestors qui détiennent la valeur"),
            "Dataroma": st.column_config.LinkColumn(display_text="fiche"),
            "OpenInsider": st.column_config.LinkColumn(display_text="fiche")})
        st.caption("Colonnes « Dirigeants » : transactions des dirigeants et administrateurs (OpenInsider). " + sg.LEGEND
                   + " Colonnes « Fonds » : superinvestors (Dataroma, déclarations trimestrielles 13F) — rond pondéré par la "
                   "conviction (poids des lignes dans le portefeuille de chaque gérant, case cochée) ; ▲ nb de fonds "
                   "qui ont acheté ou renforcé, ▼ nb qui ont vendu ou allégé ; sans conviction : rond d'après le dernier trimestre, "
                   "↗ / ↘ = solde meilleur / moins bon qu'au trimestre précédent, achats groupés = au moins 3 fonds "
                   "acheteurs dans le trimestre. La frise et la dernière opération ne concernent que "
                   "les dirigeants.")
    with t2:
        tk = st.selectbox("Valeur", ["(toutes)"] + (sorted(tx_all["ticker"].unique()) if not tx_all.empty else []),
                          key="inv_tx_tk")
        st.dataframe(an.display_tx(tx_all if tk == "(toutes)" else tx_all[tx_all["ticker"] == tk]), width="stretch",
                     hide_index=True)
        st.caption("Type « S - Sale+OE » : vente liée à une levée d'options — affichée ici, hors signal.")
    with t3:
        st.write(f"{len(raw_tx)} transactions OpenInsider lues sur 12 mois, {len(tx)} retenues pour le signal "
                 f"(+ {len(tx_oe)} ventes liées à une levée d'options, hors signal).")
        st.dataframe(pd.DataFrame(log), width="stretch", hide_index=True)

    sel = st.multiselect("Envoyer en analyse complète (onglet « Valeurs à étudier »)", list(view["ticker"]),
                         default=list(view["ticker"])[:5], key="inv_sel")
    if st.button("➜ Analyse complète", key="inv_full", disabled=not sel):
        send_to_watchlist(sel, "Screener investors", max_days, use_cache, use_oi)
    _sig_exports(show, "investors", "Screener investors (S&P 500 × superinvestors)", {"Dirigeants · net 12 mois"},
                 {"Dataroma", "OpenInsider"}, "inv",
                 f"Univers : S&P 500 détenu par ≥ {int(min_si)} superinvestor(s) (Dataroma) ; dirigeants : OpenInsider "
                 f"12 mois, seuil significatif {min_sig:,.0f} $.".replace(",", " "))


def dr_link(tickers):
    return "https://www.dataroma.com/m/stock.php?sym=" + tickers


def _sig_exports(show, stem, title, money_cols, link_cols, key, note):
    from exports import to_excel_bytes
    out = show
    stamp = f"{TODAY:%Y%m%d}"
    b = st.columns(2)
    b[0].download_button("⬇️ Excel", to_excel_bytes({title[:31]: out}, money_cols, link_cols, note),
                         f"{stem}_{stamp}.xlsx",
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch",
                         key=f"{key}_dl_xlsx")
    b[1].download_button("⬇️ CSV", out.to_csv(index=False, sep=";").encode("utf-8-sig"), f"{stem}_{stamp}.csv",
                         "text/csv", width="stretch", key=f"{key}_dl_csv")
    st.caption(note + " Informations à titre indicatif, pas un conseil en investissement.")


def _sbf120():
    fn = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sbf120.csv")
    return pd.read_csv(fn, sep=";", dtype=str) if os.path.exists(fn) else None


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def _amf_load(universe):
    import amf
    from datetime import timedelta
    bar = st.progress(0.0, text="Liste des déclarations AMF…")
    rows = amf.list_declarations(TODAY.date() - timedelta(days=366), progress=lambda m: bar.progress(0.05, text=m))
    names, missing = {}, []
    if universe == "sbf120":
        lst = _sbf120()
        jm = amf.jetons_for_isins(list(lst["isin"]), os.path.join(CACHE_DIR, "amf_jetons.json"),
                                  progress=lambda m: bar.progress(0.1, text=m))
        nm = dict(zip(lst["isin"], lst["nom"]))
        names = {j: nm[i] for i, (j, _) in jm.items()}
        missing = [nm[i] for i in lst["isin"] if i not in jm]
        rows = [r for r in rows if r["jeton"] in names]
    tx, diag = amf.load_transactions(rows, os.path.join(CACHE_DIR, "amf"), progress=lambda k, n: bar.progress(
        0.1 + 0.9 * k / max(n, 1), text=f"Déclarations AMF {k}/{n} (lues une seule fois, puis gardées sur disque)"))
    bar.empty()
    return len(rows), tx, diag, names, missing


@st.fragment      # un clic dans cet onglet ne relance que l'onglet, pas le portefeuille
def render_amf_screener():
    """Onglet « Screener AMF » : déclarations des dirigeants des sociétés cotées à Paris (base BDIF de l'AMF)."""
    import signaux as sg
    try:
        import amf  # noqa: F401
        import pypdf  # noqa: F401
    except ImportError as e:
        st.error(f"Module manquant ({e}) : placez amf.py et signaux.py dans le dossier du dashboard et installez "
                 "pypdf (pip install -r requirements.txt).")
        return
    st.caption("Déclarations des dirigeants et personnes liées (article 19 du règlement MAR) publiées par l'AMF : "
               "achats et ventes d'actions sur 12 mois. Actions gratuites ou de performance, options, conversions, "
               "nantissements, donations, parts de FCPE et droits de souscription sont exclus.")
    with st.expander("⚙️ Réglages", expanded="amf_data" not in st.session_state):
        c = st.columns(3)
        universe = c[0].radio("Sociétés", ["sbf120", "all"], key="amf_univ",
                              format_func=lambda k: {"sbf120": "SBF 120",
                                                     "all": "Toutes (Euronext Paris, Growth, Access)"}[k],
                              help="SBF 120 : liste du fichier sbf120.csv (Euronext, novembre 2026), modifiable. "
                                   "Toutes : environ 5 500 déclarations sur 12 mois — premier chargement de quelques "
                                   "minutes, puis seules les nouvelles déclarations sont lues.")
        min_sig = c[1].select_slider("Montant significatif", [10_000, 25_000, 50_000, 100_000, 250_000, 500_000],
                                     50_000, key="amf_min_sig",
                                     format_func=lambda v: f"{v:,.0f} €".replace(",", " "),
                                     help="Sous ce montant, achats (6 mois) ou ventes (3 mois) ne comptent pas dans "
                                          "le signal.")
        go = c[2].button("Lancer l'analyse", type="primary", key="amf_go")
    if universe == "sbf120" and _sbf120() is None:
        st.error("Fichier sbf120.csv introuvable dans le dossier du dashboard.")
        return
    if go:
        try:
            st.session_state["amf_data"] = (_amf_load(universe), universe)
        except Exception as e:  # noqa: BLE001
            st.error(f"AMF indisponible : {type(e).__name__} : {e}")
    if "amf_data" not in st.session_state:
        st.info("Réglez les paramètres puis cliquez sur **Lancer l'analyse**.")
        return
    (n_decl, tx, diag, names, missing), univ = st.session_state["amf_data"]
    stx = (pd.DataFrame({"key": tx["jeton"], "date": tx["date"], "person": tx["person"], "side": tx["side"],
                        "value": tx["value"], "mandate": tx["mandate"],
                        "name": tx["declarant"] + tx["fonction"].map(lambda x: f" ({x})" if x else "")})
           if not tx.empty else None)
    sig = sg.summarize(stx, today=TODAY.date(), min_sig=min_sig, cur="€")
    soc = (tx.sort_values("date_pub").groupby("jeton")["societe"].last().to_dict() if not tx.empty else {})
    last_pdf = (tx.sort_values(["date", "date_pub"]).groupby("jeton")["pdf"].last().to_dict() if not tx.empty else {})
    if univ == "sbf120":                       # toutes les valeurs de l'indice, même sans opération
        base = pd.DataFrame({"key": list(names)})
        sig = _sig_fill(base.merge(sig, on="key", how="left"))
    sig["societe"] = sig["key"].map(lambda j: names.get(j) or soc.get(j) or j)
    sig = sig.sort_values(["rank", "net12"], ascending=[True, False]).reset_index(drop=True)

    c = st.columns(6)
    c[5].metric("Alertes 7 j", f"🚨 {sig['alerte'].str.contains('🚨').sum()} · ✅ {sig['alerte'].str.contains('✅').sum()}")
    c[0].metric("Sociétés", len(sig))
    c[1].metric("🟢 favorables", int((sig["dot"] == "🟢").sum()))
    c[2].metric("🟠 mitigées", int((sig["dot"] == "🟠").sum()))
    c[3].metric("🔴 défavorables", int((sig["dot"] == "🔴").sum()))
    c[4].metric("Achats groupés", int(sig["signal"].str.contains("achats groupés").sum()))
    picks = st.pills("Filtrer", SIG_FILTERS, selection_mode="multi", key="amf_pills") or []
    view = _sig_filter(sig, picks)
    show = pd.DataFrame({
        "Alerte 7 j": view["alerte"], "Signal": view["signal"], "Société": view["societe"],
        "Achats 6 mois": view["achats_txt"],
        "Ventes 3 mois": view["ventes_txt"], "Net 12 mois": view["net12"], "12 derniers mois": view["frise"],
        "Dernière opération": view["derniere_txt"], "Opérations (12 mois)": view["n_tx"].astype(int),
        "Dernière déclaration": view["key"].map(last_pdf)})
    t1, t2, t3 = st.tabs(["Tableau", "Transactions", "Diagnostic"])
    with t1:
        st.dataframe(show, width="stretch", hide_index=True, height=min(600, 38 + 35 * max(len(show), 1)),
                     column_config=_sig_columns("€") | {
            "Signal": st.column_config.TextColumn(width="medium"),
            "Dernière déclaration": st.column_config.LinkColumn(display_text="PDF AMF")})
        st.caption(sg.LEGEND + " Ventes « (mandat) » : faites par un gérant sous mandat de gestion, moins "
                   "parlantes qu'une vente décidée par le dirigeant lui-même.")
        if missing:
            st.caption("Sans déclaration trouvée sur la base AMF (sociétés étrangères qui déclarent à leur propre "
                       "régulateur, ou aucune déclaration récente) : " + ", ".join(missing))
    with t2:
        if tx.empty:
            st.caption("Aucune transaction retenue.")
        else:
            opts = sig.loc[sig["n_tx"] > 0, ["key", "societe"]]
            pick = st.selectbox("Société", ["(toutes)"] + list(opts["societe"]), key="amf_tx_soc")
            d = tx if pick == "(toutes)" else tx[tx["jeton"] == opts.set_index("societe")["key"].get(pick)]
            d = d.sort_values("date", ascending=False)
            st.dataframe(pd.DataFrame({
                "Date": d["date"].dt.strftime("%d/%m/%Y"), "Société": d["jeton"].map(lambda j: names.get(j) or soc.get(j)),
                "Déclarant": d["declarant"], "Fonction": d["fonction"],
                "Sens": d["side"].map({"A": "🟢 achat", "V": "🔴 vente"}), "Nature": d["nature"],
                "Prix": d["prix"], "Volume": d["volume"], "Montant": d["value"], "Devise": d["devise"],
                "Lieu": d["lieu"], "Mandat": d["mandate"].map({True: "oui", False: ""}),
                "Correction": d["correction"].map({True: "oui", False: ""}), "PDF": d["pdf"]}),
                width="stretch", hide_index=True, column_config={
                    "Prix": st.column_config.NumberColumn(format="%.4f"),
                    "Volume": st.column_config.NumberColumn(format="%,.0f"),
                    "Montant": st.column_config.NumberColumn(format="%,.0f"),
                    "PDF": st.column_config.LinkColumn(display_text="PDF AMF")})
    with t3:
        ok = int((diag["motif"].fillna("") == "").sum()) if not diag.empty else 0
        st.write(f"{n_decl} déclarations sur 12 mois · {len(diag)} opérations lues · {ok} achats/ventes d'actions "
                 f"retenus ({len(tx)} après remplacement des déclarations corrigées).")
        if not diag.empty:
            st.dataframe(diag[diag["motif"].fillna("") != ""]["motif"].str.replace(r" : .*", "", regex=True)
                         .value_counts().rename_axis("Motif d'exclusion").reset_index(name="Opérations"),
                         hide_index=True)
            err = diag[diag["motif"].fillna("").str.startswith(("lecture", "aucune"))]
            if len(err):
                st.markdown(f"**{len(err)} déclaration(s) illisible(s)**")
                st.dataframe(err[["numero", "societe", "date_pub", "motif", "pdf"]], hide_index=True,
                             column_config={"pdf": st.column_config.LinkColumn(display_text="PDF")})
        st.caption(f"Source : base BDIF de l'AMF ({amf.FRONT}). Les PDF déjà lus sont gardés dans "
                   f"{os.path.join(CACHE_DIR, 'amf')} ; la liste est relue à chaque analyse (cache 6 h).")
    _sig_exports(show, "amf", "Screener AMF", {"Net 12 mois"}, {"Dernière déclaration"}, "amf",
                 f"Source : AMF (BDIF), déclarations des dirigeants sur 12 mois ; "
                 f"{'SBF 120' if univ == 'sbf120' else 'toutes sociétés'} ; seuil significatif "
                 f"{min_sig:,.0f} €.".replace(",", " "))


def peak_txt(fv, price):
    """Valeur cyclique : « 1326 (+50 %) » = juste valeur si les bénéfices actuels durent (multiples sur le BNA N+1)."""
    v = fv.get("JV si pic")
    if not v:
        return "—"
    return f"{v:.0f}" + (f" ({(v / price - 1) * 100:+.0f} %)" if price else "")


def reverse_txt(fv):
    """« 🟢 −3 % / +12 % » : croissance que suppose le cours (DCF inversé) / croissance attendue
    (DCF d'entreprise : croissance annuelle moyenne du CA sur 5 ans)."""
    gi, ga = fv.get("g implicite (%)"), fv.get("g attendue (%)")
    if gi is None or ga is None:
        return "n.d."
    d = ga - gi
    dot = "🟢" if d >= 5 else "🔴" if d <= -5 else "⚪"
    lim = " (≥ 80 %)" if gi >= 79.9 else ""
    return f"{dot} {gi:+.0f} %{lim} / {ga:+.0f} %"


def bna_evol(e0, e1):
    """« 4,44 (▲ +32 %) » : BNA N+1 et sa progression par rapport au BNA de l'exercice en cours."""
    if e1 is None:
        return "n.d."
    if not e0 or e0 <= 0:
        return f"{e1:.2f}" + (" (retour aux bénéfices)" if e0 is not None and e0 <= 0 < e1 else "")
    g = (e1 / e0 - 1) * 100
    arrow = "▲" if g > 0.5 else "▼" if g < -0.5 else "="
    return f"{e1:.2f} ({arrow} {g:+.0f} %)"


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def fund_grid():
    """Activité des superinvestors sur tout le S&P 500 (grilles Dataroma), comme le screener investors."""
    an_, _, _ = _ins_modules()
    return an_.investors_universe(0).set_index("ticker")


def _sig_settings():
    """Seuils du screener investors (mêmes réglages pour l'analyse des valeurs)."""
    try:
        return (st.session_state.get("inv_min_sig", 100_000), st.session_state.get("inv_min_value", 50_000),
                st.session_state.get("inv_excl10", False))
    except Exception:  # noqa: BLE001
        return 100_000, 50_000, False


def _inner(txt):
    """Détails d'un bloc entre parenthèses, « / » (ou « · ») séparant les blocs achats / ventes / net :
    « 78,50 M$ · 4 dirig. » -> « 78,50 M$ (4 dirig.) » ; « 418 k$ · 1 dirig. (1 mois) » -> « 418 k$ (1 dirig., 1 mois) »."""
    m = re.match(r"^(.*?) · (\d+ dirig\.)(?: \((.*)\))?$", txt or "")
    if not m:
        return txt
    return f"{m.group(1)} ({m.group(2)}" + (f", {m.group(3)}" if m.group(3) else "") + ")"


def concord_insiders(ins, tk):
    """Signal des dirigeants calculé comme dans le screener investors (signaux.summarize, mêmes seuils) :
    ventes avec levée d'options exclues, achats sous le montant minimum ignorés."""
    if not ins or ins.get("_tx") is None:
        return ins
    import signaux as sg
    min_sig, min_buy, excl10 = _sig_settings()
    rows = [t for t in ins["_tx"] if not (t["typ"] == "Sale" and t["oe"])
            and not (t["typ"] == "Buy" and t["val"] * 1e6 < min_buy)
            and not (excl10 and "10%" in (t["rel"] or ""))]
    out = dict(ins)
    oe_rows = [t for t in ins["_tx"] if t["typ"] == "Sale" and t["oe"]]
    warn = sg.top_exit_warning(pd.DataFrame({
        "date": [t["date"] for t in oe_rows], "name": [t["name"] for t in oe_rows],
        "role": [(t["rel"] or "").split(" · ")[0] for t in oe_rows], "side": "V", "oe": True,
        "pct_sold": [t.get("pct_sold") for t in oe_rows]}), today=TODAY.date()) if oe_rows else ""
    out["Sortie dirigeant"] = warn
    if not rows:
        out["Signal"] = "⚪"
        out["Résumé"] = (f"{warn} · " if warn else "") + "⚪ aucun achat ni vente significatif en bourse sur 12 mois"
        return out
    df = pd.DataFrame({"key": tk, "date": [t["date"] for t in rows], "person": [t["name"].lower() for t in rows],
                       "side": ["A" if t["typ"] == "Buy" else "V" for t in rows],
                       "value": [t["val"] * 1e6 for t in rows], "mandate": False,
                       "name": [f"{t['name']} ({t['rel'].split(' · ')[0]})" if t["rel"] else t["name"] for t in rows]})
    sm = sg.summarize(df, today=TODAY.date(), min_sig=min_sig, cur="$")
    if sm.empty:                                  # transactions toutes antérieures aux 12 mois calendaires
        out["Signal"] = "⚪"
        out["Résumé"] = (f"{warn} · " if warn else "") + "⚪ aucun achat ni vente significatif en bourse sur 12 mois"
        return out
    r = sm.iloc[0]
    net = f"{'+' if r['net12'] >= 0 else '−'}{sg.money(r['net12'], '$')}"
    out.update({"Signal": r["signal"], "Frise": r["frise"], "Dernière opération": r["derniere_txt"],
                "Alerte": r["alerte"],
                "Résumé": (f"{r['alerte']} · " if r["alerte"] else "") + (f"{warn} · " if warn else "")
                          + f"{r['signal']} · achats 6 m {_inner(r['achats_txt'])} · ventes 3 m {_inner(r['ventes_txt'])} · "
                          f"net 12 m {net}"})
    return out


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def amf_company_tx(isin):
    """Transactions AMF des dirigeants d'une société (12 mois), lues comme dans le screener AMF."""
    import amf
    from datetime import timedelta
    jeton, name, rows = amf.company_declarations(isin, TODAY.date() - timedelta(days=366),
                                                 os.path.join(CACHE_DIR, "amf_jetons.json"))
    tx, _ = amf.load_transactions(rows, os.path.join(CACHE_DIR, "amf")) if rows else (pd.DataFrame(), None)
    return jeton, name, tx


def amf_insiders(isin):
    """Signal AMF des dirigeants (valeur française lue sur Zonebourse), même calcul et même seuil que le screener
    AMF ; reprend directement les transactions du screener si la société y figure déjà."""
    import signaux as sg
    try:
        min_sig = st.session_state.get("amf_min_sig", 50_000)
    except Exception:  # noqa: BLE001
        min_sig = 50_000
    res = _guarded(f"amf_{isin}", lambda: amf_company_tx(isin))
    if res is None:
        return None
    jeton, name, tx = res
    if not jeton:
        return {"Signal": "⚪", "Résumé": "⚪ société introuvable dans la base AMF (les sociétés étrangères déclarent "
                                         "souvent à leur propre régulateur)", "Source": "déclarations AMF",
                "Achats notables": []}
    out = {"Source": "déclarations AMF", "Achats notables": [], "Achats 365j": 0.0, "Ventes 365j": 0.0, "Acheteurs": 0,
           "Achats groupés (max en 30 j)": 0, "Achats PDG / DAF (M$)": None, "Hausse médiane de détention (%)": None,
           "Ventes avec levée d'options (M$)": None, "Devise": "€"}
    if tx is None or tx.empty:
        out.update({"Signal": "⚪", "Résumé": "⚪ aucun achat ni vente d'actions déclaré à l'AMF sur 12 mois"})
        return out
    df = pd.DataFrame({"key": jeton, "date": tx["date"], "person": tx["person"], "side": tx["side"],
                       "value": tx["value"], "mandate": tx["mandate"],
                       "name": tx["declarant"] + tx["fonction"].map(lambda x: f" ({x})" if x else "")})
    sm = sg.summarize(df, today=TODAY.date(), min_sig=min_sig, cur="€")
    if sm.empty:
        out.update({"Signal": "⚪", "Résumé": "⚪ aucun achat ni vente d'actions déclaré à l'AMF sur 12 mois"})
        return out
    r = sm.iloc[0]
    b, v = tx[tx["side"] == "A"], tx[tx["side"] == "V"]
    net = f"{'+' if r['net12'] >= 0 else '−'}{sg.money(r['net12'], '€')}"
    out.update({
        "Signal": r["signal"], "Frise": r["frise"], "Dernière opération": r["derniere_txt"],
        "Alerte": r["alerte"],
        "Résumé": (f"{r['alerte']} · " if r["alerte"] else "")
                  + f"{r['signal']} · achats 6 m {_inner(r['achats_txt'])} · ventes 3 m {_inner(r['ventes_txt'])} · net 12 m {net}",
        "Achats 365j": round(b["value"].sum() / 1e6, 2), "Ventes 365j": round(v["value"].sum() / 1e6, 2),
        "Acheteurs": int(b["person"].nunique()), "Achats groupés (max en 30 j)": int(r["group_buy"]),
        "Achats notables": [f"{x.declarant} ({x.fonction}) {x.date:%d/%m/%y} : {sg.money(x.value, '€')}"
                            for x in b.sort_values("value", ascending=False).head(5).itertuples()],
        "Ventes sous mandat": bool(v["mandate"].any())})
    return out


def concord_funds(si, tk):
    """Signal des fonds calculé comme dans le screener investors (grilles Dataroma : dernier trimestre et tendance
    par rapport au trimestre précédent) ; hors S&P 500, résumé de la fiche Dataroma inchangé."""
    import signaux as sg
    try:
        g = _guarded("fund_grid", fund_grid)
    except Exception:  # noqa: BLE001
        g = None
    k = (tk or "").replace("-", ".").upper()
    if g is None or k not in g.index:
        return si
    r = g.loc[k]
    aq, rq, ap, rp = sg.fund_quarters(r["add_q"], r["red_q"], r["add_6m"], r["red_6m"])
    out = dict(si)
    counts = f"{int(r['owners'])} fonds · trim. ▲ {aq} ▼ {rq} · préc. ▲ {ap} ▼ {rp}"
    if si.get("Lignes"):                  # fiche Dataroma lue : signal pondéré par la conviction (même calcul que le
        conv = si["Conviction"]           # screener investors)
        out.update({"Signal": conv["dot"], "Grille": True, "Détenteurs": int(r["owners"]),
                    "Résumé": f"{conv['txt']} · {counts}"})
    else:                                 # sinon : nombre de fonds (grilles)
        sig = sg.fund_signal(r["add_q"], r["red_q"], r["add_6m"], r["red_6m"])
        out.update({"Signal": sig, "Grille": True, "Détenteurs": int(r["owners"]), "Résumé": f"{sig} · {counts}"})
    return out


def analyze(f, p, ov=None):
    fv, an = fair_value(f, p), analyst_stats(f)
    si, pc, ins = superinvestor_stats(f["dataroma"], f["price"]), put_call_stats(f["options"]), f["insiders"]
    ins = concord_insiders(ins, f.get("ticker") or "")
    if not ins and f.get("isin") and not str(f["isin"]).startswith("US"):    # valeur européenne (Zonebourse)
        ins = amf_insiders(f["isin"])
    si = concord_funds(si, f.get("ticker") or "")
    sc = scenarios(f, ov)
    jv = (f"{fv['JV centrale']:.0f} ({fv['JV basse']:.0f}–{fv['JV haute']:.0f})"
          if fv.get("JV centrale") else "n.d.")
    ana = f"{an['Consensus'] or '?'} · {f['target']:.0f}" if f["target"] else (an["Consensus"] or "n.d.")
    if an["Potentiel (%)"] is not None:
        ana += f" ({an['Potentiel (%)']:+.0f} %)"
    ins_txt = ins["Résumé"] if ins else "n.d."
    pcs = (f"{pc.get('P/C encours') or '–'} / {pc.get('P/C encours ±10 %') or '–'}" if pc else "n.d.")
    lev = f["net_debt"] / f["ebitda"] if f["net_debt"] is not None and f["ebitda"] and f["ebitda"] > 0 else None
    row = {"Cours": f["price"], "Devise": f.get("currency", "USD"),
           "Juste valeur (fourchette)": jv,
           "Écart JV (%)": fv.get("Écart (%)"),                      # (JV centrale / cours - 1) : >0 = sous-évaluée
           "JV si bénéfices actuels durent": peak_txt(fv, f["price"]),
           "Précision JV": precision_txt(fv) + (" ⚠️" if any(a_.startswith("⚠️") for a_ in fv.get("_alerts") or [])
                                                 else ""),
           "Analystes · objectif": ana,
           "Croissance implicite / attendue": reverse_txt(fv),
           "Initiés (12 m)": ins_txt.replace(" · ", " / "),              # « / » : plus lisible dans le tableau
           "Super investors (6 m)": si["Résumé"].replace(" · ", " / ")
           if (si.get("Grille") or f["dataroma"].get("n") is not None) and si["Résumé"] else "n.d.",
           "Put/Call (total / ±10 %)": pcs,
           "BNA N": round(f["eps_this_y"], 2) if f["eps_this_y"] is not None else None,
           "BNA N+1 (évol.)": bna_evol(f["eps_this_y"], f["eps_next_y"]),
           "Dette nette/EBITDA": lev,                                # actuel (dette nette du jour / EBITDA 12 mois)
           "Résultats": f["earnings"],
           "Pages Zonebourse": ("📄 " + (f.get("zb_date") or "")) if f.get("zb") else ""}
    return row, {"fv": fv, "an": an, "si": si, "pc": pc, "ins": ins, "sc": sc, "lev": lev}


DEFAULT_P = {"wacc": 0.09, "tg": 0.025, "gcap": 0.15, "pe_base": 18.0, "pe_min": 10.0, "pe_max": 28.0,
             "pe_cyc": 8.0, "pe_cyc_mid": 11.0, "sbc": True, "coe_bank": 0.11, "pe_bank": 8.5, "pe_bank_us": 12.0,
             "capm": True, "rf": 0.043, "erp": 0.05, "cyc_prem": 0.03, "w_dcf": 0.5}


# ============================================================
#  INTERFACE
# ============================================================
FMT = {
    "Cours": st.column_config.NumberColumn(format="%.2f"),
    "Prix d'entrée": st.column_config.NumberColumn(format="%.2f", help="prix d'entrée du fichier de positions"),
    "JV / prix d'entrée (%)": st.column_config.NumberColumn(format="%+.0f %%", help="(JV centrale / prix d'entrée − 1) : positif = acheté sous la juste valeur"),
    "Écart JV (%)": st.column_config.NumberColumn(format="%+.0f %%", help="(JV centrale / cours − 1) : positif = cours sous la juste valeur"),
    "JV si bénéfices actuels durent": st.column_config.TextColumn(help="Valeurs cycliques (⟳) : la juste valeur centrale est calculée en milieu de cycle (DCF sur la trésorerie moyenne du cycle, multiples sur le BNA moyen publié et prévu). Ici, la même juste valeur si les bénéfices actuels durent (multiples sur le BNA N+1) — l'écart entre les deux mesure le pari sur la durée du cycle. « — » : valeur non cyclique."),
    "Lecture de l'écart": st.column_config.TextColumn(width="large",
        help="⚠️ / ℹ️ en tête : ajustement important du calcul (scission, flux de trésorerie non représentatifs, DCF non "
             "retenu, redressement durable) — détail dans « 🔍 Détail ». Puis, pour une précision 🟠 / 🔴 : pourquoi le "
             "DCF et les multiples divergent (causes probables, les deux plus fortes)."),
    "Précision JV": st.column_config.TextColumn(
        help="Concordance des méthodes retenues (DCF, PER, comparables…) : écart maximal à la JV centrale. "
             "🟢 ≤ 15 % : JV fiable ; 🟠 ≤ 35 % ; 🔴 au-delà : méthodes divergentes, JV incertaine. « DCF écarté » : "
             "méthode à plus de 50 % de la médiane, exclue du calcul — signal de risque (ex. multiples élevés que le "
             "flux de trésorerie ne justifie pas)."),
    "Croissance implicite / attendue": st.column_config.TextColumn(
        help="DCF inversé. Avec états financiers : croissance annuelle moyenne du CA sur 5 ans que suppose le cours / "
             "celle du scénario central (consensus Zonebourse compris). Fiche seule : croissance du flux libre / du BNA "
             "plafonnée. 🟢 attendue ≥ implicite + 5 pts : le marché est pessimiste ; 🔴 le cours suppose plus que "
             "prévu ; ⚪ cohérent."),
    "Initiés (12 m)": st.column_config.TextColumn(width="large", help="Même calcul et mêmes seuils que le screener investors : 🟢 achats des dirigeants sur 6 mois sans vente significative sur 3 mois · 🟠 achats et ventes récentes · 🔴 ventes récentes (3 mois) supérieures aux achats · ⚪ non significatif ; « achats groupés » = ≥ 3 dirigeants en 30 j. Ventes avec levée d'options exclues."),
    "Super investors (6 m)": st.column_config.TextColumn(width="large", help="Rond pondéré par la conviction (fiche Dataroma lue) : chaque mouvement compte pour le poids qu'il représente dans le portefeuille du gérant, en points (5 au plus par fonds) — 🟢 ≥ +0,5 pt · 🔴 ≤ −0,5 pt · 🟠 achats et ventes qui se compensent · ⚪ symbolique ; puis le nombre de fonds : trim. / préc. = ▲ renforcements ▼ allègements sur 6 mois (dernier trimestre et trimestre précédent). Sans fiche lue : même calcul que le screener investors (grilles Dataroma, S&P 500) : rond d'après le dernier trimestre (🟢 plus de fonds qui renforcent, 🔴 plus qui allègent, 🟠 autant), ↗ / ↘ tendance par rapport au trimestre précédent ; trim. / préc. = ▲ renforcements ▼ allègements. Hors S&P 500 : résumé de la fiche Dataroma (dernier mouvement de chaque détenteur). % = poids de la ligne dans le portefeuille du gérant."),
    "Put/Call (total / ±10 %)": st.column_config.TextColumn(help="encours puts / calls ; >1 prudence, <0,7 optimisme"),
    "BNA N": st.column_config.NumberColumn(format="%.2f", help="BNA de l'exercice en cours (consensus)"),
    "BNA N+1 (évol.)": st.column_config.TextColumn(help="BNA prévu pour l'exercice suivant et sa progression ▲ / régression ▼ par rapport à N"),
    "Dette nette/EBITDA": st.column_config.NumberColumn(format="%.1f", help="années d'EBITDA pour rembourser la dette ; négatif = trésorerie nette"),
}


def _base(t):
    """« ARGX (EUR) » -> « ARGX »."""
    return t.split(" ")[0].upper()


def portfolio_facts(raws, zb_store):
    """Valeurs du portefeuille : fiche Finviz, combinée avec les pages Zonebourse déposées pour la même valeur
    (même ticker, même devise) — prévisions sur 3 ans, PER historique, consensus."""
    out = {}
    for t, r in raws.items():
        f = build_facts(r)
        zf = zb_store.get(t.upper())
        if zf and zf.get("zb") and zf.get("currency", "USD") == f.get("currency", "USD"):
            f = merge_facts(f, zf)
            f["source"] = (r.get("source") or "Finviz en ligne") + " + pages Zonebourse du " + (zf.get("zb_date") or "?")
        out[t] = f
    return out


def label(t, f):
    return t + (" ⟳" if f["cyclical"] else "")


def company_name(f, n=32):
    """Nom court de la société (fiche Finviz, sinon titre de la page Zonebourse), sans forme juridique."""
    import html as _h
    nm = _h.unescape(str(f.get("company") or f.get("name") or "")).strip()
    for _ in range(2):                                # « Holdings Inc », « Holdings Co. »
        nm = re.sub(r",?\s+(Inc\.?|Corp\.?|Corporation|Co\.?|Ltd\.?|PLC|plc|N\.V\.|NV|SE|S\.A\.|SA|AG|Holdings?)$",
                    "", nm).strip(" ,")
    return nm[:n] + ("…" if len(nm) > n else "")


def evaluate_memo(name, facts, p, ov, pf=None):
    """evaluate() mémorisé : tant que les valeurs, les réglages et les seuils des signaux ne changent pas, un clic
    ailleurs dans l'application (screener, sélection…) ne recalcule pas toutes les justes valeurs (DCF, Monte Carlo)."""
    import pickle
    try:
        key = hashlib.md5(pickle.dumps((facts, p, ov, pf, _sig_settings(), st.session_state.get("amf_min_sig"),
                                        TODAY.date()))).hexdigest()
    except Exception:  # noqa: BLE001 — objet non sérialisable : pas de mémorisation
        return evaluate(facts, p, ov, pf)
    memo = st.session_state.setdefault("_eval_memo", {})
    if memo.get(name, (None,))[0] != key:
        memo[name] = (key, evaluate(facts, p, ov, pf))
        memo.pop(name + "_xlsx", None)
    return copy.deepcopy(memo[name][1])          # copie : l'affichage ne modifie jamais le résultat mémorisé


def evaluate(facts, p, ov, pf=None):
    """facts : {ticker: faits} -> (tableau, détails, libellé -> ticker)."""
    rows, det = [], {}
    for t, f in facts.items():
        f.setdefault("cyclical_auto", f["cyclical"])
        cy = (ov.get(t) or {}).get("Cyclique", "auto")
        f["cyclical"] = (cy == "oui") if cy in ("oui", "non") else f["cyclical_auto"]
        row, d = analyze(f, p, ov.get(t))
        if pf is not None and not pf.empty:
            entry = None
            if "Prix d'entrée" in pf.columns:
                m = pf["Ticker"] == _base(t)
                if "Devise" in pf.columns:          # même devise que l'analyse (argenx EUR ≠ ADR USD)
                    m &= pf["Devise"] == f.get("currency", "USD")
                if m.any():
                    entry = num(pf.loc[m, "Prix d'entrée"].iloc[0])
            jvc = d["fv"].get("JV centrale")
            row = {"Cours": row.pop("Cours"), "Prix d'entrée": entry,
                   "JV / prix d'entrée (%)": (jvc / entry - 1) * 100 if jvc and entry else None, **row}
        rows.append({"Ticker": label(t, f), "Société": company_name(f), **row})
        det[t] = (f, d)
    df = pd.DataFrame(rows).sort_values("Écart JV (%)", ascending=False, na_position="last")
    if (df["Devise"] == "USD").all():
        df = df.drop(columns="Devise")
    if "Pages Zonebourse" in df and not df["Pages Zonebourse"].astype(bool).any():
        df = df.drop(columns="Pages Zonebourse")
    return df, det, {label(t, det[t][0]): t for t in det}


def dcfx_tables(res):
    """Tableaux du DCF d'entreprise : projection (scénario central), scénarios, CMPC."""
    proj = pd.DataFrame(res["scen"][1]["_rows"]).drop(columns=["_nopat"], errors="ignore")
    for c in ("CA", "EBITDA", "FCFF"):
        proj[c] = proj[c] / 1e3
    proj = proj.rename(columns={"CA": "CA (Md)", "EBITDA": "EBITDA (Md)", "FCFF": "FCFF (Md)"})
    scen = pd.DataFrame([{k: v for k, v in sc.items() if not k.startswith("_")} for sc in res["scen"]])
    w = res["wacc"]
    tot = (w["E"] + w["D"]) or 1
    wacc = pd.DataFrame([
        {"Composante": "Coût des fonds propres", "Valeur (%)": w["ke"] * 100, "Détail": w["how"]},
        {"Composante": "Coût de la dette après impôt", "Valeur (%)": w["kd"] * 100,
         "Détail": f"taux sans risque + {w['spread'] * 100:.1f} pt, impôt {res['x']['tax'] * 100:.0f} %"},
        {"Composante": "Poids des fonds propres", "Valeur (%)": w["E"] / tot * 100, "Détail": "capitalisation boursière"},
        {"Composante": "Poids de la dette", "Valeur (%)": w["D"] / tot * 100, "Détail": "dette financière brute"},
        {"Composante": "CMPC (WACC)", "Valeur (%)": w["wacc"] * 100, "Détail": "taux d'actualisation des flux"}])
    return proj, scen, wacc


def render_dcfx(f, res):
    cur = {"USD": "$", "EUR": "€"}.get(f.get("currency"), f.get("currency", "") + " ")
    st.markdown("**DCF d'entreprise** — flux de trésorerie disponibles pour l'entreprise (FCFF), actualisés au CMPC")
    c = st.columns(4)
    c[0].metric("DCF (scénarios pondérés)", f"{res['central']:.0f} {cur}")
    c[1].metric("Fourchette P10 – P90", f"{res['p10']:.0f} – {res['p90']:.0f}")
    c[2].metric("Probabilité > cours", f"{res['prob_up']:.0f} %" if res.get("prob_up") is not None else "n.d.")
    c[3].metric("CMPC", f"{res['wacc']['wacc'] * 100:.1f} %")
    proj, scen, wacc = dcfx_tables(res)
    st.dataframe(scen.round(1), hide_index=True, width="stretch")
    x = res["x"]
    st.caption(f"Hypothèses ({res['src']}) : CA de départ {x['rev0'] / 1e3:.1f} Md, croissance an 1 "
               f"{res['g_rev'] * 100:.1f} % décroissant jusqu'à {res['tg'] * 100:.1f} % en année 10"
               + (" (consensus Zonebourse pour les premières années)" if x.get("fc") else "")
               + f" ; marge EBITDA {x['m0'] * 100:.1f} % (médiane historique "
               f"{(_med(x['m_hist']) or x['m0']) * 100:.1f} %) ; capex {x['capex0'] * 100:.1f} % du CA → niveau "
               f"d'entretien ; impôt {x['tax'] * 100:.0f} % ; rémunération en actions déduite. Valeur terminale : "
               f"moyenne croissance perpétuelle et multiple de sortie {res['exit_mult']:.1f}x EBITDA. "
               f"Valeur des fonds propres = valeur d'entreprise − dette nette ({x['net_debt'] / 1e3:.1f} Md) "
               f"− minoritaires. Pessimiste / optimiste : croissance ± 4 pts, marge ± 2 pts (creux de cycle si "
               f"cyclique) ; pondération 25 / 50 / 25 ; P10–P90 : 2 000 tirages (croissance, marge, CMPC, "
               f"croissance perpétuelle, multiple).")
    with st.expander("Projection du scénario central (10 ans)"):
        st.dataframe(proj.round(2), hide_index=True, width="stretch")
    with st.expander("CMPC (WACC)"):
        st.dataframe(wacc.round(2), hide_index=True, width="stretch")
    with st.expander("Sensibilité : CMPC × croissance perpétuelle (scénario central, valeur / action)"):
        ws = [res["wacc"]["wacc"] + d_ for d_ in (-0.02, -0.01, 0.0, 0.01, 0.02)]
        tgs = [0.015, 0.02, 0.025, 0.03, 0.035]
        grid = {}
        for t_ in tgs:
            col = []
            for w_ in ws:
                if w_ <= t_ + 0.01:
                    col.append(None)
                    continue
                r = dcfx_project(x, res["g_rev"], 0.0, w_, t_, res["exit_mult"], bool(f.get("cyclical")))
                col.append((r["eq_perp"] + r["eq_exit"]) / 2 / f["shares"])
            grid[f"g {t_ * 100:.1f} %"] = col
        st.dataframe(pd.DataFrame(grid, index=[f"CMPC {w_ * 100:.1f} %" for w_ in ws]).round(0))
        st.caption(f"Cours : {f['price']:.2f}. Moyenne des deux valeurs terminales.")


METHODS = [("DCF", "DCF"), ("PER", "PER"), ("Comparables", "Comparables"), ("PER hist.", "PER historique"),
           ("P/B", "Actif net (P/B)")]


def render_fv_tables(fv, cur):
    """Juste valeur en deux tableaux : 1) valeur par action de chaque méthode (avec le multiple ou le taux utilisé)
    et synthèse ; 2) indicateurs (taux, croissances, précision)."""
    keep = set(fv.get("_keep") or [])
    mult = fv.get("_mult") or {}
    rows = []
    for k, lbl in METHODS:
        if fv.get(k) is None:
            continue
        hyp = mult.get(k, "")
        if k == "DCF":
            hyp = (f"CMPC {fv['CMPC (%)']:.1f} %" if fv.get("CMPC (%)") else
                   f"coût des fonds propres {fv['Coût fonds propres (%)']:.1f} %" if fv.get("Coût fonds propres (%)") else "")
        rows.append({"Méthode": lbl, "Hypothèse": hyp, f"Basse ({cur})": fv.get(k + " bas"),
                     f"Centrale ({cur})": fv[k], f"Haute ({cur})": fv.get(k + " haut"),
                     "Retenue": "✓" if not keep or k in keep else "écartée"})
    if fv.get("JV centrale"):
        rows.append({"Méthode": "➜ Juste valeur" + (" (milieu de cycle)" if fv.get("JV si pic") else ""),
                     "Hypothèse": "", f"Basse ({cur})": fv.get("JV basse"),
                     f"Centrale ({cur})": fv["JV centrale"], f"Haute ({cur})": fv.get("JV haute"), "Retenue": ""})
    if fv.get("JV si pic"):
        rows.append({"Méthode": "Si les bénéfices actuels durent", "Hypothèse": f"multiples sur le BNA N+1 "
                     f"{fv['_eps_peak']:.2f}", f"Basse ({cur})": fv.get("JV si pic basse"),
                     f"Centrale ({cur})": fv["JV si pic"], f"Haute ({cur})": fv.get("JV si pic haute"),
                     "Retenue": "information"})
    if rows:
        st.dataframe(pd.DataFrame(rows).round(1), hide_index=True, width="stretch")
        st.caption(f"Valeurs par action en {cur.strip()} (pas des ratios) : multiple ou taux utilisé dans « Hypothèse »."
                   + (f" Médiane des multiples : {fv['Multiples (médiane)']:.1f} ; juste valeur = "
                      f"{fv['_w_dcf'] * 100:.0f} % DCF + {(1 - fv['_w_dcf']) * 100:.0f} % multiples (réglable)."
                      if fv.get("Multiples (médiane)") and fv.get("_w_dcf") is not None else
                      " Juste valeur = médiane des méthodes (moyenne s'il n'y en a que deux)." if fv.get("JV centrale") else ""))
    skip = {m for k, _ in METHODS for m in (k, k + " bas", k + " haut")} | {
        "JV centrale", "JV basse", "JV haute", "Multiples (médiane)", "notes", "JV si pic", "JV si pic basse",
        "JV si pic haute"}
    ind = {k: v for k, v in fv.items() if k not in skip and not k.startswith("_") and isinstance(v, (int, float))}
    if ind:
        st.dataframe(pd.DataFrame([ind]).round(1), hide_index=True)


def _md(x):
    """Texte affiché en Markdown : « $ » échappé (sinon Streamlit l'interprète comme une formule)."""
    return str(x).replace("$", "\\$")


def short_source(f):
    """Source lisible : noms de fichiers Zonebourse remplacés par « Zonebourse (pages du …) »."""
    src = str(f.get("source") or ("Zonebourse" if f.get("zb") and not f.get("dcfx", {}) is None else "Finviz en ligne"))
    if not f.get("source") and f.get("zb"):
        src = "Zonebourse"
    src = re.sub(r"Zonebourse\s*:.*$", "Zonebourse", src)
    if f.get("zb_date") and "Zonebourse" in src and f["zb_date"] not in src:
        src = src.replace("Zonebourse", f"Zonebourse (pages du {f['zb_date']})", 1)
    return src


def _eps_series(f):
    """BNA : historique publié + consensus -> [(exercice, BNA, « publié » / « consensus »)]."""
    z = f.get("zb")
    out = []
    if z and z.get("bna"):
        N, L = z.get("N"), z.get("L")
        for y in sorted(z["bna"]):
            v = z["bna"][y]
            if v is None or (L and y < L - 4):
                continue
            out.append((str(y), v, "publié" if L and y <= L else "consensus"))
        return out
    for lbl, v in (f.get("eps_hist_long") or (f.get("eps_hist") or [])[-5:]):
        out.append((lbl, v, "publié"))
    if f.get("eps_this_y") is not None:
        out.append(("N (en cours)", f["eps_this_y"], "consensus"))
    if f.get("eps_next_y") is not None:
        out.append(("N+1", f["eps_next_y"], "consensus"))
    return out


def _cagr(a, b, n):
    return ((b / a) ** (1 / n) - 1) * 100 if a and b and a > 0 and b > 0 and n > 0 else None


def en_bref(f, d, cur):
    """3 à 4 phrases : juste valeur, fourchette, potentiel, écart avec les analystes, principal point d'attention."""
    fv, an = d["fv"], d["an"]
    jv, pr = fv.get("JV centrale"), f.get("price")
    if not jv or not pr:
        return "Juste valeur non calculable avec les données disponibles."
    s = [f"Juste valeur **{jv:.0f} {cur}** (fourchette {fv['JV basse']:.0f}–{fv['JV haute']:.0f}), soit "
         f"**{(jv / pr - 1) * 100:+.0f} %** par rapport au cours de {pr:.2f} {cur}."]
    keep = [lbl for k, lbl in METHODS if k in (fv.get("_keep") or [])]
    if keep:
        s.append(f"Méthodes retenues : {', '.join(keep)} ; précision {precision_txt(fv)}.")
    if f.get("target"):
        e = (jv / f["target"] - 1) * 100
        s.append(f"Objectif moyen des analystes : {f['target']:.0f} {cur} ({an.get('Consensus') or 'n.d.'}) — la juste "
                 f"valeur est {'au-dessus' if e > 0 else 'en dessous'} de {abs(e):.0f} %"
                 + (" : écart expliqué en section 7." if abs(e) > 10 else "."))
    al = fv.get("_alerts") or []
    if al:
        s.append("À noter : " + " ; ".join(a.lstrip("⚠️ℹ️ ").strip() for a in al) + ".")
    elif f.get("cyclical"):
        s.append("Valeur classée cyclique : calcul prudent (bénéfices possiblement proches d'un pic).")
    if fv.get("JV si pic") and f.get("price"):
        s.append(f"Valeur cyclique : juste valeur de milieu de cycle (BNA moyen {fv['_eps_mid']:.2f}) ; si les bénéfices "
                 f"actuels durent (BNA N+1 {fv['_eps_peak']:.2f}), elle serait de {fv['JV si pic']:.0f} {cur.strip()} "
                 f"({(fv['JV si pic'] / f['price'] - 1) * 100:+.0f} % par rapport au cours).")
    return " ".join(s)


def risk_points(f, d):
    """Points de vigilance et points favorables, déduits des données."""
    fv, ins, si = d["fv"], d["ins"], d["si"]
    bad, good = [], []
    if f.get("cyclical"):
        bad.append("valeur cyclique : multiples appliqués à des bénéfices qui peuvent être proches d'un pic")
    lev = d.get("lev")
    if lev is not None and lev > 3 and not f.get("bank"):
        bad.append(f"endettement élevé : dette nette = {lev:.1f}× l'EBITDA")
    tv = fv.get("Valeur terminale (% VE)") or fv.get("Valeur terminale (% du DCF)")
    if tv and tv > 75:
        bad.append(f"valeur très dépendante du long terme : valeur terminale = {tv:.0f} % de la valeur du DCF")
    gi, ga = fv.get("g implicite (%)"), fv.get("g attendue (%)")
    if gi is not None and ga is not None and gi - ga >= 8:
        bad.append(f"le cours suppose une croissance de {gi:.0f} %/an, contre {ga:.0f} % retenus")
    if fv.get("Dispersion (%)") and fv["Dispersion (%)"] > 35:
        bad.append(f"méthodes divergentes (±{fv['Dispersion (%)']:.0f} %) : juste valeur incertaine")
    if f.get("short") and f["short"] > 10:
        bad.append(f"vente à découvert élevée ({f['short']:.1f} % du flottant)")
    if f.get("target") and f.get("eps_next_y") and f["eps_next_y"] > 0 and f["target"] / f["eps_next_y"] > 35:
        bad.append(f"l'objectif des analystes suppose un PER de {f['target'] / f['eps_next_y']:.0f}x le BNA N+1")
    for a_ in fv.get("_alerts") or []:
        (bad if a_.startswith("⚠️") else good if "redressement" in a_ else bad).append(a_.lstrip("⚠️ℹ️ ").strip())
    if ins and ins.get("Sortie dirigeant"):
        bad.append("dirigeant de premier rang sorti après levée d'options : " + ins["Sortie dirigeant"])
    if ins and "🚨" in (ins.get("Alerte") or ""):
        bad.append("vente d'un dirigeant dans les 7 derniers jours : " + ins["Alerte"])
    if ins and "✅" in (ins.get("Alerte") or ""):
        good.append("achat d'un dirigeant dans les 7 derniers jours : " + ins["Alerte"])
    if ins and str(ins.get("Signal", "")).startswith("🟢"):
        good.append("achats des dirigeants : " + ins["Résumé"])
    elif ins and str(ins.get("Signal", "")).startswith(("🟠", "🔴")):
        bad.append("ventes récentes des dirigeants : " + ins["Résumé"])
    if si and (si.get("Lignes") or si.get("Grille")) and str(si.get("Signal", "")).startswith("🟢"):
        good.append("super investors : " + si["Résumé"])
    elif si and (si.get("Lignes") or si.get("Grille")) and str(si.get("Signal", "")).startswith("🔴"):
        bad.append("super investors : " + si["Résumé"])
    if fv.get("Écart (%)") is not None and fv["Écart (%)"] > 15 and precision_txt(fv).startswith("🟢"):
        good.append(f"juste valeur supérieure au cours de {fv['Écart (%)']:.0f} % avec des méthodes concordantes")
    if fv.get("Probabilité JV DCF > cours (%)") is not None and fv["Probabilité JV DCF > cours (%)"] >= 75:
        good.append(f"{fv['Probabilité JV DCF > cours (%)']:.0f} % des simulations du DCF dépassent le cours")
    return bad[:5], good[:4]


def football_field(fv, f, cur):
    """Graphique « terrain de football » : une barre basse–haute par méthode, la juste valeur mise en avant,
    le cours et l'objectif des analystes en lignes de référence."""
    import altair as alt
    rows = [{"Méthode": lbl, "Basse": fv.get(k + " bas"), "Haute": fv.get(k + " haut"), "Centrale": fv[k]}
            for k, lbl in METHODS if fv.get(k) is not None and fv.get(k + " bas") is not None]
    if fv.get("JV centrale"):
        rows.append({"Méthode": "➜ Juste valeur", "Basse": fv["JV basse"], "Haute": fv["JV haute"],
                     "Centrale": fv["JV centrale"]})
    if len(rows) < 2:
        return
    df = pd.DataFrame(rows)
    order = list(df["Méthode"])
    base = alt.Chart(df).encode(y=alt.Y("Méthode:N", sort=order, title=None,
                                        axis=alt.Axis(labelOverlap=False, labelLimit=200, labelFontSize=12)))
    bars = base.mark_bar(size=16, cornerRadius=3).encode(
        x=alt.X("Basse:Q", title=f"Valeur par action ({cur.strip()})"), x2="Haute:Q",
        color=alt.condition(alt.datum["Méthode"] == "➜ Juste valeur", alt.value("#2e7d32"), alt.value("#90a4ae")),
        tooltip=["Méthode", alt.Tooltip("Basse:Q", format=".0f"), alt.Tooltip("Centrale:Q", format=".0f"),
                 alt.Tooltip("Haute:Q", format=".0f")])
    ticks = base.mark_tick(color="black", thickness=2, size=18).encode(x="Centrale:Q")
    refs = [{"Repère": f"Cours {f['price']:.0f}", "x": f["price"]}] if f.get("price") else []
    if f.get("target"):
        refs.append({"Repère": f"Objectif analystes {f['target']:.0f}", "x": f["target"]})
    layers = [bars, ticks]
    if refs:
        rdf = pd.DataFrame(refs)
        layers.append(alt.Chart(rdf).mark_rule(strokeDash=[5, 4], size=2).encode(
            x="x:Q", color=alt.Color("Repère:N", scale=alt.Scale(range=["#d32f2f", "#1565c0"]),
                                     legend=alt.Legend(orient="bottom", title=None))))
    st.altair_chart(alt.layer(*layers).properties(height=48 * len(rows) + 30), width="stretch", theme=None)


# ---------- Export de la fiche détaillée (PDF / Word)
_DOTS = {"🟢": ("\x01", "2e7d32"), "🟠": ("\x02", "ef6c00"), "🔴": ("\x03", "c62828"), "⚪": ("\x04", "9e9e9e")}
_DOT_COL = {v[0]: v[1] for v in _DOTS.values()}
_EMO = {"🟢🟢": "\x01\x01", "⚠️": "Attention :",
        "ℹ️": "Info :", "✅": "", "⟳": "(cyclique)", "🏦": "(modèle financier)", "➜": "->", "→": "->", "✓": "oui",
        "≥": ">=", "≤": "<=", "▲": "+", "▼": "-", "💱": "", "📄": "", " ": " ", "\xa0": " ", "−": "-", "😨": ""}


def _clean(x):
    """Texte sans émoji ni caractère hors des polices standard des PDF ; les pastilles de couleur deviennent des
    marqueurs (\x01…\x04) que les rendus PDF et Word dessinent en ronds colorés."""
    s = "" if x is None else str(x)
    for k, v in _EMO.items():
        s = s.replace(k, v)
    for k, (m_, _) in _DOTS.items():
        s = s.replace(k, m_)
    s = s.replace("**", "")
    s = s.encode("cp1252", "ignore").decode("cp1252")
    return re.sub(r"[ \t]{2,}", " ", s).strip()


def _fmt(v, nd=1):
    if v is None or (isinstance(v, float) and v != v):
        return ""
    if isinstance(v, (int, float)):
        return f"{v:,.{nd}f}".replace(",", " ")
    return _clean(v)


def report_model(f, d, t):
    """Contenu de la fiche (mêmes sections que l'onglet Détail) : liste de blocs
    ("h", texte) · ("p", texte) · ("ul", [textes]) · ("table", en-têtes, lignes) · ("img", png)."""
    fv, an = d["fv"], d["an"]
    cur = {"USD": "$", "EUR": "€"}.get(f.get("currency"), f.get("currency", "") + " ").strip()
    B = []
    name = company_name(f, 80) or t
    B.append(("title", f"{name} ({t})" if name != t else t))
    B.append(("p", f"Analyse du {TODAY:%d/%m/%Y} — {short_source(f)}"))
    # 1
    B.append(("h", "1. Fiche d'identité"))
    lev = d.get("lev")
    rows = [["Cours", f"{f['price']:.2f} {cur}" if f.get("price") else "n.d."],
            ["Capitalisation", f"{f['mcap'] / 1e3:,.1f} Md {cur}".replace(",", " ") if f.get("mcap") else "n.d."],
            ["Dette nette / EBITDA", "sans objet (banque / assurance)" if f.get("bank") else
             "trésorerie nette" if lev is not None and lev < 0 else f"{lev:.1f}x" if lev is not None else "n.d."],
            ["Bêta", f"{f['beta']:.2f}" if f.get("beta") else "n.d."],
            ["Dividende / action", f"{f['div_ps']:.2f} {cur}" if f.get("div_ps") else "aucun"],
            ["Secteur / industrie", " · ".join(x for x in (f.get("sector"), f.get("industry")
                                                          or (f.get("zb_sector") or "").replace("-", " ")) if x) or "n.d."],
            ["Prochains résultats", f.get("earnings") or "n.d."],
            ["Classement", ", ".join(x for x in ("cyclique" if f.get("cyclical") else "",
                                                  "modèle financier (banque / assurance)" if f.get("bank") else "") if x)
             or "standard"]]
    B.append(("table", ["Élément", "Valeur"], rows))
    # 2
    B.append(("h", "2. En bref"))
    B.append(("p", en_bref(f, d, cur)))
    if fv.get("_alerts"):
        B.append(("ul", list(fv["_alerts"])))
    # 3
    B.append(("h", "3. Bénéfice par action (BNA)"))
    ser = _eps_series(f)
    if ser:
        png = _png_eps(ser, cur)
        if png:
            B.append(("img", png))
        B.append(("table", ["Exercice", f"BNA ({cur})", "Type"], [[a, _fmt(b, 2), c] for a, b, c in ser]))
        B.append(("p", "Historique : BNA comptable publié. Prévisions : consensus des analystes, en général sur le BNA "
                       "non-GAAP (ajusté), utilisé par le PER cible et les comparables."))
    else:
        B.append(("p", "Non disponible pour cette source."))
    # 4
    B.append(("h", "4. DCF"))
    res = fv.get("_dcfx")
    if res:
        x = res["x"]
        B.append(("table", ["DCF (scénarios pondérés)", "Fourchette P10 – P90", "Probabilité > cours", "CMPC"],
                  [[f"{res['central']:.0f} {cur}", f"{res['p10']:.0f} – {res['p90']:.0f}",
                    f"{res['prob_up']:.0f} %" if res.get("prob_up") is not None else "n.d.",
                    f"{res['wacc']['wacc'] * 100:.1f} %"]]))
        B.append(("table", ["Scénario", "Poids", "Croissance CA an 1 (%)", "Écart de marge (pts)", f"Valeur / action ({cur})",
                            "Valeur terminale (% VE)"],
                  [[s["Scénario"], _fmt(s["Poids"], 2), _fmt(s["Croissance CA an 1 (%)"]), _fmt(s["Écart de marge (pts)"]),
                    _fmt(s["Valeur / action"]), _fmt(s["Valeur terminale (% VE)"], 0)] for s in res["scen"]]))
        B.append(("p", f"Hypothèses ({res['src']}) : CA de départ {x['rev0'] / 1e3:.1f} Md, croissance an 1 "
                       f"{res['g_rev'] * 100:.1f} % décroissant jusqu'à {res['tg'] * 100:.1f} % en année 10 ; marge EBITDA "
                       f"{x['m0'] * 100:.1f} % ; impôt {x['tax'] * 100:.0f} % ; valeur terminale : moyenne croissance "
                       f"perpétuelle et multiple de sortie {res['exit_mult']:.1f}x EBITDA ; dette nette "
                       f"{x['net_debt'] / 1e3:.1f} Md déduite."))
    elif fv.get("_dcf"):
        fcf, g0, ke, sh = fv["_dcf"]
        B.append(("p", f"DCF simplifié (fiche seule) : flux libre {fcf / 1e3:.1f} Md, croissance {g0 * 100:.0f} %, coût des "
                       f"fonds propres {ke * 100:.1f} % -> {fv['DCF']:.0f} {cur} (fourchette {fv['DCF bas']:.0f}–"
                       f"{fv['DCF haut']:.0f})."))
    else:
        why_ = [n for n in fv["notes"] if "DCF" in n or "flux" in n or "assureur" in n]
        B.append(("p", "Pas de DCF retenu" + (" : " + " · ".join(why_) if why_ else ".")))
    # 5
    B.append(("h", "5. PER et multiples"))
    mult = fv.get("_mult") or {}
    rows = [[lbl, mult.get(k, ""), _fmt(fv[k]), "oui" if not fv.get("_keep") or k in fv["_keep"] else "écartée"]
            for k, lbl in METHODS if k != "DCF" and fv.get(k) is not None]
    if rows:
        B.append(("table", ["Méthode", "Hypothèse", f"Valeur ({cur})", "Retenue"], rows))
    if fv.get("_peers_txt"):
        B.append(("p", "Par rapport aux pairs : " + fv["_peers_txt"] + "."))
    m = [s for s in (f"PER actuel sur N+1 {f['fwd_pe']:.1f}x" if f.get("fwd_pe") else "",
                     f"PER historique médian {f['pe_hist']:.1f}x" if f.get("pe_hist") else "") if s]
    if m:
        B.append(("p", " · ".join(m) + "."))
    # 6
    B.append(("h", "6. Synthèse"))
    png = _png_football(fv, f, cur)
    if png:
        B.append(("img", png))
    keep = set(fv.get("_keep") or [])
    rows = [[lbl, (f"CMPC {fv['CMPC (%)']:.1f} %" if k == "DCF" and fv.get("CMPC (%)") else mult.get(k, "")),
             _fmt(fv.get(k + " bas")), _fmt(fv[k]), _fmt(fv.get(k + " haut")), "oui" if not keep or k in keep else "écartée"]
            for k, lbl in METHODS if fv.get(k) is not None]
    if fv.get("JV centrale"):
        rows.append(["-> Juste valeur" + (" (milieu de cycle)" if fv.get("JV si pic") else ""), "",
                     _fmt(fv["JV basse"]), _fmt(fv["JV centrale"]), _fmt(fv["JV haute"]), ""])
    if fv.get("JV si pic"):
        rows.append(["Si les bénéfices actuels durent", f"BNA N+1 {fv['_eps_peak']:.2f}", _fmt(fv.get("JV si pic basse")),
                     _fmt(fv["JV si pic"]), _fmt(fv.get("JV si pic haute")), "information"])
    if rows:
        B.append(("table", ["Méthode", "Hypothèse", f"Basse ({cur})", f"Centrale ({cur})", f"Haute ({cur})", "Retenue"], rows))
        B.append(("p", f"Valeurs par action en {cur}. Précision : {precision_txt(fv)}. Croissance implicite / attendue : "
                       f"{reverse_txt(fv)}."))
    # 7
    B.append(("h", "7. Juste valeur calculée et objectif des analystes"))
    jv = fv.get("JV centrale")
    if jv and f.get("target"):
        e = (jv / f["target"] - 1) * 100
        ast_ = analyst_style_target(f)
        rows_ = [["Juste valeur calculée", _fmt(jv), _fmt(fv.get("Écart (%)"), 0)],
                 [f"Objectif moyen des analystes ({f.get('n_analysts') or '?'} analystes)", _fmt(f["target"]),
                  _fmt(an.get("Potentiel (%)"), 0)]]
        if ast_ and not ast_.get("skip") and f.get("price"):
            rows_.append(["Objectif à 12 mois façon analystes (VE/EBIT moyen 5 ans × EBIT N+1)", _fmt(ast_["target"]),
                          _fmt((ast_["target"] / f["price"] - 1) * 100, 0)])
        B.append(("table", ["Repère", f"Valeur ({cur})", "vs cours (%)"], rows_))
        if ast_:
            B.append(("p", analyst_style_txt(ast_, f, cur)))
        br_ = target_bridge_txt(target_bridge(f, fv), f, cur, fv.get("JV si pic"))
        if br_:
            B.append(("p", br_))
        vm_ = value_markers(f, fv, cur)
        if vm_:
            B.append(("table", ["Repère", "Valeur", "Lecture"], vm_))
        pr_ = premium_reasons(f, fv)
        if pr_:
            B.append(("p", "Pourquoi cette prime ? " + premium_intro(pr_)))
            B.append(("table", ["", "Critère", "Valeur", "Lecture"],
                      [[ICON_P[r[3]], r[0], r[1], r[2]] for r in pr_["rows"]]))
            B.append(("p", pr_["verdict"] + " " + PREMIUM_NOTE))
        txt = f"Écart juste valeur / objectif : {e:+.0f} %."
        if f.get("eps_next_y") and f["eps_next_y"] > 0:
            txt += f" L'objectif correspond à un PER de {f['target'] / f['eps_next_y']:.1f}x le BNA N+1."
        B.append(("p", txt))
        causes = gap_causes(f, fv, d.get("lev")) if abs(e) > 10 else []
        if causes:
            B.append(("p", "Raisons probables de l'écart :"))
            B.append(("ul", causes))
    else:
        B.append(("p", "Objectif des analystes non disponible."))
    # 8
    B.append(("h", "8. Risques et points de vigilance"))
    bad, good = risk_points(f, d)
    if bad or good:
        B.append(("ul", [f"Vigilance : {x}" for x in bad] + [f"Favorable : {x}" for x in good]))
    else:
        B.append(("p", "Aucun point particulier détecté dans les données."))
    # 9
    B.append(("h", "9. Signaux de marché"))
    ins = d.get("ins")
    if ins:
        B.append(("p", f"Initiés ({ins.get('Source') or f.get('insider_source', 'fiche Finviz')}) : {ins['Résumé']}"))
        keys = ("Achats 365j", "Ventes 365j", "Acheteurs", "Achats groupés (max en 30 j)", "Achats PDG / DAF (M$)",
                "Hausse médiane de détention (%)", "Ventes avec levée d'options (M$)")
        B.append(("table", ["Initiés, 12 mois", "Valeur"], [[k.replace("M$", "M€") if ins.get("Devise") == "€" else k,
                                                             _fmt(ins.get(k))] for k in keys if ins.get(k) is not None]))
        if ins.get("Achats notables"):
            B.append(("ul", ["Achat : " + x for x in ins["Achats notables"]]))
    else:
        B.append(("p", "Initiés : non disponible pour cette source."))
    si = d.get("si") or {}
    B.append(("p", "Super investors (Dataroma) : " + (si.get("Résumé") if si.get("Lignes") or si.get("Grille")
                                                      else "aucune donnée.")))
    nb = f" · {f['n_analysts']} analystes" if f.get("n_analysts") else ""
    B.append(("p", f"Analystes : consensus {an.get('Consensus') or 'n.d.'}{nb}"
                   + (f", {an['Relèvements 180j']} relèvement(s) / {an['Abaissements 180j']} abaissement(s) sur 180 j"
                      if f.get("ratings") else "")))
    if an.get("Actions récentes"):
        ar = pd.DataFrame(an["Actions récentes"]).head(8)
        B.append(("table", [str(c) for c in ar.columns], [[_fmt(v) for v in r] for r in ar.itertuples(index=False)]))
    op = pd.DataFrame(f.get("options") or [])
    if not op.empty:
        op["P/C"] = (op["OI puts"].astype(float) / op["OI calls"].astype(float).replace(0, float("nan"))).round(2)
        B.append(("p", f"Options par échéance — P/C encours {d['pc'].get('P/C encours', 'n.d.')}"
                       + (f" ({d['pc']['Lecture P/C']})" if d["pc"].get("Lecture P/C") else "")))
        B.append(("table", [str(c) for c in op.columns], [[_fmt(v, 2) for v in r] for r in op.itertuples(index=False)]))
    # 10
    B.append(("h", "11. Actualités (7 derniers jours)"))
    try:
        det_, ns_ = news_recent(dict(f, ticker=f.get("ticker") or t))
    except Exception:
        det_, ns_ = None, None
    if det_ is not None and len(det_):
        B.append(("p", f"{ns_['n7']} titres sur 7 jours ({ns_['pos7']} positifs, {ns_['neg7']} négatifs) ; ton moyen "
                       f"7 j {ns_['s7']} / 30 j {ns_['s30']} (de −100 à +100). Source : {ns_['src']}."))
        sh = det_.sort_values("Date", ascending=False).head(15)
        B.append(("table", ["Date", "Ton", "Événement", "Titre", "Source"],
                  [[r_.Date.strftime("%d/%m"), f"{r_.Ton} ({r_.Score:+d})", r_.Événement, r_.Titre, r_.Source]
                   for r_ in sh.itertuples()]))
    else:
        B.append(("p", "Actualités non disponibles pour cette source."))
    B.append(("h", "10. Notes de calcul et sources"))
    B.append(("ul", list(fv["notes"])))
    B.append(("p", f"Sources : {short_source(f)}. Estimation argumentée, pas un conseil en investissement."))
    return B


def _png_eps(ser, cur):
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        return None
    fig, ax = plt.subplots(figsize=(7.2, 2.6), dpi=150)
    col = ["#607d8b" if t == "publié" else "#43a047" for _, _, t in ser]
    ax.bar([a for a, _, _ in ser], [b for _, b, _ in ser], color=col)
    ax.axhline(0, color="#999", lw=0.8)
    ax.set_ylabel(f"BNA ({cur})", fontsize=8)
    ax.tick_params(labelsize=7)
    for s_ in ("top", "right"):
        ax.spines[s_].set_visible(False)
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(color="#607d8b", label="publié"), Patch(color="#43a047", label="consensus")],
              fontsize=7, frameon=False, loc="upper left")
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    return buf.getvalue()


def _png_football(fv, f, cur):
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        return None
    rows = [(lbl, fv.get(k + " bas"), fv.get(k + " haut"), fv[k]) for k, lbl in METHODS
            if fv.get(k) is not None and fv.get(k + " bas") is not None]
    if fv.get("JV centrale"):
        rows.append(("Juste valeur", fv["JV basse"], fv["JV haute"], fv["JV centrale"]))
    if len(rows) < 2:
        return None
    fig, ax = plt.subplots(figsize=(7.2, 0.55 * len(rows) + 1.0), dpi=150)
    for i, (lbl, lo, hi, c) in enumerate(reversed(rows)):
        ax.barh(i, hi - lo, left=lo, height=0.5, color="#2e7d32" if lbl == "Juste valeur" else "#90a4ae")
        ax.plot([c, c], [i - 0.3, i + 0.3], color="black", lw=2)
    ax.set_yticks(range(len(rows)))
    ax.set_yticklabels([r[0] for r in reversed(rows)], fontsize=8)
    if f.get("price"):
        ax.axvline(f["price"], color="#d32f2f", ls="--", lw=1.5, label=f"Cours {f['price']:.0f}")
    if f.get("target"):
        ax.axvline(f["target"], color="#1565c0", ls="--", lw=1.5, label=f"Objectif analystes {f['target']:.0f}")
    ax.set_xlabel(f"Valeur par action ({cur})", fontsize=8)
    ax.tick_params(axis="x", labelsize=7)
    ax.legend(fontsize=7, frameon=False, loc="lower right")
    for s_ in ("top", "right"):
        ax.spines[s_].set_visible(False)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    return buf.getvalue()


def _docx_runs(par, text, size=None, bold=False):
    """Ajoute le texte à un paragraphe Word, les marqueurs de pastille devenant des ronds colorés."""
    from docx.shared import Pt, RGBColor
    for part in re.split(r"([\x01-\x04])", _clean(text)):
        if not part:
            continue
        r_ = par.add_run("●" if part in _DOT_COL else part)
        if part in _DOT_COL:
            r_.font.color.rgb = RGBColor.from_string(_DOT_COL[part].upper())
        if size:
            r_.font.size = Pt(size)
        r_.font.bold = bold
    return par


def report_docx(blocks):
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(1.8)
    sec.top_margin = sec.bottom_margin = Cm(1.6)
    st_ = doc.styles["Normal"]
    st_.font.name = "Calibri"
    st_.font.size = Pt(10)
    for b in blocks:
        k = b[0]
        if k == "title":
            doc.add_heading(re.sub(r"[\x01-\x04]", "", _clean(b[1])), level=0)
        elif k == "h":
            doc.add_heading(re.sub(r"[\x01-\x04]", "", _clean(b[1])), level=1)
        elif k == "p":
            _docx_runs(doc.add_paragraph(), b[1])
        elif k == "ul":
            for it in b[1]:
                _docx_runs(doc.add_paragraph(style="List Bullet"), it)
        elif k == "img":
            doc.add_picture(io.BytesIO(b[1]), width=Cm(17))
        elif k == "table":
            head, rows = b[1], b[2]
            tb = doc.add_table(rows=1, cols=len(head))
            tb.style = "Light Grid Accent 1"
            for i, h in enumerate(head):
                _docx_runs(tb.rows[0].cells[i].paragraphs[0], h, 8.5, True)
            for row in rows:
                cells = tb.add_row().cells
                for i, v in enumerate(row[:len(head)]):
                    _docx_runs(cells[i].paragraphs[0], v, 8.5)
            doc.add_paragraph()
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def report_pdf(blocks):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, ListFlowable, ListItem
    from xml.sax.saxutils import escape
    ss = getSampleStyleSheet()
    body = ParagraphStyle("b", parent=ss["Normal"], fontName="Helvetica", fontSize=9, leading=12, spaceAfter=4)
    cellst = ParagraphStyle("c", parent=body, fontSize=7.5, leading=9, spaceAfter=0)
    h1 = ParagraphStyle("h1", parent=ss["Heading2"], fontName="Helvetica-Bold", fontSize=12, spaceBefore=10, spaceAfter=4,
                        textColor=colors.HexColor("#1f3b57"))
    tit = ParagraphStyle("t", parent=ss["Title"], fontName="Helvetica-Bold", fontSize=17)
    # rond coloré : police DejaVu Sans fournie avec matplotlib (« ● »), sinon puce standard « • »
    dot_font, dot_char = "Helvetica", "•"
    try:
        import matplotlib
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        ttf = os.path.join(matplotlib.get_data_path(), "fonts", "ttf", "DejaVuSans.ttf")
        if os.path.exists(ttf):
            if "DejaVuDot" not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont("DejaVuDot", ttf))
            dot_font, dot_char = "DejaVuDot", "●"
    except Exception:
        pass

    def _pdf_txt(s_):
        t_ = escape(_clean(s_))
        for m_, col in _DOT_COL.items():
            t_ = t_.replace(m_, f'<font name="{dot_font}" color="#{col}">{dot_char}</font>')
        return t_
    P = lambda s_, st=body: Paragraph(_pdf_txt(s_), st)
    W = A4[0] - 3.4 * cm
    story = []
    for b in blocks:
        k = b[0]
        if k == "title":
            story.append(P(b[1], tit))
        elif k == "h":
            story.append(P(b[1], h1))
        elif k == "p":
            story.append(P(b[1]))
        elif k == "ul":
            story.append(ListFlowable([ListItem(P(x), leftIndent=10) for x in b[1]], bulletType="bullet",
                                      start="•", leftIndent=12))
        elif k == "img":
            img = Image(io.BytesIO(b[1]))
            r = W / img.drawWidth
            img.drawWidth, img.drawHeight = W, img.drawHeight * r
            story.append(img)
        elif k == "table":
            head, rows = b[1], b[2]
            data = [[P(h, ParagraphStyle("hc", parent=cellst, fontName="Helvetica-Bold")) for h in head]]
            data += [[P(v, cellst) for v in row[:len(head)]] for row in rows]
            tb = Table(data, colWidths=[W / len(head)] * len(head), repeatRows=1)
            tb.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e8eef4")),
                                    ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#b0bec5")),
                                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f7f9fb")])]))
            story += [tb, Spacer(1, 6)]
    buf = io.BytesIO()
    SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.7 * cm, rightMargin=1.7 * cm, topMargin=1.5 * cm,
                      bottomMargin=1.5 * cm, title="Fiche d'analyse").build(story)
    return buf.getvalue()


def render_export(f, d, t, ns=""):
    """Boutons d'export de la fiche (PDF et Word), préparés à la demande (graphiques compris)."""
    key = f"exp_{ns}_{t}_{round(d['fv'].get('JV centrale') or 0)}_{f.get('price')}"
    c1, c2, c3 = st.columns([1.3, 1, 1])
    if c1.button("📄 Préparer la fiche PDF / Word", key=f"prep_{ns}_{t}"):
        try:
            blocks = report_model(f, d, t)
            st.session_state[key] = (report_pdf(blocks), report_docx(blocks))
        except ImportError as e:
            st.error(f"Export impossible : bibliothèque manquante ({e.name}) — pip install -r requirements.txt")
        except Exception as e:
            st.error(f"Export impossible ({type(e).__name__} : {e})")
    if key in st.session_state:
        pdf, docx_ = st.session_state[key]
        fn = f"Fiche_{re.sub(r'[^A-Za-z0-9]+', '_', t)}_{TODAY:%Y%m%d}"
        c2.download_button("⬇️ PDF", pdf, f"{fn}.pdf", "application/pdf", key=f"dlp_{ns}_{t}")
        c3.download_button("⬇️ Word", docx_, f"{fn}.docx",
                           "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"dlw_{ns}_{t}")


def wacc_effect(f, fv, ref=0.085):
    """CMPC nettement au-dessus (> 9,5 %) de ce qu'utilisent souvent les analystes (8 à 9 %) : effet chiffré sur le DCF
    (scénario central recalculé à 8,5 %). Rien si le CMPC est dans cette zone : il n'explique alors pas l'écart."""
    res = fv.get("_dcfx")
    if not res or not f.get("shares") or res["wacc"]["wacc"] <= 0.095:
        return None
    x = res["x"]
    v = lambda w: (lambda r: (r["eq_perp"] + r["eq_exit"]) / 2 / f["shares"])(
        dcfx_project(x, res["g_rev"], 0.0, w, res["tg"], res["exit_mult"], bool(f.get("cyclical"))))
    try:
        a, b = v(res["wacc"]["wacc"]), v(ref)
    except Exception:
        return None
    w_dcf = fv.get("_w_dcf", 0.5)
    return (f"taux d'actualisation : CMPC {res['wacc']['wacc'] * 100:.1f} %, au-dessus des 8 à 9 % souvent retenus "
            f"par les analystes ; à {ref * 100:.1f} %, le DCF (scénario central) passerait de {a:.0f} à {b:.0f}, soit "
            f"{(b - a) * w_dcf:+.0f} sur la juste valeur")


# ---------- Section 11 du détail : actualités récentes et sentiment
NEWS_EVENTS = [   # (catégorie, motif dans le titre anglais) : événements qui font souvent bouger le cours
    ("Résultats", r"\b(earnings|results|quarter|Q[1-4]\b|beats?|miss(es|ed)?|EPS|revenue (tops|falls|rises))"),
    ("Prévisions", r"\b(guidance|outlook|forecast|raises? (its )?(full-year|outlook|guidance)|cuts? (its )?(outlook|guidance))"),
    ("Analystes", r"\b(upgrade[sd]?|downgrade[sd]?|price target|initiat(es|ed) coverage|overweight|underweight|"
                  r"outperform|underperform|buy rating|sell rating)"),
    ("Opération", r"\b(acquir(e|es|ed|ition)|merger|to buy|takeover|buyout|spin[- ]?off|divest|stake)"),
    ("Contrat", r"\b(contract|order|deal|partnership|agreement|award(ed)?|wins?)\b"),
    ("Justice / régulateur", r"\b(lawsuit|sue[sd]?|probe|investigation|SEC|DOJ|FTC|antitrust|recall|fine[sd]?|"
                             r"settle(ment|s|d)?|subpoena)\b"),
    ("Produit / autorisation", r"\b(FDA|approv(al|es|ed)|launch(es|ed)?|trial)\b"),
    ("Capital", r"\b(dividend|buyback|repurchase|offering|share sale|debt|bond|notes)\b"),
    ("Direction", r"\b(CEO|CFO|chief executive|resign(s|ed)?|appoint(s|ed)?|steps down|succession)\b"),
]


def news_event(title):
    """Catégorie de l'événement (premier motif reconnu) ou "" ."""
    for cat, pat in NEWS_EVENTS:
        if re.search(pat, title, re.I):
            return cat
    return ""


YAHOO_UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36")


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_yahoo_news(isin, name=""):
    """Titres Yahoo Finance d'une valeur sans fiche Finviz (Europe…) : ISIN -> symbole Yahoo (RI.PA…), puis
    recherche par nom ; seuls les titres que Yahoo rattache à ce symbole sont gardés.
    -> (symbole, [[iso, titre, source, lien], …]) du plus récent au plus ancien."""
    import requests

    def search(q, news=0, quotes=0):
        err = None
        for host in ("query1", "query2"):
            try:
                r = requests.get(f"https://{host}.finance.yahoo.com/v1/finance/search", timeout=15,
                                 params={"q": q, "newsCount": news, "quotesCount": quotes},
                                 headers={"User-Agent": YAHOO_UA})
                r.raise_for_status()
                return r.json()
            except Exception as e:
                err = e
        raise err

    qs = search(isin, quotes=3).get("quotes") or []
    q = next((x for x in qs if x.get("quoteType") == "EQUITY"), qs[0] if qs else None)
    if not q:
        return None, []
    sym, out, seen = q["symbol"], [], set()
    # le nom court (« PERNOD RICARD ») donne les meilleurs résultats ; le nom long (« … SA », « … Société
    # anonyme ») n'en donne parfois aucun
    for term in dict.fromkeys(t for t in (q.get("shortname"), name, q.get("longname")) if t):
        for n in search(term, news=30).get("news") or []:
            if sym not in (n.get("relatedTickers") or []) or n.get("title") in seen or not n.get("providerPublishTime"):
                continue
            seen.add(n["title"])
            out.append([datetime.fromtimestamp(n["providerPublishTime"]).isoformat(timespec="minutes"), n["title"],
                        n.get("publisher") or "Yahoo Finance", n.get("link") or ""])
        if out:
            break
    return sym, sorted(out, reverse=True)


def news_for(f):
    """Titres de la fiche Finviz ; à défaut (valeur européenne lue sur Zonebourse), titres Yahoo Finance via l'ISIN.
    -> (titres, libellé de la source, filtrer sur le nom de la société ?)"""
    if f.get("news"):
        return f["news"], "Finviz", True
    isin = f.get("isin") or ""
    if isin:
        y = _guarded(f"yahoo_{isin}", lambda: fetch_yahoo_news(isin, f.get("name") or f.get("company") or ""))
        if y and y[1]:
            return y[1], f"Yahoo Finance ({y[0]})", False       # déjà rattachés au symbole par Yahoo
    return [], "", False


def news_recent(f, days=7):
    """Titres des `days` derniers jours, avec le ton et la catégorie d'événement ; résumé 7 j / 30 j."""
    news, src, filt = news_for(f)
    if not news:
        return None, None
    tk = f.get("ticker") or ""
    ns = (news_sentiment(news, ticker=tk, company=f.get("company", "")) if filt
          else news_sentiment(news, ticker="", company=""))
    if not ns:
        return None, None
    ns["src"] = src
    det = ns["détail"].copy()
    det = det[(TODAY - det["Date"]).dt.days <= days]
    links = {n[1]: n[3] for n in news if len(n) > 3 and n[3]}
    det["Événement"] = det["Titre"].map(news_event)
    det["Ton"] = det["Score"].map(lambda v: "🟢 positif" if v >= 15 else "🔴 négatif" if v <= -15 else "⚪ neutre")
    from urllib.parse import quote_plus
    det["Lien"] = det["Titre"].map(lambda t_: links.get(t_) or
                                   "https://www.google.com/search?tbm=nws&q=" + quote_plus(t_))
    return det, ns


def _guarded(name, fn, ttl=600):
    """Appel réseau protégé : après un échec, pas de nouvel essai pendant `ttl` secondes (l'onglet reste rapide)."""
    k = f"_fail_{name}"
    if time.time() - st.session_state.get(k, 0) < ttl:
        return None
    try:
        return fn()
    except Exception:
        st.session_state[k] = time.time()
        return None


def render_news(f, ns_key=""):
    det, ns = news_recent(f)
    if det is None:
        st.caption("Actualités non disponibles : pas de fiche Finviz, et aucun titre Yahoo Finance trouvé"
                   + (" (code ISIN absent de la page Zonebourse)." if not f.get("isin") else "."))
    else:
        c = st.columns(4)
        c[0].metric("Titres sur 7 jours", ns["n7"], f"{ns['pos7']} positifs · {ns['neg7']} négatifs", delta_color="off")
        c[1].metric("Ton sur 7 jours", score_txt(ns["s7"], ns["n7"]) or "n.d.")
        c[2].metric("Ton sur 30 jours", score_txt(ns["s30"], ns["n30"]) or "n.d.")
        ev = det[det["Événement"] != ""]
        c[3].metric("Événements repérés (7 j)", len(ev))
        if len(det):
            show = det.sort_values("Date", ascending=False)[["Date", "Ton", "Score", "Événement", "Titre", "Source", "Lien"]]
            show["Date"] = show["Date"].dt.strftime("%d/%m %H:%M")
            st.dataframe(show, hide_index=True, width="stretch", key=f"news_{ns_key}", column_config={
                "Ton": st.column_config.TextColumn(help="positif : score ≥ +15 · négatif : score ≤ −15 · neutre "
                                                        "entre les deux (titre factuel, sans mot nettement positif "
                                                        "ou négatif)"),
                "Score": st.column_config.NumberColumn(format="%d", help="ton du titre, de −100 à +100 (VADER + "
                                                                           "vocabulaire financier)"),
                "Événement": st.column_config.TextColumn(help="type d'information qui fait souvent bouger le cours "
                                                              "(repéré par mots-clés)"),
                "Titre": st.column_config.TextColumn(width="large"),
                "Lien": st.column_config.LinkColumn(display_text="ouvrir", help="article d'origine (ou recherche "
                                                                                  "Google Actualités du titre)")})
        if not len(det):
            det30, _ = news_recent(f, 30)
            if det30 is not None and len(det30):
                st.caption("Aucun titre sur les 7 derniers jours — titres des 30 derniers jours :")
                show = det30.sort_values("Date", ascending=False)[["Date", "Ton", "Score", "Événement", "Titre",
                                                                  "Source", "Lien"]]
                show["Date"] = show["Date"].dt.strftime("%d/%m %H:%M")
                st.dataframe(show, hide_index=True, width="stretch", key=f"news30_{ns_key}", column_config={
                    "Titre": st.column_config.TextColumn(width="large"),
                    "Lien": st.column_config.LinkColumn(display_text="article")})
            else:
                st.caption("Aucun titre sur les 30 derniers jours.")
        if ns["src"] == "Finviz":
            st.caption("Titres de la fiche Finviz qui citent la société (Reuters, Bloomberg, Barron's, communiqués…). "
                       "Le ton est calculé titre par titre, sans comprendre le contexte : lisez les événements repérés "
                       "avant d'en tirer une conclusion.")
        else:
            st.caption(f"Titres {ns['src']}, en anglais, que Yahoo rattache à la société (code ISIN "
                       f"{f.get('isin')}) — mis en cache 1 heure. Couverture plus mince que Finviz pour les valeurs "
                       "européennes (souvent quelques titres par semaine). Le ton est calculé titre par titre, sans "
                       "comprendre le contexte.")
    ape = _guarded("ape", fetch_apewisdom)
    tk = (f.get("ticker") or "").upper()
    x = (ape or {}).get(tk)
    if x:
        m, m0 = int(x.get("mentions") or 0), int(x.get("mentions_24h_ago") or 0)
        mom = f" ({(m / m0 - 1) * 100:+.0f} % sur 24 h)" if m0 >= 10 else ""
        st.markdown(f"**Reddit** — rang **{x.get('rank')}** des actions les plus citées, {m} mentions sur 24 h{mom}. "
                    "L'attention ne dit pas le sens : à croiser avec le ton des actualités.")
    elif ape is not None:
        st.caption("Reddit : action peu ou pas citée sur 24 h (hors des ~600 premières d'ApeWisdom).")
    else:
        st.caption("Reddit (ApeWisdom) momentanément indisponible.")
    fg = _guarded("fg", fetch_fear_greed)
    if fg:
        cur = fg[0]
        st.caption(f"Contexte de marché : indice Fear & Greed {float(cur['score']):.0f}/100 "
                   f"({FG_LABELS.get(cur['rating'], cur['rating'])}) — détail dans l'onglet 😨 Fear & Greed.")


def render_detail(f, d, t=None, ns=""):
    fv = d["fv"]
    cur = {"USD": "$", "EUR": "€"}.get(f.get("currency"), f.get("currency", "") + " ")
    name = company_name(f, 60)
    # ---------- 1. Fiche d'identité
    tk = t or f.get("ticker") or ""
    st.markdown(f"### {name} ({tk})" if name and tk and name != tk else f"### {name or tk}")
    render_export(f, d, tk, ns)
    st.markdown("##### 1. Fiche d'identité")
    c = st.columns(5)
    c[0].metric("Cours", f"{f['price']:.2f} {cur}" if f["price"] else "n.d.")
    c[1].metric("Capitalisation", f"{f['mcap'] / 1e3:,.1f} Md {cur}".replace(",", " ") if f.get("mcap") else "n.d.")
    lev = d.get("lev")
    c[2].metric("Dette nette / EBITDA", "sans objet (banque / assurance)" if f.get("bank") else
                "trésorerie nette" if lev is not None and lev < 0 else f"{lev:.1f}x" if lev is not None else "n.d.")
    c[3].metric("Bêta", f"{f['beta']:.2f}" if f.get("beta") else "n.d.")
    c[4].metric("Dividende / action", f"{f['div_ps']:.2f} {cur}" if f.get("div_ps") else "aucun")
    info = [x for x in (f.get("sector"), f.get("industry") or (f.get("zb_sector") or "").replace("-", " "),
                        f"devise {f.get('currency', 'USD')}",
                        f"prochains résultats : {f['earnings']}" if f.get("earnings") else None,
                        "⟳ cyclique" if f.get("cyclical") else None, "🏦 modèle financier" if f.get("bank") else None) if x]
    st.caption(" · ".join(info) + f" — source : {short_source(f)}")
    if f.get("fx_note"):
        st.caption("💱 " + f["fx_note"])
    # ---------- 2. En bref
    st.markdown("##### 2. En bref")
    c = st.columns(4)
    jv = fv.get("JV centrale")
    c[0].metric("Juste valeur", f"{jv:.0f} {cur}" if jv else "n.d.",
                f"{fv['Écart (%)']:+.0f} %" if fv.get("Écart (%)") is not None else None)
    c[1].metric("Objectif analystes", f"{f['target']:.0f} {cur}" if f["target"] else "n.d.",
                f"{d['an']['Potentiel (%)']:+.0f} %" if d["an"]["Potentiel (%)"] is not None else None)
    c[2].metric("Précision", precision_txt(fv).split(" (")[0] if jv else "n.d.")
    c[3].metric("Croissance implicite / attendue", reverse_txt(fv))
    st.info(_md(en_bref(f, d, cur)))
    why_ = _divergence_causes(f, fv)
    if fv.get("_alerts") or why_:
        st.markdown("**Lecture de l'écart de la JV** — ajustements du calcul et raisons de l'écart entre méthodes")
        for a_ in fv.get("_alerts") or []:
            st.warning(a_)
        if why_:
            st.warning(_md("Écart entre méthodes — " + why_))
    # ---------- 3. Bénéfice par action
    st.markdown("##### 3. Bénéfice par action (BNA)")
    ser = _eps_series(f)
    if ser:
        edf = pd.DataFrame(ser, columns=["Exercice", "BNA", "Type"])
        import altair as alt
        st.altair_chart(alt.Chart(edf).mark_bar().encode(
            x=alt.X("Exercice:N", sort=list(edf["Exercice"]), title=None),
            y=alt.Y("BNA:Q", title=f"BNA ({cur.strip()})"),
            color=alt.Color("Type:N", scale=alt.Scale(domain=["publié", "consensus"], range=["#607d8b", "#43a047"]),
                            legend=alt.Legend(orient="bottom", title=None)),
            tooltip=["Exercice", alt.Tooltip("BNA:Q", format=".2f"), "Type"]).properties(height=220),
            width="stretch")
        pub = [v for _, v, t in ser if t == "publié"]
        est = [v for _, v, t in ser if t == "consensus"]
        txt = []
        g_h = _cagr(pub[0], pub[-1], len(pub) - 1) if len(pub) >= 2 else None
        if g_h is not None:
            txt.append(f"croissance annuelle publiée {g_h:+.0f} % ({ser[0][0]} → {[s_ for s_ in ser if s_[2] == 'publié'][-1][0]})")
        if pub and est:
            g_e = _cagr(pub[-1], est[-1], len(est))
            if g_e is not None:
                txt.append(f"attendue par le consensus {g_e:+.0f} %/an")
        if f.get("eps_5y") is not None:
            txt.append(f"consensus à 5 ans {f['eps_5y']:.0f} %/an")
        st.caption((("BNA " + " ; ".join(txt) + ". ") if txt else "") +
                   "Historique : BNA comptable publié. Prévisions : consensus des analystes, en général sur le BNA "
                   "non-GAAP (« ajusté » : hors rémunérations en actions, amortissements d'acquisitions, éléments "
                   "exceptionnels) — c'est lui qu'utilisent le PER cible et les comparables ; le PER historique reste "
                   "sur le BNA comptable, le DCF part des flux de trésorerie et déduit les rémunérations en actions.")
    else:
        st.caption("Non disponible pour cette source.")
    # ---------- 4. DCF
    st.markdown("##### 4. DCF")
    if fv.get("_dcfx"):
        render_dcfx(f, fv["_dcfx"])
    elif fv.get("_dcf"):
        fcf, g0, ke, sh = fv["_dcf"]
        st.caption(f"DCF simplifié (fiche seule) : flux libre {fcf / 1e3:.1f} Md, croissance {g0 * 100:.0f} % décroissant "
                   f"vers {DEFAULT_P['tg'] * 100:.1f} %, coût des fonds propres {ke * 100:.1f} % → {fv['DCF']:.0f} {cur} "
                   f"(fourchette {fv['DCF bas']:.0f}–{fv['DCF haut']:.0f}). Analyse complète en ligne : DCF d'entreprise.")
        with st.expander("Sensibilité du DCF (coût des fonds propres × croissance terminale)"):
            kes = [ke + x for x in (-0.02, -0.01, 0.0, 0.01, 0.02)]
            tgs = [0.015, 0.02, 0.025, 0.03, 0.035]
            grid = pd.DataFrame({f"g {t_ * 100:.1f} %": [dcf_value(fcf, g0, t_, k) / sh if k > t_ + 0.01 else None
                                                          for k in kes] for t_ in tgs},
                                index=[f"{k * 100:.1f} %" for k in kes])
            st.dataframe(grid.round(0))
            st.caption(f"Valeur par action selon le taux d'actualisation (lignes) et la croissance à long terme "
                       f"(colonnes) ; croissance initiale {g0 * 100:.0f} %. Cours : {f['price']:.2f}.")
    else:
        why_ = [n for n in fv["notes"] if "DCF" in n or "flux" in n or "assureur" in n]
        st.caption("Pas de DCF retenu" + (" : " + " · ".join(why_) if why_ else "."))
    # ---------- 5. PER et multiples
    st.markdown("##### 5. PER et multiples")
    mult = fv.get("_mult") or {}
    rows = [{"Méthode": lbl, "Hypothèse": mult.get(k, ""), f"Valeur ({cur.strip()})": fv[k],
             "Retenue": "✓" if not fv.get("_keep") or k in fv["_keep"] else "écartée"}
            for k, lbl in METHODS if k != "DCF" and fv.get(k) is not None]
    if rows:
        st.dataframe(pd.DataFrame(rows).round(1), hide_index=True, width="stretch")
    if fv.get("_peers_txt"):
        st.caption("Par rapport aux pairs : " + fv["_peers_txt"] + ".")
    m = []
    if f.get("fwd_pe"):
        m.append(f"PER actuel sur N+1 {f['fwd_pe']:.1f}x")
    if f.get("pe"):
        m.append(f"sur 12 mois {f['pe']:.1f}x")
    if f.get("fwd_pe") and f.get("eps_5y") and f["eps_5y"] > 0:
        m.append(f"PEG {f['fwd_pe'] / f['eps_5y']:.2f}")
    if f.get("pe_hist"):
        m.append(f"PER historique médian {f['pe_hist']:.1f}x")
    if m:
        st.caption(" · ".join(m) + ".")
    if f.get("peers"):
        with st.expander(f"Sociétés comparables ({len(f['peers'])})"):
            st.dataframe(pd.DataFrame(f["peers"]).rename(columns={
                "ticker": "Ticker", "fpe": "PER prév.", "pfcf": "P/FCF", "pe": "PER", "g": "Croiss. BNA 5 ans (%)",
                "mcap": "Capitalisation"}), hide_index=True)
    # ---------- 6. Synthèse
    st.markdown("##### 6. Synthèse")
    football_field(fv, f, cur)
    render_fv_tables(fv, cur)
    # ---------- 7. Juste valeur et objectif des analystes
    st.markdown("##### 7. Juste valeur calculée et objectif des analystes")
    if jv and f.get("target"):
        e = (jv / f["target"] - 1) * 100
        cmp_ = [{"Repère": "Juste valeur calculée", "Valeur": round(jv, 1), "vs cours (%)": fv.get("Écart (%)")},
                {"Repère": f"Objectif moyen des analystes ({f.get('n_analysts') or '?'} analystes)",
                 "Valeur": round(f["target"], 1), "vs cours (%)": d["an"].get("Potentiel (%)")}]
        ast_ = analyst_style_target(f)
        if ast_ and not ast_.get("skip") and f.get("price"):
            cmp_.append({"Repère": "Objectif à 12 mois façon analystes (VE/EBIT moyen 5 ans × EBIT N+1)",
                         "Valeur": round(ast_["target"], 1), "vs cours (%)": (ast_["target"] / f["price"] - 1) * 100})
        st.dataframe(pd.DataFrame(cmp_).round(1), hide_index=True)
        if ast_:
            st.caption(analyst_style_txt(ast_, f, cur.strip()))
        br_ = target_bridge_txt(target_bridge(f, fv), f, cur.strip(), fv.get("JV si pic"))
        if br_:
            st.caption(br_)
        vm_ = value_markers(f, fv, cur.strip())
        if vm_:
            st.dataframe(pd.DataFrame(vm_, columns=["Repère", "Valeur", "Lecture"]), hide_index=True,
                         width="stretch")
        pr_ = premium_reasons(f, fv)
        if pr_:
            with st.expander("🔍 Pourquoi cette prime ?", expanded=True):
                st.markdown(premium_intro(pr_))
                st.dataframe(pd.DataFrame([{"": ICON_P[r[3]], "Critère": r[0], "Valeur": r[1], "Lecture": r[2]}
                                           for r in pr_["rows"]]), hide_index=True, width="stretch")
                st.markdown(f"**{pr_['verdict']}**")
                st.caption(PREMIUM_NOTE)
        txt = [f"Écart juste valeur / objectif : {e:+.0f} %."]
        if f.get("eps_next_y") and f["eps_next_y"] > 0:
            txt.append(f"L'objectif correspond à un PER de {f['target'] / f['eps_next_y']:.1f}x le BNA N+1 "
                       f"(marché : {f['fwd_pe']:.1f}x)." if f.get("fwd_pe") else
                       f"L'objectif correspond à un PER de {f['target'] / f['eps_next_y']:.1f}x le BNA N+1.")
        st.caption(" ".join(txt))
        if f.get("merge_notes"):
            st.caption("Finviz + Zonebourse — " + " · ".join(f["merge_notes"]))
        if abs(e) > 10:
            causes = gap_causes(f, fv, lev)
            if causes:
                st.markdown(_md("Raisons probables de l'écart :\n" + "\n".join(f"- {x}" for x in causes)))
        st.caption("L'objectif des analystes est une comparaison utile, il vise souvent un cours dans 12 mois, "
                   "sur le BNA N+2.")
    else:
        st.caption("Objectif des analystes non disponible.")
    # ---------- 8. Risques et points de vigilance
    st.markdown("##### 8. Risques et points de vigilance")
    bad, good = risk_points(f, d)
    if bad:
        st.markdown(_md("\n".join(f"- ⚠️ {x}" for x in bad)))
    if good:
        st.markdown(_md("\n".join(f"- ✅ {x}" for x in good)))
    if not bad and not good:
        st.caption("Aucun point particulier détecté dans les données.")
    # ---------- 9. Signaux de marché (initiés, super investors, analystes, options)
    st.markdown("##### 9. Signaux de marché")
    a, b = st.columns(2)
    with a:
        ins = d["ins"]
        st.markdown(f"**Initiés ({(ins or {}).get('Source') or f.get('insider_source', 'fiche Finviz')}) — achats et "
                    "ventes en bourse**")
        if ins:
            if ins.get("Alerte"):
                (st.error if "🚨" in ins["Alerte"] else st.success)(
                    "Opération d'un dirigeant dans les 7 derniers jours : " + ins["Alerte"])
            if ins.get("Sortie dirigeant"):
                st.warning(ins["Sortie dirigeant"] + " — vente liée à une levée d'options : hors signal, mais un "
                           "dirigeant de premier rang qui cède l'essentiel de sa détention est souvent mal perçu.")
            st.markdown(ins["Résumé"])
            if ins.get("Frise"):
                st.markdown(f"12 derniers mois : {ins['Frise']} · dernière opération : {ins['Dernière opération']}")
                st.caption(("Même calcul que le screener AMF (même seuil). " if ins.get("Source") == "déclarations AMF"
                            else "Même calcul que le screener investors (mêmes seuils). ") + "🟢 achats sur 6 mois sans vente "
                           "significative sur 3 mois · 🟠 achats et ventes récentes · 🔴 ventes récentes supérieures aux "
                           "achats · ⚪ non significatif. Frise : 🟩 achats nets, 🟥 ventes nettes, ⬜ rien.")
            keep = ("Achats 365j", "Ventes 365j", "Acheteurs", "Achats groupés (max en 30 j)", "Achats PDG / DAF (M$)",
                    "Hausse médiane de détention (%)", "Ventes avec levée d'options (M$)")
            st.dataframe(pd.DataFrame([{k.replace("M$", "M€") if ins.get("Devise") == "€" else k: ins.get(k)
                                        for k in keep if ins.get(k) is not None}]).T.rename(columns={0: "12 mois"}))
            if ins.get("Devise") == "€":
                st.caption("Montants en millions d'euros (déclarations à l'AMF). Ventes sous mandat de gestion comptées"
                           + (" — il y en a sur 12 mois." if ins.get("Ventes sous mandat") else "."))
            for x in ins["Achats notables"]:
                st.caption("Achat : " + x)
            st.caption("Un achat (argent personnel du dirigeant) est un signal plus fort qu'une vente : diversification, "
                       "impôts ou plans de cession programmés expliquent souvent les ventes.")
        else:
            st.caption("Non disponible pour cette source.")
    with b:
        st.markdown("**Super investors (Dataroma, déclarations 13F)**")
        si = d["si"]
        if si.get("Grille") and not si["Lignes"]:
            st.markdown(si["Résumé"])
        if si.get("Grille"):
            st.caption("Signal calculé comme dans le screener investors : nombre de fonds qui renforcent (▲) ou allègent "
                       "(▼) au dernier trimestre, comparé au trimestre précédent (↗ / ↘).")
        if si["Lignes"]:
            st.markdown(si["Résumé"])
            if si.get("Grille"):
                st.caption("Tableau ci-dessous : fonds détenteurs et dernier mouvement déclaré par chacun (il peut "
                           "dater du trimestre précédent si le fonds n'a pas encore publié son 13F).")
            if si["Cours vs prix moyen (%)"] is not None:
                st.caption(f"Prix de détention moyen des fonds : {si['Prix de détention moyen']:.2f} — cours "
                           f"{si['Cours vs prix moyen (%)']:+d} % " + ("(sous leur prix : vous achèteriez moins cher qu'eux)"
                                                                       if si['Cours vs prix moyen (%)'] < 0 else "(au-dessus de leur prix)"))
            lg = pd.DataFrame(si["Lignes"]).sort_values("% portefeuille", ascending=False)
            st.dataframe(lg, hide_index=True, height=230, column_config={
                "% portefeuille": st.column_config.NumberColumn(format="%.2f", help="poids de l'action dans le portefeuille du gérant : conviction")})
            st.caption("Données trimestrielles (déclarations 13F publiées ~45 jours après la fin du trimestre) : un signal lent.")
        else:
            st.caption("Aucune donnée (valeur non suivie par Dataroma ou non chargée).")
    nb = f" · {f['n_analysts']} analystes" if f.get("n_analysts") else ""
    st.markdown(f"**Analystes** — consensus {d['an']['Consensus'] or 'n.d.'}{nb}"
                + (f" ({f['recom']})" if f["recom"] else "")
                + (f", {d['an']['Relèvements 180j']} relèvement(s) / {d['an']['Abaissements 180j']} abaissement(s) sur 180 j"
                   if f["ratings"] else ""))
    if d["an"]["Actions récentes"]:
        st.dataframe(pd.DataFrame(d["an"]["Actions récentes"]), hide_index=True)
    op = pd.DataFrame(f["options"])
    if not op.empty:
        st.markdown(f"**Options par échéance** — P/C encours {d['pc'].get('P/C encours', 'n.d.')}"
                    + (f" ({d['pc']['Lecture P/C']})" if d["pc"].get("Lecture P/C") else ""))
        op["P/C"] = (op["OI puts"].astype(float) / op["OI calls"].astype(float).replace(0, float("nan"))).round(2)
        st.dataframe(op, hide_index=True)
    # ---------- 10. Notes de calcul et sources
    st.markdown("##### 10. Notes de calcul et sources")
    with st.expander("Détail des hypothèses et notes du calcul"):
        st.markdown(_md("\n".join(f"- {n}" for n in fv["notes"]) or "—"))
    st.caption("Sources : " + short_source(f) + ". Estimation argumentée, pas un conseil en investissement.")
    # ---------- 11. Actualités et sentiment
    st.markdown("##### 11. Actualités et sentiment (7 derniers jours)")
    render_news(dict(f, ticker=f.get("ticker") or tk), ns)


def scenario_values(f, d):
    """Juste valeur par scénario : JV centrale × (BNA N+1 du scénario / BNA N+1 central)."""
    sc, jv = d["sc"], d["fv"].get("JV centrale")
    if sc.empty or not jv or "BNA" not in sc:
        return pd.DataFrame()
    n1 = sc[sc["Exercice"].astype(str).str.startswith("N+1")]
    c = n1.loc[n1["Scénario"] == "Central", "BNA"]
    if c.empty or not c.iloc[0]:
        return pd.DataFrame()
    rows = []
    for name in ("Pessimiste", "Central", "Optimiste"):
        e = n1.loc[n1["Scénario"] == name, "BNA"]
        if e.empty or e.iloc[0] is None:
            continue
        v = jv * e.iloc[0] / c.iloc[0]
        rows.append({"Scénario": name, "BNA N+1": round(e.iloc[0], 2), "Juste valeur": round(v, 1),
                     "Écart vs cours (%)": round((v / f["price"] - 1) * 100, 1) if f["price"] else None})
    return pd.DataFrame(rows)


def render_scenarios(f, d):
    sc = d["sc"]
    if sc.empty:
        st.info("États financiers indisponibles.")
        return
    sv = scenario_values(f, d)
    if not sv.empty:
        st.markdown("**Juste valeur selon le scénario**")
        st.dataframe(sv, hide_index=True, column_config={
            "Écart vs cours (%)": st.column_config.NumberColumn(format="%+.0f %%")})
        pes = sv.loc[sv["Scénario"] == "Pessimiste", "Écart vs cours (%)"]
        if not pes.empty and pes.iloc[0] is not None:
            st.caption("Lecture : même dans le scénario pessimiste, le cours reste sous la juste valeur — marge de sécurité."
                       if pes.iloc[0] > 0 else
                       "Lecture : dans le scénario pessimiste, le cours dépasse la juste valeur — pas de marge de sécurité.")
    with st.expander("Comment lire ce tableau ?"):
        st.markdown("""
- **Central** = prévisions du consensus ; **Pessimiste / Optimiste** = résultats inférieurs / supérieurs (±10 % sur le BNA,
  ±25 % pour une valeur cyclique, écart élargi la 2e année).
- **Juste valeur selon le scénario** = juste valeur centrale ajustée du BNA du scénario : si le BNA est 10 % plus bas,
  la juste valeur l'est aussi. Si le cours est sous la juste valeur *pessimiste*, la marge de sécurité est confortable.
- **EBITDA / capex / flux libre** : ce que l'entreprise génère et investit ; **dette nette / EBITDA** : endettement
  (< 1 faible, > 3 élevé, négatif = plus de trésorerie que de dettes). Une dette qui monte en scénario pessimiste est un signal de fragilité.
""")
    st.markdown("**Détail par exercice**")
    st.dataframe(sc.round(2), hide_index=True, width="stretch")
    unit = "Md$" if f.get("currency", "USD") == "USD" else f"Md {f.get('currency')}"
    if f.get("bank"):
        st.caption("Banque : central = consensus Zonebourse ; pessimiste / optimiste ± 10 % sur le BNA (± 15 % en N+2), "
                   "dividende au même taux de distribution, actif net ajusté du bénéfice non distribué.")
    elif f.get("zb"):
        st.caption(f"Central = consensus Zonebourse ; pessimiste / optimiste ± 10 % sur BNA et EBITDA (± 25 % si cyclique), "
                   f"élargi en N+2. Montants en {unit}.")
    else:
        cap = (f"{f['capex'] / 1e3:.1f} {unit}" if f["capex"] is not None else
               "inconnu (page sans états financiers : flux libre = marge actuelle, rachats non déduits ; "
               "cochez « compléter en ligne » pour un calcul complet)")
        st.caption(f"Départ : EBITDA 12 mois {f['ebitda'] / 1e3:.1f} {unit}, capex {cap}, "
                   f"dette nette actuelle {(f['net_debt'] or 0) / 1e3:.1f} {unit} (négatif = trésorerie nette), "
                   f"rachats + dividendes maintenus au rythme actuel.")


def detail_and_scenarios(df, det, keys, key):
    t = keys[st.selectbox("Valeur", list(df["Ticker"]), key=f"{key}_sel")]
    f, d = det[t]
    s1, s2 = st.tabs(["🔍 Détail", "📈 Scénarios"])
    with s1:
        render_detail(f, d, t, key)
    with s2:
        render_scenarios(f, d)


def main():
    st.set_page_config(page_title="Finviz Dashboard", page_icon="📊", layout="wide")
    # barres de défilement toujours visibles dans les tableaux (macOS les masque par défaut)
    st.markdown("""<style>
    [data-testid="stDataFrame"] *::-webkit-scrollbar { -webkit-appearance: none; width: 12px; height: 12px; }
    [data-testid="stDataFrame"] *::-webkit-scrollbar-track { background: rgba(128,128,128,0.12); border-radius: 6px; }
    [data-testid="stDataFrame"] *::-webkit-scrollbar-thumb { background: rgba(110,110,110,0.65); border-radius: 6px;
        border: 2px solid transparent; background-clip: padding-box; }
    [data-testid="stDataFrame"] *::-webkit-scrollbar-thumb:hover { background: rgba(80,80,80,0.85);
        background-clip: padding-box; }
    @supports (-moz-appearance: none) {
        [data-testid="stDataFrame"] * { scrollbar-width: auto; scrollbar-color: rgba(110,110,110,0.65) rgba(128,128,128,0.12); }
    }
    </style>""", unsafe_allow_html=True)
    st.title("📊 Finviz Dashboard")
    st.caption("Juste valeur · initiés & super investors · analystes · Put/Call · scénarios BNA / EBITDA / capex / dette")

    with st.sidebar:
        st.header("📁 Portefeuille")
        pf_file = st.file_uploader("Fichier de positions (xlsx)", type=["xlsx"])
        only_nyse = st.checkbox("NYSE uniquement (sinon toutes les actions en USD)", value=False)
        pf, pf_all = pd.DataFrame(), pd.DataFrame()
        if pf_file is not None:
            try:
                pf = load_portfolio(io.BytesIO(pf_file.getvalue()), only_nyse)
                pf_all = load_portfolio(io.BytesIO(pf_file.getvalue()), all_ccy=True)
                st.caption(f"{len(pf)} valeur(s) US : {', '.join(pf['Ticker'])}")
            except Exception as e:
                st.warning(f"Fichier illisible : {e}")
        default = ", ".join(pf["Ticker"]) if not pf.empty else ", ".join(DEFAULT_TICKERS)
        # le fichier a-t-il changé (nouveau fichier ou contenu différent) ? -> liste de tickers et données mises à jour
        sig = hashlib.md5(pf_file.getvalue()).hexdigest() if pf_file is not None and not pf.empty else None
        auto = False
        if "tickers_txt" not in st.session_state:
            st.session_state["tickers_txt"] = default
        if sig and sig != st.session_state.get("pf_sig"):
            old = {t.strip().upper() for t in st.session_state["tickers_txt"].split(",") if t.strip()}
            new = set(pf["Ticker"])
            if st.session_state.get("pf_sig"):                 # fichier remplacé ou modifié
                st.session_state["pf_change"] = (sorted(new - old), sorted(old - new))
            st.session_state["tickers_txt"], st.session_state["pf_sig"] = default, sig
            auto = True                                        # récupère aussitôt les données de la nouvelle liste
        tickers = [t.strip().upper() for t in st.text_area("Tickers", key="tickers_txt").split(",") if t.strip()]
        add, rem = st.session_state.pop("pf_change", ([], []))
        if add or rem:
            st.success("Fichier de positions modifié — " + " ; ".join(
                x for x in (f"ajoutées : {', '.join(add)}" if add else "", f"retirées : {', '.join(rem)}" if rem else "") if x))
        max_days = st.slider("Options : échéances jusqu'à (jours)", 30, 400, 200, 10)
        use_cache = st.checkbox(f"Utiliser le cache ({CACHE_HOURS} h)", value=True,
                                help="Décoché : tout est relu et retéléchargé — données Finviz en ligne, compléments "
                                     "des fiches déposées, et pages HTML (Finviz et Zonebourse).")
        if st.button("🧹 Vider le cache", help="Efface les données téléchargées (dossier cache_finviz) et les pages HTML "
                                              "gardées en mémoire. Les résultats affichés restent jusqu'à la prochaine "
                                              "récupération ou analyse."):
            n = clear_caches()
            st.success(f"Cache vidé ({n} fichier(s) supprimé(s)).")
        use_oi = st.checkbox("Initiés via OpenInsider (12 mois complets)", value=True,
                             help="Sinon : les ~20 dernières transactions de la fiche Finviz")
        go = st.button("📡 Récupérer les données", type="primary") or auto

        st.divider()
        st.header("🛒 Valeurs à étudier")
        st.caption("Pages enregistrées (Cmd+S) : fiche Finviz (+ pages Options), Zonebourse Finances (+ Valorisation). "
                   "Les fiches d'une même action sont combinées.")
        # formulaire : ajouter / retirer des pages ne relance pas le script (pas d'interruption d'une
        # récupération Finviz en cours) ; l'analyse ne part qu'au clic sur « Analyser »
        n_wl = st.session_state.get("wl_n", 0)
        with st.form("wl_form", border=False):
            ups = st.file_uploader("Pages HTML", type=["html", "htm"], accept_multiple_files=True, key=f"wl_up_{n_wl}")
            online = st.checkbox("Compléter en ligne (Finviz, OpenInsider, Dataroma)", value=True,
                                 help="Fiche Finviz déposée : ajoute états financiers, options, initiés et super investors. "
                                      "Action US déposée seulement en pages Zonebourse : va chercher la fiche Finviz en ligne "
                                      "et combine les deux. Valeurs européennes : sans effet.")
            c1, c2 = st.columns(2)
            run = c1.form_submit_button("🔎 Analyser", type="primary")
            clear = c2.form_submit_button("🗑️ Vider", help="Retire les pages déposées et leurs résultats (avant de redéposer des pages à jour) ; "
                                                          "les valeurs envoyées par le Screener JV restent")
        if clear:
            st.session_state["wl_n"] = n_wl + 1
            for k in ("wl_pages", "wl_errs_pages", "zb_store"):   # pages HTML seulement (valeurs du screener gardées)
                st.session_state.pop(k, None)
            st.rerun()
        if run and not ups:
            st.warning("Aucune page déposée.")
        st.divider()
        with st.expander("Réglages de la juste valeur"):
            n_reset = st.session_state.get("jv_reset", 0)

            def K(name):                          # nouvelle clé après « réinitialiser » -> valeurs par défaut
                return f"jv_{name}_{n_reset}"

            if st.button("↺ Réinitialiser les réglages", help="Revenir aux valeurs par défaut (et vider le classement cyclique forcé)"):
                st.session_state["jv_reset"] = n_reset + 1
                st.session_state["cyc_force"] = ""
                st.rerun()
            cyc_txt = st.text_input("Forcer le classement cyclique", key="cyc_force", placeholder="ex. VST:non, MU:oui",
                                    help="Le classement cyclique (⟳) est automatique. Saisissez TICKER:oui ou TICKER:non "
                                         "pour le corriger (PER cible 8x et croissance plafonnée si cyclique).")
            capm = st.checkbox("Taux propre à chaque valeur (CAPM, bêta Finviz)", value=DEFAULT_P["capm"], key=K("capm"),
                               help="Coût des fonds propres = taux sans risque + bêta × prime de risque du marché")
            rf = st.slider("Taux sans risque (emprunt d'État 10 ans)", 0.02, 0.06, DEFAULT_P["rf"], 0.0025, key=K("rf"))
            erp = st.slider("Prime de risque du marché actions", 0.03, 0.08, DEFAULT_P["erp"], 0.0025, key=K("erp"))
            p = {"capm": capm, "rf": rf, "erp": erp,
                 "cyc_prem": st.slider("Prime de risque des valeurs cycliques sans bêta", 0.0, 0.06, DEFAULT_P["cyc_prem"], 0.005, key=K("cyc_prem"),
                                       help="Ajoutée au taux unique pour les valeurs cycliques sans bêta (pages Zonebourse : "
                                            "automobile, semi-conducteurs…). Les valeurs avec bêta Finviz ne sont pas concernées."),
                 "wacc": st.slider("Coût des fonds propres (taux unique, valeurs sans bêta)", 0.06, 0.13,
                                   DEFAULT_P["wacc"], 0.0025, key=K("wacc")),
                 "w_dcf": st.slider("Poids du DCF dans la juste valeur", 0.0, 1.0, DEFAULT_P["w_dcf"], 0.1, key=K("w_dcf"),
                                    help="Avec états financiers : JV = poids × DCF d'entreprise + (1 − poids) × multiples"),
                 "tg": st.slider("Croissance terminale", 0.01, 0.04, DEFAULT_P["tg"], 0.0025, key=K("tg")),
                 "gcap": st.slider("Plafond de croissance initiale", 0.05, 0.35, DEFAULT_P["gcap"], 0.01, key=K("gcap")),
                 "pe_base": st.slider("PER de base (croissance 10 %)", 10.0, 30.0, DEFAULT_P["pe_base"], 0.5, key=K("pe_base")),
                 "pe_min": st.slider("PER minimum", 5.0, 20.0, DEFAULT_P["pe_min"], 0.5, key=K("pe_min")),
                 "pe_max": st.slider("PER maximum", 15.0, 40.0, DEFAULT_P["pe_max"], 0.5, key=K("pe_max")),
                 "pe_cyc_mid": st.slider("PER des cycliques sur BNA de milieu de cycle", 6.0, 16.0,
                                         DEFAULT_P["pe_cyc_mid"], 0.5, key=K("pe_cyc_mid"),
                                         help="BNA déjà normalisé (moyenne du cycle) : multiple normalisé 10–12x."),
                 "pe_cyc": st.slider("PER des cycliques sur BNA N+1 (pic possible)", 4.0, 15.0, DEFAULT_P["pe_cyc"], 0.5,
                                     key=K("pe_cyc"), help="Utilisé quand l'historique ne permet pas de calculer un BNA "
                                                           "de milieu de cycle."),
                 "sbc": st.checkbox("Déduire la rémunération en actions du flux libre", value=DEFAULT_P["sbc"], key=K("sbc")),
                 "coe_bank": st.slider("Banques : coût des fonds propres", 0.08, 0.15, DEFAULT_P["coe_bank"], 0.0025, key=K("coe_bank")),
                 "pe_bank": st.slider("Banques européennes : PER cible", 5.0, 14.0, DEFAULT_P["pe_bank"], 0.5, key=K("pe_bank")),
                 "pe_bank_us": st.slider("Banques et courtiers US : PER cible", 8.0, 18.0, DEFAULT_P["pe_bank_us"], 0.5,
                                         key=K("pe_bank_us"))}
        with st.expander("🏛️ SEC EDGAR (valeurs US)"):
            use_edgar = st.checkbox("Compléter avec les dépôts SEC (10-K, Form 4)", value=False, key="edgar_on",
                                    help="Historique de 10 exercices pour les valeurs cycliques, nombre d'actions et "
                                         "rachats sur 10 ans, déclarations d'initiés Form 4 quand OpenInsider ne "
                                         "répond pas. Valeurs cotées aux États-Unis seulement.")
            ua_ = st.text_input("Nom et adresse e-mail (exigés par la SEC)", key="edgar_ua",
                                placeholder="Jean Dupont jean.dupont@exemple.fr",
                                help="Envoyés uniquement à la SEC dans l'en-tête des requêtes (règle d'accès de la "
                                     "SEC) ; jamais enregistrés. Variable d'environnement SEC_USER_AGENT possible.")
            ES.configure((ua_ or os.environ.get("SEC_USER_AGENT", "")) if use_edgar else "",
                         os.path.join(CACHE_DIR, "edgar"))
            if use_edgar and not ES.enabled():
                st.caption("⚠️ saisissez un nom et une adresse e-mail valides")
        ov = {}
        for item in re.split(r"[,;\s]+", cyc_txt or ""):
            if ":" in item:
                tk, v = item.split(":", 1)
                v = v.strip().lower()
                if v in ("oui", "non", "o", "n", "yes", "no"):
                    ov[tk.strip().upper()] = {"Cyclique": "oui" if v[0] in "oy" else "non"}

    if go:
        raws, errs = {}, []
        bar = st.progress(0.0)
        for i, t in enumerate(tickers):
            try:
                raws[t] = fetch_ticker(t, max_days, use_cache, log=errs.append, openinsider=use_oi)
            except FinvizBlocked as e:
                errs.append(f"⛔ {e} — chargement interrompu ({len(tickers) - i} valeur(s) non chargée(s))")
                break
            except Exception as e:
                errs.append(f"{t} : {type(e).__name__} — {e}")
            bar.progress((i + 1) / len(tickers), text=t)
        bar.empty()
        st.session_state["raws"], st.session_state["errs"] = raws, errs
    raws = {t: r for t, r in st.session_state.get("raws", {}).items() if t in tickers}   # liste à jour uniquement

    if run and ups:
        facts, errs = {}, []
        # même fichier déposé plusieurs fois : seule la dernière version est gardée
        pages = [(u.name, u.getvalue().decode("utf-8", errors="ignore")) for u in ups]
        pages, dups = dedupe_pages(pages)
        if dups:
            errs.append(f"{len(dups)} page(s) en double ignorée(s) (la plus récente est gardée) : " + ", ".join(dups)
                        + ". Retirez-les de la liste avec la croix ✕, ou « 🗑️ Vider » avant de redéposer.")
        fv_pages = [x for x in pages if page_kind(*x) == "finviz"]
        groups = {}
        zb_store = {}
        for name, html in pages:
            kind = page_kind(name, html)
            if kind == "zonebourse":
                groups.setdefault(zb_company(html) or name, []).append((name, html))
            elif kind is None:
                errs.append(f"{name} : ni Finviz ni Zonebourse")
        for comp, grp in groups.items():
            try:
                hs_ = tuple(h for _, h in grp)
                f = (copy.deepcopy(parse_zonebourse_cached(hs_, grp[0][0])) if use_cache
                     else parse_zonebourse(list(hs_), grp[0][0]))     # cache décoché : relecture complète
                f["source"] = "Zonebourse : " + ", ".join(n for n, _ in grp)
                f.setdefault("isin", zb_isin("".join(hs_)))
                facts[f["ticker"]] = f
                zb_store[f["ticker"]] = f             # gardées pour enrichir le portefeuille
                if f["missing_finances"]:
                    errs.append(f"{comp} : onglet Finances non chargé — pas de consensus ni d'actif net ; "
                                "ajoutez-le pour une analyse complète.")
            except Exception as e:
                errs.append(f"{comp} : {e}")
        fraws = (copy.deepcopy(raws_from_finviz_pages_cached(tuple(fv_pages))) if use_cache
                 else raws_from_finviz_pages(fv_pages))
        if fv_pages and not fraws:
            errs.append("Aucune fiche Finviz (page « quote ») reconnue parmi les pages Finviz déposées.")
        for t, r in fraws.items():
            if online:
                with st.spinner(f"{t} : compléments en ligne…"):
                    try:
                        _, hit = complete_online_cached(r, t, max_days, use_cache, log=errs.append, openinsider=use_oi)
                        r["source"] = "page Finviz + compléments en ligne" + (" (cache)" if hit else "")
                    except Exception as e:
                        errs.append(f"{t} : compléments en ligne impossibles ({e})")
            fvf = build_facts(r)
            zf = facts.get(t)
            if zf and zf.get("zb") and zf.get("currency", "USD") != "USD":
                # même ticker mais pages Zonebourse dans une autre devise (ex. argenx coté à Bruxelles en EUR,
                # fiche Finviz de l'ADR en USD) : pas de combinaison, deux lignes distinctes
                errs.append(f"{t} : pages Zonebourse en {zf['currency']} et fiche Finviz en USD (ADR) — "
                            f"analyses séparées ({t} et {t} ({zf['currency']})).")
                facts[f"{t} ({zf['currency']})"] = zf
                facts[t] = fvf
            else:
                facts[t] = merge_facts(fvf, zf) if zf and zf.get("zb") else fvf
        # action US déposée seulement en pages Zonebourse : fiche Finviz récupérée en ligne puis combinée
        for t, zf in list(facts.items()):
            if online and zf.get("zb") and not zf.get("merged") and zf.get("currency") == "USD":
                with st.spinner(f"{t} : fiche Finviz en ligne…"):
                    try:
                        fvf = build_facts(fetch_ticker(t, max_days, use_cache, log=errs.append, openinsider=use_oi))
                        if fvf.get("price"):
                            facts[t] = merge_facts(fvf, zf)
                            facts[t]["source"] = "Finviz en ligne + " + (zf.get("source") or "Zonebourse")
                        else:
                            errs.append(f"{t} : fiche Finviz introuvable — analyse Zonebourse seule.")
                    except Exception as e:
                        errs.append(f"{t} : Finviz en ligne indisponible ({e}) — analyse Zonebourse seule.")
        st.session_state["wl_pages"], st.session_state["wl_errs_pages"] = facts, errs
        st.session_state["zb_store"] = zb_store

    # « Valeurs à étudier » = pages HTML analysées + valeurs envoyées par le Screener JV
    st.session_state["wl"] = {**(st.session_state.get("wl_scr") or {}), **(st.session_state.get("wl_pages") or {})}
    st.session_state["wl_errs"] = list(st.session_state.get("wl_errs_pages") or []) + \
        list(st.session_state.get("wl_errs_scr") or [])

    tab1, tab2, tab4, tab5, tab6, tab7, tab8, tab3 = st.tabs(
        ["📋 Portefeuille", "🛒 Valeurs à étudier", "😨 Fear & Greed", "🔥 Sentiment screener", "🔎 Screener JV",
         "🇺🇸 Screener investors", "🇫🇷 Screener AMF", "ℹ️ Méthode"])

    # ---------- calculs (avant les onglets : une erreur n'empêche pas l'affichage des autres onglets)
    pf_res, pf_err = None, None
    if raws:
        try:
            pf_res = evaluate_memo("pf", portfolio_facts(raws, st.session_state.get("zb_store") or {}), p, ov, pf)
        except Exception as e:
            pf_err = e
    wl_res, wl_err, wl_notes = None, None, []
    if st.session_state.get("wl"):
        try:
            wl = st.session_state["wl"]
            # valeurs déjà en portefeuille (toutes devises, même hors NYSE) : prix d'entrée
            in_pf = not pf_all.empty and bool({_base(t) for t in wl} & set(pf_all["Ticker"]))
            wl_res = evaluate_memo("wl", wl, p, ov, pf_all if in_pf else None)
            PE = "Prix d'entrée"
            for t, f in wl.items():                # position dans une autre devise : prix d'entrée non comparable
                pos = pf_all[pf_all["Ticker"] == _base(t)] if not pf_all.empty else pf_all
                if not pos.empty and f.get("currency", "USD") not in set(pos["Devise"]):
                    wl_notes.append(f"{_base(t)} : position en {pos['Devise'].iloc[0]} (prix d'entrée "
                                    f"{num(pos.iloc[0][PE]):g}), analyse en {f.get('currency', 'USD')} — "
                                    "pages d'une autre cotation (ex. ADR US). Pour comparer, enregistrez la page "
                                    "Zonebourse de la cotation où vous détenez le titre.")
        except Exception as e:
            wl_err = e

    with tab1:
        for e in st.session_state.get("errs", []):
            st.warning(e)
        if pf_err:
            st.error(f"Erreur pendant l'analyse du portefeuille : {pf_err}")
        elif not pf_res:
            st.info("👈 Chargez votre fichier de positions puis cliquez sur « Récupérer les données ».")
        else:
            df, det, keys = pf_res
            st.dataframe(df, hide_index=True, width="stretch", column_config=FMT)
            st.caption("⟳ = valeur cyclique (bénéfices proches d'un pic). 📄 = combinée avec les pages Zonebourse déposées "
                       "dans « Valeurs à étudier » (date des pages). Survolez les en-têtes pour la définition des colonnes.")
            memo = st.session_state.setdefault("_eval_memo", {})
            if "pf_xlsx" not in memo:                       # fichier reconstruit seulement si les résultats changent
                buf = io.BytesIO()
                with pd.ExcelWriter(buf) as xw:
                    df.to_excel(xw, sheet_name="Synthèse", index=False)
                    for t, (_, d) in det.items():
                        if not d["sc"].empty:
                            d["sc"].to_excel(xw, sheet_name=f"Scénarios {t}"[:31], index=False)
                        if d["fv"].get("_dcfx"):                       # modèle DCF complet par valeur
                            proj, scen, wacc = dcfx_tables(d["fv"]["_dcfx"])
                            row = 0
                            for tb in (scen, wacc, proj):
                                tb.to_excel(xw, sheet_name=f"DCF {t}"[:31], index=False, startrow=row)
                                row += len(tb) + 3
                memo["pf_xlsx"] = buf.getvalue()
            st.download_button("💾 Excel (synthèse, scénarios, modèles DCF)", memo["pf_xlsx"], "finviz_dashboard.xlsx")
            st.divider()
            detail_and_scenarios(df, det, keys, "pf")

    with tab2:
        for e in st.session_state.get("wl_errs", []):
            st.warning(e)
        if st.session_state.get("wl_scr"):
            c1, c2 = st.columns([3, 1])
            c1.caption("Venues des screeners (JV, sentiment, initiés) : " + ", ".join(st.session_state["wl_scr"]))
            if c2.button("🗑️ Retirer les valeurs du screener", key="scr_rm"):
                for k in ("wl_scr", "wl_errs_scr"):
                    st.session_state.pop(k, None)
                st.rerun()
        for e in wl_notes:
            st.info(e)
        if wl_err:
            st.error(f"Erreur pendant l'analyse des pages : {wl_err}")
        elif not wl_res:
            st.info("👈 Déposez vos pages Finviz / Zonebourse dans la barre latérale (section « 🛒 Valeurs à étudier »), "
                    "puis cliquez sur « Analyser ». « 🗑️ Vider » retire toutes les pages déposées.")
        else:
            df2, det2, keys2 = wl_res
            st.dataframe(df2, hide_index=True, width="stretch", column_config=FMT)
            st.caption("Valeurs Zonebourse : pas d'initiés, de super investors ni de Put/Call (n.d.) ; montants dans la devise de la valeur.")
            st.divider()
            detail_and_scenarios(df2, det2, keys2, "wl")

    with tab4:
        render_fear_greed()

    with tab5:
        render_reddit_top(p, use_cache, max_days, use_oi)

    with tab6:
        render_screener(p, use_cache, max_days, use_oi)

    with tab7:
        render_investors_screener(max_days, use_cache, use_oi)
    with tab8:
        render_amf_screener()

    with tab3:
        st.markdown("""
**Juste valeur** = moyenne de deux méthodes, ou médiane de trois ou quatre (DCF, PER, comparables d'industrie, PER historique
Zonebourse) ; fourchette = plus bas / plus haut des méthodes, sans celles qui s'écartent de plus de 50 % de la médiane.
Croissance retenue = BNA à 5 ans du consensus, plafonnée à la croissance du CA + 10 points (médiane 3 ans / 5 ans / 12 mois).
PER prévisionnel < 7 : bénéfices traités comme un haut de cycle (⟳). Pour Zonebourse : croissance mesurée depuis le dernier
exercice publié, croissance du DCF plafonnée par celle du CA, flux libre normalisé s'il est irrégulier, PER cible ≤ 1,25 × le PER historique.
- *DCF* : flux libre des 12 derniers mois (moyenne sur 4 périodes pour une valeur cyclique) moins la rémunération en actions,
  croissance initiale = prévision de BNA à 5 ans (plafonnée), décroissante jusqu'à la croissance terminale ; ± 1 point de coût du capital.
- *PER* : BNA N+1 du consensus × PER cible (18x pour 10 % de croissance, +0,4 point par point de croissance ; 8x pour une valeur cyclique) ; ± 15 %.
- *Comparables* : BNA N+1 × PER prévisionnel moyen de l'industrie (borné entre les PER minimum et maximum des réglages,
  ≤ 10x si cyclique) et flux libre × P/FCF moyen de l'industrie (borné 8–25x : les moyennes sont tirées par les géants) ; multiples de la page « Groups » de Finviz, mis en cache 24 h ; ± 15 %.
- *Comparables* : en priorité les sociétés comparables désignées par Finviz (lien « Peers » de la fiche, médianes de leurs
  PER prévisionnels et P/FCF) ; à défaut, moyenne de l'industrie.
- *Lecture de l'écart de la JV* (onglet Détail, section 2 ; ⚠️ dans la colonne « Précision JV » signale un ajustement
  important du calcul) : les deux causes les plus fortes de l'écart DCF / multiples, détectées
  dans les données — haut de cycle, investissements lourds (capex / amortissements), faible conversion du bénéfice en
  trésorerie, CMPC élevé, croissance anticipée par le cours, dette nette, ou à l'inverse trésorerie nette importante,
  flux supérieurs aux bénéfices, forte croissance du consensus.
- *Précision JV* : écart maximal des méthodes retenues à la JV centrale — 🟢 ≤ 15 % (méthodes concordantes), 🟠 ≤ 35 %,
  🔴 au-delà ; une méthode à plus de 50 % de la médiane est exclue et signalée (« DCF écarté »). Valeur cyclique dont le
  DCF est écarté : 🔴 « multiples du seul BNA de pic » (PER et comparables reposent sur le même bénéfice, non durable).
  Une forte décote avec des méthodes divergentes est souvent un piège (« value trap »).
- *Sensibilité du DCF* (détail d'une valeur) : valeur par action selon le taux d'actualisation et la croissance terminale ;
  part de la valeur terminale dans le DCF (> 80 % : signalée).
- *Repères complémentaires* (section 7) : **valeur sans croissance (EPV)** = EBIT normalisé (exercice en cours ; moyenne du
  cycle pour une cyclique) × (1 − impôt) ÷ CMPC − dette nette, par action → part du cours et de l'objectif payée pour la
  croissance ; **rendement implicite au cours** = taux d'actualisation qui ramène le DCF central au cours, comparé au CMPC ;
  **nombre d'actions** = évolution annuelle sur 5 ans (relution ✅ ≤ −1 %/an, dilution ⚠️ ≥ +2 %/an), rachats et rémunération
  en actions des 3 derniers exercices.
- *SEC EDGAR* (option de la barre latérale, valeurs US) : 10-K de la SEC → cycle mesuré sur 10 exercices pour les
  cycliques (marges, trésorerie, BNA corrigé des divisions d'actions) et nombre d'actions sur 10 ans ; déclarations Form 4
  lues à la SEC quand OpenInsider ne répond pas. Nom + e-mail exigés par la SEC, jamais enregistrés ; cache 1 jour.
- *DCF inversé* (colonne « Croissance implicite / attendue ») — avec états financiers : croissance du CA en année 1
  (décroissante ensuite jusqu'à la croissance perpétuelle, marges et capex inchangés) que suppose le cours dans le DCF
  d'entreprise, comparée à la croissance du CA retenue (≤ 5 % si cyclique) ; sans états financiers : croissance du flux
  libre, comparée à la croissance du BNA plafonnée : croissance annuelle que le cours suppose, à comparer à la croissance
  attendue. 🟢 le marché suppose nettement moins que prévu, 🔴 nettement plus.

**DCF d'entreprise** (dès que les états financiers sont disponibles : récupération en ligne, analyse complète, ou pages
Zonebourse Finances) : flux de trésorerie disponibles pour l'entreprise (FCFF) projetés sur 10 ans — CA de départ, croissance
(consensus Zonebourse pour les premières années, sinon croissance du CA constatée et du BNA) décroissant jusqu'à la croissance
perpétuelle ; marge de trésorerie d'exploitation (flux d'exploitation + intérêts après impôt − rémunération en actions) ou
EBITDA × conversion ; capex convergeant vers un niveau d'entretien. Actualisation au CMPC (coût des fonds propres CAPM avec bêta
ajusté, coût de la dette après impôt, poids de marché). Valeur terminale = moyenne croissance perpétuelle / multiple de sortie
(VE/EBITDA actuel du marché borné 4–12x, 9x si cyclique (appliqué à un EBITDA déjà normalisé), et au plus 1,5 × le multiple implicite de la valeur perpétuelle :
une société qui convertit peu d'EBITDA en trésorerie ne vaut pas un multiple normal). Pages Zonebourse : flux libre du consensus
pour les premières années, puis transition vers le flux « économique » NOPAT × (1 − g / RONIC) — résultat d'exploitation après
impôt moins l'investissement nécessaire à la croissance (RONIC = CMPC + 2 pts, 10 % au moins) — qui intègre loyers et
amortissements ; même formule pour la valeur perpétuelle. Valeur d'entreprise − dette nette − minoritaires = valeur des fonds propres.
Trois scénarios (pessimiste / central / optimiste, croissance ± 4 pts, marge ± 2 pts, creux de cycle si cyclique) pondérés
25 / 50 / 25 ; fourchette P10–P90 d'une simulation de Monte Carlo (2 000 tirages). JV = 50 % DCF + 50 % médiane des multiples
(réglable) ; précision = écart DCF / multiples. Sans états financiers (fiche seule) : ancien DCF simplifié et médiane des méthodes.

**Taux d'actualisation du DCF** : coût des fonds propres, car le flux libre utilisé est après intérêts (flux aux actionnaires).
Par défaut il est propre à chaque valeur (CAPM) : taux sans risque 4,3 % + bêta Finviz (borné entre 0,6 et 2) × prime de risque 5 %.
Sans bêta (pages Zonebourse), taux unique de 9 %, + 3 points pour une valeur cyclique (réglable). Le WACC est affiché à titre indicatif (dette nette, coût de la dette = taux sans
risque + 1,5 point, impôt 25 %) mais n'entre pas dans le calcul.

**😨 Fear & Greed** : indice CNN du sentiment du marché américain (0 = peur extrême, 100 = avidité extrême), calculé par
CNN à partir de sept indicateurs (momentum du S&P 500, force et largeur des cours, put/call, VIX, valeurs refuges,
obligations à haut rendement) ; jauge, historique 1 an avec le S&P 500, trois composantes ; données mises en cache 10 min.
Commentaire automatique facultatif avec une clé API Anthropic (saisie dans l'onglet, jamais enregistrée).

**Sentiment screener** : titres d'actualité de la fiche Finviz notés par VADER, un analyseur de ton complété d'un
vocabulaire financier (upgrade, beat, downgrade, miss…) ; mentions Reddit via ApeWisdom (gratuit, sans clé).
C'est un indicateur d'ambiance, pas un signal d'achat : à croiser avec la juste valeur.

**Cyclique** : croissance du BNA de l'année > 150 % (bénéfices probablement proches d'un pic), ou retour à la perte
après un exercice bénéficiaire. Exception (pages Zonebourse) : **redressement durable** si le consensus prévoit encore un
BNA en hausse d'au moins 10 % par an en N+1 et N+2, une marge d'EBITDA en progression et un CA en hausse d'au moins 5 % par
an — un pic de cycle se lit au contraire par un BNA qui plafonne ou recule. L'exception ne s'applique pas aux industries
cycliques par nature (semi-conducteurs et matériel informatique, automobile, pétrole et gaz, matériaux, chimie, mines).

**Redressement** (pages Zonebourse, valeur non cyclique) : si le BNA N+2 du consensus vaut au moins 2 fois le BNA N+1
(ou si N+1 est une perte et N+2 un bénéfice), le PER cible s'applique au BNA N+2 actualisé d'un an au coût des fonds
propres ; comparables et PER historique restent sur N+1.

**Banques et assureurs** (industrie Finviz « Bank / Capital Markets / Insurance », ou secteur Zonebourse « services bancaires »
/ « assurance ») : pas de DCF d'entreprise — leurs flux et leur dette viennent du bilan. Actif net × P/B justifié
((ROE − g) / (coût des fonds propres − g)), PER et PER historique. Banques : coût des fonds propres 11 % (CAPM aux US), PER
bancaire ; assureurs : coût des fonds propres de la valeur (≥ 8 %), PER cible selon la croissance plafonné à 1,25 × le
PER historique.

**Endettement et CMPC** : sans bêta de marché (pages Zonebourse), le taux unique vaut pour un endettement courant ; si la
dette dépasse 30 % de la capitalisation, le coût des fonds propres est majoré (bêta réendetté, plafond 20 %) et la marge
de crédit dépend de dette nette / EBITDA (1,5 à 6 points) — une dette lourde ne fait plus baisser le CMPC.

**Vague d'investissements sans consensus** (capex > 25 % du CA et > 2,5× les amortissements, sans pages Zonebourse) :
le DCF prolongerait la dépense sans voir le CA qu'elle doit produire ; il n'est pas retenu et la JV repose sur les multiples.

**Après le consensus** (années 4 à 10 du DCF) : la croissance du CA décroît linéairement vers la croissance perpétuelle
en partant au plus de 25 % — une hyper-croissance prévue sur 3 ans (ex. Broadcom, +60 %/an) n'est pas prolongée telle quelle.

**Scission ou cession** : avec Finviz (activités abandonnées dans les comptes) ou Zonebourse seul (CA prévu N < 80 % du
dernier exercice alors que la marge d'EBITDA tient et que le CA repart en N+1), le DCF, la croissance et la dette nette
repartent de la base pro forma N du consensus ; le bond du BNA entre l'exercice publié et N n'est pas lu comme un pic.
Une valeur cyclique est valorisée sur un flux normalisé et 8x le BNA N+1, et ses scénarios supposent un plateau en N+2.

**Initiés** et **Super investors** (portefeuille, valeurs à étudier, fiche détaillée) : même calcul et mêmes seuils que le
screener investors (voir « Signal des initiés » plus bas), pour que les deux onglets concordent. Initiés : OpenInsider
(formulaires 4 de la SEC, 12 mois, achats « P » et ventes « S » en bourse, ventes avec levée d'options exclues), ou à défaut la
fiche Finviz ; dédoublonnés. Super investors : grilles Dataroma du S&P 500 (dernier trimestre et trimestre précédent) ; hors
S&P 500, résumé de la fiche Dataroma (dernier mouvement déclaré par chaque fonds détenteur, déclarations 13F décalées d'un
trimestre). **Put/Call** : encours (open interest) puts / calls, toutes échéances jusqu'à N jours ;
> 1 = couverture ou prudence, < 0,7 = positionnement haussier ; la version ± 10 % ne garde que les prix d'exercice proches du cours.

**Scénarios** : EBITDA N+1 = EBITDA TTM × (1 + croissance du CA sur 12 mois, plafonnée à 50 %), N+2 = croissance à 5 ans plafonnée à 25 % ;
pessimiste −10 pts (−25 pts si cyclique), optimiste +8 pts (+15 pts). Capex = intensité actuelle (capex / CA) × CA projeté (+20 % en
pessimiste, −10 % en optimiste). Dette nette = dette nette actuelle − (EBITDA × conversion en flux d'exploitation − capex)
+ rachats et dividendes. BNA = consensus N+1, puis croissance à 5 ans ; ± 10 % (± 25 % si cyclique), élargi en N+2.

**PER historique** (si l'onglet Valorisation Zonebourse est chargé) : BNA N+1 × PER médian des exercices publiés, ± 15 %,
ajouté comme troisième méthode dans la moyenne.

**Banques, courtiers, assureurs** (page Zonebourse sans EBITDA ni flux libre, ou fiche Finviz secteur « Financial » :
banques, marchés de capitaux, assurance) : juste valeur = moyenne de l'actif net par action × P/B justifié
(ROE N+1 − g) / (coût des fonds propres − g), ± 1 point, et du BNA N+1 × PER bancaire, ± 15 %. PER bancaire : 8,5x en
Europe, 12x aux États-Unis ; coût des fonds propres : CAPM (bêta Finviz) pour les US, 11 % sinon. Le DCF n'est pas utilisé :
le flux de trésorerie d'une banque ou d'un courtier dépend des mouvements de son bilan, pas de son activité.

**Action US avec fiche Finviz et pages Zonebourse** déposées ensemble : les deux sont combinées. Finviz fournit le cours,
le bêta (taux d'actualisation), les initiés, les options, les actualités et les analystes ; Zonebourse les prévisions sur 3 ans
(scénarios), le PER historique (3e méthode) et la normalisation du flux libre. BNA homogène avec chaque multiple : BNA
Finviz (consensus ajusté) pour le PER cible et les comparables, BNA Zonebourse pour le PER historique (mêmes bases que leurs
PER). Scission ou cession détectée (activités abandonnées dans les comptes Finviz et CA prévu inférieur de plus de 3 % aux 12
derniers mois) : le DCF repart du CA et des marges du consensus (pro forma), sans l'activité sortie.

**Portefeuille + pages Zonebourse** : les pages Zonebourse déposées dans « Valeurs à étudier » servent aussi aux valeurs du
portefeuille (même ticker, même devise) : la fiche Finviz est alors combinée avec le consensus sur 3 ans, le PER historique et
l'objectif Zonebourse (colonne « Pages Zonebourse » avec la date des pages). Sans page : Finviz seul.

**Valeurs à étudier** : pages enregistrées, déposées dans la barre latérale (toujours visible, même avec un fichier de positions). Si la valeur est aussi dans le portefeuille, le prix d'entrée s'affiche avec « JV / prix d'entrée ». Une fiche Finviz seule donne le consensus, les analystes et les initiés ;
cochez « compléter en ligne » pour les états financiers (capex), les options et Dataroma. Une page Zonebourse « Finances » donne
les prévisions du consensus sur 2 à 3 ans : le scénario central reprend ces prévisions.

**🔎 Screener JV** : lit toutes les pages d'un screener Finviz (adresse collée), exclut les industries choisies
(mines, acier, charbon… par défaut), calcule une juste valeur rapide sur la fiche Finviz seule et garde les actions
dont la marge dépasse le seuil ; les meilleures partent en analyse complète dans « Valeurs à étudier ».

**Valeurs cycliques (⟳)** : la juste valeur centrale est une valeur de **milieu de cycle** — DCF sur la marge de trésorerie
moyenne du cycle (12 derniers mois exclus, jusqu'à 8 exercices, creux et pic compris) et marge d'EBITDA moyenne du cycle ;
PER cible (11x, multiple « normalisé » de milieu de cycle : 8x seulement si le BNA moyen n'est pas calculable et que le
BNA N+1 de pic sert de base) et comparables appliqués au BNA moyen (exercices publiés + N + N+1), jamais au BNA d'une année
de pic ou de creux — un multiple bas sur un bénéfice déjà normalisé pénaliserait deux fois.
Colonne « JV si bénéfices actuels durent » : même calcul avec les multiples sur le BNA N+1 — l'écart entre les deux mesure
le pari du marché sur la durée du cycle. Moins de 6 exercices : milieu de cycle approximatif (signalé).

**Comparables** : pairs désignés par Finviz (PER prévisionnel × BNA N+1, VE/EBITDA 12 mois × EBITDA 12 mois, P/FCF) ou,
pour une valeur lue sur Zonebourse, tableau des pairs de l'onglet Valorisation (PER et VE/EBITDA de l'exercice en cours ×
BNA et EBITDA du même exercice). Valeur centrale = médiane des pairs (bornée), fourchette = 1er–3e quartile ; position de
la société (prime / dans la norme / décote) affichée en section 5. VE/EBITDA écarté si la dette nette dépasse 4× l'EBITDA
(valeur des fonds propres trop sensible au multiple) ou si le VE publié est incohérent avec capitalisation + dette.
Valeur cyclique : BNA et EBITDA de milieu de cycle, multiples des pairs non plafonnés.

**Signal des initiés** (screeners investors et AMF, même règle) : 🟢 achats de dirigeants ≥ seuil sur 6 mois sans
vente significative sur 3 mois · 🟠 achats sur 6 mois mais aussi des ventes récentes (3 mois) · 🔴 ventes récentes
supérieures aux achats des 6 mois · ⚪ montants sous le seuil ou mouvements anciens. « Achats groupés » / « ventes
groupées » : au moins 3 dirigeants distincts dans la même fenêtre de 30 jours. Frise : solde mensuel achats − ventes
sur 12 mois (🟩 achats nets, 🟥 ventes nettes, ⬜ rien), pour voir d'un coup d'œil si les achats se répètent et si des
ventes sont récentes. Alerte 7 j : 🚨 vente / ✅ achat d'un dirigeant dans les 7 derniers jours (date de l'opération),
aussi affichée en tête de la colonne Initiés et dans la fiche détaillée. Ventes liées à une levée d'options : hors
signal ; mais « ⚠️ sortie PDG / DAF » si un PDG, DAF, président ou fondateur a vendu au moins la moitié de sa détention
dans les 6 derniers mois en lien avec une levée d'options (part vendue = titres vendus / (détention après + titres
vendus)) — colonne du screener investors, en tête de la colonne Initiés, point de vigilance de la fiche (valeurs
américaines : les déclarations AMF n'indiquent pas la détention restante).

**Conviction des super investors** (portefeuille, valeurs à étudier, fiche détaillée, screener investors) : chaque
mouvement du dernier 13F compte pour le poids qu'il représente dans le portefeuille du gérant — achat d'une ligne à 15 % =
+15 points, allègement de 41 % d'une ligne restée à 15 % = −10,7 points (part vendue exprimée en poids actuel), sortie
complète = au moins −0,5 point ; 5 points au plus par fonds, pour qu'un seul gérant ne décide pas de tout. 🟢 score ≥ +0,5,
🔴 ≤ −0,5, 🟠 achats et ventes significatifs qui se compensent, ⚪ mouvements symboliques. Le poids mesure la conviction du
gérant, pas le montant investi (0,6 % du portefeuille de Bridgewater = une grosse somme, mais une petite conviction).

**🇺🇸 Screener investors** : univers = valeurs du S&P 500 détenues par au moins N superinvestors (grille Dataroma
« Ownership »). Colonne Dirigeants : transactions OpenInsider sur 12 mois (levées d'options exclues, achats sous le montant
minimum ignorés, ventes toutes comptées). Colonnes Fonds : nombre de superinvestors qui renforcent / allègent la valeur
au dernier trimestre et au trimestre précédent (grilles « Last qtr buys / sells », et « 6 months » moins le dernier
trimestre) — rond d'après le dernier trimestre (🟢 plus de renforcements, 🔴 plus d'allègements, 🟠 autant, ⚪ aucun
mouvement), « ↗ en amélioration » / « ↘ en dégradation » si le solde a changé d'au moins 2 fonds par rapport au trimestre
précédent, « achats groupés » dès 3 fonds acheteurs dans le trimestre. Sentiment de l'univers : ratios
acheteurs / vendeurs des 90 derniers jours comparés aux 90 précédents.

**🇫🇷 Screener AMF** : déclarations des dirigeants publiées par l'AMF (base BDIF, article 19 du règlement MAR), sur
12 mois, pour le SBF 120 (fichier sbf120.csv) ou toutes les sociétés cotées à Paris (Euronext, Growth, Access). Chaque PDF
est lu une fois puis gardé sur disque. Retenus : acquisitions, souscriptions et cessions d'actions à prix > 0 ; exclus :
actions gratuites ou de performance, options, conversions, nantissements, prêts, donations, parts de FCPE, droits de
souscription. Une déclaration corrigée remplace l'initiale. Les ventes sous mandat de gestion sont comptées mais signalées.

**🔥 Sentiment screener** : les actions les plus citées sur Reddit (ApeWisdom, 24 h, hors ETF) sont passées au crible du ton
de leurs titres d'actualité Finviz ; sont retenues celles à l'actualité nettement positive (≥ 15 sur 30 j avec au moins 5 titres,
7 j non négatif), classées par score (60 % ton 30 j + 40 % ton 7 j, +5 si les mentions bondissent). Juste valeur indicative
(fiche Finviz seule).

Le cache des données en ligne dure 2 h (dossier ./cache_finviz). Le fichier de positions n'est pas mis en cache.

Ce n'est pas un conseil en investissement.
""")


if __name__ == "__main__":
    main()
