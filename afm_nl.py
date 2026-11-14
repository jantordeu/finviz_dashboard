"""Registre néerlandais des transactions des dirigeants (AFM, « Transacties leidinggevenden MAR 19 »).

Les sociétés cotées à Paris mais de droit néerlandais (Airbus, Stellantis, STMicroelectronics, Euronext, Pluxee,
Technip Energies…) déclarent les opérations de leurs dirigeants à l'AFM, pas à l'AMF. Ce module :
- liste les déclarations d'une société sur une période (export XML du registre, recherche par nom) ;
- lit la page de détail de chaque déclaration (instrument, nature, prix, quantité), gardée sur disque ;
- renvoie un tableau au même format que amf.load_transactions (mêmes règles d'exclusion : actions gratuites ou de
  performance, droits, options, prix nul…).
"""
from __future__ import annotations

import json
import os
import re
import time
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime
from typing import Callable, Optional
from urllib.parse import quote_plus

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE = "https://www.afm.nl"
REG = "/en/sector/registers/meldingenregisters/transacties-leidinggevenden-mar19-"
TYPE = "0ee836dc-5520-459d-bcf4-a4a689de6614"
UA = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0 Safari/537.36", "Accept-Language": "en-US,en;q=0.9"}
ENTITY = re.compile(r"\b(sa|sas|nv|n\.v\.|bv|b\.v\.|ltd|llc|gmbh|participations|holding|fund|fonds|trust|"
                    r"foundation|stichting|capital|invest|ag|plc|se|spa|s\.p\.a\.)\b", re.I)
BUY = re.compile(r"^(verwerving|aankoop|koop|acquisition|purchase|buy)", re.I)
SELL = re.compile(r"^(vervreemding|verkoop|disposal|sale|sell)", re.I)
EXCL_NATURE = re.compile(r"schenking|gift|donation|verpanding|pledge|erf|inherit|uitoefen|exercis|toeken|grant|"
                         r"conver|lening|lend|loan|dividend|bonus|inschrijving|subscription|vesting", re.I)
EXCL_INSTR = re.compile(r"performance|restricted|award|option|optie|right|recht|rsu|psu|warrant|bonus|convert|"
                        r"obligat|bond|note|unit|future|swap|cfd", re.I)
SHARE = re.compile(r"\b(aandee?l(en)?|share|shares|ordinary|gewone|equity)\b", re.I)


def keyword(name: str) -> str:
    """« STELLANTIS NV » -> « stellantis » ; « TECHNIP ENERGIES » -> « technip energies »."""
    k = re.sub(r"\b(n\.?v\.?|s\.?a\.?|s\.?e\.?|se|sa|nv|plc|ag)\b", " ", name, flags=re.I)
    return re.sub(r"\s+", " ", k).strip().lower()


def nl_num(t) -> Optional[float]:
    """« 14.363,00 » -> 14363.0 ; « 0,00 » -> 0.0 ; « 1,234.5 » -> 1234.5."""
    t = re.sub(r"[^\d,.\-]", "", str(t or ""))
    if not t:
        return None
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".") if t.rfind(",") > t.rfind(".") else t.replace(",", "")
    elif "," in t:
        t = t.replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", t):
        t = t.replace(".", "")
    try:
        return float(t)
    except ValueError:
        return None


def _get(s: requests.Session, url: str, pause: float = 0.6) -> str:
    for attempt in range(3):
        r = s.get(url, headers=UA, timeout=30)
        if r.status_code in (429, 503):
            time.sleep(min(float(r.headers.get("Retry-After") or 3 * (attempt + 1)), 20))
            continue
        r.raise_for_status()
        time.sleep(pause)
        return r.text
    raise RuntimeError("AFM : trop de requêtes, réessayez dans quelques minutes")


def parse_list_xml(xml: str) -> list[dict]:
    """Export XML du registre -> [{id, date, issuer, notifiable, last, associated, function, lei}]."""
    root = ET.fromstring(xml.encode("utf-8") if isinstance(xml, str) else xml)
    out = []
    for v in root.findall("vermelding"):
        g = lambda tag: (v.findtext(tag) or "").strip()
        try:
            d = datetime.strptime(g("transactiedatum").split(" ")[0], "%m/%d/%Y").date().isoformat()
        except ValueError:
            d = ""
        out.append({"id": g("meldingid"), "date": d, "issuer": g("uitgevendeinstelling"),
                    "notifiable": g("meldingsplichtige"), "last": g("meldingsplichtigeachternaam"),
                    "associated": g("nauwgelieerdaan"), "function": g("functie"), "lei": g("lei")})
    return out


def parse_detail(html: str) -> list[dict]:
    """Page de détail -> lignes du tableau « Transactions » (le tableau « Aggregated information » est ignoré)."""
    soup = BeautifulSoup(html, "html.parser")
    rows = []
    for h2 in soup.find_all("h2"):
        if h2.get_text(strip=True).lower() not in ("transactions", "transacties"):
            continue
        tb = h2.find_next("table")
        if tb is None:
            continue
        hdr = [th.get_text(" ", strip=True) for th in tb.find_all("th")]
        for tr in tb.find_all("tr")[1:]:
            cells = []
            for td in tr.find_all("td"):
                for sp in td.select("span.cc-mobile-title"):
                    sp.extract()
                cells.append(td.get_text(" ", strip=True))
            if len(cells) >= len(hdr):
                rows.append(dict(zip(hdr, cells)))
        break
    out = []
    for r in rows:
        g = lambda *ks: next((r[k] for k in r for kk in ks if k.lower().startswith(kk)), "")
        out.append({"instrument": g("instrument"), "isin": g("isin"), "categorie": g("transaction category",
                                                                                    "transactiecategorie"),
                    "nature": g("transaction type", "transactietype", "soort"), "programme": g("stock option",
                                                                                             "aandelenoptie"),
                    "lieu": g("trading place", "plaats"), "prix": nl_num(g("price", "prijs")),
                    "volume": nl_num(g("quantity", "aantal", "volume")), "devise": g("unit", "eenheid")})
    return out


def side_of(tx: dict) -> tuple[str, str]:
    """("A" | "V" | "", motif d'exclusion) — mêmes règles que pour l'AMF."""
    ins, nat = tx.get("instrument") or "", f"{tx.get('categorie') or ''} {tx.get('nature') or ''}".strip()
    if (tx.get("programme") or "").lower() in ("ja", "yes"):
        return "", "programme d'options / actions gratuites"
    if EXCL_INSTR.search(ins) or not SHARE.search(ins):
        return "", f"instrument : {ins or '?'}"
    if EXCL_NATURE.search(nat):
        return "", f"nature : {nat or '?'}"
    if not tx.get("prix") or not tx.get("volume"):
        return "", "prix ou volume nul"
    first = (tx.get("nature") or tx.get("categorie") or "").strip()
    if BUY.match(first) or BUY.match(tx.get("categorie") or ""):
        return "A", ""
    if SELL.match(first) or SELL.match(tx.get("categorie") or ""):
        return "V", ""
    return "", f"nature : {nat or '?'}"


def person_key(r: dict) -> tuple[str, str]:
    """(clé de regroupement, nom affiché)."""
    name = r["notifiable"]
    shown = name + (f" (liée à {r['associated']})" if r.get("associated") else "")
    if r.get("last") and not ENTITY.search(name):
        return r["last"].lower(), shown
    toks = re.findall(r"[a-z]+", name.lower())
    return (" ".join(toks[:2]) or "?"), shown


def company_declarations(s: requests.Session, name: str, since: date, till: date, cache_dir: str) -> list[dict]:
    """Déclarations d'une société sur la période (recherche par nom, contrôle du nom de l'émetteur)."""
    kw = keyword(name)
    p = os.path.join(cache_dir, f"list_{re.sub(r'[^a-z0-9]+', '_', kw)}.json")   # un fichier par société (6 h)
    if os.path.exists(p) and time.time() - os.path.getmtime(p) < 6 * 3600:
        try:
            with open(p, encoding="utf-8") as fh:
                return json.load(fh)
        except (OSError, ValueError):
            pass                                       # cache illisible : relu en ligne
    url = (f"{BASE}/export.aspx?KeyWords={quote_plus(kw)}&DateFrom={since:%d-%m-%Y}&type={TYPE}"
           f"&DateTill={till:%d-%m-%Y}&format=xml")
    rows = [r for r in parse_list_xml(_get(s, url)) if kw.split()[0] in r["issuer"].lower()]
    with open(p, "w", encoding="utf-8") as fh:
        json.dump(rows, fh)
    return rows


def detail(s: requests.Session, rid: str, cache_dir: str) -> list[dict]:
    p = os.path.join(cache_dir, f"detail_{rid}.json")
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as fh:
                return json.load(fh)
        except (OSError, ValueError):
            pass
    tx = parse_detail(_get(s, f"{BASE}{REG}/details?id={rid}"))
    with open(p, "w", encoding="utf-8") as fh:
        json.dump(tx, fh)
    return tx


def load_companies(companies: list[tuple[str, str]], since: date, cache_dir: str,
                   progress: Optional[Callable[[str], None]] = None):
    """companies : [(nom, isin)] -> (tx, diag, names {jeton: nom}, erreurs [nom: motif]) au format amf."""
    os.makedirs(cache_dir, exist_ok=True)
    s = requests.Session()
    till = date.today()
    lines, names, errs = [], {}, []
    for k, (nom, isin) in enumerate(companies):
        jeton = "AFM:" + isin
        if progress:
            progress(f"Registre AFM (Pays-Bas) : {nom} ({k + 1}/{len(companies)})")
        try:
            decl = company_declarations(s, nom, since, till, cache_dir)
        except Exception as e:  # noqa: BLE001
            errs.append(f"{nom} : registre AFM indisponible ({type(e).__name__})")
            continue
        names[jeton] = nom
        # pages de détail : 3 lectures en parallèle (celles déjà lues viennent du disque, sans requête)
        def _one(r_):
            try:
                return detail(s, r_["id"], cache_dir)
            except Exception as e:  # noqa: BLE001
                return e
        with ThreadPoolExecutor(3) as ex:
            details = list(ex.map(_one, decl))
        for r, txs in zip(decl, details):
            url = f"{BASE}{REG}/details?id={r['id']}"
            base = {"numero": r["id"], "date_pub": r["date"], "jeton": jeton, "societe": r["issuer"], "pdf": url}
            if isinstance(txs, Exception):
                lines.append({**base, "motif": f"lecture impossible : {str(txs)[:120]}"})
                continue
            if not txs:
                lines.append({**base, "motif": "aucune transaction reconnue sur la page"})
            pk, shown = person_key(r)
            for tx in txs:
                if tx.get("isin") and isin and tx["isin"] != isin and not tx["isin"].startswith("NL"):
                    side, why = "", f"autre titre ({tx['isin']})"
                else:
                    side, why = side_of(tx)
                lines.append({**base, "isin": tx.get("isin"), "declarant": shown, "fonction": r.get("function", ""),
                              "person": pk, "date": r["date"], "nature": f"{tx.get('categorie')} / {tx.get('nature')}",
                              "instrument": tx.get("instrument"), "lieu": tx.get("lieu"), "prix": tx.get("prix"),
                              "volume": tx.get("volume"), "devise": tx.get("devise"), "side": side, "motif": why,
                              "correction": False, "mandate": False, "commentaire": ""})
    diag = pd.DataFrame(lines)
    if diag.empty or "side" not in diag:
        return diag, diag, names, errs
    diag["side"] = diag["side"].fillna("")
    tx = diag[(diag["side"] != "") & diag["date"].notna() & (diag["date"] != "")].copy()
    tx["value"] = tx["prix"] * tx["volume"]
    tx = tx.drop_duplicates(subset=["jeton", "person", "date", "side", "volume", "prix"])
    tx["date"] = pd.to_datetime(tx["date"], errors="coerce")
    return tx.reset_index(drop=True), diag, names, errs
