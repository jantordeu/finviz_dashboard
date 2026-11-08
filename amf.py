"""
AMF — BDIF (bdif.amf-france.org) : déclarations des dirigeants (article 19 du règlement MAR).

API publique, sans clé :
    back/api/v1/informations?TypesInformation=DD&From=…&Size=…   liste (société, date, PDF)
    back/api/v1/documents/<chemin>                                 déclaration (PDF, formulaire standard)
Le détail (déclarant, nature, prix, volume) n'est que dans le PDF : chaque PDF est lu une fois puis gardé
sur disque (une déclaration publiée ne change plus ; une correction est une nouvelle déclaration).

Opérations retenues : acquisitions / souscriptions (achats) et cessions (ventes) d'actions, prix > 0.
Exclues : actions gratuites ou de performance, options, conversions, nantissements, prêts, donations,
parts de FCPE (épargne salariale), droits préférentiels de souscription, RSU.
"""
from __future__ import annotations

import io
import json
import os
import re
import unicodedata
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from typing import Callable, Optional

import pandas as pd
import requests

BASE = "https://bdif.amf-france.org/back/api/v1/"
FRONT = "https://bdif.amf-france.org/fr?typesInformation=DD"
HEADERS = {"User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"),
           "Accept": "application/json, application/pdf, */*"}
PARSE_VERSION = 1
MONTHS = {"janvier": 1, "fevrier": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6, "juillet": 7, "aout": 8,
          "septembre": 9, "octobre": 10, "novembre": 11, "decembre": 12}
EXCL_NATURE = re.compile(r"gratuit|performance|donation|donat|nantissement|pret|collar|exercice|conversion|"
                         r"attribution|heritage|succession|apport|echange|transfert|remise|levee|dividende|"
                         r"financement|couverture|derive", re.I)
ENTITY = {"sa", "sas", "sasu", "sarl", "sci", "ltd", "llc", "holding", "gmbh", "bv", "nv", "se", "sc", "srl",
          "limited", "inc", "trust", "fondation", "societe", "fund", "lp", "spa", "ag"}


def session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    return s


# --------------------------------------------------------------------------- #
# Liste des déclarations et identifiants AMF des sociétés
# --------------------------------------------------------------------------- #
def list_declarations(since: date, s: Optional[requests.Session] = None, page: int = 500,
                      progress: Optional[Callable[[str], None]] = None) -> list[dict]:
    """Déclarations publiées depuis `since` (toutes sociétés), de la plus récente à la plus ancienne."""
    s = s or session()
    rows, frm = [], 0
    while True:
        r = s.get(BASE + "informations", params={"TypesInformation": "DD", "From": frm, "Size": page}, timeout=40)
        r.raise_for_status()
        res = r.json().get("result") or []
        for x in res:
            socs = x.get("societes") or [{}]
            soc = next((c for c in socs if c.get("role") == "SocieteConcernee"), socs[0])
            doc = (x.get("documents") or [{}])[0]
            rows.append({"numero": x.get("numero"), "date_pub": (x.get("dateInformation") or "")[:10],
                         "jeton": soc.get("jeton") or "", "societe": soc.get("raisonSociale") or "",
                         "path": doc.get("path") or ""})
        if progress:
            progress(f"liste AMF : {len(rows)} déclarations lues")
        if len(res) < page or min(x["date_pub"] for x in rows[-len(res):]) < since.isoformat():
            break
        frm += page
    return [x for x in rows if x["date_pub"] >= since.isoformat() and x["path"] and x["numero"]]


def jeton_for_isin(isin: str, s: Optional[requests.Session] = None) -> tuple[str, str]:
    """Identifiant AMF (« jeton ») et raison sociale d'une société, d'après son code ISIN (recherche plein texte
    dans les déclarations : la société la plus fréquente parmi les résultats)."""
    s = s or session()
    r = s.get(BASE + "informations", params={"TypesInformation": "DD", "RechercheTexte": isin, "From": 0,
                                              "Size": 20}, timeout=30)
    r.raise_for_status()
    c = Counter()
    names = {}
    for x in r.json().get("result") or []:
        for soc in x.get("societes") or []:
            if soc.get("role") == "SocieteConcernee" and soc.get("jeton"):
                c[soc["jeton"]] += 1
                names[soc["jeton"]] = soc.get("raisonSociale") or ""
    if not c:
        return "", ""
    j = c.most_common(1)[0][0]
    return j, names[j]


def jetons_for_isins(isins: list[str], cache_file: str, max_age_days: int = 30,
                     progress: Optional[Callable[[str], None]] = None) -> dict[str, tuple[str, str]]:
    """{isin: (jeton, raison sociale)} — gardé sur disque `max_age_days` jours."""
    cache = {}
    if os.path.exists(cache_file):
        try:
            with open(cache_file, encoding="utf-8") as fh:
                cache = json.load(fh)
        except (OSError, ValueError):
            cache = {}
    limit = (date.today() - timedelta(days=max_age_days)).isoformat()
    todo = [i for i in isins if i not in cache or cache[i][2] < limit]
    if todo:
        s = session()
        with ThreadPoolExecutor(8) as ex:
            futs = {ex.submit(jeton_for_isin, i, s): i for i in todo}
            for k, f in enumerate(as_completed(futs), 1):
                i = futs[f]
                try:
                    j, n = f.result()
                    cache[i] = [j, n, date.today().isoformat()]
                except Exception:   # noqa: BLE001 — réessayé au prochain chargement
                    pass
                if progress:
                    progress(f"identifiants AMF : {k}/{len(todo)}")
        os.makedirs(os.path.dirname(cache_file) or ".", exist_ok=True)
        with open(cache_file, "w", encoding="utf-8") as fh:
            json.dump(cache, fh, ensure_ascii=False)
    return {i: (cache[i][0], cache[i][1]) for i in isins if i in cache and cache[i][0]}


def company_declarations(isin: str, since: date, cache_file: str,
                         s: Optional[requests.Session] = None) -> tuple[str, str, list[dict]]:
    """Déclarations d'une seule société depuis `since` : identifiant AMF retrouvé par l'ISIN, puis recherche par
    ISIN et par raison sociale (l'ISIN peut changer, ex. regroupement d'actions) ; seules les déclarations de
    cette société (même identifiant) sont gardées. -> (jeton, raison sociale, lignes au format list_declarations)."""
    s = s or session()
    jm = jetons_for_isins([isin], cache_file)
    if isin not in jm:
        return "", "", []
    jeton, name = jm[isin]
    rows = {}
    for q in dict.fromkeys(x for x in (isin, name) if x):
        r = s.get(BASE + "informations", params={"TypesInformation": "DD", "RechercheTexte": q, "From": 0,
                                                  "Size": 200}, timeout=40)
        r.raise_for_status()
        for x in r.json().get("result") or []:
            d = (x.get("dateInformation") or "")[:10]
            soc = next((c for c in x.get("societes") or [] if c.get("jeton") == jeton), None)
            doc = (x.get("documents") or [{}])[0]
            if soc and d >= since.isoformat() and doc.get("path") and x.get("numero"):
                rows[x["numero"]] = {"numero": x["numero"], "date_pub": d, "jeton": jeton,
                                     "societe": soc.get("raisonSociale") or name, "path": doc["path"]}
    return jeton, name, sorted(rows.values(), key=lambda r: r["date_pub"], reverse=True)


# --------------------------------------------------------------------------- #
# Lecture d'une déclaration (texte du PDF)
# --------------------------------------------------------------------------- #
def norm(t: str) -> str:
    t = unicodedata.normalize("NFKD", t or "")
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = t.replace("’", "'").replace("‘", "'").replace("\xa0", " ").replace(" ", " ")
    return re.sub(r"\s+", " ", t).strip()


def _between(t: str, a: str, b: str) -> str:
    m = re.search(a + r"\s*(.*?)\s*(?:" + b + ")", t, re.S)
    return m.group(1).strip() if m else ""


def _num(x: str) -> Optional[float]:
    x = (x or "").replace(" ", "").replace(",", ".")
    try:
        return float(x)
    except ValueError:
        return None


def _date_fr(x: str) -> Optional[str]:
    m = re.match(r"\s*(\d{1,2})(?:er)? ([a-z]+) (\d{4})", (x or "").lower())
    if m and m.group(2) in MONTHS:
        try:
            return date(int(m.group(3)), MONTHS[m.group(2)], int(m.group(1))).isoformat()
        except ValueError:
            return None
    m = re.match(r"\s*(\d{2})/(\d{2})/(\d{4})", x or "")
    return f"{m.group(3)}-{m.group(2)}-{m.group(1)}" if m else None


def parse_declaration(text: str) -> dict:
    """Champs d'une déclaration (texte brut du PDF) ; une déclaration peut contenir plusieurs transactions."""
    t = norm(text)
    k = t.find('"Les donnees')
    t = t[:k] if k > 0 else t
    isin = re.search(r"\b([A-Z]{2}[A-Z0-9]{9}\d)\s*-\s*DD\d+", t)
    decl = _between(t, r"ETROITEMENT LIEE\s*:", r"NOTIFICATION INITIALE")
    out = {
        "isin": isin.group(1) if isin else "",
        "declarant": decl,
        "modification": _between(t, r"MODIFICATION\s*:", r"COORDONNEES"),
        "emetteur": _between(t, r"EMETTEUR NOM\s*:", r"LEI\s*:|DETAIL DE LA TRANSACTION"),
        "commentaire": t.split("COMMENTAIRES :", 1)[1].strip() if "COMMENTAIRES :" in t else "",
        "tx": [],
    }
    for b in re.split(r"DATE DE LA TRANSACTION\s*:", t)[1:]:
        agg = re.search(r"INFORMATIONS AGREGEES\s*PRIX\s*:\s*([\d ]+(?:[.,]\d+)?)\s*([A-Za-z]+)\s*VOLUME\s*:\s*"
                        r"([\d ]+(?:[.,]\d+)?)", b)
        if not agg:
            agg = re.search(r"PRIX UNITAIRE\s*:\s*([\d ]+(?:[.,]\d+)?)\s*([A-Za-z]+)\s*VOLUME\s*:\s*"
                            r"([\d ]+(?:[.,]\d+)?)", b)
        aga = re.search(r"PERFORMANCES\s*:\s*(OUI|NON)", b)
        out["tx"].append({
            "date": _date_fr(b[:40]),
            "lieu": _between(b, r"LIEU DE LA TRANSACTION\s*:", r"NATURE DE LA TRANSACTION"),
            "nature": _between(b, r"NATURE DE LA TRANSACTION\s*:", r"DESCRIPTION DE L'INSTRUMENT"),
            "instrument": _between(b, r"DESCRIPTION DE L'INSTRUMENT FINANCIER\s*:",
                                   r"CODE D'IDENTIFICATION|INFORMATION DETAILLEE"),
            "prix": _num(agg.group(1)) if agg else None,
            "devise": agg.group(2) if agg else "",
            "volume": _num(agg.group(3)) if agg else None,
            "aga": aga.group(1) if aga else "",
        })
    return out


def side_of(tx: dict) -> tuple[str, str]:
    """("A" | "V" | "", motif d'exclusion)."""
    n, i = norm(tx.get("nature")).lower(), norm(tx.get("instrument")).lower()
    if tx.get("aga") == "OUI":
        return "", "actions gratuites / options"
    if not re.match(r"actions?\b", i) or re.search(r"gratuit|preferentiel|fcpe|restricted|droit", i):
        return "", f"instrument : {tx.get('instrument') or '?'}"
    if EXCL_NATURE.search(n):
        return "", f"nature : {tx.get('nature') or '?'}"
    if not tx.get("prix") or not tx.get("volume"):
        return "", "prix ou volume nul"
    if re.match(r"(acquisition|achat|souscription)", n):
        return "A", ""
    if re.match(r"(cession|vente)", n):
        return "V", ""
    return "", f"nature : {tx.get('nature') or '?'}"


def person_key(declarant: str) -> tuple[str, str, str]:
    """(clé de regroupement, nom affiché, fonction) — « Jean-François DESCAVES » et « jean descaves » ont la même
    clé (nom de famille) ; une société (holding…) garde son nom complet."""
    name, _, role = (declarant or "").partition(",")
    toks = re.findall(r"[a-z]+", norm(name).lower())
    if not toks:
        return "?", name.strip(), role.strip()
    if set(toks) & ENTITY:                       # société : « FLAMBOYANTS CONSULTING LTD (Holding) » -> 2 premiers mots
        key = " ".join(toks[:2])
    else:                                        # personne : nom de famille
        key = toks[-1]
    return key, name.strip(), role.strip()


def pdf_text(data: bytes) -> str:
    from pypdf import PdfReader
    return " ".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(data)).pages)


# --------------------------------------------------------------------------- #
# Téléchargement + cache disque
# --------------------------------------------------------------------------- #
def fetch_declaration(row: dict, cache_dir: str, s: requests.Session) -> dict:
    fn = os.path.join(cache_dir, f"{row['numero']}.json")
    if os.path.exists(fn):
        try:
            with open(fn, encoding="utf-8") as fh:
                d = json.load(fh)
            if d.get("v") == PARSE_VERSION:
                return d
        except (OSError, ValueError):
            pass
    r = s.get(BASE + "documents/" + row["path"], timeout=40)
    r.raise_for_status()
    d = parse_declaration(pdf_text(r.content))
    d["v"] = PARSE_VERSION
    with open(fn, "w", encoding="utf-8") as fh:
        json.dump(d, fh, ensure_ascii=False)
    return d


def load_transactions(rows: list[dict], cache_dir: str, workers: int = 8,
                      progress: Optional[Callable[[int, int], None]] = None
                      ) -> tuple[pd.DataFrame, pd.DataFrame]:
    """→ (transactions retenues, diagnostic : une ligne par opération lue, avec motif d'exclusion)."""
    os.makedirs(cache_dir, exist_ok=True)
    s = session()
    lines = []
    with ThreadPoolExecutor(workers) as ex:
        futs = {ex.submit(fetch_declaration, r, cache_dir, s): r for r in rows}
        for k, f in enumerate(as_completed(futs), 1):
            r = futs[f]
            if progress:
                progress(k, len(rows))
            base = {"numero": r["numero"], "date_pub": r["date_pub"], "jeton": r["jeton"],
                    "societe": r["societe"], "pdf": BASE + "documents/" + r["path"]}
            try:
                d = f.result()
            except Exception as e:  # noqa: BLE001
                lines.append({**base, "motif": f"lecture impossible : {str(e)[:120]}"})
                continue
            if not d["tx"]:
                lines.append({**base, "motif": "aucune transaction reconnue dans le PDF"})
                continue
            pk, pname, role = person_key(d["declarant"])
            corr = not norm(d["modification"]).lower().startswith("notification initiale")
            mandate = bool(re.search(r"mandat", norm(d["commentaire"]), re.I))
            for tx in d["tx"]:
                side, why = side_of(tx)
                lines.append({**base, "isin": d["isin"], "declarant": pname, "fonction": role, "person": pk,
                              "date": tx["date"], "nature": tx["nature"], "instrument": tx["instrument"],
                              "lieu": tx["lieu"], "prix": tx["prix"], "volume": tx["volume"],
                              "devise": tx["devise"], "side": side, "motif": why, "correction": corr,
                              "mandate": mandate, "commentaire": d["commentaire"][:300]})
    diag = pd.DataFrame(lines)
    if diag.empty:
        return diag, diag
    for c in ("side", "motif", "date"):
        if c not in diag:
            diag[c] = ""
    diag["side"] = diag["side"].fillna("")
    tx = diag[(diag["side"] != "") & diag["date"].notna() & (diag["date"] != "")].copy()
    tx["value"] = tx["prix"] * tx["volume"]
    # correction : remplace la déclaration initiale de la même opération (même déclarant, date, sens, volume)
    tx["_k"] = (tx["jeton"] + "|" + tx["person"] + "|" + tx["date"] + "|" + tx["side"] + "|"
                + tx["volume"].round(0).astype("Int64").astype(str))
    latest = tx.sort_values(["date_pub", "numero"]).groupby("_k")["numero"].transform("last")
    tx = tx[tx["numero"] == latest.reindex(tx.index)].drop(columns="_k")    # deux ordres identiques d'un même PDF restent distincts
    tx["date"] = pd.to_datetime(tx["date"], errors="coerce")
    return tx.reset_index(drop=True), diag
