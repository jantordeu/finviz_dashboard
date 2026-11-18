#!/usr/bin/env python3
"""
zb_fetch.py - telecharge les pages Zonebourse (Valorisation + Finances) avec un vrai navigateur Chrome,
comme une session d'abonne ordinaire, et les enregistre en HTML pour finviz_dashboard.py.

Installation (une fois) :
    pip install playwright
    (Google Chrome doit etre installe ; sinon :  playwright install chromium)

Utilisation :
    python zb_fetch.py login          ouvre Chrome : connectez-vous a Zonebourse (compte abonne), puis fermez la
                                      fenetre. Le mot de passe n'est ni lu ni enregistre par ce script : la session
                                      reste dans le profil du navigateur (dossier zb_profile/).
    python zb_fetch.py status         verifie que la session est reconnue comme abonne
    python zb_fetch.py list           liste les valeurs (zonebourse_urls.csv) et l'age des pages deja telechargees
    python zb_fetch.py fetch          telecharge toutes les valeurs dont les pages ont plus de 12 h
    python zb_fetch.py fetch AXA SANOFI --max-age 0      valeurs choisies (extrait du nom), toujours retelechargees

Sortie : zb_pages/<valeur>/valorisation.html et finances.html (+ zb_pages/_journal.csv).

Prudence : les conditions de Zonebourse/MarketScreener peuvent interdire l'extraction automatique. Le script va a
rythme humain (pause aleatoire de 5 a 11 s entre pages, 2 pages par valeur) et s'arrete des qu'il voit un blocage
(403, 429, captcha). Usage personnel seulement.
"""
from __future__ import annotations

import argparse
import csv
import os
import random
import re
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
URLS_FILE = HERE / "zonebourse_urls.csv"
USER_URLS_FILE = HERE / "zonebourse_urls_perso.csv"   # adresses ajoutees depuis le dashboard (jamais ecrasees par une mise a jour)
PROFILE_DIR = Path(os.environ.get("ZB_PROFILE", HERE / "zb_profile"))
PAGES_DIR = Path(os.environ.get("ZB_PAGES", HERE / "zb_pages"))
TABS = ("valorisation", "finances")
TITLE_MARK = {"valorisation": "Ratios de Valorisation", "finances": "Prévisions"}   # idem finviz_dashboard.page_identity
HOME = "https://www.zonebourse.com/"
PAUSE = (5.0, 11.0)
MAX_PAGES_PER_RUN = 120


def slugify(nom: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", nom.lower()).strip("-")


def load_urls(path: Path | None = None) -> dict[str, tuple[str, str]]:
    """{slug: (nom, url de base terminee par /)} : fichier fourni + adresses personnelles, doublons ignores."""
    return {r["slug"]: (r["nom"], r["url"]) for r in load_rows(path)}


URL_RE = re.compile(r"^https://(?:www\.)?zonebourse\.com/cours/action/([A-Za-z0-9-]+-\d+)/?")
_LEGAL = {"sa", "se", "nv", "n", "v", "s", "a", "plc", "inc", "corp", "corporation", "co", "ltd", "ag", "holdings",
          "holding", "class", "adr", "group", "groupe", "cgde", "company", "the"}


def norm_name(nom: str) -> list[str]:
    """Nom de societe -> mots significatifs (sans accents, ponctuation ni forme juridique)."""
    import unicodedata
    t = unicodedata.normalize("NFKD", nom or "").encode("ascii", "ignore").decode().lower()
    t = t.replace("&", " and ")
    return [w for w in re.sub(r"[^a-z0-9]+", " ", t).split() if w not in _LEGAL]


def _read_rows(path: Path) -> list[dict]:
    out = []
    with open(path, encoding="utf-8-sig", newline="") as fh:
        for row in csv.DictReader(fh, delimiter=";"):
            nom, url = (row.get("nom") or "").strip(), (row.get("url") or "").strip()
            if nom and url:
                out.append({"nom": nom, "url": url.rstrip("/") + "/", "slug": slugify(nom),
                            "ticker": (row.get("ticker") or "").strip().upper()})
    return out


def load_rows(path: Path | None = None) -> list[dict]:
    """Lignes [{nom, url, ticker, slug}] : adresses personnelles (zonebourse_urls_perso.csv) puis fichier fourni
    (zonebourse_urls.csv). Une adresse personnelle remplace la ligne fournie de meme adresse."""
    files = [path] if path else [f for f in (USER_URLS_FILE, URLS_FILE) if f.exists()]
    out, slugs, urls = [], set(), set()
    for f in files:
        for r in _read_rows(f):
            if r["slug"] in slugs or r["url"].lower() in urls:
                continue
            slugs.add(r["slug"]); urls.add(r["url"].lower())
            out.append(r)
    return out


def match_position(ticker: str, nom: str, rows: list[dict] | None = None) -> dict | None:
    """Ligne du fichier d'adresses correspondant a une position : 1) colonne ticker, 2) nom identique,
    3) l'un des noms commence par l'autre (LVMH / LVMH Moet Hennessy), 4) noms tres proches."""
    import difflib
    rows = load_rows() if rows is None else rows
    tk = (ticker or "").strip().upper()
    for r in rows:
        if tk and r["ticker"] == tk:
            return r
    w = norm_name(nom)
    if not w:
        return None
    best, score = None, 0.0
    for r in rows:
        if r["ticker"]:                                  # ligne deja attribuee a un autre ticker
            continue
        x = norm_name(r["nom"])
        if not x:
            continue
        if x == w:
            return r
        n = min(len(x), len(w))
        sc = 0.9 if x[:n] == w[:n] and len(" ".join(x[:n])) >= 4 else \
            difflib.SequenceMatcher(None, " ".join(x), " ".join(w)).ratio()
        if sc > score:
            best, score = r, sc
    return best if score >= 0.88 else None


def add_url(nom: str, url: str, ticker: str = "", path: Path | None = None) -> dict:
    """Ajoute une adresse (ou attache un ticker a une adresse connue) dans zonebourse_urls_perso.csv.
    Leve ValueError si l'adresse n'est pas une fiche Zonebourse (https://www.zonebourse.com/cours/action/NOM-1234/)."""
    m = URL_RE.match((url or "").strip())
    if not m:
        raise ValueError("adresse attendue : https://www.zonebourse.com/cours/action/NOM-1234/")
    base = f"https://www.zonebourse.com/cours/action/{m.group(1)}/"
    path = path or USER_URLS_FILE
    tk = (ticker or "").strip().upper()
    mine = _read_rows(path) if path.exists() else []
    known = next((r for r in load_rows() if r["url"].lower() == base.lower()), None) if path == USER_URLS_FILE else None
    for r in mine:
        if tk and r["ticker"] == tk and r["url"].lower() != base.lower():
            r["ticker"] = ""                              # ce ticker pointait vers une autre adresse : corrige
    row = next((r for r in mine if r["url"].lower() == base.lower()), None)
    if row:
        row["ticker"] = tk or row["ticker"]
    else:
        nom = (known["nom"] if known else (nom or "").strip()) or re.sub(r"-\d+$", "", m.group(1)).replace("-", " ").title()
        row = {"nom": nom.replace(";", " "), "url": base, "ticker": tk, "slug": slugify(nom)}
        mine.append(row)
    tmp = path.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["nom", "url", "ticker"])
        for r in mine:
            w.writerow([r["nom"], r["url"], r["ticker"]])
    os.replace(tmp, path)
    return row


def page_path(slug: str, tab: str) -> Path:
    return PAGES_DIR / slug / f"{tab}.html"


def age_hours(path: Path) -> float | None:
    return (time.time() - path.stat().st_mtime) / 3600 if path.exists() else None


def check_page(html: str, tab: str) -> tuple[bool, str]:
    """Page complete et reconnue par le dashboard ? -> (ok, raison)."""
    m = re.search(r"<title>(.*?)</title>", html, re.S)
    title = re.sub(r"\s+", " ", m.group(1)).strip() if m else ""
    low = html.lower()
    if any(k in low for k in ("captcha", "access denied", "are you a robot", "cf-challenge", "just a moment")) \
            and len(html) < 200_000:
        return False, "blocage anti-robot (captcha / challenge)"
    if TITLE_MARK[tab] not in title:
        return False, f"titre inattendu « {title[:60]} » (page non reconnue)"
    if not re.search(r"\d{2}:\d{2}:\d{2}\s+\d{2}/\d{2}/\d{4}", html):
        return False, "horodatage de cotation absent (page incomplete)"
    if len(html) < 80_000:
        return False, f"page trop courte ({len(html) // 1000} Ko)"
    return True, "ok"


def logged_in(html: str) -> bool | None:
    """True / False si un indice de connexion est visible dans la page, None si on ne peut pas conclure."""
    low = html.lower()
    if any(k in low for k in ("déconnexion", "deconnexion", "/logout", "se déconnecter", "sign out")):
        return True
    if any(k in low for k in ("se connecter", "connexion", "sign in", "/login")):
        return False
    return None


def profile_busy() -> bool:
    """Le profil est-il deja ouvert par un Chrome (fenetre de connexion restee ouverte) ?"""
    lock = PROFILE_DIR / "SingletonLock"
    if not os.path.lexists(lock):
        return False
    try:
        pid = int(os.readlink(lock).rsplit("-", 1)[1])
        os.kill(pid, 0)                                 # le processus existe-t-il encore ?
        return True
    except (OSError, ValueError, IndexError):
        try:
            os.unlink(lock)                             # verrou perime (Chrome ferme brutalement)
        except OSError:
            pass
        return False


class FetchError(Exception):
    pass


def _launch(pw, headless: bool):
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    if profile_busy():
        raise FetchError("La fenetre Chrome de connexion Zonebourse est encore ouverte : quittez ce Chrome "
                         "(Cmd+Q, pas seulement la croix), puis relancez.")
    kw = dict(user_data_dir=str(PROFILE_DIR), headless=headless, locale="fr-FR", timezone_id="Europe/Paris",
              viewport={"width": 1440, "height": 900}, args=["--disable-blink-features=AutomationControlled"],
              # bac a sable de Chrome conserve (sinon bandeau « --no-sandbox ») ; pas de bandeau « logiciel de test »
              chromium_sandbox=not os.environ.get("ZB_NO_SANDBOX"), ignore_default_args=["--enable-automation"])
    try:
        return pw.chromium.launch_persistent_context(channel="chrome", **kw)        # vrai Chrome installe
    except Exception as e1:
        try:
            return pw.chromium.launch_persistent_context(**kw)                      # Chromium de Playwright
        except Exception:
            raise FetchError("Impossible d'ouvrir Chrome avec le profil Zonebourse (" + str(e1).splitlines()[0][:120]
                             + "). Verifiez que Google Chrome est installe et qu'aucune fenetre de connexion "
                             "Zonebourse n'est restee ouverte ; sinon lancez :  playwright install chromium")


def _human_scroll(page):
    """Defile la page par paliers (charge les tableaux differes), comme un lecteur."""
    try:
        h = page.evaluate("document.body.scrollHeight")
        y = 0
        while y < h:
            y += random.randint(500, 900)
            page.mouse.wheel(0, random.randint(500, 900))
            page.wait_for_timeout(random.randint(150, 400))
            h = page.evaluate("document.body.scrollHeight")
            if y > 30000:
                break
        page.evaluate("window.scrollTo(0, 0)")
    except Exception:
        pass


def cmd_login(a) -> int:
    from playwright.sync_api import sync_playwright
    print("Une fenetre Chrome s'ouvre : acceptez les cookies, connectez-vous a Zonebourse avec votre compte abonne,\n"
          "puis quittez ce Chrome avec Cmd+Q. (Votre mot de passe n'est pas lu par ce script.)")
    with sync_playwright() as pw:
        ctx = _launch(pw, headless=False)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(HOME, wait_until="domcontentloaded")
        try:
            while ctx.pages:
                time.sleep(1)
        except KeyboardInterrupt:
            pass
        try:
            ctx.close()
        except Exception:
            pass
    print("Session enregistree dans", PROFILE_DIR)
    return 0


def cmd_status(a) -> int:
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        ctx = _launch(pw, headless=a.headless)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        resp = page.goto(HOME, wait_until="domcontentloaded", timeout=45000)
        page.wait_for_timeout(2500)
        st = resp.status if resp else None
        li = logged_in(page.content())
        ctx.close()
    if st and st >= 400:
        print(f"BLOQUE : HTTP {st}")
        return 2
    msg = {True: "session reconnue (connecte)", False: "NON connecte : lancez  python zb_fetch.py login",
           None: "etat de connexion indeterminable (verifiez avec 'login')"}[li]
    print("STATUS", msg)
    return 0 if li else 3


def cmd_list(a) -> int:
    urls = load_urls(Path(a.urls) if a.urls else None)
    for slug, (nom, _u) in urls.items():
        ages = []
        for t in TABS:
            h = age_hours(page_path(slug, t))
            ages.append(f"{t[:4]} {h:.0f} h" if h is not None else f"{t[:4]} -")
        print(f"{nom:34s} {'  '.join(ages)}")
    return 0


def select(urls, names):
    if not names:
        return list(urls)
    wanted = [slugify(n) for n in names]
    return [s for s in urls if any(w in s for w in wanted)]


def cmd_fetch(a) -> int:
    from playwright.sync_api import sync_playwright
    urls = load_urls(Path(a.urls) if a.urls else None)
    todo = []
    slugs = [x for x in (a.slugs or "").split(",") if x in urls] if a.slugs else select(urls, a.names)
    for slug in slugs:
        for tab in TABS:
            h = age_hours(page_path(slug, tab))
            if h is None or h > a.max_age:
                todo.append((slug, tab))
    todo = todo[: a.limit]
    if not todo:
        print("DONE rien a telecharger (pages recentes)")
        return 0
    print(f"START {len(todo)} page(s)", flush=True)
    PAGES_DIR.mkdir(parents=True, exist_ok=True)
    nok, nerr, blocked = 0, 0, False
    with sync_playwright() as pw:
        ctx = _launch(pw, headless=a.headless)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        for i, (slug, tab) in enumerate(todo, 1):
            nom, base = urls[slug]
            url = base + tab + "/"
            status, reason, html = None, "", ""
            try:
                resp = page.goto(url, wait_until="domcontentloaded", timeout=60000)
                status = resp.status if resp else None
                try:
                    page.wait_for_load_state("networkidle", timeout=20000)
                except Exception:
                    pass
                _human_scroll(page)
                page.wait_for_timeout(random.randint(600, 1500))
                html = page.content()
            except Exception as e:
                reason = f"{type(e).__name__}: {str(e)[:80]}"
            if status in (401, 403, 429):
                print(f"BLOCK {i}/{len(todo)} {nom} {tab} HTTP {status} - arret (compte non reconnu ou limitation)",
                      flush=True)
                blocked = True
                break
            ok, why = check_page(html, tab) if html else (False, reason or "page vide")
            if ok:
                if logged_in(html) is False:
                    why = "ok (attention : connexion non confirmee, donnees possiblement tronquees)"
                p = page_path(slug, tab)
                p.parent.mkdir(parents=True, exist_ok=True)
                tmp = p.with_suffix(".tmp")
                tmp.write_text(html, encoding="utf-8")
                os.replace(tmp, p)
                nok += 1
            else:
                nerr += 1
            with open(PAGES_DIR / "_journal.csv", "a", encoding="utf-8") as fh:
                fh.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')};{slug};{tab};{status};{'ok' if ok else 'ERR'};{why}\n")
            print(f"{'OK  ' if ok else 'ERR '} {i}/{len(todo)} {nom} {tab} - {why}", flush=True)
            if i < len(todo):
                time.sleep(random.uniform(*PAUSE))
        ctx.close()
    print(f"DONE {nok} page(s) enregistree(s), {nerr} en erreur" + (", BLOQUE" if blocked else ""), flush=True)
    return 2 if blocked else (1 if nerr and not nok else 0)


def main() -> int:
    ap = argparse.ArgumentParser(description="Telecharge les pages Zonebourse (session d'abonne, navigateur reel)")
    sub = ap.add_subparsers(dest="cmd", required=True)
    for n in ("login", "status", "list", "fetch"):
        s = sub.add_parser(n)
        s.add_argument("--urls", help="fichier nom;url (defaut : zonebourse_urls.csv)")
        s.add_argument("--headless", action="store_true", help="navigateur invisible (moins fiable)")
        if n == "fetch":
            s.add_argument("names", nargs="*", help="extraits de nom ; vide = toutes les valeurs")
            s.add_argument("--slugs", help="valeurs exactes (identifiants separes par des virgules)")
            s.add_argument("--max-age", type=float, default=12.0, help="ne retelecharge pas les pages plus recentes (heures)")
            s.add_argument("--limit", type=int, default=MAX_PAGES_PER_RUN, help="nombre maximal de pages par lancement")
    a = ap.parse_args()
    try:
        return {"login": cmd_login, "status": cmd_status, "list": cmd_list, "fetch": cmd_fetch}[a.cmd](a)
    except FetchError as e:
        print("ERREUR", e, flush=True)
        return 4


if __name__ == "__main__":
    sys.exit(main())
