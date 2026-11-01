"""
Scraper Dataroma : croise la page « S&P500 Grid » (achats/ventes des superinvestors)
avec le tableau « Total insider Buys/Sells last 6 months » de chaque fiche action.

Utilisation en ligne de commande :
    python dataroma_scraper.py                 # page Last qtr buys, solde > 0
    python dataroma_scraper.py --grid h --all  # Last 6 months buys, toutes les valeurs
"""
from __future__ import annotations

import argparse
import io
import re
import time
from datetime import date
from typing import Callable, Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE_URL = "https://www.dataroma.com/m/"

# Pages de la grille S&P 500 (paramètre ?s=)
GRIDS = {
    "q": "Last qtr buys",
    "h": "Last 6 months buys",
    "sq": "Last qtr sells",
    "sh": "Last 6 months sells",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

COLUMNS = {
    "ticker": "Ticker",
    "name": "Société",
    "sector": "Secteur",
    "si_count": "Nb superinvestors",
    "ins_buy_n": "Achats initiés (nb)",
    "ins_buy_usd": "Achats initiés ($)",
    "ins_sell_n": "Ventes initiés (nb)",
    "ins_sell_usd": "Ventes initiés ($)",
    "balance_usd": "Solde ($)",
    "url": "Lien Dataroma",
}


# --------------------------------------------------------------------------- #
# Réseau
# --------------------------------------------------------------------------- #
def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    return s


def fetch(session: requests.Session, url: str, params: dict | None = None,
          retries: int = 3, pause: float = 2.0) -> str:
    last_err: Exception | None = None
    for attempt in range(retries):
        try:
            r = session.get(url, params=params, timeout=20)
            r.raise_for_status()
            return r.text
        except requests.RequestException as e:
            last_err = e
            time.sleep(pause * (attempt + 1))
    raise RuntimeError(f"Échec de téléchargement {url} {params or ''} : {last_err}")


# --------------------------------------------------------------------------- #
# Parsing (fonctions pures, testables sur des pages HTML enregistrées)
# --------------------------------------------------------------------------- #
def _num(text: str | None) -> float:
    t = re.sub(r"[^0-9.\-]", "", text or "")
    try:
        return float(t) if t not in ("", "-", ".") else 0.0
    except ValueError:
        return 0.0


def parse_grid(html: str) -> list[dict]:
    """Retourne [{ticker, name, sector, si_count}] depuis une page grid.php."""
    soup = BeautifulSoup(html, "html.parser")
    rows = []
    for a in soup.select('a[href*="stock.php?sym="]'):
        div = a.parent.find("div") if a.parent else None
        if div is None:
            continue
        text = div.get_text(" ", strip=True)
        m = re.search(r"(?:Buys/Adds|Sells/Reductions|Sells|Buys)[^:]*:\s*(\d+)", text)
        sector = re.search(r"\(([^)]+)\)", text)
        rows.append({
            "ticker": a.get_text(strip=True),
            "name": re.sub(r"\s+", " ", div.b.get_text(strip=True)) if div.b else "",
            "sector": sector.group(1) if sector else "",
            "si_count": int(m.group(1)) if m else 0,
        })
    return rows


def parse_insider_summary(html: str) -> Optional[dict]:
    """Lit le tableau #ins_sum d'une page stock.php. None si absent."""
    soup = BeautifulSoup(html, "html.parser")
    table = soup.select_one("#ins_sum")
    if table is None:
        return None

    def row(cls: str) -> tuple[float, float]:
        tds = table.select(f"tr.{cls} td.num")
        if len(tds) < 2:
            return 0.0, 0.0
        return _num(tds[0].get_text()), _num(tds[1].get_text())

    bn, bv = row("buys")
    sn, sv = row("sells")
    return {"ins_buy_n": int(bn), "ins_buy_usd": bv,
            "ins_sell_n": int(sn), "ins_sell_usd": sv}


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #
def scrape(grid: str = "q",
           min_si_count: int = 2,
           delay: float = 0.5,
           limit: Optional[int] = None,
           progress: Optional[Callable[[int, int, str], None]] = None) -> pd.DataFrame:
    """Télécharge la grille puis le résumé insider de chaque valeur retenue."""
    if grid not in GRIDS:
        raise ValueError(f"grid doit être parmi {list(GRIDS)}")
    session = make_session()
    stocks = [s for s in parse_grid(fetch(session, BASE_URL + "grid.php", {"s": grid}))
              if s["si_count"] >= min_si_count]
    if limit:
        stocks = stocks[:limit]

    out = []
    for i, s in enumerate(stocks, start=1):
        if progress:
            progress(i, len(stocks), s["ticker"])
        try:
            summary = parse_insider_summary(
                fetch(session, BASE_URL + "stock.php", {"sym": s["ticker"]}))
        except RuntimeError:
            summary = None
        s.update(summary or {"ins_buy_n": None, "ins_buy_usd": None,
                             "ins_sell_n": None, "ins_sell_usd": None})
        s["url"] = f"{BASE_URL}stock.php?sym={s['ticker']}"
        out.append(s)
        time.sleep(delay)

    df = pd.DataFrame(out)
    if df.empty:
        return df
    df["balance_usd"] = df["ins_buy_usd"].fillna(0) - df["ins_sell_usd"].fillna(0)
    return df[list(COLUMNS)]


def positive_balance(df: pd.DataFrame) -> pd.DataFrame:
    return df[df["balance_usd"] > 0].reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Exports
# --------------------------------------------------------------------------- #
def to_display(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns=COLUMNS)


def to_excel_bytes(df: pd.DataFrame, title: str = "Dataroma") -> bytes:
    from openpyxl.styles import Alignment, Font, PatternFill

    buf = io.BytesIO()
    disp = to_display(df)
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        disp.to_excel(xw, index=False, sheet_name="Résultats")
        ws = xw.sheets["Résultats"]
        hdr_font = Font(name="Arial", bold=True, color="FFFFFF")
        for c in ws[1]:
            c.font = hdr_font
            c.fill = PatternFill("solid", fgColor="1F4E78")
            c.alignment = Alignment(wrap_text=True, vertical="center")
        cols = list(disp.columns)
        ci = {name: idx + 1 for idx, name in enumerate(cols)}
        buy_col = ws.cell(1, ci[COLUMNS["ins_buy_usd"]]).column_letter
        sell_col = ws.cell(1, ci[COLUMNS["ins_sell_usd"]]).column_letter
        for r in range(2, ws.max_row + 1):
            for c in ws[r]:
                c.font = Font(name="Arial")
            # Solde recalculé par formule
            ws.cell(r, ci[COLUMNS["balance_usd"]]).value = f"={buy_col}{r}-{sell_col}{r}"
            for key in ("ins_buy_usd", "ins_sell_usd", "balance_usd"):
                ws.cell(r, ci[COLUMNS[key]]).number_format = "$#,##0;($#,##0);-"
            link = ws.cell(r, ci[COLUMNS["url"]])
            link.hyperlink = link.value
            link.font = Font(name="Arial", color="0563C1", underline="single")
        widths = [9, 34, 22, 14, 12, 16, 12, 16, 16, 46]
        for idx, w in enumerate(widths[:len(cols)], start=1):
            ws.column_dimensions[ws.cell(1, idx).column_letter].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        note = ws.max_row + 2
        ws.cell(note, 1, f"Source : dataroma.com – {title} – relevé le "
                         f"{date.today():%d/%m/%Y}. Solde = achats $ − ventes $ "
                         "(insiders, 6 derniers mois).").font = Font(name="Arial", italic=True, size=9)
    return buf.getvalue()


def to_html_page(df: pd.DataFrame, title: str = "Dataroma") -> str:
    disp = to_display(df).copy()
    disp[COLUMNS["url"]] = disp[COLUMNS["url"]].map(
        lambda u: f'<a href="{u}" target="_blank">fiche</a>')
    money = [COLUMNS[k] for k in ("ins_buy_usd", "ins_sell_usd", "balance_usd")]
    fmt = {c: (lambda v: "" if pd.isna(v) else f"${v:,.0f}".replace(",", " ")) for c in money}
    styler = (disp.style.format(fmt)
              .map(lambda v: "color:#1a7f37;font-weight:600" if isinstance(v, (int, float)) and v > 0
                   else ("color:#cf222e" if isinstance(v, (int, float)) and v < 0 else ""),
                   subset=[COLUMNS["balance_usd"]])
              .hide(axis="index"))
    table = styler.to_html(escape=False)
    today = f"{date.today():%d/%m/%Y}"
    return f"""<!doctype html>
<html lang="fr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
 body{{font-family:Arial,Helvetica,sans-serif;margin:24px;color:#1f2328;background:#fff}}
 h1{{font-size:20px;margin:0 0 4px}} p.sub{{color:#57606a;margin:0 0 16px;font-size:13px}}
 .wrap{{overflow-x:auto}}
 table{{border-collapse:collapse;font-size:13px;width:100%}}
 th{{background:#1f4e78;color:#fff;text-align:left;padding:8px;position:sticky;top:0;cursor:pointer}}
 td{{padding:6px 8px;border-bottom:1px solid #d0d7de;white-space:nowrap}}
 tr:nth-child(even) td{{background:#f6f8fa}}
 td:nth-child(n+4){{text-align:right}}
</style></head><body>
<h1>{title}</h1>
<p class="sub">{len(df)} valeurs · Source : dataroma.com · relevé le {today} ·
Solde = achats − ventes d'initiés sur 6 mois · cliquer un en-tête pour trier</p>
<div class="wrap">{table}</div>
<script>
document.querySelectorAll('th').forEach((th,i)=>th.addEventListener('click',()=>{{
 const tb=th.closest('table').tBodies[0], rows=[...tb.rows], asc=th.dataset.asc!=='1';
 const v=r=>{{const t=r.cells[i].innerText.replace(/[$\\s]/g,'');const n=parseFloat(t);return isNaN(n)?t:n}};
 rows.sort((a,b)=>{{const x=v(a),y=v(b);return (x>y?1:x<y?-1:0)*(asc?1:-1)}});
 th.dataset.asc=asc?'1':'0'; rows.forEach(r=>tb.appendChild(r));
}}));
</script>
</body></html>"""


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main() -> None:
    p = argparse.ArgumentParser(description="Dataroma : superinvestors × insiders")
    p.add_argument("--grid", default="q", choices=list(GRIDS), help="page de la grille S&P 500")
    p.add_argument("--min", type=int, default=2, help="nb minimum de superinvestors (défaut 2)")
    p.add_argument("--all", action="store_true", help="garder aussi les soldes ≤ 0")
    p.add_argument("--delay", type=float, default=0.5, help="pause entre requêtes (s)")
    p.add_argument("--limit", type=int, default=None, help="limiter le nb de valeurs (test)")
    p.add_argument("--out", default="dataroma_inities", help="préfixe des fichiers de sortie")
    a = p.parse_args()

    df = scrape(a.grid, a.min, a.delay, a.limit,
                progress=lambda i, n, t: print(f"\r[{i}/{n}] {t:<8}", end="", flush=True))
    print()
    if not a.all:
        df = positive_balance(df)
    title = f"Dataroma – {GRIDS[a.grid]} – solde initiés {'(tous)' if a.all else '> 0'}"
    with open(f"{a.out}.xlsx", "wb") as f:
        f.write(to_excel_bytes(df, title))
    with open(f"{a.out}.html", "w", encoding="utf-8") as f:
        f.write(to_html_page(df, title))
    print(f"{len(df)} valeurs → {a.out}.xlsx et {a.out}.html")


if __name__ == "__main__":
    main()
