#!/usr/bin/env python3
"""
data_yfinance.py - telecharge les valeurs du CAC 40 depuis Yahoo Finance (yfinance).

Fichiers du projet (dans le meme dossier) :
    data_yfinance.py   ce script
    data.txt           liste des tickers, au format  ["AC.PA", "AI.PA", ...]
                       (un ticker par ligne, sans crochets, fonctionne aussi)
    requirements.txt   dependances :  pip install -r requirements.txt

Sortie (creee a cote du script) :
    data/<TICKER>.csv      un CSV par valeur : date, open, high, low, close, volume

Lancement dans PyCharm : clic droit sur data_yfinance.py -> Run. Aucun argument requis.
En ligne de commande :  python data_yfinance.py --period 5y --out-dir data

Notes :
  - Les prix sont AJUSTES (dividendes, divisions d'actions) : auto_adjust=True. Ils different donc
    legerement des cotations brutes d'un site de bourse.
  - Les symboles Yahoo peuvent changer ou etre errones : le script liste ceux qui ne renvoient rien.
  - Les doublons de data.txt sont ignores ; un indice ecrit "^FCHI.PA" est corrige en "^FCHI"
    (les indices Yahoo n'ont pas de suffixe .PA). Un indice n'a pas de volume (valeur 0).
  - Ce n'est pas un conseil en investissement.

Utilise aussi par finviz_dashboard.py (section 12 « Graphique » de chaque valeur etudiee) : le dashboard
importe download() et csv_name() depuis ce fichier et ecrit ses CSV dans data/.
"""
from __future__ import annotations

import argparse
import ast
import sys
import time
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent

# --- Parametres par defaut (modifiables ici ou en ligne de commande) -------------------------------
TICKERS_FILE = HERE / "data.txt"
OUT_DIR = HERE / "data"
PERIOD = "4y"        # 1y, 2y, 3y, 4y, 5y, 10y, max
INTERVAL = "1d"
MIN_BARS = 150       # en dessous, la valeur est signalee (trop courte pour l'analyse de ranges)
# -------------------------------------------------------------------------------------------------

PRICE_COLS = ["open", "high", "low", "close"]
FIELDS = PRICE_COLS + ["volume"]


def normalize_ticker(t: str) -> str:
    """Majuscules ; un indice Yahoo s'ecrit ^FCHI (pas ^FCHI.PA)."""
    t = t.strip().upper()
    if t.startswith("^") and t.endswith(".PA"):
        t = t[: -len(".PA")]
    return t


def read_tickers(path: Path):
    """Lit data.txt : liste Python  ["AC.PA", ...]  (virgule finale acceptee) ou un ticker par ligne.
    Retourne (tickers uniques dans l'ordre, doublons ignores, corrections [(avant, apres)])."""
    text = path.read_text(encoding="utf-8").strip()
    if not text:
        raise ValueError(f"{path.name} est vide.")
    try:
        parsed = ast.literal_eval(text)
        if isinstance(parsed, (list, tuple)):
            raw = [str(t) for t in parsed]
        else:
            raise ValueError
    except (ValueError, SyntaxError):
        raw = [line.strip().strip('",\'[]') for line in text.splitlines()]
    out: list[str] = []
    dups: list[str] = []
    fixes: list[tuple[str, str]] = []
    for r in raw:
        before = r.strip().upper()
        if not before:
            continue
        t = normalize_ticker(before)
        if t != before:
            fixes.append((before, t))
        if t in out:
            dups.append(t)
        else:
            out.append(t)
    if not out:
        raise ValueError(f"Aucun ticker lisible dans {path.name}.")
    return out, dups, fixes


def csv_name(ticker: str) -> str:
    """Nom de fichier sur : BNP.PA -> BNP.csv ; ^FCHI -> FCHI.csv."""
    name = ticker[: -len(".PA")] if ticker.endswith(".PA") else ticker
    name = "".join(c if (c.isalnum() or c in "-_") else "_" for c in name.lstrip("^"))
    return f"{name}.csv"


def clean(df: pd.DataFrame | None) -> pd.DataFrame | None:
    """Normalise un DataFrame OHLCV : colonnes en minuscules, index = date sans fuseau, sans doublons."""
    if df is None or len(df) == 0:
        return None
    df = df.copy()
    df.columns = [str(c).lower() for c in df.columns]
    if not all(c in df.columns for c in PRICE_COLS):
        return None
    df = df[[c for c in FIELDS if c in df.columns]].dropna(subset=PRICE_COLS)
    idx = pd.to_datetime(df.index)
    if getattr(idx, "tz", None) is not None:
        idx = idx.tz_localize(None)
    df.index = idx.normalize()
    df.index.name = "date"
    df = df[~df.index.duplicated(keep="last")].sort_index()
    return df if len(df) else None


def split_batch(data: pd.DataFrame | None, tickers: list[str]) -> dict[str, pd.DataFrame]:
    """Decoupe le resultat de yf.download (colonnes MultiIndex) en {ticker: DataFrame propre}."""
    out: dict[str, pd.DataFrame] = {}
    if data is None or len(data) == 0:
        return out
    if isinstance(data.columns, pd.MultiIndex):
        top = {str(x).lower() for x in data.columns.get_level_values(0)}
        if top & {"close", "open"}:  # niveaux (champ, ticker) -> on inverse en (ticker, champ)
            data = data.swaplevel(0, 1, axis=1)
        present = set(data.columns.get_level_values(0))
        for t in tickers:
            if t in present:
                d = clean(data[t])
                if d is not None:
                    out[t] = d
    elif len(tickers) == 1:
        d = clean(data)
        if d is not None:
            out[tickers[0]] = d
    return out


def download(tickers: list[str], period: str, interval: str):
    """Retourne ({ticker: DataFrame}, [tickers en echec])."""
    try:
        import yfinance as yf
    except ImportError:
        sys.exit("yfinance n'est pas installe. Lancez :  pip install -r requirements.txt")

    print(f"Telechargement de {len(tickers)} valeurs (periode {period}, intervalle {interval})...")
    try:
        data = yf.download(tickers, period=period, interval=interval, group_by="ticker",
                           auto_adjust=True, progress=False, threads=True)
    except Exception as e:  # reseau, limitation de debit...
        print(f"  echec de la requete groupee : {e}")
        data = None
    got = split_batch(data, tickers)

    manquants = [t for t in tickers if t not in got]
    if manquants and got:
        print(f"  {len(manquants)} valeur(s) absente(s) de la requete groupee : nouvel essai une par une...")
    for t in manquants:
        try:
            d = yf.download(t, period=period, interval=interval, auto_adjust=True, progress=False)
            if isinstance(d.columns, pd.MultiIndex):
                d.columns = d.columns.get_level_values(0)
            d = clean(d)
            if d is not None:
                got[t] = d
        except Exception as e:
            print(f"  {t} : {e}")
        time.sleep(0.3)  # evite de saturer Yahoo
    return got, [t for t in tickers if t not in got]


def save(got: dict[str, pd.DataFrame], out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    for t, d in got.items():
        d.to_csv(out_dir / csv_name(t))


def main() -> int:
    ap = argparse.ArgumentParser(description="Telecharge le CAC 40 (yfinance) en CSV par valeur")
    ap.add_argument("--tickers-file", type=Path, default=TICKERS_FILE)
    ap.add_argument("--out-dir", type=Path, default=OUT_DIR)
    ap.add_argument("--period", default=PERIOD)
    ap.add_argument("--interval", default=INTERVAL)
    a = ap.parse_args()

    if not a.tickers_file.exists():
        print(f"Fichier introuvable : {a.tickers_file}")
        return 1
    tickers, dups, fixes = read_tickers(a.tickers_file)
    print(f"{len(tickers)} tickers uniques lus dans {a.tickers_file.name}")
    if dups:
        print(f"  doublons ignores : {', '.join(sorted(set(dups)))}")
    for before, after in fixes:
        print(f"  corrige : {before} -> {after}")

    got, failed = download(tickers, a.period, a.interval)
    if not got:
        print("\nAucune donnee recue. Verifiez la connexion Internet, mettez yfinance a jour "
              "(pip install -U yfinance) ou reessayez plus tard (limitation de debit de Yahoo).")
        return 1

    save(got, a.out_dir)
    print(f"\n{len(got)}/{len(tickers)} valeurs ecrites dans '{a.out_dir}'")
    short = []
    for t, d in got.items():
        print(f"  {t:<8} {len(d):>5} barres   {d.index.min().date()} -> {d.index.max().date()}")
        if len(d) < MIN_BARS:
            short.append(t)
    if short:
        print(f"\nAttention : moins de {MIN_BARS} barres pour : {', '.join(short)}")
    if failed:
        print(f"\nSans donnees ({len(failed)}) : {', '.join(failed)}")
        print("  -> symbole Yahoo perime ou errone ? A verifier sur finance.yahoo.com")
    return 0


if __name__ == "__main__":
    sys.exit(main())
