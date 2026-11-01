"""
Deux analyses SÉPARÉES sur le même univers, puis un tableau final qui les réunit.

Univers : valeurs du S&P 500 de la grille Dataroma (ex. « Last qtr buys »)
          détenues par au moins N superinvestors (≈ 175 valeurs pour N = 2).

A. Dataroma    : tableau « Total insider Buys/Sells last 6 months » de chaque fiche
                 → retenue si solde $ > 0 (méthode du fichier initial).
B. OpenInsider : transactions des N derniers jours, levées d'options exclues
                 → retenue selon le critère choisi (acheteurs distincts, transactions, $).

Tableau final : colonnes du tableau initial Dataroma + indicateurs OpenInsider,
                avec la colonne « Retenue par » (Dataroma, OpenInsider ou les deux).

Ligne de commande :
    python analyse.py
    python analyse.py --critere usd --jours 120 --sans-10pct
    python analyse.py --sans-dataroma          # OpenInsider seul (≈ 5 s)
    python analyse.py --montant-min 500000          # forte conviction
    python analyse.py --univers-si                  # OpenInsider limité aux valeurs des superinvestors
"""
from __future__ import annotations

import argparse
from datetime import date

import pandas as pd

import dataroma_scraper as dr
import openinsider as oi
from exports import to_excel_bytes, to_html_page

# Colonnes du tableau initial Dataroma, puis OpenInsider
LABELS = {
    "ticker": "Ticker",
    "name": "Société",
    "sector": "Secteur",
    # Superinvestors (fonds, déclarations 13F) — Dataroma
    "si_count": "Superinvestors acheteurs (3 mois) · Dataroma",
    # Initiés (dirigeants, administrateurs) — Dataroma, 6 mois
    "ins_buy_n": "Initiés 6 mois · nb achats · Dataroma",
    "ins_buy_usd": "Initiés 6 mois · achats ($) · Dataroma",
    "ins_sell_n": "Initiés 6 mois · nb ventes · Dataroma",
    "ins_sell_usd": "Initiés 6 mois · ventes ($) · Dataroma",
    "dr_balance_usd": "Initiés 6 mois · solde ($) · Dataroma",
    # Initiés (dirigeants, administrateurs) — OpenInsider, période choisie
    "oi_signal": "Initiés période · résumé · OpenInsider",
    "buyers": "Initiés période · acheteurs distincts · OpenInsider",
    "sellers": "Initiés période · vendeurs distincts · OpenInsider",
    "buy_tx": "Initiés période · nb achats · OpenInsider",
    "sell_tx": "Initiés période · nb ventes · OpenInsider",
    "buy_usd": "Initiés période · achats ($) · OpenInsider",
    "sell_usd": "Initiés période · ventes ($) · OpenInsider",
    "balance_usd": "Initiés période · solde ($) · OpenInsider",
    "max_delta_own": "Initiés période · max ΔOwn (%) · OpenInsider",
    "cluster": "Initiés période · achat groupé (nb en 30 j) · OpenInsider",
    "atypical": "Cas atypique · OpenInsider",
    "buyer_names": "Initiés période · noms des acheteurs · OpenInsider",
    "selected_by": "Retenue par",
    "url": "Lien Dataroma",
    "oi_url": "Lien OpenInsider",
}
TX_LABELS = {
    "ticker": "Ticker", "trade_date": "Date transaction", "filing_date": "Date déclaration",
    "insider": "Initié", "title": "Fonction", "trade_type": "Type", "price": "Prix ($)",
    "qty": "Quantité", "owned": "Détenu après", "delta_own": "ΔOwn (%)", "value": "Montant ($)",
}
MONEY = {LABELS[k] for k in ("ins_buy_usd", "ins_sell_usd", "dr_balance_usd",
                             "buy_usd", "sell_usd", "balance_usd")} | {TX_LABELS["value"]}
LINKS = {LABELS["url"], LABELS["oi_url"]}


def sp500(grid: str = "q") -> pd.DataFrame:
    """Les 500 valeurs de la grille Dataroma, avec leur nb de superinvestors."""
    html = dr.fetch(dr.make_session(), dr.BASE_URL + "grid.php", {"s": grid})
    return pd.DataFrame(dr.parse_grid(html))


def universe(grid: str = "q", min_si: int = 1, base: pd.DataFrame | None = None) -> pd.DataFrame:
    df = sp500(grid) if base is None else base
    return df[df["si_count"] >= min_si].reset_index(drop=True)


def run_dataroma(uni: pd.DataFrame, delay: float = 0.5, progress=None) -> pd.DataFrame:
    """Analyse A : résumé initiés 6 mois de chaque fiche Dataroma."""
    s, out = dr.make_session(), []
    for i, t in enumerate(uni["ticker"], start=1):
        if progress:
            progress(i, len(uni), t)
        try:
            summ = dr.parse_insider_summary(dr.fetch(s, dr.BASE_URL + "stock.php", {"sym": t}))
        except RuntimeError:
            summ = None
        out.append({"ticker": t, **(summ or {"ins_buy_n": None, "ins_buy_usd": None,
                                             "ins_sell_n": None, "ins_sell_usd": None})})
        dr.time.sleep(delay)
    df = pd.DataFrame(out, columns=["ticker", "ins_buy_n", "ins_buy_usd", "ins_sell_n", "ins_sell_usd"])
    for c in ("ins_buy_n", "ins_buy_usd", "ins_sell_n", "ins_sell_usd"):
        df[c] = pd.to_numeric(df[c], errors="coerce")       # NaN (case vide) si fiche illisible
    df["dr_balance_usd"] = df["ins_buy_usd"] - df["ins_sell_usd"]
    return df


def fetch_openinsider(uni: pd.DataFrame, days: int = 90, delay: float = 0.5, progress=None,
                      years_back: int = 0, periods_back: int = 0):
    """Téléchargement OpenInsider (levées d'options exclues) → (transactions, journal).
    years_back=1   : même période, un an plus tôt ;
    periods_back=1 : les `days` jours qui précèdent la période actuelle."""
    return oi.fetch_transactions(uni["ticker"].tolist(), days, delay, progress=progress,
                                 period=oi.date_range(days, years_back, periods_back=periods_back))


SENT_ROWS = [("ratio_ins", "Acheteurs / vendeurs distincts (ratio)"),
             ("ratio_tx", "Nb achats / nb ventes (ratio)"),
             ("ratio_usd", "Montant achats / ventes (ratio)"),
             ("acheteurs", "Acheteurs distincts"), ("vendeurs", "Vendeurs distincts"),
             ("achats", "Nb achats"), ("ventes", "Nb ventes"),
             ("achats_usd", "Achats ($)"), ("ventes_usd", "Ventes ($)"),
             ("valeurs_pos", "Valeurs avec acheteurs > vendeurs")]


COMPARE = {"an": "Même période l'an dernier",
           "precedente": "Période précédente (consécutive)",
           "aucune": "Aucune comparaison"}
COMPARE_SHORT = {"an": "vs an dernier", "precedente": "vs période préc."}


def compare_range(days: int, mode: str):
    if mode == "an":
        return oi.date_range(days, 1)
    if mode == "precedente":
        return oi.date_range(days, periods_back=1)
    return None


def sentiment_table(now: dict, prev: dict | None, days: int, mode: str = "an") -> pd.DataFrame:
    p_now = oi.date_range(days)
    p_prev = compare_range(days, mode) or oi.date_range(days, 1)
    rows = []
    for k, label in SENT_ROWS:
        a, b = now.get(k), (prev or {}).get(k)
        var = (a / b - 1) if (a is not None and b not in (None, 0)) else None
        rows.append({"Indicateur": label,
                     f"{p_now[0]:%d/%m/%Y} → {p_now[1]:%d/%m/%Y}": a,
                     f"{p_prev[0]:%d/%m/%Y} → {p_prev[1]:%d/%m/%Y}": b,
                     "Variation": var})
    return pd.DataFrame(rows)


DEFAULT_MIN_BUY = 50_000   # achats plus petits ignorés (achat médian S&P 500 ≈ 90 000 $, oct. 2026)


def filter_tx(tx: pd.DataFrame, exclude_10pct: bool = False,
              min_value: float = DEFAULT_MIN_BUY) -> pd.DataFrame:
    """Filtres appliqués après téléchargement (instantanés)."""
    if tx.empty:
        return tx
    if exclude_10pct:
        tx = tx[~tx["title"].str.contains("10%", na=False)]
    if min_value:   # seuil appliqué aux ACHATS seulement : toutes les ventes restent comptées
        tx = tx[(tx["type"] != "P") | (tx["value"] >= min_value)]
    return tx


def run_openinsider(uni: pd.DataFrame, days: int = 90, exclude_10pct: bool = False,
                    delay: float = 0.5, progress=None, min_value: float = DEFAULT_MIN_BUY):
    """Analyse B complète : (indicateurs par valeur, transactions retenues, journal)."""
    tx, log = fetch_openinsider(uni, days, delay, progress)
    tx = filter_tx(tx, exclude_10pct, min_value)
    return oi.summarize(tx, uni["ticker"].tolist()), tx, log


def combine(uni: pd.DataFrame, dr_df: pd.DataFrame | None, oi_df: pd.DataFrame,
            criterion: str = "buyers", min_delta_own: float = 0,
            cluster_only: bool = False, dr_universe: set | None = None) -> pd.DataFrame:
    """dr_universe : tickers sur lesquels porte la sélection Dataroma (univers superinvestors).
    Les fiches lues en complément hors univers sont affichées mais ne comptent pas
    comme « retenues par Dataroma »."""
    df = uni.copy()
    if dr_df is not None:
        df = df.merge(dr_df.drop_duplicates("ticker", keep="last"), on="ticker", how="left")
        dr_ok = df["dr_balance_usd"].fillna(0) > 0
        if dr_universe is not None:
            dr_ok &= df["ticker"].isin(dr_universe)
    else:
        dr_ok = pd.Series(False, index=df.index)
    df = df.merge(oi_df, on="ticker", how="left")
    oi_ok = oi.criterion_mask(df, criterion)
    if min_delta_own:
        oi_ok &= df["max_delta_own"].fillna(-1e9) >= min_delta_own
    if cluster_only:
        oi_ok &= df["cluster"] >= 2
    df["dr_ok"], df["oi_ok"] = dr_ok.values, oi_ok.values
    df["selected_by"] = [("Les deux" if a and b else "Dataroma" if a else
                          "OpenInsider" if b else "") for a, b in zip(dr_ok, oi_ok)]
    df["url"] = dr.BASE_URL + "stock.php?sym=" + df["ticker"]
    df["oi_url"] = "http://openinsider.com/" + df["ticker"]
    order = {"Les deux": 0, "OpenInsider": 1, "Dataroma": 2, "": 3}
    df["_o"] = df["selected_by"].map(order)
    return df.sort_values(["_o", "buyers", "balance_usd"], ascending=[True, False, False]) \
             .drop(columns="_o").reset_index(drop=True)


def display(df: pd.DataFrame) -> pd.DataFrame:
    cols = [c for c in LABELS if c in df.columns]
    return df[cols].rename(columns=LABELS)


def display_tx(tx: pd.DataFrame) -> pd.DataFrame:
    if tx is None or tx.empty:
        return pd.DataFrame(columns=list(TX_LABELS.values()))
    return tx.sort_values(["ticker", "trade_date"], ascending=[True, False])[list(TX_LABELS)] \
             .rename(columns=TX_LABELS)


def make_note(grid, min_si, days, criterion, excl10, with_dr, scope="si", min_value=DEFAULT_MIN_BUY) -> str:
    return (f"Univers : dataroma.com « {dr.GRIDS[grid]} », ≥ {min_si} superinvestors"
            + (" ; OpenInsider sur tout le S&P 500. " if scope == "sp500" else ". ")
            + ("Dataroma : solde initiés 6 mois > 0. " if with_dr else "")
            + f"OpenInsider : déclarations des {days} derniers jours, levées d'options exclues"
            + (", actionnaires >10 % exclus" if excl10 else "")
            + (f", achats < {min_value:,.0f} $ ignorés".replace(",", " ") if min_value else "")
            + f", critère « {oi.CRITERIA[criterion]} ». Relevé le {date.today():%d/%m/%Y}.")


def export_excel(final: pd.DataFrame, tx: pd.DataFrame, note: str,
                 sent: pd.DataFrame | None = None) -> bytes:
    sel = final[final["selected_by"] != ""]
    sheets = {"Sélection": display(sel), "Toutes les valeurs": display(final),
              "Transactions OpenInsider": display_tx(tx)}
    if sent is not None:
        sheets["Sentiment"] = sent
    return to_excel_bytes(
        sheets,
        MONEY, LINKS, note,
        formulas={LABELS["dr_balance_usd"]: (LABELS["ins_buy_usd"], LABELS["ins_sell_usd"]),
                  LABELS["balance_usd"]: (LABELS["buy_usd"], LABELS["sell_usd"])})


def export_html(sel: pd.DataFrame, title: str, note: str) -> str:
    return to_html_page(display(sel), title, MONEY, LINKS,
                        positive_cols={LABELS["dr_balance_usd"], LABELS["balance_usd"]},
                        subtitle=note)


def main() -> None:
    p = argparse.ArgumentParser(description="Dataroma et OpenInsider, analyses séparées")
    p.add_argument("--grid", default="q", choices=list(dr.GRIDS))
    p.add_argument("--min", type=int, default=1)
    p.add_argument("--jours", type=int, default=90)
    p.add_argument("--critere", default="buyers", choices=list(oi.CRITERIA))
    p.add_argument("--delta-min", type=float, default=0, help="ΔOwn minimum (%%) côté OpenInsider")
    p.add_argument("--groupes", action="store_true", help="seulement les achats groupés")
    p.add_argument("--sans-10pct", action="store_true")
    p.add_argument("--sans-dataroma", action="store_true", help="ne pas lire les fiches Dataroma")
    p.add_argument("--univers-si", action="store_true",
                   help="OpenInsider limité aux valeurs des superinvestors (défaut : tout le S&P 500)")
    p.add_argument("--comparer", default="an", choices=list(COMPARE),
                   help="sentiment comparé à : an (même période l'an dernier), "
                        "precedente (N jours précédents) ou aucune")
    p.add_argument("--sans-comparaison", action="store_true", help="équivaut à --comparer aucune")
    p.add_argument("--montant-min", type=float, default=DEFAULT_MIN_BUY,
                   help="ignorer les achats < X $ (défaut 50 000 ; 0 = aucun)")
    p.add_argument("--out", default="selection_inities")
    a = p.parse_args()

    base = sp500(a.grid)
    uni = universe(a.grid, a.min, base)
    oi_uni = uni if a.univers_si else base
    print(f"Univers Dataroma : {len(uni)} valeurs · OpenInsider : {len(oi_uni)} valeurs")
    dr_df = None
    if not a.sans_dataroma:
        dr_df = run_dataroma(uni, progress=lambda i, n, t: print(f"\rDataroma [{i}/{n}] {t:<8}", end=""))
        print()
    oi_df, tx, log = run_openinsider(oi_uni, a.jours, a.sans_10pct, min_value=a.montant_min,
                                     progress=lambda i, n, t: print(f"\rOpenInsider [{i}/{n}]", end=""))
    print(f"\nOpenInsider : {len(tx)} transactions lues")
    if tx.empty:
        print("⚠️  Aucune transaction OpenInsider. Journal :")
        for e in log:
            print("   ", e)
    final = combine(oi_uni, dr_df, oi_df, a.critere, a.delta_min, a.groupes)
    prev = None
    mode = "aucune" if a.sans_comparaison else a.comparer
    if mode != "aucune":
        prev_raw, _ = fetch_openinsider(oi_uni, a.jours, years_back=int(mode == "an"),
                                        periods_back=int(mode == "precedente"))
        prev = oi.sentiment(filter_tx(prev_raw, a.sans_10pct, a.montant_min))
    sent = sentiment_table(oi.sentiment(tx), prev, a.jours, mode)
    print("\nSentiment des initiés :")
    print(sent.to_string(index=False))
    sel = final[final["selected_by"] != ""]
    note = make_note(a.grid, a.min, a.jours, a.critere, a.sans_10pct, dr_df is not None,
                     "si" if a.univers_si else "sp500", a.montant_min)
    with open(f"{a.out}.xlsx", "wb") as f:
        f.write(export_excel(final, tx, note, sent))
    with open(f"{a.out}.html", "w", encoding="utf-8") as f:
        f.write(export_html(sel, "Superinvestors × initiés", note))
    print(sel["selected_by"].value_counts().to_string())
    print(f"→ {a.out}.xlsx / {a.out}.html")


if __name__ == "__main__":
    main()
