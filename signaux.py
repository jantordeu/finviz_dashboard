"""
Signal des initiés, commun aux screeners AMF (Euronext) et investors (NYSE / Nasdaq).

Entrée : une transaction par ligne, colonnes
    key      identifiant de la valeur (ticker, ou identifiant AMF de la société)
    date     date de la transaction (Timestamp)
    person   déclarant (nom normalisé : sert à compter les dirigeants distincts)
    side     "A" achat / "V" vente
    value    montant (devise de la valeur)
    mandate  vente faite sous mandat de gestion / plan programmé (bool)
    name     (facultatif) nom et fonction affichés pour la dernière opération

Règles (fenêtres glissantes à partir d'aujourd'hui) :
    🟢 favorable     achats ≥ seuil sur 6 mois, pas de vente significative sur 3 mois
    🟠 mitigé        achats ≥ seuil sur 6 mois ET ventes ≥ seuil sur 3 mois, ventes ≤ achats
    🔴 défavorable   ventes ≥ seuil sur 3 mois, supérieures aux achats des 6 mois
    ⚪ non significatif (montants sous le seuil, ou mouvements anciens seulement)
    + « achats groupés »  : ≥ 3 acheteurs distincts en 30 jours (6 derniers mois)
    + « ventes groupées » : ≥ 3 vendeurs distincts en 30 jours (3 derniers mois)
"""
from __future__ import annotations

from datetime import date
import re
from typing import Optional

import pandas as pd

GROUP_N = 3          # dirigeants distincts pour un mouvement « groupé »
GROUP_DAYS = 30
BUY_DAYS = 182       # achats : 6 mois
SELL_DAYS = 91       # ventes récentes : 3 mois
ALERT_DAYS = 7       # alerte : transaction d'un dirigeant dans les 7 derniers jours
ALERT_SELL, ALERT_BUY = "🚨", "✅"
EXIT_DAYS = 182      # sortie d'un dirigeant de premier rang après levée d'options : 6 derniers mois
EXIT_PCT = 50.0      # part de la détention vendue
TOP_ROLE = re.compile(r"\b(?:CEO|Chief Executive|CFO|Chief Financial|President|Pres|Chairman|COB|Founder|PDG|DG|"
                      r"Directeur g[ée]n[ée]ral|Directeur financier|Pr[ée]sident)\b", re.I)
TREND_MIN = 2        # fonds : écart de solde (renforcements − allègements) entre deux trimestres pour une tendance

RANK = {"🟢": 0, "🟠": 1, "🔴": 2, "⚪": 3}
LEGEND = ("🟢 achats sur 6 mois, sans vente significative sur 3 mois · 🟠 achats sur 6 mois mais aussi des ventes "
          "récentes (3 mois) · 🔴 ventes récentes supérieures aux achats · ⚪ mouvements non significatifs ou anciens. "
          "« achats groupés » / « ventes groupées » : au moins 3 dirigeants distincts dans la même fenêtre de 30 jours. "
          "Frise : un carré par mois sur 12 mois, du plus ancien au plus récent — 🟩 achats nets, 🟥 ventes nettes, "
          "⬜ aucune opération. Alerte 7 j : 🚨 vente / ✅ achat d'un dirigeant dans les 7 derniers jours.")


def money(v: float, cur: str = "€") -> str:
    v = abs(v or 0)
    if v >= 1e9:
        return f"{v / 1e9:.1f} Md{cur}".replace(".", ",")
    if v >= 1e6:
        return f"{v / 1e6:.2f} M{cur}".replace(".", ",")
    return f"{v / 1e3:.0f} k{cur}"


def _group(d: pd.DataFrame) -> int:
    """Nb max de déclarants distincts dans une fenêtre de GROUP_DAYS jours."""
    if d.empty:
        return 0
    best = 0
    for s in d["date"].unique():
        w = d[(d["date"] >= s) & (d["date"] <= s + pd.Timedelta(days=GROUP_DAYS))]
        best = max(best, w["person"].nunique())
    return best


def _alert(d: pd.DataFrame, today: pd.Timestamp, cur: str) -> str:
    """« 🚨 vente 152 k€ (04/11) » / « ✅ achat 1,2 M€ (05/11) » : transactions des ALERT_DAYS derniers jours
    (date de la transaction) ; vide sinon."""
    w = d[d["date"] > today - pd.Timedelta(days=ALERT_DAYS)]
    parts = []
    for side, icon, lbl in (("V", ALERT_SELL, "vente"), ("A", ALERT_BUY, "achat")):
        x = w[w["side"] == side]
        if len(x):
            parts.append(f"{icon} {lbl} {money(x['value'].sum(), cur)} ({x['date'].max():%d/%m})")
    return " · ".join(parts)


def _months(today: pd.Timestamp) -> list[pd.Period]:
    p = today.to_period("M")
    return [p - i for i in range(11, -1, -1)]


def summarize(tx: pd.DataFrame, today: Optional[date] = None, min_sig: float = 50_000,
              cur: str = "€") -> pd.DataFrame:
    """Une ligne par valeur : signal, achats 6 mois, ventes 3 mois, solde 12 mois, barres mensuelles."""
    today = pd.Timestamp(today or date.today()).normalize()
    cols = ["key", "signal", "dot", "rank", "buy6", "buyers6", "months6", "sell3", "sellers3", "sell3_mandate",
            "buy12", "sell12", "net12", "monthly", "frise", "group_buy", "group_sell", "last_date", "last_side",
            "achats_txt", "ventes_txt", "derniere_txt", "n_tx", "alerte"]
    if tx is None or tx.empty:
        return pd.DataFrame(columns=cols)
    tx = tx.copy()
    tx["date"] = pd.to_datetime(tx["date"], errors="coerce")
    tx = tx.dropna(subset=["date"])
    months = _months(today)                       # 12 mois calendaires : même fenêtre pour net, frise, dernière opération
    tx = tx[(tx["date"] <= today + pd.Timedelta(days=1)) & (tx["date"] >= months[0].start_time)]
    if "mandate" not in tx:
        tx["mandate"] = False
    out = []
    for key, d in tx.groupby("key", sort=False):
        b = d[d["side"] == "A"]
        s = d[d["side"] == "V"]
        b6 = b[b["date"] > today - pd.Timedelta(days=BUY_DAYS)]
        s3 = s[s["date"] > today - pd.Timedelta(days=SELL_DAYS)]
        buy6, sell3 = float(b6["value"].sum()), float(s3["value"].sum())
        buy12, sell12 = float(b["value"].sum()), float(s["value"].sum())
        gb, gs = _group(b6), _group(s3)
        B, S = buy6 >= min_sig, sell3 >= min_sig
        if S and sell3 > buy6:
            dot = "🔴"
        elif B and S:
            dot = "🟠"
        elif B:
            dot = "🟢"
        else:
            dot = "⚪"
        tags = []
        if gb >= GROUP_N:
            tags.append("achats groupés")
        if gs >= GROUP_N:
            tags.append("ventes groupées")
        mshare = float(s3.loc[s3["mandate"].astype(bool), "value"].sum()) / sell3 if sell3 else 0.0
        per = d.assign(m=d["date"].dt.to_period("M"),
                       v=d["value"].where(d["side"] == "A", -d["value"])).groupby("m")["v"].sum()
        last = d.sort_values("date").iloc[-1]
        months6 = b6["date"].dt.to_period("M").nunique()
        out.append({
            "key": key, "dot": dot, "rank": RANK[dot],
            "signal": dot + (" " + " · ".join(tags) if tags else ""),
            "buy6": buy6, "buyers6": int(b6["person"].nunique()), "months6": int(months6),
            "sell3": sell3, "sellers3": int(s3["person"].nunique()), "sell3_mandate": mshare,
            "buy12": buy12, "sell12": sell12, "net12": buy12 - sell12,
            "monthly": [float(per.get(m, 0.0)) for m in months],
            "frise": "".join("🟩" if per.get(m, 0.0) > 0 else "🟥" if per.get(m, 0.0) < 0 else "⬜" for m in months),
            "group_buy": gb, "group_sell": gs,
            "last_date": last["date"], "last_side": last["side"],
            "achats_txt": (f"{money(buy6, cur)} · {b6['person'].nunique()} dirig. ({months6} mois)"
                           if buy6 else "—"),
            "ventes_txt": (f"{money(sell3, cur)} · {s3['person'].nunique()} dirig."
                           + (" (mandat)" if mshare >= 0.5 else "") if sell3 else "—"),
            "derniere_txt": f"{'achat' if last['side'] == 'A' else 'vente'} {last['date']:%d/%m/%y}"
                            + (f" — {last['name']}" if str(last.get("name") or "").strip() else ""),
            "n_tx": len(d),
            "alerte": _alert(d, today, cur),
        })
    df = pd.DataFrame(out, columns=cols)
    df["_s"] = df["net12"].where(df["rank"] != 2, -df["sell3"])
    return df.sort_values(["rank", "_s"], ascending=[True, False]).drop(columns="_s").reset_index(drop=True)


def _role_label(role: str) -> str:
    r = role or ""
    if re.search(r"\b(CEO|Chief Executive|PDG|DG|Directeur g)", r, re.I):
        return "PDG"
    if re.search(r"\b(CFO|Chief Financial|Directeur financier)", r, re.I):
        return "DAF"
    if re.search(r"\b(Chairman|COB|Pr[ée]sident|President|Pres)\b", r, re.I):
        return "Président"
    return "Fondateur"


def top_exit_warning(rows: pd.DataFrame, today: Optional[date] = None) -> str:
    """« ⚠️ PDG Jane Doe : −85 % de sa détention (levée d'options, 12/03/26) » si un dirigeant de premier rang
    (PDG, DAF, président, fondateur) a vendu, dans les 6 derniers mois et en lien avec une levée d'options, au moins
    la moitié de sa détention. N'entre pas dans le signal 🟢/🟠/🔴/⚪ (information à part).
    rows : colonnes date, name, role, side ("V"), oe (bool), pct_sold (% de la détention vendue)."""
    if rows is None or rows.empty:
        return ""
    today = pd.Timestamp(today or date.today()).normalize()
    r = rows.copy()
    r["date"] = pd.to_datetime(r["date"], errors="coerce")
    r = r[(r["side"] == "V") & r["oe"].astype(bool) & (r["date"] > today - pd.Timedelta(days=EXIT_DAYS))
          & (r["pct_sold"].fillna(0) >= EXIT_PCT) & r["role"].fillna("").str.contains(TOP_ROLE)]
    if r.empty:
        return ""
    out = []
    for name, g in r.sort_values("pct_sold", ascending=False).groupby("name", sort=False):
        x = g.iloc[0]
        out.append(f"⚠️ {_role_label(x['role'])} {name} : −{min(x['pct_sold'], 100):.0f} % de sa détention "
                   f"(levée d'options, {x['date']:%d/%m/%y})")
    return " · ".join(out[:2])


def fund_quarters(add_q: int, red_q: int, add_6m: int, red_6m: int) -> tuple[int, int, int, int]:
    """(renforcements, allègements) du dernier trimestre et du trimestre précédent : les grilles « 6 mois » de
    Dataroma incluent le dernier trimestre, le trimestre précédent s'en déduit par différence."""
    add_q, red_q, add_6m, red_6m = (int(x or 0) for x in (add_q, red_q, add_6m, red_6m))
    return add_q, red_q, max(add_6m - add_q, 0), max(red_6m - red_q, 0)


def fund_signal(add_q: int, red_q: int, add_6m: int, red_6m: int) -> str:
    """Signal des superinvestors (Dataroma), d'après le dernier trimestre :
    🟢 plus de fonds qui renforcent que de fonds qui allègent · 🔴 l'inverse · 🟠 autant des deux · ⚪ aucun mouvement ;
    + tendance par rapport au trimestre précédent (« ↗ en amélioration » / « ↘ en dégradation », écart ≥ 2 fonds)
    + « achats groupés » (≥ 3 fonds acheteurs dans le trimestre) / « ventes groupées »."""
    aq, rq, ap, rp = fund_quarters(add_q, red_q, add_6m, red_6m)
    q, p = aq - rq, ap - rp
    dot = "🟢" if q > 0 else "🔴" if q < 0 else "🟠" if aq else "⚪"
    tags = []
    if q - p >= TREND_MIN:
        tags.append("↗ en amélioration")
    elif p - q >= TREND_MIN:
        tags.append("↘ en dégradation")
    tags += (["achats groupés"] if aq >= GROUP_N else []) + (["ventes groupées"] if rq >= GROUP_N else [])
    return dot + (" " + " · ".join(tags) if tags else "")
