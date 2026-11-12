# Finviz Dashboard — contexte du projet

## Objet
Dashboard Streamlit en français (`finviz_dashboard.py` + `requirements.txt`) pour suivre un portefeuille d'actions US
(fichier Excel du courtier, devise USD, hors ETF) et étudier de nouvelles valeurs. Utilisateur francophone, réponses
concises et pratiques. Déploiement prévu : GitHub (dépôt privé) + Streamlit Community Cloud. Plus d'installateur Mac.

## Sources de données
- Finviz (cloudscraper) : fiche « quote », états financiers `/api/statement` (IA, BA, CA), options, actualités,
  sociétés comparables (lien « Peers » → screener v=121), multiples d'industrie (page Groups, cache 24 h), screener.
- OpenInsider (initiés 12 mois), Dataroma (super investors), ApeWisdom (rangs Reddit).
- Yahoo Finance (API de recherche publique, sans clé, cache 1 h) : actualités de la section 11 pour les valeurs sans
  fiche Finviz (Europe) — ISIN lu sur la page Zonebourse → symbole Yahoo (RI.PA…) → recherche par nom court, titres
  gardés seulement si Yahoo les rattache au symbole ; titres en anglais ; à défaut de titre sur 7 j, liste 30 j.
- Zonebourse : uniquement des pages HTML enregistrées et déposées par l'utilisateur (onglets Finances + Valorisation) ;
  jamais de récupération automatique.
- AMF — base BDIF (API publique sans clé `bdif.amf-france.org/back/api/v1/`) : déclarations des dirigeants
  (article 19 MAR) ; liste JSON + un PDF par déclaration (formulaire standard, lu avec pypdf, gardé sur disque dans
  `cache_finviz/amf/`). Société identifiée par son « jeton » AMF (l'ISIN peut changer : Capital B, regroupement
  d'actions → FR0014019Y19). SBF 120 : fichier `sbf120.csv` (Euronext, 08/11/2026 ; l'API Euronext est chiffrée).
- Cache disque 2 h (`./cache_finviz`).

## Onglets
Portefeuille · Valeurs à étudier (pages HTML + valeurs envoyées par le Screener JV, gardées séparément) · Sentiment
(VADER + lexique financier) · Sentiment screener (Reddit : 10 valeurs les plus citées à l'actualité positive) ·
Screener JV (lit un screener Finviz, exclut des industries, JV rapide, envoi en analyse complète) · Screener investors (🇺🇸 S&P 500 détenu par les
superinvestors, grille Dataroma « Ownership » ; colonnes Dirigeants = OpenInsider 12 mois et Fonds = renforcements /
allègements Dataroma, rond du dernier trimestre + tendance ↗/↘ vs trimestre précédent) · Screener AMF (🇫🇷 SBF 120 ou toutes sociétés de Paris) · Méthode. Les deux screeners partagent
le module signaux.py : 🟢 achats ≥ seuil sur 6 mois sans vente significative sur 3 mois, 🟠 achats 6 mois + ventes
récentes, 🔴 ventes 3 mois > achats 6 mois, ⚪ non significatif ; « achats / ventes groupés » ≥ 3 dirigeants en 30 j ;
frise 12 mois 🟩/🟥/⬜. Ventes sous mandat de gestion comptées mais signalées. Choix de l'utilisateur : deux onglets
séparés (pas de tableau commun NYSE / Euronext). Concordance : les colonnes Initiés et Super investors du portefeuille /
valeurs à étudier / fiche détaillée utilisent le même calcul et les mêmes seuils que le screener investors
(concord_insiders, concord_funds). Valeurs européennes (Zonebourse, ISIN hors US) : signal AMF des dirigeants
(amf_insiders → amf.company_declarations : jeton via ISIN, recherche ISIN + raison sociale), seuil du screener AMF.
Ventes avec levée d'options : hors signal (choix de l'utilisateur) ; avertissement « ⚠️ sortie PDG / DAF » si PDG/DAF/
président/fondateur vend ≥ 50 % de sa détention en 6 mois après levée d'options (US seulement). Alerte 7 j 🚨 vente / ✅ achat.
Fusion OpenInsider + Finviz : ligne Finviz ignorée si même dirigeant, même sens, ±3 jours (_oi_covers). Le tableau « Total insider Buys/Sells » de Dataroma = dirigeants, pas les fonds.

## Juste valeur (JV)
- Avec états financiers : **DCF d'entreprise** (FCFF sur 10 ans, CMPC avec bêta ajusté Blume, valeur terminale =
  moyenne croissance perpétuelle / multiple de sortie cohérent avec la trésorerie, 3 scénarios 25/50/25, Monte Carlo
  2 000 tirages P10–P90) ; Finviz : flux calés sur la trésorerie d'exploitation réelle ; Zonebourse : consensus puis
  NOPAT × (1 − g/RONIC).
- BNA homogène : Finviz (ajusté) pour PER cible et comparables, Zonebourse pour le PER historique ; scission/cession
  détectée (activités abandonnées + CA prévu < 12 derniers mois − 3 %) → DCF pro forma sur le consensus ; sociétés
  « asset light » (incorporels ≥ 40 % de l'actif, capex < 3 % du CA) : capex d'entretien = capex réel × 1,25.
- **JV = 50 % DCF + 50 % médiane des multiples** (PER cible, comparables, PER historique) — poids réglable.
- Sans états financiers (fiche seule) : médiane DCF simplifié / PER / comparables.
- Banques, courtiers, assureurs : actif net × P/B justifié + PER bancaire (8,5x Europe, 12x US).
- Cycliques (⟳) : bond du BNA > 150 % hors secteurs défensifs, PER prévisionnel < 7, rechute en perte ; PER 8x,
  croissance ≤ 5 %, marge ramenée à la médiane du cycle. Correction manuelle « TICKER:oui/non ».
- Colonnes : Précision JV (🟢 ≤ 15 %, 🟠 ≤ 35 %, 🔴 au-delà — écart DCF / multiples), Lecture de l'écart (causes
  chiffrées), Croissance implicite / attendue (DCF inversé), prix d'entrée et JV / prix d'entrée.

## Screener Finviz par défaut
https://finviz.com/screener/stocks?v=111&f=an_recom_buybetter%2Ccap_midover%2Cfa_debteq_u1%2Cfa_estltgrowth_o10%2Cfa_fpe_u25%2Cfa_opermargin_o10%2Cfa_pfcf_u25%2Cfa_roi_o10%2Csh_avgvol_o200%2Cta_sma200_pa&ft=4&preset=s154743054

## Décisions de l'utilisateur
Pas de FMP, Finnhub, Twelve Data, StockAInsights, Quiver ni opérations des élus du Congrès (signal lent, payant), FinBERT, NLP Reddit, lecture de PDF ; pas de colonne « JV Finviz (portefeuille) » ;
pas de PER sur BNA N+2. Les clés d'API éventuelles ne sont jamais écrites dans le script.

## Pistes ouvertes
Historique long des valeurs US (cycliques, PER historique) : API officielle SEC EDGAR (gratuite), si besoin ;
Exclure le pétrole et le gaz par défaut dans le Screener JV ; comparables et modèle de dividendes pour les banques ;
vérifier que Finviz accepte les requêtes depuis Streamlit Cloud.


## Valeurs cycliques (09/11/2026)
JV centrale = milieu de cycle : DCF sur la marge de trésorerie moyenne du cycle (jusqu'à 8 exercices, moyenne et non
médiane) et marge d'EBITDA moyenne ; PER cible et comparables sur le BNA moyen (publiés + N + N+1, cycle_eps).
Colonne « JV si bénéfices actuels durent » = même synthèse avec les multiples sur le BNA N+1 (JV si pic). Ex. SanDisk :
JV 510 $ (milieu de cycle) / 1 400 $ (si le pic dure) pour un cours de 886 $.

## Conviction des super investors (09/11/2026)
signaux.conviction_signal : chaque mouvement du 13F = poids de la ligne dans le portefeuille du gérant (points ;
allègement x % d'une ligne restée à w % = −w·x/(100−x)), plafond 5 pts par fonds ; 🟢 ≥ +0,5, 🔴 ≤ −0,5, 🟠 compensé,
⚪ symbolique. Utilisé dans l'analyse (fiche Dataroma) et le screener investors (case « pondérer par la conviction »,
fiches Dataroma lues, cache 6 h). Ex. Micron : 🔴 −5,0 pts (Tepper allège 41 % d'une ligne à 15 %).

## Redressements et dette (10/11/2026, cas Clariane)
- Zonebourse : bond du BNA après pertes / BNA quasi nul dans un secteur non cyclique (hors CYCLICAL_ZB) = redressement
  (f['recovery']), pas cyclique : PER sur BNA N+1.
- DCF d'entreprise négatif : DCF écarté, plus de DCF simplifié de secours (il ignorait la dette).
- PER cible réduit de 10 % par point de dette nette / EBITDA au-delà de 3× (au plus −40 %), valeurs non cycliques.
- Croissance en % sur base quasi nulle affichée « non significative » ; raisons de l'écart identiques écran / PDF (gap_causes).
- Clariane : JV 10,2 € -> 4,6 € (objectif analystes 4,7 €). Boeing 108 -> 102 $.

## Objectif « façon analystes » (10/11/2026)
Section 7 : VE/EBIT moyen des 5 derniers exercices (onglet Valorisation Zonebourse) × EBIT N+1 − dette nette N+1, par
action (analyst_style_target) = cours visé à 12 mois ; reproduit l'objectif d'ASMI (1 269 € / 1 164 €) et de Vertiv
(324 $ / 337 $). Garde-fous : EBIT N+1 ≥ 2,5× EBIT publié et multiple passé > 2× l'actuel (Nvidia) ; écart > 35 % avec
l'objectif réel (Clariane) -> explication seulement. Même multiple × EBIT en cours ≈ notre JV (ASMI 864 €, Vertiv 223 $).

## Comparables (10/11/2026)
comparables() : pairs Finviz (vue 152 : PER prév., P/FCF, VE/EBITDA) ou pairs Zonebourse (zb_peers, onglet Valorisation :
PER (N), VE/EBITDA (N)) ; médiane bornée, fourchette Q1–Q3, position prime / norme / décote. VE/EBITDA écarté si dette
nette > 4× EBITDA ou VE incohérent. Cyclique : BNA/EBITDA de milieu de cycle, pas de plafond. Effet : ASMI 846 -> 748 €,
NVIDIA 346 -> 302 $, S&P Global (ZB seul) 403 -> 318 $ (PER hist. écarté par la médiane de 3), Seagate (ZB) 482 -> 277 $.

## PER des cycliques sur BNA de milieu de cycle (10/11/2026)
- Avant : 8x appliqué au BNA de milieu de cycle → double pénalisation (bénéfice déjà normalisé + multiple de bas de cycle).
- Maintenant : `pe_cyc_mid` = 11x (fourchette usuelle 10–12x, curseur 6–16) quand `cycle_eps` fournit un BNA milieu de cycle ; 8x (`pe_cyc`) seulement en repli sur le BNA N+1 (pic possible), p. ex. Renault (BNA milieu de cycle négatif).
- Effet : SanDisk JV 510 → 580 $, pic 1 400 → 1 618 $ ; Seagate (page Finviz) 116 → 129 $. Sur les pages Zonebourse, la ligne PER monte (Seagate 133 → 183, BESI 27 → 37, First Solar 90 → 123) mais la JV centrale ne bouge pas : la médiane des 3 multiples reste les Comparables ; « JV si pic » Seagate 410 → 447.

## Multiple de sortie du DCF des cycliques : 7x → 9x (10/11/2026)
- Le plafond s'applique à un EBITDA déjà normalisé sur le cycle ; 7x faisait double pénalisation. Règle : min(VE/EBITDA du marché borné 4–12x, 9x), toujours ≤ 1,5 × le multiple implicite de la valeur perpétuelle.
- Effet limité : Seagate DCF 73 → 78 $ (JV 129 → 132 $, page Finviz) ; Seagate ZB DCF 296 → 304 (JV 277 → 281) ; BESI DCF 56 → 59 (JV 77 → 79). SanDisk, First Solar, Renault inchangés (VE/EBITDA du marché déjà < 7x).

## Pont de l'objectif des analystes (11/11/2026)
- `target_bridge(f, fv)` / `target_bridge_txt(b, f, cur, jv_pic)` : objectif × titres − trésorerie (ou + dette) nette N+1 = VE implicite → VE/EBIT N+1 et PER N+1 ; mêmes multiples implicites pour la JV. Affiché en section 7 (écran : légende ; PDF : paragraphe), sous l'objectif « façon analystes ».
- Conclusion automatique : écart < 5 % → cohérents ; cyclique avec objectif > JV → écart dû à la base de bénéfices (milieu de cycle vs BNA N+1) ; sinon prime/décote de multiple (la trésorerie/dette est comptée des deux côtés). VE/EBIT masqué si < 2x (ex. Renault, trésorerie de la banque captive). Banques : rien.
- Exemple Siemens Energy : objectif 193 € ↔ 15,8x EBIT N+1 / 23,4x BNA N+1 ; JV 131 € ↔ 10,3x / 16,0x → prime de 47 %.

## « Pourquoi cette prime ? » (11/11/2026)
- `premium_reasons(f, fv)` → affiché (section 7, expander + PDF) si PER implicite de l'objectif ≥ +10 % vs JV, ou PER de la valeur ≥ +25 % vs médiane des pairs (ZB, sinon Finviz). Pas pour banques/assurances.
- Critères (✅ justifie / ➖ neutre / ⚠️ fragilise) : croissance du BNA dernier publié → N+1 (≥ 20 % ✅, < 10 % ⚠️ ; base < 15 % du BNA N+1 → « n.s. », redressement), PEG (≤ 1,2 ✅, > 2 ⚠️), croissance EBIT (note levier/exceptionnels si ≫ BNA), ROE N+1 (≥ 20 % ✅, < 10 % ⚠️), dette nette/EBITDA (≤ 1,5x ✅, > 3x ⚠️), conversion FCF/résultat net (≥ 90 % ✅, < 60 % ⚠️), pertes sur 5 ans ⚠️, cyclique ⚠️.
- Verdicts : cyclique → prime liée au cycle ; croissance (max BNA/EBIT) ≥ 15 % et plus de ✅ que de ⚠️ → cohérente ; < 10 % → peu expliquée (redressement, thème, optimisme) ; sinon en partie justifiée.
- Note fixe : carnet de commandes, thèmes, révisions de consensus non mesurés.
- 11/11 (suite) : conversion FCF = médiane des ratios annuels (une année atypique, ex. Engie 2025 FCF −8,7 Md€, ne fausse plus) ; nouveau critère « Investissements / EBITDA » N–N+1 (≤ 30 % ✅, > 60 % ⚠️). `premium_reasons(..., force=True)` pour comparer des valeurs sans prime.
- Comparaison Siemens Energy / Engie : SE 23,4x BNA N+1 implicite (croissance BNA +36 %/an, trésorerie nette, capex 21 % de l'EBITDA) ; Engie 14,9x (BNA +16 % sur base 2025 basse mais EBIT +7 %, dette 3,5x EBITDA, capex 72 % de l'EBITDA), JV 29 € ≈ objectif 30 €.

## Lot « outils complémentaires » (11/11/2026) — retenus après revue de 6 dépôts GitHub
1. **EPV** `epv_value(f, fv)` : EBIT normalisé × (1 − impôt) ÷ CMPC − dette nette (ZB : EBIT N ; cyclique : moyenne du cycle ; Finviz : EBIT 12 m ou marge moyenne − amortissements). Part du cours / de l'objectif payée pour la croissance ; alerte si trésorerie nette > 50 % de la capitalisation (Renault, Technip Energies). Ligne aussi dans « Pourquoi cette prime ? » (≤ 40 % ✅, > 65 % ⚠️).
2. **Rendement implicite** `dcfx_implied_return(f, res)` → `fv["_irr"]` = (taux, borne) : CMPC qui égalise le scénario central du DCF au cours (bissection 40 itérations, bornes tg+1 pt / 40 %). 🟢 si ≥ CMPC + 1,5 pt, 🔴 si ≤ CMPC − 1,5 pt.
3. **Dilution** `net_dilution(f)` / `dilution_txt` : TCAM du nombre d'actions (≤ 6 exercices publiés ; ZB « titres », Finviz `cap_hist` « Shares Outstanding », ou 10-K si plus long), corrigé des divisions (`ES.split_adjusted_shares`), rachats et SBC moyens 3 ans ; note « émissions d'actions » si ≥ 4 %/an.
   Affichage : tableau « Repère / Valeur / Lecture » (`value_markers`) en section 7 sous le pont de l'objectif (écran + PDF).
4. **Blocages** : `_get` → `FinvizBlocked` (403/503 avec challenge Cloudflare « Just a moment » ou cf-mitigated, 429 après une attente Retry-After ≤ 30 s) ; boucles de chargement et « Envoyer en analyse » interrompues au premier blocage.
5. **EDGAR** `edgar_sec.py` (requests seul, pas edgartools : Python 3.10 OK) : `annual_facts` (companyfacts XBRL, 10-K, durée 340–390 j, dépôt le plus récent par clôture, millions), `form4_rows` (submissions → XML Form 4, P/S agrégés par déclarant/date, « avec levée d'options » si code M, cache disque par numéro de dépôt), ≤ 8 req/s, UA « nom e-mail » saisi dans la barre latérale (expander « 🏛️ SEC EDGAR », jamais stocké, ou env SEC_USER_AGENT). Intégration : `raw["edgar"]` (ONLINE_KEYS, `_ensure_edgar` aussi sur cache), `f["edgar"]`, `long_history(f)` → `f["dcfx_long"]`, `f["eps_hist_long_all"]` (10 exercices : Finviz récents + 10-K anciens, BNA = RN ÷ actions corrigées ramenées à la base Finviz) appliqués dans `fair_value` si cyclique ; Form 4 en secours dans `add_openinsider` (`_edgar_insiders`).
   ⚠️ Non testé en réel : la SEC est inaccessible depuis la session de développement (proxy) ; testé sur données simulées au format officiel.

## Carte du portefeuille (11/11/2026) — supprimée à la demande de l'utilisateur (encombrante, peu utile)
