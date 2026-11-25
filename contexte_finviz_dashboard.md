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

## Sociétés comparables : tableau + pairs européens (13/11/2026)
- `zb_peers` lit le drapeau de chaque ligne (`i.flag__xx`) → `cc` ; `EUROPE_CC`, `CC_NAMES`, `cc_region`. PAGE_PARSE_VERSION 4 (+ `parse_zonebourse_cached(..., version=4)`).
- `comparables` (branche Zonebourse) : société européenne → seuls les pairs européens, s'il y en a ≥ 3 (sinon tous, noté) ; sociétés US inchangées. `fv["_peers_tab"]` = (lignes, multiples de la société, libellé PER).
- `peers_rows(fv, f)` : tableau Société / Pays / PER / VE/EBITDA / VE/CA (ou P/FCF Finviz) / Retenu (✓ ou « autre région ») + médiane des retenus + la société ; section 5 (expander « Sociétés comparables (n retenues sur m) ») et PDF.
- Effets : Bouygues comparables 109,7 → 89,3 € (toujours écartés : Ferrovial 58x, Hochtief 26x, ACS 22x), Engie 45,8 → 34,4 (JV 29,3 → 27,8), Technip Energies 37,4 → 25,9 (JV 45,0 → 42,6), Sanofi 185 → 154 (JV inchangée). ASM, BESI, Siemens Energy, Clariane, Renault, Air France : < 3 pairs européens → tous conservés.

## Screener AMF : sociétés néerlandaises via le registre AFM (13/11/2026)
- Les sociétés de droit néerlandais du SBF 120 (Airbus, Stellantis, STMicro, Euronext, Pluxee, Technip Energies) déclarent à l'AFM, pas à l'AMF → absentes de la BDIF.
- Nouveau module `afm_nl.py` : export XML du registre « Transacties leidinggevenden MAR 19 » (`/export.aspx?KeyWords=<nom>&DateFrom=jj-mm-aaaa&DateTill=…&type=0ee836dc-5520-459d-bcf4-a4a689de6614&format=xml` → meldingid, transactiedatum M/D/YYYY, émetteur, déclarant, nom, personne liée, fonction, LEI) puis page de détail `/en/sector/registers/meldingenregisters/transacties-leidinggevenden-mar19-/details?id=<meldingid>` (tableau « Transactions » : Instrument type, ISIN, Transaction category, Transaction type, Stock option program, Trading place, Price, Quantity, Unit ; nombres au format néerlandais 14.363,00).
- Règles : achat = Verwerving/Aankoop, vente = Vervreemding/Verkoop ; exclus : programme d'options « Ja », instruments performance/restricted/award/option/right/warrant…, non-actions, schenking/verpanding/uitoefening…, prix ou volume nul. Liste en cache 6 h, détails en cache disque définitif (cache_finviz/afm_nl).
- `_amf_load` (univers SBF 120) : ISIN NL absents de la BDIF → `afm_nl.load_companies` ; jeton « AFM:<ISIN> » ; mêmes colonnes que amf.load_transactions. Luxembourg (CSSF) et Belgique (FSMA) restent listés sous le tableau.
- ⚠️ Testé sur les pages réelles déposées (Technip Energies) avec réseau simulé : afm.nl inaccessible depuis la session de développement.

## Faux cyclique corrigé : EBIT en hausse régulière (14/11/2026)
- Palo Alto Networks classée cyclique à tort : BNA comptable 1,60 → 0,40 → 1,54 (amortissements et charges de l'acquisition CyberArk, émission d'actions) → bond > 150 %. EBIT pourtant en hausse chaque année (2,2 → 2,7 → 3,4 → 4,2 → 5,0 Md$).
- `parse_zonebourse` : `f["ebit_steady"]` = EBIT de L−4 à N+1 (≥ 5 valeurs), positif, jamais en baisse de plus de 5 %, secteur hors CYCLICAL_ZB, hors banque → `cyclical = False` + `steady_note` (note + alerte ℹ️ dans fair_value) ; prioritaire aussi dans la fusion Finviz + Zonebourse.
- Effet sur les pages déposées : seul PANW change (JV pages Zonebourse 46 → 77 $). Bouygues, Oracle, S&P Global « steady » mais déjà non cycliques.

## Section 12 : graphique des cours via yfinance (14/11/2026)
- Nouveau dossier `yfinance/` (sans __init__.py, n'écrase pas la bibliothèque yfinance) : `data_yfinance.py` (script de l'utilisateur, inchangé hormis une ligne de docstring ; lancé seul il lit data.txt et écrit data/<TICKER>.csv). Le dashboard le charge par `importlib` (`_yf_module`) et utilise `download()` / `save()` (CSV copiés dans yfinance/data/).
- `yahoo_symbol(f)` : ticker Finviz pour une valeur US (BRK.B → BRK-B), sinon recherche Yahoo par ISIN (cache 24 h). `yf_history(sym, "2y")` (st.cache_data 1 h) ; `money_flow_index(df, 14)` ; `price_chart` : Plotly 3 panneaux (chandeliers + MM 20 j, volumes colorés hausse/baisse, MFI avec zones 80/20), week-ends masqués.
- `render_price_chart` (fragment) après « 11. Actualités et sentiment » : choix 3 mois / 6 mois / 1 an / 2 ans (indicateurs calculés sur 2 ans), légende : clôture vs MM20, lecture du MFI. requirements : yfinance>=0.2.40.
- Testé avec des cours simulés (Yahoo inaccessible depuis la session de développement).
- Symbole Yahoo (14/11) : recherche Yahoo par ISIN (10 cotations), `pick_yahoo_symbol` : 1) racine du symbole = ticker Zonebourse (STMPA → STMPA.PA, ENR → ENR.DE) ; 2) place préférée selon le pays de l'ISIN (`YAHOO_EXCH` : FR → PAR, NL/LU → PAR puis AMS, DE → GER…) ; 3) première action. Champ « Symbole Yahoo » modifiable dans la section 12 (correction gardée pour la session).

## Simplification de l'interface (14/11/2026)
- Barre latérale : expander « Réglages de la juste valeur » supprimé (p = DEFAULT_P : hypothèses et scénarios pessimiste/central/optimiste par défaut) ; champ « Forcer le classement cyclique » supprimé.
- Tableaux Portefeuille et Valeurs à étudier : colonne case à cocher « Cyclique » juste après « Ticker » (à la place du sigle ⟳ dans le ticker) ; `cyc_table` (st.data_editor, seule la case modifiable) → `st.session_state["cyc_ov"]` {ticker: "oui"/"non"}, supprimé quand on revient au classement automatique ; `ov` construit depuis cyc_ov ; légende des corrections sous le tableau. Le screener JV garde ⟳ et son filtre.

## Audit du code et revue de la méthode (14/11/2026, modèle Fable)
### Corrections appliquées (aucune juste valeur modifiée — vérifié sur 31 valeurs)
- Caches : `_read_json` (fichier illisible = cache manquant, plus d'erreur pendant 2 h après une écriture interrompue) et `_write_json` (écriture atomique) pour fetch_ticker, complete_online_cached, fiche du screener ; lectures tolérantes dans afm_nl et edgar_sec.
- `yf_history` : seuls les succès sont mis en mémoire (`_yf_history_cached`) ; après un échec, nouvel essai possible au bout de 2 min (avant : échec gardé 1 h).
- « Vider le cache » vide aussi grilles Dataroma, AMF, Yahoo (actualités, cours), Fear & Greed, et lève les blocages « _fail_ ».
- afm_nl : pages de détail lues par 3 en parallèle ; fichier de liste unique par société (avant : un fichier par jour).
- Affichage DCF : « croissance du CA an 1 » = celle réellement projetée (consensus), note et tableau des scénarios (`res["g1"]`).
### Constats de méthode NON appliqués (à décider) — mesurés sur 31 valeurs Zonebourse
- Calibrage : JV/objectif médian 0,80 ; croissance modérée 0,92 ; croissance CA ≥ 12 % 0,67 ; cycliques 0,59.
- M1 « trou de trésorerie » années 4–7 du DCF : après le consensus, FCFF = NOPAT × (1 − g / RONIC) avec RONIC fixe 11 % → flux quasi nul quand g ≈ 10 % (GE Vernova 10,1 → 0,7 Md$ en année 6, Schneider 7,6 → 3,0 Md€). Variante testée (RONIC observé en fin de consensus, plafonné 50 %, décroissant vers 11 % en année 10) : DCF +5 à +35 % sur les valeurs de croissance, JV/objectif croissance 0,67 → 0,73. Code de la variante : /tmp (à refaire) — remplacer `drv` dans dcfx_project.
- M2 acquisition en cours d'exercice : CA consensus avec la cible mais dette nette d'avant (Constellation/Calpine : 19,4 Md$ = 55 $/action ; Engie : 16,6 Md€ = 6,8 €/action). Règle proposée : sorties hors flux libre et dividendes en N > 10 % de la capitalisation et dette N > dette L → les ajouter à la dette nette du DCF.
- M3 cycliques avec consensus : le DCF reprend 3 ans de consensus de pic (SanDisk : JV 1 628 $ par Zonebourse contre 580 $ par Finviz) ; milieu de cycle calculé avec N et N+1 prévus ; historique court → JV à marquer « peu fiable ».
- M4 chocs de marge des scénarios et du Monte Carlo en points absolus (±2 pts = −29 % d'EBITDA pour Boeing, −3 % pour Broadcom) → fourchettes trop larges pour les faibles marges, trop étroites pour les fortes.
- M5 date de valeur : DCF daté de la fin du dernier exercice publié, non capitalisé jusqu'à aujourd'hui (≈ +5 % en fin d'année).
- M6 « précision ±x % » = accord entre méthodes, pas incertitude (SanDisk ±8 %).
- Lenteur : options Finviz = une requête par échéance (≈ 13 sur 200 jours, 15–20 s par valeur) ; limiter aux mensuelles + 2 hebdomadaires diviserait par ~2 le chargement.

## M1 et M2 appliqués (14/11/2026)
- M1 `dcfx_project` : après le consensus, RONIC = celui observé en fin de consensus (g / taux de réinvestissement, borné [RONIC long terme ; 50 %], 50 % si flux libre ≥ NOPAT) décroissant linéairement vers max(CMPC + 2 pts, 10 %) en année 10. GE Vernova années 4–7 : 9,8 / 9,1 / 7,9 / 8,7 Md$ (avant : 6,7 / 3,5 / 0,7 / 3,0).
- M2 `parse_zonebourse` (hors scission) : sorties hors flux libre et dividendes en N = (dette N − dette L) + flux libre N − dividendes N ; si > 10 % de la capitalisation et dette N > dette L → ajoutées à `dcfx["net_debt"]` (`acq_adj`, `acq_note` ; note + alerte ℹ️). Déclenché sur Constellation (19,4 Md$), Engie (16,6 Md€), Air France-KLM (1,5 Md€).
- Effets (JV) : ASM 748→799, Broadcom 443→506, Coherent 243→279, GE Vernova 585→618, NVIDIA 302→340, Oracle 255→304, Schneider 246→257, Siemens Energy 131→141, Vertiv 202→225, argenx 671→735, Palo Alto 77→82 ; Constellation 265→246, Engie 28→25, Air France 28→25. JV/objectif médian : croissance 0,67→0,73, croissance modérée 0,92→0,88.

## M3 appliqué : cycliques (14/11/2026)
- `dcfx_project` : pour une cyclique, une marge du consensus supérieure à la marge de milieu de cycle n'est reprise qu'à moitié (`k_c` = moyenne consensus / milieu de cycle ; flux libre du consensus réduit dans la même proportion ; même décote `k_last` sur le point de départ de l'après-consensus). Consensus sous le milieu de cycle (creux) : inchangé.
- Cyclique avec moins de 4 exercices publiés : `fv["_low_conf"]`, alerte « ⚠️ juste valeur peu fiable », précision « 🔴 peu fiable (cycle : n exercice(s) publié(s)) », plus de point fort « x % des simulations dépassent le cours ».
- Effets (pages Zonebourse) : SanDisk 1 633 → 1 507 $ (peu fiable ; 580 $ par le chemin Finviz, inchangé), Seagate 288 → 256 $, First Solar 165 → 157 $, BESI 80 → 79 €, Renault inchangé.
- Non fait : M4 (chocs de marge relatifs), M5 (date de valeur — voir discussion : capitaliser au coût des fonds propres moins le rendement du dividende depuis la clôture, mois de clôture à lire sur Zonebourse).

## CVS Health : faux cyclique sur le chemin Zonebourse (14/11/2026)
- Bond du BNA > 150 % = cyclique seulement si secteur CYCLICAL_ZB ou EBITDA N ≥ 1,4 × EBITDA L (sinon `steady_note` : dépréciations / charges exceptionnelles). CVS : BNA 1,39 → 7,00, EBITDA 17,1 → 20,7 Md$ → non cyclique ; JV pages Zonebourse 135 $ (82–200), objectif 110, cours 89. Aucune autre des 31 valeurs ne change.
- Fiche Finviz seule de CVS : JV 205 $ (DCF 245, fourchette 74–419). Causes identifiées, NON corrigées : croissance du CA an 1 de 10,1 % (formule = CA historique + 2,5 pts quand la croissance du BNA est plafonnée ; consensus +2 à +5 %), CMPC 6,8 % (bêta 0,58), choc de marge ±2 pts sur une marge de 4,7 % (scénarios 57 / 231 / 463 $ ; M4), PER cible 19,2x (croissance du BNA depuis une base déprimée) contre 10x pour le marché.

## Trois corrections après le cas CVS (14/11/2026)
- Rebond ≠ croissance : `rebound_growth(f)` — dernier BNA publié < 60 % du BNA N prévu et un exercice antérieur ≥ 1,5× plus haut → croissance du BNA = N → N+2 par an (Zonebourse) ou N → N+1 (Finviz) ; utilisée par `growth_expected` (PER cible, croissance du CA du DCF Finviz) ; note dans fair_value. Déclenché : CVS (82 → 10 %), Constellation (31 → 15 %), Aptiv, Clariane, Palo Alto, BESI, Renault (les trois derniers sans effet : plafond CA + 10 ou cyclique).
- Plancher du CMPC : `DEFAULT_P["wacc_min"] = 0.075` dans `wacc_calc` (mention dans « how » quand il joue ; surtout chemin Finviz, bêta bas : CVS 6,8 → 7,5 %).
- Chocs de marge proportionnels : SCEN = ±10 % de la marge de référence (`margin_ref` : 1re année du consensus, sinon marge actuelle, sinon médiane historique, min 2 %) ; Monte Carlo écart-type 7,5 % de la marge (`MC_MARGIN_SD`). Avant : ±2 pts et 1,5 pt pour tous. Fourchettes : Bouygues 39–79 → 44–74, Boeing 76–140 → 86–130 ; valeurs centrales quasi inchangées.
- CVS par les pages Zonebourse : JV 128 $ (101–181), objectif 110 $, cours 89 $. Constellation 246 → 223 $.

## Profil de volume dans le graphique (15/11/2026)
- `volume_profile(df, bins=40)` : volume de chaque séance réparti uniformément entre plus bas et plus haut, part acheteuse = (clôture − bas) / (haut − bas) (Yahoo quotidien : pas de volumes bid/ask réels) ; renvoie niveaux, volumes acheteur/vendeur, prix le plus traité.
- `price_chart` : grille 3 × 2, profil en barres horizontales empilées à gauche des chandeliers (même axe de prix), ligne pointillée du prix le plus traité ; recalculé pour la période choisie (3 mois … 2 ans), pas lors d'un zoom à la souris.
- Profil de volume, version 2 (15/11) : couleur selon le sens de la séance (clôture ≥ ouverture = volume à la hausse, vert ; sinon à la baisse, rouge) ; zone de valeur = niveaux contigus autour du prix le plus traité concentrant 70 % du volume (couleurs vives dedans, atténuées dehors ; bande grise sur les chandeliers) ; la légende indique les bornes de la zone et la position du cours. `volume_profile` renvoie 7 valeurs.
- Graphique : période « 3 ans » ajoutée (téléchargement 3y) ; testé sur 767 séances réelles de Sanofi fournies par l'utilisateur (prix le plus traité 85,69 € ; zone de valeur 74,89–93,40 € = 70 % du volume).

## Symbole Yahoo des valeurs européennes (16/11/2026)
- Problème signalé sur Sanofi (cause exacte inconnue : la fiche PDF ne contient pas le message). Nouvelle résolution `yahoo_candidates(f)`, sans dépendre de la recherche Yahoo : valeur en USD → ticker ; sinon 1) table `yfinance/tickers_sbf120.csv` (fichier de l'utilisateur, colonnes Ticker,Entreprise, BOM géré ; correspondance par code Zonebourse + nom, sinon nom), 2) recherche Yahoo par ISIN seulement si la table ne donne rien, 3) code Zonebourse + suffixes du pays de l'ISIN (`YF_SUFFIX` : FR .PA ; NL .AS/.PA/.BR ; DE .DE…).
- `render_price_chart` essaie les candidats dans l'ordre jusqu'à obtenir des cours, retient celui qui marche (session), et affiche chaque essai et son erreur en cas d'échec.
- Ne pas confondre avec `sbf120.csv` à la racine (nom;isin, screener AMF).

## Corrections cycliques en sortie de crise (Stellantis)
1. `cycle_eps` : les exercices **publiés** dont le BNA dépasse 3 × la médiane sont exclus du milieu de cycle (les prévisions restent, pour ne pas pénaliser une croissance structurelle type Seagate). Note affichée.
2. Le plafond « PER cible ≤ 1,25 × PER historique » s'applique désormais aussi aux cycliques.
3. Comparables Zonebourse : VE/EBITDA des pairs écarté si le VE/EBITDA de la société < 1x (trésorerie nette non libre) → PER seul.
Effet sur les 32 fiches : seule Stellantis bouge (JV 12,9 → 6,2 €).

## Téléchargement Zonebourse par navigateur réel (zb_fetch.py)
- `zb_fetch.py` (Playwright, Chrome réel, profil persistant `zb_profile/`) : commandes `login` (connexion manuelle, mot de passe jamais lu), `status`, `list`, `fetch [noms] --max-age H`. Pages enregistrées dans `zb_pages/<valeur>/valorisation.html` et `finances.html`, journal `_journal.csv`. Pauses 5–11 s, arrêt sur 401/403/429/captcha, validation du titre/horodatage/taille.
- `zonebourse_urls.csv` (nom;url) : CAC 40 + Aptiv, Vertiv, Nvidia, Meta, Boston Scientific, Constellation, CVS, Technip Energies, Clariane, Nexans, Adyen.
- Dashboard : panneau « 🌐 Zonebourse (abonné) » (`zb_fetch_panel`) dans « Valeurs à étudier » : Connexion, Vérifier, choix des valeurs, Télécharger ; les pages prêtes s'ajoutent aux pages déposées au clic sur « Analyser ». Pas de champ identifiant/mot de passe (choix volontaire).
- Testé en local sur un serveur factice (téléchargement, validation, AppTest) ; non testé sur le vrai site (connexion, anti-robot, indicateur « connecté »).

## Correctifs du 16/11 (soir)
- `zb_isin` : ISIN lu d'abord dans la pastille de l'en-tête, puis les mots-clés, puis le plus fréquent (les pages rendues par Playwright contiennent les ISIN d'autres valeurs). PAGE_PARSE_VERSION = 5.
- Graphique section 12 : clés de saisie et de figure par ticker + ISIN (plus de symbole partagé entre valeurs, plus d'identifiant en double).
- Téléchargement Zonebourse en arrière-plan (`zb_fetch_progress`, fragment 3 s) ; Chrome avec bac à sable, sans `--enable-automation`.

## Portefeuille USA + Europe automatisé (17/11)
- Fichier de positions : toutes les actions (toutes devises). Valeurs USA : Finviz + OpenInsider + Dataroma, combinées aux pages Zonebourse ; valeurs Europe (ou USA sans fiche Finviz) : pages Zonebourse seules (+ AMF via l'ISIN). `portfolio_facts(raws, zb_store, pf_all, zmap)` -> (faits, messages) ; clés = ticker de la position ; pages d'une autre devise que la position refusées.
- Pages Zonebourse du portefeuille : `zb_pages/<slug>/` (lues par `zb_disk_facts`, relues si les fichiers changent), sinon pages déposées (`zb_store`). « Récupérer les données » (automatique à chaque nouveau fichier) lance Finviz puis `zb_start` en arrière-plan ; fin du téléchargement -> relance du dashboard (`zbf_seen`).
- Rapprochement position ↔ adresse : `zb_fetch.match_position` (colonne ticker, nom identique, préfixe, similarité ≥ 0,88). Adresses manquantes : formulaire dans la barre latérale -> `zb_fetch.add_url` écrit `zonebourse_urls_perso.csv` (nom;url;ticker), jamais dans le zip ; correction d'un rapprochement par le même moyen. `fetch --slugs a,b`.
- Barre latérale : 1 Portefeuille, 2 Zonebourse (`zb_section`), 3 Valeurs à étudier (pages + « Tickers USA sans page Zonebourse »), 4 Options (OpenInsider et compléments en ligne cochés par défaut, échéances, NYSE, cache), SEC EDGAR. Liste « Tickers » du portefeuille supprimée.

## 18/11 : initiés AFM dans les fiches, état du téléchargement
- `amf_insiders(isin, name)` : société absente de la base AMF et ISIN « NL » -> registre AFM (`afm_company_tx`, cache 6 h, même calcul de signal) ; source « registre AFM (Pays-Bas) ». Vaut pour le portefeuille et les valeurs à étudier (Adyen, STMicroelectronics…).
- `zb_status_badge` (fragment 3 s) sous « Récupérer les données » : téléchargement Zonebourse en cours (i/n), terminé à hh:mm, ou interrompu.
- Encadré des adresses manquantes : Entrée ne valide plus (`enter_to_submit=False`), téléchargement lancé à l'ajout (`zb_new`). Échec Yahoo : pas de nouvel essai pendant 30 min.

## Sociétés sans bénéfice : juste valeur spéculative (18/11)
- `speculative_value(f, out)` appelé par `fair_value` quand aucune méthode classique ne donne de JV, BNA N+1 ≤ 0, pages Zonebourse avec comparables. Trois méthodes (`SPEC_METHODS`) : « VE/CA » (médiane des comparables × CA N+1), « EBITDA futur » (VE/EBITDA médian × EBITDA de la dernière année prévue, actualisé 11 %), « Marge norm. » (CA prolongé de 2 ans à la croissance du consensus bornée 5–15 %, marge 15 % [12–18], multiple 15, actualisé 11 %) ; dette nette Zonebourse déduite ; réglages dans `SPEC`. JV = médiane (`synthese`), `_spec=True`, alerte « valeur spéculative », précision « 🔴 spéculative ».
- `_spec_tgt` : multiples que suppose l'objectif des analystes (VE/CA N, N+1, VE/EBITDA dernière année), affiché en section 7 et dans la fiche PDF.
- `zb["ca"]` (chiffre d'affaires par exercice) ajouté ; PAGE_PARSE_VERSION = 6.
- Tests : Intuitive Machines 13,9 $ (Zonebourse seul, 177 M d'actions) ou 10,8 $ (combiné Finviz, 228 M d'actions) vs objectif 25,2 ; Redwire 6,2 $ vs 12,3. Les 32 autres fiches inchangées.

## Ventes dirigeants sur 6 mois + prix moyen des fonds (Dataroma)
- `signaux.SELL_DAYS` = 182 (avant 91) : achats, ventes de dirigeants et mouvements de fonds Dataroma couvrent tous 6 mois. Libellés « ventes 6 mois » partout (tableaux, aides, méthode, fiche). Un 🟢 exige désormais l'absence de vente significative sur 6 mois (certains 🟢 passent 🟠/🔴).
- Screener investisseurs : colonne « Fonds · prix moyen ($) » = « Hold Price » de la fiche Dataroma (un prix moyen par valeur, pas par fonds). `_inv_conviction` renvoie (rows, hold) ; `st.session_state["inv_hold"]`. Remplie seulement si « Fonds : pondérer par la conviction » est cochée (une fiche lue par ticker).

## Compatibilité Streamlit ancien (lancement terminal / anaconda)
- Erreur « 'stretch' has type str, but expected int » : l'anaconda du Mac a un Streamlit < 1.50, PyCharm un récent. `_compat_streamlit()` (début de finviz_dashboard.py) traduit `width="stretch"` en `use_container_width=True` si version < 1.50. Recommandé quand même : `pip install -U "streamlit>=1.66"` dans l'anaconda (autres fonctions récentes utilisées).
- À faire au prochain update : expliquer la case « Fonds : pondérer par la conviction » dans la section Méthode.

## « Too many open files » (macOS, 256 fichiers par défaut)
- Cause : `_session()` créait une session cloudscraper par valeur (ex. `_inv_conviction`, analyse de nombreuses pages) sans la fermer. Maintenant une session par thread, réutilisée (`_TLS`). `_raise_file_limit()` relève la limite à 4096 au démarrage. Lectures `open(...)` sans `with` remplacées par `Path.read_*`.

## Peter Lynch Fair Value (recoupement, hors JV)
- `peter_lynch(f)` : BNA 12 m × croissance (%), PEG = 1 ; croissance de l'EBITDA PAR ACTION (comme GuruFocus ; repli BNA) = moyenne (prévue Zonebourse 3 ans, passée jusqu'à 5 ans), plafonnée de façon DYNAMIQUE (`pl_cap`) : 20 % + 0,5 pt par point de croissance réelle au-delà de 20 %, jusqu'à 30 % (atteint à 40 %/an) ; BNA 12 m = dernier exercice publié avancé vers l'exercice en cours au prorata de l'année écoulée (⚠️ = plafond atteint). n.d. si BNA ≤ 0 ou prévu > +100 %/an (rebond). Colonne « Peter Lynch » du tableau portefeuille/valeurs à étudier et ligne caption dans « En bref » de la fiche. N'entre pas dans la JV. Hypothèse « moyenne » pour « pondéré à la hausse par la croissance passée » (à confirmer).

## Peter Lynch — contrôle (audit) et corrections
- BNA 12 mois : prorata par `fiscal_elapsed(f)` = temps écoulé depuis la clôture du dernier exercice publié (mois de « Période Fiscale » + date de publication, stockés dans `zb["pubL"]`, `zb["fy_month"]` ; PAGE_PARSE_VERSION 7), à la date des pages. Avant : jour de l'année civile -> faux en janvier-février (BNA en retard d'un an) et pour les exercices décalés (NVDA, ORCL, ENR…).
- Non calculée, motif affiché (`{"na": …}`) : cyclique ; rebond (croissance prévue du BNA ou de l'EBITDA/action > 100 %/an) ; bénéfices attendus en baisse ; croissance < `PL_G_MIN` = 10 %.
- Plafond dynamique `pl_cap` : 20 % + 0,5 pt par point au-delà de 20 %, max 30 % (à 40 %/an).
- Limite connue : BNA Zonebourse comptable très bas (ex. PANW 0,73) -> valeur sans signification.
- Peter Lynch aussi dans la fiche PDF/Word : paragraphe en fin de section 6 (`peter_lynch_phrase`, partagé avec l'écran).
- Peter Lynch, écran = PDF : section 6 « Synthèse » des deux côtés — ligne « Peter Lynch (recoupement) » (Retenue : information) dans le tableau des méthodes quand elle est calculée + phrase de détail sous le tableau ; retirée de « En bref ». `render_fv_tables(fv, cur, f)`.
- Libellés « Peter Lynch (GuruFocus) » partout (colonne, tableau section 6, phrase, PDF) pour rappeler l'origine de la méthode.
- 21/11 : Peter Lynch MASQUÉE (`PL_SHOW = False`) à la demande de l'utilisateur : elle perturbait la lecture (ex. GOOGL : 435 $ chez nous, 300 $ sur GuruFocus, JV 317 $). Code conservé ; ne pas la réafficher sans avoir reproduit le calcul exact de GuruFocus.

## Repère « DCF GuruFocus (bénéfices) » (21/11, à la demande de l'utilisateur)
- `gf_dcf(f)` : formule exacte du calculateur GuruFocus (vérifiée sur Alphabet : BNA 12,02, +20 %, 11 % -> 375,6 $ pour 375,5 $ affichés) : 10 ans de BNA croissant à g1 puis 10 ans à 4 %, actualisés à d ; rien après 20 ans ; dette ignorée. g1 = croissance passée du BNA (plus longue période publiée dans Zonebourse, 3 à 10 ans — souvent 4-5 ans, GuruFocus : 10 ans) bornée 5–20 % (5 % sans historique) ; d = plafond(taux 10 ans `DEFAULT_P["rf"]`) + 6 % = 11 %. BNA de départ = `eps_12m(f)` (partagé avec Peter Lynch), ramené au BNA N+1 si l'exercice en cours le dépasse (exceptionnel).
- Affichage section 6 (écran, PDF, Word) : ligne « DCF GuruFocus (bénéfices) — repère » (Retenue : information) + phrase de détail. `GF_SHOW`. Hors JV. Écarts possibles avec le site : BNA hors exceptionnels (Zonebourse = BNA publié), durée de l'historique.

## Révisions du consensus (21/11)
- Les pages Zonebourse téléchargées sont ÉCRASÉES à chaque téléchargement (zb_pages/<valeur>/finances.html) : pas d'historique des pages. D'où un fichier de relevés : `data/consensus_hist.csv` (jamais dans le zip, donc conservé d'une mise à jour à l'autre) — colonnes key (ISIN sinon ticker) ; ticker ; date des pages ; an ; bna ; ca ; titres ; cur. Un relevé par valeur et par date (exercices N à N+2), écrit par `cons_record(f)` appelé dans `analyze`.
- `cons_revisions(f)` : compare le même exercice (N+1, à défaut N) au plus ancien relevé des 100 derniers jours (≥ 5 jours). 🟢 BNA ≥ +2 %, 🔴 ≤ −2 %, ⚪ sinon ; n.d. au 1er relevé, si le relevé précédent est trop récent/ancien, ou si le nombre d'actions a changé de plus de 20 % (division du titre).
- Affichage : colonne « Révisions du consensus » (tableau portefeuille / valeurs à étudier, après BNA N+1) ; phrase en fin de section 3 de la fiche (écran, PDF, Word).
- Testé sur les pages réelles à plusieurs dates (ex. NVDA 05/11 -> 19/11 : BNA 2028 +3,9 %, CA +4,7 % 🟢).
- Doublon `_cagr` supprimé (une seule définition).

## Gérants de référence (Baggr)
`REF_FUNDS` (finviz_dashboard.py) : Ackman (Pershing Square), Greenberg (Brave Warrior), Rolfe (Wedgewood), Spier (Aquamarine), Abel/Berkshire.
Sélection manuelle sur Baggr (≤ 60 lignes ; Spier et Abel battent SPY sur 1 an). Repérés par nom/fonds dans les lignes Dataroma
(`parse_dataroma` garde maintenant la colonne « Fonds »). Affichage : fiche (« Gérants de référence » : poids + dernier mouvement)
et colonne « Gérants de référence » du screener investors (nécessite la lecture des fiches Dataroma). N'entre pas dans les signaux.

## Onglet « ⭐ Gérants de référence »
`REF_MANAGERS` (codes Dataroma : psc, CCM, WP, aq, BRK) ; `parse_dataroma_manager` lit holdings.php?m=CODE (ticker, % portefeuille,
mouvement du trimestre, cours déclaré / actuel) ; `_ref_load` (cache 6 h) ; `render_ref_managers`. Initiés : `an.fetch_openinsider` sur les
tickers détenus (12 mois, hors levées d'options) puis `sg.summarize` ; pas de repli Finviz dans cet onglet (« — » si OpenInsider est down).
Vues : par gérant, achats/renforcements, allègements/sorties, titres communs, tableau complet.

### Rotation des gérants de référence
Liste éditable dans l'onglet (expander « Gérants suivis ») : retirer, ajouter depuis le catalogue Dataroma (`_ref_catalog`, managers.php) ou
saisie manuelle (nom · fonds + code Dataroma). Enregistrée dans `data/ref_managers.json` (`ref_managers_get/save`, défaut `REF_DEFAULT`).
`ref_keys()` déduit les mots de repérage (nom de famille + début du fonds) pour la colonne « Gérants de référence » du screener et des fiches.
Limite : seuls les gérants présents sur Dataroma (13F US) sont ajoutables.

## Lecture simple de l'écart JV / objectif (section 7)
`simple_mults(f)` (10/12/15x si cyclique, sinon ~0,8 / 1 / 1,25 × PER historique médian, base bornée 10–32) ; `simple_reading(f, fv, cur)` :
phrase « pour justifier le cours il faut un BNA normal de … », tableau BNA requis (cours / JV / objectif), scénarios BNA × multiples
(consensus N+2, N+1, 12 derniers mois, milieu de cycle ou moyenne, point bas ; non cyclique : « déception −20 % » à la place de l'historique lointain),
valorisation comparée (PER 12 mois et N+1 de la valeur et de ses pairs Finviz / Zonebourse + médiane). Écran, PDF et Word. `SIMPLE_NOTE`.
L'ancienne analyse (pont, repères, prime, causes) suit sous « Analyse détaillée ».

## Sentiment des actualités : vocabulaire boursier élargi
`FIN_LEXICON` complété (fall, decline, dip, sink, slide, plunge, fine, clawback… / rise, climb, rebound…) et `FIN_NEUTRAL` : mots que VADER note
positifs à tort en bourse (« shares », « advantage » de Medicare Advantage, « united », « trust », « care », « securities »…) retirés du lexique.
Cas d'origine : « UnitedHealth, Humana and CVS Shares Fall as Medicare Advantage… Threaten $107 Billion » noté +15, désormais −62.
Limite : seul le titre est noté ; pas de compréhension du sujet (« yields fall » compte comme une baisse).

## Redressement après pertes (cas Intel, 23/11/2026)
Problème : fiche fusionnée Finviz + Zonebourse classée « cyclique » par Finviz (bond du BNA) alors que Zonebourse reconnaissait un
redressement -> PER de pic 8x sur un BNA de creux, JV 18 $ affichée ±18 %. Corrections : `merge_facts` — `zb["turnaround"]` l'emporte
(sauf `low_fpe`), clés turnaround/recovery/recovery_note/zb_sector reportées ; `recent_losses(f)` (dernier exercice publié ou exercice
en cours en perte, BNA N+1 > 0) -> alerte « société en redressement… JV peu fiable », `out["_trough"]`, précision 🔴
« peu fiable (redressement : pertes récentes) » ; cyclique en sortie de pertes : multiple normalisé (pe_cyc_mid 11x) au lieu de 8x.
Intel pages Zonebourse seules : JV 25,6 (DCF 29, PER 16x × 1,37 = 22, PER hist. 17,5 ; comparables 44 écartés).
Suite (même jour) : fiche fusionnée INTC, DCF Finviz 3 $ (pessimiste 0,6 / central 9,1 / optimiste −5,9, CMPC 12,9 %) pesait 50 % -> JV 16.
Règle ajoutée dans `fair_value` : en redressement (`recent_losses`), si les 3 scénarios du DCF ne sont pas dans l'ordre
(optimiste < central ou central < pessimiste), DCF non retenu (`_no_dcf_trough`), JV = multiples. Attendu INTC fusionné : ~28 $ (🔴 peu fiable).

## Screener investors élargi au Nasdaq 100 (24/11/2026)
`nasdaq100.csv` (ticker;nom;secteur — 101 lignes, composition Wikipédia au 22/10/2026, à réviser après la reconstitution de décembre) ;
`ndx_list()`, `INV_UNIVERSES` (défaut « S&P 500 + Nasdaq 100 »), `_inv_load(min_si, univers)` : grilles Dataroma S&P 500 + pour les valeurs
du Nasdaq 100 absentes de la grille, `_ndx_extra` lit stock.php?sym=… (détenteurs, ▲ Buy/Add, ▼ Reduce/Sell ; trimestre précédent « n.d. »).
Colonne « Indice », case « Nasdaq 100 seulement », titres d'export selon l'univers. Testé avec des pages simulées uniquement.

## Grille « croissance de qualité » (fiche, sous « 2. En bref » ; aussi dans l'export PDF/Word)
- `growth_grid(f, fv, cur)` : 4 blocs (Qualité : ROIC − CMPC, tendance de la marge d'EBITDA, FCF / résultat net ;
  Solidité : dette nette / EBITDA, couverture des intérêts ; Croissance : CA, BNA hors rachats d'actions, FCF contre BNA ;
  Prix : PEG, rendement du FCF). Chaque critère ✅ / ➖ / ⚠️ (seuils dans `GRID_NOTE`), verdict : « Candidate » si
  JV > cours, ≥ 70 % de critères favorables, aucun ⚠️ sur la solidité ni sur le ROIC.
- ROIC : champ Finviz `ROIC` (`f["roic_fv"]`), sinon estimé (Zonebourse : EBIT × 0,75 ÷ (fonds propres + dette nette) ;
  Finviz : EBIT après impôt ÷ (actif net + dette nette)). CMPC = celui du DCF, sinon 9 % par défaut (signalé).
- Effet de rebond (cyclique, redressement, pertes récentes) : BNA et PEG neutralisés (➖) avec une note.
- Non applicable aux banques / assurances. Critères non mesurables = ligne absente.
- Screener (idée, non implémentée) : pré-tri par les colonnes Finviz (ROIC, Debt/Eq, croissance BNA/CA, PEG) puis grille
  complète sur les 20–30 meilleures seulement.

## Screener JV : filtres de croissance Finviz + grille
- `SCREENER_URL` par défaut : BNA et CA trimestriels > 0, ratio de liquidité > 1, ROIC > 10 %, dette/fonds propres < 1 (court + long terme), CA 5 ans > 10 %, croissance BNA
  attendue > 10 %, PEG < 2, BNA N+1 > 10 %, capitalisation moyenne et plus ; pas de filtre de place (NYSE + Nasdaq).
- `screen_values` ajoute « Grille croissance » (ex. « 6/8 · ✅ Candidate ») et « Points faibles (grille) » (fiche Finviz seule,
  CMPC 9 % par défaut) ; case « Seulement les candidates » dans l'onglet.

## Screener investors : croisement fonds × dirigeants
- `cross_label(fonds_dot, dirigeants_dot, net12)` -> colonne « Croisement fonds × dirigeants » : 🟢🟢 double confirmation ;
  🟢🟠 confirmation partielle (🟠 dirigeants à net acheteur) ; 🟢⚪ / 🟢🔴 fonds seuls (les ventes de dirigeants sont la norme :
  ~65 % des valeurs, autant avec fonds 🟢 qu'avec fonds ⚪) ; 🔴🟢 divergence ; « dirigeants seuls (achats) ».
- Filtres (rangée de pastilles « Croisement fonds × dirigeants », sous les filtres dirigeants ; Secteur et Nasdaq 100 sur la ligne suivante) : « 🟢 fonds sans vente de dirigeants » (fonds 🟢, dirigeants ≠ 🔴), « 🟢 fonds + achats de dirigeants » (🟢🟢 + 🟢🟠).

## Correctifs DCF (cas Marvell, 24/11/2026) — fiche Finviz + Zonebourse
- `asset_light` élargi : incorporels ≥ 40 % de l'actif, capex < 8 % du CA (avant : 3 %) et < la moitié des amortissements.
  Avant, pour une société bâtie par acquisitions avec 5 % de capex (Marvell : amortissements 13 % du CA, surtout des
  incorporels rachetés), le capex projeté montait jusqu'aux amortissements (14 % du CA en année 10) : flux en baisse
  alors que le CA triplait, valeur terminale 28 % de la VE, DCF 37 $. Après : capex terminal = 1,25 × capex → DCF 88 $.
- `merge_facts` : quand les 12 derniers mois Finviz débordent sur l'exercice N (q > 0), la dernière année du consensus
  n'était lue qu'en partie (2 années sur 3). Le consensus est prolongé d'un point (même croissance, plafonnée à 25 %,
  5 % si cyclique) pour reprendre les 3 années.
- Marvell : JV 76 $ → 101 $ (55–119). Reste prudent : CMPC 13,6 % (bêta 2,40), rémunération en actions déduite des flux,
  PER historique sur BNA N+1.

## Screener investors : reprise après interruption
- `_dr_store()` (cache_resource) garde chaque fiche Dataroma lue 6 h : si le chargement est interrompu (clic dans un autre
  onglet), la relecture reprend où elle s'était arrêtée. `st.empty()` avant la relecture : plus de double barre.

## Grille banques / assurances (`growth_grid_fin`)
- `growth_grid` renvoie la grille financière quand `f["bank"]` ou `f["insurer"]` : ROE − coût des fonds propres (`_ke_bank`,
  stocké par `fair_value_bank`), régularité du ROE (5 exercices, Zonebourse), BNA hors rachats, revenus, actif net par action
  (3 ans), rendement · distribution, P/B payé contre P/B justifié (`_pb_just`), PER N+1 contre son historique.
- Non disponibles : CET1 / Solvabilité II, ratio coûts / revenus, ratio combiné, qualité des actifs (rappelés dans le verdict).
- Note propre à chaque grille dans `gg["note"]` (GRID_NOTE / GRID_NOTE_FIN). Le Screener JV utilise aussi cette grille (🏦).

## Vérification d'ensemble du 25/11/2026
- Analyse statique (pyflakes, ruff, pylint -E) sur les 10 fichiers : aucune erreur bloquante (noms indéfinis, appels
  incorrects) ; restes sans effet : quelques variables et imports inutilisés.
- Corrigé : doublon de `_cagr` (une seule définition gardée) ; grille croissance : pertes sur l'un des 3 derniers exercices
  publiés = effet de rebond (BNA et PEG neutralisés, ex. Marvell PEG 0,29 n'est plus ✅).
- Tests d'exécution hors ligne : 13 cas (Intel, Marvell, ASML, STMicro ; Finviz seul, Zonebourse seul, fusion ; cyclique
  inversé ; banque / assureur forcés) × calcul, fiche à l'écran, PDF, Word ; application entière et 9 onglets ; Screener JV,
  Screener investors (filtres de croisement), Gérants de référence avec pages simulées. Scripts : /tmp/zbt (non livrés).
