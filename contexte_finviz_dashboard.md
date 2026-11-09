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
