# Finviz Dashboard — contexte du projet

## Objet
Dashboard Streamlit en français (`finviz_dashboard.py` + `requirements.txt`) pour suivre un portefeuille d'actions US
(fichier Excel du courtier, devise USD, hors ETF) et étudier de nouvelles valeurs. Utilisateur francophone, réponses
concises et pratiques. Déploiement prévu : GitHub (dépôt privé) + Streamlit Community Cloud. Plus d'installateur Mac.

## Sources de données
- Finviz (cloudscraper) : fiche « quote », états financiers `/api/statement` (IA, BA, CA), options, actualités,
  sociétés comparables (lien « Peers » → screener v=121), multiples d'industrie (page Groups, cache 24 h), screener.
- OpenInsider (initiés 12 mois), Dataroma (super investors), ApeWisdom (rangs Reddit).
- Zonebourse : uniquement des pages HTML enregistrées et déposées par l'utilisateur (onglets Finances + Valorisation) ;
  jamais de récupération automatique.
- Cache disque 2 h (`./cache_finviz`).

## Onglets
Portefeuille · Valeurs à étudier (pages HTML + valeurs envoyées par le Screener JV, gardées séparément) · Sentiment
(VADER + lexique financier) · Sentiment screener (Reddit : 10 valeurs les plus citées à l'actualité positive) ·
Screener JV (lit un screener Finviz, exclut des industries, JV rapide, envoi en analyse complète) · Screener initiés
(superinvestors Dataroma × achats des dirigeants OpenInsider, sentiment des initiés, exports ; modules analyse.py,
openinsider.py, dataroma_scraper.py, exports.py dans le même dossier) · Méthode.

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
