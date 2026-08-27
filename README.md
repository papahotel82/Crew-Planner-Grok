# Crew Planner – PWA Planning Équipage

Application Progressive Web App pour analyser les plannings de travail d’une équipe équipage à partir de fichiers Excel (.xlsx) au format « Crew Duties ».

## Fonctionnalités

- **Import** de un ou plusieurs fichiers XLSX
- **Consolidation avec résolution de conflits** : en cas de chevauchement personne+date, une modale demande le choix (garder l’existant / le nouveau / choix ligne par ligne)
- **Sauvegarde locale** (localStorage)
- **Filtres** : personne, plage de dates, code
- **Configuration des jours non travaillés** (modale cases à cocher)
- **Indicateurs** avec menu déroulant de mode de ratio :
  - % sur la période
  - % vs jours de l’année
  - Taux annualisé (période)
- **Graphiques** :
  - Camembert : répartition des codes **ou** non travaillés équipe **ou** comparaison 2 personnes
  - Histogramme : non travaillés / total par personne
- **Calendrier heatmap** (section repliable) : vue mensuelle colorée par type de jour
- **Tableau détail** (section repliable, en fin de page)
- **Export** CSV
- **3 thèmes** : sombre, clair, haut contraste
- **Responsive / tactile / PWA** installable

## Règles métier

- **Cellule vide** = jour travaillé (sans vol ni bureau O / training T / visite médicale M…)
- Codes non travaillés par défaut : OFF, H, SICK, MED, CVR, Rest Period, Days Off, Vacation

## Utilisation

Servir le dossier via HTTP (ex. `python3 -m http.server 8080`) puis ouvrir l’URL. Importer le fichier `sample-crew-duties.xlsx` pour tester.
