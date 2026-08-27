# Crew Planner – PWA Planning Équipage

Application Progressive Web App (PWA) pour analyser les plannings de travail d’une équipe équipage à partir de fichiers Excel (.xlsx) au format « Crew Duties ».

## Fonctionnalités

- **Import** de un ou plusieurs fichiers XLSX (consolidation automatique par personne + date)
- **Sauvegarde locale** (localStorage) – les données persistent entre les sessions
- **Filtres** : par personne, plage de dates, code de service
- **Configuration des jours non travaillés** via une modale (cases à cocher)
- **Indicateurs** :
  - Nombre de jours person-jour
  - Jours travaillés / non travaillés
  - Ratio non travaillés / période chargée
  - Ratio non travaillés / jours de l’année (indication)
- **Graphiques** : camembert (répartition des codes) + histogramme (non travaillés par personne)
- **Export** CSV
- **3 thèmes** : sombre, clair, haut contraste (malvoyants)
- **Responsive** & tactile (téléphone, tablette, desktop)
- **Installable** en tant qu’application (manifest + service worker)

## Structure des fichiers Excel attendue

- Lignes d’en-tête contenant « Cockpit CPT… » ou « Cockpit FO… » avec les dates en colonnes (format JJ-MM-AAAA)
- Lignes de personnes : nom en colonne A, codes de service dans les colonnes de dates
- Lignes de notes éventuelles (« flight ») sous les personnes
- Légende en bas (ignorée par le parseur)

Cellules **vides** = jours travaillés (vol).

## Codes non travaillés par défaut

`OFF`, `H`, `SICK`, `MED`, `CVR`, `Rest Period`, `Days Off`, `Vacation`

Modifiables dans l’application.

## Utilisation

1. Ouvrir `index.html` dans un navigateur moderne (ou servir le dossier via un serveur HTTP pour le service worker).
2. Cliquer sur **Importer** et sélectionner un ou plusieurs fichiers `.xlsx`.
3. Ajuster les filtres et la configuration des jours non travaillés.
4. Consulter les stats, graphiques et le tableau détaillé.
5. Exporter en CSV si besoin.

Pour une vraie PWA (hors ligne + installation) : servir le dossier avec un serveur local (ex. `npx serve .` ou `python -m http.server`).

## Fichier d’exemple

`sample-crew-duties.xlsx` – planning Q1 2026 (4 personnes).

## Technologies

- HTML / CSS / JavaScript vanilla
- SheetJS (xlsx) pour le parsing Excel
- Chart.js pour les graphiques
- localStorage pour la persistance
- Service Worker + Web App Manifest
