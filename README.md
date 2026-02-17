# Carbonometre (MVP local)

Application web locale pour saisir des postes d'emissions, calculer le CO2e et exporter/reimporter un fichier Excel.

## Lancer en local

Option rapide:

```bash
./launch.sh
```

Option manuelle (`pip`):

```bash
python3 -m venv carbo
source carbo/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Puis ouvrir `http://localhost:8501`.

## Fonctionnalites principales

- Modes: `item_unique`, `bilan_personnel`, `bilan_projet`
- Postes: `achats`, `domicile_travail`, `campagnes_terrain`, `missions`, `heures_calcul`
- Identite: anonyme par defaut, equipe optionnelle, type de poste
- `achats`: type d'achat via menu deroulant avec facteur par defaut pre-rempli
- `missions`: calcul via `Moulinette_missions` a partir des villes/pays de depart et d'arrivee, avec cartes
- Export/reimport Excel pour reprendre un dossier
- Sauvegarde locale:
  - `bilan_personnel`: `nom_bilan_annuel.xlsx`
  - `bilan_projet`: `nom_nom_projet_projet.xlsx`

## Arborescence de stockage local

Les formulaires sont sauvegardes dans:

```text
./<lab_id>/<annee>/<equipe>/*.xlsx
```

Par defaut (CEREGE):

```text
./cerege/<annee>/<equipe>/*.xlsx
```

## Vue d'ensemble du labo

- Consolidation automatique de tous les formulaires locaux du dossier du labo courant
- Histogramme d'evolution interannuelle
- Filtres d'analyse (annee, poste, equipe, statut, type de poste)

## Configuration multi-labos

L'application charge une config labo depuis `labs/<lab_id>.toml`.

- valeur par defaut: `lab_id=cerege`
- logo, equipes et palette de couleurs sont parametres par labo

Pour lancer un autre labo sans modifier le code:

```bash
LAB_ID=monlabo streamlit run app.py
```

## Notes

- Aucun champ `notes` dans les formulaires.
- Facteurs d'emission v1 simplifies et modifiables dans le formulaire.
