# Publipostage DOETH

Automatisation de la génération d'attestations DOETH (Déclaration Obligatoire d'Emploi des Travailleurs Handicapés) à partir de fichiers Excel. Produit des attestations Word et/ou PDF, regroupées par SIRET, prêtes à être envoyées aux clients.

---

## Sommaire

1. [Vue d'ensemble](#vue-densemble)
2. [Prérequis](#prérequis)
3. [Installation](#installation)
4. [Configuration](#configuration)
5. [Utilisation](#utilisation)
6. [Architecture](#architecture)
7. [Flux de traitement](#flux-de-traitement)
8. [Format des données d'entrée](#format-des-données-dentrée)
9. [Développement](#développement)
10. [Dépannage](#dépannage)

---

## Vue d'ensemble

L'application lit un fichier Excel BOETH, nettoie et regroupe les données par SIRET, puis génère une attestation par entreprise cliente. Chaque attestation inclut l'en-tête société, les informations légales, le tableau des bénéficiaires avec leurs ETP, et la signature du représentant légal.

**Formats de sortie :** Word (`.docx`), PDF (`.pdf`), ou les deux simultanément.

**Deux modes d'accès :**
- Interface graphique (`gui.py`) pour une utilisation quotidienne
- Ligne de commande (`main.py`) pour l'automatisation et les tests

---

## Prérequis

| Composant | Version minimale | Rôle |
|---|---|---|
| Python | 3.12 | Runtime |
| Microsoft Word | 2016+ | Conversion PDF (COM Windows) |
| uv | latest | Gestionnaire d'environnement |

**Dépendances Python** (installées automatiquement) :

```
pandas >= 2.0
python-docx >= 1.1
openpyxl >= 3.1
pyyaml >= 6.0
pywin32 >= 308
colorama >= 0.4
```

---

## Installation

```bash
# 1. Cloner le dépôt
git clone <URL_REPO>
cd publipostage_doeth

# 2. Installer les dépendances
uv sync

# 3. Vérifier l'installation
uv run python main.py --help
```

**Structure des dossiers à créer** (ignorés par Git, à initialiser une fois) :

```bash
mkdir -p data/input data/processed data/output logs
touch data/input/.gitkeep data/processed/.gitkeep data/output/.gitkeep
```

---

## Configuration

Tous les paramètres sont centralisés dans `config.yaml` à la racine :

```yaml
paths:
  base_dir: D:\Data\01_Projects\Active\publipostage_doeth
  input_dir:     ${paths.data_dir}/input
  processed_dir: ${paths.data_dir}/processed
  output_dir:    ${paths.data_dir}/output
  logs_dir:      ${paths.base_dir}/logs

resources:
  logo_path:      ${paths.resources_dir}/images/Entete_GI.png
  signature_path: ${paths.resources_dir}/images/RL_LG.png

defaults:
  input_filename: "Liste des BOETH par RGP CLI avec adresses V2.xlsx"
  excel_sheet:    "Feuil1"
  csv_separator:  ";"
  date_format:    "%d/%m/%Y"

document:
  font_size:       10
  table_font_size:  8
  margins:          1.5   # cm

representant:
  nom:    "Loïc GALLERAND"
  adresse: "233 rue de Chateaugiron a Rennes (35000)"
  siret:   "49342093900057"
```

Adapter `base_dir` et `representant` a votre environnement. Les autres chemins sont resolus dynamiquement.

---

## Utilisation

### Interface graphique

```bash
uv run python gui.py
```

Les parametres disponibles dans l'interface :

| Champ | Description |
|---|---|
| Fichier Excel | Chemin vers le fichier BOETH source |
| Feuille Excel | Nom de l'onglet a traiter |
| Dossier de sortie | Ou enregistrer les attestations |
| Logo / Signature | Images a integrer dans les documents |
| Format de sortie | Word, PDF, ou Les deux |
| Ignorer traitement Excel | Utiliser un CSV deja genere |
| Mode debug | Logs detailles (niveau DEBUG) |

### Ligne de commande

```bash
# Traitement complet avec les valeurs du config.yaml
uv run python main.py

# Fichier et feuille explicites
uv run python main.py \
  --input "data/input/Liste des BOETH par RGP CLI avec adresses V2.xlsx" \
  --sheet "Liste des BOETH par RGP CLI ave"

# Choisir le format de sortie
uv run python main.py --format pdf      # PDF uniquement
uv run python main.py --format both     # Word + PDF
uv run python main.py --format docx     # Word uniquement (defaut)

# Sauter le traitement Excel - utiliser un CSV existant
uv run python main.py \
  --skip-processing \
  --csv-path "data/processed/processed_20260224_121027.csv"

# Mode debug
uv run python main.py --debug
```

**Reference complete des arguments :**

| Argument | Valeurs | Defaut | Description |
|---|---|---|---|
| `--input` | chemin | config.yaml | Fichier Excel source |
| `--sheet` | texte | config.yaml | Nom de la feuille |
| `--output-dir` | chemin | config.yaml | Dossier de sortie |
| `--format` | `docx` `pdf` `both` | `docx` | Format des attestations |
| `--skip-processing` | flag | False | Passer directement a la generation |
| `--csv-path` | chemin | auto | CSV source (avec `--skip-processing`) |
| `--debug` | flag | False | Logs niveau DEBUG |

### Fichiers produits

```
data/
├── processed/
│   └── processed_YYYYMMDD_HHMMSS.csv     # Donnees intermediaires
└── output/
    ├── 1_Attestation DOETH_2025_ADAPEI.docx
    ├── 1_Attestation DOETH_2025_ADAPEI.pdf
    ├── 2_Attestation DOETH_2025_SNEF.docx
    └── ...
logs/
    └── publipostage_doeth_prod_YYYYMMDD.log
```

---

## Architecture

```
publipostage_doeth/
│
├── config.yaml              # Configuration centralisee
│
├── data/                    # <- ignore par Git
│   ├── input/               # Fichiers Excel sources
│   ├── processed/           # CSV intermediaires horodates
│   └── output/              # Attestations generees
│
├── resources/               # <- ignore par Git
│   └── images/
│       ├── Entete_GI.png    # Logo en-tete
│       └── RL_LG.png        # Signature representant legal
│
├── logs/                    # <- ignore par Git
│
├── data_processor.py        # Etapes 1-6 : Excel -> CSV propre
├── document_generator.py    # Generation Word, enum OutputFormat
├── pdf_converter.py         # Conversion DOCX->PDF via COM Word
├── config.py                # Chargement et resolution config.yaml
├── logger.py                # Logger colore + rotation fichier
├── error_handling.py        # Decorateurs et handlers d'erreurs
├── gui.py                   # Interface graphique Tkinter
├── main.py                  # Point d'entree CLI
│
├── pyproject.toml
├── requirements.txt
├── .gitignore
└── README.md
```

**Separation des responsabilites :**

| Module | Role |
|---|---|
| `config.py` | Config - chargement unique au demarrage, resolution des variables |
| `main.py` | Orchestrateur - sequence les etapes, gere les parametres CLI |
| `data_processor.py` | Processor - logique de transformation des donnees |
| `document_generator.py` | Processor - creation des documents Word |
| `pdf_converter.py` | Processor - conversion batch DOCX->PDF, instance COM unique |

---

## Flux de traitement

```
Fichier Excel
     |
     v
[1] Chargement           load_excel_data()
     |
     v
[2] Nettoyage            clean_and_transform_data()
     |  . Creation SIRET (SIREN + NIC, validation numerique)
     |  . Formatage DATE_NAISSANCE -> str uniforme
     |  . Conversion ETP_ANNUEL / NB_HEURES en float, NaN -> 0
     |
     v
[3] Agregation           aggregate_data()
     |  . groupby par toutes les colonnes cles
     |  . sum(ETP_ANNUEL, NB_HEURES)
     |  . ANNEE force en int
     |
     v
[4] Filtrage             filter_data()
     |  . Exclusion CODE_REGROUPEMENT = 'DIFFUS'
     |
     v
[5] Enrichissement       add_processing_columns()
     |  . Tri SIRET / NOM / PRENOM
     |  . NOUVEAU_GROUPE, FIN_GROUPE
     |
     v
[6] Sauvegarde CSV       save_processed_data()   -> data/processed/
     |  . QUOTE_NONNUMERIC (symetrique lecture/ecriture)
     |
     v
[7] Generation Word      generer_attestations_doeth()
     |  . 1 document par SIRET unique
     |  . Logo, en-tete client, references legales
     |  . Tableau beneficiaires + total ETP
     |  . Signature representant legal
     |
     v
[8] Conversion PDF       convert_batch()          (si format != docx)
     |  . Instance Word COM ouverte une seule fois pour le batch
     |  . Fermeture garantie (context manager RAII)
     |
     v
Attestations finales     -> data/output/
```

---

## Format des donnees d'entree

Le fichier Excel doit contenir les colonnes suivantes (noms exacts) :

| Colonne | Type | Description |
|---|---|---|
| `CODE_REGROUPEMENT` | str | Code groupe (ex: `ADAPE01`, `DIFFUS`) |
| `REGROUPEMENT` | str | Libelle du regroupement |
| `SIREN` | int | Code SIREN 9 chiffres |
| `NIC` | int | Code NIC 5 chiffres |
| `NOM_CLIENT` | str | Raison sociale du client |
| `ADRESSE_CLIENT` | str | Adresse postale |
| `CP_CLIENT` | int | Code postal |
| `VILLE_CLIENT` | str | Ville |
| `APE` | str | Code APE |
| `NOM` | str | Nom du beneficiaire |
| `PRENOM` | str | Prenom du beneficiaire |
| `DATE_NAISSANCE` | date | Format JJ/MM/AAAA |
| `ANNEE` | int | Annee de declaration |
| `QUALIFICATION` | str | Qualification du poste |
| `ETP_MAJORE` | str | `oui` / `non` |
| `ETP_ANNUEL` | float | ETP annuel du beneficiaire |
| `NB_HEURES` | float | Nombre d'heures travaillees |

**Regles de qualite appliquees automatiquement :**

- `CODE_REGROUPEMENT = 'DIFFUS'` : lignes exclues silencieusement.
- SIREN/NIC non numeriques (ex: `ST MALO`) : exclus avec WARNING dans les logs.
- `ETP_ANNUEL` / `NB_HEURES` manquants : remplaces par `0`.
- `DATE_NAISSANCE` non convertible : remplacee par une chaine vide.

---

## Developpement

### Branches Git

| Branche | Usage |
|---|---|
| `main` | Production stable - generation Word validee (`v1.0-word-stable`) |
| `develop` | Developpement actif - export PDF + selection format |

### Workflow

```bash
# Travailler sur develop
git checkout develop

# Fusionner et tagger une release sur main
git checkout main
git merge develop
git tag v1.x-description
git push origin main --tags
```

### Tests manuels sans IHM

```bash
# Pipeline complet
uv run python main.py \
  --input "data/input/fichier.xlsx" \
  --sheet "Nom feuille" \
  --format both \
  --debug

# Generation seule (CSV deja pret)
uv run python main.py \
  --skip-processing \
  --csv-path "data/processed/processed_YYYYMMDD_HHMMSS.csv" \
  --format pdf
```

---

## Depannage

**`'<' not supported between instances of 'int' and 'datetime.datetime'`**
Colonne `DATE_NAISSANCE` : valeurs non converties restant en type `int` lors du groupby. Resolu par `.fillna('')` post-strftime dans `data_processor.py`.

**`'float' object is not iterable` dans python-docx**
Cellule contenant `NaN` (cellule Excel vide). Resolu par guard `val == val` dans `document_generator.py`.

**SIRET invalide du type `00ST MALO9521Z`**
SIREN non numerique zero-padde. Resolu par validation `str.isdigit()` dans `create_siret_column()`. Les lignes invalides sont exclues avec un WARNING.

**Adresse client absente dans le Word genere**
Cles `ADRESSE CLIENT` (espace) au lieu de `ADRESSE_CLIENT` (underscore). Resolu dans `add_client_header()`.

**PDF non genere - `Impossible d'ouvrir Microsoft Word`**
Microsoft Word doit etre installe sur le poste. La conversion PDF utilise l'automation COM Windows (`pywin32`), sans dependance tierce supplementaire.