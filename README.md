# IA_Reporting
Projet de reporting porté sur l'intelligence artificielle

## Description

Ce projet contient un script Python pour analyser des liasses fiscales au format Excel (Système Normal OHADA/DGI).

Le script `analyse_liasse.py` permet de :
- 📊 Explorer la structure du fichier Excel (feuilles, dimensions, colonnes)
- 🔍 Analyser les données (types, valeurs manquantes, statistiques)
- 📝 Générer un rapport de synthèse détaillé
- 💾 Exporter les données nettoyées en CSV (optionnel)
- ⚠️  Gérer les erreurs de lecture et les feuilles vides

## Fichier analysé

Le script analyse le fichier : `Liasse Système Normal DCP PROV 2024 Actualisée (1) (1).xlsx`

Ce fichier doit être présent à la racine du dépôt.

## Prérequis

- Python 3.8 ou supérieur
- pip (gestionnaire de paquets Python)

## Installation

1. Cloner le dépôt :
```bash
git clone https://github.com/Guillaume225/IA_Reporting.git
cd IA_Reporting
```

2. Installer les dépendances :
```bash
pip install -r requirements.txt
```

Les dépendances installées sont :
- `pandas` : Pour la manipulation et l'analyse des données
- `openpyxl` : Pour la lecture des fichiers Excel

## Utilisation

### Exécuter le script d'analyse

```bash
python analyse_liasse.py
```

Le script va :
1. Lire le fichier Excel
2. Analyser chaque feuille
3. Générer un rapport détaillé dans `rapport_analyse.txt`
4. Proposer d'exporter les données en CSV

### Fichiers générés

- **rapport_analyse.txt** : Rapport complet de l'analyse avec :
  - Résumé global du fichier
  - Analyse détaillée de chaque feuille
  - Types de données par colonne
  - Statistiques descriptives
  - Valeurs manquantes
  - Aperçu des premières lignes

- **exports_csv/** (optionnel) : Dossier contenant les exports CSV de chaque feuille

## Structure du projet

```
IA_Reporting/
│
├── analyse_liasse.py                              # Script principal d'analyse
├── requirements.txt                                # Dépendances Python
├── README.md                                       # Documentation
├── .gitignore                                      # Fichiers à ignorer
└── Liasse Système Normal DCP PROV 2024 Actualisée (1) (1).xlsx  # Fichier Excel à analyser
```

## Fonctionnalités détaillées

### Exploration de la structure
- Liste de toutes les feuilles disponibles
- Dimensions (lignes × colonnes) de chaque feuille
- Noms des colonnes de chaque feuille

### Analyse des données
- Identification des types de données par colonne
- Détection des valeurs manquantes (nombre et pourcentage)
- Calcul des statistiques descriptives pour les colonnes numériques :
  - Somme
  - Moyenne
  - Minimum
  - Maximum
  - Médiane

### Gestion des erreurs
- Vérification de l'existence du fichier
- Gestion des feuilles vides
- Gestion des erreurs de lecture
- Messages d'erreur explicites en français

## Auteur

Guillaume225

## Licence

Ce projet est open source.
