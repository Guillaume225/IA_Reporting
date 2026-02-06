#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script d'analyse de la liasse fiscale Excel
Analyse le fichier 'Liasse Système Normal DCP PROV 2024 Actualisée (1) (1).xlsx'
et génère un rapport détaillé
"""

import pandas as pd
import os
import sys
from datetime import datetime


# Nom du fichier Excel à analyser
FICHIER_EXCEL = "Liasse Système Normal DCP PROV 2024 Actualisée (1) (1).xlsx"
FICHIER_RAPPORT = "rapport_analyse.txt"


def afficher_separateur(fichier=None):
    """Affiche un séparateur visuel"""
    separateur = "=" * 80
    print(separateur)
    if fichier:
        fichier.write(separateur + "\n")


def afficher_titre(titre, fichier=None):
    """Affiche un titre formaté"""
    print(f"\n{titre}")
    print("-" * len(titre))
    if fichier:
        fichier.write(f"\n{titre}\n")
        fichier.write("-" * len(titre) + "\n")


def verifier_fichier_existe(chemin_fichier):
    """Vérifie si le fichier existe"""
    if not os.path.exists(chemin_fichier):
        print(f"❌ ERREUR : Le fichier '{chemin_fichier}' n'a pas été trouvé.")
        print(f"   Répertoire actuel : {os.getcwd()}")
        return False
    return True


def lire_fichier_excel(chemin_fichier):
    """
    Lit le fichier Excel et retourne un dictionnaire de DataFrames
    
    Args:
        chemin_fichier: Chemin vers le fichier Excel
        
    Returns:
        dict: Dictionnaire {nom_feuille: DataFrame}
    """
    try:
        print(f"📖 Lecture du fichier : {chemin_fichier}")
        # Lire toutes les feuilles du fichier Excel
        donnees = pd.read_excel(chemin_fichier, sheet_name=None, engine='openpyxl')
        print(f"✅ Fichier lu avec succès : {len(donnees)} feuille(s) trouvée(s)")
        return donnees
    except FileNotFoundError:
        print(f"❌ ERREUR : Fichier non trouvé - {chemin_fichier}")
        return None
    except Exception as e:
        print(f"❌ ERREUR lors de la lecture : {str(e)}")
        return None


def analyser_structure_feuille(nom_feuille, df):
    """
    Analyse la structure d'une feuille Excel
    
    Args:
        nom_feuille: Nom de la feuille
        df: DataFrame pandas
        
    Returns:
        dict: Informations sur la structure
    """
    info = {
        'nom': nom_feuille,
        'dimensions': df.shape,
        'nb_lignes': df.shape[0],
        'nb_colonnes': df.shape[1],
        'colonnes': list(df.columns),
        'vide': df.empty
    }
    return info


def analyser_donnees_feuille(df):
    """
    Analyse les données d'une feuille
    
    Args:
        df: DataFrame pandas
        
    Returns:
        dict: Analyse des données
    """
    analyse = {
        'types_colonnes': {},
        'valeurs_manquantes': {},
        'statistiques_numeriques': {}
    }
    
    # Analyser chaque colonne
    for colonne in df.columns:
        # Type de données
        analyse['types_colonnes'][colonne] = str(df[colonne].dtype)
        
        # Valeurs manquantes
        nb_manquantes = df[colonne].isna().sum()
        pct_manquantes = (nb_manquantes / len(df)) * 100 if len(df) > 0 else 0
        analyse['valeurs_manquantes'][colonne] = {
            'nombre': nb_manquantes,
            'pourcentage': round(pct_manquantes, 2)
        }
        
        # Statistiques pour colonnes numériques
        if pd.api.types.is_numeric_dtype(df[colonne]):
            try:
                stats = {
                    'somme': df[colonne].sum(),
                    'moyenne': df[colonne].mean(),
                    'min': df[colonne].min(),
                    'max': df[colonne].max(),
                    'mediane': df[colonne].median()
                }
                analyse['statistiques_numeriques'][colonne] = stats
            except Exception as e:
                analyse['statistiques_numeriques'][colonne] = f"Erreur: {str(e)}"
    
    return analyse


def afficher_analyse_feuille(nom_feuille, df, fichier_rapport):
    """
    Affiche et écrit l'analyse complète d'une feuille
    
    Args:
        nom_feuille: Nom de la feuille
        df: DataFrame pandas
        fichier_rapport: Objet fichier pour écrire le rapport
    """
    afficher_separateur(fichier_rapport)
    afficher_titre(f"📋 Feuille : {nom_feuille}", fichier_rapport)
    
    # Structure
    structure = analyser_structure_feuille(nom_feuille, df)
    
    if structure['vide']:
        message = "⚠️  Cette feuille est vide"
        print(message)
        fichier_rapport.write(message + "\n")
        return
    
    print(f"\n📏 Dimensions : {structure['nb_lignes']} lignes × {structure['nb_colonnes']} colonnes")
    fichier_rapport.write(f"\nDimensions : {structure['nb_lignes']} lignes × {structure['nb_colonnes']} colonnes\n")
    
    # Liste des colonnes
    print(f"\n📊 Colonnes ({len(structure['colonnes'])}) :")
    fichier_rapport.write(f"\nColonnes ({len(structure['colonnes'])}) :\n")
    for i, col in enumerate(structure['colonnes'], 1):
        ligne = f"  {i}. {col}"
        print(ligne)
        fichier_rapport.write(ligne + "\n")
    
    # Analyse des données
    analyse = analyser_donnees_feuille(df)
    
    # Types de données
    print(f"\n🔤 Types de données :")
    fichier_rapport.write(f"\nTypes de données :\n")
    for col, dtype in analyse['types_colonnes'].items():
        ligne = f"  • {col}: {dtype}"
        print(ligne)
        fichier_rapport.write(ligne + "\n")
    
    # Valeurs manquantes
    print(f"\n❓ Valeurs manquantes :")
    fichier_rapport.write(f"\nValeurs manquantes :\n")
    for col, info in analyse['valeurs_manquantes'].items():
        if info['nombre'] > 0:
            ligne = f"  • {col}: {info['nombre']} ({info['pourcentage']}%)"
            print(ligne)
            fichier_rapport.write(ligne + "\n")
    
    # Statistiques numériques
    if analyse['statistiques_numeriques']:
        print(f"\n📈 Statistiques des colonnes numériques :")
        fichier_rapport.write(f"\nStatistiques des colonnes numériques :\n")
        for col, stats in analyse['statistiques_numeriques'].items():
            print(f"\n  {col} :")
            fichier_rapport.write(f"\n  {col} :\n")
            if isinstance(stats, dict):
                for stat, valeur in stats.items():
                    ligne = f"    - {stat}: {valeur:,.2f}" if isinstance(valeur, (int, float)) else f"    - {stat}: {valeur}"
                    print(ligne)
                    fichier_rapport.write(ligne + "\n")
            else:
                print(f"    {stats}")
                fichier_rapport.write(f"    {stats}\n")
    
    # Aperçu des données
    print(f"\n👀 Aperçu des premières lignes :")
    fichier_rapport.write(f"\nAperçu des premières lignes :\n")
    apercu = df.head(5).to_string()
    print(apercu)
    fichier_rapport.write(apercu + "\n")


def exporter_csv(nom_feuille, df, dossier="exports_csv"):
    """
    Exporte une feuille en fichier CSV
    
    Args:
        nom_feuille: Nom de la feuille
        df: DataFrame pandas
        dossier: Dossier de destination
    """
    try:
        # Créer le dossier s'il n'existe pas
        if not os.path.exists(dossier):
            os.makedirs(dossier)
        
        # Nettoyer le nom de fichier
        nom_fichier = nom_feuille.replace("/", "_").replace("\\", "_").replace(" ", "_")
        chemin_csv = os.path.join(dossier, f"{nom_fichier}.csv")
        
        # Exporter
        df.to_csv(chemin_csv, index=False, encoding='utf-8-sig')
        print(f"  ✅ Exporté : {chemin_csv}")
        return True
    except Exception as e:
        print(f"  ❌ Erreur export {nom_feuille} : {str(e)}")
        return False


def generer_rapport_complet(chemin_fichier):
    """
    Génère un rapport complet d'analyse du fichier Excel
    
    Args:
        chemin_fichier: Chemin vers le fichier Excel
    """
    # Vérifier l'existence du fichier
    if not verifier_fichier_existe(chemin_fichier):
        return False
    
    # Lire le fichier
    donnees = lire_fichier_excel(chemin_fichier)
    if donnees is None:
        return False
    
    # Ouvrir le fichier de rapport
    try:
        with open(FICHIER_RAPPORT, 'w', encoding='utf-8') as fichier_rapport:
            # En-tête du rapport
            afficher_separateur(fichier_rapport)
            titre = "RAPPORT D'ANALYSE DE LA LIASSE FISCALE"
            print(f"\n{titre}")
            print("=" * len(titre))
            fichier_rapport.write(f"\n{titre}\n")
            fichier_rapport.write("=" * len(titre) + "\n")
            
            # Informations générales
            date_analyse = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            print(f"📅 Date d'analyse : {date_analyse}")
            print(f"📁 Fichier analysé : {chemin_fichier}")
            print(f"📊 Nombre de feuilles : {len(donnees)}")
            
            fichier_rapport.write(f"\nDate d'analyse : {date_analyse}\n")
            fichier_rapport.write(f"Fichier analysé : {chemin_fichier}\n")
            fichier_rapport.write(f"Nombre de feuilles : {len(donnees)}\n")
            
            # Résumé global
            afficher_titre("📌 RÉSUMÉ GLOBAL", fichier_rapport)
            total_lignes = sum(df.shape[0] for df in donnees.values())
            total_colonnes = sum(df.shape[1] for df in donnees.values())
            print(f"Total de lignes : {total_lignes}")
            print(f"Total de colonnes : {total_colonnes}")
            fichier_rapport.write(f"Total de lignes : {total_lignes}\n")
            fichier_rapport.write(f"Total de colonnes : {total_colonnes}\n")
            
            print("\nListe des feuilles :")
            fichier_rapport.write("\nListe des feuilles :\n")
            for i, (nom, df) in enumerate(donnees.items(), 1):
                ligne = f"  {i}. {nom} ({df.shape[0]} lignes × {df.shape[1]} colonnes)"
                print(ligne)
                fichier_rapport.write(ligne + "\n")
            
            # Analyse détaillée de chaque feuille
            afficher_titre("\n📋 ANALYSE DÉTAILLÉE PAR FEUILLE", fichier_rapport)
            
            for nom_feuille, df in donnees.items():
                afficher_analyse_feuille(nom_feuille, df, fichier_rapport)
            
            # Pied de page
            afficher_separateur(fichier_rapport)
            message_fin = f"\n✅ Rapport généré avec succès : {FICHIER_RAPPORT}"
            print(message_fin)
            fichier_rapport.write(message_fin + "\n")
            
        # Proposer l'export CSV
        print("\n" + "=" * 80)
        reponse = input("\n💾 Voulez-vous exporter les données en CSV ? (o/n) : ").strip().lower()
        if reponse in ['o', 'oui', 'y', 'yes']:
            print("\n📤 Export des feuilles en CSV...")
            for nom_feuille, df in donnees.items():
                if not df.empty:
                    exporter_csv(nom_feuille, df)
            print("\n✅ Export CSV terminé")
        
        return True
        
    except Exception as e:
        print(f"\n❌ ERREUR lors de la génération du rapport : {str(e)}")
        return False


def main():
    """Fonction principale"""
    print("\n" + "=" * 80)
    print("🔍 ANALYSE DE LA LIASSE FISCALE")
    print("=" * 80 + "\n")
    
    # Chemin du fichier
    chemin_fichier = os.path.join(os.getcwd(), FICHIER_EXCEL)
    
    # Générer le rapport
    succes = generer_rapport_complet(chemin_fichier)
    
    if succes:
        print("\n" + "=" * 80)
        print("✅ Analyse terminée avec succès !")
        print("=" * 80 + "\n")
        sys.exit(0)
    else:
        print("\n" + "=" * 80)
        print("❌ L'analyse a échoué")
        print("=" * 80 + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
