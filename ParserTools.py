import os
import sys
import json
import pandas as pd

# 📌 Chemins et fichiers
path_excel = "translate.xlsx"  # Fichier Excel à traiter
sheet_name = "translate"        # Nom de la feuille Excel
path_output = "json_output"     # Dossier où sauvegarder les JSON

# 📂 Vérifier si le fichier Excel existe
if not os.path.exists(path_excel):
    print(f"❌ ERREUR : Le fichier '{path_excel}' n'existe pas.")
    sys.exit(1)

# 📖 Lire le fichier Excel
try:
    df = pd.read_excel(path_excel, sheet_name=sheet_name, header=0)  # Lire le fichier avec pandas
    df.columns = df.columns.str.strip()  # Supprime les espaces cachés autour des noms de colonnes
    df.columns = df.columns.str.upper()  # Met tous les noms de colonnes en majuscules

    print(f"📌 Colonnes détectées : {df.columns.tolist()}")  # Afficher les colonnes trouvées

    # 🔍 Vérifier si "KEYS" existe après nettoyage
    if "KEYS" not in df.columns:
        print("❌ ERREUR : La colonne 'Keys' est introuvable. Vérifiez votre fichier Excel.")
        sys.exit(1)

    # 📌 Vérifier si les langues FR et EN existent
    languages = [lang for lang in ["FR", "EN"] if lang in df.columns]

    if not languages:
        print("❌ ERREUR : Les colonnes 'FR' et 'EN' sont introuvables !")
        sys.exit(1)

    print(f"✅ Langues détectées : {languages}")

    # 🏗 Transformer les données en dictionnaires JSON
    translations = {lang: dict(zip(df["KEYS"], df[lang].fillna(""))) for lang in languages} 
                        # Créer un dictionnaire pour chaque langue # A simplifier
                        #.fillna : remplace les valeurs NaN par une chaîne vide
                        #.zip : fusionne les colonnes KEYS et la langue actuelle
                        #.dict : convertit les tuples en dictionnaires
                        #.df.set_index : définit la colonne KEYS comme index

    # 📂 Créer le dossier de sortie s'il n'existe pas
    os.makedirs(path_output, exist_ok=True)

    # 💾 Sauvegarder chaque langue en fichier JSON
    for lang, data in translations.items():
        json_path = os.path.join(path_output, f"{lang.lower()}.json")
        with open(json_path, "w", encoding="utf-8") as json_file:
            json.dump(data, json_file, ensure_ascii=False, indent=4)
        print(f"✅ Fichier JSON créé : {json_path}")

except Exception as e:
    print(f"❌ Erreur lors du traitement du fichier Excel : {e}")
    sys.exit(1)
