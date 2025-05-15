import os
import sys
import json
import pandas as pd

path_excel = "translate.xlsx"  
sheet_name = "translate"        
path_output = "json_output"     


if not os.path.exists(path_excel):
    print(f"❌ ERREUR : Le fichier '{path_excel}' n'existe pas.")
    sys.exit(1)


try:
    df = pd.read_excel(path_excel, sheet_name=sheet_name, header=0)  
    df.columns = df.columns.str.strip()  
    df.columns = df.columns.str.upper()  

    print(f"📌 Colonnes détectées : {df.columns.tolist()}")  

    
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
                        

  
    os.makedirs(path_output, exist_ok=True)

  
    for lang, data in translations.items():
        json_path = os.path.join(path_output, f"{lang.lower()}.json")
        with open(json_path, "w", encoding="utf-8") as json_file:
            json.dump(data, json_file, ensure_ascii=False, indent=4)
        print(f"✅ Fichier JSON créé : {json_path}")

except Exception as e:
    print(f"❌ Erreur lors du traitement du fichier Excel : {e}")
    sys.exit(1)
