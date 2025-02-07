```shell
python -m venv excel-to-json  # créer un environnement virtuel surlequel on va travailler
excel-to-json/Scripts/activate # activier l'env virtuel
py -m pip install -U pip setuptools wheel #met à jour les outils essentiels pour la gestion des paquets Python : pip, setuptools et wheel, garantissant ainsi que vous utilisez leurs dernières versions.
pip install -r requirements.txt
```
# Excel to JSON Converter

Ce projet permet de convertir un fichier Excel de traduction (ou autre type de données structurées) en plusieurs fichiers JSON, organisés par langue. Chaque fichier JSON contiendra des identifiants et leurs traductions/mots de passe dans la langue correspondante.

## Fonctionnalités

- Lecture d'un fichier Excel contenant des données structurées avec une colonne d'identifiants et plusieurs colonnes pour les traductions/mots de passe dans différentes langues.
- Création de fichiers JSON pour chaque langue, avec les identifiants comme clés et les traductions/mots de passe comme valeurs.

## Prérequis

Avant d'utiliser ce projet, vous devez avoir installé les bibliothèques suivantes :

- `openpyxl` : pour la lecture et l'écriture des fichiers Excel.
- `json` : pour manipuler les fichiers JSON.
- `re` : pour les expressions régulières (utilisé pour valider les codes de langue).

Vous pouvez installer la bibliothèque `openpyxl` avec la commande suivante si vous ne l'avez pas déjà :

```bash
pip install openpyxl
