import os
import re
import sys
import json
import openpyxl
from openpyxl import load_workbook


def transform_excel_to_json(path_excel, sheet_name):
    workbook = load_workbook(filename = path_excel)
    worksheet = workbook[sheet_name]
    print(worksheet.max_row, worksheet.max_column)
    print(worksheet.cell(row=1, column=1).value)
    object_json = {}
    #on parcourt la colonne des indentifiants
    for col in range(2, worksheet.max_column + 1):
        language_code = worksheet.cell(row=1, column=col).value # on récupère le code de la langue
        if language_code and len(language_code) == 2:
            object_json[language_code] = {}
            for row in range(2, worksheet.max_row + 1):
                identifier = worksheet.cell(row=row, column=1).value
                password = worksheet.cell(row=row, column=col).value # on récupère le mot de passe
                object_json[language_code][identifier] = password
    print(f"object_json: {object_json}")
    return object_json

def create_json_files(path_output, json_object):
    for key in json_object.keys():
        if re.search(r'\w{2}', key): # on vérifie que la clé est bien un code de langue
            create_json_file(os.path.join(path_output, f"{key}.json" ), json_object[key])

def create_json_file(path_json, json_obj):
    with open(path_json, "w") as file:
        json.dump(json_obj, file, indent = 4)


#def file_exist(path):
    #return os.path.exists(path) 


def parser_translate_excel_to_json(path_excel, sheet_name, path_output):
    json_object = transform_excel_to_json(path_excel, sheet_name)
    create_json_files(path_output, json_object)

if __name__ == "__main__":

    path_excel = "translate.xlsx"
    sheet_name = "translate"
    path_output = r"C:\workspace\excel_to_json"
    if not os.path.exists(path_excel): # tester que le fichier excel existe
        sys.exit(0)
    # tester que output exist
    parser_translate_excel_to_json(path_excel, sheet_name, path_output)


