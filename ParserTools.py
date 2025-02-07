import os
import re
import sys
import json
import openpyxl
from openpyxl import load_workbook


def transform_excel_to_json(path_excel, sheet_name):
    wb = load_workbook(filename = path_excel)
    ws = wb[sheet_name]
    print(ws.max_row, ws.max_column)
    print(ws.cell(row=1, column=1).value)
    object_json = {}
    #on parcourt la colonne des indentifiants
    for col in range(2, ws.max_column + 1):
        language_code = ws.cell(row=1, column=col).value # on récupère le code de la langue
        if language_code and len(language_code) == 2:
            object_json[language_code] = {}
            for row in range(2, ws.max_row + 1):
                identifier = ws.cell(row=row, column=1).value
                password = ws.cell(row=row, column=col).value # on récupère le mot de passe
                object_json[language_code][identifier] = password
    print(f"object_json: {object_json}")
    return object_json

    # return {
    #     "FR": {
    #         "identifier": "identifiant",
    #         "password": "mot de passe"
    #     },
    #     "EN": {
    #         "identifier": "identifier",
    #         "password": "password"
    #     },
    # }

def create_json_files(path_output, json_object):
    for key in json_object.keys():
        if re.search(r'\w{2}', key):
            create_json_file(os.path.join(path_output, f"{key}.json" ), json_object[key])

def create_json_file(path_json, json_obj):
    with open(path_json, "w") as file:
        json.dump(json_obj, file, indent = 4)


def file_exist(path):
    return True


def parser_translate_excel_to_json(path_excel, sheet_name, path_output):
    json_object = transform_excel_to_json(path_excel, sheet_name)
    create_json_files(path_output, json_object)

if __name__ == "__main__":

    path_excel = "translate.xlsx"
    sheet_name = "translate"
    #  path_output = r"C:\Users\emerisie\AppData\Local\TESTFACTORY\workspace\PrjFormationPython"
    path_output = r"C:\workspace\excel_to_json"
    if not(file_exist(path_excel)):
        sys.exit(0)
    # tester que output exist
    parser_translate_excel_to_json(path_excel, sheet_name, path_output)


