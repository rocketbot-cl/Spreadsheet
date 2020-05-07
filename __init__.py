# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Funcion que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"

    pip install <package> -t .

"""
import os.path

base_path = tmp_global_obj["basepath"]
cur_path = base_path + 'modules' + os.sep + \
           'Spreadsheet' + os.sep + 'libs' + os.sep
sys.path.append(cur_path)
import gspread
from collections import namedtuple

"""
    Obtengo el modulo que fueron invocados
"""
module = GetParams("module")

global gc

open_option = {
    "title": gc.open,
    "key": gc.open_by_key,
    "url": gc.open_by_url
}

if module == "config":
    credential_path = GetParams("credentials_path")

    try:
        print(credential_path)
        if not os.path.exists(credential_path):
            raise Exception(
                "El archivo de credenciales no existe en la ruta especificada")

        gc = gspread.service_account(credential_path)

    except Exception as e:
        PrintException()
        raise e

if module == "CreateSheet":
    if not gc:
        raise Exception(
            "No hay credenciales ni token válidos, por favor configure sus credenciales")

    try:
        ss_name = GetParams('ss_id')
        search_by = GetParams("by")
        name = GetParams('name')
        result = GetParams("result")

        if not search_by:
            search_by = "title"
        spreadsheet = open_option[search_by](ss_name)
        sheet = spreadsheet.add_worksheet(title=name, rows=100, cols=20)

        if result:
            SetVar(result, sheet.id)
    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if module == "DeleteSheet":
    if not gc:
        raise Exception(
            "No hay credenciales ni token válidos, por favor configure sus credenciales")

    ss_name = GetParams('ss_id')
    search_by = GetParams("by")
    sheet_id = GetParams('sheetName')

    try:
        if not search_by:
            search_by = "title"
        spreadsheet = open_option[search_by](ss_name)
        Sheet = namedtuple("Sheet", "id")(sheet_id)  # Object with id attribute
        sheet = spreadsheet.del_worksheet(Sheet)  # Need an object with an id attribute

    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e
if module == "UpdateRange":

    if not gc:
        raise Exception(
            "No hay credenciales ni token válidos, por favor configure sus credenciales")

    ss_name = GetParams('ss_id')
    search_by = GetParams("by")
    sheet_name = GetParams('sheetName')
    range_ = GetParams('range')
    text = GetParams('text')

    try:

        if ":" in range_:
            text = eval(text) if text.startswith("[[") else [eval(text)]
        if not search_by:
            search_by = "title"
        spreadsheet = open_option[search_by](ss_name)
        sheet = spreadsheet.worksheet(sheet_name)

        sheet.update(range_, text)

    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e
if module == "ReadCells":

    if not gc:
        raise Exception(
            "No hay credenciales ni token válidos, por favor configure sus credenciales")

    ss_name = GetParams('ss_id')
    search_by = GetParams("by")
    sheet_name = GetParams('sheetName')
    range_ = GetParams('range')
    result = GetParams('result')

    try:

        spreadsheet = open_option[search_by](ss_name)
        sheet = spreadsheet.worksheet(sheet_name)
        print(sheet.get_all_values(value_render_option="UNFORMATTED_VALUE"))
        value = sheet.get(range_)

        if result:
            SetVar(result, value)
    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e
if module == "GetSheets":

    if not gc:
        raise Exception(
            "No hay credenciales ni token válidos, por favor configure sus credenciales")

    ss_name = GetParams('ss_id')
    search_by = GetParams("by")
    result = GetParams('result')

    spreadsheet = open_option[search_by](ss_name)
    sheets = spreadsheet.worksheets()
    sheets = [sheet.title for sheet in sheets]

    if result:
        SetVar(result, sheets)
if module == "CountCells":

    if not gc:
        raise Exception(
            "No hay credenciales ni token válidos, por favor configure sus credenciales")

    ss_name = GetParams('ss_id')
    search_by = GetParams("by")
    sheet_name = GetParams('sheetName')
    result = GetParams('result')

    try:

        spreadsheet = open_option[search_by](ss_name)
        sheet = spreadsheet.worksheet(sheet_name)
        cells = sheet.get_all_values(value_render_option="UNFORMATTED_VALUE")

        rows_count = len(cells)
        cols_count = max([len(cell) for cell in cells])
        value = rows_count, cols_count

        if result:
            SetVar(result, value)
    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if module == "DeleteColumn":
    if not creds:
        raise Exception(
            "No hay credenciales ni token válidos, por favor configure sus credenciales")

    ss_id = GetParams('ss_id')
    sheet = GetParams("sheetName")
    col = GetParams('column').lower()
    blank = GetParams('blank')

    try:
        service = discovery.build('sheets', 'v4', credentials=creds)

        data = service.spreadsheets().get(spreadsheetId=ss_id).execute()

        for element in data["sheets"]:
            if element["properties"]["title"] == sheet:
                sheet_id = element["properties"]["index"]

        abc = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u',
               'v', 'w', 'x', 'y', 'z']

        around_abc = len(col) - 1
        col = col[-1]
        col_index = around_abc * len(abc) + abc.index(col)
        print(col_index)

        if blank:
            shiftDimension = "ROWS"
        else:
            shiftDimension = "COLUMNS"

        body = {
            "requests": [{
                "deleteRange": {
                    "range": {
                        "sheetId": sheet_id,
                        "startColumnIndex": col_index,
                        "endColumnIndex": col_index + 1
                    },
                    "shiftDimension": shiftDimension
                }
            }]
        }

        request = service.spreadsheets().batchUpdate(spreadsheetId=ss_id, body=body)
        response = request.execute()

    except Exception as e:
        PrintException()
        raise e
if module == "DeleteRow":
    if not creds:
        raise Exception(
            "No hay credenciales ni token válidos, por favor configure sus credenciales")

    ss_id = GetParams('ss_id')
    sheet = GetParams("sheetName")
    row = GetParams('row')
    blank = GetParams('blank')

    try:

        service = discovery.build('sheets', 'v4', credentials=creds)

        data = service.spreadsheets().get(spreadsheetId=ss_id).execute()

        for element in data["sheets"]:
            if element["properties"]["title"] == sheet:
                sheet_id = element["properties"]["sheetId"]

        if blank:
            shiftDimension = "COLUMNS"
        else:
            shiftDimension = "ROWS"

        if row.find("-") == -1:
            row = int(row)
            body = {
                "requests": [{
                    "deleteRange": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": row - 1,
                            "endRowIndex": row
                        },
                        "shiftDimension": shiftDimension
                    }
                }]
            }
        else:
            row = row.split("-")
            body = {
                "requests": [{
                    "deleteRange": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": int(row[0]) - 1,
                            "endRowIndex": int(row[1])
                        },
                        "shiftDimension": shiftDimension
                    }
                }]
            }

        request = service.spreadsheets().batchUpdate(spreadsheetId=ss_id, body=body)
        response = request.execute()

    except Exception as e:
        PrintException()
        raise e
