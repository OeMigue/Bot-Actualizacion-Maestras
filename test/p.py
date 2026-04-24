import sys, pythoncom, time, os, psutil
import pandas as pd
import win32com.client.gencache
from tabulate import tabulate

sys.stdout.reconfigure(encoding="utf8")


RUTA_ENTRADA  = r"O:\Gerencia Contraloria\Analitica Contraloria\Tiendas Franquicias\MAESTRAS BASE DE DATOS"
RUTA_SALIDA = r"O:\Gerencia Contraloria\Analitica Contraloria\Automatiaciones Ambiente Pruebas\Carpeta Miguel Cardona\Bot Actualizacion Maestras\output"
RUTA_DOCS = r"O:\Gerencia Contraloria\Analitica Contraloria\Automatiaciones Ambiente Pruebas\Carpeta Miguel Cardona\Bot Actualizacion Maestras\docs"
RUTA_001 = r"O:\Gerencia Contraloria\Analitica Contraloria\Automatiaciones Ambiente Pruebas\Carpeta Miguel Cardona\Bot Actualizacion Maestras\docs\001.xlsx"
RUTA_002 = r"O:\Gerencia Contraloria\Analitica Contraloria\Automatiaciones Ambiente Pruebas\Carpeta Miguel Cardona\Bot Actualizacion Maestras\docs\002.xlsx"

def cerrar_instancias_excel():
    for proceso in psutil.process_iter(['name']):
        try:
            if proceso.info['name'] == 'EXCEL.EXE':
                proceso.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass


def eliminar_espacios(texto):
    return texto.replace(" ", "")


def remplazo_datos():
    excel = None
    try:
        cerrar_instancias_excel()
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        excel.ScreenUpdating = True
        excel.DisplayAlerts = False
        df_001 = pd.read_excel(RUTA_001)
        workbook = excel.Workbooks.Open(RUTA_002)
        hoja = workbook.Worksheets(1)
        excel.ScreenUpdating = False
        hoja.Cells.Clear()

        df_001 = df_001.fillna("")
        df_001 = df_001.astype(object)

        data = [df_001.columns.tolist()] + df_001.values.tolist()
        data = tuple([tuple(row) for row in data])

        hoja.Range('A1').Resize(len(data), len(data[0])).Value = data
        for col, nombre_col in enumerate(df_001.columns, start = 1):
            hoja.Cells(1, col).Value = nombre_col

        for fila, row in enumerate(df_001.values, start = 2):
            for col,valor in enumerate(row, start=1):
                hoja.Cells(fila,col).Value = valor
        excel.ScreenUpdating = True
        workbook.Save()
        workbook.Close(False)

        df_verificacion = pd.read_excel(RUTA_002)

        print("Original:", df_001.shape)
        print("Destino:", df_verificacion.shape)

        print("Escribiendo en:", RUTA_002)
        print("Columnas:", df_001.columns)
        print("Filas:", len(df_001))

        print(type(data))
        print(type(data[0]))
        print(data[0])
        print(data[1])

    except Exception as e:
        print(f'❌ ERROR ----- {e}')
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
remplazo_datos()