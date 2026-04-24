import sys, pythoncom, time, os, psutil, tabulate
import pandas as pd
import win32com.client.gencache


sys.path.append(r"O:\Gerencia Contraloria\Analitica Contraloria\Automatiaciones Ambiente Pruebas\Carpeta Miguel Cardona\Bot Actualizacion Maestras\config")
sys.path.append(r"O:\Gerencia Contraloria\Analitica Contraloria\Automatiaciones Ambiente Pruebas\Carpeta Miguel Cardona\Bot Actualizacion Maestras\src")
sys.path.append(r"O:\Gerencia Contraloria\Analitica Contraloria\Automatiaciones Ambiente Pruebas\Carpeta Miguel Cardona\Funciones")
sys.stdout.reconfigure(encoding="utf8")

RUTA_ENTRADA  = r"O:\Gerencia Contraloria\Analitica Contraloria\Automatiaciones Ambiente Pruebas\Carpeta Miguel Cardona\Bot Actualizacion Maestras\input"
RUTA_SALIDA   = r"O:\Gerencia Contraloria\Analitica Contraloria\Automatiaciones Ambiente Pruebas\Carpeta Miguel Cardona\Bot Actualizacion Maestras\output"


def eliminar_espacios(texto):
    return texto.replace(" ", "")

def cerrar_instancias_excel():
    for proceso in psutil.process_iter(['name']):
        try:
            if proceso.info['name'] == 'EXCEL.EXE':
                proceso.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

# def inicio_y_filtrado_de_excel():
#     excel = None
#     try:
#         pythoncom.CoInitialize()
#         excel = win32com.client.Dispatch("Excel.Application")
#         excel.Visible = True
#         excel.ScreenUpdating = True
#         excel.DisplayAlerts = False
#         for archivo in os.listdir(RUTA_ENTRADA):            
#             maestra = os.path.join(RUTA_ENTRADA, archivo)

#             if not os.path.isfile(maestra):
#                 continue
#             nombre, extencion = os.path.splitext(archivo)
#             if archivo.startswith('~$') or extencion != '.xlsx':
#                 continue
#     except Exception as e:
#         print(f'❌ ERROR ----- {e}')
#     finally:
#         if excel is not None:
#             excel.Quit()
#         pythoncom.CoUninitialize()

def actualizar_maestras_bd():
    """Actualiza (RefreshAll) todas las maestras en RUTA_ENTRADA y las guarda."""
    excel = None
    try:
        cerrar_instancias_excel()
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        
        excel.Visible = True              # Recomendado: más rápido y limpio
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False

        print("🔄 Iniciando RefreshAll de todas las maestras en RUTA_ENTRADA...")

        for archivo in sorted(os.listdir(RUTA_ENTRADA)):   # sorted para orden consistente
            if not archivo.endswith('.xlsx') or archivo.startswith('~$'):
                continue

            ruta_maestra = os.path.join(RUTA_ENTRADA, archivo)
            nombre = os.path.splitext(archivo)[0]

            print(f"📊 Actualizando: {nombre}")

            workbook = excel.Workbooks.Open(ruta_maestra)
            workbook.RefreshAll()

            # Esperar a que termine de calcular
            while excel.CalculationState != 0:
                time.sleep(1)

            workbook.Save()
            workbook.Close(False)

            print(f"   ✅ {nombre} actualizado correctamente")

        print("🎉 Todas las maestras han sido actualizadas con RefreshAll.")

    except Exception as e:
        print(f'❌ ERROR en actualizar_maestras_bd ----- {e}')
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()



# def actualizar_maestras_excels(dfs):
#     excel = None
#     try:
#         cerrar_instancias_excel()
#         pythoncom.CoInitialize()
#         excel = win32com.client.Dispatch("Excel.Application")
#         excel.Visible = True
#         excel.ScreenUpdating = True
#         excel.DisplayAlerts = False
#         for archivo in os.listdir(RUTA_ENTRADA):            
#             maestra = os.path.join(RUTA_ENTRADA, archivo)

#             if not os.path.isfile(maestra):
#                 continue
#             nombre, extencion = os.path.splitext(archivo)
#             if archivo.startswith('~$') or extencion != '.xlsx':
#                 continue
            
#             workbook = excel.Workbooks.Open(maestra)
#             for nombre, df in dfs.items():
#                 nombre_df = nombre
#                 df_maestra = df
#                 # return nombre_df, df_maestra
            
#                 print(f"\n===== {nombre} =====")
#                 print(tabulate(df.head(), headers="keys", tablefmt="psql"))
#             hoja = workbook.Worksheets("ARCHIVO EXCEL")


#             while excel.CalculationState != 0:
#                 time.sleep(1)
#             workbook.Save()
#             workbook.Close(False)
#     except Exception as e:
#         print(f'❌ ERROR ----- {e}')
#     finally:
#         if excel is not None:
#             excel.Quit()
#         pythoncom.CoUninitialize()
def actualizar_maestras_excels():
    excel = None
    try:
        cerrar_instancias_excel()
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        
        excel.Visible = False
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False

        print("🔄 Iniciando actualización de archivos en RUTA_SALIDA...")

        # ==================== ARCHIVOS INDIVIDUALES ====================
        mapping = {
            "MAESTRA AMERICANINO": os.path.join(RUTA_SALIDA, "Maestra Americanino Marzo.xlsb"),
            "MAESTRA CHEVIGNON":  os.path.join(RUTA_SALIDA, "Maestra Chevignon Marzo.xlsb"),
            "MAESTRA ESPRIT":     os.path.join(RUTA_SALIDA, "Maestra Esprit Marzo.xlsb"),
            "MAESTRA NAF NAF":    os.path.join(RUTA_SALIDA, "Maestra Naf Naf Marzo.xlsb"),
        }

        # ==================== BRAND STORE (solo en RUTA_SALIDA) ====================
        brand_store_path = os.path.join(RUTA_SALIDA, "Maestra Brand Store Marzo.xlsx")

        # ==================== 1. ACTUALIZAR LOS 4 ARCHIVOS INDIVIDUALES ====================
        for nombre_origen, ruta_destino in mapping.items():
            ruta_origen = os.path.join(RUTA_ENTRADA, f"{nombre_origen}.xlsx")

            print(f"📤 Copiando: {nombre_origen}  →  {os.path.basename(ruta_destino)}")

            origen_wb = excel.Workbooks.Open(ruta_origen)
            origen_sheet = origen_wb.Worksheets("ARCHIVO EXCEL")
            used_range = origen_sheet.UsedRange

            dest_wb = excel.Workbooks.Open(ruta_destino)
            dest_sheet = dest_wb.Worksheets("MAESTRAS")
            dest_sheet.Cells.Clear()

            used_range.Copy(Destination=dest_sheet.Range("A1"))

            # Guardar con la extensión correcta
            file_format = 50 if ruta_destino.endswith('.xlsb') else 51
            dest_wb.SaveAs(Filename=ruta_destino, FileFormat=file_format)
            dest_wb.Close(False)
            origen_wb.Close(False)

            print(f"   ✅ Guardado correctamente")

        # ==================== 2. ACTUALIZAR BRAND STORE ====================
        print(f"\n📤 Actualizando Maestra Brand Store (5 bloques)...")

        brand_wb = excel.Workbooks.Open(brand_store_path)
        brand_sheet = brand_wb.Worksheets("MAESTRAS")
        brand_sheet.Cells.Clear()
        current_row = 1

        orden_brand_store = [
            "MAESTRA ESPRIT",
            "MAESTRA DISANDINA",
            "MAESTRA CHEVIGNON",
            "MAESTRA AMERICANINO",
            "MAESTRA ONEIL"
        ]

        for nombre_origen in orden_brand_store:
            ruta_origen = os.path.join(RUTA_ENTRADA, f"{nombre_origen}.xlsx")

            print(f"   → Pegando bloque: {nombre_origen}")

            origen_wb = excel.Workbooks.Open(ruta_origen)
            origen_sheet = origen_wb.Worksheets("ARCHIVO EXCEL")
            used_range = origen_sheet.UsedRange

            paste_range = brand_sheet.Cells(current_row, 1)
            used_range.Copy(Destination=paste_range)

            current_row += used_range.Rows.Count + 2   # 2 filas en blanco entre bloques

            origen_wb.Close(False)

        # Guardar Brand Store
        brand_wb.SaveAs(Filename=brand_store_path, FileFormat=51)
        brand_wb.Close(False)

        print("✅ Maestra Brand Store actualizada con los 5 bloques en orden")
        print("🎉 ¡Proceso completado exitosamente!")

    except Exception as e:
        print(f'❌ ERROR ----- {e}')
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()

def main():
    actualizar_maestras_bd()
    actualizar_maestras_excels()

if __name__ == '__main__':
    main()