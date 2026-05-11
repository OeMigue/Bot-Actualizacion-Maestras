import sys, pythoncom, time, os, psutil
import pandas as pd
import glob
from datetime import datetime
import win32com.client.gencache

sys.path.append(r"O:\Gerencia Contraloria\Analitica Contraloria\Automatiaciones Ambiente Pruebas\Carpeta Miguel Cardona\Bot Actualizacion Maestras\config")
sys.path.append(r"O:\Gerencia Contraloria\Analitica Contraloria\Automatiaciones Ambiente Pruebas\Carpeta Miguel Cardona\Bot Actualizacion Maestras\src")
sys.path.append(r"O:\Gerencia Contraloria\Analitica Contraloria\Automatiaciones Ambiente Pruebas\Carpeta Miguel Cardona\Funciones")
sys.stdout.reconfigure(encoding="utf8")

RUTA_ENTRADA  = r"O:\Gerencia Contraloria\Analitica Contraloria\Tiendas Franquicias\MAESTRAS BASE DE DATOS"
RUTA_SALIDA   = r"O:\Gerencia Contraloria\Analitica Contraloria\Tiendas Franquicias\MAESTRAS\MAESTRAS EXCEL"

def eliminar_espacios(texto):
    return texto.replace(" ", "")

def cerrar_instancias_excel():
    for proceso in psutil.process_iter(['name']):
        try:
            if proceso.info['name'] == 'EXCEL.EXE':
                proceso.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

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
            # cerrar_instancias_excel()

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

MESES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

def obtener_mes_siguiente():
    mes_actual = datetime.now().month
    mes_siguiente = (mes_actual % 12) + 1
    return MESES[mes_siguiente]

def buscar_archivo_existente(patron):
    resultados = glob.glob(os.path.join(RUTA_SALIDA, patron))
    if resultados:
        return resultados[0]
    return None

def actualizar_maestras_excels():
    excel = None
    try:
        cerrar_instancias_excel()
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")

        excel.Visible = False
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False

        mes_nuevo = obtener_mes_siguiente()
        print(f"📅 Mes destino detectado: {mes_nuevo}")
        print("🔄 Iniciando actualización de archivos en RUTA_SALIDA...")

        # ==================== MAPPING AUTOMÁTICO ====================
        mapping = {
            "MAESTRA AMERICANINO": {
                "archivo_actual": buscar_archivo_existente("Maestra Americanino *.xlsb"),
                "archivo_nuevo":  os.path.join(RUTA_SALIDA, f"Maestra Americanino {mes_nuevo}.xlsb"),
            },
            "MAESTRA CHEVIGNON": {
                "archivo_actual": buscar_archivo_existente("Maestra Chevignon *.xlsb"),
                "archivo_nuevo":  os.path.join(RUTA_SALIDA, f"Maestra Chevignon {mes_nuevo}.xlsb"),
            },
            "MAESTRA ESPRIT": {
                "archivo_actual": buscar_archivo_existente("Maestra Esprit *.xlsb"),
                "archivo_nuevo":  os.path.join(RUTA_SALIDA, f"Maestra Esprit {mes_nuevo}.xlsb"),
            },
            "MAESTRA NAF NAF": {
                "archivo_actual": buscar_archivo_existente("Maestra Naf Naf *.xlsb"),
                "archivo_nuevo":  os.path.join(RUTA_SALIDA, f"Maestra Naf Naf {mes_nuevo}.xlsb"),
            },
        }

        # ==================== 1. ARCHIVOS INDIVIDUALES ====================
        for nombre_origen, rutas in mapping.items():
            ruta_actual = rutas["archivo_actual"]
            ruta_nueva  = rutas["archivo_nuevo"]

            if ruta_actual is None:
                print(f"⚠️  No se encontró archivo existente para {nombre_origen}, se omite.")
                continue

            ruta_origen = os.path.join(RUTA_ENTRADA, f"{nombre_origen}.xlsx")

            print(f"📤 Copiando: {nombre_origen}  →  {os.path.basename(ruta_nueva)}")

            print(f"   → [1] Abriendo origen: {ruta_origen}")
            origen_wb = excel.Workbooks.Open(ruta_origen)
            print(f"   → [2] Origen abierto OK")

            origen_sheet = origen_wb.Worksheets("ARCHIVO EXCEL")
            print(f"   → [3] Hoja origen OK")

            used_range = origen_sheet.UsedRange
            print(f"   → [4] UsedRange OK")

            print(f"   → [5] Abriendo destino: {ruta_actual}")
            dest_wb = excel.Workbooks.Open(ruta_actual)
            print(f"   → [6] Destino abierto OK")

            dest_sheet = dest_wb.Worksheets("Maestra")
            print(f"   → [7] Hoja destino OK")

            dest_sheet.Cells.Clear()
            print(f"   → [8] Clear OK")

            used_range.Copy(Destination=dest_sheet.Range("A1"))
            print(f"   → [9] Copy OK")

            file_format = 50 if ruta_nueva.endswith('.xlsb') else 51

            if ruta_actual == ruta_nueva:
                dest_wb.Save()
                print(f"   ✅ Guardado correctamente (mismo nombre)")
            else:
                print(f"   → [10] Ejecutando SaveAs...")
                dest_wb.SaveAs(Filename=ruta_nueva, FileFormat=file_format)
                dest_wb.Close(False)
                origen_wb.Close(False)
                os.remove(ruta_actual)
                print(f"   ✅ Guardado como {os.path.basename(ruta_nueva)} y eliminado {os.path.basename(ruta_actual)}")
                continue

            dest_wb.Close(False)
            origen_wb.Close(False)

        # ==================== 2. BRAND STORE ====================
        print(f"\n📤 Actualizando Maestra Brand Store...")

        brand_actual = buscar_archivo_existente("Maestra Brand Store *.xlsb")
        print(f"   → [BS-1] Archivo encontrado: {brand_actual}")

        brand_nueva = os.path.join(RUTA_SALIDA, f"Maestra Brand Store {mes_nuevo}.xlsb")

        if brand_actual is None:
            print("⚠️  No se encontró el archivo Maestra Brand Store, se omite.")
        else:
            print(f"   → [BS-2] Abriendo Brand Store...")
            brand_wb    = excel.Workbooks.Open(brand_actual)
            print(f"   → [BS-3] Abierto OK")

            brand_sheet = brand_wb.Worksheets("MAESTRA")
            print(f"   → [BS-4] Hoja OK")

            brand_sheet.Cells.Clear()
            print(f"   → [BS-5] Clear OK")
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
                print(f"   → [BS-6] Abriendo bloque: {nombre_origen}")

                origen_wb    = excel.Workbooks.Open(ruta_origen)
                print(f"   → [BS-7] Abierto OK")

                origen_sheet = origen_wb.Worksheets("ARCHIVO EXCEL")
                print(f"   → [BS-8] Hoja origen OK")

                used_range   = origen_sheet.UsedRange
                print(f"   → [BS-9] UsedRange OK")

                paste_range = brand_sheet.Cells(current_row, 1)
                used_range.Copy(Destination=paste_range)
                print(f"   → [BS-10] Copiado OK, fila actual: {current_row}")

                current_row += used_range.Rows.Count + 2
                origen_wb.Close(False)

            if brand_actual == brand_nueva:
                brand_wb.Save()
                print(f"   ✅ Brand Store guardado (mismo nombre)")
            else:
                print(f"   → [BS-11] Ejecutando SaveAs...")
                brand_wb.SaveAs(Filename=brand_nueva, FileFormat=50)
                brand_wb.Close(False)
                os.remove(brand_actual)
                print(f"   ✅ Brand Store guardado como {os.path.basename(brand_nueva)} y eliminado {os.path.basename(brand_actual)}")

        print("\n🎉 ¡Proceso completado exitosamente!")

    except Exception as e:
        print(f'❌ ERROR ----- {e}')
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
    
DESTINATARIOS = [
    "aprendizprogramacion2@gco.com.co"
]

ASUNTO_CORREO = "Maestras actualizadas - Junio 2026"

CUERPO_CORREO = """
Buen día,

Se adjuntan las maestras actualizadas correspondientes al mes de Junio.

Saludos,
Miguel Cardona
"""

def enviar_maestras_correo():
    outlook = None
    try:
        print("\n📧 Iniciando envío de correo con maestras adjuntas...")

        # Buscar todos los archivos en RUTA_SALIDA que sean maestras
        adjuntos = []
        for archivo in os.listdir(RUTA_SALIDA):
            ruta_archivo = os.path.join(RUTA_SALIDA, archivo)
            # Solo archivos Excel, ignorar carpetas y temporales
            if os.path.isfile(ruta_archivo) and not archivo.startswith('~$'):
                if archivo.endswith('.xlsx') or archivo.endswith('.xlsb'):
                    adjuntos.append(ruta_archivo)

        if not adjuntos:
            print("⚠️  No se encontraron archivos para adjuntar en RUTA_SALIDA.")
            return

        print(f"   → {len(adjuntos)} archivo(s) encontrado(s) para adjuntar")

        # Iniciar Outlook
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        correo  = outlook.CreateItem(0)  # 0 = MailItem

        # Destinatarios
        for destinatario in DESTINATARIOS:
            correo.Recipients.Add(destinatario)

        # Asunto y cuerpo
        correo.Subject = ASUNTO_CORREO
        correo.Body    = CUERPO_CORREO

        # Adjuntar archivos
        for ruta_adjunto in adjuntos:
            correo.Attachments.Add(ruta_adjunto)
            print(f"   → Adjuntando: {os.path.basename(ruta_adjunto)}")

        # Enviar
        correo.Send()
        print("✅ Correo enviado exitosamente!")
        print(f"   → Destinatarios: {', '.join(DESTINATARIOS)}")

    except Exception as e:
        print(f"❌ ERROR en enviar_maestras_correo ----- {e}")
    finally:
        if outlook is not None:
            outlook.Quit()
        pythoncom.CoUninitialize()

def main():
    actualizar_maestras_bd()
    actualizar_maestras_excels()
    enviar_maestras_correo()

if __name__ == '__main__':
    main()