import sys
sys.stdout.reconfigure(encoding='utf-8')

def cerrar_instancias_excel():
    import psutil
    for proceso in psutil.process_iter(['name']):
        try:
            if proceso.info['name'] == 'EXCEL.EXE':
                proceso.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

def actualizar_maestras_bd(ruta_fuente):
    import xlwings as xw
    import time
    import os
    try:
        for archivo in os.listdir(ruta_fuente):
            ruta_archivo = os.path.join(ruta_fuente, archivo)
            wb = xw.Book(ruta_archivo, visible=False)
            wb.api.RefreshAll()
            time.sleep(3)
            wb.save()
            wb.close()
    except Exception as e:
        print(f'Error al actualizar {ruta_fuente}: {e}')
    finally:
        wb.close()
    print(f'Archivo ')