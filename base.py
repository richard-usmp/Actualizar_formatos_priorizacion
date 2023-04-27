from os import path, remove
from shutil import copyfile
import pandas as pd
import xlwings as xw
import numpy as np

def lastRow(ws, col=1):
    lwr_r_cell = ws.cells.last_cell
    lwr_row = lwr_r_cell.row
    lwr_cell = ws.range((lwr_row, col))

    if lwr_cell.value is None:
        lwr_cell = lwr_cell.end('up')

    return lwr_cell.row

def lastColumn(ws, row=1):
    lwr_r_cell = ws.cells.last_cell
    lwr_col = lwr_r_cell.column
    lwr_cell = ws.range((row, lwr_col))

    if lwr_cell.value is None:
        lwr_cell = lwr_cell.end('left')

    return lwr_cell.column

def leer_excel_simple(ruta,hoja=None,f_inicio=1, c_inicio=1,is_encuesta=False):
    header = 1

    app = xw.App(visible= False)
    app.display_alerts = False
    wb_api = app.books.api.Open(ruta, UpdateLinks=False, ReadOnly=True)
    wb = xw.Book(impl=xw._xlwindows.Book(xl=wb_api))
    
    ws = wb.sheets[0] if hoja is None else wb.sheets(hoja)
    # Obteneiendo rangos
    lr = lastRow(ws,c_inicio)
    lc = lastColumn(ws,f_inicio)

    # Caso encuesta
    if is_encuesta:
        header = 2 

    df = ws.range((f_inicio,c_inicio),(lr,lc)).options(pd.DataFrame, index=False,empty=np.nan, header=header).value

    wb.close()
    app.kill()

    return df

def df_a_excel(ruta, nom_hoja, df, f_ini = 1, c_ini = 1):

    # Abriendo la instancia de Excel
    app = xw.App(visible=False)
    app.display_alerts = False

    # Abriendo el libro
    wb = app.books.open(ruta)
    ws = wb.sheets(nom_hoja)
    
    # Pegando la información
    ws.range((f_ini,c_ini)).options(index=False, header = False).value = df

    # Guardando y cerrando el archivo
    wb.save()
    wb.close()
    app.kill()

def copia_pega(ruta_origen, ruta_destino):
# Limpiando el archivo anterior de la carpeta en caso exista
    try:
        remove(ruta_destino)
        print('\nSe removió archivo anterior')
    except:
        print('\nNo se encontró archivo anterior')
        # Copiando los formatos a la carpeta output
        copyfile(ruta_origen,ruta_destino)