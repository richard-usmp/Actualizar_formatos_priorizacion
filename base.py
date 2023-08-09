from calendar import monthrange
from datetime import date
from os import path, remove
import os
from shutil import copyfile
import pandas as pd
import xlwings as xw
from xlwings.constants import DeleteShiftDirection
import numpy as np

PATH_BA = r'\\130.1.22.103\P&A-Interno\08. People Analytics\12. Iniciativas\2. Base de activos\4. OUTPUTS'

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

def df_a_excel_header(ruta, nom_hoja, df, f_ini = 1, c_ini = 1):

    # Abriendo la instancia de Excel
    app = xw.App(visible=False)
    app.display_alerts = False

    # Abriendo el libro
    wb = app.books.open(ruta)
    ws = wb.sheets(nom_hoja)
    
    # Pegando la información
    ws.range((f_ini,c_ini)).options(index=False, header = True).value = df

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

def leer_ba():
    d_today = date.today()

    t_today = get_fecha_activos(d_today.year,d_today.month,d_today.day)

    panio = t_today[0]
    pmes = t_today[1]
    pdia = t_today[2]-1

    ruta = os.path.join(PATH_BA,completa_ruta(panio,pmes,pdia),'Base de Activos GDH.xlsx')
    print(ruta)
    df_activos = leer_excel_simple(ruta,hoja='BD ACTIVOS')
    df_activos = df_activos[df_activos['Tipo_Preper'].isin(['Proveedor','Orgánico'])]
    df_activos = df_activos[['Correo electronico','Matrícula']]

    # Renombramos columnas
    dic_columns = {
        df_activos.columns[0]: 'CORREO',
        df_activos.columns[1]: 'MATRICULA'
    }
    df_activos.rename(columns=dic_columns, inplace=True)

    # Homologamos Correo y Matricula
    df_activos['CORREO'] = [ str(x).lower().strip() for x in df_activos['CORREO']]
    df_activos['MATRICULA'] = [ str(x)[1:] for x in df_activos['MATRICULA']]

    return df_activos

def completa_ruta(anio,mes,dia):
    dic_meses = {
        1 : '1. ENE',
        2 : '2. FEB',
        3 : '3. MAR',
        4 : '4. ABR',
        5 : '5. MAY',
        6 : '6. JUN',
        7 : '7. JUL',
        8 : '8. AGO',
        9 : '9. SEP',
        10 : '10. OCT',
        11 : '11. NOV',
        12 : '12. DIC'
    }
    folder_name_mes = dic_meses.get(mes)
    folder_name_dia = get_folder_dia(dia)
    
    ruta = '{anio_f}\{mes_f}\{dia_f}'.format(anio_f = anio, mes_f = folder_name_mes,dia_f = folder_name_dia)
    return ruta

def get_fecha_activos(anio,mes,dia):
    s_dia = get_folder_dia(dia)

    if s_dia == '31':
        mr = monthrange(anio,mes)
        dia = mr[1]
    elif s_dia == '06':
        dia = 7
    else:
        dia = int(dia)

    return (anio,mes,dia)

def get_folder_dia(dia):
    if dia >= 1 and dia < 7 : 
        folder_name_dia = '01'
    elif dia >= 7 and dia < 20:
        folder_name_dia = '06'
    elif dia >=20 and dia <27:
        folder_name_dia = '20'
    else:
        folder_name_dia = '31'
    
    return folder_name_dia

def elimina_col_excel_res_lid(ruta, nom_hoja, cant_capa):

    # Abriendo la instancia de Excel
    app = xw.App(visible=False)
    app.display_alerts = False

    # Abriendo el libro
    wb = app.books.open(ruta)
    ws = wb.sheets(nom_hoja)

    # Eliminando registros
    col_ini = 4 + cant_capa
    col_final = col_ini + (19 - cant_capa - 1)
    letra_col_ini = numero_a_letra_excel(col_ini)
    letra_col_final = numero_a_letra_excel(col_final)
    ws.range('{}:{}'.format(letra_col_ini, letra_col_final)).api.Delete(DeleteShiftDirection.xlShiftToLeft)

    # Guardando y cerrando el archivo
    wb.save()
    wb.close()
    app.kill()

def elimina_filas_excel_res_lid(ruta, nom_hoja):

    # Abriendo la instancia de Excel
    app = xw.App(visible=False)
    app.display_alerts = False

    # Abriendo el libro
    wb = app.books.open(ruta)
    ws = wb.sheets(nom_hoja)
    
    # Eliminando registros
    fila_ini = lastRow(ws, col=2) + 1
    if lastRow(ws, col=2) > 2: ws.range('{}:{}'.format(fila_ini, 1000)).api.Delete(DeleteShiftDirection.xlShiftUp) 

    # Guardando y cerrando el archivo
    wb.save()
    wb.close()
    app.kill()

def elimina_col_excel(ruta, nom_hoja, cant_capa):

    # Abriendo la instancia de Excel
    app = xw.App(visible=False)
    app.display_alerts = False

    # Abriendo el libro
    wb = app.books.open(ruta)
    ws = wb.sheets(nom_hoja)

    # Eliminando registros
    col_ini = 9 + cant_capa + 2
    col_final = col_ini + (19 - cant_capa - 1)
    letra_col_ini = numero_a_letra_excel(col_ini)
    letra_col_final = numero_a_letra_excel(col_final)
    ws.range('{}:{}'.format(letra_col_ini, letra_col_final)).api.Delete(DeleteShiftDirection.xlShiftToLeft)

    # Guardando y cerrando el archivo
    wb.save()
    wb.close()
    app.kill()

def elimina_filas_excel(ruta, nom_hoja):

    # Abriendo la instancia de Excel
    app = xw.App(visible=False)
    app.display_alerts = False

    # Abriendo el libro
    wb = app.books.open(ruta)
    ws = wb.sheets(nom_hoja)
    
    # Eliminando registros
    fila_ini = lastRow(ws, col=3) + 1
    if lastRow(ws, col=3) > 2: ws.range('{}:{}'.format(fila_ini, 1000)).api.Delete(DeleteShiftDirection.xlShiftUp) 

    # Guardando y cerrando el archivo
    wb.save()
    wb.close()
    app.kill()

def numero_a_letra_excel(numero):
    letras = []
    while numero > 0:
        numero -= 1
        letras.append(chr(ord('A') + numero % 26))
        numero //= 26
    letras.reverse()
    letra_excel = ''.join(letras)
    return letra_excel

def getColumnName(n):
    # initialize output string as empty
    result = ''
    while n > 0:
        # find the index of the next letter and concatenate the letter
        # to the solution
        # here index 0 corresponds to 'A', and 25 corresponds to 'Z'
        index = (n - 1) % 26
        result += chr(index + ord('A'))
        n = (n - 1) // 26
    return result[::-1]