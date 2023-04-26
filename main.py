from os import path, remove
from shutil import copyfile
import pandas as pd
import xlwings as xw
import numpy as np
from datetime import datetime

def main():
    fec_hoy = datetime.today()
    fecha_hoy_format = fec_hoy.strftime('%Y%m%d')
    ruta1 = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Entrada\PRUEBA_PLANTILLA_PRIORIZACIÓN.xlsx'
    ruta_principal = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Entrada'
    nombre_archivo = 'PRIORIZACIÓN_PO_{}.xlsx'.format(fecha_hoy_format)
    ruta_out_f = path.join(ruta_principal,nombre_archivo)

    lt_in_colaborador_curso = leer_excel_simple(ruta1, 'LT_IN_COLABORADOR_CURSO')
    lt_out_capacidad_enfoque = leer_excel_simple(ruta1, 'LT_OUT_CAPACIDAD_ENFOQUE')
    lt_out_compromiso = leer_excel_simple(ruta1, 'LT_OUT_COMPROMISO')
    lt_in_capacidad = leer_excel_simple(ruta1, 'LT_IN_CAPACIDAD')
    curso_priorizado = leer_excel_simple(ruta1, 'CURSO_PRIORIZADO')
    capacidad_enfoque = leer_excel_simple(ruta1, 'CAPACIDAD_ENFOQUE')
    
    #curso_priorizado
    lt_in_colaborador_curso.loc[lt_in_colaborador_curso['Columna1'] == '#N/D', 'Columna1'] = pd.NA
    lt_in_colaborador_curso.dropna(subset=['Columna1'], inplace=True)
    lt_in_colaborador_curso.reset_index(drop=True, inplace=True)

    lt_in_colaborador_curso.loc[lt_in_colaborador_curso['Columna2'] == '#N/D', 'Columna2'] = pd.NA
    lt_in_colaborador_curso.dropna(subset=['Columna2'], inplace=True)
    lt_in_colaborador_curso.reset_index(drop=True, inplace=True)


    #capacidad_enfoque
    lt_out_capacidad_enfoque.loc[lt_out_capacidad_enfoque['Columna1'] == '#N/D', 'Columna1'] = pd.NA
    lt_out_capacidad_enfoque.dropna(subset=['Columna1'], inplace=True)
    lt_out_capacidad_enfoque.reset_index(drop=True, inplace=True)

    df_2 = pd.merge(lt_out_capacidad_enfoque, lt_in_capacidad, how='left', on='ID_CAPACIDAD')
    #df_2 = capacidad_enfoque[['MATRICULA', 'CAPACIDAD']]
    print(df_2)

    #compromiso
    lt_out_compromiso.loc[lt_out_compromiso['Columna1'] == '#N/D', 'Columna1'] = pd.NA
    lt_out_compromiso.dropna(subset=['Columna1'], inplace=True)
    lt_out_compromiso.reset_index(drop=True, inplace=True)

    #excel
    copia_pega(ruta1, ruta_out_f)
    df_a_excel(ruta_out_f, 'CURSO_PRIORIZADO', lt_in_colaborador_curso[['MATRICULA']], f_ini = 2, c_ini = 2)
    df_a_excel(ruta_out_f, 'CURSO_PRIORIZADO', lt_in_colaborador_curso[['COD_CURSO']], f_ini = 2, c_ini = 6)

    df_a_excel(ruta_out_f, 'CAPACIDAD_ENFOQUE', df_2[['MATRICULA']],f_ini = 2, c_ini = 2)
    df_a_excel(ruta_out_f, 'CAPACIDAD_ENFOQUE', df_2[['CAPACIDAD']],f_ini = 2, c_ini = 6)

    df_a_excel(ruta_out_f, 'COMPROMISO', lt_out_compromiso[['MATRICULA']],f_ini = 2, c_ini = 2)
    df_a_excel(ruta_out_f, 'COMPROMISO', lt_out_compromiso[['N_COMPROMISO']],f_ini = 2, c_ini = 6)
    df_a_excel(ruta_out_f, 'COMPROMISO', lt_out_compromiso[['ACCION']],f_ini = 2, c_ini = 7)
    df_a_excel(ruta_out_f, 'COMPROMISO', lt_out_compromiso[['RECURSO']],f_ini = 2, c_ini = 8)
    df_a_excel(ruta_out_f, 'COMPROMISO', lt_out_compromiso[['FECHA_INI']],f_ini = 2, c_ini = 9)
    df_a_excel(ruta_out_f, 'COMPROMISO', lt_out_compromiso[['FECHA_FIN']],f_ini = 2, c_ini = 10)
    df_a_excel(ruta_out_f, 'COMPROMISO', lt_out_compromiso[['COMPROMISO']],f_ini = 2, c_ini = 11)


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

if __name__ == '__main__':
    main()