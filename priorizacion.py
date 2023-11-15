from os import path, remove
from shutil import copyfile
import time
import pandas as pd
import xlwings as xw
import numpy as np
from datetime import datetime
from base import copia_pega, df_a_excel, leer_excel_simple
import win32com.client
from constantes import PATH_BA

def priorizacion():
    chapter = input(
        '''
        PRIORIZACIÓN
        --------------
        Chapter:

        '''
    )
    fecha_filtro = '29/08/2023' #CAMBIAR EN BASE A LO NECESARIO, DEJAR EN BLANCO SI SE DESEA COGERTODO
    ruta_plantilla = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Entrada\PRUEBA_PLANTILLA_PRIORIZACIÓN.xlsx'
    xlapp = win32com.client.DispatchEx("Excel.Application")
    wb = xlapp.Workbooks.Open(ruta_plantilla)
    xlapp.Visible = True
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    wb.Save()
    xlapp.Quit()
    fec_hoy = datetime.today()
    fecha_hoy_format = fec_hoy.strftime('%Y%m%d')
    ruta_principal = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Salida'
    nombre_archivo = 'PRIORIZACIÓN_{}_{}.xlsx'.format(chapter, fecha_hoy_format)
    ruta_out_f = path.join(ruta_principal, nombre_archivo)
    ruta_ba = PATH_BA

    colaboradores = leer_excel_simple(ruta_plantilla, 'COLABORADORES')
    cursos = leer_excel_simple(ruta_plantilla, 'CURSOS')
    lt_in_colaborador_curso = leer_excel_simple(ruta_plantilla, 'LT_IN_COLABORADOR_CURSO')
    lt_out_capacidad_enfoque = leer_excel_simple(ruta_plantilla, 'LT_OUT_CAPACIDAD_ENFOQUE')
    lt_out_compromiso = leer_excel_simple(ruta_plantilla, 'LT_OUT_COMPROMISO')
    lt_in_capacidad = leer_excel_simple(ruta_plantilla, 'LT_IN_CAPACIDAD')
    curso_priorizado = leer_excel_simple(ruta_plantilla, 'CURSO_PRIORIZADO')
    capacidad_enfoque = leer_excel_simple(ruta_plantilla, 'CAPACIDAD_ENFOQUE')
    base_activos = leer_excel_simple(ruta_ba, 'BD ACTIVOS')
    base_activos.rename(columns={'Matrícula': 'MATRICULA'}, inplace=True)

    #filtros por fecha
    if fecha_filtro != '':
        lt_out_compromiso['Created'] = pd.to_datetime(lt_out_compromiso['Created'], format='%d/%m/%Y')
        lt_out_capacidad_enfoque['Created'] = pd.to_datetime(lt_out_capacidad_enfoque['Created'], format='%d/%m/%Y')
        lt_in_colaborador_curso['Created'] = pd.to_datetime(lt_in_colaborador_curso['Created'], format='%d/%m/%Y')

        fecha_filtro_format = pd.to_datetime(fecha_filtro, format='%d/%m/%Y')

        lt_out_compromiso = lt_out_compromiso[lt_out_compromiso['Created'] >= fecha_filtro_format]
        lt_out_capacidad_enfoque = lt_out_capacidad_enfoque[lt_out_capacidad_enfoque['Created'] >= fecha_filtro_format]
        lt_in_colaborador_curso = lt_in_colaborador_curso[lt_in_colaborador_curso['Created'] >= fecha_filtro_format]
    
    #curso_priorizado
    lt_in_colaborador_curso_filtrado = pd.merge(cursos[['COD_CURSO']], lt_in_colaborador_curso, how='left', on='COD_CURSO')
    lt_in_colaborador_curso_filtrado = lt_in_colaborador_curso_filtrado.drop_duplicates(subset=['MATRICULA','COD_CURSO', 'N_SESION', 'N_SESION_COMPLETADA', 'FECHA_INICIO', 'FECHA_FIN'])

    lt_in_colaborador_curso_filtrado = pd.merge(colaboradores[['MATRICULA']], lt_in_colaborador_curso_filtrado, how='left', on='MATRICULA')
    lt_in_colaborador_curso_filtrado = lt_in_colaborador_curso_filtrado.drop_duplicates(subset=['MATRICULA', 'COD_CURSO', 'N_SESION', 'N_SESION_COMPLETADA', 'FECHA_INICIO', 'FECHA_FIN'])
    lt_in_colaborador_curso_filtrado = lt_in_colaborador_curso_filtrado.dropna(subset=['COD_CURSO'])

    #capacidad_enfoque
    lt_out_capacidad_enfoque_filtrado = pd.merge(colaboradores[['MATRICULA']], lt_out_capacidad_enfoque, how='left', on='MATRICULA')
    lt_out_capacidad_enfoque_filtrado = lt_out_capacidad_enfoque_filtrado.drop_duplicates(subset=['MATRICULA', 'ID_CAPACIDAD'])
    lt_out_capacidad_enfoque_filtrado = lt_out_capacidad_enfoque_filtrado.dropna(subset=['ID_CAPACIDAD'])

    df_2 = pd.merge(lt_out_capacidad_enfoque_filtrado, lt_in_capacidad, how='left', on='ID_CAPACIDAD')

    #compromiso
    lt_out_compromiso_filtrado = pd.merge(colaboradores[['MATRICULA']], lt_out_compromiso, how='left', on='MATRICULA')
    #lt_out_compromiso_filtrado = lt_out_compromiso_filtrado.drop_duplicates(subset=['MATRICULA', 'N_COMPROMISO', 'ACCION', 'RECURSO', 'FECHA_INI', 'FECHA_FIN', 'COMPROMISO'])
    lt_out_compromiso_filtrado = lt_out_compromiso_filtrado.dropna(subset=['COMPROMISO'])

    lt_out_compromiso_filtrado['FECHA_INI'] = [x.strftime("%d/%m/%Y") for x in lt_out_compromiso_filtrado['FECHA_INI']]
    lt_out_compromiso_filtrado['FECHA_FIN'] = [x.strftime("%d/%m/%Y") for x in lt_out_compromiso_filtrado['FECHA_FIN']]

    #actualizar FLAG_PRIORIZACIÓN
    df_3 = pd.merge(lt_in_colaborador_curso_filtrado[['MATRICULA']], colaboradores, how='left', on='MATRICULA')
    df_3 = df_3.drop_duplicates(subset=['MATRICULA', 'NOMBRE', 'ROL', 'ESTADO', 'MATRICULA_CALIFICADOR', 'NOMBRE_CALIFICADOR', 'ROL_CALIFICADOR', 'CHAPTER', 'FLAG_PRIORIZACIÓN', 'FLAG_EXCLUSIÓN', 'MOTIVO_EXCLUSIÓN'])

    for i, matricula in enumerate(colaboradores['MATRICULA']):
        if matricula in df_3['MATRICULA'].values:
            colaboradores.loc[i, 'FLAG_PRIORIZACIÓN'] = 'SI'

    #actualizar FLAG_EXCLUSIÓN
    df_4 = pd.merge(base_activos[['MATRICULA']], colaboradores, how='left', on='MATRICULA')
    df_4 = df_4.drop_duplicates(subset=['MATRICULA', 'NOMBRE', 'ROL', 'ESTADO', 'MATRICULA_CALIFICADOR', 'NOMBRE_CALIFICADOR', 'ROL_CALIFICADOR', 'CHAPTER', 'FLAG_PRIORIZACIÓN', 'FLAG_EXCLUSIÓN', 'MOTIVO_EXCLUSIÓN'])

    for i, matricula in enumerate(colaboradores['MATRICULA']):
        if matricula in df_4['MATRICULA'].values:
            colaboradores.loc[i, 'ESTADO'] = 'ACTIVO'
            colaboradores.loc[i, 'FLAG_EXCLUSIÓN'] = 'NO'
        else:
            colaboradores.loc[i, 'ESTADO'] = 'INACTIVO'
            colaboradores.loc[i, 'FLAG_EXCLUSIÓN'] = 'SI'
    
    #crear excel final
    copia_pega(ruta_plantilla, ruta_out_f)
    df_a_excel(ruta_out_f, 'CURSO_PRIORIZADO', lt_in_colaborador_curso_filtrado[['MATRICULA']], f_ini = 2, c_ini = 2)
    df_a_excel(ruta_out_f, 'CURSO_PRIORIZADO', lt_in_colaborador_curso_filtrado[['COD_CURSO']], f_ini = 2, c_ini = 6)

    df_a_excel(ruta_out_f, 'CAPACIDAD_ENFOQUE', df_2[['MATRICULA']],f_ini = 2, c_ini = 2)
    df_a_excel(ruta_out_f, 'CAPACIDAD_ENFOQUE', df_2[['CAPACIDAD']],f_ini = 2, c_ini = 6)

    df_a_excel(ruta_out_f, 'COMPROMISO', lt_out_compromiso_filtrado[['MATRICULA']],f_ini = 2, c_ini = 2)
    df_a_excel(ruta_out_f, 'COMPROMISO', lt_out_compromiso_filtrado[['N_COMPROMISO']],f_ini = 2, c_ini = 6)
    df_a_excel(ruta_out_f, 'COMPROMISO', lt_out_compromiso_filtrado[['ACCION']],f_ini = 2, c_ini = 7)
    df_a_excel(ruta_out_f, 'COMPROMISO', lt_out_compromiso_filtrado[['RECURSO']],f_ini = 2, c_ini = 8)
    df_a_excel(ruta_out_f, 'COMPROMISO', lt_out_compromiso_filtrado[['FECHA_INI']],f_ini = 2, c_ini = 9)
    df_a_excel(ruta_out_f, 'COMPROMISO', lt_out_compromiso_filtrado[['FECHA_FIN']],f_ini = 2, c_ini = 10)
    df_a_excel(ruta_out_f, 'COMPROMISO', lt_out_compromiso_filtrado[['COMPROMISO']],f_ini = 2, c_ini = 11)

    df_a_excel(ruta_out_f, 'COLABORADORES', colaboradores, f_ini = 2, c_ini = 1)

    xlapp = win32com.client.DispatchEx("Excel.Application")
    wb = xlapp.Workbooks.Open(ruta_out_f)
    xlapp.Visible = True
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    wb.Save()
    xlapp.Quit()

if __name__ == '__main__':
    priorizacion()