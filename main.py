from os import path, remove
from shutil import copyfile
import pandas as pd
import xlwings as xw
import numpy as np
from datetime import datetime
from base import copia_pega, df_a_excel, leer_excel_simple

def main():
    fec_hoy = datetime.today()
    fecha_hoy_format = fec_hoy.strftime('%Y%m%d')
    ruta1 = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Entrada\PRUEBA_PLANTILLA_PRIORIZACIÓN.xlsx'
    ruta_principal = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Salida'
    nombre_archivo = 'PRIORIZACIÓN_PO_{}.xlsx'.format(fecha_hoy_format)
    ruta_out_f = path.join(ruta_principal, nombre_archivo)

    colaboradores = leer_excel_simple(ruta1, 'COLABORADORES')
    cursos = leer_excel_simple(ruta1, 'CURSOS')
    lt_in_colaborador_curso = leer_excel_simple(ruta1, 'LT_IN_COLABORADOR_CURSO')
    lt_out_capacidad_enfoque = leer_excel_simple(ruta1, 'LT_OUT_CAPACIDAD_ENFOQUE')
    lt_out_compromiso = leer_excel_simple(ruta1, 'LT_OUT_COMPROMISO')
    lt_in_capacidad = leer_excel_simple(ruta1, 'LT_IN_CAPACIDAD')
    curso_priorizado = leer_excel_simple(ruta1, 'CURSO_PRIORIZADO')
    capacidad_enfoque = leer_excel_simple(ruta1, 'CAPACIDAD_ENFOQUE')
    
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
    lt_out_compromiso_filtrado = lt_out_compromiso_filtrado.drop_duplicates(subset=['MATRICULA', 'N_COMPROMISO', 'ACCION', 'RECURSO', 'FECHA_INI', 'FECHA_FIN', 'COMPROMISO'])
    lt_out_compromiso_filtrado = lt_out_compromiso_filtrado.dropna(subset=['N_COMPROMISO'])

    lt_out_compromiso_filtrado['FECHA_INI'] = [x.strftime("%d/%m/%Y") for x in lt_out_compromiso_filtrado['FECHA_INI']]
    lt_out_compromiso_filtrado['FECHA_FIN'] = [x.strftime("%d/%m/%Y") for x in lt_out_compromiso_filtrado['FECHA_FIN']]

    #actualizar FLAG_PRIORIZACIÓN
    df_3 = pd.merge(lt_in_colaborador_curso_filtrado[['MATRICULA']], colaboradores, how='left', on='MATRICULA')
    df_3 = df_3.drop_duplicates(subset=['MATRICULA', 'NOMBRE', 'ROL', 'ESTADO', 'MATRICULA_CALIFICADOR', 'NOMBRE_CALIFICADOR', 'ROL_CALIFICADOR', 'CHAPTER', 'FLAG_PRIORIZACIÓN', 'FLAG_EXCLUSIÓN', 'MOTIVO_EXCLUSIÓN'])

    for i, matricula in enumerate(colaboradores['MATRICULA']):
        if matricula in df_3['MATRICULA'].values:
            colaboradores.loc[i, 'FLAG_PRIORIZACIÓN'] = 'SI'
    
    #crear excel final
    copia_pega(ruta1, ruta_out_f)
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



if __name__ == '__main__':
    main()