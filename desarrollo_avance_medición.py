from datetime import datetime
from os import path
import pandas as pd
import win32com
from base import copia_pega, df_a_excel, leer_excel_simple
from enumeraciones import ETipoEva
from constantes import PATH_BA

def des_avance_medicion():
    chapter = 'DESARROLLO_NOVIEMBRE'
    ruta_plantilla = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Entrada\PLANTILLA_AVANCE_MEDICION_NOVIEMBRE.xlsx'
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
    nombre_archivo = '0AVANCE_MEDICION_{}_{}.xlsx'.format(chapter, fecha_hoy_format)
    ruta_out_f = path.join(ruta_principal, nombre_archivo)
    ruta_ba = PATH_BA

    #base = leer_excel_simple(ruta_plantilla, 'BASE')
    lt_out_comportamiento = leer_excel_simple(ruta_plantilla, 'LT_OUT_COMPORTAMIENTO')
    lt_out_capacidad = leer_excel_simple(ruta_plantilla, 'LT_OUT_CAPACIDAD')
    base_activos = leer_excel_simple(ruta_ba, 'BD ACTIVOS')
    base_activos.rename(columns={'Matrícula': 'MATRICULA'}, inplace=True)
    
    lt_out_capacidad.rename(columns={'ID_CAPACIDAD':'ID_COMPORTAMIENTO'}, inplace=True)
    columnas_concatenadas = ['MATRICULA_CALIFICADOR', 'MATRICULA_CALIFICADO', 'ID_COMPORTAMIENTO']

    df_concatenado = pd.concat([lt_out_comportamiento[columnas_concatenadas], lt_out_capacidad[columnas_concatenadas]], ignore_index=True)

    copia_pega(ruta_plantilla, ruta_out_f)
    df_a_excel(ruta_out_f, 'BASE', df_concatenado, f_ini = 2)

if __name__ == '__main__':
    des_avance_medicion()