from datetime import datetime
from os import path
import pandas as pd
import win32com
from base import copia_pega, df_a_excel, leer_excel_simple
from enumeraciones import ETipoEva
from constantes import PATH_BA

def avance_medición():
    chapter = input(
        '''
        AVANCE DE MEDICIÓN
        --------------------
        Chapter:
        
        '''
    )
    ruta_plantilla = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Entrada\PLANTILLA_AVANCE_MEDICION.xlsx'
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
    nombre_archivo = 'AVANCE_MEDICION_{}_{}.xlsx'.format(chapter, fecha_hoy_format)
    ruta_out_f = path.join(ruta_principal, nombre_archivo)
    ruta_ba = PATH_BA

    base = leer_excel_simple(ruta_plantilla, 'BASE')
    lt_out_comportamiento = leer_excel_simple(ruta_plantilla, 'LT_OUT_COMPORTAMIENTO')
    lt_out_capacidad = leer_excel_simple(ruta_plantilla, 'LT_OUT_CAPACIDAD')
    base_activos = leer_excel_simple(ruta_ba, 'BD ACTIVOS')
    base_activos.rename(columns={'Matrícula': 'MATRICULA'}, inplace=True)

    lt_out_comportamiento = lt_out_comportamiento.drop_duplicates(subset=['MATRICULA_CALIFICADOR', 'MATRICULA_CALIFICADO', 'CHAPTER_CALIFICADO'])
    lt_out_capacidad = lt_out_capacidad.drop_duplicates(subset=['MATRICULA_CALIFICADOR', 'MATRICULA_CALIFICADO', 'CHAPTER_CALIFICADO'])

    x_menu = input('''
          MENÚ
          ----------------------
          1. FLAG_EVALUACION y PO
          2. FLAG_AUTOEVALUACION y PO
          3. FLAG_EVALUACION y otros
          4. FLAG_AUTOEVALUACION y otros
          5. FLAG_EVALUACION, FLAG_AUTOEVALUACION y otros

          ELIJA UNA OPCIÓN: 
          ''')
    dic = {
        '1':ETipoEva.FLAG_EVALUACION_y_PO,
        '2':ETipoEva.FLAG_AUTOEVALUACION_y_PO,
        '3':ETipoEva.FLAG_EVALUACION_y_otros,
        '4':ETipoEva.FLAG_AUTOEVALUACION_y_otros,
        '5':ETipoEva.AUTOEVALUACION_Y_EVALUACION_y_otros
    }
    en = dic.get(x_menu)

    if en is not None:
        if en == ETipoEva.FLAG_EVALUACION_y_PO:
            
            base['CONCAT'] = base.MATRICULA_CALIFICADOR.str.cat(base.MATRICULA_CALIFICADO.str.cat(base.CHAPTER_CALIFICADO, sep=''), sep='')
            lt_out_capacidad['CONCAT'] = lt_out_capacidad.MATRICULA_CALIFICADOR.str.cat(lt_out_capacidad.MATRICULA_CALIFICADO.str.cat(lt_out_capacidad.CHAPTER_CALIFICADO, sep=''), sep='')

            df_1 = pd.merge(lt_out_capacidad[['CONCAT']], base, how='left', on='CONCAT')
            df_1 = df_1.dropna(subset=['MATRICULA_CALIFICADOR'])

            for i, matricula in enumerate(base['CONCAT']):
                if matricula in df_1['CONCAT'].values:
                    base.loc[i, 'FLAG_EVALUACION'] = 'SI'
        
        elif en == ETipoEva.FLAG_AUTOEVALUACION_y_PO:

            base['CONCAT'] = base.MATRICULA_CALIFICADO.str.cat(base.MATRICULA_CALIFICADO.str.cat(base.CHAPTER_CALIFICADO, sep=''), sep='')
            lt_out_capacidad['CONCAT'] = lt_out_capacidad.MATRICULA_CALIFICADOR.str.cat(lt_out_capacidad.MATRICULA_CALIFICADO.str.cat(lt_out_capacidad.CHAPTER_CALIFICADO, sep=''), sep='')
            
            df_2 = pd.merge(lt_out_capacidad[['CONCAT']], base, how='left', on='CONCAT')
            df_2 = df_2.dropna(subset=['MATRICULA_CALIFICADOR'])

            for i, matricula in enumerate(base['CONCAT']):
                if matricula in df_2['CONCAT'].values:
                    base.loc[i, 'FLAG_AUTOEVALUACION'] = 'SI'

        elif en == ETipoEva.FLAG_EVALUACION_y_otros:

            base['CONCAT'] = base.MATRICULA_CALIFICADOR.str.cat(base.MATRICULA_CALIFICADO.str.cat(base.CHAPTER_CALIFICADO, sep=''), sep='')
            lt_out_comportamiento['CONCAT'] = lt_out_comportamiento.MATRICULA_CALIFICADOR.str.cat(lt_out_comportamiento.MATRICULA_CALIFICADO.str.cat(lt_out_comportamiento.CHAPTER_CALIFICADO, sep=''), sep='')
            
            df_3 = pd.merge(lt_out_comportamiento[['CONCAT']], base, how='left', on='CONCAT')
            df_3 = df_3.dropna(subset=['MATRICULA_CALIFICADOR'])

            for i, matricula in enumerate(base['CONCAT']):
                if matricula in df_3['CONCAT'].values:
                    base.loc[i, 'FLAG_EVALUACION'] = 'SI'


        elif en == ETipoEva.FLAG_AUTOEVALUACION_y_otros:

            base['CONCAT'] = base.MATRICULA_CALIFICADO.str.cat(base.MATRICULA_CALIFICADO.str.cat(base.CHAPTER_CALIFICADO, sep=''), sep='')
            lt_out_comportamiento['CONCAT'] = lt_out_comportamiento.MATRICULA_CALIFICADOR.str.cat(lt_out_comportamiento.MATRICULA_CALIFICADO.str.cat(lt_out_comportamiento.CHAPTER_CALIFICADO, sep=''), sep='')
            
            df_4 = pd.merge(lt_out_comportamiento[['CONCAT']], base, how='left', on='CONCAT')
            df_4 = df_4.dropna(subset=['MATRICULA_CALIFICADOR'])

            for i, matricula in enumerate(base['CONCAT']):
                if matricula in df_4['CONCAT'].values:
                    base.loc[i, 'FLAG_AUTOEVALUACION'] = 'SI'
        
        elif en == ETipoEva.AUTOEVALUACION_Y_EVALUACION_y_otros:
            base['CONCAT'] = base.MATRICULA_CALIFICADO.str.cat(base.MATRICULA_CALIFICADO.str.cat(base.CHAPTER_CALIFICADO, sep=''), sep='')
            lt_out_comportamiento['CONCAT'] = lt_out_comportamiento.MATRICULA_CALIFICADOR.str.cat(lt_out_comportamiento.MATRICULA_CALIFICADO.str.cat(lt_out_comportamiento.CHAPTER_CALIFICADO, sep=''), sep='')
            
            df_2 = pd.merge(lt_out_comportamiento[['CONCAT']], base, how='left', on='CONCAT')
            df_2 = df_2.dropna(subset=['MATRICULA_CALIFICADOR'])

            for i, matricula in enumerate(base['CONCAT']):
                if matricula in df_2['CONCAT'].values:
                    base.loc[i, 'FLAG_AUTOEVALUACION'] = 'SI'


            base['CONCAT'] = base.MATRICULA_CALIFICADOR.str.cat(base.MATRICULA_CALIFICADO.str.cat(base.CHAPTER_CALIFICADO, sep=''), sep='')
            lt_out_comportamiento['CONCAT'] = lt_out_comportamiento.MATRICULA_CALIFICADOR.str.cat(lt_out_comportamiento.MATRICULA_CALIFICADO.str.cat(lt_out_comportamiento.CHAPTER_CALIFICADO, sep=''), sep='')

            df_1 = pd.merge(lt_out_comportamiento[['CONCAT']], base, how='left', on='CONCAT')
            df_1 = df_1.dropna(subset=['MATRICULA_CALIFICADOR'])

            for i, matricula in enumerate(base['CONCAT']):
                if matricula in df_1['CONCAT'].values:
                    base.loc[i, 'FLAG_EVALUACION'] = 'SI'

        #actualizar FLAG_EXCLUSIÓN
        base.rename(columns={'MATRICULA_CALIFICADOR': 'MATRICULA'}, inplace=True)
        df_4 = pd.merge(base_activos[['MATRICULA']], base, how='left', on='MATRICULA')
        if 'FLAG_AUTOEVALUACION' in df_4 and 'FLAG_EVALUACION' in df_4:
            df_4 = df_4.drop_duplicates(subset=['ESTADO', 'MATRICULA', 'NOMBRES_CALIFICADOR', 'CORREO_CALIFICADOR', 'ROL_CALIFICADOR', 'MATRICULA_CALIFICADO', 'NOMBRES_CALIFICADO', 'CHAPTER_CALIFICADO', 'ROL_CALIFICADO', 'CORREO_CALIFICADO', 'FLAG_AUTOEVALUACION', 'FLAG_EVALUACION', 'FLAG_EXCLUSIÓN'])
        elif 'FLAG_EVALUACION' in df_4:
            df_4 = df_4.drop_duplicates(subset=['ESTADO', 'MATRICULA', 'NOMBRES_CALIFICADOR', 'CORREO_CALIFICADOR', 'ROL_CALIFICADOR', 'MATRICULA_CALIFICADO', 'NOMBRES_CALIFICADO', 'CHAPTER_CALIFICADO', 'ROL_CALIFICADO', 'CORREO_CALIFICADO', 'FLAG_EVALUACION', 'FLAG_EXCLUSIÓN'])
        elif 'FLAG_AUTOEVALUACION' in df_4:
            df_4 = df_4.drop_duplicates(subset=['ESTADO', 'MATRICULA', 'NOMBRES_CALIFICADOR', 'CORREO_CALIFICADOR', 'ROL_CALIFICADOR', 'MATRICULA_CALIFICADO', 'NOMBRES_CALIFICADO', 'CHAPTER_CALIFICADO', 'ROL_CALIFICADO', 'CORREO_CALIFICADO', 'FLAG_AUTOEVALUACION', 'FLAG_EXCLUSIÓN'])

        for i, matricula in enumerate(base['MATRICULA']):
            if matricula in df_4['MATRICULA'].values:
                base.loc[i, 'FLAG_EXCLUSIÓN'] = 'NO'
            else:
                base.loc[i, 'FLAG_EXCLUSIÓN'] = 'SI'
        
        base = base.drop(['CONCAT'], axis=1)
        copia_pega(ruta_plantilla, ruta_out_f)
        df_a_excel(ruta_out_f, 'BASE', base, f_ini = 2)

        wb = xlapp.Workbooks.Open(ruta_out_f)
        xlapp.Visible = True
        wb.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        wb.Save()
        xlapp.Quit()

if __name__ == '__main__':
    avance_medición()