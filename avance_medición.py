from datetime import datetime
from os import path
import pandas as pd
from base import copia_pega, df_a_excel, leer_excel_simple
from enumeraciones import ETipoEva

def avance_medición():
    fec_hoy = datetime.today()
    fecha_hoy_format = fec_hoy.strftime('%Y%m%d')
    ruta1 = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Entrada\PLANTILLA_AVANCE_MEDICION.xlsx'
    ruta_principal = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Salida'
    nombre_archivo = 'AVANCE_MEDICION_PO_{}.xlsx'.format(fecha_hoy_format)
    ruta_out_f = path.join(ruta_principal, nombre_archivo)

    base = leer_excel_simple(ruta1, 'BASE')
    lt_out_comportamiento = leer_excel_simple(ruta1, 'LT_OUT_COMPORTAMIENTO')
    lt_out_capacidad = leer_excel_simple(ruta1, 'LT_OUT_CAPACIDAD')

    lt_out_comportamiento = lt_out_comportamiento.drop_duplicates(subset=['MATRICULA_CALIFICADOR', 'MATRICULA_CALIFICADO', 'CHAPTER_CALIFICADO'])
    lt_out_capacidad = lt_out_capacidad.drop_duplicates(subset=['MATRICULA_CALIFICADOR', 'MATRICULA_CALIFICADO', 'CHAPTER_CALIFICADO'])

    x_menu = input('''
          MENÚ
          ----------------------
          1. FLAG_EVALUACION y PO
          2. FLAG_AUTOEVALUACION y PO
          3. FLAG_EVALUACION y otros
          4. FLAG_AUTOEVALUACION y otros
          5. FLAG_MEDICIÓN y otros

          ELIJA UNA OPCIÓN: 
          ''')
    dic = {
        '1':ETipoEva.FLAG_EVALUACION_y_PO,
        '2':ETipoEva.FLAG_AUTOEVALUACION_y_PO,
        '3':ETipoEva.FLAG_EVALUACION_y_otros,
        '4':ETipoEva.FLAG_AUTOEVALUACION_y_otros,
        '5':ETipoEva.FLAG_MEDICIÓN_y_otros
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
            lt_out_capacidad['CONCAT'] = lt_out_capacidad.MATRICULA_CALIFICADO.str.cat(lt_out_capacidad.MATRICULA_CALIFICADO.str.cat(lt_out_capacidad.CHAPTER_CALIFICADO, sep=''), sep='')
            
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
            lt_out_comportamiento['CONCAT'] = lt_out_comportamiento.MATRICULA_CALIFICADO.str.cat(lt_out_comportamiento.MATRICULA_CALIFICADO.str.cat(lt_out_comportamiento.CHAPTER_CALIFICADO, sep=''), sep='')
            
            df_4 = pd.merge(lt_out_comportamiento[['CONCAT']], base, how='left', on='CONCAT')
            df_4 = df_4.dropna(subset=['MATRICULA_CALIFICADOR'])

            for i, matricula in enumerate(base['CONCAT']):
                if matricula in df_4['CONCAT'].values:
                    base.loc[i, 'FLAG_AUTOEVALUACION'] = 'SI'
        
        elif en == ETipoEva.FLAG_MEDICIÓN_y_otros:

            base['CONCAT'] = base.MATRICULA_CALIFICADOR.str.cat(base.MATRICULA_CALIFICADO.str.cat(base.CHAPTER_CALIFICADO, sep=''), sep='')
            lt_out_comportamiento['CONCAT'] = lt_out_comportamiento.MATRICULA_CALIFICADOR.str.cat(lt_out_comportamiento.MATRICULA_CALIFICADO.str.cat(lt_out_comportamiento.CHAPTER_CALIFICADO, sep=''), sep='')
            
            df_4 = pd.merge(lt_out_comportamiento[['CONCAT']], base, how='left', on='CONCAT')
            df_4 = df_4.dropna(subset=['MATRICULA_CALIFICADOR'])

            for i, matricula in enumerate(base['CONCAT']):
                if matricula in df_4['CONCAT'].values:
                    base.loc[i, 'FLAG_MEDICIÓN'] = 'SI'

        
        base = base.drop(['CONCAT'], axis=1)
        copia_pega(ruta1, ruta_out_f)
        df_a_excel(ruta_out_f, 'BASE', base, f_ini = 2)

if __name__ == '__main__':
    avance_medición()