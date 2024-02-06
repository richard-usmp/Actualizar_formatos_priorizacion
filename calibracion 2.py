from datetime import date, datetime
from getpass import getuser
import pyodbc
import pandas as pd,csv
import numpy as np
import csv
import os
import win32com
from base import apply_logic_DE, apply_logic_capacidades, apply_logic_dimension, copia_pega, df_a_excel, leer_excel_simple, elimina_col_excel, elimina_filas_excel, df_a_excel_header, elimina_col_excel_res_lid, elimina_filas_excel_res_lid
from constantes import PATH_BA
import xlwings as xw
from xlwings.utils import col_name

_conn_params = {
    "server": 'PAUTGSQLP43',
    "database": 'BCP_GDH_PA_STAGE',
    "trusted_connection": "Yes",
    "driver": "{SQL Server}",
}

def calibracion():
    ruta_ba = PATH_BA
    ruta_plantilla = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Entrada\PLANTILLA_RESULTADOS_POSTCALIBRACION_2.xlsx'
    fec_hoy = datetime.today()
    fecha_hoy_format = fec_hoy.strftime('%Y%m%d')
    ruta_BD = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Entrada\Base de datos.xlsx'
    
    base_activos = leer_excel_simple(ruta_ba, 'BD ACTIVOS')
    base_activos.rename(columns={'Matrícula': 'MATRICULA'}, inplace=True)
    resultados_general = leer_excel_simple(ruta_BD, 'Hoja1')

    chapters1 = resultados_general['CHAPTER'].unique()

    #array_chapters = ['ANALISTA DE SEGURIDAD DE INFRAESTRUCTURA','.NET','PROJECT MANAGER','DATA MODELER','MAINFRAME','DATA ARCHITECT','OPERATIONAL RISK','ANALYTICS TRANSLATOR','FRONTEND WEB','VISUAL COMMUNICATION DESIGN','RIESGO DE FRAUDE','DEVELOPER SALESFORCE','CIBERSEGURIDAD','INGENIERÍA FINANCIERA','NETWORKING - TELEPHONY & COLLABORATION','PARAMETRIZATION RISK','CYBERCRIME INVESTIGATIONS','SOLUTION ARCHITECTURE','CONSULTORÍA DE PROCESOS','SMART AUTOMATION','TECHNOLOGY ARCHITECTURE','DATA GOVERNANCE EXPERT','CONSULTORÍA DE SALESFORCE','GROWTH HACKING','ARQUITECTURA SALESFORCE','RIESGOS DE ALM & LIQUIDEZ','RIESGOS DE INVERSIONES Y DERIVADOS']

    for chapter in chapters1:
        print(chapter)
        ruta_principal = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Salida'
        nombre_archivo = 'RESULTADOS_CALIBRACION_{}_{}.xlsx'.format(chapter, fecha_hoy_format)
        ruta_out_f = os.path.join(ruta_principal, nombre_archivo)

        query_resultados = '''
            DECLARE	@CHAPTER varchar(100) = '{}'

            DECLARE @PK_CHAPTER INT = (
            SELECT PK_CHAPTER FROM BCP_GDH_PA_DW.PDIGITAL.D_CHAPTER WHERE DESCRIPCION = @CHAPTER)

            DECLARE @FECHA_EVA INT = (
            SELECT MAX(FK_FECHA) FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE WHERE FK_CHAPTER=@PK_CHAPTER AND CATEGORIA_EVALUACION='OFICIAL'
            )

            DECLARE @FECHA_EVA_PENULTIMA INT = (
            SELECT MAX(FK_FECHA) FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE WHERE FK_CHAPTER=@PK_CHAPTER AND CATEGORIA_EVALUACION='OFICIAL' AND FK_FECHA<@FECHA_EVA
            )

            ;WITH BASE_COLABORADOR AS(
                SELECT DISTINCT FRE.MATRICULA,UC.Nombre + ' ' + UC.Ape_Paterno + ' ' + UC.Ape_Materno NOMBRE_COMPLETO
                FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE FRE
                LEFT JOIN BCP_GDH_PA_DW.GENERAL.D_COLABORADOR UC ON FRE.MATRICULA=RIGHT(UC.MATRICULA,6)
                WHERE FK_FECHA = @FECHA_EVA AND FK_CHAPTER=@PK_CHAPTER
            ),
            BASE_LIDER AS(
                SELECT DISTINCT FRE.MATRICULA_CALIFICADOR,UC.Nombre + ' ' + UC.Ape_Paterno + ' ' + UC.Ape_Materno NOMBRE_COMPLETO
                FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE FRE
                LEFT JOIN BCP_GDH_PA_DW.GENERAL.D_COLABORADOR UC ON FRE.MATRICULA_CALIFICADOR=RIGHT(UC.MATRICULA,6)
	            WHERE (FK_FECHA = @FECHA_EVA OR FK_FECHA = @FECHA_EVA_PENULTIMA) AND FK_CHAPTER=@PK_CHAPTER
            ),
            ROL_CHAPTER AS(
            SELECT FK_CHAPTER,ROL,DESCRIPCION_ROL,AGRUPACION,FK_FECHA_INI,COALESCE(FK_FECHA_FIN,CAST(convert(varchar,GETDATE(),112) AS INT)) FK_FECHA_FIN FROM BCP_GDH_PA_DW.PDIGITAL.F_ROL_CHAPTER WHERE FK_CHAPTER=@PK_CHAPTER
            ),
            CATEGORIA_CAPACIDAD AS
            (
                SELECT FK_CAPACIDAD,
                    FK_CHAPTER,
                    ROL,
                    CATEGORIA_CAPACIDAD,
                    SUBCATEGORIA_CAPACIDAD,
					FK_FECHA_INI,
					COALESCE(FK_FECHA_FIN,CAST(convert(varchar,GETDATE(),112) AS INT)) FK_FECHA_FIN
                FROM BCP_GDH_PA_DW.PDIGITAL.F_CAPACIDAD_CATEGORIA FC
                WHERE FK_CHAPTER=@PK_CHAPTER
            ),
            COLABORADOR_CAPACIDAD_INI AS(
            SELECT *
            FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_CAPACIDAD FR
            WHERE FR.FK_CHAPTER=@PK_CHAPTER AND ((FR.FK_FECHA = @FECHA_EVA) OR (FR.FK_FECHA = @FECHA_EVA_PENULTIMA AND MATRICULA_CALIFICADOR<>MATRICULA))
            ),COLABORADOR_CAPACIDAD_FIN AS(
            SELECT FR.MATRICULA_CALIFICADOR,
                LID.NOMBRE_COMPLETO AS NOMBRES_CALIFICADOR,
                coalesce (FRCL.DESCRIPCION_ROL, FR.ROL_CALIFICADOR) ROL_CALIFICADOR,
                DP.DESCRIPCION CHAPTER,
                FR.MATRICULA,
                BC.NOMBRE_COMPLETO AS NOMBRES_CALIFICADO,
                coalesce (FRCC.DESCRIPCION_ROL, FR.ROL) ROL,
                DC.DESCRIPCION,
				PK_CAPACIDAD,
                CC.CATEGORIA_CAPACIDAD,
                CC.SUBCATEGORIA_CAPACIDAD,
                FR.N_NIVEL,
                FR.NIVEL,
                FR.FLAG_CONOCIMIENTO,
                FR.FK_FECHA,
                CASE
                    WHEN FR.FK_FECHA=@FECHA_EVA_PENULTIMA THEN 'EVALUACION ANTERIOR'
                    ELSE FR.TIPO_EVALUACION
                END EVALUACION
            FROM COLABORADOR_CAPACIDAD_INI FR
                INNER JOIN BASE_COLABORADOR BC ON FR.MATRICULA = BC.MATRICULA
                LEFT JOIN BASE_LIDER LID ON FR.MATRICULA_CALIFICADOR=RIGHT(LID.MATRICULA_CALIFICADOR,6)
                LEFT JOIN BCP_GDH_PA_DW.PDIGITAL.D_CHAPTER DP ON FR.FK_CHAPTER=DP.PK_CHAPTER
                LEFT JOIN BCP_GDH_PA_DW.PDIGITAL.D_CAPACIDAD DC ON FR.FK_CAPACIDAD=DC.PK_CAPACIDAD
                LEFT JOIN CATEGORIA_CAPACIDAD CC ON FR.FK_CAPACIDAD=CC.FK_CAPACIDAD AND FR.ROL=CC.ROL AND FR.FK_FECHA BETWEEN CC.FK_FECHA_INI AND CC.FK_FECHA_FIN
                LEFT JOIN ROL_CHAPTER FRCC ON FR.ROL=FRCC.ROL AND FR.FK_FECHA BETWEEN FRCC.FK_FECHA_INI AND FRCC.FK_FECHA_FIN
                LEFT JOIN ROL_CHAPTER FRCL ON FR.ROL_CALIFICADOR=FRCL.ROL AND FR.FK_FECHA BETWEEN FRCL.FK_FECHA_INI AND FRCL.FK_FECHA_FIN
            ),COLABORADOR_COMPORTAMIENTO AS(
            SELECT FR.MATRICULA,
                FR.ROL,
                FR.MATRICULA_CALIFICADOR,
                FR.ROL_CALIFICADOR,
                DC.DESCRIPCION AS COMPORTAMIENTO,
                DP.DESCRIPCION,
                FR.N_NIVEL,
                FR.FK_FECHA
            FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_COMPORTAMIENTO FR
                INNER JOIN BASE_COLABORADOR CC ON FR.MATRICULA = CC.MATRICULA
                LEFT JOIN BCP_GDH_PA_DW.PDIGITAL.D_COMPORTAMIENTO DC ON FR.FK_COMPORTAMIENTO=DC.PK_COMPORTAMIENTO
                LEFT JOIN BCP_GDH_PA_DW.PDIGITAL.D_CHAPTER DP ON FR.FK_CHAPTER=DP.PK_CHAPTER
			WHERE FR.FK_CHAPTER=@PK_CHAPTER AND ((FR.FK_FECHA = @FECHA_EVA) OR (FR.FK_FECHA = @FECHA_EVA_PENULTIMA AND MATRICULA_CALIFICADOR<>FR.MATRICULA))
            ),PIVOT_COMPORTAMIENTO AS(
            SELECT MATRICULA,
                ROL,
                MATRICULA_CALIFICADOR,
                ROL_CALIFICADOR,
                DESCRIPCION,
                [Domain expertise] AS N_NIVELDOMAINEXPERTISE,
                [Análisis y solución de problemas] AS N_NIVELRESOL,
                [Liderazgo y comunicación] AS N_NIVELLIDERAZG,
                [Fit cultural] AS N_NIVELCULTURAL,
                FK_FECHA
            FROM COLABORADOR_COMPORTAMIENTO
            PIVOT
            (
                AVG(N_NIVEL)
                FOR COMPORTAMIENTO IN ([Análisis y solución de problemas], [Liderazgo y comunicación], [Fit cultural], [Domain expertise])
            ) AS T
            ),
            COLABORADOR_EXPERTISE AS(
            SELECT RE.MATRICULA,
                RE.MATRICULA_CALIFICADOR,
                DP.DESCRIPCION,
                RE.N_NIVEL,
                RE.NIVEL,
                RE.FK_FECHA
            FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE RE
                INNER JOIN BASE_COLABORADOR CC ON RE.MATRICULA = CC.MATRICULA
                LEFT JOIN BCP_GDH_PA_DW.PDIGITAL.D_CHAPTER DP ON RE.FK_CHAPTER=DP.PK_CHAPTER
            WHERE RE.FK_CHAPTER=@PK_CHAPTER AND (RE.FK_FECHA = @FECHA_EVA OR RE.FK_FECHA = @FECHA_EVA_PENULTIMA AND MATRICULA_CALIFICADOR<>RE.MATRICULA)
            )
            SELECT CCF.*,PC.N_NIVELCULTURAL,PC.N_NIVELDOMAINEXPERTISE,PC.N_NIVELLIDERAZG,PC.N_NIVELRESOL,CE.N_NIVEL N_NIVELEXPERTISE
            FROM COLABORADOR_CAPACIDAD_FIN CCF
            LEFT JOIN PIVOT_COMPORTAMIENTO PC ON CCF.FK_FECHA=PC.FK_FECHA AND CCF.MATRICULA_CALIFICADOR=PC.MATRICULA_CALIFICADOR AND CCF.MATRICULA=PC.MATRICULA
            LEFT JOIN COLABORADOR_EXPERTISE CE ON CCF.FK_FECHA=CE.FK_FECHA AND CCF.MATRICULA_CALIFICADOR=CE.MATRICULA_CALIFICADOR AND CCF.MATRICULA=CE.MATRICULA;
        '''.format(chapter)

        #df_base = select(query_resultados)
        df_base = resultados_general[resultados_general['CHAPTER'] == chapter]
        df_base = pd.merge(df_base, base_activos[['MATRICULA', 'Correo electronico','Grado Salarial']], on='MATRICULA', how='left')
        df_base = df_base.sort_values(by=['MATRICULA','EVALUACION'])

        copia_pega(ruta_plantilla, ruta_out_f)

        df1 = df_base[['MATRICULA_CALIFICADOR', 'NOMBRES_CALIFICADOR', 'ROL_CALIFICADOR', 'CHAPTER', 'MATRICULA', 'NOMBRES_CALIFICADO', 'Correo electronico']]
        df2 = df_base[['ROL', 'DESCRIPCION', 'CATEGORIA_CAPACIDAD', 'SUBCATEGORIA_CAPACIDAD']]
        df3 = df_base[['N_NIVEL', 'NIVEL', 'FLAG_CONOCIMIENTO', 'N_NIVELDOMAINEXPERTISE']]
        df4 = df_base[['N_NIVELRESOL']]
        df5 = df_base[['N_NIVELLIDERAZG']]
        df6 = df_base[['N_NIVELCULTURAL']]
        df7 = df_base[['N_NIVELEXPERTISE']]
        df8 = df_base[['EVALUACION']]

        df_a_excel(ruta_out_f, 'BASE', df1, f_ini = 2, c_ini = 1)
        df_a_excel(ruta_out_f, 'BASE', df2, f_ini = 2, c_ini = 12)
        df_a_excel(ruta_out_f, 'BASE', df3, f_ini = 2, c_ini = 18)
        df_a_excel(ruta_out_f, 'BASE', df4, f_ini = 2, c_ini = 23)
        df_a_excel(ruta_out_f, 'BASE', df5, f_ini = 2, c_ini = 25)
        df_a_excel(ruta_out_f, 'BASE', df6, f_ini = 2, c_ini = 27)
        df_a_excel(ruta_out_f, 'BASE', df7, f_ini = 2, c_ini = 29)
        df_a_excel(ruta_out_f, 'BASE', df8, f_ini = 2, c_ini = 31)

        #LLENADO DE CUADRO RESUMEN TMs
        df_capacidades = df_base[['DESCRIPCION','SUBCATEGORIA_CAPACIDAD']].sort_values(by='SUBCATEGORIA_CAPACIDAD', ascending=False)
        df_capacidades = df_capacidades.drop_duplicates(subset = 'DESCRIPCION')
        array_capacidades = df_capacidades['DESCRIPCION'].unique()
        q_capacidades = len(array_capacidades)
        print(q_capacidades)
        q_capacidades_core = len(df_capacidades[df_capacidades['SUBCATEGORIA_CAPACIDAD'] == 'PRINCIPAL'].index)
        df_a_excel(ruta_out_f, 'Resumen TMs', array_capacidades, f_ini = 5, c_ini = 11)
        df_a_excel(ruta_out_f, 'Resumen TMs', df_base[['MATRICULA_CALIFICADOR','NOMBRES_CALIFICADOR', 'ROL_CALIFICADOR', 'MATRICULA', 'NOMBRES_CALIFICADO', 'ROL', 'Grado Salarial','EVALUACION']].drop_duplicates(), f_ini = 6, c_ini = 3)

        # Abrir el archivo de Excel
        app = xw.App(visible=False)
        workbook = app.books.open(ruta_out_f)
        worksheet = workbook.sheets['Resumen TMs']

        #capacidad cuadro Resumen TMs
        color = (255, 165, 0)
        starting_cell = 'K5'
        row_index = worksheet.range(starting_cell).row
        col_index = worksheet.range(starting_cell).column

        for i in range(q_capacidades_core):
            cell = worksheet.cells(row_index, col_index + i)
            cell.color = color

        rango1 = 'K6:AN1000'
        valores_copia = worksheet.range(rango1).value
        worksheet.range(rango1).value = valores_copia

        rango2 = 'AP6:AT1000'
        valores_copia2 = worksheet.range(rango2).value
        worksheet.range(rango2).value = valores_copia2

        workbook.save()
        workbook.close()
        app.quit()
        #app.kill()


        elimina_col_excel(ruta_out_f, 'Resumen TMs', q_capacidades)
        elimina_filas_excel(ruta_out_f, 'Resumen TMs')

        #REEMPLAZANDO LÍDERES EN LA AUTO
        df_resumen_tms = leer_excel_simple(ruta_out_f, 'Resumen TMs', f_inicio=5, c_inicio=2)
        df_calificador_calificado = df_resumen_tms[['MatriculaCalificador', 'NombresCalificador', 'MatriculaCalificado', 'NombresCalificado', 'TipoEvaluacion']]
        df_calificador_calificado_eva = df_resumen_tms[df_resumen_tms['TipoEvaluacion'] == 'EVALUACION']
            
        df_calificador_calificado_new = pd.merge(df_calificador_calificado, df_calificador_calificado_eva, how='left', on='MatriculaCalificado')
        df_calificador_calificado_new = df_calificador_calificado_new[['MatriculaCalificador_y','NombresCalificador_y','RolCalificador']]
        df_a_excel(ruta_out_f, 'Resumen TMs', df_calificador_calificado_new, f_ini = 6, c_ini = 3)


        #ALERTAS
        resumen_tms_para_alerta = df_resumen_tms.drop(['Chapter', 'MatriculaCalificador', 'NombresCalificador', 'RolCalificador', 'NombresCalificado', 'RolCalificado', 
                                                                'GS Calificado', 'Promedio', 'Alerta', 'Comentario'], axis=1)
        duplicados = resumen_tms_para_alerta[(resumen_tms_para_alerta['TipoEvaluacion'] != 'AUTOEVALUACION') & (resumen_tms_para_alerta['MatriculaCalificado'].duplicated(keep=False))]
        
        diferencias_list = []
        if duplicados.empty:
            print('Dataframe duplicados vacio, no hay comparacion entre evaluacion nueva vs evaluación anterior.')
        else:
            for index, row in duplicados.iterrows():
                duplicate_rows = resumen_tms_para_alerta[(resumen_tms_para_alerta['MatriculaCalificado'] == row['MatriculaCalificado']) & (resumen_tms_para_alerta['TipoEvaluacion'] != 'AUTOEVALUACION')]
                duplicate_rows = duplicate_rows.sort_values(by='TipoEvaluacion')
                if len(duplicate_rows) == 2:
                    first_row = duplicate_rows.iloc[0]
                    second_row = duplicate_rows.iloc[1]

                    diferencias = {'MatriculaCalificado': row['MatriculaCalificado'], 'TipoEvaluacion': row['TipoEvaluacion']}
                    for col in resumen_tms_para_alerta.columns:
                        if col != 'MatriculaCalificado':
                            if isinstance(first_row[col], (int, float)) and isinstance(second_row[col], (int, float)):
                                diferencias[col] = first_row[col] - second_row[col]
                            # else:
                            #     diferencias[col] = None

                    if diferencias:
                        diferencias_list.append(diferencias)
            diferencias_df = pd.DataFrame(diferencias_list)

            if not(diferencias_df.empty):
                diferencias_df_new = diferencias_df.copy()
                diferencias_df_new['Concatenadas_2'] = ''
                for col in diferencias_df.columns[1:]:
                    if col == 'NivelDomainExpertise':
                        diferencias_df[col] = diferencias_df[col].apply(apply_logic_DE)
                    elif col == 'Análisis y solución de problemas':
                        diferencias_df[col] = diferencias_df[col].apply(apply_logic_dimension)
                    elif col == 'Liderazgo y Comunicación':
                        diferencias_df[col] = diferencias_df[col].apply(apply_logic_dimension)
                    elif col == 'Fit Cultural':
                        diferencias_df[col] = diferencias_df[col].apply(apply_logic_dimension)
                    elif col == 'Nivel general':
                        diferencias_df[col] = diferencias_df[col].apply(apply_logic_dimension)
                    elif col == 'TipoEvaluacion':
                        pass
                    else:
                        diferencias_df[col] = diferencias_df[col].apply(apply_logic_capacidades)

                for fila in range(len(diferencias_df_new.index)):
                    cont=0
                    for columnas in range(2,q_capacidades):
                        if diferencias_df_new.iloc[fila, columnas] > 0: cont += 1
                    if cont >= 4: diferencias_df_new.at[fila, 'Concatenadas_2'] = 'Subió 4 o más capacidades'

                diferencias_df['Concatenadas'] = diferencias_df.iloc[:, 2:].apply(lambda row: '' if all(val == '' for val in row) else ', '.join(set(val for val in row if val != '')), axis=1)
                new_df = diferencias_df[['MatriculaCalificado','TipoEvaluacion', 'Concatenadas']]
                new_df = pd.merge(new_df,diferencias_df_new,how='left', on=['MatriculaCalificado','TipoEvaluacion'])
                
                new_df['Concatenadas_final'] = np.where(new_df['Concatenadas_2'].fillna('-') != '-', new_df['Concatenadas'] +', '+ new_df['Concatenadas_2'].fillna('-'), new_df['Concatenadas'])
                new_df['Concatenadas_final'] = np.where(new_df['Concatenadas_2'].fillna('-') != '-', new_df['Concatenadas'] +', '+ new_df['Concatenadas_2'].fillna('-'), new_df['Concatenadas'])
                new_df['Concatenadas_final'] = new_df['Concatenadas_final'].apply(lambda x: x.lstrip(', '))


                new_df_merge = pd.merge(df_calificador_calificado, new_df, how='left', on='MatriculaCalificado')
                new_df_merge = new_df_merge.drop_duplicates(subset=['MatriculaCalificador', 'NombresCalificador', 'MatriculaCalificado', 'NombresCalificado', 'TipoEvaluacion_x'])

                df_a_excel(ruta_out_f, 'Resumen TMs', new_df_merge[['Concatenadas_final']], f_ini = 6, c_ini = (17 + q_capacidades)) #

                app = xw.App(visible=False)
                workbook = app.books.open(ruta_out_f)
                worksheet = workbook.sheets['Resumen TMs']

                starting_cell = 'B6'
                row_index_eva = worksheet.range(starting_cell).row
                col_index_eva = worksheet.range(starting_cell).column

                for i, tipo_eva in enumerate(new_df_merge['TipoEvaluacion_x']):
                    if tipo_eva == 'EVALUACION ANTERIOR' or tipo_eva == 'AUTOEVALUACION':
                        cell = worksheet.range(row_index_eva + i, col_index_eva + 15 + q_capacidades)
                        cell.font.color = (255, 255, 255)

                workbook.save()
                workbook.close()
                app.quit()
                #app.kill()

        #LLENADO RESUMEN LÍDERES
        resumen_tms = df_resumen_tms
        resumen_tms = resumen_tms[(resumen_tms['TipoEvaluacion'] != 'EVALUACION ANTERIOR') & (resumen_tms['TipoEvaluacion'] != 'AUTOEVALUACION')]
        base_merge = df_calificador_calificado[(df_calificador_calificado['TipoEvaluacion'] != 'EVALUACION ANTERIOR') & (df_calificador_calificado['TipoEvaluacion'] != 'AUTOEVALUACION')]
        cant_evaluados = base_merge['NombresCalificador'].value_counts().reset_index()
        cant_evaluados.columns = ['NombresCalificador', 'Cant_evaluados']
        cant_evaluados = cant_evaluados.sort_values(by='NombresCalificador')
        resumen_tms['NombresCalificador'] = pd.to_numeric(resumen_tms['NombresCalificador'], errors='coerce')#
        promedios = resumen_tms.groupby('NombresCalificador').mean()
        promedios2 = promedios.drop(['Promedio','Comentario','NivelDomainExpertise','Análisis y solución de problemas','Liderazgo y Comunicación','Fit Cultural','Nivel general'], axis=1)

        df_a_excel(ruta_out_f, 'Resumen Líderes', cant_evaluados[['NombresCalificador']], f_ini = 5, c_ini = 2)
        df_a_excel(ruta_out_f, 'Resumen Líderes', cant_evaluados[['Cant_evaluados']], f_ini = 5, c_ini = 3)
        df_a_excel_header(ruta_out_f, 'Resumen Líderes', promedios2, f_ini = 4, c_ini = 4)
        df_a_excel(ruta_out_f, 'Resumen Líderes', promedios[['NivelDomainExpertise']], f_ini = 5, c_ini = 33)

        elimina_col_excel_res_lid(ruta_out_f, 'Resumen Líderes', q_capacidades)
        elimina_filas_excel_res_lid(ruta_out_f, 'Resumen Líderes')
        
        app_lid = xw.App(visible=False)
        workbook_lid = app_lid.books.open(ruta_out_f)
        worksheet_lid = workbook_lid.sheets['Resumen Líderes']

        color_lid = (255, 165, 0)
        starting_cell_lid = 'D4'
        row_index_lid = worksheet_lid.range(starting_cell_lid).row
        col_index_lid = worksheet_lid.range(starting_cell_lid).column
        for i in range(q_capacidades_core):
            cell_lid = worksheet_lid.cells(row_index_lid, col_index_lid + i)
            cell_lid.color = color_lid

        workbook_lid.save()
        workbook_lid.close()
        app_lid.quit()
        #app_lid.kill()

        #ACTUALIZAR EXCEL TABLAS DINAMICAS Y LISTAS
        xlapp = win32com.client.DispatchEx("Excel.Application")
        wb = xlapp.Workbooks.Open(ruta_out_f)
        xlapp.Visible = False
        wb.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        wb.Save()
        xlapp.Quit()


def select(q,t_params=()):
    cnxn = pyodbc.connect(**_conn_params)
    cnxn.autocommit = True
    c = cnxn.cursor()
    c.execute(q,t_params)

    columns = [column[0] for column in c.description]
    results = []
    for row in c.fetchall():
        results.append(dict(zip(columns, row)))

    cnxn.close()

    return pd.DataFrame(results)

if __name__ == '__main__':
    calibracion()
