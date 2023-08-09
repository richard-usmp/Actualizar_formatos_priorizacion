from datetime import date, datetime
from getpass import getuser
import pyodbc
import pandas as pd,csv
import numpy as np
import csv
import os
import win32com
from base import copia_pega, df_a_excel, leer_excel_simple, elimina_col_excel, elimina_filas_excel, df_a_excel_header, elimina_col_excel_res_lid, elimina_filas_excel_res_lid
from constantes import PATH_BA
import xlwings as xw
from xlwings.utils import col_name

_conn_params = {
    "server": 'PUGINSQLP01',
    "database": 'BCP_GDH_PA_STAGE',
    "trusted_connection": "Yes",
    "driver": "{SQL Server}",
}

def calibracion():
    ruta_ba = PATH_BA
    ruta_plantilla = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Entrada\PLANTILLA_RESULTADOS_POSTCALIBRACION.xlsx'
    ruta_BD = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Entrada\Base de datos.xlsx' #BORRAR DESPUES DE CONECTAR CON LA DB
    fec_hoy = datetime.today()
    fecha_hoy_format = fec_hoy.strftime('%Y%m%d')
    
    base_activos = leer_excel_simple(ruta_ba, 'BD ACTIVOS')
    base_activos.rename(columns={'Matrícula': 'MATRICULA'}, inplace=True)
    resultados_capacidad_categoria = leer_excel_simple(ruta_BD, 'Hoja1') #BORRAR DESPUES DE CONECTAR CON LA DB
    resultados_resultado_comportamiento = leer_excel_simple(ruta_BD, 'Hoja2') #BORRAR DESPUES DE CONECTAR CON LA DB
    resultados_resultado_expertise = leer_excel_simple(ruta_BD, 'Hoja3') #BORRAR DESPUES DE CONECTAR CON LA DB

    # resultados_capacidad_categoria = select(query1) #QUITAR COMENTARIO DESPUES DE CONECTAR CON LA DB
    # resultados_resultado_comportamiento = select(query2) #QUITAR COMENTARIO DESPUES DE CONECTAR CON LA DB
    # resultados_resultado_expertise = select(query3) #QUITAR COMENTARIO DESPUES DE CONECTAR CON LA DB

    chapters1 = resultados_capacidad_categoria['DESCRIPCION'].unique()
    chapters2 = resultados_resultado_comportamiento['DESCRIPCION'].unique()
    chapters3 = resultados_resultado_expertise['DESCRIPCION'].unique()
    for chapter in chapters1:
        print(chapter)
        ruta_principal = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Salida'
        nombre_archivo = 'RESULTADOS_{}_{}_PRECALIBRACION.xlsx'.format(chapter, fecha_hoy_format)
        ruta_out_f = os.path.join(ruta_principal, nombre_archivo)
        resultados_capacidad_categoria_chapter = resultados_capacidad_categoria[resultados_capacidad_categoria['DESCRIPCION'] == chapter]
        resultados_resultado_comportamiento_chapter = resultados_resultado_comportamiento[resultados_resultado_comportamiento['DESCRIPCION'] == chapter]
        resultados_resultado_expertise_chapter = resultados_resultado_expertise[resultados_resultado_expertise['DESCRIPCION'] == chapter]

        query1 = '''
            DECLARE	@CHAPTER varchar(100) = 'BACKEND JAVA'

            DECLARE @PK_CHAPTER INT = (
            SELECT PK_CHAPTER FROM BCP_GDH_PA_DW.PDIGITAL.D_CHAPTER WHERE DESCRIPCION = @CHAPTER)

            DECLARE @FECHA_EVA INT = (
            SELECT MAX(FK_FECHA) FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE WHERE FK_CHAPTER=@PK_CHAPTER)

			DECLARE @FECHA_EVA_PENULTIMA INT = (
			SELECT MAX(FK_FECHA) FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE WHERE FK_CHAPTER=@PK_CHAPTER AND FK_FECHA<(SELECT MAX(FK_FECHA) FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE WHERE FK_CHAPTER=@PK_CHAPTER))

            ;WITH TCC AS
            (
                SELECT FK_CAPACIDAD,
                    FK_CHAPTER,
                    ROL,
                    CATEGORIA_CAPACIDAD,
                    SUBCATEGORIA_CAPACIDAD
                FROM BCP_GDH_PA_DW.PDIGITAL.F_CAPACIDAD_CATEGORIA FC
                WHERE FK_FECHA_FIN IS NULL
            )
            SELECT FR.MATRICULA_CALIFICADOR,
                LID.Nombres + ' ' + LID.Ape_Paterno + ' ' + LID.Ape_Materno AS NOMBRES_CALIFICADOR,
                FR.ROL_CALIFICADOR,
                DP.DESCRIPCION,
                FR.MATRICULA,
                CAL.Nombres + ' ' + CAL.Ape_Paterno + ' ' + CAL.Ape_Materno AS NOMBRES_CALIFICADO,
                FR.ROL,
                DC.DESCRIPCION,
                TC.CATEGORIA_CAPACIDAD,
                TC.SUBCATEGORIA_CAPACIDAD,
                FR.N_NIVEL,
                FR.NIVEL,
                FR.FLAG_CONOCIMIENTO,
				FK_FECHA
            FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_CAPACIDAD FR
                LEFT JOIN BCP_GDH_PA_DW.PDIGITAL.D_CHAPTER DP ON FR.FK_CHAPTER=DP.PK_CHAPTER
                LEFT JOIN BCP_GDH_PA_DW.PDIGITAL.D_CAPACIDAD DC ON FR.FK_CAPACIDAD=DC.PK_CAPACIDAD
                LEFT JOIN TCC TC ON FR.FK_CAPACIDAD=TC.FK_CAPACIDAD AND
                                    FR.FK_CHAPTER=TC.FK_CHAPTER AND
                                    FR.ROL=TC.ROL
                LEFT JOIN BCP_GDH_DW_UGI.UGI.COLABORADORES LID ON FR.MATRICULA_CALIFICADOR=RIGHT(LID.Matricula,6)
                LEFT JOIN BCP_GDH_DW_UGI.UGI.COLABORADORES CAL ON FR.MATRICULA=RIGHT(CAL.Matricula,6)
            WHERE FK_FECHA = @FECHA_EVA AND FR.FK_CHAPTER=@PK_CHAPTER or FK_FECHA = @FECHA_EVA_PENULTIMA AND FR.FK_CHAPTER=@PK_CHAPTER;
        '''.format(chapter)
        query2 = '''
            DECLARE	@CHAPTER varchar(100) = '{}'

            DECLARE @PK_CHAPTER INT = (
            SELECT PK_CHAPTER FROM BCP_GDH_PA_DW.PDIGITAL.D_CHAPTER WHERE DESCRIPCION = @CHAPTER)

            DECLARE @FECHA_EVA INT = (
            SELECT MAX(FK_FECHA) FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE WHERE FK_CHAPTER=@PK_CHAPTER)

            DECLARE @FECHA_EVA_PENULTIMA INT = (
			SELECT MAX(FK_FECHA) FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE WHERE FK_CHAPTER=@PK_CHAPTER AND FK_FECHA<(SELECT MAX(FK_FECHA) FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE WHERE FK_CHAPTER=@PK_CHAPTER))

            ;WITH T_COMPORTAMIENTO AS
            (
            SELECT FR.MATRICULA,
                FR.ROL,
                FR.MATRICULA_CALIFICADOR,
                FR.ROL_CALIFICADOR,
                DC.DESCRIPCION AS COMPORTAMIENTO,
				DP.DESCRIPCION,
                N_NIVEL,
				FK_FECHA
            FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_COMPORTAMIENTO FR
                LEFT JOIN BCP_GDH_PA_DW.PDIGITAL.D_COMPORTAMIENTO DC ON FR.FK_COMPORTAMIENTO=DC.PK_COMPORTAMIENTO
                LEFT JOIN BCP_GDH_PA_DW.PDIGITAL.D_CHAPTER DP ON FR.FK_CHAPTER=DP.PK_CHAPTER
            WHERE FK_FECHA = @FECHA_EVA AND FR.FK_CHAPTER=@PK_CHAPTER or FK_FECHA = @FECHA_EVA_PENULTIMA AND FR.FK_CHAPTER=@PK_CHAPTER
            )
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
            FROM T_COMPORTAMIENTO
            PIVOT
            (
                AVG(N_NIVEL)
                FOR COMPORTAMIENTO IN ([Análisis y solución de problemas], [Liderazgo y comunicación], [Fit cultural], [Domain expertise])
            ) AS T
        '''.format(chapter)
        query3 = '''
            DECLARE	@CHAPTER varchar(100) = '{}'

            DECLARE @PK_CHAPTER INT = (
            SELECT PK_CHAPTER FROM BCP_GDH_PA_DW.PDIGITAL.D_CHAPTER WHERE DESCRIPCION = @CHAPTER)

            DECLARE @FECHA_EVA INT = (
            SELECT MAX(FK_FECHA) FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE WHERE FK_CHAPTER=@PK_CHAPTER)

            DECLARE @FECHA_EVA_PENULTIMA INT = (
			SELECT MAX(FK_FECHA) FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE WHERE FK_CHAPTER=@PK_CHAPTER AND FK_FECHA<(SELECT MAX(FK_FECHA) FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE WHERE FK_CHAPTER=@PK_CHAPTER))

            SELECT MATRICULA,
                MATRICULA_CALIFICADOR,
				DP.DESCRIPCION,
                N_NIVEL,
                NIVEL,
				FK_FECHA
            FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE RE
				LEFT JOIN BCP_GDH_PA_DW.PDIGITAL.D_CHAPTER DP ON RE.FK_CHAPTER=DP.PK_CHAPTER
            WHERE FK_FECHA = @FECHA_EVA AND FK_CHAPTER=@PK_CHAPTER or FK_FECHA = @FECHA_EVA_PENULTIMA AND FK_CHAPTER=@PK_CHAPTER;
        '''.format(chapter)

        resultados_resultado_comportamiento_chapter = resultados_resultado_comportamiento_chapter.drop(['ROL', 'ROL_CALIFICADOR'], axis=1)
        resultados_capacidad_categoria_chapter['CONCAT'] = resultados_capacidad_categoria_chapter.MATRICULA.str.cat(resultados_capacidad_categoria_chapter.MATRICULA_CALIFICADOR, sep='')
        resultados_resultado_comportamiento_chapter['CONCAT'] = resultados_resultado_comportamiento_chapter.MATRICULA.str.cat(resultados_resultado_comportamiento_chapter.MATRICULA_CALIFICADOR, sep='')
        resultados_resultado_expertise_chapter['CONCAT'] = resultados_resultado_expertise_chapter.MATRICULA.str.cat(resultados_resultado_expertise_chapter.MATRICULA_CALIFICADOR, sep='')

        df_1 = pd.merge(resultados_resultado_comportamiento_chapter, resultados_capacidad_categoria_chapter, how='left', on='CONCAT')
        df_2 = pd.merge(resultados_resultado_expertise_chapter, df_1, how='left', on='CONCAT')

        df_resultado = pd.merge(df_2, base_activos[['MATRICULA', 'Correo electronico']], on='MATRICULA', how='left')

        df_resultado['EVALUACION'] = ['AUTOEVALUACIÓN' if matricula == m_calificador else 'EVALUACIÓN' for matricula, m_calificador in zip(df_resultado['MATRICULA'], df_resultado['MATRICULA_CALIFICADOR'])]

        max_fecha = df_resultado['FECHA'].max()
        df_resultado['FLAG_NUMBER_EVALUACION'] = ['EVALUACION ANTERIOR' if valor < max_fecha else '' for valor in df_resultado['FECHA']]

        df_resultado = df_resultado[~((df_resultado['EVALUACION'] == 'AUTOEVALUACIÓN') & (df_resultado['FLAG_NUMBER_EVALUACION'] == 'EVALUACION ANTERIOR'))]
        df_resultado.reset_index(drop=True, inplace=True)

        df_resultado['EVALUACION'] = [valor if flag == 'EVALUACION ANTERIOR' else evaluacion for valor, flag, evaluacion in zip(df_resultado['FLAG_NUMBER_EVALUACION'], df_resultado['FLAG_NUMBER_EVALUACION'], df_resultado['EVALUACION'])]

        copia_pega(ruta_plantilla, ruta_out_f)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['MATRICULA_CALIFICADOR']], f_ini = 2, c_ini = 1)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['NOMBRES_CALIFICADOR']], f_ini = 2, c_ini = 2)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['ROL_CALIFICADOR']], f_ini = 2, c_ini = 3)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['DESCRIPCION']], f_ini = 2, c_ini = 4)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['MATRICULA']], f_ini = 2, c_ini = 5)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['NOMBRES_CALIFICADO']], f_ini = 2, c_ini = 6)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['Correo electronico']], f_ini = 2, c_ini = 7)
        #df_a_excel(ruta_out_f, 'BASE', [['']], f_ini = 2, c_ini = 8)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['ROL']], f_ini = 2, c_ini = 12)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['DESCRIPCION2']], f_ini = 2, c_ini = 13)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['CATEGORIA_CAPACIDAD']], f_ini = 2, c_ini = 14)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['SUBCATEGORIA_CAPACIDAD']], f_ini = 2, c_ini = 15)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['N_NIVEL_y']], f_ini = 2, c_ini = 18)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['NIVEL_y']], f_ini = 2, c_ini = 19)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['FLAG_CONOCIMIENTO']], f_ini = 2, c_ini = 20)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['N_NIVELDOMAINEXPERTISE']], f_ini = 2, c_ini = 21)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['N_NIVELRESOL']], f_ini = 2, c_ini = 23)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['N_NIVELLIDERAZG']], f_ini = 2, c_ini = 25)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['N_NIVELCULTURAL']], f_ini = 2, c_ini = 27)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['N_NIVEL_x']], f_ini = 2, c_ini = 29)
        df_a_excel(ruta_out_f, 'BASE', df_resultado[['EVALUACION']], f_ini = 2, c_ini = 31)

        #LLENADO DE CUADRO RESUMEN TMs
        base = leer_excel_simple(ruta_out_f, 'BASE')
        base1 = base.drop_duplicates(subset=['MatriculaCalificador', 'NombresCalificador', 'MatriculaCalificado', 'NombresCalificado', 'TipoEvaluacion'])
        base_solo_eva = base1[base1['TipoEvaluacion'] == 'EVALUACIÓN']

        base_merge = pd.merge(base1, base_solo_eva, how='left', on='MatriculaCalificado')


        #GS Calificado cuadro Resumen TMs
        base_activos.rename(columns={'Nombre completo': 'NombresCalificado_x'}, inplace=True)
        gs_Calificado = pd.merge(base_merge[['NombresCalificado_x']], base_activos[['NombresCalificado_x', 'Grado Salarial']], on='NombresCalificado_x', how='left')

        # Abrir el archivo de Excel
        app = xw.App(visible=False)
        workbook = app.books.open(ruta_out_f)
        worksheet = workbook.sheets['Resumen TMs']

        #capacidad cuadro Resumen TMs
        color = (255, 165, 0)
        starting_cell = 'K5'
        row_index = worksheet.range(starting_cell).row
        col_index = worksheet.range(starting_cell).column
        capacidades = df_resultado.drop(['MATRICULA', 'MATRICULA_CALIFICADOR', 'N_NIVEL_x', 'NIVEL_x', 'CONCAT', 'MATRICULA_x', 'MATRICULA_CALIFICADOR_x', 'Correo electronico', 
                                        'N_NIVELDOMAINEXPERTISE', 'N_NIVELRESOL', 'N_NIVELLIDERAZG', 'N_NIVELCULTURAL', 'MATRICULA_CALIFICADOR_y', 'NOMBRES_CALIFICADOR', 
                                        'ROL_CALIFICADOR', 'DESCRIPCION', 'MATRICULA_y', 'NOMBRES_CALIFICADO', 'ROL', 'CATEGORIA_CAPACIDAD', 'N_NIVEL_y', 'NIVEL_y', 'FLAG_CONOCIMIENTO'], axis=1)
        capacidades = capacidades.sort_values(by='SUBCATEGORIA_CAPACIDAD', ascending=False)
        capacidades_no_duplicada = capacidades.drop_duplicates(subset = 'DESCRIPCION2')
        capacidades_principal = capacidades[capacidades['SUBCATEGORIA_CAPACIDAD'] == 'PRINCIPAL']
        capacidades_traspuesta = capacidades_no_duplicada[['DESCRIPCION2']].T
        for i, capa in enumerate(capacidades_no_duplicada['DESCRIPCION2']):
            if capa in capacidades_principal[['DESCRIPCION2']].values:
                cell = worksheet.cells(row_index, col_index + i)
                cell.color = color

        workbook.save()
        workbook.close()
        app.quit()
        app.kill()
        
        df_a_excel(ruta_out_f, 'Resumen TMs', capacidades_traspuesta, f_ini = 5, c_ini = 11)
        df_a_excel(ruta_out_f, 'Resumen TMs', base_merge[['MatriculaCalificador_x']], f_ini = 6, c_ini = 3)
        df_a_excel(ruta_out_f, 'Resumen TMs', base_merge[['NombresCalificador_x']], f_ini = 6, c_ini = 4)
        df_a_excel(ruta_out_f, 'Resumen TMs', base_merge[['RolCalificador_x']], f_ini = 6, c_ini = 5)
        df_a_excel(ruta_out_f, 'Resumen TMs', base_merge[['MatriculaCalificado']], f_ini = 6, c_ini = 6)
        df_a_excel(ruta_out_f, 'Resumen TMs', base_merge[['NombresCalificado_x']], f_ini = 6, c_ini = 7)
        df_a_excel(ruta_out_f, 'Resumen TMs', base_merge[['RolCalificado_x']], f_ini = 6, c_ini = 8)
        df_a_excel(ruta_out_f, 'Resumen TMs', gs_Calificado[['Grado Salarial']], f_ini = 6, c_ini = 9)
        df_a_excel(ruta_out_f, 'Resumen TMs', base_merge[['TipoEvaluacion_x']], f_ini = 6, c_ini = 10)

        # Abrir el archivo de Excel
        app = xw.App(visible=False)
        workbook = app.books.open(ruta_out_f)
        worksheet = workbook.sheets['Resumen TMs']

        rango1 = 'K6:AC133'
        valores_copia = worksheet.range(rango1).value
        worksheet.range(rango1).value = valores_copia

        rango2 = 'AE6:AI133'
        valores_copia2 = worksheet.range(rango2).value
        worksheet.range(rango2).value = valores_copia2
        
        workbook.save()
        workbook.close()
        app.quit()
        app.kill()

        cant_capa = capacidades_no_duplicada['DESCRIPCION2'].count()

        elimina_col_excel(ruta_out_f, 'Resumen TMs', cant_capa)
        elimina_filas_excel(ruta_out_f, 'Resumen TMs')

        resumen_tms_para_alerta = leer_excel_simple(ruta_out_f, 'Resumen TMs', f_inicio=5, c_inicio=2)
        resumen_tms_para_alerta = resumen_tms_para_alerta.drop(['Chapter', 'MatriculaCalificador', 'NombresCalificador', 'RolCalificador', 'NombresCalificado', 'RolCalificado', 
                                                                'GS Calificado', 'Promedio', 'Alerta', 'Comentario'], axis=1)
        duplicados = resumen_tms_para_alerta[resumen_tms_para_alerta['MatriculaCalificado'].duplicated(keep=False)]

        diferencias_list = []

        for index, row in duplicados.iterrows():
            duplicate_rows = resumen_tms_para_alerta[resumen_tms_para_alerta['MatriculaCalificado'] == row['MatriculaCalificado']]
            if len(duplicate_rows) == 2:
                first_row = duplicate_rows.iloc[0]
                second_row = duplicate_rows.iloc[1]

                diferencias = {'MatriculaCalificado': row['MatriculaCalificado']}
                for col in resumen_tms_para_alerta.columns:
                    if col != 'MatriculaCalificado':
                        if isinstance(first_row[col], (int, float)) and isinstance(second_row[col], (int, float)):
                            diferencias[col] = second_row[col] - first_row[col]
                        else:
                            diferencias[col] = None

                #diferencias = {'MatriculaCalificado': row['MatriculaCalificado'],
                #       **{col: first_row[col] - second_row[col] if isinstance(first_row[col], (int, float)) and isinstance(second_row[col], (int, float)) else None for col in resumen_tms_para_alerta.columns}}
                if diferencias:
                    diferencias_list.append(diferencias)

        diferencias_df = pd.DataFrame(diferencias_list)

        for col in diferencias_df.columns[1:]:
            if col == 'NivelDomainExpertise':
                diferencias_df[col] = diferencias_df[col].apply(apply_logic_DE)
            elif col == 'Análisis y solución de problemas':
                diferencias_df[col] = diferencias_df[col].apply(apply_logic_analisis)
            elif col == 'Liderazgo y Comunicación':
                diferencias_df[col] = diferencias_df[col].apply(apply_logic_liderazgo)
            elif col == 'Fit Cultural':
                diferencias_df[col] = diferencias_df[col].apply(apply_logic_fit)
            elif col == 'Nivel general':
                diferencias_df[col] = diferencias_df[col].apply(apply_logic_nivel)
            else:
                diferencias_df[col] = diferencias_df[col].apply(apply_logic_capacidades)

        #diferencias_df['Concatenadas'] = diferencias_df.iloc[:, 1:].apply(lambda row: ' '.join(map(str, row)), axis=1)
        diferencias_df['Concatenadas'] = diferencias_df.iloc[:, 1:].apply(lambda row: ', '.join(map(str, [val for val in row if val != ''])), axis=1)
        new_df = diferencias_df[['MatriculaCalificado', 'Concatenadas']]

        new_df_merge = pd.merge(base_merge, new_df, how='left', on='MatriculaCalificado')
        new_df_merge = new_df_merge.drop_duplicates(subset=['MatriculaCalificador_x', 'NombresCalificador_x', 'MatriculaCalificado', 'NombresCalificado_x', 'TipoEvaluacion_x'])
        #new_df_merge.to_excel('new_df_merge.xlsx')

        df_a_excel(ruta_out_f, 'Resumen TMs', new_df_merge[['Concatenadas']], f_ini = 6, c_ini = 34) #mejorar c_ini; falta borrar alertas de evaluación anterior

        #LLENADO RESUMEN LÍDERES
        resumen_tms = leer_excel_simple(ruta_out_f, 'Resumen TMs', f_inicio=5, c_inicio=2)
        cant_evaluados = base_merge['NombresCalificador_x'].value_counts().reset_index()
        cant_evaluados.columns = ['NombresCalificador_x', 'Cant_evaluados']
        cant_evaluados = cant_evaluados.sort_values(by='NombresCalificador_x')
        promedios = resumen_tms.groupby('NombresCalificador').mean()
        promedios2 = promedios.drop(['Promedio','Comentario','NivelDomainExpertise','Análisis y solución de problemas','Liderazgo y Comunicación','Fit Cultural','Nivel general'], axis=1)

        df_a_excel(ruta_out_f, 'Resumen Líderes', cant_evaluados[['NombresCalificador_x']], f_ini = 5, c_ini = 2)
        df_a_excel(ruta_out_f, 'Resumen Líderes', cant_evaluados[['Cant_evaluados']], f_ini = 5, c_ini = 3)
        df_a_excel_header(ruta_out_f, 'Resumen Líderes', promedios2, f_ini = 4, c_ini = 4)
        df_a_excel(ruta_out_f, 'Resumen Líderes', promedios[['NivelDomainExpertise']], f_ini = 5, c_ini = 24)

        elimina_col_excel_res_lid(ruta_out_f, 'Resumen Líderes', cant_capa)
        elimina_filas_excel_res_lid(ruta_out_f, 'Resumen Líderes')

        #falta: cambiar colores a los PRINCIPALES en resumen lideres

        #ACTUALIZAR EXCEL TABLAS DINAMICAS Y LISTAS
        xlapp = win32com.client.DispatchEx("Excel.Application")
        wb = xlapp.Workbooks.Open(ruta_out_f)
        xlapp.Visible = False
        wb.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        wb.Save()
        xlapp.Quit()

def apply_logic_capacidades(valor):
    if pd.isna(valor) or valor == 0 or valor == 1:
        return ''
    elif valor > 1:
        return 'Subió capacidad más de un nivel'
    else:
        return 'Bajó capacidad'

def apply_logic_DE(valor):
    if pd.isna(valor) or valor == 0 or valor == 1:
        return ''
    elif valor > 1:
        return 'Subió DE más de un nivel'
    else:
        return 'Bajó DE'
    
def apply_logic_analisis(valor):
    if pd.isna(valor) or valor == 0 or valor == 1:
        return ''
    elif valor > 1:
        return 'Subió Análisis y solución de problemas más de un nivel'
    else:
        return 'Bajó Análisis y solución de problemas'
    
def apply_logic_liderazgo(valor):
    if pd.isna(valor) or valor == 0 or valor == 1:
        return ''
    elif valor > 1:
        return 'Subió Liderazgo y Comunicación más de un nivel'
    else:
        return 'Bajó Liderazgo y Comunicación'
    
def apply_logic_fit(valor):
    if pd.isna(valor) or valor == 0 or valor == 1:
        return ''
    elif valor > 1:
        return 'Subió Fit Cultural más de un nivel'
    else:
        return 'Bajó Fit Cultural'
    
def apply_logic_nivel(valor):
    if pd.isna(valor) or valor == 0 or valor == 1:
        return ''
    elif valor > 1:
        return 'Subió Nivel general más de un nivel'
    else:
        return 'Bajó Nivel general'
    
def crear_csv(df,file_name):
    file_name = file_name
    df.to_csv(file_name,index=False,sep='|',encoding='UTF-16',header=False,quotechar='`',quoting=csv.QUOTE_NONNUMERIC)

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

    return results

if __name__ == '__main__':
    calibracion()