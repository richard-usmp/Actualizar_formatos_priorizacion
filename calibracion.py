from datetime import date, datetime
from getpass import getuser
import pyodbc
import pandas as pd,csv
from concurrent import futures
import csv
import os
import win32com
from base import copia_pega, df_a_excel, leer_excel_simple, leer_ba
from constantes import PATH_BA
import xlwings as xw

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
        nombre_archivo = 'RESULTADOS_{}_{}_POSTCALIBRACION.xlsx'.format(chapter, fecha_hoy_format)
        ruta_out_f = os.path.join(ruta_principal, nombre_archivo)
        resultados_capacidad_categoria_chapter = resultados_capacidad_categoria[resultados_capacidad_categoria['DESCRIPCION'] == chapter]
        resultados_resultado_comportamiento_chapter = resultados_resultado_comportamiento[resultados_resultado_comportamiento['DESCRIPCION'] == chapter]
        resultados_resultado_expertise_chapter = resultados_resultado_expertise[resultados_resultado_expertise['DESCRIPCION'] == chapter]

        query1 = '''
            DECLARE	@CHAPTER varchar(100) = '{}'

            DECLARE @PK_CHAPTER INT = (
            SELECT PK_CHAPTER FROM BCP_GDH_PA_DW.PDIGITAL.D_CHAPTER WHERE DESCRIPCION = @CHAPTER)

            DECLARE @FECHA_EVA INT = (
            SELECT MAX(FK_FECHA) FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE WHERE FK_CHAPTER=@PK_CHAPTER)


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
                FR.FLAG_CONOCIMIENTO
            FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_CAPACIDAD FR
                LEFT JOIN BCP_GDH_PA_DW.PDIGITAL.D_CHAPTER DP ON FR.FK_CHAPTER=DP.PK_CHAPTER
                LEFT JOIN BCP_GDH_PA_DW.PDIGITAL.D_CAPACIDAD DC ON FR.FK_CAPACIDAD=DC.PK_CAPACIDAD
                LEFT JOIN TCC TC ON FR.FK_CAPACIDAD=TC.FK_CAPACIDAD AND
                                    FR.FK_CHAPTER=TC.FK_CHAPTER AND
                                    FR.ROL=TC.ROL
                LEFT JOIN BCP_GDH_DW_UGI.UGI.COLABORADORES LID ON FR.MATRICULA_CALIFICADOR=RIGHT(LID.Matricula,6)
                LEFT JOIN BCP_GDH_DW_UGI.UGI.COLABORADORES CAL ON FR.MATRICULA=RIGHT(CAL.Matricula,6)
            WHERE FK_FECHA = @FECHA_EVA AND FR.FK_CHAPTER=@PK_CHAPTER;
        '''.format(chapter)
        query2 = '''
            DECLARE	@CHAPTER varchar(100) = '{}'

            DECLARE @PK_CHAPTER INT = (
            SELECT PK_CHAPTER FROM BCP_GDH_PA_DW.PDIGITAL.D_CHAPTER WHERE DESCRIPCION = @CHAPTER)

            DECLARE @FECHA_EVA INT = (
            SELECT MAX(FK_FECHA) FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE WHERE FK_CHAPTER=@PK_CHAPTER)

            ;WITH T_COMPORTAMIENTO AS
            (
            SELECT FR.MATRICULA,
                FR.ROL,
                FR.MATRICULA_CALIFICADOR,
                FR.ROL_CALIFICADOR,
                DC.DESCRIPCION AS COMPORTAMIENTO,
                N_NIVEL
            FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_COMPORTAMIENTO FR
                LEFT JOIN BCP_GDH_PA_DW.PDIGITAL.D_COMPORTAMIENTO DC ON FR.FK_COMPORTAMIENTO=DC.PK_COMPORTAMIENTO
                LEFT JOIN BCP_GDH_PA_DW.PDIGITAL.D_CHAPTER DP ON FR.FK_CHAPTER=DP.PK_CHAPTER
            WHERE FK_FECHA = @FECHA_EVA AND FR.FK_CHAPTER=@PK_CHAPTER
            )
            SELECT MATRICULA,
                ROL,
                MATRICULA_CALIFICADOR,
                ROL_CALIFICADOR,
				DESCRIPCION,
                [Domain expertise] AS N_NIVELDOMAINEXPERTISE,
                [Análisis y solución de problemas] AS N_NIVELRESOL,
                [Liderazgo y comunicación] AS N_NIVELLIDERAZG,
                [Fit cultural] AS N_NIVELCULTURAL
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

            SELECT MATRICULA,
                MATRICULA_CALIFICADOR,
				DP.DESCRIPCION,
                N_NIVEL,
                NIVEL
            FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE RE
				LEFT JOIN BCP_GDH_PA_DW.PDIGITAL.D_CHAPTER DP ON RE.FK_CHAPTER=DP.PK_CHAPTER
            WHERE FK_FECHA = @FECHA_EVA AND FK_CHAPTER=@PK_CHAPTER;
        '''.format(chapter)

        resultados_resultado_comportamiento_chapter = resultados_resultado_comportamiento_chapter.drop(['ROL', 'ROL_CALIFICADOR'], axis=1)
        resultados_capacidad_categoria_chapter['CONCAT'] = resultados_capacidad_categoria_chapter.MATRICULA.str.cat(resultados_capacidad_categoria_chapter.MATRICULA_CALIFICADOR, sep='')
        resultados_resultado_comportamiento_chapter['CONCAT'] = resultados_resultado_comportamiento_chapter.MATRICULA.str.cat(resultados_resultado_comportamiento_chapter.MATRICULA_CALIFICADOR, sep='')
        resultados_resultado_expertise_chapter['CONCAT'] = resultados_resultado_expertise_chapter.MATRICULA.str.cat(resultados_resultado_expertise_chapter.MATRICULA_CALIFICADOR, sep='')

        df_1 = pd.merge(resultados_resultado_comportamiento_chapter, resultados_capacidad_categoria_chapter, how='left', on='CONCAT')
        df_2 = pd.merge(resultados_resultado_expertise_chapter, df_1, how='left', on='CONCAT')

        df_resultado = pd.merge(df_2, base_activos[['MATRICULA', 'Correo electronico']], on='MATRICULA', how='left')

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

        #LLENADO DE CUADRO RESUMEN TMs
        resumen_tms = df_resultado.drop_duplicates(subset=['MATRICULA_CALIFICADOR', 'NOMBRES_CALIFICADOR', 'ROL_CALIFICADOR', 'MATRICULA', 'NOMBRES_CALIFICADO', 'ROL'])
        base = leer_excel_simple(ruta_out_f, 'BASE')
        base1 = base.drop_duplicates(subset=['MatriculaCalificador', 'NombresCalificador', 'MatriculaCalificado', 'NombresCalificado', 'TipoEvaluacion'])
        base_solo_eva = base1[base1['TipoEvaluacion'] != 'AUTOEVALUACIÓN']

        base_merge = pd.merge(base_solo_eva, base1, how='left', on='MatriculaCalificado')

        #GS Calificado cuadro Resumen TMs
        base_activos.rename(columns={'Nombre completo': 'NombresCalificado_x'}, inplace=True)
        gs_Calificado = pd.merge(base_merge[['NombresCalificado_x']], base_activos[['NombresCalificado_x', 'Grado Salarial']], on='NombresCalificado_x', how='left')

        # Abrir el archivo de Excel
        app = xw.App(visible=False)
        workbook = app.books.open(ruta_out_f)
        worksheet = workbook.sheets['Resumen TMs'] 

        #capacidad cuadro Resumen TMs
        orange_color = (255, 165, 0)
        starting_cell = 'K5'
        row_index = worksheet.range(starting_cell).row
        col_index = worksheet.range(starting_cell).column
        capacidades = df_resultado.drop(['MATRICULA', 'MATRICULA_CALIFICADOR', 'N_NIVEL_x', 'NIVEL_x', 'CONCAT', 'MATRICULA_x', 'MATRICULA_CALIFICADOR_x', 'Correo electronico', 
                                        'N_NIVELDOMAINEXPERTISE', 'N_NIVELRESOL', 'N_NIVELLIDERAZG', 'N_NIVELCULTURAL', 'MATRICULA_CALIFICADOR_y', 'NOMBRES_CALIFICADOR', 
                                        'ROL_CALIFICADOR', 'DESCRIPCION', 'MATRICULA_y', 'NOMBRES_CALIFICADO', 'ROL', 'CATEGORIA_CAPACIDAD', 'N_NIVEL_y', 'NIVEL_y', 'FLAG_CONOCIMIENTO'], axis=1)
        capacidades = capacidades.drop_duplicates(subset = ['DESCRIPCION2', 'SUBCATEGORIA_CAPACIDAD'])
        capacidades_principal = capacidades[capacidades['SUBCATEGORIA_CAPACIDAD'] == 'PRINCIPAL']
        capacidades_no_duplicada = capacidades.drop_duplicates(subset = 'DESCRIPCION2')
        capacidades_traspuesta = capacidades_no_duplicada[['DESCRIPCION2']].T
        for i, capa in enumerate(capacidades_no_duplicada['DESCRIPCION2']):
            if capa in capacidades_principal[['DESCRIPCION2']].values:
                cell = worksheet.cells(row_index, col_index + i)
                cell.color = orange_color
        
        workbook.save()
        workbook.close()
        app.quit()    
        
        df_a_excel(ruta_out_f, 'Resumen TMs', capacidades_traspuesta, f_ini = 5, c_ini = 11)
        df_a_excel(ruta_out_f, 'Resumen TMs', base_merge[['MatriculaCalificador_x']], f_ini = 6, c_ini = 3)
        df_a_excel(ruta_out_f, 'Resumen TMs', base_merge[['NombresCalificador_x']], f_ini = 6, c_ini = 4)
        df_a_excel(ruta_out_f, 'Resumen TMs', base_merge[['RolCalificador_x']], f_ini = 6, c_ini = 5)
        df_a_excel(ruta_out_f, 'Resumen TMs', base_merge[['MatriculaCalificado']], f_ini = 6, c_ini = 6)
        df_a_excel(ruta_out_f, 'Resumen TMs', base_merge[['NombresCalificado_x']], f_ini = 6, c_ini = 7)
        df_a_excel(ruta_out_f, 'Resumen TMs', base_merge[['RolCalificado_x']], f_ini = 6, c_ini = 8)
        df_a_excel(ruta_out_f, 'Resumen TMs', gs_Calificado[['Grado Salarial']], f_ini = 6, c_ini = 9)
        df_a_excel(ruta_out_f, 'Resumen TMs', base_merge[['TipoEvaluacion_y']], f_ini = 6, c_ini = 10)

        #ACTUALIZAR EXCEL TABLAS DINAMICAS Y LISTAS
        xlapp = win32com.client.DispatchEx("Excel.Application")
        wb = xlapp.Workbooks.Open(ruta_out_f)
        xlapp.Visible = True
        wb.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        wb.Save()
        xlapp.Quit()

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