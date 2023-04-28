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

_conn_params = {
    "server": 'PUGINSQLP01',
    "database": 'BCP_GDH_PA_STAGE',
    "trusted_connection": "Yes",
    "driver": "{SQL Server}",
}

def calibracion():
    chapter = input(
        '''
        CALIBRACIÓN:
        --------------------
        Chapter:
        
        '''
    )
    ruta_ba = PATH_BA
    ruta_plantilla = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Entrada\PLANTILLA_RESULTADOS_POSTCALIBRACION.xlsx'
    ruta_BD = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Entrada\Base de datos.xlsx' #BORRAR DESPUES DE CONECTAR CON LA DB
    fec_hoy = datetime.today()
    fecha_hoy_format = fec_hoy.strftime('%Y%m%d')
    ruta_principal = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Salida'
    nombre_archivo = 'RESULTADOS_DATA_ENGINEER_{}_POSTCALIBRACION.xlsx'.format(fecha_hoy_format)  
    ruta_out_f = os.path.join(ruta_principal, nombre_archivo)

    base_activos = leer_excel_simple(ruta_ba, 'BD ACTIVOS')
    base_activos.rename(columns={'Matrícula': 'MATRICULA'}, inplace=True)
    base = leer_excel_simple(ruta_plantilla, 'BASE')
    resumen = leer_excel_simple(ruta_plantilla, 'Resumen TMs')
    resultados_capacidad_categoria = leer_excel_simple(ruta_BD, 'Hoja1') #BORRAR DESPUES DE CONECTAR CON LA DB
    resultados_resultado_comportamiento = leer_excel_simple(ruta_BD, 'Hoja2') #BORRAR DESPUES DE CONECTAR CON LA DB
    resultados_resultado_expertise = leer_excel_simple(ruta_BD, 'Hoja3') #BORRAR DESPUES DE CONECTAR CON LA DB


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
            N_NIVEL,
            NIVEL
        FROM BCP_GDH_PA_DW.PDIGITAL.F_RESULTADO_EXPERTISE
        WHERE FK_FECHA = @FECHA_EVA AND FK_CHAPTER=@PK_CHAPTER;
    '''.format(chapter)

    # resultados_capacidad_categoria = select(query1) #QUITAR COMENTARIO DESPUES DE CONECTAR CON LA DB
    # resultados_resultado_comportamiento = select(query2) #QUITAR COMENTARIO DESPUES DE CONECTAR CON LA DB
    # resultados_resultado_expertise = select(query3) #QUITAR COMENTARIO DESPUES DE CONECTAR CON LA DB

    resultados_resultado_comportamiento = resultados_resultado_comportamiento.drop(['ROL', 'MATRICULA_CALIFICADOR', 'ROL_CALIFICADOR'], axis=1)

    df_1 = pd.merge(resultados_resultado_comportamiento, resultados_capacidad_categoria, how='left', on='MATRICULA')
    df_2 = pd.merge(resultados_resultado_expertise, df_1, how='left', on='MATRICULA')
    # QUITAR EXCLUIDOS
    # df_2 = pd.merge(base_activos[['Matrícula']], df_2, how='left', on='MATRICULA')
    # df_2 = df_2.dropna(subset=['MATRICULA_CALIFICADOR'])
    df_resultado = pd.merge(df_2, base_activos[['MATRICULA', 'Correo electronico']], on='MATRICULA', how='left')

    gs_Calificado = df_resultado.drop_duplicates(subset="NOMBRES_CALIFICADO")
    gs_Calificado = pd.merge(gs_Calificado[['NOMBRES_CALIFICADO']], base_activos[['Nombre completo', 'Grado Salarial']], on='NOMBRES_CALIFICADO', how='left') #bug

    capacidades = df_resultado.drop(['MATRICULA_CALIFICADOR', 'NOMBRES_CALIFICADOR', 'ROL_CALIFICADOR', 'DESCRIPCION', 'MATRICULA', 'NOMBRES_CALIFICADO', 'Correo electronico', 
                                     'ROL', 'CATEGORIA_CAPACIDAD', 'SUBCATEGORIA_CAPACIDAD', 'N_NIVEL_y', 'NIVEL_y', 'FLAG_CONOCIMIENTO', 'N_NIVELDOMAINEXPERTISE', 
                                     'N_NIVELRESOL', 'N_NIVELLIDERAZG', 'N_NIVELCULTURAL', 'N_NIVEL_x', 'NIVEL_x'], axis=1)
    capacidades = capacidades.drop_duplicates(subset = "DESCRIPCION2").T
    print(gs_Calificado)

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
    df_a_excel(ruta_out_f, 'BASE', df_resultado[['N_NIVEL_y']], f_ini = 2, c_ini = 17)
    df_a_excel(ruta_out_f, 'BASE', df_resultado[['NIVEL_y']], f_ini = 2, c_ini = 18)
    df_a_excel(ruta_out_f, 'BASE', df_resultado[['FLAG_CONOCIMIENTO']], f_ini = 2, c_ini = 19)
    df_a_excel(ruta_out_f, 'BASE', df_resultado[['N_NIVELDOMAINEXPERTISE']], f_ini = 2, c_ini = 20)
    df_a_excel(ruta_out_f, 'BASE', df_resultado[['N_NIVELRESOL']], f_ini = 2, c_ini = 22)
    df_a_excel(ruta_out_f, 'BASE', df_resultado[['N_NIVELLIDERAZG']], f_ini = 2, c_ini = 24)
    df_a_excel(ruta_out_f, 'BASE', df_resultado[['N_NIVELCULTURAL']], f_ini = 2, c_ini = 26)
    df_a_excel(ruta_out_f, 'BASE', df_resultado[['N_NIVEL_x']], f_ini = 2, c_ini = 28)
    df_a_excel(ruta_out_f, 'Resumen TMs', capacidades, f_ini = 5, c_ini = 8)
    df_a_excel(ruta_out_f, 'Resumen TMs', gs_Calificado[['Grado Salarial']], f_ini = 4, c_ini = 7)

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