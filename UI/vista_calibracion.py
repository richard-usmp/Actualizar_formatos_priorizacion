import sys
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *
from __feature__ import true_property
from datetime import datetime
from os import path
import pandas as pd
import win32com
from base import apply_logic_DE, apply_logic_capacidades, apply_logic_dimension, copia_pega, df_a_excel, leer_excel_simple, elimina_col_excel, elimina_filas_excel, df_a_excel_header, elimina_col_excel_res_lid, elimina_filas_excel_res_lid
from enumeraciones import ETipoEva
from constantes import PATH_BA
import csv
import os
import xlwings as xw

class Window_calibracion(QMainWindow):
    def setupUI(self):
        #super().__init__()

        self.setWindowTitle("Calibración")
        self.setFixedSize(1280, 720)

        self.fr_calibracion = QFrame(self)
        self.fr_calibracion.geometry=QRect(64, 36, 1152, 648)
        self.fr_calibracion.styleSheet="background: white;"

        # self.label = QLabel(self.fr_calibracion)
        # self.label.text = "Selecciona los chapters:" 
        # self.label.geometry = QRect(0, 10, 500, 30)
        # self.label.styleSheet = "color: gray; font-size: 20px; font-weight: bold;"

        # self.grid_checktext = QFrame(self.fr_calibracion)
        # self.grid_checktext.geometry=QRect(0, 50, 700, 250)
        # self.grid_checktext.styleSheet="background: white;"

        # self.grid_layout = QGridLayout(self.grid_checktext)
        # self.selected_items = []
        # self.checkboxes = []
        # radio_chapter = [
        #     "FRAUD INVESTIGATION", "BACKEND JAVA", "DATA STEWARD", "ARQUITECTURA DE SEGURIDAD",
        #     "TESTING", "DEVSECOPS", "ARQUITECTURA DE NEGOCIOS", "SAP-SIGA",
        #     "DATA ENGINEER", "END USER - DISPOSITIVOS END USER", "PHYSICAL EXPERIENCE DESIGN",
        #     "GENESYS", "FRONTEND GROWTH SPECIALIST", "FRONTEND MOBILE", "DIGITAL COMMUNICATIONS"
        # ]

        # row = 0
        # col = 0
        # for text in radio_chapter:
        #     checkbox = QCheckBox(text)
        #     self.grid_layout.addWidget(checkbox, row, col)
        #     self.checkboxes.append(checkbox)
        #     col += 1
        #     if col == 3:
        #         col = 0
        #         row += 1

        self.boton_importar_BD = QPushButton(self.fr_calibracion)
        self.boton_importar_BD.text = "Importar Excel BD"
        self.boton_importar_BD.clicked.connect(self.abrir_archivo_excel_BD)
        self.boton_importar_BD.geometry = QRect(100, 440, 235, 23)
        
        self.boton_ingresar_plantilla = QPushButton(self.fr_calibracion)
        self.boton_ingresar_plantilla.text = "Importar plantilla de calibración"
        self.boton_ingresar_plantilla.clicked.connect(self.abrir_archivo_excel)
        self.boton_ingresar_plantilla.geometry = QRect(100, 500, 235, 23)

        self.submit_button = QPushButton(self.fr_calibracion)
        self.submit_button.text = "Generar reporte de calibración"
        self.submit_button.clicked.connect(self.on_submit)
        self.submit_button.geometry=QRect(100, 560, 235, 23)

        self.boton_priorizacion = QPushButton(self.fr_calibracion)
        self.boton_priorizacion.text = "Ir a priorización"
        self.boton_priorizacion.geometry=QRect(650, 500, 205, 23)
        self.boton_priorizacion.styleSheet = "background: #33E9FF;"

        self.boton_avance = QPushButton(self.fr_calibracion)
        self.boton_avance.text = "Ir a avance medición"
        self.boton_avance.geometry=QRect(650, 600, 205, 23)
        self.boton_avance.styleSheet = "background: #33E9FF;"

    def on_submit(self):
        options = QFileDialog.Options()
        ruta_principal = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta", options=options)
        #self.selected_items = [checkbox.text for checkbox in self.checkboxes if str(checkbox.checkState()) == 'CheckState.Checked']

        ruta_ba = PATH_BA
        fec_hoy = datetime.today()
        fecha_hoy_format = fec_hoy.strftime('%Y%m%d')
        base_activos = leer_excel_simple(ruta_ba, 'BD ACTIVOS')
        base_activos.rename(columns={'Matrícula': 'MATRICULA'}, inplace=True)
        resultados_capacidad_categoria = leer_excel_simple(ruta_BD, 'Hoja1')
        resultados_resultado_comportamiento = leer_excel_simple(ruta_BD, 'Hoja2')
        resultados_resultado_expertise = leer_excel_simple(ruta_BD, 'Hoja3')

        chapters1 = resultados_capacidad_categoria['DESCRIPCION'].unique()
        for chapter in chapters1:
            print(chapter)
            nombre_archivo = 'RESULTADOS_{}_{}_PRECALIBRACION.xlsx'.format(chapter, fecha_hoy_format)
            ruta_out_f = os.path.join(ruta_principal, nombre_archivo)
            resultados_capacidad_categoria_chapter = resultados_capacidad_categoria[resultados_capacidad_categoria['DESCRIPCION'] == chapter]
            resultados_resultado_comportamiento_chapter = resultados_resultado_comportamiento[resultados_resultado_comportamiento['DESCRIPCION'] == chapter]
            resultados_resultado_expertise_chapter = resultados_resultado_expertise[resultados_resultado_expertise['DESCRIPCION'] == chapter]

            resultados_resultado_comportamiento_chapter = resultados_resultado_comportamiento_chapter.drop(['ROL', 'ROL_CALIFICADOR'], axis=1)
            resultados_capacidad_categoria_chapter['CONCAT'] = resultados_capacidad_categoria_chapter['MATRICULA'] + resultados_capacidad_categoria_chapter['MATRICULA_CALIFICADOR'] + resultados_capacidad_categoria_chapter['FECHA'].astype(str)
            resultados_resultado_comportamiento_chapter['CONCAT'] = resultados_resultado_comportamiento_chapter['MATRICULA'] + resultados_resultado_comportamiento_chapter['MATRICULA_CALIFICADOR'] + resultados_resultado_comportamiento_chapter['FECHA'].astype(str)
            resultados_resultado_expertise_chapter['CONCAT'] = resultados_resultado_expertise_chapter['MATRICULA'] + resultados_resultado_expertise_chapter['MATRICULA_CALIFICADOR'] + resultados_resultado_expertise_chapter['FECHA'].astype(str)

            df_1 = pd.merge(resultados_resultado_comportamiento_chapter, resultados_capacidad_categoria_chapter, how='left', on='CONCAT')
            df_2 = pd.merge(resultados_resultado_expertise_chapter, df_1, how='left', on='CONCAT')

            df_resultado = pd.merge(df_2, base_activos[['MATRICULA', 'Correo electronico']], on='MATRICULA', how='left')
            
            copia_pega(ruta_plantilla, ruta_out_f)
            df1 = df_resultado[['MATRICULA_CALIFICADOR', 'NOMBRES_CALIFICADOR', 'ROL_CALIFICADOR', 'DESCRIPCION', 'MATRICULA', 'NOMBRES_CALIFICADO', 'Correo electronico']]
            df2 = df_resultado[['ROL', 'DESCRIPCION2', 'CATEGORIA_CAPACIDAD', 'SUBCATEGORIA_CAPACIDAD']]
            df3 = df_resultado[['N_NIVEL_y', 'NIVEL_y', 'FLAG_CONOCIMIENTO', 'N_NIVELDOMAINEXPERTISE']]
            df4 = df_resultado[['N_NIVELRESOL']]
            df5 = df_resultado[['N_NIVELLIDERAZG']]
            df6 = df_resultado[['N_NIVELCULTURAL']]
            df7 = df_resultado[['N_NIVEL_x']]
            df8 = df_resultado[['EVALUACION']]

            df_a_excel(ruta_out_f, 'BASE', df1, f_ini = 2, c_ini = 1)
            df_a_excel(ruta_out_f, 'BASE', df2, f_ini = 2, c_ini = 12)
            df_a_excel(ruta_out_f, 'BASE', df3, f_ini = 2, c_ini = 18)
            df_a_excel(ruta_out_f, 'BASE', df4, f_ini = 2, c_ini = 23)
            df_a_excel(ruta_out_f, 'BASE', df5, f_ini = 2, c_ini = 25)
            df_a_excel(ruta_out_f, 'BASE', df6, f_ini = 2, c_ini = 27)
            df_a_excel(ruta_out_f, 'BASE', df7, f_ini = 2, c_ini = 29)
            df_a_excel(ruta_out_f, 'BASE', df8, f_ini = 2, c_ini = 31)

            #LLENADO DE CUADRO RESUMEN TMs
            base = leer_excel_simple(ruta_out_f, 'BASE')
            base1 = base.drop_duplicates(subset=['MatriculaCalificador', 'NombresCalificador', 'MatriculaCalificado', 'NombresCalificado', 'TipoEvaluacion'])
            base_solo_eva = base1[base1['TipoEvaluacion'] == 'EVALUACION']

            base_merge = pd.merge(base1, base_solo_eva, how='left', on='MatriculaCalificado')
            #base_merge.to_excel('base_merge{}.xlsx'.format(chapter))

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
            df_a_excel(ruta_out_f, 'Resumen TMs', base_merge[['MatriculaCalificador_y','NombresCalificador_y', 'RolCalificador_x', 'MatriculaCalificado', 'NombresCalificado_x', 'RolCalificado_x']], f_ini = 6, c_ini = 3)
            df_a_excel(ruta_out_f, 'Resumen TMs', gs_Calificado[['Grado Salarial']], f_ini = 6, c_ini = 9)
            df_a_excel(ruta_out_f, 'Resumen TMs', base_merge[['TipoEvaluacion_x']], f_ini = 6, c_ini = 10)


            # Abrir el archivo de Excel
            app = xw.App(visible=False)
            workbook = app.books.open(ruta_out_f)
            worksheet = workbook.sheets['Resumen TMs']

            rango1 = 'K6:AH1000'
            valores_copia = worksheet.range(rango1).value
            worksheet.range(rango1).value = valores_copia

            rango2 = 'AJ6:AN1000'
            valores_copia2 = worksheet.range(rango2).value
            worksheet.range(rango2).value = valores_copia2
            
            workbook.save()
            workbook.close()
            app.quit()
            app.kill()

            cant_capa = capacidades_no_duplicada['DESCRIPCION2'].count()
            print('cant_capa')
            print(cant_capa)

            elimina_col_excel(ruta_out_f, 'Resumen TMs', cant_capa)
            elimina_filas_excel(ruta_out_f, 'Resumen TMs')
            
            #ALERTAS
            resumen_tms_para_alerta = leer_excel_simple(ruta_out_f, 'Resumen TMs', f_inicio=5, c_inicio=2)
            columns_to_drop = ['Chapter', 'MatriculaCalificador', 'NombresCalificador', 'RolCalificador', 'NombresCalificado', 
                            'RolCalificado', 'GS Calificado', 'Promedio', 'Alerta', 'Comentario']
            for columna in columns_to_drop:
                if columna in resumen_tms_para_alerta.columns:
                    resumen_tms_para_alerta = resumen_tms_para_alerta.drop(columna, axis=1)
                else:
                    print('Error: No existe la columna {} en la hoja Resumen TM.'.format(columna))

            duplicados = resumen_tms_para_alerta[
                (resumen_tms_para_alerta['TipoEvaluacion'] != 'AUTOEVALUACION') &
                resumen_tms_para_alerta['MatriculaCalificado'].duplicated(keep=False)
            ]

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

                        if diferencias:
                            diferencias_list.append(diferencias)

                diferencias_df = pd.DataFrame(diferencias_list)
                diferencias_df_new = diferencias_df.copy()

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
                        print('TO DO')
                    else:
                        diferencias_df[col] = diferencias_df[col].apply(apply_logic_capacidades)

                for fila in range(len(diferencias_df_new.index)):
                    cont=0
                    for columnas in range(2,cant_capa):
                        if diferencias_df_new.iloc[fila, columnas] > 0: cont += 1
                    if cont >= 4: diferencias_df_new.at[fila, 'Concatenadas_2'] = 'Subió 4 o más capacidades'

                diferencias_df['Concatenadas'] = diferencias_df.iloc[:, 2:].apply(lambda row: '' if all(val == '' for val in row) else ', '.join(set(val for val in row if val != '')), axis=1)
                new_df = diferencias_df[['MatriculaCalificado', 'TipoEvaluacion', 'Concatenadas']]
                new_df = pd.merge(new_df, diferencias_df_new, how='left', on=['MatriculaCalificado', 'TipoEvaluacion'])
                new_df['Concatenadas_final'] = new_df['Concatenadas'].fillna('') + (', ' + new_df['Concatenadas_2']).fillna('')
                new_df['Concatenadas_final'] = new_df['Concatenadas_final'].apply(lambda x: x.lstrip(', ')) #si sale algo mal en alertas, comentar esta fila

                new_df_merge = pd.merge(base_merge, new_df, how='left', on='MatriculaCalificado')
                new_df_merge = new_df_merge.drop_duplicates(subset=['MatriculaCalificador_x', 'NombresCalificador_x', 'MatriculaCalificado', 'NombresCalificado_x', 'TipoEvaluacion_x'])
                new_df_merge.loc[new_df_merge['TipoEvaluacion_x'] == 'AUTOEVALUACION', 'Concatenadas'] = ''

                df_a_excel(ruta_out_f, 'Resumen TMs', new_df_merge[['Concatenadas_final']], f_ini = 6, c_ini = (17+cant_capa))

                app = xw.App(visible=False)
                workbook = app.books.open(ruta_out_f)
                worksheet = workbook.sheets['Resumen TMs']

                starting_cell = 'B6'
                row_index_eva = worksheet.range(starting_cell).row
                col_index_eva = worksheet.range(starting_cell).column

                for i, tipo_eva in enumerate(new_df_merge['TipoEvaluacion_x']):
                    if tipo_eva == 'EVALUACION ANTERIOR':
                        cell = worksheet.range(row_index_eva + i, col_index_eva + 15 + cant_capa)
                        cell.font.color = (255, 255, 255)

                workbook.save()
                workbook.close()
                app.quit()
                app.kill()

            #LLENADO RESUMEN LÍDERES
            resumen_tms = leer_excel_simple(ruta_out_f, 'Resumen TMs', f_inicio=5, c_inicio=2)
            resumen_tms = resumen_tms[(resumen_tms['TipoEvaluacion'] != 'EVALUACION ANTERIOR') & (resumen_tms['TipoEvaluacion'] != 'AUTOEVALUACION')]
            base_merge = base_merge[(base_merge['TipoEvaluacion_x'] != 'EVALUACION ANTERIOR') & (base_merge['TipoEvaluacion_x'] != 'AUTOEVALUACION')]
            cant_evaluados = base_merge['NombresCalificador_y'].value_counts().reset_index()
            cant_evaluados.columns = ['NombresCalificador_y', 'Cant_evaluados']
            cant_evaluados = cant_evaluados.sort_values(by='NombresCalificador_y')
            promedios = resumen_tms.groupby('NombresCalificador').mean()
            promedios2 = promedios.drop(['Promedio','Comentario','NivelDomainExpertise','Análisis y solución de problemas','Liderazgo y Comunicación','Fit Cultural','Nivel general'], axis=1)

            df_a_excel(ruta_out_f, 'Resumen Líderes', cant_evaluados[['NombresCalificador_y']], f_ini = 5, c_ini = 2)
            df_a_excel(ruta_out_f, 'Resumen Líderes', cant_evaluados[['Cant_evaluados']], f_ini = 5, c_ini = 3)
            df_a_excel_header(ruta_out_f, 'Resumen Líderes', promedios2, f_ini = 4, c_ini = 4)
            df_a_excel(ruta_out_f, 'Resumen Líderes', promedios[['NivelDomainExpertise']], f_ini = 5, c_ini = 28)

            elimina_col_excel_res_lid(ruta_out_f, 'Resumen Líderes', cant_capa)
            elimina_filas_excel_res_lid(ruta_out_f, 'Resumen Líderes')
            
            app_lid = xw.App(visible=False)
            workbook_lid = app_lid.books.open(ruta_out_f)
            worksheet_lid = workbook_lid.sheets['Resumen Líderes']

            color_lid = (255, 165, 0)
            starting_cell_lid = 'D4'
            row_index_lid = worksheet_lid.range(starting_cell_lid).row
            col_index_lid = worksheet_lid.range(starting_cell_lid).column
            capacidades_principal_drop = capacidades_principal.drop_duplicates(subset='DESCRIPCION2')
            for i in range(len(capacidades_principal_drop.index)):
                cell_lid = worksheet_lid.cells(row_index_lid, col_index_lid + i)
                cell_lid.color = color_lid

            workbook_lid.save()
            workbook_lid.close()
            app_lid.quit()
            app_lid.kill()

            #ACTUALIZAR EXCEL TABLAS DINAMICAS Y LISTAS
            xlapp = win32com.client.DispatchEx("Excel.Application")
            wb = xlapp.Workbooks.Open(ruta_out_f)
            xlapp.Visible = False
            wb.RefreshAll()
            xlapp.CalculateUntilAsyncQueriesDone()
            wb.Save()
            xlapp.Quit()

        #LLENADO RESUMEN LÍDERES
        resumen_tms = leer_excel_simple(ruta_out_f, 'Resumen TMs', f_inicio=5, c_inicio=2)
        resumen_tms = resumen_tms[(resumen_tms['TipoEvaluacion'] != 'EVALUACION ANTERIOR') & (resumen_tms['TipoEvaluacion'] != 'AUTOEVALUACION')]
        base_merge = base_merge[(base_merge['TipoEvaluacion_x'] != 'EVALUACION ANTERIOR') & (base_merge['TipoEvaluacion_x'] != 'AUTOEVALUACION')]
        cant_evaluados = base_merge['NombresCalificador_y'].value_counts().reset_index()
        cant_evaluados.columns = ['NombresCalificador_y', 'Cant_evaluados']
        cant_evaluados = cant_evaluados.sort_values(by='NombresCalificador_y')
        promedios = resumen_tms.groupby('NombresCalificador').mean()
        promedios2 = promedios.drop(['Promedio','Comentario','NivelDomainExpertise','Análisis y solución de problemas','Liderazgo y Comunicación','Fit Cultural','Nivel general'], axis=1)

        df_a_excel(ruta_out_f, 'Resumen Líderes', cant_evaluados[['NombresCalificador_y']], f_ini = 5, c_ini = 2)
        df_a_excel(ruta_out_f, 'Resumen Líderes', cant_evaluados[['Cant_evaluados']], f_ini = 5, c_ini = 3)
        df_a_excel_header(ruta_out_f, 'Resumen Líderes', promedios2, f_ini = 4, c_ini = 4)
        df_a_excel(ruta_out_f, 'Resumen Líderes', promedios[['NivelDomainExpertise']], f_ini = 5, c_ini = 28)

        elimina_col_excel_res_lid(ruta_out_f, 'Resumen Líderes', cant_capa)
        elimina_filas_excel_res_lid(ruta_out_f, 'Resumen Líderes')
        
        app_lid = xw.App(visible=False)
        workbook_lid = app_lid.books.open(ruta_out_f)
        worksheet_lid = workbook_lid.sheets['Resumen Líderes']

        color_lid = (255, 165, 0)
        starting_cell_lid = 'D4'
        row_index_lid = worksheet_lid.range(starting_cell_lid).row
        col_index_lid = worksheet_lid.range(starting_cell_lid).column
        capacidades_principal_drop = capacidades_principal.drop_duplicates(subset='DESCRIPCION2')
        for i in range(len(capacidades_principal_drop.index)):
            cell_lid = worksheet_lid.cells(row_index_lid, col_index_lid + i)
            cell_lid.color = color_lid

        workbook_lid.save()
        workbook_lid.close()
        app_lid.quit()
        app_lid.kill()

        #ACTUALIZAR EXCEL TABLAS DINAMICAS Y LISTAS
        xlapp = win32com.client.DispatchEx("Excel.Application")
        wb = xlapp.Workbooks.Open(ruta_out_f)
        xlapp.Visible = False
        wb.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        wb.Save()
        xlapp.Quit()

    def abrir_archivo_excel(self):
        global ruta_plantilla
        ruta_plantilla = QFileDialog.getOpenFileName(self, 'Abrir archivo', 'C:\\', 'Excel (*.xls *.xlsx)')
        ruta_plantilla = ruta_plantilla[0]
    
    def abrir_archivo_excel_BD(self):
        global ruta_BD
        ruta_BD = QFileDialog.getOpenFileName(self, 'Abrir archivo', 'C:\\', 'Excel (*.xls *.xlsx)')
        ruta_BD = ruta_BD[0]