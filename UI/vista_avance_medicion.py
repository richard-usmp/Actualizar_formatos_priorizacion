import sys
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *
from __feature__ import true_property
from datetime import datetime
from os import path
import pandas as pd
import win32com
from base import copia_pega, df_a_excel, leer_excel_simple
from enumeraciones import ETipoEva
from constantes import PATH_BA

formatted_date = ''

class Window_avance_medicion(QMainWindow):
    def setupUI(self):
        #super().__init__()

        self.setWindowTitle("Avance de medición")
        self.setFixedSize(1280, 720)

        self.fr_avance_medicion = QFrame(self)
        self.fr_avance_medicion.geometry=QRect(64, 36, 1152, 648)
        self.fr_avance_medicion.styleSheet="background: white;"

        self.label = QLabel(self.fr_avance_medicion)
        self.label.text = "Ingresa el nombre del chapter:" 
        self.label.geometry = QRect(0, 10, 500, 30)
        self.label.styleSheet = "color: gray; font-size: 20px; font-weight: bold;"

        self.chapter_input = QLineEdit(self.fr_avance_medicion, placeholderText ="Chapter")
        self.chapter_input.geometry=QRect(20, 60, 360, 40)

        #calendario
        self.label_calendario = QLabel(self.fr_avance_medicion)
        self.label_calendario.text = "Selecciona una fecha: (opcional)" 
        self.label_calendario.geometry = QRect(0, 100, 500, 30)
        self.label_calendario.styleSheet = "color: gray; font-size: 20px; font-weight: bold;"
        self.calendar = QCalendarWidget(self.fr_avance_medicion)
        self.calendar.selectionChanged.connect(self.on_date_selected)
        self.calendar.geometry = QRect(20, 150, 400, 300)
        self.calendar.styleSheet = "background: gray;"
        
        self.selected_date_label = QLabel(self.fr_avance_medicion)
        self.selected_date_label.geometry = QRect(50, 460, 500, 30)

        self.label_opciones = QLabel(self.fr_avance_medicion)
        self.label_opciones.text = "Opciones de reporte:" 
        self.label_opciones.geometry = QRect(600, 200, 400, 30)
        self.label_opciones.styleSheet = "color: gray; font-size: 20px; font-weight: bold;"

        self.combo_box = QComboBox(self.fr_avance_medicion)
        self.combo_box.geometry = QRect(620, 260, 400, 30)
        self.combo_box.addItem("1. FLAG_EVALUACION y PO")
        self.combo_box.addItem("2. FLAG_AUTOEVALUACION y PO")
        self.combo_box.addItem("3. FLAG_EVALUACION y otros")
        self.combo_box.addItem("4. FLAG_AUTOEVALUACION y otros")
        self.combo_box.addItem("5. FLAG_EVALUACION, FLAG_AUTOEVALUACION y otros")

        self.boton_ingresar_plantilla = QPushButton(self.fr_avance_medicion)
        self.boton_ingresar_plantilla.text = "Importar plantilla de avance de medición"
        self.boton_ingresar_plantilla.clicked.connect(self.abrir_archivo_excel)
        self.boton_ingresar_plantilla.geometry = QRect(100, 500, 235, 23)

        self.submit_button = QPushButton(self.fr_avance_medicion)
        self.submit_button.text = "Generar reporte de avance de medición"
        self.submit_button.clicked.connect(self.on_submit)
        self.submit_button.geometry=QRect(100, 560, 235, 23)

        self.boton_priorizacion = QPushButton(self.fr_avance_medicion)
        self.boton_priorizacion.text = "Ir a priorización"
        self.boton_priorizacion.geometry=QRect(650, 500, 205, 23)
        self.boton_priorizacion.styleSheet = "background: #33E9FF;"

        self.boton_calibracion = QPushButton(self.fr_avance_medicion)
        self.boton_calibracion.text = "Ir a calibración"
        self.boton_calibracion.geometry=QRect(650, 600, 205, 23)
        self.boton_calibracion.styleSheet = "background: #33E9FF;"

    def on_submit(self):
        options = QFileDialog.Options()
        ruta_principal = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta", options=options)
        chapter = self.chapter_input.text

        selected_option = self.combo_box.currentIndex + 1
        dic = {
            '1':ETipoEva.FLAG_EVALUACION_y_PO,
            '2':ETipoEva.FLAG_AUTOEVALUACION_y_PO,
            '3':ETipoEva.FLAG_EVALUACION_y_otros,
            '4':ETipoEva.FLAG_AUTOEVALUACION_y_otros,
            '5':ETipoEva.AUTOEVALUACION_Y_EVALUACION_y_otros
        }
        en = dic.get(str(selected_option))

        xlapp = win32com.client.DispatchEx("Excel.Application")
        wb = xlapp.Workbooks.Open(ruta_plantilla)
        xlapp.Visible = True
        wb.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        wb.Save()
        xlapp.Quit()
        fec_hoy = datetime.today()
        fecha_hoy_format = fec_hoy.strftime('%Y%m%d')
        #ruta_principal = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Salida'
        nombre_archivo = 'AVANCE_MEDICION_{}_{}.xlsx'.format(chapter, fecha_hoy_format)
        ruta_out_f = path.join(ruta_principal, nombre_archivo)
        ruta_ba = PATH_BA

        base = leer_excel_simple(ruta_plantilla, 'BASE')
        lt_out_comportamiento = leer_excel_simple(ruta_plantilla, 'LT_OUT_COMPORTAMIENTO')
        lt_out_capacidad = leer_excel_simple(ruta_plantilla, 'LT_OUT_CAPACIDAD')
        base_activos = leer_excel_simple(ruta_ba, 'BD ACTIVOS')
        base_activos.rename(columns={'Matrícula': 'MATRICULA'}, inplace=True)

        #filtros por fecha
        # if formatted_date != '':
        #     lt_out_comportamiento['Created'] = pd.to_datetime(lt_out_comportamiento['Created'], format='%d/%m/%Y')
        #     lt_out_capacidad['Created'] = pd.to_datetime(lt_out_capacidad['Created'], format='%d/%m/%Y')

        #     fecha_filtro_format = pd.to_datetime(formatted_date, format='%d/%m/%Y')

        #     lt_out_comportamiento = lt_out_comportamiento[lt_out_comportamiento['Created'] >= fecha_filtro_format]
        #     lt_out_capacidad = lt_out_capacidad[lt_out_capacidad['Created'] >= fecha_filtro_format]

        if formatted_date != '':
            lt_out_comportamiento['Modified'] = pd.to_datetime(lt_out_comportamiento['Modified'], format='%d/%m/%Y')
            lt_out_capacidad['Modified'] = pd.to_datetime(lt_out_capacidad['Modified'], format='%d/%m/%Y')

            fecha_filtro_format = pd.to_datetime(formatted_date, format='%d/%m/%Y')

            lt_out_comportamiento = lt_out_comportamiento[lt_out_comportamiento['Modified'] >= fecha_filtro_format]
            lt_out_capacidad = lt_out_capacidad[lt_out_capacidad['Modified'] >= fecha_filtro_format]

        if en is not None:
            if en == ETipoEva.FLAG_EVALUACION_y_PO:
                
                base['CONCAT'] = base.MATRICULA_CALIFICADOR.str.cat(base.MATRICULA_CALIFICADO.str.cat(base.CHAPTER_CALIFICADO, sep=''), sep='')
                lt_out_capacidad['CONCAT'] = lt_out_capacidad.MATRICULA_CALIFICADOR.str.cat(lt_out_capacidad.MATRICULA_CALIFICADO.str.cat(lt_out_capacidad.CHAPTER_CALIFICADO, sep=''), sep='')

                df_1 = pd.merge(lt_out_capacidad[['CONCAT']], base, how='left', on='CONCAT')
                df_1 = df_1.dropna(subset=['MATRICULA_CALIFICADOR'])

                conteo = df_1['MATRICULA_CALIFICADO'].value_counts()
                valores_conteo_4 = conteo[conteo == 4].index
                df_1 = df_1[df_1['MATRICULA_CALIFICADO'].isin(valores_conteo_4)]

                for i, matricula in enumerate(base['CONCAT']):
                    if matricula in df_1['CONCAT'].values:
                        base.loc[i, 'FLAG_EVALUACION'] = 'SI'
            
            elif en == ETipoEva.FLAG_AUTOEVALUACION_y_PO:

                base['CONCAT'] = base.MATRICULA_CALIFICADO.str.cat(base.MATRICULA_CALIFICADO.str.cat(base.CHAPTER_CALIFICADO, sep=''), sep='')
                lt_out_capacidad['CONCAT'] = lt_out_capacidad.MATRICULA_CALIFICADOR.str.cat(lt_out_capacidad.MATRICULA_CALIFICADO.str.cat(lt_out_capacidad.CHAPTER_CALIFICADO, sep=''), sep='')
                
                df_2 = pd.merge(lt_out_capacidad[['CONCAT']], base, how='left', on='CONCAT')
                df_2 = df_2.dropna(subset=['MATRICULA_CALIFICADOR'])

                conteo = df_2['MATRICULA_CALIFICADO'].value_counts()
                valores_conteo_4 = conteo[conteo == 4].index
                df_2 = df_2[df_2['MATRICULA_CALIFICADO'].isin(valores_conteo_4)]

                for i, matricula in enumerate(base['CONCAT']):
                    if matricula in df_2['CONCAT'].values:
                        base.loc[i, 'FLAG_AUTOEVALUACION'] = 'SI'

            elif en == ETipoEva.FLAG_EVALUACION_y_otros:

                base['CONCAT'] = base.MATRICULA_CALIFICADOR.str.cat(base.MATRICULA_CALIFICADO.str.cat(base.CHAPTER_CALIFICADO, sep=''), sep='')
                lt_out_comportamiento['CONCAT'] = lt_out_comportamiento.MATRICULA_CALIFICADOR.str.cat(lt_out_comportamiento.MATRICULA_CALIFICADO.str.cat(lt_out_comportamiento.CHAPTER_CALIFICADO, sep=''), sep='')
                
                df_3 = pd.merge(lt_out_comportamiento[['CONCAT']], base, how='left', on='CONCAT')
                df_3 = df_3.dropna(subset=['MATRICULA_CALIFICADOR'])

                conteo = df_3['MATRICULA_CALIFICADO'].value_counts()
                valores_conteo_3 = conteo[conteo == 3].index
                df_3 = df_3[df_3['MATRICULA_CALIFICADO'].isin(valores_conteo_3)]

                for i, matricula in enumerate(base['CONCAT']):
                    if matricula in df_3['CONCAT'].values:
                        base.loc[i, 'FLAG_EVALUACION'] = 'SI'


            elif en == ETipoEva.FLAG_AUTOEVALUACION_y_otros:

                base['CONCAT'] = base.MATRICULA_CALIFICADO.str.cat(base.MATRICULA_CALIFICADO.str.cat(base.CHAPTER_CALIFICADO, sep=''), sep='')
                lt_out_comportamiento['CONCAT'] = lt_out_comportamiento.MATRICULA_CALIFICADOR.str.cat(lt_out_comportamiento.MATRICULA_CALIFICADO.str.cat(lt_out_comportamiento.CHAPTER_CALIFICADO, sep=''), sep='')
                
                df_4 = pd.merge(lt_out_comportamiento[['CONCAT']], base, how='left', on='CONCAT')
                df_4 = df_4.dropna(subset=['MATRICULA_CALIFICADOR'])

                conteo = df_4['MATRICULA_CALIFICADO'].value_counts()
                valores_conteo_3 = conteo[conteo == 3].index
                df_4 = df_4[df_4['MATRICULA_CALIFICADO'].isin(valores_conteo_3)]

                for i, matricula in enumerate(base['CONCAT']):
                    if matricula in df_4['CONCAT'].values:
                        base.loc[i, 'FLAG_AUTOEVALUACION'] = 'SI'
            
            elif en == ETipoEva.AUTOEVALUACION_Y_EVALUACION_y_otros:
                base['CONCAT'] = base.MATRICULA_CALIFICADO.str.cat(base.MATRICULA_CALIFICADO.str.cat(base.CHAPTER_CALIFICADO, sep=''), sep='')
                lt_out_comportamiento['CONCAT'] = lt_out_comportamiento.MATRICULA_CALIFICADOR.str.cat(lt_out_comportamiento.MATRICULA_CALIFICADO.str.cat(lt_out_comportamiento.CHAPTER_CALIFICADO, sep=''), sep='')
                
                df_2 = pd.merge(lt_out_comportamiento[['CONCAT']], base, how='left', on='CONCAT')
                df_2 = df_2.dropna(subset=['MATRICULA_CALIFICADOR'])

                conteo = df_2['MATRICULA_CALIFICADO'].value_counts()
                valores_conteo_3 = conteo[conteo == 3].index
                df_2 = df_2[df_2['MATRICULA_CALIFICADO'].isin(valores_conteo_3)]

                for i, matricula in enumerate(base['CONCAT']):
                    if matricula in df_2['CONCAT'].values:
                        base.loc[i, 'FLAG_AUTOEVALUACION'] = 'SI'


                base['CONCAT'] = base.MATRICULA_CALIFICADOR.str.cat(base.MATRICULA_CALIFICADO.str.cat(base.CHAPTER_CALIFICADO, sep=''), sep='')
                lt_out_comportamiento['CONCAT'] = lt_out_comportamiento.MATRICULA_CALIFICADOR.str.cat(lt_out_comportamiento.MATRICULA_CALIFICADO.str.cat(lt_out_comportamiento.CHAPTER_CALIFICADO, sep=''), sep='')

                df_1 = pd.merge(lt_out_comportamiento[['CONCAT']], base, how='left', on='CONCAT')
                df_1 = df_1.dropna(subset=['MATRICULA_CALIFICADOR'])

                conteo = df_1['MATRICULA_CALIFICADO'].value_counts()
                valores_conteo_3 = conteo[conteo == 3].index
                df_1 = df_1[df_1['MATRICULA_CALIFICADO'].isin(valores_conteo_3)]

                for i, matricula in enumerate(base['CONCAT']):
                    if matricula in df_1['CONCAT'].values:
                        base.loc[i, 'FLAG_EVALUACION'] = 'SI'

            #actualizar FLAG_EXCLUSIÓN
            base.rename(columns={'MATRICULA_CALIFICADO': 'MATRICULA'}, inplace=True)
            df_5 = pd.merge(base_activos[['MATRICULA']], base, how='left', on='MATRICULA')
            if 'FLAG_AUTOEVALUACION' in df_5 and 'FLAG_EVALUACION' in df_5:
                df_5 = df_5.drop_duplicates(subset=['ESTADO', 'MATRICULA_CALIFICADOR', 'NOMBRES_CALIFICADOR', 'CORREO_CALIFICADOR', 'ROL_CALIFICADOR', 'MATRICULA', 'NOMBRES_CALIFICADO', 'CHAPTER_CALIFICADO', 'ROL_CALIFICADO', 'CORREO_CALIFICADO', 'FLAG_AUTOEVALUACION', 'FLAG_EVALUACION', 'FLAG_EXCLUSIÓN'])
            elif 'FLAG_EVALUACION' in df_5:
                df_5 = df_5.drop_duplicates(subset=['ESTADO', 'MATRICULA_CALIFICADOR', 'NOMBRES_CALIFICADOR', 'CORREO_CALIFICADOR', 'ROL_CALIFICADOR', 'MATRICULA', 'NOMBRES_CALIFICADO', 'CHAPTER_CALIFICADO', 'ROL_CALIFICADO', 'CORREO_CALIFICADO', 'FLAG_EVALUACION', 'FLAG_EXCLUSIÓN'])
            elif 'FLAG_AUTOEVALUACION' in df_5:
                df_5 = df_5.drop_duplicates(subset=['ESTADO', 'MATRICULA_CALIFICADOR', 'NOMBRES_CALIFICADOR', 'CORREO_CALIFICADOR', 'ROL_CALIFICADOR', 'MATRICULA', 'NOMBRES_CALIFICADO', 'CHAPTER_CALIFICADO', 'ROL_CALIFICADO', 'CORREO_CALIFICADO', 'FLAG_AUTOEVALUACION', 'FLAG_EXCLUSIÓN'])

            for i, matricula in enumerate(base['MATRICULA']):
                if matricula in df_5['MATRICULA'].values:
                    base.loc[i, 'ESTADO'] = 'ACTIVO'
                    base.loc[i, 'FLAG_EXCLUSIÓN'] = 'NO'
                else:
                    base.loc[i, 'ESTADO'] = 'INACTIVO'
                    base.loc[i, 'FLAG_EXCLUSIÓN'] = 'SI'
            
            base = base.drop(['CONCAT'], axis=1)
            copia_pega(ruta_plantilla, ruta_out_f)
            df_a_excel(ruta_out_f, 'BASE', base, f_ini = 2)

            xlapp = win32com.client.DispatchEx("Excel.Application")
            wb = xlapp.Workbooks.Open(ruta_out_f)
            xlapp.Visible = True
            wb.RefreshAll()
            xlapp.CalculateUntilAsyncQueriesDone()
            wb.Save()
            xlapp.Quit()

            QMessageBox.information(self, "Diálogo informativo", "Se generó el reporte en {}".format(ruta_out_f))

    def abrir_archivo_excel(self):
        global ruta_plantilla
        ruta_plantilla = QFileDialog.getOpenFileName(self, 'Abrir archivo', 'C:\\', 'Excel (*.xls *.xlsx)')
        ruta_plantilla = ruta_plantilla[0]

    def on_date_selected(self):
        selected_date = self.calendar.selectedDate
        global formatted_date
        formatted_date = selected_date.toString("dd/MM/yyyy")
        self.selected_date_label.text = f"Fecha seleccionada: {formatted_date}"