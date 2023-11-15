import sys
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *
from __feature__ import true_property
from os import path, remove
import pandas as pd
from datetime import datetime
from base import copia_pega, df_a_excel, leer_excel_simple
import win32com.client
from constantes import PATH_BA

formatted_date = ''

class Window_priorizacion(QMainWindow):
    def setupUI(self):
        #super().__init__()

        self.setWindowTitle("Priorización")
        self.setFixedSize(1280, 720)

        self.fr_priorizacion = QFrame(self)
        self.fr_priorizacion.geometry=QRect(64, 36, 1152, 648)
        self.fr_priorizacion.styleSheet="background: white;"

        self.label = QLabel(self.fr_priorizacion)
        self.label.text = "Ingresa el nombre del chapter:" 
        self.label.geometry = QRect(0, 10, 500, 30)
        self.label.styleSheet = "color: gray; font-size: 20px; font-weight: bold;"

        self.chapter_input = QLineEdit(self.fr_priorizacion, placeholderText ="Chapter")
        self.chapter_input.geometry=QRect(20, 60, 360, 40)

        #calendario
        self.label_calendario = QLabel(self.fr_priorizacion)
        self.label_calendario.text = "Selecciona una fecha: (Opcional)"
        self.label_calendario.geometry = QRect(0, 100, 500, 30)
        self.label_calendario.styleSheet = "color: gray; font-size: 20px; font-weight: bold;"
        self.calendar = QCalendarWidget(self.fr_priorizacion)
        self.calendar.selectionChanged.connect(self.on_date_selected)
        self.calendar.geometry = QRect(20, 150, 400, 300)
        self.calendar.styleSheet = "background: gray;"
        
        self.selected_date_label = QLabel(self.fr_priorizacion)
        self.selected_date_label.geometry = QRect(50, 460, 500, 30)

        self.boton_ingresar_plantilla = QPushButton(self.fr_priorizacion)
        self.boton_ingresar_plantilla.text = "Importar plantilla de priorización"
        self.boton_ingresar_plantilla.clicked.connect(self.abrir_archivo_excel)
        self.boton_ingresar_plantilla.geometry = QRect(100, 500, 205, 23)

        self.submit_button = QPushButton(self.fr_priorizacion)
        self.submit_button.text = "Generar reporte de priorización"
        self.submit_button.clicked.connect(self.on_submit)
        self.submit_button.geometry=QRect(100, 560, 205, 23)

        self.boton_avance = QPushButton(self.fr_priorizacion)
        self.boton_avance.text = "Ir a avance medición"
        self.boton_avance.geometry=QRect(650, 500, 205, 23)
        self.boton_avance.styleSheet = "background: #33E9FF;"
        
        self.boton_calibracion = QPushButton(self.fr_priorizacion)
        self.boton_calibracion.text = "Ir a calibración"
        self.boton_calibracion.geometry=QRect(650, 600, 205, 23)
        self.boton_calibracion.styleSheet = "background: #33E9FF;"

    def on_submit(self):
        options = QFileDialog.Options()
        ruta_principal = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta", options=options)
        chapter = self.chapter_input.text

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
        nombre_archivo = 'PRIORIZACIÓN_{}_{}.xlsx'.format(chapter, fecha_hoy_format)
        ruta_out_f = path.join(ruta_principal, nombre_archivo)
        ruta_ba = PATH_BA
        colaboradores = leer_excel_simple(ruta_plantilla, 'COLABORADORES')
        cursos = leer_excel_simple(ruta_plantilla, 'CURSOS')
        lt_in_colaborador_curso = leer_excel_simple(ruta_plantilla, 'LT_IN_COLABORADOR_CURSO')
        lt_out_capacidad_enfoque = leer_excel_simple(ruta_plantilla, 'LT_OUT_CAPACIDAD_ENFOQUE')
        lt_out_compromiso = leer_excel_simple(ruta_plantilla, 'LT_OUT_COMPROMISO')
        lt_in_capacidad = leer_excel_simple(ruta_plantilla, 'LT_IN_CAPACIDAD')
        base_activos = leer_excel_simple(ruta_ba, 'BD ACTIVOS')
        base_activos.rename(columns={'Matrícula': 'MATRICULA'}, inplace=True)

        #filtros por fecha
        if formatted_date != '':
            lt_out_compromiso['Created'] = pd.to_datetime(lt_out_compromiso['Created'], format='%d/%m/%Y')
            lt_out_capacidad_enfoque['Created'] = pd.to_datetime(lt_out_capacidad_enfoque['Created'], format='%d/%m/%Y')
            lt_in_colaborador_curso['Created'] = pd.to_datetime(lt_in_colaborador_curso['Created'], format='%d/%m/%Y')

            fecha_filtro_format = pd.to_datetime(formatted_date, format='%d/%m/%Y')

            lt_out_compromiso = lt_out_compromiso[lt_out_compromiso['Created'] >= fecha_filtro_format]
            lt_out_capacidad_enfoque = lt_out_capacidad_enfoque[lt_out_capacidad_enfoque['Created'] >= fecha_filtro_format]
            lt_in_colaborador_curso = lt_in_colaborador_curso[lt_in_colaborador_curso['Created'] >= fecha_filtro_format]
        
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
        #lt_out_compromiso_filtrado = lt_out_compromiso_filtrado.drop_duplicates(subset=['MATRICULA', 'N_COMPROMISO', 'ACCION', 'RECURSO', 'FECHA_INI', 'FECHA_FIN', 'COMPROMISO'])
        lt_out_compromiso_filtrado = lt_out_compromiso_filtrado.dropna(subset=['COMPROMISO'])

        lt_out_compromiso_filtrado['FECHA_INI'] = [x.strftime("%d/%m/%Y") for x in lt_out_compromiso_filtrado['FECHA_INI']]
        lt_out_compromiso_filtrado['FECHA_FIN'] = [x.strftime("%d/%m/%Y") for x in lt_out_compromiso_filtrado['FECHA_FIN']]

        #actualizar FLAG_PRIORIZACIÓN
        df_3 = pd.merge(lt_in_colaborador_curso_filtrado[['MATRICULA']], colaboradores, how='left', on='MATRICULA')
        df_3 = df_3.drop_duplicates(subset=['MATRICULA', 'NOMBRE', 'ROL', 'ESTADO', 'MATRICULA_CALIFICADOR', 'NOMBRE_CALIFICADOR', 'ROL_CALIFICADOR', 'CHAPTER', 'FLAG_PRIORIZACIÓN', 'FLAG_EXCLUSIÓN', 'MOTIVO_EXCLUSIÓN'])

        for i, matricula in enumerate(colaboradores['MATRICULA']):
            if matricula in df_3['MATRICULA'].values:
                colaboradores.loc[i, 'FLAG_PRIORIZACIÓN'] = 'SI'

        #actualizar FLAG_EXCLUSIÓN
        df_4 = pd.merge(base_activos[['MATRICULA']], colaboradores, how='left', on='MATRICULA')
        df_4 = df_4.drop_duplicates(subset=['MATRICULA', 'NOMBRE', 'ROL', 'ESTADO', 'MATRICULA_CALIFICADOR', 'NOMBRE_CALIFICADOR', 'ROL_CALIFICADOR', 'CHAPTER', 'FLAG_PRIORIZACIÓN', 'FLAG_EXCLUSIÓN', 'MOTIVO_EXCLUSIÓN'])

        for i, matricula in enumerate(colaboradores['MATRICULA']):
            if matricula in df_4['MATRICULA'].values:
                colaboradores.loc[i, 'ESTADO'] = 'ACTIVO'
                colaboradores.loc[i, 'FLAG_EXCLUSIÓN'] = 'NO'
            else:
                colaboradores.loc[i, 'ESTADO'] = 'INACTIVO'
                colaboradores.loc[i, 'FLAG_EXCLUSIÓN'] = 'SI'
        
        #crear excel final
        copia_pega(ruta_plantilla, ruta_out_f)
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