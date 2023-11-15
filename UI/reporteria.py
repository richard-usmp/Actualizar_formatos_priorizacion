import sys
from PySide6.QtWidgets import QApplication, QMessageBox
from vista_priorizacion import Window_priorizacion
from vista_avance_medicion import Window_avance_medicion
from vista_calibracion import Window_calibracion

class reporteria:
    def __init__(self):
        
        #ventanas
        self.vista_priorizacion = Window_priorizacion()
        self.vista_avance_medicion = Window_avance_medicion()
        self.vista_calibracion = Window_calibracion()

        self.vista_priorizacion.setupUI()
        self.vista_avance_medicion.setupUI()
        self.vista_calibracion.setupUI()

        #cambiar a ventanas
        self.vista_priorizacion.boton_avance.clicked.connect(self.entrar_avance_medicion)
        self.vista_priorizacion.boton_calibracion.clicked.connect(self.entrar_calibracion)
        
        self.vista_avance_medicion.boton_priorizacion.clicked.connect(self.entrar_priorizacion)
        self.vista_avance_medicion.boton_calibracion.clicked.connect(self.entrar_calibracion)
        
        self.vista_calibracion.boton_avance.clicked.connect(self.entrar_avance_medicion)
        self.vista_calibracion.boton_priorizacion.clicked.connect(self.entrar_priorizacion)


    def entrar_avance_medicion(self):
        self.vista_avance_medicion.show()
        self.vista_priorizacion.hide()
        self.vista_calibracion.hide()
    
    def entrar_priorizacion(self):
        self.vista_avance_medicion.hide()
        self.vista_priorizacion.show()
        self.vista_calibracion.hide()
    
    def entrar_calibracion(self):
        self.vista_avance_medicion.hide()
        self.vista_priorizacion.hide()
        self.vista_calibracion.show()

app = QApplication(sys.argv)
inicio = reporteria()
inicio.vista_priorizacion.show()
sys.exit(app.exec())