import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton

class MyForm(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Formulario PySide6")
        self.setGeometry(100, 100, 400, 200)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()

        label = QLabel("Ingresa tu nombre:")
        layout.addWidget(label)

        self.text_input = QLineEdit()
        layout.addWidget(self.text_input)

        submit_button = QPushButton("Enviar")
        submit_button.clicked.connect(self.on_submit)
        layout.addWidget(submit_button)

        central_widget.setLayout(layout)

    def on_submit(self):
        user_input = self.text_input.text()
        print(f"Nombre ingresado: {user_input}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyForm()
    window.show()
    sys.exit(app.exec_())
