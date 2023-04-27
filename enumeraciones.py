from enum import Enum

class ETipoEva(Enum):
    FLAG_EVALUACION_y_PO = 1
    FLAG_AUTOEVALUACION_y_PO = 2
    FLAG_EVALUACION_y_otros = 3
    FLAG_AUTOEVALUACION_y_otros = 4
    FLAG_MEDICIÓN_y_otros = 5