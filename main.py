import pandas as pd
import xlwings as xw
import numpy as np

# lt_out_capacidad_enfoque = pd.read_excel("Excel_Entrada\PRUEBA_PLANTILLA_PRIORIZACIÓN.xlsx", sheet_name='LT_OUT_CAPACIDAD_ENFOQUE')
# lt_out_compromiso = pd.read_excel("Excel_Entrada\PRUEBA_PLANTILLA_PRIORIZACIÓN.xlsx", sheet_name='LT_OUT_COMPROMISO')
# lt_in_capacidad = pd.read_excel("Excel_Entrada\PRUEBA_PLANTILLA_PRIORIZACIÓN.xlsx", sheet_name='LT_IN_CAPACIDAD')

def main():
    ruta = r'D:\Ricardo\Documentos\Actualizar_formatos_priorizacion\Excel_Entrada\COLABORADOR_CURSO.xlsx'


    lt_in_colaborador_curso = leer_excel_simple(ruta, 'LT_IN_COLABORADOR_CURSO')

    print(lt_in_colaborador_curso)


def lastRow(ws, col=1):
    lwr_r_cell = ws.cells.last_cell
    lwr_row = lwr_r_cell.row
    lwr_cell = ws.range((lwr_row, col))

    if lwr_cell.value is None:
        lwr_cell = lwr_cell.end('up')

    return lwr_cell.row

def lastColumn(ws, row=1):
    lwr_r_cell = ws.cells.last_cell
    lwr_col = lwr_r_cell.column
    lwr_cell = ws.range((row, lwr_col))

    if lwr_cell.value is None:
        lwr_cell = lwr_cell.end('left')

    return lwr_cell.column

def leer_excel_simple(ruta,hoja=None,f_inicio=1, c_inicio=1,is_encuesta=False):
    header = 1

    app = xw.App(visible= False)
    app.display_alerts = False
    wb_api = app.books.api.Open(ruta, UpdateLinks=False, ReadOnly=True)
    wb = xw.Book(impl=xw._xlwindows.Book(xl=wb_api))
    
    ws = wb.sheets[0] if hoja is None else wb.sheets(hoja)
    # Obteneiendo rangos
    lr = lastRow(ws,c_inicio)
    lc = lastColumn(ws,f_inicio)

    # Caso encuesta
    if is_encuesta:
        header = 2 

    df = ws.range((f_inicio,c_inicio),(lr,lc)).options(pd.DataFrame, index=False,empty=np.nan, header=header).value

    wb.close()
    app.kill()

    return df

if __name__ == '__main__':
    main()
# lt_in_colaborador_curso.loc[lt_in_colaborador_curso['Columna2'] == '#N/D', 'Columna2'] = pd.NA
# lt_in_colaborador_curso.dropna(subset=['Columna2'], inplace=True)
# lt_in_colaborador_curso.reset_index(drop=True, inplace=True)

# lt_in_colaborador_curso.loc[lt_in_colaborador_curso['Columna4'] == '#N/D', 'Columna4'] = pd.NA
# lt_in_colaborador_curso.dropna(subset=['Columna4'], inplace=True)
# lt_in_colaborador_curso.reset_index(drop=True, inplace=True)

# curso_priorizado = lt_in_colaborador_curso[['Columna1']].copy()
# curso_priorizado.rename(columns={'Columna1': 'Columna2'}, inplace=True)

# curso_priorizado['Columna6'] = lt_in_colaborador_curso['Columna3'].copy()

# # Filtrar y eliminar filas con valor "#N/D" en la columna 2 de "LT_OUT_CAPACIDAD_ENFOQUE"
# lt_out_capacidad_enfoque.loc[lt_out_capacidad_enfoque['Columna2'] == '#N/D', 'Columna2'] = pd.NA
# lt_out_capacidad_enfoque.dropna(subset=['Columna2'], inplace=True)
# lt_out_capacidad_enfoque.reset_index(drop=True, inplace=True)

# # Copiar columna A de "LT_OUT_CAPACIDAD_ENFOQUE" a columna B de "CAPACIDAD_ENFOQUE"
# capacidad_enfoque = lt_out_capacidad_enfoque[['Columna1']].copy()
# capacidad_enfoque.rename(columns={'Columna1': 'Columna2'}, inplace=True)

# # Copiar columna C de "LT_OUT_CAPACIDAD_ENFOQUE" a columna G de "CAPACIDAD_ENFOQUE"
# capacidad_enfoque['Columna7'] = lt_out_capacidad_enfoque['Columna3'].copy()

# # Buscar valores de la columna G de "CAPACIDAD_ENFOQUE" en la columna 2 de "LT_IN_CAPACIDAD"
# capacidad_enfoque['Columna6'] = capacidad_enfoque['Columna7'].map(lt_in_capacidad.set_index('Columna2')['Columna3'])


# curso_priorizado.to_excel('result.xlsx', index=False)