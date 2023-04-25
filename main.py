import pandas as pd

lt_in_colaborador_curso = pd.read_excel('Excel_Entrada\PRUEBA_PLANTILLA_PRIORIZACIÓN.xlsx', sheet_name='LT_IN_COLABORADOR_CURSO')
lt_out_capacidad_enfoque = pd.read_excel("Excel_Entrada\PRUEBA_PLANTILLA_PRIORIZACIÓN.xlsx", sheet_name='LT_OUT_CAPACIDAD_ENFOQUE')
lt_out_compromiso = pd.read_excel("Excel_Entrada\PRUEBA_PLANTILLA_PRIORIZACIÓN.xlsx", sheet_name='LT_OUT_COMPROMISO')
lt_in_capacidad = pd.read_excel("Excel_Entrada\PRUEBA_PLANTILLA_PRIORIZACIÓN.xlsx", sheet_name='LT_IN_CAPACIDAD')

lt_in_colaborador_curso.loc[lt_in_colaborador_curso['Columna2'] == '#N/D', 'Columna2'] = pd.NA
lt_in_colaborador_curso.dropna(subset=['Columna2'], inplace=True)
lt_in_colaborador_curso.reset_index(drop=True, inplace=True)

lt_in_colaborador_curso.loc[lt_in_colaborador_curso['Columna4'] == '#N/D', 'Columna4'] = pd.NA
lt_in_colaborador_curso.dropna(subset=['Columna4'], inplace=True)
lt_in_colaborador_curso.reset_index(drop=True, inplace=True)

curso_priorizado = lt_in_colaborador_curso[['Columna1']].copy()
curso_priorizado.rename(columns={'Columna1': 'Columna2'}, inplace=True)

curso_priorizado['Columna6'] = lt_in_colaborador_curso['Columna3'].copy()

# Filtrar y eliminar filas con valor "#N/D" en la columna 2 de "LT_OUT_CAPACIDAD_ENFOQUE"
lt_out_capacidad_enfoque.loc[lt_out_capacidad_enfoque['Columna2'] == '#N/D', 'Columna2'] = pd.NA
lt_out_capacidad_enfoque.dropna(subset=['Columna2'], inplace=True)
lt_out_capacidad_enfoque.reset_index(drop=True, inplace=True)

# Copiar columna A de "LT_OUT_CAPACIDAD_ENFOQUE" a columna B de "CAPACIDAD_ENFOQUE"
capacidad_enfoque = lt_out_capacidad_enfoque[['Columna1']].copy()
capacidad_enfoque.rename(columns={'Columna1': 'Columna2'}, inplace=True)

# Copiar columna C de "LT_OUT_CAPACIDAD_ENFOQUE" a columna G de "CAPACIDAD_ENFOQUE"
capacidad_enfoque['Columna7'] = lt_out_capacidad_enfoque['Columna3'].copy()

# Buscar valores de la columna G de "CAPACIDAD_ENFOQUE" en la columna 2 de "LT_IN_CAPACIDAD"
capacidad_enfoque['Columna6'] = capacidad_enfoque['Columna7'].map(lt_in_capacidad.set_index('Columna2')['Columna3'])


curso_priorizado.to_excel('result.xlsx', index=False)