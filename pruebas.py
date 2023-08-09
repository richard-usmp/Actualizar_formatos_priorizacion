import pandas as pd
from base import leer_excel_simple

# Leer el archivo de Excel
ruta_BD = 'D:\BCP Effio\Documents\Actualizar_formatos_priorizacion\Excel_Entrada\Base de datos.xlsx'
df = leer_excel_simple(ruta_BD, 'Hoja1')

# Obtener la lista de chapters únicos
chapters = df['DESCRIPCION'].unique()

# Generar un archivo de Excel separado por cada chapter
for chapter in chapters:
    # Filtrar el dataframe por el chapter actual
    df_chapter = df[df['DESCRIPCION'] == chapter]

    # Crear un nuevo archivo de Excel para el chapter actual
    nombre_archivo = f'{chapter}.xlsx'
    writer = pd.ExcelWriter(nombre_archivo)
    df_chapter.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()

    print(f'Se ha generado el archivo de Excel para el chapter "{chapter}".')

print('Se han generado todos los archivos de Excel por chapter.')

import pandas as pd

# Lee el archivo de Excel y especifica que la hoja se encuentra en la fila 5 y columna 2
ruta_archivo = 'Ruta/del/archivo/Resumen TMs.xlsx'
nombre_hoja = 'NombreDeLaHoja'
fila_inicio = 5
columna_inicio = 'B'  # Columna 'B' es la columna 2

# Lee el contenido de la hoja de Excel a partir de la fila y columna especificadas
data_frame = pd.read_excel(ruta_archivo, sheet_name=nombre_hoja, header=None, skiprows=range(0, fila_inicio-1), usecols=columna_inicio)

# Ahora, si deseas guardar estos datos en un nuevo archivo Excel:
ruta_nuevo_archivo = 'Ruta/del/archivo/NuevoResumen.xlsx'
nombre_nueva_hoja = 'NuevaHoja'

data_frame.to_excel(ruta_nuevo_archivo, sheet_name=nombre_nueva_hoja, index=False, header=False)

