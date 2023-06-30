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
