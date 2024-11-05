import os
import pandas as pd

# Cargar los datos desde el archivo Excel
file_path = './MedicamentosDataset.xlsx'  # Asegúrate de colocar la ruta correcta del archivo
df = pd.read_excel(file_path, engine='openpyxl')

# Convertir 'FECHA' a tipo datetime y extraer el mes y el año
df['FECHA'] = pd.to_datetime(df['FECHA'])
df['MES'] = df['FECHA'].dt.month
df['AÑO'] = df['FECHA'].dt.year
df['PIEZAS SURTIDAS'] = df['PIEZAS SURTIDAS'].fillna(0)

# Agrupar por medicamento, mes y año para obtener el total mensual por medicamento
df_agrupado = df.groupby(['DESCRIPCION', 'MES', 'AÑO']).agg({'PIEZAS SURTIDAS': 'sum'}).reset_index()

# Guardar el DataFrame en un nuevo archivo Excel
output_file = './medicamentos_surtidos_mensual.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df_agrupado.to_excel(writer, index=False, sheet_name='Surtidos Mensuales')

# Abrir el archivo de Excel automáticamente
os.startfile(os.path.normpath(output_file))

print(f"Los datos se han guardado y el archivo '{output_file}' se ha abierto.")
