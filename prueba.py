import os
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from datetime import timedelta

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

# Codificar 'DESCRIPCION' como variable categórica y guardar el mapeo
df_agrupado['DESCRIPCION'] = df_agrupado['DESCRIPCION'].astype('category')
descripcion_map = dict(enumerate(df_agrupado['DESCRIPCION'].cat.categories))
df_agrupado['DESCRIPCION'] = df_agrupado['DESCRIPCION'].cat.codes

# Definir características (X) y variable objetivo (y)
X = df_agrupado[['MES', 'AÑO', 'DESCRIPCION']]
y = df_agrupado['PIEZAS SURTIDAS']

# Dividir los datos en conjuntos de entrenamiento y prueba
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Crear el modelo de predicción
model = RandomForestRegressor(n_estimators=100, random_state=42)
model.fit(X_train, y_train)

# Pedir al usuario la fecha para la predicción
mes_pred = int(input("Ingresa el mes para la predicción (1-12): "))
anio_pred = int(input("Ingresa el año para la predicción: "))

# Crear un DataFrame para predecir para todos los medicamentos en la fecha especificada
medicamentos_unicos = df_agrupado['DESCRIPCION'].unique()
X_pred = pd.DataFrame({
    'MES': [mes_pred] * len(medicamentos_unicos),
    'AÑO': [anio_pred] * len(medicamentos_unicos),
    'DESCRIPCION': medicamentos_unicos
})

# Realizar las predicciones
predicciones = model.predict(X_pred)

# Crear DataFrame de salida con el nombre del medicamento, la cantidad predicha y la fecha de predicción
resultado = pd.DataFrame({
    'Nombre del Medicamento': X_pred['DESCRIPCION'].map(descripcion_map),
    'Cantidad Necesaria a Surtir (Predicción)': predicciones.astype(int),
    'Fecha de Predicción': pd.to_datetime(dict(year=X_pred['AÑO'], month=X_pred['MES'], day=[1]*len(X_pred)))
})

# Guardar el DataFrame en un archivo Excel
output_file = './predicciones.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    resultado.to_excel(writer, index=False, sheet_name='Surtido de Medicamentos')

# Abrir el archivo de Excel automáticamente
os.startfile(os.path.normpath(output_file))

print(f"Los datos se han guardado y el archivo '{output_file}' se ha abierto.")
