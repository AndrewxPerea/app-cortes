import pandas as pd
import re
from funciones import procesar_archivo_csv_solo

# Cargar archivos
abonados = pd.read_excel('abonados.xlsx', usecols=[0])
saeplus = pd.read_excel('saeplus.xlsx')

# Renombrar la columna del archivo abonados
abonados = abonados.rename(columns={abonados.columns[0]: 'abonados'})
saeplus = saeplus.rename(columns={saeplus.columns[0]: 'abonados'})

# Unir abonados y saeplus por la columna abonados
data = pd.merge(abonados, saeplus, on="abonados", how="inner")
data['EQUIPO MACO'] = data['EQUIPO MAC'].astype(str).str[-8:]

# Procesar el archivo CSV de OLT
olt = procesar_archivo_csv_solo('olt.csv')

# Fusionar los datos de abonados y OLT
if not data.empty and not olt.empty:
    resultado = pd.merge(data, olt, how='right', left_on='EQUIPO MACO', right_on='NSN', suffixes=('_abonados', '_cortes'))
    resultado = resultado.dropna(subset=['EQUIPO MACO'])
    resultado.columns = resultado.columns.str.lower()

# Función para extraer el valor de velocidad (número y MG) de 'detalle suscripcion'
def extraer_velocidad(detalle):
    # Buscar el patrón con o sin espacios entre el número y 'MG'
    match = re.search(r'(\d+)\s*MG', detalle.upper())  # Convertir el texto a mayúsculas para evitar problemas de may/min
    if match:
        return match.group(1) + 'MG'  # Retorna el número encontrado y 'MG' pegados
    return None  # Retorna None si no encuentra el patrón

# Aplicar la extracción de velocidad a la columna 'detalle suscripcion'
resultado['velocidad_detalle'] = resultado['detalle suscripcion'].apply(extraer_velocidad)

# Revisar si algunos valores de 'velocidad_detalle' fueron correctamente extraídos
print(resultado[['detalle suscripcion', 'velocidad_detalle']].head(10))

# Filtrar los abonados que no coincidan en velocidad entre 'velocidad_detalle' y 'service port download speed'
abonados_filtrados = resultado[
    (resultado['velocidad_detalle'] != resultado['service port download speed'])
]

# Guardar los resultados filtrados en un archivo Excel
abonados_filtrados.to_excel('fusion_resultado.xlsx', index=False)

# Mostrar los primeros 5 resultados
print(abonados_filtrados.head(5))
