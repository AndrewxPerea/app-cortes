import pandas as pd
from funciones import procesar_archivo_excel_solo, procesar_archivo_csv_solo

# Archivos de entrada
drive_file = 'drive.xlsx'
saeplus_file = 'saeplus.xlsx'
epayco_file = 'epayco.xlsx'
olt_file = 'olt.csv'
# Función para cargar archivos
def cargar_archivos():
    try:
        df_drive = pd.read_excel(drive_file)
        df_saeplus = pd.read_excel(saeplus_file)
        df_epayco = pd.read_excel(epayco_file)
        df_abonado_cortes = procesar_archivo_excel_solo(saeplus_file)
        df_olt_cortes = procesar_archivo_csv_solo(olt_file)
        return df_drive, df_saeplus, df_epayco, df_abonado_cortes, df_olt_cortes
    except Exception as e:
        print(f"Error al cargar archivos: {e}")
        return None, None, None, None, None, None

# Cargar datos
df_drive, df_saeplus, df_epayco, df_abonado_cortes, df_olt_cortes = cargar_archivos()

# Renombrar y normalizar columnas
def normalizar_columnas(df, rename_col):
    df = df.rename(columns={df.columns[0]: rename_col})
    df.columns = df.columns.str.lower()
    return df
df_drive = normalizar_columnas(df_drive, 'abonados')
df_saeplus = normalizar_columnas(df_saeplus, 'abonados')
df_epayco = normalizar_columnas(df_epayco, 'abonados')


# Merge de DataFrames
df_resultado1 = pd.merge(df_drive, df_saeplus, on="abonados", how="inner")
df_resultado2 = pd.merge(df_drive, df_epayco, on="abonados", how="inner")
df_resultado3 = pd.merge(
    df_abonado_cortes, df_olt_cortes, how='right', 
    left_on='EQUIPO MACO', right_on='NSN', suffixes=('_abonados', '_cortes')
).dropna(subset=['EQUIPO MACO'])
df_resultado3.columns = df_resultado3.columns.str.lower()

# Selección de columnas relevantes
df_resultado1 = df_resultado1[['abonados', 'documento_x', 'nombre_x', 'apellido_x', 'observaciones', 'estatus_y', 'detalle suscripcion_x']]
df_resultado2 = df_resultado2[['abonados', 'documento_x', 'nombre', 'apellido', 'observaciones', 'estatus', 'detalle suscripcion']]
df_resultado3 = df_resultado3[['n° abonado', 'documento', 'nombre', 'apellido', 'estatus', 'status', 'olt', 'catv', 'administrative status', 'detalle suscripcion']]

#Filtra abonados que en workdrive no tienen observaciones y que su estatus es activo en saeplus
df_resultado1 = df_resultado1[(df_resultado1['observaciones'].isna() | (df_resultado1['observaciones'] == '')) & 
                            (df_resultado1['estatus_y'] == 'ACTIVO')]

# Filtra abonados que en workdrive no tienen observaciones y que hayan pagado en epayco
abonados_epayco = df_resultado2[(df_resultado2['observaciones'].isna() | (df_resultado2['observaciones'] == '')) ]

#filtra abonados que en saeplus esten activos  y que en olt no esten habilitados
df_resultado3 = df_resultado3 [
                (df_resultado3 ['estatus'].str.lower().isin(['activo']) == True) & 
                ((df_resultado3 ['administrative status'].str.lower() != 'enabled') |
                (df_resultado3 ['catv'].str.lower() !=  'enabled') ) |(
                (df_resultado3 ['status'].str.lower() !=  'online')
                )    # Filtra solo los que tienen "status" como "online"
                ]

# Filtra abonados que en workdrive no tienen observaciones y que su estatus es activo en saeplus con la olt
df_resultado3 =df_resultado3.rename(columns={df_resultado3.columns[0]: 'abonados'})
df_resultado4 = pd.merge(df_resultado3, df_resultado1, on="abonados" , how="inner")

# Filtrar filas donde 'detalle de plan' contiene '@' y 'catv' está en 'enabled'
df_resultado5 = df_resultado4[
    (df_resultado4['detalle suscripcion'].str.contains('@', na=False)) & 
    (df_resultado4['catv'].str.lower() == 'enabled')   |
    ~(df_resultado4['detalle suscripcion'].str.contains('@', na=False)) &
    (df_resultado4['catv'].str.lower() != 'enabled')
]


# Guardar los resultados en diferentes pestañas de un archivo Excel
output_file = "resultado_merge.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df_resultado1.to_excel(writer, sheet_name='Reconexion sin observaciones', index=False)
    abonados_epayco.to_excel(writer, sheet_name='Pagos de epayco', index=False)
    df_resultado4.to_excel(writer, sheet_name='prueba 1', index=False)
    df_resultado5.to_excel(writer, sheet_name='Abonados sin activar', index=False)
