import pandas as pd
from funciones import procesar_archivo_excel_solo, procesar_archivo_csv_solo

saeplus = 'saeplus.xlsx'
cortes ='cortes.xlsx'
olt = 'olt.csv'

df_saeplus = procesar_archivo_excel_solo(saeplus)
df_cortes = procesar_archivo_excel_solo(cortes)
df_olt = procesar_archivo_csv_solo(olt)


if not df_cortes.empty and not df_saeplus.empty and not df_olt.empty:
            resultado = pd.merge( df_saeplus, df_cortes, how='right', left_on='N° Abonado', right_on='N° Abonado')
            resultado = resultado.dropna(subset=['N° Abonado'])
            print(resultado.head(1))
            resultado =pd.merge(resultado, df_olt, left_on='EQUIPO MACO_y', right_on='NSN', suffixes=('_abonados', '_cortes'))
            resultado = resultado.dropna(subset=['EQUIPO MACO_y'])
            resultado.columns = resultado.columns.str.lower()
            resultado_filtrado = resultado[
                (resultado['observaciones'].isna()) &
                (resultado['estatus_x'] == 'CORTADO') & 
                ((resultado['catv'] == 'Enabled') |
                 (resultado['administrative status'] == 'Enabled'))
            ]

            print(resultado_filtrado.head(10))