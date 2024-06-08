from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
import io

app = Flask(__name__)

# Variable global para almacenar el archivo Excel resultante
resultado_excel = None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/reconexiones')
def reconexiones():
    return render_template('reconexiones.html')

@app.route('/procesar', methods=['POST'])
def procesar_archivos():
    global resultado_excel
    abonados_file = request.files['abonados']
    cortes_file = request.files['cortes']

    try:
        df_cortes = pd.read_excel(cortes_file, usecols=[0, 1, 2, 3, 4])
        df_abonados = pd.read_excel(abonados_file)
    except Exception as e:
        return render_template('error.html', error=str(e))

    df_cortes = df_cortes.rename(columns={df_cortes.columns[0]: 'Abonados'})
    df_abonados = df_abonados.rename(columns={df_abonados.columns[0]: 'Abonados'})
    df_cortes.columns = df_cortes.columns.str.lower()
    df_abonados.columns = df_abonados.columns.str.lower()

    df_resultado = pd.merge(df_cortes, df_abonados, on="abonados", how="inner")

    df_resultado = df_resultado[['abonados', 'documento_x', 'nombre_x', 'apellido_x', 'observaciones', 'estatus']]
    df_resultado = df_resultado[(df_resultado['observaciones'].isna() | (df_resultado['observaciones'] == '')) & 
                                (df_resultado['estatus'] == 'ACTIVO')]

    if df_resultado.empty:
        return render_template('exitoso.html')

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_resultado.to_excel(writer, index=False, sheet_name='Resultado')
    output.seek(0)

    resultado_excel = output

    return render_template('resultado.html', data=df_resultado.to_dict(orient='records'))

def procesar_archivo_csv(archivo):
    try:
        df = pd.read_csv(archivo, usecols=['SN', 'Name', 'OLT', 'CATV', "Administrative status", 'Service port upload speed'])
        df['codigo'] = df['Name'].str.split(' - ').str[0]
        return df
    except Exception as e:
        print(f"Ocurrió un error al procesar {archivo}:", e)
        return pd.DataFrame()  # Devolver un DataFrame vacío en caso de error

def procesar_archivo_excel(archivo):
    try:
        df = pd.read_excel(archivo, usecols=[0, 1, 2, 3, 4])
        df['Nombre'] = df.apply(lambda row: ' '.join([str(row['Nombre']), str(row['Apellido'])]), axis=1)
        df.drop(columns=['Apellido'], inplace=True)
        return df
    except Exception as e:
        print(f"Ocurrió un error al procesar {archivo}:", e)
        return pd.DataFrame()  # Devolver un DataFrame vacío en caso de error

@app.route('/cortes', methods=['GET', 'POST'])
def cortes():
    global resultado_excel
    if request.method == 'POST':
        abonados_file = request.files['abonados']
        cortes_file = request.files['cortes']

        df_abonados = procesar_archivo_excel(abonados_file)
        df_cortes = procesar_archivo_csv(cortes_file)

        if not df_abonados.empty and not df_cortes.empty:
            resultado = pd.merge(df_abonados, df_cortes, how='right', left_on='N° Abonado', right_on='codigo', suffixes=('_abonados', '_cortes'))
            resultado = resultado.dropna(subset=['N° Abonado'])
            resultado.columns = resultado.columns.str.lower()

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                resultado.to_excel(writer, index=False, sheet_name='Resultado')
            output.seek(0)

            resultado_excel = output

            resultado_filtrado = resultado[
                (resultado['observaciones'].isna()) & 
                ((resultado['catv'] == 'Enabled') |
                 (resultado['administrative status'] == 'Enabled'))
            ]

            output_filtrado = io.BytesIO()
            with pd.ExcelWriter(output_filtrado, engine='xlsxwriter') as writer_filtrado:
                resultado_filtrado.to_excel(writer_filtrado, index=False, sheet_name='Resultado Filtrado')
            output_filtrado.seek(0)

            resultado_excel = output_filtrado

            return render_template('resultado.html', data=resultado_filtrado.to_dict(orient='records'))

    return render_template('cortes.html')

@app.route('/descargar_resultado')
def descargar_resultado():
    global resultado_excel
    if resultado_excel:
        return send_file(resultado_excel, as_attachment=True, download_name='resultado.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
