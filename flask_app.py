from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
import io

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/cpendiente')
def inicio():
    return render_template('cpendiente.html')

# Variable global para almacenar el archivo Excel resultante
resultado_excel = None

# Auditoria cortes colocando pendiente
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
    

    # Realizar la fusión de los DataFrames basada en la columna "Abonados"
    df_resultado = pd.merge(df_cortes, df_abonados, on="abonados", how="inner")

    # Filtrar los registros con observaciones vacías o NaN y con Estatus igual a "ACTIVO"
    df_resultado = df_resultado[['abonados', 'documento_x', 'nombre_x', 'apellido_x', 'observaciones', 'estatus']]
    df_resultado = df_resultado[(df_resultado['observaciones'].isna() | (df_resultado['observaciones'] == '')) & 
                                (df_resultado['estatus'] == 'ACTIVO')]

    if df_resultado.empty:
        return render_template('exitoso.html') #return

    # Guardar el DataFrame en un archivo Excel en la memoria
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df_resultado.to_excel(writer, index=False, sheet_name='Resultado')
    writer.close()
    output.seek(0)

    # Guardar el archivo Excel en la variable global
    resultado_excel = output

    return render_template('resultado.html', data=df_resultado.to_dict(orient='records'))

@app.route('/descargar_resultado')
def descargar_resultado():
    global resultado_excel
    if resultado_excel:
        return send_file(resultado_excel, as_attachment=True, download_name='resultado.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
