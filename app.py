from flask import Flask, render_template, request, redirect, url_for
import pandas as pd


app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/procesar', methods=['POST'])
def procesar_archivos():
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

    return render_template('resultado.html', data=df_resultado.to_html(classes='table table-bordered table-success table-striped text-center  table-hover'))
    
    

if __name__ == '__main__':
    app.run(debug=True)
