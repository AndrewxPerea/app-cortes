import os
from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
import io
from funciones import procesar_excel, procesar_archivo_csv_solo, procesar_archivo_excel_solo


app = Flask(__name__)

# Variable global para almacenar el archivo Excel resultante
resultado_excel = None

#index

@app.route('/')
def index():
    return render_template('index.html')

#pagina reconexiones

@app.route('/reconexiones',)
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

    return render_template('resultado.html', data=df_resultado.to_dict(orient='records'), columns=df_resultado.columns)


@app.route('/solointernet', methods=['GET', 'POST'])
def solointernet():
    global resultado_excel
    if request.method == 'POST':
        abonados_file = request.files['abonados_solointernet']
        cortes_file = request.files['olt']

        df_abonados = procesar_archivo_excel_solo(abonados_file)
        df_cortes = procesar_archivo_csv_solo(cortes_file)

        if not df_abonados.empty and not df_cortes.empty:
            resultado = pd.merge(df_abonados, df_cortes, how='right', left_on='EQUIPO MACO', right_on='NSN', suffixes=('_abonados', '_cortes'))
            resultado = resultado.dropna(subset=['EQUIPO MACO'])
            resultado.columns = resultado.columns.str.lower()

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                resultado.to_excel(writer, index=False, sheet_name='Resultado')

                abonados_filtrados = resultado[
                    (resultado['detalle suscripcion'].str.contains('@', na=False)) & 
                    (resultado['catv'] == 'Enabled')
                ]
                columnas_deseadas = [
                'n° abonado', 'documento', 'nombre', 'apellido',
                'estatus', 'equipo maco', 'detalle suscripcion', 'sn', 'olt', 
                'catv', 'administrative status'
            ]
                abonados_filtrados = abonados_filtrados[columnas_deseadas]
                if not abonados_filtrados.empty:
                    abonados_filtrados.to_excel(writer, index=False, sheet_name='Abonados Filtrados')

            output.seek(0)
            resultado_excel = output

            return render_template('resultado.html', data=abonados_filtrados.to_dict(orient='records'), columns=abonados_filtrados.columns)

    return render_template('solointernet.html')


@app.route('/noactivos', methods=['GET', 'POST'])
def noactivos():
    global resultado_excel
    if request.method == 'POST':
        abonados_file = request.files['abonados_solointernet']
        cortes_file = request.files['olt']

        try:

            df_abonados = procesar_archivo_excel_solo(abonados_file)
            df_cortes = procesar_archivo_csv_solo(cortes_file)
        except Exception as e:
            return render_template('error.html', error=str(e))

        if not df_abonados.empty and not df_cortes.empty:
            resultado = pd.merge(df_abonados, df_cortes, how='right', left_on='EQUIPO MACO', right_on='NSN', suffixes=('_abonados', '_cortes'))
            resultado = resultado.dropna(subset=['EQUIPO MACO'])
            resultado.columns = resultado.columns.str.lower()

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                resultado.to_excel(writer, index=False, sheet_name='Resultado')

                abonados_filtrados = resultado[
                (resultado['estatus'].str.lower().isin(['activo', 'por instalar']) == False) &  # Filtra los que no sean "estatus activo" o "estatus por instalar"
                ((resultado['administrative status'].str.lower() == 'enabled') |  # Y que tengan "administrative status" o "catv" en "enabled"
                (resultado['catv'].str.lower() == 'enabled'))]            
                
                if not abonados_filtrados.empty:
                    abonados_filtrados.to_excel(writer, index=False, sheet_name='Abonados Filtrados')
                
                columnas_deseadas = [
                    'n° abonado', 'documento', 'nombre', 'apellido',
                    'estatus', 'status', 'equipo maco', 'sn', 'olt', 
                    'catv', 'administrative status'
                ]
                abonados_filtrados = abonados_filtrados[columnas_deseadas]
                
                if not abonados_filtrados.empty:
                    abonados_filtrados.to_excel(writer, index=False, sheet_name='Abonados Filtrados')
                

            output.seek(0)
            resultado_excel = output

            return render_template('resultado.html', data=abonados_filtrados.to_dict(orient='records'), columns=abonados_filtrados.columns)

    return render_template('noactivos.html')

@app.route('/cortes', methods=['GET', 'POST'])
def cortes():
    global resultado_excel
    if request.method == 'POST':
        abonados_file = request.files['abonados']
        cortes_file = request.files['cortes']
        sae_file = request.files['asaeplus']
        
        df_cortes = procesar_archivo_excel_solo(abonados_file)
        df_olt = procesar_archivo_csv_solo(cortes_file)
        df_saeplus = procesar_archivo_excel_solo(sae_file)

        if not df_cortes.empty and not df_saeplus.empty and not df_olt.empty:
            resultado = pd.merge( df_saeplus, df_cortes, how='right', left_on='N° Abonado', right_on='N° Abonado')
            resultado = resultado.dropna(subset=['N° Abonado'])
            resultado =pd.merge(resultado, df_olt, left_on='EQUIPO MACO_y', right_on='NSN', suffixes=('_abonados', '_cortes'))
            resultado = resultado.dropna(subset=['EQUIPO MACO_y'])
            resultado.columns = resultado.columns.str.lower()
            columnas_deseadas = [
                    'n° abonado', 'documento_x', 'nombre_x', 'apellido_x',
                    'estatus_x', 'observaciones', 'sn', 'olt', 
                    'catv', 'administrative status'
                ]
            resultado_filtrado = resultado[columnas_deseadas]
            resultado_filtrado = resultado_filtrado [
                (resultado['observaciones'].isna()) &
                (resultado['estatus_x'] == 'CORTADO') & 
                ((resultado['catv'] == 'Enabled') |
                (resultado['administrative status'] == 'Enabled'))
            ]
          
            output_filtrado = io.BytesIO()
            with pd.ExcelWriter(output_filtrado, engine='xlsxwriter') as writer_filtrado:
                resultado_filtrado.to_excel(writer_filtrado, index=False, sheet_name='Resultado Filtrado')
            output_filtrado.seek(0)

            resultado_excel = output_filtrado
            

            return render_template('resultado.html', data=resultado_filtrado.to_dict(orient='records'), columns=resultado_filtrado.columns)

    return render_template('cortes.html')



@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'No file part'

        file = request.files['file']

        if file.filename == '':
            return 'No selected file'

        if file:
            # Crear el directorio uploads si no existe
            upload_dir = 'uploads'
            if not os.path.exists(upload_dir):
                os.makedirs(upload_dir)

            # Guardar archivo subido temporalmente
            file_path = os.path.join(upload_dir, file.filename)
            file.save(file_path)

            try:
                # Procesar el archivo Excel
                processed_file = procesar_excel(file_path)

                # Eliminar el archivo después de procesarlo
                os.remove(file_path)

                # Devolver el archivo procesado
                return send_file(processed_file, as_attachment=True)
            except Exception as e:
                # Asegurarse de eliminar el archivo si ocurre un error
                os.remove(file_path)
                return f"Error processing file: {str(e)}"

    return render_template('upload.html')


@app.route('/descargar_resultado')
def descargar_resultado():
    global resultado_excel
    if resultado_excel:
        return send_file(resultado_excel, as_attachment=True, download_name='resultado.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)

