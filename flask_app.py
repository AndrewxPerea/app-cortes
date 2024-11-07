import os
from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
import io
from funciones import procesar_excel, procesar_archivo_csv_solo, procesar_archivo_excel_solo
import re


app = Flask(__name__)

# Variable global para almacenar el archivo Excel resultante
resultado_excel = None

#index

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/mayuscula')
def mayuscula():
    return render_template('mayus.html')



# reconexiones___________________________________________________

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

#solo @ ________________________________________________________________________


@app.route('/solointernet', methods=['GET', 'POST'])
def solointernet():
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
                resultado.to_excel(writer, index=False, sheet_name='Todos los abonados')

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
                    abonados_filtrados.to_excel(writer, index=False, sheet_name='Abonados solo @ Con catv')

            output.seek(0)
            resultado_excel = output

            return render_template('resultado.html', data=abonados_filtrados.to_dict(orient='records'), columns=abonados_filtrados.columns)
    return render_template('solointernet.html')

#No activos ___________________________________________________________________

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
                ((resultado['administrative status'].str.lower() == 'enabled') |
                (resultado['catv'].str.lower() == 'enabled')) &
                (resultado['status'].str.lower() == 'online')  # Filtra solo los que tienen "status" como "online"
                ]
                
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
            import time
            time.sleep(3)
            return render_template('resultado.html', data=abonados_filtrados.to_dict(orient='records'), columns=abonados_filtrados.columns)

    return render_template('noactivos.html')
#Cortes____________________________________________________________________________________________


@app.route('/cortes', methods=['GET', 'POST'])
def cortes():
    global resultado_excel
    if request.method == 'POST':
        abonados_file = request.files['abonados']
        cortes_file = request.files['cortes']
        sae_file = request.files['asaeplus']
        try:
            df_cortes = procesar_archivo_excel_solo(abonados_file)
            df_olt = procesar_archivo_csv_solo(cortes_file)
            df_saeplus = procesar_archivo_excel_solo(sae_file)
        except Exception as e:
            return render_template('error.html', error=str(e))

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


#Mensajes/________________________________________________________________________________

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

#///////////// comparador de planes
@app.route('/verificar_velocidad', methods=['GET', 'POST'])
def verificar_velocidad():
    global resultado_excel
    if request.method == 'POST':
        saeplus_file = request.files['saeplus']
        olt_file = request.files['olt']
        try:
            # Leer los archivos cargados
            saeplus = procesar_archivo_excel_solo(saeplus_file)
            olt = procesar_archivo_csv_solo(olt_file)
        except Exception as e:
         return render_template('error.html', error=str(e))



        if not saeplus.empty and not olt.empty:
            resultado = pd.merge(saeplus, olt, how='right', left_on='EQUIPO MACO', right_on='NSN', suffixes=('_abonados', '_OLT'))
            resultado = resultado.dropna(subset=['EQUIPO MACO'])
            resultado.columns = resultado.columns.str.lower()

            # Función para extraer la velocidad
            def extraer_velocidad(detalle):
                match = re.search(r'(\d+)\s*MG', detalle.upper())
                if match:
                    return match.group(1) + 'MG'
                return None

            # Aplicar la extracción de velocidad
            resultado['velocidad_detalle'] = resultado['detalle suscripcion'].apply(extraer_velocidad)

            # Filtrar los abonados que no coinciden en velocidad
            abonados_filtrados = resultado[
                (resultado['velocidad_detalle'] != resultado['service port download speed'])
            ]

            columnas_deseadas = [
                    'n° abonado', 'documento', 'nombre', 'name','estatus',
                    'detalle suscripcion', 'nombre franquicia', 'equipo maco', 'sn', 'olt', 
                    'service port upload speed', 'service port download speed'
                ]
            abonados_filtrados = abonados_filtrados[columnas_deseadas]
            output_filtrado = io.BytesIO()
            with pd.ExcelWriter(output_filtrado, engine='xlsxwriter') as writer_filtrado:
                abonados_filtrados.to_excel(writer_filtrado, index=False, sheet_name='Resultado Filtrado')
            output_filtrado.seek(0)

            resultado_excel = output_filtrado


            return render_template('resultado.html', data=abonados_filtrados.to_dict(orient='records'), columns=abonados_filtrados.columns)



    return render_template('verificar_velocidad.html')



@app.route('/diferentes', methods=['GET', 'POST'])
def diferentes():
    if request.method == 'POST':
        saeplus = request.files['saeplus']
        olt = request.files['olt']
        try:
            saeplus = procesar_archivo_excel_solo(saeplus)
            olt2 = procesar_archivo_csv_solo(olt)
            olt = olt2[olt2['Status'] == 'Online']
        except Exception as e:
            return render_template('error.html', error=str(e))

        if not saeplus.empty and not olt.empty:
            # Fusionar solo los registros coincidentes
            resultado = pd.merge(saeplus, olt, how='right', left_on='EQUIPO MACO', right_on='NSN', suffixes=('_abonados', '_cortes'))
            resultado= resultado[resultado['Status'] == 'Online']
            resultado = resultado.dropna(subset=['EQUIPO MACO'])
            resultado.columns = resultado.columns.str.lower()

        if not saeplus.empty and not olt2.empty:
            # Confirmar que existen las columnas 'EQUIPO MACO' y 'NSN'
            if 'EQUIPO MACO' in saeplus.columns and 'NSN' in olt2.columns:
                # Hacer la fusión con 'indicator=True' para identificar los registros coincidentes
                resultado = pd.merge(saeplus, olt2, how='outer', left_on='EQUIPO MACO', right_on='NSN', suffixes=('_abonados', '_cortes'), indicator=True)

                if '_merge' in resultado.columns:
                    # Filtrar registros donde 'EQUIPO MACO' y 'NSN' no coinciden
                    resultado_diferente = resultado[resultado['_merge'] != 'both']
                    resultado_todos_diferentes = resultado_diferente
                    resultado_solo_equipo_mac = resultado_diferente.dropna(subset=['EQUIPO MAC'])
                    resultado_solo_nsn = resultado_diferente.dropna(subset=['SN'])

                    # Crear el archivo Excel en memoria
                    output_filtrado = io.BytesIO()
                    with pd.ExcelWriter(output_filtrado, engine='xlsxwriter') as writer:
                        resultado_todos_diferentes.to_excel(writer, sheet_name='Todos los Registros Diferentes', index=False)
                        resultado_solo_equipo_mac.to_excel(writer, sheet_name='Solo en saeplus', index=False)
                        resultado_solo_nsn.to_excel(writer, sheet_name='Solo en olt', index=False)

                    output_filtrado.seek(0)

                    # Redirige a la página de resultados y prepara la descarga
                    return send_file(output_filtrado, download_name="olt_diferente.xlsx", as_attachment=True)

    return render_template('diferentes.html')



#Descargas /////////////////////////////////////////////////////////////
@app.route('/descargar_resultado')
def descargar_resultado():
    global resultado_excel
    if resultado_excel:
        return send_file(resultado_excel, as_attachment=True, download_name='resultado.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)

