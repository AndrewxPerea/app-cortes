import os
from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
import io
from funciones import procesar_excel, procesar_archivo_csv_solo, procesar_archivo_excel_solo, normalizar_columnas
import re
import time

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Variable global para almacenar el archivo Excel resultante
resultado_excel = None

# Home page______________________________________________________

@app.route('/')
def index():
    return render_template('index.html')

# Maysusculas
@app.route('/mayuscula')
def mayuscula():
    return render_template('mayus.html')

# Tutoriales______________________________________________________
@app.route('/tutoriales')
def tutoriales():
    return render_template('tutoriales.html')


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
        df_cortes = pd.read_excel(cortes_file)
        df_abonados = pd.read_excel(abonados_file)
    except Exception as e:
        return render_template('error.html', error=str(e))

    df_cortes = normalizar_columnas(df_cortes, 'Abonados')
    df_abonados = normalizar_columnas(df_abonados, 'Abonados')
    df_cortes.columns = df_cortes.columns.str.lower()
    df_abonados.columns = df_abonados.columns.str.lower()

    df_resultado = pd.merge(df_cortes, df_abonados, on="abonados", how="inner")

    df_resultado = df_resultado[['abonados', 'documento_x',
                                 'nombre_x', 'apellido_x', 'observaciones', 'estatus_y']]
    df_resultado = df_resultado[(df_resultado['observaciones'].isna() | (df_resultado['observaciones'] == '')) &
                                (df_resultado['estatus_y'] == 'ACTIVO')]

    if df_resultado.empty:
        return render_template('exitoso.html')

    # Número de casos encontrados
    num_casos = df_resultado.shape[0]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_resultado.to_excel(writer, index=False, sheet_name='Resultado')
    output.seek(0)

    resultado_excel = output

    time.sleep(3)
    return render_template('resultado.html', data=df_resultado.to_dict(orient='records'), columns=df_resultado.columns, num_casos=num_casos)

# solo @ ________________________________________________________________________


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
            resultado = pd.merge(df_abonados, df_cortes, how='right',
                                 left_on='EQUIPO MACO', right_on='NSN', suffixes=('_abonados', '_cortes'))
            resultado = resultado.dropna(subset=['EQUIPO MACO'])
            resultado.columns = resultado.columns.str.lower()

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                resultado.to_excel(writer, index=False,
                                   sheet_name='Todos los abonados')

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
                    abonados_filtrados.to_excel(
                        writer, index=False, sheet_name='Abonados solo @ Con catv')

            output.seek(0)
            resultado_excel = output
           # Número de casos encontrados
            num_casos = abonados_filtrados.shape[0]

            time.sleep(3)
            return render_template('resultado.html', data=abonados_filtrados.to_dict(orient='records'), columns=abonados_filtrados.columns, num_casos=num_casos)
    return render_template('solointernet.html')

# No activos ___________________________________________________________________


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
            resultado = pd.merge(df_abonados, df_cortes, how='right',
                                 left_on='EQUIPO MACO', right_on='NSN', suffixes=('_abonados', '_cortes'))
            resultado = resultado.dropna(subset=['EQUIPO MACO'])
            resultado.columns = resultado.columns.str.lower()

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                resultado.to_excel(writer, index=False, sheet_name='Resultado')

                abonados_filtrados = resultado[
                    # Filtra los que no sean "estatus activo" o "estatus por instalar"
                    (resultado['estatus'].str.lower().isin(['activo', 'por instalar']) == False) &
                    ((resultado['administrative status'].str.lower() == 'enabled') |
                     (resultado['catv'].str.lower() == 'enabled') |
                        (resultado['status'].str.lower() == 'online'))    # Filtra solo los que tienen "status" como "online"
                ]

                if not abonados_filtrados.empty:
                    abonados_filtrados.to_excel(
                        writer, index=False, sheet_name='Abonados Filtrados')

                columnas_deseadas = [
                    'n° abonado', 'documento', 'nombre', 'apellido',
                    'estatus', 'status', 'equipo maco', 'sn', 'olt',
                    'catv', 'administrative status'
                ]
                abonados_filtrados = abonados_filtrados[columnas_deseadas]

                if not abonados_filtrados.empty:
                    abonados_filtrados.to_excel(
                        writer, index=False, sheet_name='Abonados Filtrados')
              # Número de casos encontrados
            num_casos = abonados_filtrados.shape[0]

            output.seek(0)
            resultado_excel = output

            time.sleep(3)
            return render_template('resultado.html', data=abonados_filtrados.to_dict(orient='records'), columns=abonados_filtrados.columns, num_casos=num_casos)

    return render_template('noactivos.html')
# Cortes____________________________________________________________________________________________


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
            resultado = pd.merge(
                df_saeplus, df_cortes, how='right', left_on='N° Abonado', right_on='N° Abonado')
            resultado = resultado.dropna(subset=['N° Abonado'])
            resultado = pd.merge(resultado, df_olt, left_on='EQUIPO MACO_y',
                                 right_on='NSN', suffixes=('_abonados', '_cortes'))
            resultado = resultado.dropna(subset=['EQUIPO MACO_y'])
            resultado.columns = resultado.columns.str.lower()

            columnas_deseadas = [
                'n° abonado', 'documento_x', 'nombre_x', 'apellido_x',
                'estatus_x', 'observaciones', 'sn', 'olt',
                'catv', 'administrative status', 'status', 'ingeniero'
            ]

            resultado_filtrado = resultado[columnas_deseadas]
            resultado_filtrado = resultado_filtrado[
                (resultado_filtrado['observaciones'].isna()) &
                (resultado_filtrado['estatus_x'] != 'ACTIVO') &
                ((resultado_filtrado['status'] == 'Online') |
                 (resultado_filtrado['catv'] != 'Disabled') |
                 (resultado_filtrado['administrative status'] == 'Enabled'))
            ]
            resultado_filtrado = resultado_filtrado.dropna(
                subset=['estatus_x'])

            output_filtrado = io.BytesIO()
            with pd.ExcelWriter(output_filtrado, engine='xlsxwriter') as writer_filtrado:
                resultado_filtrado.to_excel(
                    writer_filtrado, index=False, sheet_name='Resultado Filtrado')
            output_filtrado.seek(0)
            num_casos = resultado_filtrado.shape[0]
            resultado_excel = output_filtrado

            time.sleep(3)

            return render_template('resultado.html', data=resultado_filtrado.to_dict(orient='records'), columns=resultado_filtrado.columns, num_casos=num_casos)

    return render_template('cortes.html')


# Mensajes/________________________________________________________________________________

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

        time.sleep(3)
    return render_template('upload.html')

# comparador de planes


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
            resultado = pd.merge(saeplus, olt, how='right', left_on='EQUIPO MACO',
                                 right_on='NSN', suffixes=('_abonados', '_OLT'))
            resultado = resultado.dropna(subset=['EQUIPO MACO'])
            resultado.columns = resultado.columns.str.lower()

            # Función para extraer la velocidad

            def extraer_velocidad(detalle):
                match = re.search(r'(\d+)\s*MG', detalle.upper())
                if match:
                    return match.group(1) + 'MG'
                return None

            # Aplicar la extracción de velocidad
            resultado['velocidad_detalle'] = resultado['detalle suscripcion'].apply(
                extraer_velocidad)

            # Filtrar los abonados que no coinciden en velocidad
            abonados_filtrados = resultado[
                (resultado['velocidad_detalle'] != resultado['service port download speed']) &
                (resultado['estatus'] == 'ACTIVO')
            ]

            columnas_deseadas = [
                'n° abonado', 'documento', 'nombre', 'name', 'estatus',
                'detalle suscripcion', 'nombre franquicia', 'equipo maco', 'sn', 'olt',
                'service port upload speed', 'service port download speed', 'tipo tecnología.'
            ]
            abonados_filtrados = abonados_filtrados[columnas_deseadas]
            output_filtrado = io.BytesIO()
            with pd.ExcelWriter(output_filtrado, engine='xlsxwriter') as writer_filtrado:
                abonados_filtrados.to_excel(
                    writer_filtrado, index=False, sheet_name='Resultado Filtrado')
            output_filtrado.seek(0)
            num_casos = abonados_filtrados.shape[0]
            resultado_excel = output_filtrado

            time.sleep(3)
            return render_template('resultado.html', data=abonados_filtrados.to_dict(orient='records'), columns=abonados_filtrados.columns, num_casos=num_casos)

           # Número de casos encontrados

    return render_template('verificar_velocidad.html')


@app.route('/diferentes', methods=['GET', 'POST'])
def diferentes():
    global resultado_excel  # Variable global para almacenar el archivo generado

    if request.method == 'POST':
        saeplus = request.files['saeplus']
        olt = request.files['olt']
        try:
            saeplus = procesar_archivo_excel_solo(saeplus)
            olt2 = procesar_archivo_csv_solo(olt)
        except Exception as e:
            return render_template('error.html', error=str(e))

        if not saeplus.empty and not olt2.empty:
            # Confirmar que existen las columnas 'EQUIPO MACO' y 'NSN'
            if 'EQUIPO MACO' in saeplus.columns and 'NSN' in olt2.columns:
                # Hacer la fusión con 'indicator=True' para identificar los registros coincidentes
                resultado = pd.merge(saeplus, olt2, how='outer', left_on='EQUIPO MACO', right_on='NSN', suffixes=(
                    '_abonados', '_cortes'), indicator=True)

                if '_merge' in resultado.columns:
                    # Filtrar registros donde 'EQUIPO MACO' y 'NSN' no coinciden
                    resultado_diferente = resultado[resultado['_merge'] != 'both']

                    resultado_todos_diferentes = resultado_diferente.dropna(
                        axis=1, how='all')
                    resultado_solo_equipo_mac = resultado_diferente.dropna(subset=[
                                                                           'EQUIPO MACO'])
                    resultado_solo_equipo_mac = resultado_solo_equipo_mac.dropna(
                        axis=1, how='all')
                    resultado_solo_nsn = resultado_diferente.dropna(subset=[
                                                                    'NSN'])
                    resultado_solo_nsn = resultado_solo_nsn.dropna(
                        axis=1, how='all')

                    # Crear el archivo Excel en memoria
                    output_filtrado = io.BytesIO()
                    with pd.ExcelWriter(output_filtrado, engine='xlsxwriter') as writer:
                        resultado_todos_diferentes.to_excel(
                            writer, sheet_name='Todos los Registros Diferentes', index=False)
                        resultado_solo_equipo_mac.to_excel(
                            writer, sheet_name='Solo en saeplus', index=False)
                        resultado_solo_nsn.to_excel(
                            writer, sheet_name='Solo en olt', index=False)

                    output_filtrado.seek(0)
                    resultado_excel = output_filtrado  # Guardar el archivo en la variable global
                    print("Archivo generado correctamente")

                    # Redirigir a la página de resultados
                    # Número de casos encontrados
                    num_casos = resultado_diferente.shape[0]
                    return render_template('resultado.html', data=resultado_diferente.to_dict(orient='records'), columns=resultado_diferente.columns, num_casos=num_casos)

    return render_template('diferentes.html')


@app.route('/auditoria_reconexiones', methods=['GET', 'POST'])
def auditoria_reconexiones():
    global resultado_excel  # Variable global para almacenar el archivo generado

    if request.method == 'POST':
        drive = request.files['drive']
        saeplus = request.files['saeplus']
        epayco = request.files['epayco']
        olt = request.files['smartolt']

        try:
            # Procesamiento de archivos
            df_drive = pd.read_excel(drive)
            df_saeplus = pd.read_excel(saeplus)
            df_epayco = pd.read_excel(epayco)
            df_abonado_cortes = procesar_archivo_excel_solo(saeplus)
            df_olt_cortes = procesar_archivo_csv_solo(olt)
        except Exception as e:
            return render_template('error.html', error=f"Error en el procesamiento: {e}")

        if not df_drive.empty and not df_saeplus.empty and not df_epayco.empty and not df_abonado_cortes.empty and not df_olt_cortes.empty:
            # Normalización de columnas
            df_drive = normalizar_columnas(df_drive, 'abonados')
            df_saeplus = normalizar_columnas(df_saeplus, 'abonados')
            df_epayco = normalizar_columnas(df_epayco, 'abonados')

            # Merge de DataFrames
            df_resultado1 = pd.merge(
                df_drive, df_saeplus, on="abonados", how="inner")
            df_resultado2 = pd.merge(
                df_drive, df_epayco, on="abonados", how="inner")
            df_resultado3 = pd.merge(
                df_abonado_cortes, df_olt_cortes, how='right',
                left_on='EQUIPO MACO', right_on='NSN', suffixes=('_abonados', '_cortes')
            ).dropna(subset=['EQUIPO MACO'])
            df_resultado3.columns = df_resultado3.columns.str.lower()

            # Selección de columnas relevantes
            df_resultado1 = df_resultado1[['abonados', 'documento_x', 'nombre_x',
                                           'apellido_x', 'observaciones', 'estatus_y', 'detalle suscripcion_x']]
            df_resultado2 = df_resultado2[['abonados', 'documento_x', 'nombre',
                                           'apellido', 'observaciones', 'estatus', 'detalle suscripcion']]
            df_resultado3 = df_resultado3[['n° abonado', 'documento', 'nombre', 'apellido',
                                           'estatus', 'status', 'olt', 'catv', 'administrative status', 'detalle suscripcion']]

            reconexiones= df_resultado1[(df_resultado1['observaciones'].isna() | (df_resultado1['observaciones'] == '')) &
                                          (df_resultado1['estatus_y'] == 'ACTIVO')]
            df_resultado1 = df_resultado1[(df_resultado1['estatus_y'] == 'ACTIVO')]

            abonados_epayco = df_resultado2[(df_resultado2['observaciones'].isna() | (
                df_resultado2['observaciones'] == ''))]

            df_resultado3 = df_resultado3[
                (df_resultado3['estatus'].str.lower().isin(['activo'])) &
                ((df_resultado3['administrative status'].str.lower() != 'enabled') |
                 (df_resultado3['catv'].str.lower() != 'enabled')) |
                (df_resultado3['status'].str.lower() != 'online')
            ]

            df_resultado3 = df_resultado3.rename(
                columns={df_resultado3.columns[0]: 'abonados'})
            df_resultado3 = pd.merge(
                df_resultado3, df_resultado1, on="abonados", how="inner")

            desactivado = df_resultado3[
                ((df_resultado3['status'].str.lower() != 'online') |  # Cambiar a df_resultado3
                 ((df_resultado3['detalle suscripcion'].str.contains('@', na=False)) &
                    ((df_resultado3['catv'].str.lower() == 'enabled')) | (df_resultado3['administrative status'].str.lower() != 'enabled')) |
                    (~df_resultado3['detalle suscripcion'].str.contains('@', na=False) &
                     ((df_resultado3['catv'].str.lower() != 'enabled')) | (df_resultado3['administrative status'].str.lower() != 'enabled')))
            ]
            try:
                # Generar el archivo Excel en memoria
                output_filtrado = io.BytesIO()
                with pd.ExcelWriter(output_filtrado, engine='xlsxwriter') as writer:
                    reconexiones.to_excel(
                        writer, sheet_name='Reconexion sin observaciones', index=False)
                    abonados_epayco.to_excel(
                        writer, sheet_name='Pagos de epayco', index=False)
                    desactivado.to_excel(
                        writer, sheet_name='Abonados sin activar', index=False)

                output_filtrado.seek(0)
                resultado_excel = output_filtrado  # Guardar el archivo en la variable global
                print("Archivo generado correctamente")

                # Redirigir a la página de resultados
                num_casos = desactivado.shape[0]  # Número de casos encontrados
                return render_template('resultado.html', data=desactivado.to_dict(orient='records'), columns=desactivado.columns, num_casos=num_casos)

            except Exception as e:
                print(f"Error al enviar el archivo: {e}")
                return render_template('error.html', error=f"Error al enviar el archivo: {e}")

    return render_template('auditoria_reconexiones.html')

# Descargas /////////////////////////////////////////////////////////////


@app.route('/descargar_resultado')
def descargar_resultado():
    global resultado_excel
    if resultado_excel:
        return send_file(resultado_excel, as_attachment=True, download_name='resultado.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True)
