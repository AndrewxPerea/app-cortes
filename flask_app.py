import os
import tempfile

from flask import Flask, render_template, request, send_file, session

from services.atenuaciones import procesar_atenuaciones
from services.auditoria_reconexiones import procesar_auditoria_reconexiones
from services.cortes import procesar_cortes
from services.diferentes import procesar_diferentes
from services.navegacion import procesar_sin_navegar
from services.reconexiones import procesar_reconexiones
from services.upload_validation import validar_archivos_requeridos
from services.velocidad import procesar_verificacion_velocidad

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'dev-secret-key-change-me')
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def limpiar_resultado_sesion():
    ruta_anterior = session.pop('resultado_excel_path', None)
    if ruta_anterior and os.path.exists(ruta_anterior):
        os.remove(ruta_anterior)


def guardar_resultado_excel(output):
    limpiar_resultado_sesion()
    output.seek(0)
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
        temp_file.write(output.getvalue())
        session['resultado_excel_path'] = temp_file.name


def renderizar_resultado(resultado, columns=None):
    data = resultado['data']
    return render_template(
        'resultado.html',
        data=data.to_dict(orient='records'),
        columns=columns or resultado.get('columns', data.columns),
        num_casos=resultado['num_casos']
    )


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/reconexiones')
def reconexiones():
    return render_template('reconexiones.html')


@app.route('/procesar', methods=['POST'])
def procesar_archivos():
    try:
        archivos = validar_archivos_requeridos(
            request.files,
            [
                ('abonados', {'.xlsx', '.xls'}, 'de abonados SAEPlus'),
                ('cortes', {'.xlsx', '.xls'}, 'de cortes o workdrive'),
            ]
        )
        resultado = procesar_reconexiones(
            archivos['abonados'],
            archivos['cortes']
        )
    except Exception as e:
        return render_template('error.html', error=str(e))

    if resultado['data'].empty:
        return render_template('exitoso.html')

    guardar_resultado_excel(resultado['excel'])
    return renderizar_resultado(resultado)


@app.route('/cortes', methods=['GET', 'POST'])
def cortes():
    if request.method == 'POST':
        try:
            archivos = validar_archivos_requeridos(
                request.files,
                [
                    ('abonados', {'.xlsx', '.xls'}, 'de corte de abonados'),
                    ('cortes', {'.csv'}, 'de SmartOLT'),
                    ('asaeplus', {'.xlsx', '.xls'}, 'de abonados SAEPlus'),
                ]
            )
            resultado = procesar_cortes(
                archivos['abonados'],
                archivos['cortes'],
                archivos['asaeplus']
            )
        except Exception as e:
            return render_template('error.html', error=str(e))

        guardar_resultado_excel(resultado['excel'])
        return renderizar_resultado(resultado)

    return render_template('cortes.html')


@app.route('/verificar_velocidad', methods=['GET', 'POST'])
def verificar_velocidad():
    if request.method == 'POST':
        try:
            archivos = validar_archivos_requeridos(
                request.files,
                [
                    ('saeplus', {'.xlsx', '.xls'}, 'SAEPlus'),
                    ('olt', {'.csv'}, 'de SmartOLT'),
                ]
            )
            resultado = procesar_verificacion_velocidad(
                archivos['saeplus'],
                archivos['olt']
            )
        except Exception as e:
            return render_template('error.html', error=str(e))

        guardar_resultado_excel(resultado['excel'])
        return renderizar_resultado(resultado)

    return render_template('verificar_velocidad.html')


@app.route('/diferentes', methods=['GET', 'POST'])
def diferentes():
    if request.method == 'POST':
        try:
            archivos = validar_archivos_requeridos(
                request.files,
                [
                    ('saeplus', {'.xlsx', '.xls'}, 'SAEPlus'),
                    ('olt', {'.csv'}, 'de SmartOLT'),
                ]
            )
            resultado = procesar_diferentes(
                archivos['saeplus'],
                archivos['olt']
            )
        except Exception as e:
            return render_template('error.html', error=str(e))

        guardar_resultado_excel(resultado['excel'])
        return renderizar_resultado(resultado)

    return render_template('diferentes.html')


@app.route('/sin_navegar', methods=['GET', 'POST'])
def sin_navegar():
    if request.method == 'POST':
        try:
            archivos = validar_archivos_requeridos(
                request.files,
                [
                    ('abonados', {'.xlsx', '.xls'}, 'de abonados SAEPlus'),
                    ('olt', {'.csv'}, 'de SmartOLT'),
                ]
            )
            resultado = procesar_sin_navegar(
                archivos['abonados'],
                archivos['olt']
            )
        except Exception as e:
            return render_template('error.html', error=f"Error en el procesamiento: {e}")

        guardar_resultado_excel(resultado['excel'])
        return renderizar_resultado(resultado)

    return render_template('sin_navegar.html')


@app.route('/auditoria_reconexiones', methods=['GET', 'POST'])
def auditoria_reconexiones():
    if request.method == 'POST':
        try:
            archivos = validar_archivos_requeridos(
                request.files,
                [
                    ('drive', {'.xlsx', '.xls'}, 'de drive o workdrive'),
                    ('saeplus', {'.xlsx', '.xls'}, 'SAEPlus'),
                    ('epayco', {'.xlsx', '.xls'}, 'de ePayco'),
                    ('smartolt', {'.csv'}, 'de SmartOLT'),
                ]
            )
            resultado = procesar_auditoria_reconexiones(
                archivos['drive'],
                archivos['saeplus'],
                archivos['epayco'],
                archivos['smartolt']
            )
        except Exception as e:
            return render_template('error.html', error=f"Error en el procesamiento: {e}")

        guardar_resultado_excel(resultado['excel'])
        return renderizar_resultado(resultado)

    return render_template('auditoria_reconexiones.html')


@app.route('/atenuaciones', methods=['GET', 'POST'])
def atenuaciones():
    if request.method == 'POST':
        try:
            archivos = validar_archivos_requeridos(
                request.files,
                [
                    ('olt_csv', {'.csv'}, 'de OLT'),
                ]
            )
            resultado = procesar_atenuaciones(archivos['olt_csv'])
        except Exception as e:
            return render_template('error.html', error=f"Error en el procesamiento: {e}")

        guardar_resultado_excel(resultado['excel'])
        return renderizar_resultado(resultado)

    return render_template('atenuaciones.html')


@app.route('/descargar_resultado')
def descargar_resultado():
    resultado_excel_path = session.get('resultado_excel_path')
    if not resultado_excel_path:
        return render_template('error.html', error="No hay archivo para descargar.")
    if os.path.exists(resultado_excel_path):
        return send_file(
            resultado_excel_path,
            as_attachment=True,
            download_name='resultado.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    session.pop('resultado_excel_path', None)
    return render_template('error.html', error="El archivo generado ya no está disponible. Vuelve a procesarlo.")


if __name__ == '__main__':
    app.run(debug=True)
