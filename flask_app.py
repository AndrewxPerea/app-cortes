import os
import tempfile
import unicodedata
from datetime import datetime

from flask import Flask, render_template, request, send_file, session

from services.atenuaciones import procesar_atenuaciones
from services.auditoria_reconexiones import procesar_auditoria_reconexiones
from services.comparativo_precintos import procesar_comparativo_precintos
from services.coincidencia_en_fila import procesar_coincidencia_en_fila
from services.coincidencias_saeplus_smartolt import procesar_coincidencias_saeplus_smartolt
from services.cortes import procesar_cortes
from services.diferentes import procesar_diferentes
from services.navegacion import (
    procesar_navegacion_activos_sin_catv,
    procesar_navegacion_activos_sin_navegar,
    procesar_navegacion_desactivos_con_internet,
    procesar_navegacion_solo_con_arroba_y_catv_activo,
    procesar_sin_navegar,
)
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
    session.pop('resultado_excel_name', None)
    if ruta_anterior and os.path.exists(ruta_anterior):
        os.remove(ruta_anterior)


def normalizar_nombre_archivo(texto):
    texto = unicodedata.normalize('NFKD', texto)
    texto = ''.join(char for char in texto if not unicodedata.combining(char))
    texto = texto.lower().strip()
    texto = texto.replace('@', 'arroba')
    texto = ''.join(char if char.isalnum() else '_' for char in texto)
    while '__' in texto:
        texto = texto.replace('__', '_')
    return texto.strip('_') or 'resultado'


def construir_nombre_descarga(nombre_base):
    fecha = datetime.now().strftime('%Y-%m-%d')
    return f"{normalizar_nombre_archivo(nombre_base)}_{fecha}.xlsx"


def guardar_resultado_excel(output, nombre_descarga=None):
    limpiar_resultado_sesion()
    output.seek(0)
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
        temp_file.write(output.getvalue())
        session['resultado_excel_path'] = temp_file.name
        session['resultado_excel_name'] = nombre_descarga or construir_nombre_descarga('resultado')


def renderizar_resultado(resultado, columns=None):
    data = resultado['data']
    return render_template(
        'resultado.html',
        data=data.to_dict(orient='records'),
        columns=columns or resultado.get('columns', data.columns),
        num_casos=resultado['num_casos'],
        nombre_descarga=session.get('resultado_excel_name', 'resultado.xlsx')
    )


def guardar_y_renderizar_resultado(resultado, nombre_descarga_base):
    guardar_resultado_excel(
        resultado['excel'],
        construir_nombre_descarga(nombre_descarga_base)
    )
    return renderizar_resultado(resultado)


def ejecutar_procesamiento_navegacion(procesador):
    archivos = validar_archivos_requeridos(
        request.files,
        [
            ('abonados', {'.xlsx', '.xls'}, 'de abonados SAEPlus'),
            ('olt', {'.csv'}, 'de SmartOLT'),
        ]
    )
    return procesador(
        archivos['abonados'],
        archivos['olt']
    )


def renderizar_formulario_analisis(
    titulo,
    descripcion,
    endpoint,
    nombre_descarga_base,
    criterios,
    resultado_esperado,
    archivos_requeridos,
    pasos,
    campos_archivo,
    volver_url='/',
    volver_texto='Volver al inicio',
    enlaces_relacionados=None,
):
    return render_template(
        'navegacion_individual.html',
        titulo=titulo,
        descripcion=descripcion,
        action_url=endpoint,
        criterios=criterios,
        resultado_esperado=resultado_esperado,
        nombre_descarga=construir_nombre_descarga(nombre_descarga_base),
        archivos_requeridos=archivos_requeridos,
        pasos=pasos,
        campos_archivo=campos_archivo,
        volver_url=volver_url,
        volver_texto=volver_texto,
        enlaces_relacionados=enlaces_relacionados or [],
    )


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/reconexiones')
def reconexiones():
    return renderizar_formulario_analisis(
        'Auditoría de Reconexiones',
        'Cruza el archivo de Workdrive con el archivo de abonados de SAEPlus para identificar reconexiones activas que todavía no tienen observación registrada.',
        '/procesar',
        'Auditoria de reconexiones',
        [
            'Coincidencia por número de abonado entre Workdrive y SAEPlus.',
            'Observaciones vacías o sin diligenciar en el archivo de cortes.',
            'Estado ACTIVO del abonado en SAEPlus.',
        ],
        'Se genera un Excel con los abonados que aparecen como reconectados o activos, pero aún no tienen observación en el control operativo.',
        [
            'Archivo de abonados exportado desde SAEPlus en formato Excel.',
            'Archivo de cortes o Workdrive en formato Excel.',
        ],
        [
            'Exporta el archivo de abonados desde SAEPlus.',
            'Exporta el archivo de cortes o reconexiones desde Workdrive.',
            'Abre ambos archivos y guárdalos nuevamente antes de cargarlos.',
            'Carga primero el archivo de SAEPlus y luego el archivo de Workdrive.',
            'Ejecuta el análisis y revisa el Excel resultante para gestión y seguimiento.',
        ],
        [
            {'id': 'abonados', 'name': 'abonados', 'label': 'Archivo de Abonados (SAEPlus)', 'accept': '.xlsx,.xls'},
            {'id': 'cortes', 'name': 'cortes', 'label': 'Archivo de Cortes o Workdrive', 'accept': '.xlsx,.xls'},
        ],
    )


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

    return guardar_y_renderizar_resultado(resultado, 'Auditoria de reconexiones')


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

        return guardar_y_renderizar_resultado(resultado, 'Auditoria de abonados en corte')

    return renderizar_formulario_analisis(
        'Auditoría de Abonados en Corte',
        'Detecta abonados con estado comercial inactivo o en corte que todavía presentan servicio operativo en la red.',
        '/cortes',
        'Auditoria de abonados en corte',
        [
            'Abonados sin observación registrada en el archivo de corte.',
            'Estados diferentes de ACTIVO en la base analizada.',
            'Equipos con estado Online, CATV habilitado o Administrative Status habilitado en SmartOLT.',
        ],
        'Se genera un Excel con los abonados que deben revisarse porque siguen con servicio en red aun cuando el estado comercial indica corte o suspensión.',
        [
            'Archivo de corte de abonados o Workdrive en formato Excel.',
            'Archivo exportado desde SmartOLT en formato CSV.',
            'Archivo de abonados exportado desde SAEPlus en formato Excel.',
        ],
        [
            'Exporta el listado de abonados en corte desde Workdrive o tu fuente operativa.',
            'Exporta el listado de abonados desde SmartOLT en formato CSV.',
            'Exporta el listado actualizado de abonados desde SAEPlus.',
            'Abre los archivos Excel y guárdalos nuevamente antes de cargarlos.',
            'Carga los tres archivos en el orden indicado y ejecuta el análisis.',
        ],
        [
            {'id': 'abonados', 'name': 'abonados', 'label': 'Archivo de Corte de Abonados (Excel)', 'accept': '.xlsx,.xls'},
            {'id': 'cortes', 'name': 'cortes', 'label': 'Archivo de Abonados en SmartOLT (CSV)', 'accept': '.csv'},
            {'id': 'asaeplus', 'name': 'asaeplus', 'label': 'Lista de Abonados SAEPlus (Excel)', 'accept': '.xlsx,.xls'},
        ],
    )


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

        return guardar_y_renderizar_resultado(resultado, 'Verificar velocidad de abonados')

    return renderizar_formulario_analisis(
        'Verificar Velocidad de Abonados',
        'Compara la velocidad definida en el detalle de suscripción de SAEPlus contra la velocidad configurada en SmartOLT y además valida que los planes que contienen la palabra solo no tengan CATV activo.',
        '/verificar_velocidad',
        'Verificar velocidad de abonados',
        [
            'Abonados con estado ACTIVO en SAEPlus.',
            'Extracción automática de la velocidad comercial desde el detalle de suscripción.',
            'Diferencias entre la velocidad comercial y la velocidad de descarga configurada en SmartOLT.',
            'Planes cuyo detalle de suscripción contiene la palabra solo y que no deberían tener CATV habilitado.',
        ],
        'Se genera un Excel con los abonados activos cuyo plan comercial no coincide con la velocidad técnica o que tienen CATV activo cuando el detalle de suscripción indica un plan solo internet.',
        [
            'Archivo de abonados exportado desde SAEPlus en formato Excel.',
            'Archivo de equipos o abonados exportado desde SmartOLT en formato CSV.',
        ],
        [
            'Exporta el archivo de abonados desde SAEPlus.',
            'Exporta el archivo desde SmartOLT en formato CSV.',
            'Abre el archivo de SAEPlus y guárdalo nuevamente antes de cargarlo.',
            'Carga ambos archivos y ejecuta el análisis.',
            'Valida el Excel generado para corregir diferencias de velocidad.',
        ],
        [
            {'id': 'saeplus', 'name': 'saeplus', 'label': 'Archivo SAEPlus (Excel)', 'accept': '.xlsx,.xls'},
            {'id': 'olt', 'name': 'olt', 'label': 'Archivo OLT (CSV)', 'accept': '.csv'},
        ],
    )


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

        return guardar_y_renderizar_resultado(resultado, 'Equipos que no coinciden')

    return renderizar_formulario_analisis(
        'Equipos que No Coinciden',
        'Compara los equipos registrados en SAEPlus contra los equipos reportados por SmartOLT para detectar diferencias entre ambas fuentes.',
        '/diferentes',
        'Equipos que no coinciden',
        [
            'Equipos presentes en SAEPlus pero ausentes en SmartOLT.',
            'Equipos presentes en SmartOLT pero ausentes en SAEPlus.',
            'Cruce por los identificadores EQUIPO MACO y NSN.',
        ],
        'Se genera un Excel con una hoja general de diferencias y hojas separadas para equipos solo en SAEPlus y solo en SmartOLT.',
        [
            'Archivo de abonados exportado desde SAEPlus en formato Excel.',
            'Archivo de abonados exportado desde SmartOLT en formato CSV.',
        ],
        [
            'Exporta el archivo de abonados desde SAEPlus.',
            'Exporta el archivo de abonados desde SmartOLT en formato CSV.',
            'Abre el Excel de SAEPlus y guárdalo nuevamente.',
            'Carga ambos archivos y ejecuta el análisis.',
            'Revisa el Excel descargado para depurar diferencias entre inventarios.',
        ],
        [
            {'id': 'saeplus', 'name': 'saeplus', 'label': 'Archivo de Abonados (Excel)', 'accept': '.xlsx,.xls'},
            {'id': 'olt', 'name': 'olt', 'label': 'Archivo de Abonados en SmartOLT (CSV)', 'accept': '.csv'},
        ],
    )


@app.route('/coincidencias_saeplus_smartolt', methods=['GET', 'POST'])
def coincidencias_saeplus_smartolt():
    if request.method == 'POST':
        try:
            archivos = validar_archivos_requeridos(
                request.files,
                [
                    ('saeplus', {'.xlsx', '.xls'}, 'SAEPlus'),
                    ('olt', {'.csv'}, 'de SmartOLT'),
                ]
            )
            resultado = procesar_coincidencias_saeplus_smartolt(
                archivos['saeplus'],
                archivos['olt']
            )
        except Exception as e:
            return render_template('error.html', error=str(e))

        return guardar_y_renderizar_resultado(resultado, 'Equipos que coinciden')

    return renderizar_formulario_analisis(
        'Equipos que Coinciden',
        'Cruza SAEPlus y SmartOLT para devolver únicamente los registros que sí coinciden entre ambas fuentes.',
        '/coincidencias_saeplus_smartolt',
        'Equipos que coinciden',
        [
            'Coincidencia exacta entre los campos EQUIPO MACO de SAEPlus y NSN de SmartOLT.',
            'Conservación de la información de ambas fuentes en una sola tabla.',
            'Exclusión de registros que existan solo en una de las dos bases.',
        ],
        'Se genera un Excel con los equipos y abonados que sí empatan entre SAEPlus y SmartOLT.',
        [
            'Archivo de abonados exportado desde SAEPlus en formato Excel.',
            'Archivo de abonados exportado desde SmartOLT en formato CSV.',
        ],
        [
            'Exporta el archivo de abonados desde SAEPlus.',
            'Exporta el archivo de abonados desde SmartOLT en formato CSV.',
            'Abre el Excel de SAEPlus y guárdalo nuevamente antes de cargarlo.',
            'Carga ambos archivos en el formulario.',
            'Ejecuta el análisis y descarga el Excel con las coincidencias encontradas.',
        ],
        [
            {'id': 'saeplus', 'name': 'saeplus', 'label': 'Archivo de Abonados (Excel)', 'accept': '.xlsx,.xls'},
            {'id': 'olt', 'name': 'olt', 'label': 'Archivo de Abonados en SmartOLT (CSV)', 'accept': '.csv'},
        ],
    )


@app.route('/comparativo_precintos', methods=['GET', 'POST'])
def comparativo_precintos():
    if request.method == 'POST':
        try:
            archivos = validar_archivos_requeridos(
                request.files,
                [
                    ('saeplus', {'.xlsx', '.xls'}, 'SAEPlus'),
                    ('olt', {'.csv'}, 'de SmartOLT'),
                ]
            )
            resultado = procesar_comparativo_precintos(
                archivos['saeplus'],
                archivos['olt']
            )
        except Exception as e:
            return render_template('error.html', error=str(e))

        return guardar_y_renderizar_resultado(resultado, 'Comparativo de precintos')

    return renderizar_formulario_analisis(
        'Comparativo de Precintos',
        'Compara SAEPlus y SmartOLT para identificar qué abonados coinciden y cuáles no coinciden, ayudando a revisar estados y detectar posibles precintos perdidos.',
        '/comparativo_precintos',
        'Comparativo de precintos',
        [
            'Cruce entre SAEPlus y SmartOLT por los campos EQUIPO MACO y NSN.',
            'Se conservan únicamente los abonados que coinciden entre SAEPlus y SmartOLT.',
            'Visualización del estatus comercial de SAEPlus y del status técnico de SmartOLT en el mismo comparativo.',
            'Resaltado de los abonados de SAEPlus cuyo campo precinto está vacío.',
            'Orden alfabético por la columna estatus y, como segundo criterio, por la columna precinto de menor a mayor.',
        ],
        'Se genera un Excel con una sola pestaña llamada Coinciden, enfocada únicamente en los abonados coincidentes y ordenada por estatus y precinto.',
        [
            'Archivo de abonados exportado desde SAEPlus en formato Excel.',
            'Archivo de abonados exportado desde SmartOLT en formato CSV.',
        ],
        [
            'Exporta el archivo de abonados desde SAEPlus incluyendo la columna precinto.',
            'Exporta el archivo de abonados desde SmartOLT en formato CSV.',
            'Abre el Excel de SAEPlus y guárdalo nuevamente antes de cargarlo.',
            'Carga ambos archivos en el formulario.',
            'Ejecuta el análisis y revisa la pestaña Coinciden del Excel resultante para validar estados, precintos vacíos y organización del listado.',
        ],
        [
            {'id': 'saeplus', 'name': 'saeplus', 'label': 'Archivo de Abonados SAEPlus (Excel)', 'accept': '.xlsx,.xls'},
            {'id': 'olt', 'name': 'olt', 'label': 'Archivo de Abonados SmartOLT (CSV)', 'accept': '.csv'},
        ],
    )


@app.route('/sin_navegar', methods=['GET', 'POST'])
def sin_navegar():
    if request.method == 'POST':
        try:
            resultado = ejecutar_procesamiento_navegacion(procesar_sin_navegar)
        except Exception as e:
            return render_template('error.html', error=f"Error en el procesamiento: {e}")

        return guardar_y_renderizar_resultado(resultado, 'Auditoria general de abonados')

    return renderizar_formulario_analisis(
        'Auditoría General de Navegación',
        'Genera un consolidado con cuatro revisiones técnicas y comerciales para detectar inconsistencias de navegación, CATV y estado de servicio.',
        '/sin_navegar',
        'Auditoria general de abonados',
        [
            'Abonados activos que no están navegando correctamente.',
            'Abonados desactivos que siguen con servicio online.',
            'Abonados activos sin CATV cuando deberían tenerlo.',
            'Abonados con planes especiales con @ que conservan CATV activo.',
        ],
        'Se descarga un solo Excel con cuatro hojas, una por cada análisis de navegación incluido en la auditoría general.',
        [
            'Archivo de abonados exportado desde SAEPlus en formato Excel.',
            'Archivo de abonados exportado desde SmartOLT en formato CSV.',
        ],
        [
            'Exporta el archivo de abonados desde SAEPlus para la franquicia que deseas revisar.',
            'Abre el Excel de SAEPlus y guárdalo nuevamente antes de cargarlo.',
            'Exporta el archivo de abonados desde SmartOLT en formato CSV.',
            'Carga ambos archivos y ejecuta la auditoría general.',
            'Descarga el Excel consolidado o utiliza los accesos directos a los análisis individuales si necesitas trabajar un caso específico.',
        ],
        [
            {'id': 'abonados', 'name': 'abonados', 'label': 'Archivo de Abonados (Excel)', 'accept': '.xlsx,.xls'},
            {'id': 'olt', 'name': 'olt', 'label': 'Archivo de Abonados en SmartOLT (CSV)', 'accept': '.csv'},
        ],
        enlaces_relacionados=[
            {'href': '/navegacion_activos_sin_navegar', 'label': 'Activos sin navegar'},
            {'href': '/navegacion_desactivos_con_internet', 'label': 'Desactivos con internet'},
            {'href': '/navegacion_activos_sin_catv', 'label': 'Activos sin CATV'},
            {'href': '/navegacion_solo_con_arroba_y_catv_activo', 'label': 'Solo con @ y CATV activo'},
        ],
    )


@app.route('/navegacion_activos_sin_navegar', methods=['GET', 'POST'])
def navegacion_activos_sin_navegar():
    if request.method == 'POST':
        try:
            resultado = ejecutar_procesamiento_navegacion(procesar_navegacion_activos_sin_navegar)
        except Exception as e:
            return render_template('error.html', error=f"Error en el procesamiento: {e}")

        guardar_resultado_excel(
            resultado['excel'],
            construir_nombre_descarga('Activos sin navegar')
        )
        return renderizar_resultado(resultado)

    return renderizar_formulario_analisis(
        'Activos sin navegar',
        'Identifica abonados en estado ACTIVO que tienen el servicio administrativo deshabilitado o no aparecen online en SmartOLT.',
        '/navegacion_activos_sin_navegar',
        'Activos sin navegar',
        [
            'Abonados con estado ACTIVO en SAEPlus.',
            'Casos donde el campo Administrative Status aparece en Disabled.',
            'Casos donde el abonado no figura como Online en SmartOLT.',
        ],
        'Se genera un Excel con los abonados activos que requieren revisión porque no están navegando correctamente en red.',
        [
            'Archivo de abonados exportado desde SAEPlus en formato Excel.',
            'Archivo de abonados exportado desde SmartOLT en formato CSV.',
        ],
        [
            'Exporta el archivo de abonados desde SAEPlus para la franquicia que deseas revisar.',
            'Abre el Excel de SAEPlus y guárdalo nuevamente antes de cargarlo en la herramienta.',
            'Exporta el archivo de abonados desde SmartOLT en formato CSV.',
            'Carga ambos archivos y ejecuta el análisis.',
            'Revisa la tabla en pantalla y descarga el Excel generado para el seguimiento operativo.',
        ],
        [
            {'id': 'abonados', 'name': 'abonados', 'label': 'Archivo de Abonados (Excel)', 'accept': '.xlsx,.xls'},
            {'id': 'olt', 'name': 'olt', 'label': 'Archivo de Abonados en SmartOLT (CSV)', 'accept': '.csv'},
        ],
        volver_url='/sin_navegar',
        volver_texto='Volver a la auditoría general'
    )


@app.route('/navegacion_desactivos_con_internet', methods=['GET', 'POST'])
def navegacion_desactivos_con_internet():
    if request.method == 'POST':
        try:
            resultado = ejecutar_procesamiento_navegacion(procesar_navegacion_desactivos_con_internet)
        except Exception as e:
            return render_template('error.html', error=f"Error en el procesamiento: {e}")

        guardar_resultado_excel(
            resultado['excel'],
            construir_nombre_descarga('Desactivos con internet')
        )
        return renderizar_resultado(resultado)

    return renderizar_formulario_analisis(
        'Desactivos con internet',
        'Muestra abonados que no están ACTIVO ni POR INSTALAR en SAEPlus, pero siguen apareciendo online en SmartOLT.',
        '/navegacion_desactivos_con_internet',
        'Desactivos con internet',
        [
            'Abonados cuyo estado en SAEPlus es diferente de ACTIVO y POR INSTALAR.',
            'Casos donde el equipo sigue apareciendo Online en SmartOLT.',
            'Situaciones que pueden indicar servicio activo en red para un abonado deshabilitado comercialmente.',
        ],
        'Se descarga un Excel con los abonados que deben validarse por posible inconsistencia entre el estado comercial y el estado técnico.',
        [
            'Archivo de abonados exportado desde SAEPlus en formato Excel.',
            'Archivo de abonados exportado desde SmartOLT en formato CSV.',
        ],
        [
            'Exporta el archivo de abonados desde SAEPlus para la franquicia que deseas revisar.',
            'Abre el Excel de SAEPlus y guárdalo nuevamente antes de cargarlo en la herramienta.',
            'Exporta el archivo de abonados desde SmartOLT en formato CSV.',
            'Carga ambos archivos y ejecuta el análisis.',
            'Valida el Excel generado para revisar posibles servicios activos en abonados deshabilitados.',
        ],
        [
            {'id': 'abonados', 'name': 'abonados', 'label': 'Archivo de Abonados (Excel)', 'accept': '.xlsx,.xls'},
            {'id': 'olt', 'name': 'olt', 'label': 'Archivo de Abonados en SmartOLT (CSV)', 'accept': '.csv'},
        ],
        volver_url='/sin_navegar',
        volver_texto='Volver a la auditoría general'
    )


@app.route('/navegacion_activos_sin_catv', methods=['GET', 'POST'])
def navegacion_activos_sin_catv():
    if request.method == 'POST':
        try:
            resultado = ejecutar_procesamiento_navegacion(procesar_navegacion_activos_sin_catv)
        except Exception as e:
            return render_template('error.html', error=f"Error en el procesamiento: {e}")

        guardar_resultado_excel(
            resultado['excel'],
            construir_nombre_descarga('Activos sin CATV')
        )
        return renderizar_resultado(resultado)

    return renderizar_formulario_analisis(
        'Activos sin CATV',
        'Lista abonados activos con planes sin arroba (@) que tienen CATV deshabilitado.',
        '/navegacion_activos_sin_catv',
        'Activos sin CATV',
        [
            'Abonados con estado ACTIVO en SAEPlus.',
            'Planes que no contienen el símbolo @ en el detalle de suscripción.',
            'Casos donde CATV figura como Disabled en SmartOLT.',
        ],
        'Se genera un Excel con los abonados activos que deberían tener CATV habilitado según su plan, pero en red aparecen sin ese servicio.',
        [
            'Archivo de abonados exportado desde SAEPlus en formato Excel.',
            'Archivo de abonados exportado desde SmartOLT en formato CSV.',
        ],
        [
            'Exporta el archivo de abonados desde SAEPlus para la franquicia que deseas revisar.',
            'Abre el Excel de SAEPlus y guárdalo nuevamente antes de cargarlo en la herramienta.',
            'Exporta el archivo de abonados desde SmartOLT en formato CSV.',
            'Carga ambos archivos y ejecuta el análisis.',
            'Descarga el Excel para validar ajustes de CATV sobre abonados activos.',
        ],
        [
            {'id': 'abonados', 'name': 'abonados', 'label': 'Archivo de Abonados (Excel)', 'accept': '.xlsx,.xls'},
            {'id': 'olt', 'name': 'olt', 'label': 'Archivo de Abonados en SmartOLT (CSV)', 'accept': '.csv'},
        ],
        volver_url='/sin_navegar',
        volver_texto='Volver a la auditoría general'
    )


@app.route('/navegacion_solo_con_arroba_y_catv_activo', methods=['GET', 'POST'])
def navegacion_solo_con_arroba_y_catv_activo():
    if request.method == 'POST':
        try:
            resultado = ejecutar_procesamiento_navegacion(procesar_navegacion_solo_con_arroba_y_catv_activo)
        except Exception as e:
            return render_template('error.html', error=f"Error en el procesamiento: {e}")

        guardar_resultado_excel(
            resultado['excel'],
            construir_nombre_descarga('Solo con arroba y CATV activo')
        )
        return renderizar_resultado(resultado)

    return renderizar_formulario_analisis(
        'Solo con @ y CATV activo',
        'Encuentra abonados con planes que contienen arroba (@) y que conservan CATV activo.',
        '/navegacion_solo_con_arroba_y_catv_activo',
        'Solo con arroba y CATV activo',
        [
            'Abonados cuyo detalle de suscripción contiene el símbolo @.',
            'Casos donde CATV aparece como Enabled en SmartOLT.',
            'Situaciones que pueden requerir revisión de la configuración del servicio contratado.',
        ],
        'Se descarga un Excel con los abonados cuyo plan indica una condición especial con @ y que actualmente mantienen CATV activo.',
        [
            'Archivo de abonados exportado desde SAEPlus en formato Excel.',
            'Archivo de abonados exportado desde SmartOLT en formato CSV.',
        ],
        [
            'Exporta el archivo de abonados desde SAEPlus para la franquicia que deseas revisar.',
            'Abre el Excel de SAEPlus y guárdalo nuevamente antes de cargarlo en la herramienta.',
            'Exporta el archivo de abonados desde SmartOLT en formato CSV.',
            'Carga ambos archivos y ejecuta el análisis.',
            'Revisa el resultado para validar casos con planes especiales y CATV habilitado.',
        ],
        [
            {'id': 'abonados', 'name': 'abonados', 'label': 'Archivo de Abonados (Excel)', 'accept': '.xlsx,.xls'},
            {'id': 'olt', 'name': 'olt', 'label': 'Archivo de Abonados en SmartOLT (CSV)', 'accept': '.csv'},
        ],
        volver_url='/sin_navegar',
        volver_texto='Volver a la auditoría general'
    )


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

        return guardar_y_renderizar_resultado(resultado, 'Auditoria reconexiones con epayco')

    return renderizar_formulario_analisis(
        'Auditoría de Reconexiones con ePayco',
        'Cruza información de Workdrive, SAEPlus, ePayco y SmartOLT para auditar reconexiones, pagos y activaciones pendientes.',
        '/auditoria_reconexiones',
        'Auditoria reconexiones con epayco',
        [
            'Reconexiones sin observaciones registradas.',
            'Pagos detectados en ePayco que requieren contraste con SAEPlus.',
            'Abonados activos que aún no están correctamente habilitados en red.',
            'Cruce operativo entre plataformas comercial, financiera y técnica.',
        ],
        'Se genera un Excel con varias hojas para revisar reconexiones sin observación, pagos ePayco, abonados sin activar y pagos detectados solo en SAEPlus.',
        [
            'Archivo de Drive o Workdrive en formato Excel.',
            'Archivo de abonados exportado desde SAEPlus en formato Excel.',
            'Archivo de pagos o abonados exportado desde ePayco en formato Excel.',
            'Archivo de abonados exportado desde SmartOLT en formato CSV.',
        ],
        [
            'Exporta los archivos de Workdrive, SAEPlus y ePayco en formato Excel.',
            'Exporta el archivo de SmartOLT en formato CSV.',
            'Abre los archivos Excel y guárdalos nuevamente antes de cargarlos.',
            'Carga los cuatro archivos en el formulario respetando cada campo.',
            'Ejecuta el análisis y revisa cada hoja del Excel para seguimiento por área.',
        ],
        [
            {'id': 'drive', 'name': 'drive', 'label': 'Archivo de Drive o Workdrive', 'accept': '.xlsx,.xls'},
            {'id': 'saeplus', 'name': 'saeplus', 'label': 'Archivo de Abonados SAEPlus', 'accept': '.xlsx,.xls'},
            {'id': 'epayco', 'name': 'epayco', 'label': 'Archivo de ePayco', 'accept': '.xlsx,.xls'},
            {'id': 'smartolt', 'name': 'smartolt', 'label': 'Archivo de Abonados SmartOLT', 'accept': '.csv'},
        ],
    )


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

        return guardar_y_renderizar_resultado(resultado, 'Auditoria de atenuaciones')

    return renderizar_formulario_analisis(
        'Auditoría de Atenuaciones',
        'Analiza señales ópticas por OLT, board y port para construir una agenda priorizada de daños o degradaciones en red.',
        '/atenuaciones',
        'Auditoria de atenuaciones',
        [
            'Promedios de señal 1310 y 1490 por puerto.',
            'Clasificación de puertos en niveles de prioridad alta, media o baja.',
            'Separación por tipo de daño dominante: fibra, energía, mixto o sin corte.',
        ],
        'Se genera un Excel con una agenda general priorizada y hojas adicionales por tipo de daño detectado.',
        [
            'Archivo exportado desde la OLT en formato CSV.',
        ],
        [
            'Exporta el archivo CSV desde la OLT con el detalle de ONUs.',
            'Verifica que el CSV contenga columnas como OLT, Board, Port, Signal 1310, Signal 1490, Status y Address.',
            'Carga el archivo en el formulario.',
            'Ejecuta el análisis.',
            'Descarga el Excel para priorizar la atención operativa por puerto y tipo de daño.',
        ],
        [
            {'id': 'olt_csv', 'name': 'olt_csv', 'label': 'Archivo OLT (CSV)', 'accept': '.csv'},
        ],
    )


@app.route('/coincidencia_en_fila', methods=['GET', 'POST'])
def coincidencia_en_fila():
    if request.method == 'POST':
        try:
            archivos = validar_archivos_requeridos(
                request.files,
                [
                    ('archivo_excel', {'.xlsx', '.xls'}, 'Excel con las columnas ABONADO y n° abonado'),
                ]
            )
            resultado = procesar_coincidencia_en_fila(archivos['archivo_excel'])
        except Exception as e:
            return render_template('error.html', error=f"Error en el procesamiento: {e}")

        return guardar_y_renderizar_resultado(resultado, 'Coincidencia en fila')

    return renderizar_formulario_analisis(
        'Coincidencia en Fila',
        'Revisa un archivo Excel y devuelve las filas en las que las columnas ABONADO y n° abonado tienen exactamente el mismo valor.',
        '/coincidencia_en_fila',
        'Coincidencia en fila',
        [
            'Búsqueda de las columnas ABONADO y n° abonado, incluso con pequeñas variaciones en el encabezado.',
            'Comparación directa entre ambos campos por fila.',
            'Conservación de todas las columnas originales de las filas coincidentes.',
        ],
        'Se genera un Excel con las filas coincidentes y, si el archivo tiene varias hojas, se conservan separadas en la descarga.',
        [
            'Archivo Excel que contenga las columnas ABONADO y n° abonado.',
        ],
        [
            'Ubica el archivo Excel que deseas validar.',
            'Verifica que el libro contenga las columnas ABONADO y n° abonado.',
            'Carga el archivo en el formulario.',
            'Ejecuta el análisis.',
            'Descarga el Excel resultante con las coincidencias encontradas.',
        ],
        [
            {'id': 'archivo_excel', 'name': 'archivo_excel', 'label': 'Archivo Excel', 'accept': '.xlsx,.xls'},
        ],
    )


@app.route('/descargar_resultado')
def descargar_resultado():
    resultado_excel_path = session.get('resultado_excel_path')
    resultado_excel_name = session.get('resultado_excel_name', 'resultado.xlsx')
    if not resultado_excel_path:
        return render_template('error.html', error="No hay archivo para descargar.")
    if os.path.exists(resultado_excel_path):
        return send_file(
            resultado_excel_path,
            as_attachment=True,
            download_name=resultado_excel_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    session.pop('resultado_excel_path', None)
    return render_template('error.html', error="El archivo generado ya no está disponible. Vuelve a procesarlo.")


if __name__ == '__main__':
    app.run(debug=True)
