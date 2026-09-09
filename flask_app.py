import os
import shutil
import tempfile
import unicodedata
from datetime import datetime

import pandas as pd
from flask import Flask, redirect, render_template, request, send_file, session

from services.atenuaciones import procesar_atenuaciones
from services.auditoria_reconexiones import procesar_auditoria_reconexiones
from services.comparativo_equipos import procesar_comparativo_equipos
from services.comparativo_precintos import procesar_comparativo_precintos
from services.coincidencia_en_fila import procesar_coincidencia_en_fila
from services.cortes import procesar_cortes
from services.jobs import procesar_comparativo_estadisticos_olt, procesar_estadisticos_olt
from services.navegacion import (
    procesar_navegacion_catv_y_planes,
    procesar_navegacion_estado_servicio,
    procesar_sin_navegar,
)
from services.recurrencias import PROCESSING_YEAR, procesar_recurrencias
from services.reconexiones import procesar_reconexiones
from services.upload_validation import validar_archivo_opcional, validar_archivos_requeridos
from services.velocidad import procesar_verificacion_velocidad

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'dev-secret-key-change-me')
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def limpiar_resultado_sesion():
    ruta_anterior = session.pop('resultado_excel_path', None)
    session.pop('resultado_excel_name', None)
    session.pop('resultado_excel_id', None)
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
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx', dir=app.config['UPLOAD_FOLDER']) as temp_file:
        shutil.copyfileobj(output, temp_file)
        session['resultado_excel_path'] = temp_file.name
        session['resultado_excel_id'] = os.path.basename(temp_file.name)
        session['resultado_excel_name'] = nombre_descarga or construir_nombre_descarga('resultado')


def renderizar_resultado(resultado, columns=None):
    data = resultado['data'].copy()
    total_registros = len(data)
    limite_preview = 500
    data = data.head(limite_preview)
    data = data.where(pd.notna(data), '')
    return render_template(
        'resultado.html',
        data=data.to_dict(orient='records'),
        columns=columns or resultado.get('columns', data.columns),
        num_casos=resultado['num_casos'],
        nombre_descarga=session.get('resultado_excel_name', 'resultado.xlsx'),
        file_id=resultado.get('file_id') or session.get('resultado_excel_id'),
        summary=resultado.get('summary'),
        total_registros=total_registros,
        registros_mostrados=len(data),
        limite_preview=limite_preview,
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
    campos_input=None,
    campos_texto=None,
    volver_url='/',
    volver_texto='Volver al inicio',
    enlaces_relacionados=None,
    vista_compacta=False,
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
        campos_input=campos_input or [],
        campos_texto=campos_texto or [],
        volver_url=volver_url,
        volver_texto=volver_texto,
        enlaces_relacionados=enlaces_relacionados or [],
        vista_compacta=vista_compacta,
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


@app.route('/comparativo_equipos', methods=['GET', 'POST'])
def comparativo_equipos():
    if request.method == 'POST':
        try:
            archivos = validar_archivos_requeridos(
                request.files,
                [
                    ('saeplus', {'.xlsx', '.xls'}, 'SAEPlus'),
                    ('olt', {'.csv'}, 'de SmartOLT'),
                ]
            )
            resultado = procesar_comparativo_equipos(
                archivos['saeplus'],
                archivos['olt']
            )
        except Exception as e:
            return render_template('error.html', error=str(e))

        return guardar_y_renderizar_resultado(resultado, 'Comparativo de equipos')

    return renderizar_formulario_analisis(
        'Comparativo de Equipos',
        'Cruza SAEPlus y SmartOLT en una sola ejecución para revisar tanto los equipos que coinciden como los que no coinciden entre ambas fuentes.',
        '/comparativo_equipos',
        'Comparativo de equipos',
        [
            'Primera hoja con los equipos que sí coinciden entre SAEPlus y SmartOLT.',
            'Segunda hoja con los equipos que no coinciden entre ambas fuentes.',
            'Cruce por los identificadores EQUIPO MACO y NSN.',
        ],
        'Se genera un Excel con dos hojas: Equipos que coinciden y Equipos que no coinciden.',
        [
            'Archivo de abonados exportado desde SAEPlus en formato Excel.',
            'Archivo de abonados exportado desde SmartOLT en formato CSV.',
        ],
        [
            'Exporta el archivo de abonados desde SAEPlus.',
            'Exporta el archivo de abonados desde SmartOLT en formato CSV.',
            'Abre el Excel de SAEPlus y guárdalo nuevamente antes de cargarlo.',
            'Carga ambos archivos en el formulario.',
            'Ejecuta el análisis y revisa ambas hojas del Excel para validar coincidencias y diferencias.',
        ],
        [
            {'id': 'saeplus', 'name': 'saeplus', 'label': 'Archivo de Abonados (Excel)', 'accept': '.xlsx,.xls'},
            {'id': 'olt', 'name': 'olt', 'label': 'Archivo de Abonados en SmartOLT (CSV)', 'accept': '.csv'},
        ],
    )


@app.route('/diferentes', methods=['GET', 'POST'])
def diferentes():
    if request.method == 'POST':
        return comparativo_equipos()
    return redirect('/comparativo_equipos')


@app.route('/coincidencias_saeplus_smartolt', methods=['GET', 'POST'])
def coincidencias_saeplus_smartolt():
    if request.method == 'POST':
        return comparativo_equipos()
    return redirect('/comparativo_equipos')


@app.route('/comparativo_precintos', methods=['GET', 'POST'])
def comparativo_precintos():
    if request.method == 'POST':
        try:
            archivos = validar_archivos_requeridos(
                request.files,
                [
                    ('saeplus', {'.xlsx', '.xls'}, 'SAEPlus'),
                ]
            )
            archivo_olt = validar_archivo_opcional(
                request.files,
                'olt',
                {'.csv'},
                'de SmartOLT'
            )
            texto_precintos = request.form.get('precintos_texto', '')
            if archivo_olt is None and not str(texto_precintos or '').strip():
                raise ValueError(
                    "Debes cargar el archivo de SmartOLT o escribir al menos un precinto para validar."
                )
            resultado = procesar_comparativo_precintos(
                archivos['saeplus'],
                archivo_olt,
                texto_precintos
            )
        except Exception as e:
            return render_template('error.html', error=str(e))

        return guardar_y_renderizar_resultado(resultado, 'Comparativo de precintos')

    return renderizar_formulario_analisis(
        'Comparativo de Precintos',
        'Ayuda a decidir si un corte puede corresponder a un precinto no actualizado en SAEPlus. Primero valida los precintos enviados por los técnicos y luego lista los abonados sin precinto que además aparecen con contexto técnico en SmartOLT, conservando el estatus real que tengan en SAEPlus.',
        '/comparativo_precintos',
        'Comparativo de precintos',
        [
            'Compara los precintos escritos en la web contra la columna precinto de SAEPlus, sin filtrar por estatus del abonado en esa hoja inicial.',
            'Usa todos los estatus que vengan en SAEPlus en Precintos cargados y conserva Todos los estados SmartOLT cuando se carga ese archivo.',
            'Si cargas SmartOLT y una lista de precintos, agrega una hoja nueva con alertas Offline, LOS y Power fail cercanas por dirección o barrio, priorizando Power fail.',
            'En Alertas cercanas a precintos se excluyen los estados POR SUSPENDER y SUSPENDIDO del lado del abonado alertado para dejar solo casos operativos.',
            'Cada abonado alertado aparece una sola vez en esa hoja, aunque coincida con varios precintos o varias direcciones cercanas.',
            'Si no cargas SmartOLT pero sí escribes precintos, igual valida los precintos contra SAEPlus.',
        ],
        'Se genera un Excel listo para descargar desde la página de resultado. Sin SmartOLT devuelve Precintos cargados; con SmartOLT agrega Alertas cercanas a precintos y Todos los estados SmartOLT.',
        [
            'Archivo de abonados exportado desde SAEPlus en formato Excel.',
            'Archivo de abonados exportado desde SmartOLT en formato CSV. Este archivo es opcional si solo vas a validar precintos contra SAEPlus.',
            'Campo opcional para escribir o pegar números de precinto directamente desde la página.',
        ],
        [
            'Exporta el archivo de abonados desde SAEPlus incluyendo la columna precinto.',
            'Exporta el archivo de abonados desde SmartOLT en formato CSV solo si también quieres el contexto técnico.',
            'Abre el Excel de SAEPlus y guárdalo nuevamente antes de cargarlo.',
            'Pega los precintos del técnico si quieres validar cuáles ya están registrados, incluso cuando no cargues SmartOLT.',
            'Ejecuta el análisis, revisa la página de resultado y luego descarga el Excel.',
        ],
        [
            {'id': 'saeplus', 'name': 'saeplus', 'label': 'Archivo de Abonados SAEPlus (Excel)', 'accept': '.xlsx,.xls'},
            {'id': 'olt', 'name': 'olt', 'label': 'Archivo de Abonados SmartOLT (CSV)', 'accept': '.csv', 'required': False, 'help': 'Opcional si solo vas a validar los precintos contra SAEPlus.'},
        ],
        campos_texto=[
            {
                'id': 'precintos_texto',
                'name': 'precintos_texto',
                'label': 'Precintos para comparar (Opcional)',
                'rows': 8,
                'placeholder': 'Ejemplo:\nPREC-10\nPREC-22\nPREC-99',
                'help': 'Puedes pegar uno por línea o separados por comas, punto y coma, espacios o tabulaciones.',
            },
        ],
        vista_compacta=True,
    )


@app.route('/precintos_cercanos')
def redirigir_precintos_cercanos():
    return redirect('/comparativo_precintos')


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
            {'href': '/navegacion_estado_servicio', 'label': 'Estado de navegación y servicio'},
            {'href': '/navegacion_catv_y_planes', 'label': 'Validación CATV y planes'},
        ],
    )


@app.route('/navegacion_estado_servicio', methods=['GET', 'POST'])
def navegacion_estado_servicio():
    if request.method == 'POST':
        try:
            resultado = ejecutar_procesamiento_navegacion(procesar_navegacion_estado_servicio)
        except Exception as e:
            return render_template('error.html', error=f"Error en el procesamiento: {e}")

        guardar_resultado_excel(
            resultado['excel'],
            construir_nombre_descarga('Estado de navegación y servicio')
        )
        return renderizar_resultado(resultado)

    return renderizar_formulario_analisis(
        'Estado de navegación y servicio',
        'Fusiona en un solo Excel los abonados activos que no están navegando correctamente y los abonados desactivos que siguen con internet en SmartOLT.',
        '/navegacion_estado_servicio',
        'Estado de navegación y servicio',
        [
            'Primera hoja: Activos sin navegar.',
            'Segunda hoja: Desactivos con internet.',
            'Cruce entre SAEPlus y SmartOLT con enfoque operativo en estado comercial vs estado técnico.',
        ],
        'Se genera un Excel con dos hojas: Activos sin navegar y Desactivos con internet.',
        [
            'Archivo de abonados exportado desde SAEPlus en formato Excel.',
            'Archivo de abonados exportado desde SmartOLT en formato CSV.',
        ],
        [
            'Exporta el archivo de abonados desde SAEPlus para la franquicia que deseas revisar.',
            'Abre el Excel de SAEPlus y guárdalo nuevamente antes de cargarlo en la herramienta.',
            'Exporta el archivo de abonados desde SmartOLT en formato CSV.',
            'Carga ambos archivos y ejecuta el análisis.',
            'Revisa las dos hojas del Excel para validar abonados activos sin navegación y desactivos con internet.',
        ],
        [
            {'id': 'abonados', 'name': 'abonados', 'label': 'Archivo de Abonados (Excel)', 'accept': '.xlsx,.xls'},
            {'id': 'olt', 'name': 'olt', 'label': 'Archivo de Abonados en SmartOLT (CSV)', 'accept': '.csv'},
        ],
        volver_url='/sin_navegar',
        volver_texto='Volver a la auditoría general'
    )


@app.route('/navegacion_activos_sin_navegar', methods=['GET', 'POST'])
def navegacion_activos_sin_navegar():
    if request.method == 'POST':
        return navegacion_estado_servicio()
    return redirect('/navegacion_estado_servicio')


@app.route('/navegacion_desactivos_con_internet', methods=['GET', 'POST'])
def navegacion_desactivos_con_internet():
    if request.method == 'POST':
        return navegacion_estado_servicio()
    return redirect('/navegacion_estado_servicio')


@app.route('/navegacion_catv_y_planes', methods=['GET', 'POST'])
def navegacion_catv_y_planes():
    if request.method == 'POST':
        try:
            resultado = ejecutar_procesamiento_navegacion(procesar_navegacion_catv_y_planes)
        except Exception as e:
            return render_template('error.html', error=f"Error en el procesamiento: {e}")

        guardar_resultado_excel(
            resultado['excel'],
            construir_nombre_descarga('Validación CATV y planes')
        )
        return renderizar_resultado(resultado)

    return renderizar_formulario_analisis(
        'Validación CATV y planes',
        'Fusiona en un solo Excel los abonados activos sin CATV y los casos de planes con @ que mantienen CATV activo.',
        '/navegacion_catv_y_planes',
        'Validación CATV y planes',
        [
            'Primera hoja: Activos sin CATV.',
            'Segunda hoja: Solo con @ y CATV activo.',
            'Cruce entre el plan comercial y la configuración de CATV en SmartOLT.',
        ],
        'Se genera un Excel con dos hojas: Activos sin CATV y Solo con @ y CATV activo.',
        [
            'Archivo de abonados exportado desde SAEPlus en formato Excel.',
            'Archivo de abonados exportado desde SmartOLT en formato CSV.',
        ],
        [
            'Exporta el archivo de abonados desde SAEPlus para la franquicia que deseas revisar.',
            'Abre el Excel de SAEPlus y guárdalo nuevamente antes de cargarlo en la herramienta.',
            'Exporta el archivo de abonados desde SmartOLT en formato CSV.',
            'Carga ambos archivos y ejecuta el análisis.',
            'Revisa las dos hojas del Excel para validar casos de CATV faltante o CATV activo en planes con @.',
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
        return navegacion_catv_y_planes()
    return redirect('/navegacion_catv_y_planes')


@app.route('/navegacion_solo_con_arroba_y_catv_activo', methods=['GET', 'POST'])
def navegacion_solo_con_arroba_y_catv_activo():
    if request.method == 'POST':
        return navegacion_catv_y_planes()
    return redirect('/navegacion_catv_y_planes')


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


@app.route('/estadisticos_olt', methods=['GET', 'POST'])
def estadisticos_olt():
    if request.method == 'POST':
        try:
            archivos = validar_archivos_requeridos(
                request.files,
                [
                    ('olt', {'.csv'}, 'de SmartOLT'),
                ]
            )
            resultado = procesar_estadisticos_olt(archivos['olt'])
        except Exception as e:
            return render_template('error.html', error=f"Error en el procesamiento: {e}")

        return guardar_y_renderizar_resultado(resultado, 'Estadisticos OLT')

    return render_template(
        'estadisticos_olt.html',
        nombre_descarga=construir_nombre_descarga('Estadisticos OLT'),
    )


@app.route('/comparativo_estadisticos_olt', methods=['GET', 'POST'])
def comparativo_estadisticos_olt():
    if request.method == 'POST':
        try:
            archivos = validar_archivos_requeridos(
                request.files,
                [
                    ('olt_antiguo', {'.csv'}, 'antiguo de SmartOLT'),
                    ('olt_nuevo', {'.csv'}, 'nuevo de SmartOLT'),
                ]
            )
            resultado = procesar_comparativo_estadisticos_olt(
                archivos['olt_antiguo'],
                archivos['olt_nuevo']
            )
        except Exception as e:
            return render_template('error.html', error=f"Error en el procesamiento: {e}")

        return guardar_y_renderizar_resultado(resultado, 'Comparativo Estadisticos OLT')

    return render_template(
        'comparativo_estadisticos_olt.html',
        nombre_descarga=construir_nombre_descarga('Comparativo Estadisticos OLT'),
    )


@app.route('/recurrencias', methods=['GET', 'POST'])
def recurrencias():
    if request.method == 'POST':
        try:
            archivos = validar_archivos_requeridos(
                request.files,
                [
                    ('ordenes', {'.xlsx', '.xlsm'}, 'Excel de ordenes de servicio'),
                ]
            )
            year_raw = str(request.form.get('year') or PROCESSING_YEAR).strip()
            year = int(year_raw)
            if year < 1900 or year > 2100:
                raise ValueError("El año debe estar entre 1900 y 2100.")

            resultado = procesar_recurrencias(archivos['ordenes'], year=year)
        except Exception as e:
            return render_template('error.html', error=f"Error en el procesamiento: {e}")

        return guardar_y_renderizar_resultado(resultado, 'Recurrencias ordenes de servicio')

    return renderizar_formulario_analisis(
        'Recurrencias de Servicio',
        'Transforma un Excel de ordenes de servicio, calcula visitas cronologicas por abonado y genera un archivo enriquecido listo para Power BI.',
        '/recurrencias',
        'Recurrencias ordenes de servicio',
        [
            'Valida que existan todas las columnas obligatorias del archivo de ordenes.',
            'Limpia encabezados, espacios y valores de texto antes de transformar.',
            'Construye Fecha con DÍA, MES y el año indicado en el formulario.',
            'Ordena cada ABONADO por fecha para calcular visita, recurrencia y cambios de ingeniero.',
            'Conserva todas las columnas originales y agrega las columnas calculadas al final.',
        ],
        'Se genera un Excel con todas las ordenes transformadas y enriquecidas para cargar directamente en Power BI.',
        [
            'Archivo Excel de ordenes de servicio en formato .xlsx o .xlsm.',
        ],
        [
            'Exporta el archivo de ordenes de servicio desde la fuente operativa.',
            'Abre el Excel y guardalo nuevamente si viene de una descarga automatica.',
            'Carga el archivo en esta pagina.',
            'Confirma el año que se usara para construir la columna Fecha.',
            'Ejecuta el proceso y descarga el Excel listo para Power BI.',
        ],
        [
            {'id': 'ordenes', 'name': 'ordenes', 'label': 'Archivo de Ordenes de Servicio', 'accept': '.xlsx,.xlsm'},
        ],
        campos_input=[
            {
                'id': 'year',
                'name': 'year',
                'label': 'Año de las visitas',
                'type': 'number',
                'value': PROCESSING_YEAR,
                'min': 1900,
                'max': 2100,
                'required': True,
                'help': 'Este año se combina con las columnas DÍA y MES para construir Fecha.',
            },
        ],
        vista_compacta=True,
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
        'Revisa un archivo Excel y compara toda la columna ABONADO contra toda la columna n° abonado para separar los registros que sí coinciden y los que no coinciden, conservando las columnas originales del archivo.',
        '/coincidencia_en_fila',
        'Coincidencia en fila',
        [
            'Búsqueda de las columnas ABONADO y n° abonado, incluso con pequeñas variaciones en el encabezado.',
            'Comparación global entre ambas columnas, no solo por la misma fila.',
            'Hoja Coinciden basada en n° abonado con las columnas originales del archivo y el detalle de la coincidencia encontrada en ABONADO.',
            'Hoja No coinciden basada en n° abonado para revisar los registros que no aparecen en la columna ABONADO.',
        ],
        'Se genera un Excel con dos hojas: Coinciden y No coinciden.',
        [
            'Archivo Excel que contenga las columnas ABONADO y n° abonado.',
        ],
        [
            'Ubica el archivo Excel que deseas validar.',
            'Verifica que el libro contenga las columnas ABONADO y n° abonado.',
            'Carga el archivo en el formulario.',
            'Ejecuta el análisis.',
            'Descarga el Excel resultante para revisar primero Coinciden y luego No coinciden, siempre tomando como base la columna n° abonado.',
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
