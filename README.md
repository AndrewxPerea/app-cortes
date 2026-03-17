# app-cortes

Aplicacion Flask para auditorias operativas basadas en cruces de archivos Excel y CSV. La app permite cargar reportes de distintas fuentes, filtrarlos con `pandas` y descargar un Excel con los hallazgos.

## Requisitos

- Python 3.11 recomendado
- Dependencias del archivo [`requirements.txt`](c:/Users/SUPPORT-WILSON/Desktop/Andres_Perea/Desarrollo/app-cortes/requirements.txt)

## Instalacion

```bash
pip install -r requirements.txt
```

## Ejecucion local

```bash
python flask_app.py
```

La aplicacion inicia en modo debug y normalmente queda disponible en `http://127.0.0.1:5000/`.

## Ejecutar pruebas

```bash
python -m unittest discover -s tests -v
```

## Variables de entorno

- `FLASK_SECRET_KEY`: clave para la sesion de Flask. En produccion debe definirse con un valor seguro.

## Estructura principal

- [`flask_app.py`](c:/Users/SUPPORT-WILSON/Desktop/Andres_Perea/Desarrollo/app-cortes/flask_app.py): rutas Flask, validacion y generacion de reportes.
- [`funciones.py`](c:/Users/SUPPORT-WILSON/Desktop/Andres_Perea/Desarrollo/app-cortes/funciones.py): utilidades para lectura de archivos y normalizacion.
- [`services/`](c:/Users/SUPPORT-WILSON/Desktop/Andres_Perea/Desarrollo/app-cortes/services): logica de negocio separada por auditoria y helpers compartidos.
- [`templates/`](c:/Users/SUPPORT-WILSON/Desktop/Andres_Perea/Desarrollo/app-cortes/templates): vistas HTML.
- [`static/`](c:/Users/SUPPORT-WILSON/Desktop/Andres_Perea/Desarrollo/app-cortes/static): estilos y JS del frontend.
- [`conver.py`](c:/Users/SUPPORT-WILSON/Desktop/Andres_Perea/Desarrollo/app-cortes/conver.py): script auxiliar de prueba, no hace parte del flujo principal Flask.

## Flujo general

1. El usuario entra a una ruta de auditoria.
2. Sube uno o varios archivos Excel/CSV.
3. La app procesa los archivos con `pandas`, cruza columnas clave y aplica filtros.
4. Se muestra una tabla HTML con resultados.
5. Se habilita la descarga de un archivo `resultado.xlsx`.

El resultado descargable ya no se comparte globalmente entre usuarios. Cada sesion guarda su propio archivo temporal.

## Rutas principales

- `/`: pagina principal.
- `/reconexiones`: formulario de auditoria de reconexiones.
- `/procesar`: procesa el formulario de reconexiones.
- `/cortes`: auditoria de abonados en corte.
- `/verificar_velocidad`: compara plan contratado vs velocidad configurada.
- `/diferentes`: detecta equipos que no coinciden entre SAEPlus y OLT.
- `/atenuaciones`: analiza señales 1310/1490 y genera agenda priorizada por daño.
- `/sin_navegar`: auditoria general de navegacion.
- `/auditoria_reconexiones`: cruza drive, SAEPlus, ePayco y SmartOLT.
- `/descargar_resultado`: descarga el ultimo Excel generado en la sesion actual.

## Archivos y columnas esperadas

La app depende fuertemente de nombres exactos de columnas. Si una fuente cambia encabezados, es probable que el proceso falle.

### Reconexiones

Ruta: `/procesar`

Archivos:
- `abonados`: Excel de abonados SAEPlus
- `cortes`: Excel de cortes o workdrive

Columnas clave esperadas tras normalizacion:
- primera columna convertible a `abonados`
- columnas resultantes como `documento_x`, `nombre_x`, `apellido_x`, `observaciones`, `estatus_y`

Salida:
- casos con `observaciones` vacia y `estatus_y == ACTIVO`

### Cortes

Ruta: `/cortes`

Archivos:
- `abonados`: Excel de corte de abonados
- `cortes`: CSV de SmartOLT
- `asaeplus`: Excel de abonados SAEPlus

Columnas clave esperadas:
- en Excel: `EQUIPO MAC`, `N° Abonado`
- en CSV: `SN`
- en el resultado: `estatus`, `observaciones`, `catv`, `administrative status`, `status`, `ingeniero`

Salida:
- abonados con observacion vacia, no activos en sistema y con algun indicio de servicio activo en red

### Verificar Velocidad

Ruta: `/verificar_velocidad`

Archivos:
- `saeplus`: Excel SAEPlus
- `olt`: CSV SmartOLT

Columnas clave esperadas:
- `EQUIPO MAC`
- `SN`
- `detalle suscripcion`
- `service port download speed`
- `service port upload speed`
- `estatus`

Salida:
- abonados activos cuya velocidad extraida del detalle no coincide con la velocidad configurada en OLT

### Diferentes

Ruta: `/diferentes`

Archivos:
- `saeplus`: Excel SAEPlus
- `olt`: CSV SmartOLT

Columnas clave esperadas:
- `EQUIPO MAC`
- `SN`

Salida:
- hoja `Diferentes`
- hoja `Solo en SAEPLUS`
- hoja `Solo en OLT`

### Atenuaciones

Ruta: `/atenuaciones`

Archivos:
- `olt_csv`: CSV exportado desde OLT

Columnas clave esperadas:
- `OLT`
- `Board`
- `Port`
- `Signal 1310`
- `Signal 1490`
- `ONU external ID`
- `Status`
- `Address`

Salida:
- hojas por OLT y señal evaluada
- `AGENDA_GENERAL`
- agendas separadas por `FIBRA`, `ENERGIA`, `MIXTO` y `SIN_CORTE`

### Sin Navegar

Ruta: `/sin_navegar`

Archivos:
- `abonados`: Excel SAEPlus
- `olt`: CSV SmartOLT

Columnas clave esperadas:
- `EQUIPO MAC`
- `SN`
- `estatus`
- `status`
- `detalle suscripcion`
- `catv`
- `administrative status`
- `board`
- `port`

Salida:
- `Activos sin navegar`
- `Desactivos con internet`
- `Activos sin Catv`
- `Solo con @ y catv activo`

### Auditoria Reconexiones

Ruta: `/auditoria_reconexiones`

Archivos:
- `drive`: Excel de drive o workdrive
- `saeplus`: Excel SAEPlus
- `epayco`: Excel ePayco
- `smartolt`: CSV SmartOLT

Columnas clave esperadas:
- primera columna convertible a `abonados` en `drive`, `saeplus` y `epayco`
- `EQUIPO MAC` en SAEPlus
- `SN` en SmartOLT

Salida:
- `Reconexion sin observaciones`
- `Pagos de epayco`
- `Abonados sin activar`
- `pagos saeplus`

## Notas tecnicas

- La app usa archivos temporales por sesion para la descarga del Excel.
- El procesamiento se hace en memoria con `BytesIO` y `pandas`.
- La validacion de columnas es basica y ocurre antes de algunos cruces criticos.
- `flask_app.py` ahora funciona como capa web ligera y delega el procesamiento a modulos en `services/`.
- `funciones.py` todavia contiene utilidades no conectadas al flujo principal, como `procesar_excel` y `clasificar_estado_potencia`.

## Riesgos actuales

- No hay pruebas automatizadas.
- La logica depende de encabezados exactos y formatos muy especificos de archivos de entrada.
- El modo debug esta activo en la ejecucion directa.

## Siguiente mejora recomendada

Agregar pruebas automatizadas para los modulos en [`services/`](c:/Users/SUPPORT-WILSON/Desktop/Andres_Perea/Desarrollo/app-cortes/services) y para los casos borde de lectura y validacion de archivos.
