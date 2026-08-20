# app-cortes

Aplicacion Flask para auditorias operativas basadas en cruces de archivos Excel y CSV.
El aplicativo recibe reportes de SAEPlus, SmartOLT, Workdrive/Drive y ePayco, procesa
los datos con `pandas` y entrega archivos Excel con hallazgos listos para gestion.

## Documentacion

- [Guia de usuario](docs/guia_usuario.md): uso de cada pantalla, entradas, reglas y salidas.
- [Arquitectura](docs/arquitectura.md): estructura del proyecto, flujo web y contrato de servicios.
- [Referencia tecnica](docs/referencia_tecnica.md): modulos, funciones principales, rutas y scripts auxiliares.
- [Columnas y salidas](docs/columnas_y_salidas.md): matriz rapida de archivos requeridos, columnas y hojas generadas.

## Requisitos

- Python 3.11 recomendado.
- Dependencias definidas en `requirements.txt`.

Instalacion:

```bash
pip install -r requirements.txt
```

## Ejecucion local

```bash
python flask_app.py
```

Por defecto la aplicacion inicia en modo debug y queda disponible en:

```text
http://127.0.0.1:5000/
```

Para produccion se recomienda definir `FLASK_SECRET_KEY` y ejecutar con un servidor WSGI
como `gunicorn` en Linux.

## Pruebas

```bash
python -m unittest discover -s tests -v
```

Las pruebas cubren servicios de negocio, validaciones, lectura segura de Excel, cruces
SAEPlus/SmartOLT, precintos, navegacion, atenuaciones, estadisticos OLT, recurrencias y
coincidencia en fila.

## Variables de entorno

- `FLASK_SECRET_KEY`: clave de sesion de Flask. En desarrollo existe un valor por defecto,
  pero en produccion debe configurarse con un valor seguro.

## Flujo general

1. El usuario abre una pantalla de auditoria.
2. Carga uno o varios archivos `.xlsx`, `.xls`, `.xlsm` o `.csv`, segun la auditoria.
3. Flask valida que existan los archivos y que tengan la extension esperada.
4. La ruta llama un servicio en `services/`.
5. El servicio lee los archivos, valida columnas, cruza datos y genera un `BytesIO` Excel.
6. La ruta guarda el Excel temporalmente en `uploads/` asociado a la sesion.
7. La pantalla de resultado muestra una vista previa y permite descargar el Excel.

## Rutas principales

| Ruta | Funcion |
| --- | --- |
| `/` | Pagina principal. |
| `/reconexiones` y `/procesar` | Auditoria simple de reconexiones. |
| `/cortes` | Auditoria de abonados en corte con servicio activo en red. |
| `/verificar_velocidad` | Compara plan SAEPlus contra velocidad SmartOLT y CATV. |
| `/comparativo_equipos` | Cruza equipos que coinciden y no coinciden entre SAEPlus y SmartOLT. |
| `/comparativo_precintos` | Valida precintos cargados y alertas cercanas con SmartOLT. |
| `/sin_navegar` | Auditoria general de navegacion en cuatro hojas. |
| `/navegacion_estado_servicio` | Subanalisis de activos sin navegar y desactivos con internet. |
| `/navegacion_catv_y_planes` | Subanalisis de CATV y planes con arroba. |
| `/auditoria_reconexiones` | Cruce Workdrive, SAEPlus, ePayco y SmartOLT. |
| `/atenuaciones` | Agenda por senales opticas 1310/1490. |
| `/estadisticos_olt` | Resumen y priorizacion de puertos SmartOLT. |
| `/recurrencias` | Transformacion de ordenes de servicio para Power BI. |
| `/coincidencia_en_fila` | Compara columnas ABONADO y numero de abonado en un Excel. |
| `/descargar_resultado` | Descarga el ultimo Excel generado en la sesion actual. |

Tambien existen rutas historicas que redirigen o reutilizan pantallas actuales:
`/diferentes`, `/coincidencias_saeplus_smartolt`, `/precintos_cercanos`,
`/navegacion_activos_sin_navegar`, `/navegacion_desactivos_con_internet`,
`/navegacion_activos_sin_catv` y `/navegacion_solo_con_arroba_y_catv_activo`.

## Estructura

```text
.
|-- flask_app.py                 # Rutas Flask y orquestacion web
|-- funciones.py                 # Lectura segura de Excel/CSV y utilidades compartidas
|-- services/                    # Logica de negocio por auditoria
|-- templates/                   # Vistas Jinja2
|-- static/                      # CSS, JS y favicon
|-- tests/                       # Pruebas unitarias de servicios
|-- uploads/                     # Archivos temporales de descarga por sesion
|-- BuscarV.py                   # CLI auxiliar para agregar ciudad por prefijo de abonado
|-- tmp_*.py                     # Scripts temporales de diagnostico
`-- requirements.txt             # Dependencias Python
```

## Notas operativas

- Los cruces dependen de encabezados y formatos exportados por las fuentes operativas.
- Los Excel descargados se generan en memoria y se guardan temporalmente por sesion.
- La carpeta `uploads/` contiene artefactos de ejecucion; no debe usarse como fuente de datos.
- `Auxiliar_pruebas.py` es un script peligroso que desinstala paquetes del entorno; no hace
  parte del aplicativo Flask.
