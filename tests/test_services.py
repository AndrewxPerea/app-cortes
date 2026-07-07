import io
import zipfile
import unittest
from xml.etree import ElementTree as ET

import pandas as pd

from funciones import abrir_excel_seguro, leer_excel_seguro
from services.common import excel_desde_dataframe, excel_desde_hojas, validar_columnas
from services.atenuaciones import procesar_atenuaciones
from services.auditoria_reconexiones import procesar_auditoria_reconexiones
from services.ciudades_abonado import agregar_columna_ciudad
from services.comparativo_equipos import procesar_comparativo_equipos
from services.comparativo_precintos import procesar_comparativo_precintos
from services.coincidencia_en_fila import procesar_coincidencia_en_fila
from services.coincidencias_saeplus_smartolt import procesar_coincidencias_saeplus_smartolt
from services.cortes import procesar_cortes
from services.diferentes import procesar_diferentes
from services.jobs import job_estadisticos_olt, procesar_estadisticos_olt
from services.navegacion import (
    procesar_navegacion_catv_y_planes,
    procesar_navegacion_estado_servicio,
    procesar_navegacion_activos_sin_catv,
    procesar_navegacion_activos_sin_navegar,
    procesar_navegacion_desactivos_con_internet,
    procesar_navegacion_solo_con_arroba_y_catv_activo,
    procesar_sin_navegar,
)
from services.precintos_cercanos import (
    cargar_saeplus_con_ubicacion,
    procesar_precintos_cercanos,
    referencia_direccion,
)
from services.reconexiones import procesar_reconexiones
from services.upload_validation import validar_archivos_requeridos
from services.velocidad import (
    catv_esta_activo,
    es_plan_solo_internet,
    extraer_velocidad,
    normalizar_velocidad_red,
    procesar_verificacion_velocidad,
)


def excel_buffer(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer


def csv_buffer(df):
    buffer = io.StringIO()
    df.to_csv(buffer, index=False)
    buffer.seek(0)
    return buffer


def excel_buffer_con_fill_invalido(df):
    buffer = excel_buffer(df)
    entrada = io.BytesIO(buffer.getvalue())
    salida = io.BytesIO()

    with zipfile.ZipFile(entrada, 'r') as origen, zipfile.ZipFile(salida, 'w', zipfile.ZIP_DEFLATED) as destino:
        for nombre in origen.namelist():
            contenido = origen.read(nombre)
            if nombre == 'xl/styles.xml':
                root = ET.fromstring(contenido)
                ns = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                fills = root.find('a:fills', ns)
                if fills is not None and len(list(fills)) > 0:
                    fills.remove(list(fills)[0])
                contenido = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            destino.writestr(nombre, contenido)

    salida.seek(0)
    return salida


def navegacion_archivos():
    abonados = pd.DataFrame({
        'EQUIPO MAC': [
            'AA:BB:CC:11:22:12345678',
            'AA:BB:CC:11:22:87654321',
            'AA:BB:CC:11:22:ABCDEF12',
            'AA:BB:CC:11:22:11223344',
        ],
        'N° Abonado': [5001, 5002, 5003, 5004],
        'documento': ['90', '91', '92', '93'],
        'nombre': ['Luisa', 'Mario', 'Sara', 'Ana'],
        'estatus': ['ACTIVO', 'SUSPENDIDO', 'ACTIVO', 'ACTIVO'],
        'detalle suscripcion': ['GPON 50 MG', 'GPON 30 MG', 'GPON 70 MG', 'PLAN @ 100 MG'],
        'nombre franquicia': ['Norte', 'Centro', 'Sur', 'Centro'],
        'tipo tecnología.': ['GPON', 'GPON', 'GPON', 'GPON'],
        'catv abonado': ['Enabled', 'Enabled', 'Disabled', 'Enabled'],
        'administrative status abonado': ['Enabled', 'Enabled', 'Enabled', 'Enabled'],
    })
    olt = pd.DataFrame({
        'SN': ['SN12345678', 'SN87654321', 'SNABCDEF12', 'SN11223344'],
        'name': ['ONT Luisa', 'ONT Mario', 'ONT Sara', 'ONT Ana'],
        'status': ['Offline', 'Online', 'Online', 'Online'],
        'olt': ['OLT-3', 'OLT-4', 'OLT-5', 'OLT-6'],
        'board': ['1', '2', '3', '4'],
        'port': ['10', '11', '12', '13'],
        'service port upload speed': ['10MG', '10MG', '20MG', '20MG'],
        'service port download speed': ['50MG', '30MG', '70MG', '100MG'],
        'catv': ['Enabled', 'Enabled', 'Disabled', 'Enabled'],
        'administrative status': ['Disabled', 'Enabled', 'Enabled', 'Enabled'],
    })
    return excel_buffer(abonados), csv_buffer(olt)


class CommonServicesTests(unittest.TestCase):
    def test_validar_columnas_no_falla_si_todas_existen(self):
        df = pd.DataFrame({'abonados': [1], 'nombre': ['Ana']})
        validar_columnas(df, ['abonados', 'nombre'], 'prueba')

    def test_validar_columnas_lanza_error_si_faltan_columnas(self):
        df = pd.DataFrame({'abonados': [1]})

        with self.assertRaises(ValueError) as ctx:
            validar_columnas(df, ['abonados', 'nombre'], 'prueba')

        self.assertIn('nombre', str(ctx.exception))

    def test_excel_desde_dataframe_genera_un_excel_legible(self):
        original = pd.DataFrame({'abonados': [1, 2], 'estatus': ['ACTIVO', 'SUSPENDIDO']})

        output = excel_desde_dataframe(original, 'Resultado')
        leido = pd.read_excel(output)

        self.assertEqual(leido.to_dict(orient='records'), original.to_dict(orient='records'))

    def test_excel_desde_hojas_genera_varias_hojas(self):
        output = excel_desde_hojas([
            ('Hoja1', pd.DataFrame({'a': [1]})),
            ('Hoja2', pd.DataFrame({'b': [2]})),
        ])

        excel = pd.ExcelFile(output)
        self.assertEqual(excel.sheet_names, ['Hoja1', 'Hoja2'])

    def test_leer_excel_seguro_repara_fill_invalido(self):
        archivo = excel_buffer_con_fill_invalido(pd.DataFrame({'EQUIPO MAC': ['AA:BB:CC:11:22:12345678']}))

        df = leer_excel_seguro(archivo)

        self.assertEqual(df['EQUIPO MAC'].tolist(), ['AA:BB:CC:11:22:12345678'])

    def test_abrir_excel_seguro_repara_fill_invalido(self):
        archivo = excel_buffer_con_fill_invalido(pd.DataFrame({'ABONADO': ['CH100']}))

        excel = abrir_excel_seguro(archivo)

        self.assertEqual(excel.sheet_names, ['Sheet1'])

    def test_validar_archivos_requeridos_acepta_archivo_valido(self):
        class Archivo:
            filename = 'reporte.csv'

        archivos = validar_archivos_requeridos(
            {'olt_csv': Archivo()},
            [('olt_csv', {'.csv'}, 'de OLT')]
        )

        self.assertIn('olt_csv', archivos)

    def test_validar_archivos_requeridos_rechaza_extension_invalida(self):
        class Archivo:
            filename = 'reporte.xlsx'

        with self.assertRaises(ValueError) as ctx:
            validar_archivos_requeridos(
                {'olt_csv': Archivo()},
                [('olt_csv', {'.csv'}, 'de OLT')]
            )

        self.assertIn('.csv', str(ctx.exception))

    def test_referencia_direccion_detecta_cruces_similares(self):
        self.assertEqual(referencia_direccion('CR 16A #27B-48'), '16 27')
        self.assertEqual(referencia_direccion('CL 27 # 16A'), '16 27')

    def test_cargar_saeplus_con_ubicacion_consolida_duplicados_con_primer_valor_util(self):
        saeplus_base = pd.DataFrame({
            'EQUIPO MAC': ['AA:BB:CC:11:22:12345678'],
            'N° Abonado': [9001],
            'documento': ['10'],
            'nombre': ['Ana'],
            'estatus': ['ACTIVO'],
            'precinto': [''],
        })
        saeplus_ubicacion = pd.DataFrame({
            'N° Abonado': [9001, 9001],
            'Barrio': ['', 'Centro'],
            'Dirección': ['', 'Cra 1 # 10-20'],
            'Ciudad': ['', 'Pereira'],
        })

        saeplus_excel = excel_desde_hojas([
            ('Base', saeplus_base),
            ('Ubicaciones', saeplus_ubicacion),
        ])

        resultado = cargar_saeplus_con_ubicacion(saeplus_excel, requerir_ubicacion=True)

        self.assertEqual(resultado['barrio'].tolist(), ['Centro'])
        self.assertEqual(resultado['direccion'].tolist(), ['Cra 1 # 10-20'])
        self.assertEqual(resultado['ciudad'].tolist(), ['Pereira'])


class ReconexionesServiceTests(unittest.TestCase):
    def test_procesar_reconexiones_filtra_solo_activos_sin_observacion(self):
        cortes = pd.DataFrame({
            'Cuenta': [1001, 1002, 1003],
            'documento': ['1', '2', '3'],
            'nombre': ['Ana', 'Luis', 'Marta'],
            'apellido': ['A', 'B', 'C'],
            'estatus': ['EN REVISION', 'EN REVISION', 'EN REVISION'],
            'observaciones': [None, 'ya revisado', ''],
        })
        abonados = pd.DataFrame({
            'Cuenta': [1001, 1002, 1003],
            'documento': ['1', '2', '3'],
            'nombre': ['Ana', 'Luis', 'Marta'],
            'apellido': ['A', 'B', 'C'],
            'estatus': ['ACTIVO', 'ACTIVO', 'SUSPENDIDO'],
        })

        resultado = procesar_reconexiones(excel_buffer(abonados), excel_buffer(cortes))

        self.assertEqual(resultado['num_casos'], 1)
        self.assertEqual(resultado['data']['abonados'].tolist(), [1001])


class VelocidadServiceTests(unittest.TestCase):
    def test_extraer_velocidad_detecta_patron_mg(self):
        self.assertEqual(extraer_velocidad('Plan Hogar 200 MG promo'), '200MG')

    def test_extraer_velocidad_retorna_none_si_no_hay_velocidad(self):
        self.assertIsNone(extraer_velocidad('Plan sin valor numerico'))

    def test_extraer_velocidad_detecta_solo_arroba_sin_mg(self):
        self.assertEqual(extraer_velocidad('SOLO @ 100 $70.000'), '100MG')

    def test_es_plan_solo_internet_detecta_palabra_solo(self):
        self.assertTrue(es_plan_solo_internet('SOLO INTERNET 100 MG'))
        self.assertFalse(es_plan_solo_internet('PLAN HOGAR 100 MG'))

    def test_catv_esta_activo_acepta_enable_y_enabled(self):
        self.assertTrue(catv_esta_activo('Enable'))
        self.assertTrue(catv_esta_activo('Enabled'))
        self.assertFalse(catv_esta_activo('Disabled'))

    def test_normalizar_velocidad_red_detecta_valor_numerico(self):
        self.assertEqual(normalizar_velocidad_red('100MG'), '100MG')
        self.assertEqual(normalizar_velocidad_red('100'), '100MG')

    def test_procesar_verificacion_velocidad_detecta_desajuste(self):
        saeplus = pd.DataFrame({
            'EQUIPO MAC': ['AA:BB:CC:11:22:12345678'],
            'N° Abonado': [2001],
            'documento': ['55'],
            'nombre': ['Carla'],
            'estatus': ['ACTIVO'],
            'detalle suscripcion': ['GPON 100 MG $82.000'],
            'nombre franquicia': ['Centro'],
            'tipo tecnología.': ['GPON'],
        })
        olt = pd.DataFrame({
            'SN': ['SN12345678'],
            'name': ['ONT Carla'],
            'olt': ['OLT-1'],
            'service port upload speed': ['10MG'],
            'service port download speed': ['50MG'],
            'catv': ['Disabled'],
        })

        resultado = procesar_verificacion_velocidad(excel_buffer(saeplus), csv_buffer(olt))

        self.assertEqual(resultado['num_casos'], 1)
        self.assertEqual(resultado['data']['n° abonado'].tolist(), [2001])
        self.assertEqual(resultado['data']['motivo_revision'].tolist(), ['Velocidad no coincide'])

    def test_procesar_verificacion_velocidad_detecta_solo_internet_con_catv_activo(self):
        saeplus = pd.DataFrame({
            'EQUIPO MAC': ['AA:BB:CC:11:22:87654321'],
            'N° Abonado': [2002],
            'documento': ['56'],
            'nombre': ['Diego'],
            'estatus': ['ACTIVO'],
            'detalle suscripcion': ['SOLO INTERNET 100 MG $80.000'],
            'nombre franquicia': ['Centro'],
            'tipo tecnología.': ['GPON'],
        })
        olt = pd.DataFrame({
            'SN': ['SN87654321'],
            'name': ['ONT Diego'],
            'olt': ['OLT-2'],
            'service port upload speed': ['10MG'],
            'service port download speed': ['100MG'],
            'catv': ['Enable'],
        })

        resultado = procesar_verificacion_velocidad(excel_buffer(saeplus), csv_buffer(olt))

        self.assertEqual(resultado['num_casos'], 1)
        self.assertEqual(resultado['data']['n° abonado'].tolist(), [2002])
        self.assertEqual(
            resultado['data']['motivo_revision'].tolist(),
            ['Plan solo internet con CATV activo']
        )

    def test_procesar_verificacion_velocidad_no_marca_solo_internet_si_catv_esta_disabled_y_velocidad_coincide(self):
        saeplus = pd.DataFrame({
            'EQUIPO MAC': ['AA:BB:CC:11:22:6B520246'],
            'N° Abonado': ['C016454'],
            'documento': ['5982345.0'],
            'nombre': ['ESTEBAN'],
            'estatus': ['ACTIVO'],
            'detalle suscripcion': ['SOLO @ 100 $70.000'],
            'nombre franquicia': ['CARTAGO'],
            'tipo tecnología.': ['SMARTOLT_1_CARTAGO'],
        })
        olt = pd.DataFrame({
            'SN': ['XPON6B520246'],
            'name': ['C016454 - 5982345 - ESTEBAN SOGAMOSO GONZALEZ'],
            'olt': ['301_OLT_1_CARTAGO'],
            'service port upload speed': ['100MG'],
            'service port download speed': ['100MG'],
            'catv': ['Disabled'],
        })

        resultado = procesar_verificacion_velocidad(excel_buffer(saeplus), csv_buffer(olt))

        self.assertEqual(resultado['num_casos'], 0)


class CortesServiceTests(unittest.TestCase):
    def test_procesar_cortes_detecta_abonado_con_servicio_activo_en_red(self):
        cortes = pd.DataFrame({
            'EQUIPO MAC': ['AA:BB:CC:11:22:12345678'],
            'N° Abonado': [3001],
            'documento': ['80'],
            'nombre': ['Pedro'],
            'apellido': ['Lopez'],
            'Estatus': ['ACTIVO'],
            'observaciones': [None],
            'Ingeniero': ['Carlos'],
        })
        saeplus = pd.DataFrame({
            'EQUIPO MAC': ['AA:BB:CC:11:22:12345678'],
            'N° Abonado': [3001],
            'documento': ['80'],
            'nombre': ['Pedro'],
            'apellido': ['Lopez'],
            'Estatus': ['SUSPENDIDO'],
        })
        olt = pd.DataFrame({
            'SN': ['SN12345678'],
            'olt': ['OLT-9'],
            'catv': ['Enabled'],
            'administrative status': ['Enabled'],
            'status': ['Online'],
        })

        resultado = procesar_cortes(excel_buffer(cortes), csv_buffer(olt), excel_buffer(saeplus))

        self.assertEqual(resultado['num_casos'], 1)
        self.assertEqual(resultado['data']['estatus'].tolist(), ['SUSPENDIDO'])
        self.assertNotIn('ingeniero', resultado['data'].columns)
        self.assertEqual(resultado['data']['n° abonado'].tolist(), [3001])


class DiferentesServiceTests(unittest.TestCase):
    def test_procesar_diferentes_devuelve_registros_no_coincidentes(self):
        saeplus = pd.DataFrame({
            'EQUIPO MAC': ['AA:BB:CC:11:22:12345678', 'AA:BB:CC:11:22:87654321'],
            'N° Abonado': [4001, 4002],
        })
        olt = pd.DataFrame({
            'SN': ['SN12345678', 'SN99999999'],
            'olt': ['OLT-1', 'OLT-2'],
        })

        resultado = procesar_diferentes(excel_buffer(saeplus), csv_buffer(olt))

        self.assertEqual(resultado['num_casos'], 2)
        self.assertIn('_merge', resultado['data'].columns)

    def test_procesar_coincidencias_saeplus_smartolt_devuelve_solo_registros_que_empatan(self):
        saeplus = pd.DataFrame({
            'EQUIPO MAC': ['AA:BB:CC:11:22:12345678', 'AA:BB:CC:11:22:87654321'],
            'N° Abonado': [4001, 4002],
            'nombre': ['Ana', 'Luis'],
        })
        olt = pd.DataFrame({
            'SN': ['SN12345678', 'SN99999999'],
            'olt': ['OLT-1', 'OLT-2'],
            'status': ['Online', 'Offline'],
        })

        resultado = procesar_coincidencias_saeplus_smartolt(excel_buffer(saeplus), csv_buffer(olt))

        self.assertEqual(resultado['num_casos'], 1)
        self.assertEqual(resultado['data']['N° Abonado'].tolist(), [4001])
        self.assertEqual(resultado['data']['NSN'].tolist(), ['12345678'])

    def test_procesar_comparativo_equipos_retorna_dos_hojas(self):
        saeplus = pd.DataFrame({
            'EQUIPO MAC': ['AA:BB:CC:11:22:12345678', 'AA:BB:CC:11:22:87654321'],
            'N° Abonado': [4001, 4002],
            'nombre': ['Ana', 'Luis'],
        })
        olt = pd.DataFrame({
            'SN': ['SN12345678', 'SN99999999'],
            'olt': ['OLT-1', 'OLT-2'],
            'status': ['Online', 'Offline'],
        })

        resultado = procesar_comparativo_equipos(excel_buffer(saeplus), csv_buffer(olt))

        self.assertEqual(resultado['num_casos'], 3)
        self.assertEqual(resultado['data']['N° Abonado'].tolist(), [4001])

        excel = pd.ExcelFile(resultado['excel'])
        self.assertEqual(excel.sheet_names, ['Equipos que coinciden', 'Equipos que no coinciden'])
        self.assertEqual(excel.parse('Equipos que coinciden')['N° Abonado'].tolist(), [4001])
        hoja_no_coinciden = excel.parse('Equipos que no coinciden')
        self.assertEqual(hoja_no_coinciden.shape[0], 2)
        self.assertIn('no coincide en', hoja_no_coinciden.columns)
        self.assertEqual(
            sorted(hoja_no_coinciden['no coincide en'].tolist()),
            ['No coincide en SAEPlus', 'No coincide en SmartOLT']
        )

    def test_procesar_comparativo_equipos_ordena_n_abonado_mixto_sin_error(self):
        saeplus = pd.DataFrame({
            'EQUIPO MAC': [
                'AA:BB:CC:11:22:12345678',
                'AA:BB:CC:11:22:87654321',
                'AA:BB:CC:11:22:11111111',
            ],
            'N° Abonado': [4001, 'C016454', 'SR006144'],
            'nombre': ['Ana', 'Luis', 'Marta'],
        })
        olt = pd.DataFrame({
            'SN': ['SN12345678', 'SN87654321', 'SN11111111', 'SN22222222'],
            'olt': ['OLT-1', 'OLT-2', 'OLT-3', 'OLT-4'],
            'status': ['Online', 'Offline', 'Online', 'Offline'],
        })

        resultado = procesar_comparativo_equipos(excel_buffer(saeplus), csv_buffer(olt))

        self.assertEqual(resultado['data']['N° Abonado'].tolist(), [4001, 'C016454', 'SR006144'])

        excel = pd.ExcelFile(resultado['excel'])
        self.assertEqual(
            excel.parse('Equipos que coinciden')['N° Abonado'].tolist(),
            [4001, 'C016454', 'SR006144']
        )

    def test_procesar_comparativo_precintos_devuelve_solo_candidatos_operativos(self):
        saeplus = pd.DataFrame({
            'EQUIPO MAC': [
                'AA:BB:CC:11:22:12345678',
                'AA:BB:CC:11:22:87654321',
                'AA:BB:CC:11:22:11111111',
                'AA:BB:CC:11:22:22222222',
                'AA:BB:CC:11:22:33333333',
                'AA:BB:CC:11:22:44444444',
            ],
            'N° Abonado': [4101, 4102, 4103, 4104, 4105, 4106],
            'documento': ['10', '20', '30', '40', '50', '60'],
            'nombre': ['Ana', 'Luis', 'Carla', 'Pedro', 'Marta', 'Sofia'],
            'estatus': ['ACTIVO', 'ACTIVO', 'ACTIVO', 'SUSPENDIDO', 'ACTIVO', 'CORTADO'],
            'Barrio': ['Centro', 'La Floresta', 'Alamos', 'Galan', 'Bosques', 'Portal'],
            'Dirección': [
                'Cra 1 # 10-20',
                'Calle 8 # 15-30',
                'Mz 4 Casa 9',
                'Cra 7 # 2-10',
                'Calle 100 # 50-20',
                'Cra 9 # 12-34',
            ],
            'Ciudad': ['Cartago', 'Pereira', 'Dosquebradas', 'Santa Rosa', 'Pereira', 'Cartago'],
            'precinto': ['', '', '', '', 'PREC-40', ''],
        })
        olt = pd.DataFrame({
            'SN': ['SN12345678', 'SN87654321', 'SN11111111', 'SN22222222', 'SN99999999', 'SN44444444'],
            'name': ['ONT Ana', 'ONT Luis', 'ONT Carla', 'ONT Pedro', 'ONT Extra', 'ONT Sofia'],
            'status': ['LOS', 'Power fail', 'Offline', 'LOS', 'Offline', 'Offline'],
            'olt': ['OLT-1', 'OLT-1', 'OLT-2', 'OLT-3', 'OLT-4', 'OLT-5'],
        })

        resultado = procesar_comparativo_precintos(excel_buffer(saeplus), csv_buffer(olt))

        self.assertEqual(resultado['num_casos'], 5)
        self.assertIn('prioridad', resultado['data'].columns)
        self.assertIn('hallazgo', resultado['data'].columns)
        self.assertIn('barrio', resultado['data'].columns)
        self.assertIn('dirección', resultado['data'].columns)
        self.assertIn('ciudad', resultado['data'].columns)
        self.assertIn('referencia dirección', resultado['data'].columns)
        self.assertNotIn('name', resultado['data'].columns)

        self.assertEqual(
            resultado['data']['n° abonado'].astype(str).tolist(),
            ['4101', '4104', '4102', '4106', '4103']
        )
        self.assertEqual(
            resultado['data']['prioridad'].tolist(),
            ['Alta', 'Alta', 'Media', 'Baja', 'Baja']
        )
        self.assertEqual(
            resultado['data']['status'].tolist(),
            ['LOS', 'LOS', 'Power fail', 'Offline', 'Offline']
        )
        self.assertEqual(resultado['data'].iloc[0]['barrio'], 'Centro')
        self.assertEqual(resultado['data'].iloc[0]['dirección'], 'Cra 1 # 10-20')
        self.assertEqual(resultado['data'].iloc[0]['referencia dirección'], '1 10')
        self.assertNotIn('4105', resultado['data']['n° abonado'].dropna().astype(str).tolist())
        self.assertIn('4104', resultado['data']['n° abonado'].dropna().astype(str).tolist())
        self.assertIn('4106', resultado['data']['n° abonado'].dropna().astype(str).tolist())
        self.assertEqual(
            resultado['data'].loc[resultado['data']['n° abonado'].astype(str) == '4103', 'hallazgo'].tolist(),
            ['Abonado sin precinto en SAEPlus y con OFFLINE en SmartOLT']
        )
        self.assertEqual(
            resultado['data'].loc[resultado['data']['n° abonado'].astype(str) == '4106', 'estatus'].tolist(),
            ['CORTADO']
        )

        excel = pd.ExcelFile(resultado['excel'])
        self.assertEqual(excel.sheet_names, ['Todos los estados SmartOLT'])
        hoja_todos = excel.parse('Todos los estados SmartOLT')
        self.assertEqual(hoja_todos['n° abonado'].astype(str).tolist(), ['4101', '4104', '4102', '4106', '4103'])

    def test_procesar_comparativo_precintos_incluye_cercania_y_precintos_cargados(self):
        saeplus_base = pd.DataFrame({
            'EQUIPO MAC': [
                'AA:BB:CC:11:22:12345678',
                'AA:BB:CC:11:22:87654321',
                'AA:BB:CC:11:22:11111111',
                'AA:BB:CC:11:22:22222222',
                'AA:BB:CC:11:22:33333333',
            ],
            'N° Abonado': [7001, 7002, 7003, 7004, 7005],
            'documento': ['11', '22', '33', '44', '55'],
            'nombre': ['Ana', 'Luis', 'Carla', 'Pedro', 'Marta'],
            'estatus': ['ACTIVO', 'ACTIVO', 'ACTIVO', 'ACTIVO', 'ACTIVO'],
            'precinto': ['', 'PREC-10', '', 'PREC-20', ''],
        })
        saeplus_ubicacion = pd.DataFrame({
            'N° Abonado': [7001, 7002, 7003, 7004, 7005],
            'Barrio': ['Centro', 'Centro', 'Bosques', 'Centro', 'Alamos'],
            'Dirección': [
                'Cra 1 # 10-20',
                'Cra 1 # 10-25',
                'Calle 8 # 15-30',
                'Cra 1 # 11-05',
                'Mz 9 Casa 1',
            ],
            'Ciudad': ['Pereira', 'Pereira', 'Cartago', 'Pereira', 'Cartago'],
        })
        olt = pd.DataFrame({
            'SN': ['SN12345678', 'SN87654321', 'SN11111111', 'SN22222222', 'SN33333333'],
            'name': ['ONT Ana', 'ONT Luis', 'ONT Carla', 'ONT Pedro', 'ONT Marta'],
            'status': ['Offline', 'Offline', 'LOS', 'Power fail', 'Online'],
            'olt': ['OLT-1', 'OLT-1', 'OLT-2', 'OLT-3', 'OLT-4'],
        })
        precintos_texto = 'PREC-10\nPREC-99'

        saeplus_excel = excel_desde_hojas([
            ('Base', saeplus_base),
            ('Ubicaciones', saeplus_ubicacion),
        ])
        resultado = procesar_comparativo_precintos(
            saeplus_excel,
            csv_buffer(olt),
            precintos_texto
        )

        excel = pd.ExcelFile(resultado['excel'])
        self.assertEqual(
            excel.sheet_names,
            ['Precintos cargados', 'Alertas cercanas a precintos', 'Todos los estados SmartOLT']
        )

        hoja_precintos = excel.parse('Precintos cargados', dtype=str).fillna('')
        self.assertEqual(
            hoja_precintos.columns.tolist(),
            ['precinto cargado', 'coincide en saeplus', 'precinto saeplus', 'n° abonado', 'documento', 'nombre', 'estatus', 'ciudad', 'barrio', 'dirección']
        )
        self.assertEqual(hoja_precintos['precinto cargado'].tolist(), ['PREC-10', 'PREC-99'])
        self.assertEqual(hoja_precintos['coincide en saeplus'].tolist(), ['Si', 'No'])
        self.assertEqual(
            hoja_precintos.loc[
                hoja_precintos['precinto cargado'] == 'PREC-10',
                'n° abonado'
            ].astype(str).str.replace('.0', '', regex=False).tolist(),
            ['7002']
        )

        hoja_alertas = excel.parse('Alertas cercanas a precintos', dtype=str).fillna('')
        self.assertEqual(
            hoja_alertas.columns.tolist(),
            [
                'prioridad',
                'criterio ubicación',
                'precinto cargado',
                'precinto saeplus',
                'n° abonado referencia',
                'ciudad referencia',
                'barrio referencia',
                'dirección referencia',
                'n° abonado alerta',
                'documento alerta',
                'nombre alerta',
                'estatus alerta',
                'status smartolt alerta',
                'ciudad alerta',
                'barrio alerta',
                'dirección alerta',
                'precinto alerta saeplus',
                'equipo maco alerta',
                'sn alerta',
                'olt alerta',
            ]
        )
        self.assertEqual(hoja_alertas['prioridad'].tolist(), ['Alta', 'Baja', 'Baja'])
        self.assertEqual(hoja_alertas['status smartolt alerta'].tolist(), ['Power fail', 'Offline', 'Offline'])
        self.assertEqual(hoja_alertas['n° abonado alerta'].tolist(), ['7004', '7001', '7002'])
        self.assertEqual(hoja_alertas['nombre alerta'].tolist(), ['Pedro', 'Ana', 'Luis'])
        self.assertEqual(
            hoja_alertas['criterio ubicación'].tolist(),
            ['Mismo barrio', 'Mismo barrio y dirección aproximada', 'Mismo barrio y dirección aproximada']
        )

        hoja_todos = excel.parse('Todos los estados SmartOLT', dtype=str).fillna('')
        self.assertEqual(hoja_todos['n° abonado'].tolist(), ['7005', '7003', '7001'])
        self.assertEqual(hoja_todos['status'].tolist(), ['Online', 'LOS', 'Offline'])

    def test_procesar_comparativo_precintos_compara_precintos_numericos_pegados_desde_web(self):
        saeplus = pd.DataFrame({
            'EQUIPO MAC': [
                'AA:BB:CC:11:22:12345678',
                'AA:BB:CC:11:22:87654321',
                'AA:BB:CC:11:22:11111111',
            ],
            'N° Abonado': [8001, 8002, 8003],
            'documento': ['10', '20', '30'],
            'nombre': ['Ana', 'Luis', 'Carla'],
            'estatus': ['ACTIVO', 'ACTIVO', 'SUSPENDIDO'],
            'Barrio': ['Centro', 'Centro', 'Bosques'],
            'Dirección': ['Cra 1 # 10-20', 'Cra 2 # 20-30', 'Calle 8 # 15-30'],
            'Ciudad': ['Pereira', 'Pereira', 'Cartago'],
            'precinto': [2211095.0, 13557.0, 2284686.0],
        })
        olt = pd.DataFrame({
            'SN': ['SN12345678', 'SN87654321', 'SN11111111'],
            'name': ['ONT Ana', 'ONT Luis', 'ONT Carla'],
            'status': ['Online', 'Online', 'Offline'],
            'olt': ['OLT-1', 'OLT-2', 'OLT-3'],
        })

        resultado = procesar_comparativo_precintos(
            excel_buffer(saeplus),
            csv_buffer(olt),
            '2211095\tNo\n013557\tNo\n2227145\tNo'
        )

        excel = pd.ExcelFile(resultado['excel'])
        hoja_precintos = excel.parse('Precintos cargados', dtype=str).fillna('')

        self.assertEqual(hoja_precintos['precinto cargado'].tolist(), ['2211095', '013557', '2227145'])
        self.assertEqual(hoja_precintos['coincide en saeplus'].tolist(), ['Si', 'Si', 'No'])
        self.assertEqual(
            hoja_precintos['estatus'].tolist(),
            ['ACTIVO', 'ACTIVO', '']
        )
        self.assertEqual(
            hoja_precintos.loc[hoja_precintos['precinto cargado'] == '013557', 'n° abonado'].str.replace('.0', '', regex=False).tolist(),
            ['8002']
        )

    def test_procesar_comparativo_precintos_funciona_sin_smartolt_si_hay_precintos(self):
        saeplus_base = pd.DataFrame({
            'EQUIPO MAC': [
                'AA:BB:CC:11:22:12345678',
                'AA:BB:CC:11:22:87654321',
                'AA:BB:CC:11:22:11111111',
            ],
            'N° Abonado': [8051, 8052, 8053],
            'documento': ['10', '20', '30'],
            'nombre': ['Ana', 'Luis', 'Carla'],
            'estatus': ['ACTIVO', 'ACTIVO', 'ACTIVO'],
            'precinto': ['PREC-10', '', ''],
        })
        saeplus_ubicacion = pd.DataFrame({
            'N° Abonado': [8051, 8052, 8053],
            'Barrio': ['Centro', 'Centro', 'Bosques'],
            'Dirección': ['Cra 1 # 10-20', 'Cra 1 # 10-25', 'Calle 8 # 15-30'],
            'Ciudad': ['Pereira', 'Pereira', 'Cartago'],
        })

        saeplus_excel = excel_desde_hojas([
            ('Base', saeplus_base),
            ('Ubicaciones', saeplus_ubicacion),
        ])
        resultado = procesar_comparativo_precintos(
            saeplus_excel,
            None,
            'PREC-10\nPREC-99'
        )

        self.assertEqual(resultado['num_casos'], 2)
        self.assertEqual(resultado['data']['precinto cargado'].tolist(), ['PREC-10', 'PREC-99'])

        excel = pd.ExcelFile(resultado['excel'])
        self.assertEqual(excel.sheet_names, ['Precintos cargados'])
        hoja_precintos = excel.parse('Precintos cargados', dtype=str).fillna('')

        self.assertEqual(hoja_precintos['coincide en saeplus'].tolist(), ['Si', 'No'])

    def test_precintos_cargados_conserva_cualquier_estatus_de_saeplus(self):
        saeplus_base = pd.DataFrame({
            'EQUIPO MAC': [
                'AA:BB:CC:11:22:12345678',
                'AA:BB:CC:11:22:87654321',
            ],
            'N° Abonado': [8061, 8062],
            'documento': ['10', '20'],
            'nombre': ['Ana', 'Luis'],
            'estatus': ['SUSPENDIDO', 'RETIRADO'],
            'precinto': ['PREC-10', 'PREC-20'],
        })
        saeplus_ubicacion = pd.DataFrame({
            'N° Abonado': [8061, 8062],
            'Barrio': ['Centro', 'Bosques'],
            'Dirección': ['Cra 1 # 10-20', 'Calle 8 # 15-30'],
            'Ciudad': ['Pereira', 'Cartago'],
        })

        saeplus_excel = excel_desde_hojas([
            ('Base', saeplus_base),
            ('Ubicaciones', saeplus_ubicacion),
        ])

        resultado = procesar_comparativo_precintos(
            saeplus_excel,
            None,
            'PREC-10\nPREC-20'
        )

        hoja_precintos = pd.ExcelFile(resultado['excel']).parse('Precintos cargados', dtype=str).fillna('')

        self.assertEqual(hoja_precintos['coincide en saeplus'].tolist(), ['Si', 'Si'])
        self.assertEqual(hoja_precintos['estatus'].tolist(), ['SUSPENDIDO', 'RETIRADO'])

    def test_precintos_cargados_toma_el_ultimo_estatus_en_duplicados(self):
        saeplus_base = pd.DataFrame({
            'EQUIPO MAC': [
                'AA:BB:CC:11:22:12345678',
                'AA:BB:CC:11:22:12345678',
            ],
            'N° Abonado': [8063, 8063],
            'documento': ['10', '10'],
            'nombre': ['Ana', 'Ana'],
            'estatus': ['ACTIVO', 'POR SUSPENDER'],
            'precinto': ['PREC-10', 'PREC-10'],
        })
        saeplus_ubicacion = pd.DataFrame({
            'N° Abonado': [8063],
            'Barrio': ['Centro'],
            'Dirección': ['Cra 1 # 10-20'],
            'Ciudad': ['Pereira'],
        })

        saeplus_excel = excel_desde_hojas([
            ('Base', saeplus_base),
            ('Ubicaciones', saeplus_ubicacion),
        ])

        resultado = procesar_comparativo_precintos(
            saeplus_excel,
            None,
            'PREC-10'
        )

        hoja_precintos = pd.ExcelFile(resultado['excel']).parse('Precintos cargados', dtype=str).fillna('')

        self.assertEqual(hoja_precintos['coincide en saeplus'].tolist(), ['Si'])
        self.assertEqual(hoja_precintos['estatus'].tolist(), ['POR SUSPENDER'])

    def test_precintos_cargados_prioriza_por_suspender_y_no_descarta_sin_equipo(self):
        saeplus_base = pd.DataFrame({
            'EQUIPO MAC': [
                None,
                'AA:BB:CC:11:22:87654321',
                'AA:BB:CC:11:22:11111111',
            ],
            'N° Abonado': ['C015233', 'C003411', 'C015814'],
            'documento': ['14568397', '1006318164', '1112779494'],
            'nombre': ['NILSON DUVAN', 'CAMILA', 'LUIS FERNEY'],
            'estatus': ['POR SUSPENDER', 'POR SUSPENDER', 'ACTIVO'],
            'precinto': ['2446047', '2147228', '2147228'],
            'Barrio': ['EL COFRE', 'SANTA MARIA', 'EL COFRE'],
            'Dirección': ['CR 4-# 25-32, EL COFRE', 'CR 3 B-# 23-26 MZ 16, SANTA MARIA', 'CR 4 CL 25-52, EL COFRE'],
            'Ciudad': ['CARTAGO', 'CARTAGO', 'CARTAGO'],
        })

        resultado = procesar_comparativo_precintos(
            excel_buffer(saeplus_base),
            None,
            '2446047\n2147228'
        )

        hoja_precintos = pd.ExcelFile(resultado['excel']).parse('Precintos cargados', dtype=str).fillna('')

        self.assertEqual(hoja_precintos['precinto cargado'].tolist(), ['2446047', '2147228'])
        self.assertEqual(hoja_precintos['coincide en saeplus'].tolist(), ['Si', 'Si'])
        self.assertEqual(hoja_precintos['estatus'].tolist(), ['POR SUSPENDER', 'POR SUSPENDER'])
        self.assertEqual(hoja_precintos['n° abonado'].tolist(), ['C015233', 'C003411'])

    def test_procesar_comparativo_precintos_filtra_por_olt_zona_y_barrio_cuando_hay_precintos(self):
        saeplus_base = pd.DataFrame({
            'EQUIPO MAC': [
                'AA:BB:CC:11:22:12345678',
                'AA:BB:CC:11:22:87654321',
                'AA:BB:CC:11:22:11111111',
                'AA:BB:CC:11:22:22222222',
            ],
            'N° Abonado': [8101, 8102, 8103, 8104],
            'documento': ['11', '22', '33', '44'],
            'nombre': ['Ana', 'Luis', 'Carla', 'Pedro'],
            'estatus': ['ACTIVO', 'ACTIVO', 'ACTIVO', 'ACTIVO'],
            'precinto': ['', 'PREC-10', '', ''],
        })
        saeplus_ubicacion = pd.DataFrame({
            'N° Abonado': [8101, 8102, 8103, 8104],
            'Barrio': ['Centro', 'Centro', 'Bosques', 'Centro'],
            'Dirección': [
                'Cra 9 # 1-10',
                'Calle 50 # 20-10',
                'Cra 1 # 1-10',
                'Avenida 3 # 40-10',
            ],
            'Ciudad': ['Pereira', 'Pereira', 'Pereira', 'Pereira'],
        })
        olt = pd.DataFrame({
            'SN': ['SN12345678', 'SN87654321', 'SN11111111', 'SN22222222'],
            'name': ['ONT Ana', 'ONT Luis', 'ONT Carla', 'ONT Pedro'],
            'status': ['Offline', 'Online', 'LOS', 'Offline'],
            'olt': ['OLT-1', 'OLT-1', 'OLT-1', 'OLT-1'],
            'board': ['1', '1', '1', '2'],
            'port': ['3', '3', '3', '3'],
        })

        saeplus_excel = excel_desde_hojas([
            ('Base', saeplus_base),
            ('Ubicaciones', saeplus_ubicacion),
        ])
        resultado = procesar_comparativo_precintos(
            saeplus_excel,
            csv_buffer(olt),
            'PREC-10'
        )

        excel = pd.ExcelFile(resultado['excel'])
        hoja_alertas = excel.parse('Alertas cercanas a precintos', dtype=str).fillna('')

        self.assertEqual(hoja_alertas['n° abonado alerta'].tolist(), ['8101', '8104'])
        self.assertEqual(hoja_alertas['criterio ubicación'].tolist(), ['Mismo barrio', 'Mismo barrio'])

    def test_procesar_comparativo_precintos_agrega_hoja_con_todos_los_estados_smartolt(self):
        saeplus_base = pd.DataFrame({
            'EQUIPO MAC': [
                'AA:BB:CC:11:22:12345678',
                'AA:BB:CC:11:22:87654321',
                'AA:BB:CC:11:22:11111111',
            ],
            'N° Abonado': [8201, 8202, 8203],
            'documento': ['11', '22', '33'],
            'nombre': ['Ana', 'Luis', 'Carla'],
            'estatus': ['ACTIVO', 'ACTIVO', 'CORTADO'],
            'precinto': ['', 'PREC-10', ''],
        })
        saeplus_ubicacion = pd.DataFrame({
            'N° Abonado': [8201, 8202, 8203],
            'Barrio': ['Centro', 'Centro', 'Centro'],
            'Dirección': [
                'Cra 1 # 10-20',
                'Cra 1 # 10-25',
                'Cra 1 # 10-40',
            ],
            'Ciudad': ['Pereira', 'Pereira', 'Pereira'],
        })
        olt = pd.DataFrame({
            'SN': ['SN12345678', 'SN87654321', 'SN11111111'],
            'name': ['ONT Ana', 'ONT Luis', 'ONT Carla'],
            'status': ['Online', 'Online', 'LOS'],
            'olt': ['OLT-1', 'OLT-1', 'OLT-1'],
        })

        saeplus_excel = excel_desde_hojas([
            ('Base', saeplus_base),
            ('Ubicaciones', saeplus_ubicacion),
        ])
        resultado = procesar_comparativo_precintos(
            saeplus_excel,
            csv_buffer(olt),
            'PREC-10'
        )

        excel = pd.ExcelFile(resultado['excel'])
        hoja_alertas = excel.parse('Alertas cercanas a precintos', dtype=str).fillna('')
        hoja_todos = excel.parse('Todos los estados SmartOLT', dtype=str).fillna('')

        self.assertEqual(hoja_alertas['n° abonado alerta'].tolist(), ['8203'])
        self.assertEqual(hoja_todos['n° abonado'].tolist(), ['8201', '8203'])
        self.assertEqual(hoja_todos['status'].tolist(), ['Online', 'LOS'])
        self.assertEqual(hoja_todos['prioridad'].tolist(), ['Alta', 'Alta'])

    def test_procesar_comparativo_precintos_incluye_por_suspender_y_por_cortar(self):
        saeplus_base = pd.DataFrame({
            'EQUIPO MAC': [
                'AA:BB:CC:11:22:12345678',
                'AA:BB:CC:11:22:87654321',
                'AA:BB:CC:11:22:11111111',
            ],
            'N° Abonado': [8251, 8252, 8253],
            'documento': ['11', '22', '33'],
            'nombre': ['Referencia', 'Ana', 'Luis'],
            'estatus': ['ACTIVO', 'POR SUSPENDER', 'POR CORTAR'],
            'precinto': ['PREC-10', '', ''],
        })
        saeplus_ubicacion = pd.DataFrame({
            'N° Abonado': [8251, 8252, 8253],
            'Barrio': ['Centro', 'Centro', 'Centro'],
            'Dirección': [
                'Cra 1 # 10-25',
                'Cra 1 # 10-20',
                'Cra 1 # 10-40',
            ],
            'Ciudad': ['Pereira', 'Pereira', 'Pereira'],
        })
        olt = pd.DataFrame({
            'SN': ['SN12345678', 'SN87654321', 'SN11111111'],
            'name': ['ONT Referencia', 'ONT Ana', 'ONT Luis'],
            'status': ['Online', 'Offline', 'LOS'],
            'olt': ['OLT-1', 'OLT-1', 'OLT-1'],
        })

        saeplus_excel = excel_desde_hojas([
            ('Base', saeplus_base),
            ('Ubicaciones', saeplus_ubicacion),
        ])
        resultado = procesar_comparativo_precintos(
            saeplus_excel,
            csv_buffer(olt),
            'PREC-10'
        )

        excel = pd.ExcelFile(resultado['excel'])
        hoja_alertas = excel.parse('Alertas cercanas a precintos', dtype=str).fillna('')
        hoja_todos = excel.parse('Todos los estados SmartOLT', dtype=str).fillna('')

        self.assertEqual(hoja_alertas['estatus alerta'].tolist(), ['POR CORTAR'])
        self.assertEqual(hoja_todos['estatus'].tolist(), ['POR CORTAR', 'POR SUSPENDER'])

    def test_alertas_cercanas_a_precintos_excluye_suspendidos_y_deja_un_solo_registro_por_abonado(self):
        saeplus_base = pd.DataFrame({
            'EQUIPO MAC': [
                'AA:BB:CC:11:22:12345678',
                'AA:BB:CC:11:22:87654321',
                'AA:BB:CC:11:22:11111111',
                'AA:BB:CC:11:22:22222222',
            ],
            'N° Abonado': [8301, 8302, 8303, 8304],
            'documento': ['11', '22', '33', '44'],
            'nombre': ['Precinto Uno', 'Precinto Dos', 'Alerta Valida', 'Alerta Excluida'],
            'estatus': ['ACTIVO', 'ACTIVO', 'ACTIVO', 'POR SUSPENDER'],
            'precinto': ['PREC-10', 'PREC-20', '', ''],
        })
        saeplus_ubicacion = pd.DataFrame({
            'N° Abonado': [8301, 8302, 8303, 8304],
            'Barrio': ['Centro', 'Centro', 'Centro', 'Centro'],
            'Dirección': [
                'Cra 1 # 10-25',
                'Cra 1 # 10-26',
                'Cra 1 # 10-20',
                'Cra 1 # 10-21',
            ],
            'Ciudad': ['Pereira', 'Pereira', 'Pereira', 'Pereira'],
        })
        olt = pd.DataFrame({
            'SN': ['SN12345678', 'SN87654321', 'SN11111111', 'SN22222222'],
            'name': ['ONT Ref 1', 'ONT Ref 2', 'ONT Alerta Valida', 'ONT Alerta Excluida'],
            'status': ['Online', 'Online', 'Offline', 'Power fail'],
            'olt': ['OLT-1', 'OLT-1', 'OLT-1', 'OLT-1'],
        })

        saeplus_excel = excel_desde_hojas([
            ('Base', saeplus_base),
            ('Ubicaciones', saeplus_ubicacion),
        ])
        resultado = procesar_comparativo_precintos(
            saeplus_excel,
            csv_buffer(olt),
            'PREC-10\nPREC-20'
        )

        hoja_alertas = pd.ExcelFile(resultado['excel']).parse('Alertas cercanas a precintos', dtype=str).fillna('')

        self.assertEqual(hoja_alertas['n° abonado alerta'].tolist(), ['8303'])
        self.assertEqual(hoja_alertas['estatus alerta'].tolist(), ['ACTIVO'])
        self.assertEqual(hoja_alertas['status smartolt alerta'].tolist(), ['Offline'])

    def test_procesar_precintos_cercanos_relaciona_ubicacion_y_alertas(self):
        saeplus_base = pd.DataFrame({
            'EQUIPO MAC': [
                'AA:BB:CC:11:22:12345678',
                'AA:BB:CC:11:22:87654321',
                'AA:BB:CC:11:22:11111111',
                'AA:BB:CC:11:22:22222222',
                'AA:BB:CC:11:22:33333333',
            ],
            'N° Abonado': [7001, 7002, 7003, 7004, 7005],
            'documento': ['11', '22', '33', '44', '55'],
            'nombre': ['Ana', 'Luis', 'Carla', 'Pedro', 'Marta'],
            'estatus': ['ACTIVO', 'ACTIVO', 'ACTIVO', 'ACTIVO', 'ACTIVO'],
            'precinto': ['', 'PREC-10', '', 'PREC-20', ''],
        })
        saeplus_ubicacion = pd.DataFrame({
            'N° Abonado': [7001, 7002, 7003, 7004, 7005],
            'Barrio': ['Centro', 'Centro', 'Bosques', 'Centro', 'Alamos'],
            'Dirección': [
                'Cra 1 # 10-20',
                'Cra 1 # 10-25',
                'Calle 8 # 15-30',
                'Cra 1 # 11-05',
                'Mz 9 Casa 1',
            ],
            'Ciudad': ['Pereira', 'Pereira', 'Cartago', 'Pereira', 'Cartago'],
        })
        olt = pd.DataFrame({
            'SN': ['SN12345678', 'SN87654321', 'SN11111111', 'SN22222222', 'SN33333333'],
            'name': ['ONT Ana', 'ONT Luis', 'ONT Carla', 'ONT Pedro', 'ONT Marta'],
            'status': ['Online', 'Offline', 'LOS', 'Power fail', 'Online'],
            'olt': ['OLT-1', 'OLT-1', 'OLT-2', 'OLT-3', 'OLT-4'],
        })

        saeplus_excel = excel_desde_hojas([
            ('Base', saeplus_base),
            ('Ubicaciones', saeplus_ubicacion),
        ])
        resultado = procesar_precintos_cercanos(saeplus_excel, csv_buffer(olt))

        self.assertEqual(resultado['num_casos'], 2)
        self.assertEqual(resultado['data']['n° abonado'].tolist(), ['7001', '7003'])
        self.assertEqual(resultado['data']['cantidad alertas cercanas'].tolist(), [2, 1])
        self.assertEqual(resultado['data'].iloc[0]['barrio'], 'Centro')
        self.assertEqual(resultado['data'].iloc[0]['referencia dirección'], '1 10')
        self.assertIn('7002', resultado['data'].iloc[0]['abonados alerta cercanos'])
        self.assertIn('7004', resultado['data'].iloc[0]['abonados alerta cercanos'])
        self.assertEqual(
            resultado['data'].loc[resultado['data']['n° abonado'] == '7003', 'criterios detectados'].tolist(),
            ['Mismo abonado']
        )

        excel = pd.ExcelFile(resultado['excel'])
        self.assertEqual(excel.sheet_names, ['Zonas sugeridas', 'Detalle cruce'])
        hoja_detalle = excel.parse('Detalle cruce')
        self.assertIn('Mismo barrio y dirección aproximada', hoja_detalle['criterio cercanía'].tolist())
        self.assertIn('Mismo abonado', hoja_detalle['criterio cercanía'].tolist())
        self.assertNotIn(7005, hoja_detalle['n° abonado sin precinto'].dropna().tolist())


class NavegacionServiceTests(unittest.TestCase):
    def test_procesar_sin_navegar_detecta_activo_sin_navegacion(self):
        abonados = pd.DataFrame({
            'EQUIPO MAC': ['AA:BB:CC:11:22:12345678'],
            'N° Abonado': [5001],
            'documento': ['90'],
            'nombre': ['Luisa'],
            'estatus': ['ACTIVO'],
            'detalle suscripcion': ['GPON 50 MG'],
            'nombre franquicia': ['Norte'],
            'tipo tecnología.': ['GPON'],
            'catv abonado': ['Disabled'],
            'administrative status abonado': ['Disabled'],
        })
        olt = pd.DataFrame({
            'SN': ['SN12345678'],
            'name': ['ONT Luisa'],
            'status': ['Offline'],
            'olt': ['OLT-3'],
            'board': ['1'],
            'port': ['10'],
            'service port upload speed': ['10MG'],
            'service port download speed': ['50MG'],
            'catv': ['Disabled'],
            'administrative status': ['Disabled'],
        })

        resultado = procesar_sin_navegar(excel_buffer(abonados), csv_buffer(olt))

        self.assertEqual(resultado['num_casos'], 2)
        self.assertEqual(resultado['data']['n° abonado'].tolist(), [5001])

    def test_procesar_navegacion_activos_sin_navegar(self):
        resultado = procesar_navegacion_activos_sin_navegar(*navegacion_archivos())

        self.assertEqual(resultado['num_casos'], 1)
        self.assertEqual(resultado['data']['n° abonado'].tolist(), [5001])

    def test_procesar_navegacion_desactivos_con_internet(self):
        resultado = procesar_navegacion_desactivos_con_internet(*navegacion_archivos())

        self.assertEqual(resultado['num_casos'], 1)
        self.assertEqual(resultado['data']['n° abonado'].tolist(), [5002])

    def test_procesar_navegacion_estado_servicio_retorna_dos_hojas(self):
        resultado = procesar_navegacion_estado_servicio(*navegacion_archivos())

        self.assertEqual(resultado['num_casos'], 2)
        self.assertEqual(resultado['data']['n° abonado'].tolist(), [5001])

        excel = pd.ExcelFile(resultado['excel'])
        self.assertEqual(excel.sheet_names, ['Activos sin navegar', 'Desactivos con internet'])
        self.assertEqual(excel.parse('Activos sin navegar')['n° abonado'].tolist(), [5001])
        self.assertEqual(excel.parse('Desactivos con internet')['n° abonado'].tolist(), [5002])

    def test_procesar_navegacion_activos_sin_catv(self):
        resultado = procesar_navegacion_activos_sin_catv(*navegacion_archivos())

        self.assertEqual(resultado['num_casos'], 1)
        self.assertEqual(resultado['data']['n° abonado'].tolist(), [5003])

    def test_procesar_navegacion_solo_con_arroba_y_catv_activo(self):
        resultado = procesar_navegacion_solo_con_arroba_y_catv_activo(*navegacion_archivos())

        self.assertEqual(resultado['num_casos'], 1)
        self.assertEqual(resultado['data']['n° abonado'].tolist(), [5004])

    def test_procesar_navegacion_catv_y_planes_retorna_dos_hojas(self):
        resultado = procesar_navegacion_catv_y_planes(*navegacion_archivos())

        self.assertEqual(resultado['num_casos'], 2)
        self.assertEqual(resultado['data']['n° abonado'].tolist(), [5003])

        excel = pd.ExcelFile(resultado['excel'])
        self.assertEqual(excel.sheet_names, ['Activos sin Catv', 'Solo con @ y catv activo'])
        self.assertEqual(excel.parse('Activos sin Catv')['n° abonado'].tolist(), [5003])
        self.assertEqual(excel.parse('Solo con @ y catv activo')['n° abonado'].tolist(), [5004])


class AuditoriaReconexionesServiceTests(unittest.TestCase):
    def test_procesar_auditoria_reconexiones_detecta_abonado_sin_activar(self):
        drive = pd.DataFrame({
            'Cuenta': [6001],
            'documento': ['11'],
            'nombre': ['Ana'],
            'apellido': ['Ruiz'],
            'observaciones': [None],
            'estatus drive': ['PENDIENTE'],
            'detalle suscripcion drive': ['PLAN @ 100 MG'],
            'saldo drive': [0],
        })
        saeplus = pd.DataFrame({
            'Cuenta': [6001],
            'documento': ['11'],
            'nombre': ['Ana'],
            'apellido': ['Ruiz'],
            'estatus': ['ACTIVO'],
            'detalle suscripcion': ['PLAN @ 100 MG'],
            'saldo': [0],
            'observaciones sae': [None],
            'EQUIPO MAC': ['AA:BB:CC:11:22:12345678'],
            'N° Abonado': [6001],
        })
        epayco = pd.DataFrame({
            'Cuenta': [6001],
            'documento': ['22'],
            'nombre': ['Otro'],
            'apellido': ['Caso'],
            'observaciones': ['ok'],
            'estatus': ['ACTIVO'],
            'detalle suscripcion': ['PLAN 50 MG'],
        })
        olt = pd.DataFrame({
            'SN': ['SN12345678'],
            'status': ['Offline'],
            'olt': ['OLT-4'],
            'catv': ['Disabled'],
            'administrative status': ['Disabled'],
        })

        resultado = procesar_auditoria_reconexiones(
            excel_buffer(drive),
            excel_buffer(saeplus),
            excel_buffer(epayco),
            csv_buffer(olt)
        )

        self.assertEqual(resultado['num_casos'], 1)
        self.assertEqual(resultado['data']['abonados'].tolist(), [6001])


class AtenuacionesServiceTests(unittest.TestCase):
    def test_procesar_atenuaciones_genera_agenda_con_prioridad(self):
        olt = pd.DataFrame({
            'OLT': ['OLT-A', 'OLT-A', 'OLT-A'],
            'Board': ['1', '1', '1'],
            'Port': ['1', '1', '1'],
            'Signal 1310': [-34, -33, -32],
            'Signal 1490': [-27, -27, -27],
            'ONU external ID': ['ONU1', 'ONU2', 'ONU3'],
            'Status': ['LOS', 'LOS', 'Online'],
            'Address': ['Zona 1', 'Zona 1', 'Zona 1'],
        })

        resultado = procesar_atenuaciones(csv_buffer(olt))

        self.assertGreater(resultado['num_casos'], 0)
        self.assertIn('prioridad', resultado['data'].columns)
        self.assertEqual(resultado['data'].iloc[0]['danio_dominante'], 'FIBRA')


class EstadisticosOLTServiceTests(unittest.TestCase):
    def test_procesar_estadisticos_olt_genera_resumen_y_hojas(self):
        olt = pd.DataFrame({
            'SN': ['SN1', 'SN2', 'SN3'],
            'OLT': ['OLT-A', 'OLT-A', 'OLT-B'],
            'Board': ['1', '1', '2'],
            'Port': ['1', '1', '3'],
            'Status': ['Online', 'Offline', 'Online'],
            'Signal': ['Warning', 'Critical', 'Normal'],
            'Signal 1310': [-25, '-31', -28],
            'Signal 1490': [-24, -29, -27],
            'CATV': ['Enabled', 'Disabled', 'Enabled'],
            'Administrative status': ['Enabled', 'Disabled', 'Enabled'],
        })

        resultado = procesar_estadisticos_olt(csv_buffer(olt))

        self.assertEqual(resultado['num_casos'], 3)
        self.assertEqual(resultado['summary']['Total OLT'], 2)
        self.assertEqual(resultado['summary']['Online'], 2)
        self.assertEqual(resultado['summary']['Offline'], 1)
        self.assertEqual(resultado['summary']['Warning'], 1)
        self.assertEqual(resultado['summary']['Critical'], 1)
        self.assertEqual(resultado['data']['total abonados'].tolist(), [2, 1])

        excel = pd.ExcelFile(resultado['excel'])
        self.assertEqual(
            excel.sheet_names,
            [
                'Resumen por OLT',
                'Status por OLT',
                'Signal por OLT',
                'Board Port',
                'CATV',
                'Administrative Status',
            ]
        )
        resumen = excel.parse('Resumen por OLT')
        self.assertEqual(resumen.loc[0, 'total abonados'], 2)
        self.assertEqual(resumen.loc[0, 'peor signal 1310'], -31)

    def test_job_estadisticos_olt_lanza_error_si_faltan_columnas(self):
        with self.assertRaises(ValueError) as ctx:
            job_estadisticos_olt(pd.DataFrame({'olt': ['OLT-A']}))

        self.assertIn('sn', str(ctx.exception))


class CoincidenciaEnFilaServiceTests(unittest.TestCase):
    def test_procesar_coincidencia_en_fila_devuelve_hojas_de_coinciden_y_no_coinciden(self):
        abonados = pd.DataFrame({
            'ABONADO': ['000123', '789', None],
            'n° abonado': [123, 456, 999],
            'documento': ['10', '20', '30'],
        })

        resultado = procesar_coincidencia_en_fila(excel_buffer(abonados))

        self.assertEqual(resultado['num_casos'], 1)
        self.assertEqual(resultado['data']['valor comparado'].tolist(), ['123'])
        self.assertIn('hoja origen', resultado['data'].columns)
        self.assertIn('documento', resultado['data'].columns)
        self.assertEqual(resultado['data']['valor en ABONADO'].tolist(), ['000123'])
        self.assertEqual(resultado['data']['filas ABONADO'].tolist(), ['2'])
        self.assertEqual(resultado['data']['fila n° abonado'].astype(str).tolist(), ['2'])

        excel = pd.ExcelFile(resultado['excel'])
        self.assertEqual(excel.sheet_names, ['Coinciden', 'No coinciden'])
        hoja_coinciden = excel.parse('Coinciden', dtype=str).fillna('')
        hoja_no_coinciden = excel.parse('No coinciden', dtype=str).fillna('')

        self.assertEqual(hoja_coinciden['valor comparado'].tolist(), ['123'])
        self.assertEqual(hoja_coinciden['documento'].tolist(), ['10'])
        self.assertEqual(hoja_no_coinciden['valor comparado'].tolist(), ['456', '999'])
        self.assertEqual(hoja_no_coinciden['documento'].tolist(), ['20', '30'])

    def test_procesar_coincidencia_en_fila_retorna_cero_si_no_hay_match(self):
        abonados = pd.DataFrame({
            'ABONADO': [111, 222],
            'n° abonado': [333, 444],
            'documento': ['10', '20'],
        })

        resultado = procesar_coincidencia_en_fila(excel_buffer(abonados))

        self.assertEqual(resultado['num_casos'], 0)
        self.assertEqual(
            resultado['data'].columns.tolist(),
            ['hoja origen', 'fila n° abonado', 'valor comparado', 'valor en ABONADO', 'filas ABONADO', 'cantidad en ABONADO', 'ABONADO', 'n° abonado', 'documento']
        )

        excel = pd.ExcelFile(resultado['excel'])
        self.assertEqual(excel.sheet_names, ['Coinciden', 'No coinciden'])
        hoja_no_coinciden = excel.parse('No coinciden', dtype=str).fillna('')
        self.assertEqual(hoja_no_coinciden['valor comparado'].tolist(), ['333', '444'])

    def test_procesar_coincidencia_en_fila_normaliza_decimales_y_ceros(self):
        abonados = pd.DataFrame({
            'ABONADO': ['00123.00', "'000456", 'C016454'],
            'n° abonado': [123, '456.0', ' c016454 '],
            'documento': ['10', '20', '30'],
        })

        resultado = procesar_coincidencia_en_fila(excel_buffer(abonados))

        self.assertEqual(resultado['num_casos'], 3)
        self.assertEqual(resultado['data']['valor comparado'].tolist(), ['123', '456', 'C016454'])

    def test_procesar_coincidencia_en_fila_detecta_match_en_filas_distintas(self):
        abonados = pd.DataFrame({
            'ABONADO': ['SR006144', 'X', 'SR002666'],
            'n° abonado': ['Y', 'SR006144', 'SR002666'],
        })

        resultado = procesar_coincidencia_en_fila(excel_buffer(abonados))

        self.assertEqual(resultado['num_casos'], 2)
        self.assertEqual(resultado['data']['valor comparado'].tolist(), ['SR006144', 'SR002666'])
        self.assertEqual(
            resultado['data'].loc[resultado['data']['valor comparado'] == 'SR006144', 'filas ABONADO'].tolist(),
            ['2']
        )
        self.assertEqual(
            resultado['data'].loc[resultado['data']['valor comparado'] == 'SR006144', 'fila n° abonado'].astype(str).tolist(),
            ['3']
        )


class CiudadesAbonadoServiceTests(unittest.TestCase):
    def test_agregar_columna_ciudad_asigna_ciudad_por_prefijo(self):
        abonados = pd.DataFrame({
            'ABONADO': ['CH100', 'C0123', 'DQ888', 'SR111', 'VG222', 'TC999', 'TCF777', 'PQ555', 'SG444', 'ZZ000'],
        })

        resultado = agregar_columna_ciudad(abonados)

        self.assertEqual(
            resultado['CIUDAD'].tolist(),
            [
                'Chinchiná',
                'Cartago',
                'Dosquebradas',
                'Santa Rosa',
                'Virginia',
                'Pereira',
                'PereiraCentro',
                'Parque Industrial',
                'Guaviare',
                None,
            ]
        )


if __name__ == '__main__':
    unittest.main()
