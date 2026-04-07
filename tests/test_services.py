import io
import unittest

import pandas as pd

from services.common import excel_desde_dataframe, excel_desde_hojas, validar_columnas
from services.atenuaciones import procesar_atenuaciones
from services.auditoria_reconexiones import procesar_auditoria_reconexiones
from services.ciudades_abonado import agregar_columna_ciudad
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
            'estatus': ['SUSPENDIDO'],
            'observaciones': [None],
            'ingeniero': ['Carlos'],
        })
        saeplus = pd.DataFrame({
            'EQUIPO MAC': ['AA:BB:CC:11:22:12345678'],
            'N° Abonado': [3001],
            'documento': ['80'],
            'nombre': ['Pedro'],
            'apellido': ['Lopez'],
            'estatus sae': ['SUSPENDIDO'],
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

    def test_procesar_comparativo_precintos_devuelve_solo_coincidencias_ordenadas(self):
        saeplus = pd.DataFrame({
            'EQUIPO MAC': [
                'AA:BB:CC:11:22:12345678',
                'AA:BB:CC:11:22:87654321',
                'AA:BB:CC:11:22:11111111',
                'AA:BB:CC:11:22:22222222',
                'AA:BB:CC:11:22:33333333',
            ],
            'N° Abonado': [4101, 4102, 4103, 4104, 4105],
            'documento': ['10', '20', '30', '40', '50'],
            'nombre': ['Ana', 'Luis', 'Carla', 'Pedro', 'Marta'],
            'estatus': ['ACTIVO', 'ACTIVO', 'EN REVISION', 'SUSPENDIDO', 'ACTIVO'],
            'Barrio': ['Centro', 'La Floresta', 'Alamos', 'Galan', 'Bosques'],
            'Dirección': [
                'Cra 1 # 10-20',
                'Calle 8 # 15-30',
                'Mz 4 Casa 9',
                'Cra 7 # 2-10',
                'Calle 100 # 50-20',
            ],
            'Ciudad': ['Cartago', 'Pereira', 'Dosquebradas', 'Santa Rosa', 'Pereira'],
            'precinto': ['PREC-22', 'PREC-05', '', 'PREC-99', 'PREC-40'],
        })
        olt = pd.DataFrame({
            'SN': ['SN12345678', 'SN87654321', 'SN11111111', 'SN22222222', 'SN99999999'],
            'name': ['ONT Ana', 'ONT Luis', 'ONT Carla', 'ONT Pedro', 'ONT Extra'],
            'status': ['Online', 'Online', 'Offline', 'Offline', 'Offline'],
            'olt': ['OLT-1', 'OLT-1', 'OLT-2', 'OLT-3', 'OLT-4'],
        })

        resultado = procesar_comparativo_precintos(excel_buffer(saeplus), csv_buffer(olt))

        self.assertEqual(resultado['num_casos'], 4)
        self.assertIn('observacion_precinto', resultado['data'].columns)
        self.assertIn('barrio', resultado['data'].columns)
        self.assertIn('dirección', resultado['data'].columns)
        self.assertIn('ciudad', resultado['data'].columns)

        self.assertEqual(
            resultado['data']['n° abonado'].tolist(),
            [4102, 4101, 4103, 4104]
        )
        self.assertEqual(
            resultado['data']['estatus'].tolist(),
            ['ACTIVO', 'ACTIVO', 'EN REVISION', 'SUSPENDIDO']
        )
        self.assertEqual(
            resultado['data']['precinto'].fillna('').tolist(),
            ['PREC-05', 'PREC-22', '', 'PREC-99']
        )
        self.assertTrue((resultado['data']['resultado_comparativo'] == 'Coincide').all())
        self.assertEqual(resultado['data'].iloc[0]['barrio'], 'La Floresta')
        self.assertEqual(resultado['data'].iloc[0]['dirección'], 'Calle 8 # 15-30')
        self.assertEqual(resultado['data'].iloc[0]['ciudad'], 'Pereira')
        self.assertNotIn(4105, resultado['data']['n° abonado'].dropna().tolist())
        self.assertEqual(
            resultado['data'].loc[resultado['data']['n° abonado'] == 4103, 'observacion_precinto'].tolist(),
            ['Precinto vacío en SAEPlus']
        )

        excel = pd.ExcelFile(resultado['excel'])
        self.assertEqual(excel.sheet_names, ['Coinciden'])
        hoja_coinciden = excel.parse('Coinciden')
        self.assertEqual(hoja_coinciden['n° abonado'].tolist(), [4102, 4101, 4103, 4104])


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

    def test_procesar_navegacion_activos_sin_catv(self):
        resultado = procesar_navegacion_activos_sin_catv(*navegacion_archivos())

        self.assertEqual(resultado['num_casos'], 1)
        self.assertEqual(resultado['data']['n° abonado'].tolist(), [5003])

    def test_procesar_navegacion_solo_con_arroba_y_catv_activo(self):
        resultado = procesar_navegacion_solo_con_arroba_y_catv_activo(*navegacion_archivos())

        self.assertEqual(resultado['num_casos'], 1)
        self.assertEqual(resultado['data']['n° abonado'].tolist(), [5004])


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


class CoincidenciaEnFilaServiceTests(unittest.TestCase):
    def test_procesar_coincidencia_en_fila_devuelve_filas_coincidentes(self):
        abonados = pd.DataFrame({
            'ABONADO': ['000123', '789', None],
            'n° abonado': [123, 456, 999],
            'documento': ['10', '20', '30'],
        })

        resultado = procesar_coincidencia_en_fila(excel_buffer(abonados))

        self.assertEqual(resultado['num_casos'], 1)
        self.assertEqual(int(resultado['data']['ABONADO'].iloc[0]), 123)
        self.assertEqual(resultado['data']['documento'].tolist(), [10])
        self.assertIn('hoja origen', resultado['data'].columns)

    def test_procesar_coincidencia_en_fila_retorna_cero_si_no_hay_match(self):
        abonados = pd.DataFrame({
            'ABONADO': [111, 222],
            'n° abonado': [333, 444],
            'documento': ['10', '20'],
        })

        resultado = procesar_coincidencia_en_fila(excel_buffer(abonados))

        self.assertEqual(resultado['num_casos'], 0)
        self.assertEqual(resultado['data'].columns.tolist(), ['hoja origen', 'ABONADO', 'n° abonado', 'documento'])


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
