import io
import unittest

import pandas as pd

from services.common import excel_desde_dataframe, excel_desde_hojas, validar_columnas
from services.atenuaciones import procesar_atenuaciones
from services.auditoria_reconexiones import procesar_auditoria_reconexiones
from services.cortes import procesar_cortes
from services.diferentes import procesar_diferentes
from services.navegacion import procesar_sin_navegar
from services.reconexiones import procesar_reconexiones
from services.upload_validation import validar_archivos_requeridos
from services.velocidad import extraer_velocidad, procesar_verificacion_velocidad


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
        })

        resultado = procesar_verificacion_velocidad(excel_buffer(saeplus), csv_buffer(olt))

        self.assertEqual(resultado['num_casos'], 1)
        self.assertEqual(resultado['data']['n° abonado'].tolist(), [2001])


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


if __name__ == '__main__':
    unittest.main()
