from pathlib import Path
import sys

import pandas as pd

from funciones import abrir_excel_seguro
from services.comparativo_precintos import procesar_comparativo_precintos
from services.comparativo_precintos import normalizar_precinto
from services.precintos_cercanos import normalizar_encabezado


PRECINTOS_OBJETIVO = [
    '1360005',
    '2446381',
    '2147228',
    '2444945',
    '2444969',
    '2192758',
    '1360082',
    '2194108',
    '2444367',
    '2192294',
    '2445811',
    '2446047',
    '2197651',
    '2365335',
]


def main():
    if len(sys.argv) < 2:
        print('Uso: python tmp_revisar_precintos_saeplus.py <ruta_excel>')
        sys.exit(1)

    ruta = Path(' '.join(sys.argv[1:]).strip().strip('"'))
    excel = abrir_excel_seguro(ruta)
    objetivos = {normalizar_precinto(valor): valor for valor in PRECINTOS_OBJETIVO}

    print(f'ARCHIVO: {ruta}')
    print('HOJAS:', excel.sheet_names)

    hallazgos = []
    for nombre_hoja in excel.sheet_names:
        df = excel.parse(sheet_name=nombre_hoja)
        df.columns = [normalizar_encabezado(columna) for columna in df.columns]

        if 'precinto' not in df.columns:
            continue

        trabajo = df.copy()
        trabajo['precinto_normalizado'] = trabajo['precinto'].apply(normalizar_precinto)
        encontrados = trabajo[trabajo['precinto_normalizado'].isin(objetivos)].copy()
        if encontrados.empty:
            continue

        encontrados.insert(0, 'hoja', nombre_hoja)
        hallazgos.append(encontrados)

    if not hallazgos:
        print('SIN_HALLAZGOS')
        return

    resultado = pd.concat(hallazgos, ignore_index=True, sort=False)
    columnas_preferidas = [
        'hoja',
        'precinto',
        'precinto_normalizado',
        'estatus',
        'n abonado',
        'documento',
        'nombre',
        'equipo mac',
        'equipo maco',
        'barrio',
        'direccion',
        'ciudad',
    ]
    columnas = [columna for columna in columnas_preferidas if columna in resultado.columns]
    if 'fila_excel' not in resultado.columns:
        resultado['fila_excel'] = resultado.index + 2
        columnas.append('fila_excel')

    print('TOTAL_HALLAZGOS:', len(resultado))
    print(resultado[columnas].to_string(index=False))

    faltantes = [
        original
        for normalizado, original in objetivos.items()
        if normalizado not in set(resultado['precinto_normalizado'].tolist())
    ]
    print('FALTANTES:', faltantes)

    print('\n--- RESULTADO DEL SERVICIO ---')
    with open(ruta, 'rb') as archivo:
        contenido = archivo.read()

    resultado_servicio = procesar_comparativo_precintos(
        pd.io.common.BytesIO(contenido),
        None,
        '\n'.join(PRECINTOS_OBJETIVO)
    )
    hoja_precintos = pd.ExcelFile(resultado_servicio['excel']).parse('Precintos cargados', dtype=str).fillna('')
    print(hoja_precintos.to_string(index=False))


if __name__ == '__main__':
    main()
