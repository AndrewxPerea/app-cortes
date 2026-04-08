from pathlib import Path
import sys

import pandas as pd

from services.coincidencia_en_fila import normalizar_valor


def main():
    path = Path(sys.argv[1].strip().strip('"'))
    df = pd.read_excel(path)

    print("FILAS:", len(df))
    print("COLUMNAS:", list(df.columns))

    for posicion in [317, 446]:
        for indice in [posicion, posicion - 1]:
            if 0 <= indice < len(df):
                fila = df.iloc[indice]
                abonado = fila.get('ABONADO')
                numero = fila.get('n° abonado')
                print(
                    f"POSICION_SOLICITADA={posicion} INDICE_REAL={indice} "
                    f"ABONADO={abonado!r} N_ABONADO={numero!r} "
                    f"NORM_ABONADO={normalizar_valor(abonado)!r} "
                    f"NORM_N_ABONADO={normalizar_valor(numero)!r}"
                )

    if 'ABONADO' in df.columns and 'n° abonado' in df.columns:
        for objetivo in ['SR006144', 'SR002666']:
            mascara_abonado = df['ABONADO'].astype(str).str.strip().eq(objetivo)
            mascara_numero = df['n° abonado'].astype(str).str.strip().eq(objetivo)
            print(f"BUSQUEDA_{objetivo}_ABONADO:", df.index[mascara_abonado].tolist()[:20])
            print(f"BUSQUEDA_{objetivo}_N_ABONADO:", df.index[mascara_numero].tolist()[:20])
            comunes = sorted(set(df.index[mascara_abonado].tolist()) & set(df.index[mascara_numero].tolist()))
            print(f"BUSQUEDA_{objetivo}_COINCIDEN_MISMA_FILA:", comunes[:20])

        normalizados = pd.DataFrame({
            'ABONADO': df['ABONADO'],
            'n° abonado': df['n° abonado'],
            'abonado_norm': df['ABONADO'].map(normalizar_valor),
            'n_abonado_norm': df['n° abonado'].map(normalizar_valor),
        })
        coincidencias = normalizados[
            normalizados['abonado_norm'].notna() &
            (normalizados['abonado_norm'] == normalizados['n_abonado_norm'])
        ].copy()
        print("TOTAL_COINCIDENCIAS:", len(coincidencias))
        if not coincidencias.empty:
            print(coincidencias.head(20).to_string())


if __name__ == "__main__":
    main()
