import argparse

from services.ciudades_abonado import guardar_ciudades_abonado


def parse_args():
    parser = argparse.ArgumentParser(
        description=(
            "Agrega la columna CIUDAD a partir del prefijo de la columna ABONADO "
            "y genera un nuevo archivo Excel."
        )
    )
    parser.add_argument(
        'input',
        nargs='?',
        default='Libro3.xlsx',
        help='Ruta del archivo Excel de entrada. Por defecto: Libro3.xlsx',
    )
    parser.add_argument(
        'output',
        nargs='?',
        default='Libro3_con_ciudad.xlsx',
        help='Ruta del archivo Excel de salida. Por defecto: Libro3_con_ciudad.xlsx',
    )
    return parser.parse_args()


def main():
    args = parse_args()
    resultado = guardar_ciudades_abonado(args.input, args.output)
    print(
        f"Archivo generado: {args.output} | "
        f"Hojas procesadas: {len(resultado['data']['HOJA ORIGEN'].unique())} | "
        f"Filas procesadas: {resultado['num_casos']}"
    )


if __name__ == '__main__':
    main()
