import pandas as pd

def procesar_excel(archivo_excel):
    df = pd.read_excel(archivo_excel)

    # Definir el mapeo de valores del plan
    valor_plan_mapping = {
        '100 MG': 106000,
        '100 MG PA 5': 117000,
        '100 MG PA 6': 128000,
        '15 MG': 71000,
        '200 MG': 204000,
        '200 MG PA 5': 215000,
        '300 MG': 314000,
        '400 MG': 418000,
        '50 MG': 77000,
        '50 MG PA 5': 88000,
        'SOLO @ 50 MG': 64000,
        '30 MG': 70000,
        '70 MG': 87000
    }

    # Procesar cada fila
    df['Valor Plan'] = df['Plan Nuevo'].map(valor_plan_mapping)
    df['Nombre Cliente'] = df['Nombre Cliente'].astype(str)
    df['Nombre Cliente'] = df['Nombre Cliente'].apply(lambda x: x.split()[0].capitalize())
    df['Plan Nuevo'] = df['Plan Nuevo'].astype(str)

    # Generar el mensaje personalizado
    df['Mensaje'] = df.apply(lambda row: (
            f"{row['Nombre Cliente']}, TuCable te informa que el estado de tu solicitud de cambio de plan a "
            f"{row['Plan Nuevo']} Mbps de solo internet, por un valor mensual de $ {row['Valor Plan']} ha sido efectuado exitosamente. "
            "Con esto, procedemos a finalizar tu petición. ¡Te deseamos un feliz día!"
            if '@' in row['Plan Nuevo'] else
            f"{row['Nombre Cliente']}, TuCable te informa que el estado de tu solicitud de cambio de plan a "
            f"{row['Plan Nuevo']}Mbps de internet, por un valor mensual de $ {row['Valor Plan']} ha sido efectuado exitosamente. "
            "Con esto, procedemos a finalizar tu petición. ¡Te deseamos un feliz día!"
        ), axis=1)

    # Guardar el resultado en un nuevo archivo Excel
    output_file = "resultado_procesado.xlsx"
    df.to_excel(output_file, index=False)

    return output_file



