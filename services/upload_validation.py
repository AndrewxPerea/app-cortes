import os


def validar_archivos_requeridos(request_files, definiciones):
    archivos = {}

    for nombre_campo, extensiones, etiqueta in definiciones:
        archivo = request_files.get(nombre_campo)

        if archivo is None:
            raise ValueError(f"Falta cargar el archivo {etiqueta}.")

        nombre = (archivo.filename or "").strip()
        if not nombre:
            raise ValueError(f"Falta seleccionar el archivo {etiqueta}.")

        extension = os.path.splitext(nombre)[1].lower()
        if extension not in extensiones:
            extensiones_texto = ", ".join(sorted(extensiones))
            raise ValueError(
                f"El archivo {etiqueta} debe tener una de estas extensiones: {extensiones_texto}."
            )

        archivos[nombre_campo] = archivo

    return archivos
