from fastapi import UploadFile


def ejecutar_conciliacion(
    archivos: list[UploadFile]
) -> dict:

    nombres = []

    for archivo in archivos:

        nombres.append(
            archivo.filename
        )

    return {
        "estado": "OK",
        "cantidad_archivos": len(archivos),
        "archivos_recibidos": nombres
    }
