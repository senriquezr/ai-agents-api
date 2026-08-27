from pathlib import Path
from tempfile import TemporaryDirectory
from typing import BinaryIO

from fastapi import UploadFile


EXTENSIONES_PERMITIDAS = {".xlsx", ".xlsm"}


def guardar_archivo(
    archivo: UploadFile,
    carpeta: Path,
    nombre_destino: str,
) -> Path:
    nombre_original = archivo.filename or nombre_destino
    extension = Path(nombre_original).suffix.lower()

    if extension not in EXTENSIONES_PERMITIDAS:
        raise ValueError(
            f"El archivo {nombre_original} debe ser .xlsx o .xlsm."
        )

    ruta_destino = carpeta / f"{nombre_destino}{extension}"

    contenido = archivo.file.read()

    if not contenido:
        raise ValueError(
            f"El archivo {nombre_original} está vacío."
        )

    if not contenido.startswith(b"PK"):
        raise ValueError(
            f"El archivo {nombre_original} no es un Excel válido."
        )

    ruta_destino.write_bytes(contenido)

    return ruta_destino


def ejecutar_conciliacion(
    archivo_bdep: UploadFile,
    archivo_sap: UploadFile,
    archivo_ep: UploadFile,
    archivo_ip: UploadFile,
    archivo_re: UploadFile,
) -> dict:
    archivos_recibidos = {
        "bdep": archivo_bdep.filename,
        "sap": archivo_sap.filename,
        "ep": archivo_ep.filename,
        "ip": archivo_ip.filename,
        "re": archivo_re.filename,
    }

    with TemporaryDirectory() as directorio_temporal:
        carpeta = Path(directorio_temporal)

        rutas = {
            "bdep": guardar_archivo(
                archivo_bdep,
                carpeta,
                "bdep",
            ),
            "sap": guardar_archivo(
                archivo_sap,
                carpeta,
                "sap",
            ),
            "ep": guardar_archivo(
                archivo_ep,
                carpeta,
                "ep",
            ),
            "ip": guardar_archivo(
                archivo_ip,
                carpeta,
                "ip",
            ),
            "re": guardar_archivo(
                archivo_re,
                carpeta,
                "re",
            ),
        }

        tamanos = {
            tipo: ruta.stat().st_size
            for tipo, ruta in rutas.items()
        }

        return {
            "estado": "ARCHIVOS_RECIBIDOS",
            "mensaje": (
                "Los cinco archivos Excel fueron recibidos "
                "y validados correctamente."
            ),
            "archivos": archivos_recibidos,
            "tamanos_bytes": tamanos,
        }
