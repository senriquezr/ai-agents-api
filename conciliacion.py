from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
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

    ruta_destino.write_bytes(contenido)

    return ruta_destino


def ejecutar_conciliacion(
    archivo_bdep: UploadFile,
    archivo_sap: UploadFile,
    archivo_ep: UploadFile,
    archivo_ip: UploadFile,
    archivo_re: UploadFile,
) -> dict:

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

        bdep_df = pd.read_excel(rutas["bdep"])
        sap_df = pd.read_excel(rutas["sap"])
        ep_df = pd.read_excel(rutas["ep"])
        ip_df = pd.read_excel(rutas["ip"])
        re_df = pd.read_excel(rutas["re"])

        return {
            "estado": "OK",
            "mensaje": (
                "Los cinco archivos Excel fueron "
                "recibidos y leídos correctamente."
            ),
            "filas": {
                "bdep": len(bdep_df),
                "sap": len(sap_df),
                "ep": len(ep_df),
                "ip": len(ip_df),
                "re": len(re_df),
            },
            "columnas": {
                "bdep": list(bdep_df.columns),
                "sap": list(sap_df.columns),
                "ep": list(ep_df.columns),
                "ip": list(ip_df.columns),
                "re": list(re_df.columns),
            }
        }
