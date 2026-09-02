from __future__ import annotations

import base64
import binascii
from io import BytesIO
from typing import Any

import pandas as pd


EXTENSIONES_PERMITIDAS = (".xlsx", ".xlsm", ".xls")


def _nombre(archivo: Any) -> str:
    nombre = getattr(archivo, "name", None)
    if not nombre and isinstance(archivo, dict):
        nombre = archivo.get("name")
    if not nombre:
        raise ValueError("Uno de los archivos no tiene nombre.")
    return str(nombre)


def _contenido_base64(archivo: Any) -> str:
    contenido = getattr(archivo, "contentBytes", None)
    if contenido is None and isinstance(archivo, dict):
        contenido = archivo.get("contentBytes")

    if not contenido:
        raise ValueError(f"El archivo {_nombre(archivo)} no contiene contentBytes.")

    contenido = str(contenido).strip()

    # Acepta tanto Base64 puro como data URI:
    # data:application/vnd...;base64,AAAA...
    if contenido.startswith("data:") and "," in contenido:
        contenido = contenido.split(",", 1)[1]

    return contenido


def _decodificar_excel(archivo: Any, etiqueta: str) -> tuple[str, bytes]:
    nombre = _nombre(archivo)

    if not nombre.lower().endswith(EXTENSIONES_PERMITIDAS):
        raise ValueError(
            f"{etiqueta}: '{nombre}' no es un Excel permitido "
            "(.xlsx, .xlsm o .xls)."
        )

    contenido = _contenido_base64(archivo)

    try:
        datos = base64.b64decode(contenido, validate=True)
    except (binascii.Error, ValueError) as error:
        raise ValueError(
            f"{etiqueta}: el contenido de '{nombre}' no es Base64 válido."
        ) from error

    if not datos:
        raise ValueError(f"{etiqueta}: '{nombre}' está vacío.")

    return nombre, datos


def _inspeccionar_excel(datos: bytes, etiqueta: str, nombre: str) -> dict:
    try:
        archivo_memoria = BytesIO(datos)
        excel = pd.ExcelFile(archivo_memoria)
        hojas = list(excel.sheet_names)

        if not hojas:
            raise ValueError("El libro no contiene hojas.")

        # Lee una muestra de la primera hoja para validar que pandas puede abrirla.
        archivo_memoria.seek(0)
        muestra = pd.read_excel(
            archivo_memoria,
            sheet_name=hojas[0],
            nrows=5,
        )

        return {
            "tipo": etiqueta,
            "name": nombre,
            "bytes": len(datos),
            "hojas": hojas,
            "primera_hoja": hojas[0],
            "columnas_detectadas": [str(c) for c in muestra.columns],
        }

    except Exception as error:
        raise ValueError(
            f"{etiqueta}: no se pudo leer '{nombre}' como archivo Excel: {error}"
        ) from error


def ejecutar_conciliacion(
    bdep: Any,
    sap: Any,
    ep: Any,
    ip: Any,
    re: Any,
) -> dict:
    """
    Punto de entrada llamado por FastAPI.

    Esta versión valida y decodifica los cinco archivos que llegan desde
    Copilot Studio. La lógica contable específica debe implementarse después
    de la sección de validación usando los bytes/dataframes ya disponibles.
    """

    entradas = {
        "BDEP": bdep,
        "SAP": sap,
        "EP": ep,
        "IP": ip,
        "RE": re,
    }

    binarios: dict[str, bytes] = {}
    archivos_validados: list[dict] = []

    for etiqueta, archivo in entradas.items():
        nombre, datos = _decodificar_excel(archivo, etiqueta)
        binarios[etiqueta] = datos
        archivos_validados.append(
            _inspeccionar_excel(
                datos=datos,
                etiqueta=etiqueta,
                nombre=nombre,
            )
        )

    # ------------------------------------------------------------------
    # AQUÍ VA LA LÓGICA REAL DE CONCILIACIÓN.
    #
    # Ejemplo de cómo convertir cualquiera de los cinco archivos a DataFrame:
    #
    # df_bdep = pd.read_excel(BytesIO(binarios["BDEP"]))
    # df_sap  = pd.read_excel(BytesIO(binarios["SAP"]))
    # df_ep   = pd.read_excel(BytesIO(binarios["EP"]))
    # df_ip   = pd.read_excel(BytesIO(binarios["IP"]))
    # df_re   = pd.read_excel(BytesIO(binarios["RE"]))
    #
    # No se inventa aquí la lógica de negocio porque el conciliacion.py
    # proporcionado originalmente solo contaba los archivos recibidos.
    # ------------------------------------------------------------------

    return {
        "estado": "OK",
        "mensaje": (
            "Los cinco archivos fueron recibidos, decodificados y "
            "validados como Excel correctamente."
        ),
        "cantidad_archivos": len(archivos_validados),
        "archivos_recibidos": archivos_validados,
    }
