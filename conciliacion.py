from __future__ import annotations

import base64
import binascii
from io import BytesIO

import pandas as pd


EXTENSIONES_PERMITIDAS = (".xlsx", ".xlsm", ".xls")


def _decodificar_excel(nombre: str, contenido: str, etiqueta: str) -> tuple[str, bytes]:
    if not nombre:
        raise ValueError(f"{etiqueta}: falta el nombre del archivo.")
    if not nombre.lower().endswith(EXTENSIONES_PERMITIDAS):
        raise ValueError(f"{etiqueta}: '{nombre}' no es un Excel permitido (.xlsx, .xlsm o .xls).")
    if not contenido:
        raise ValueError(f"{etiqueta}: '{nombre}' no contiene contentBytes.")

    contenido = str(contenido).strip()
    if contenido.startswith("data:") and "," in contenido:
        contenido = contenido.split(",", 1)[1]

    try:
        datos = base64.b64decode(contenido, validate=True)
    except (binascii.Error, ValueError) as error:
        raise ValueError(f"{etiqueta}: el contenido de '{nombre}' no es Base64 válido.") from error

    if not datos:
        raise ValueError(f"{etiqueta}: '{nombre}' está vacío.")

    return nombre, datos


def _validar_excel(datos: bytes, etiqueta: str, nombre: str) -> dict:
    try:
        buffer = BytesIO(datos)
        excel = pd.ExcelFile(buffer)
        hojas = list(excel.sheet_names)
        if not hojas:
            raise ValueError("El libro no contiene hojas.")

        buffer.seek(0)
        muestra = pd.read_excel(buffer, sheet_name=hojas[0], nrows=5)

        return {
            "tipo": etiqueta,
            "name": nombre,
            "bytes": len(datos),
            "primera_hoja": hojas[0],
            "numero_hojas": len(hojas),
            "numero_columnas_detectadas": len(muestra.columns),
        }
    except Exception as error:
        raise ValueError(f"{etiqueta}: no se pudo leer '{nombre}' como Excel: {error}") from error


def ejecutar_conciliacion(
    bdep_name: str,
    bdep_contentBytes: str,
    sap_name: str,
    sap_contentBytes: str,
    ep_name: str,
    ep_contentBytes: str,
    ip_name: str,
    ip_contentBytes: str,
    re_name: str,
    re_contentBytes: str,
) -> dict:
    entradas = {
        "BDEP": (bdep_name, bdep_contentBytes),
        "SAP": (sap_name, sap_contentBytes),
        "EP": (ep_name, ep_contentBytes),
        "IP": (ip_name, ip_contentBytes),
        "RE": (re_name, re_contentBytes),
    }

    validados = {}
    for etiqueta, (nombre, contenido) in entradas.items():
        nombre_ok, datos = _decodificar_excel(nombre, contenido, etiqueta)
        validados[etiqueta] = _validar_excel(datos, etiqueta, nombre_ok)

    return {
        "estado": "OK",
        "mensaje": "Los cinco archivos fueron recibidos y validados correctamente.",
        "cantidad_archivos": 5,
        "bdep_recibido": validados["BDEP"]["name"],
        "sap_recibido": validados["SAP"]["name"],
        "ep_recibido": validados["EP"]["name"],
        "ip_recibido": validados["IP"]["name"],
        "re_recibido": validados["RE"]["name"],
    }
