from __future__ import annotations

import base64
import json
from io import BytesIO
from pathlib import Path

import pandas as pd


def get_ci(d, *names):
    m = {str(k).lower(): v for k, v in d.items()}
    for n in names:
        if n.lower() in m:
            return m[n.lower()]
    return None


def extraer_registros(x):
    if isinstance(x, list):
        return [v for v in x if isinstance(v, dict)]
    if isinstance(x, dict):
        for k in ("value", "attachments", "items", "files"):
            v = get_ci(x, k)
            if isinstance(v, list):
                return [i for i in v if isinstance(i, dict)]
        if get_ci(x, "name", "filename") is not None:
            return [x]
        for v in x.values():
            r = extraer_registros(v)
            if r:
                return r
    return []


def decode_content(v, nombre):
    if isinstance(v, dict):
        v = get_ci(v, "content", "contentbytes", "value", "$content")
    if v is None:
        raise ValueError(f"No se encontró contenido binario para {nombre}.")
    s = str(v).strip()
    if s.startswith("data:") and "," in s:
        s = s.split(",", 1)[1]
    try:
        return base64.b64decode(s, validate=True)
    except Exception as exc:
        raise ValueError(f"El contenido de {nombre} no llegó como Base64 válido.") from exc


def tipo_por_nombre(nombre):
    u = nombre.upper()
    for t in ("BDEP", "SAP", "EP", "IP", "RE"):
        if t in u:
            return t
    return None


def validar_excel(nombre, data, tipo):
    if not nombre.lower().endswith((".xlsx", ".xlsm", ".xls")):
        raise ValueError(f"{tipo}: {nombre} no es un Excel permitido.")
    try:
        libro = pd.ExcelFile(BytesIO(data))
        hoja = libro.sheet_names[0]
        muestra = pd.read_excel(BytesIO(data), sheet_name=hoja, nrows=20)
    except Exception as exc:
        raise ValueError(f"{tipo}: no se pudo abrir {nombre}: {exc}") from exc
    return {
        "tipo": tipo,
        "archivo": nombre,
        "primera_hoja": hoja,
        "numero_hojas": len(libro.sheet_names),
        "columnas_detectadas": len(muestra.columns),
        "filas_muestra": len(muestra),
    }


def ejecutar_conciliacion(attachments_json: str, ruta_salida: str | Path) -> dict:
    try:
        payload = json.loads(attachments_json)
    except Exception as exc:
        raise ValueError("attachmentsJson no contiene JSON válido.") from exc

    regs = extraer_registros(payload)
    if len(regs) != 5:
        raise ValueError(f"Se esperaban 5 adjuntos y se recibieron {len(regs)}.")

    encontrados = {}
    for r in regs:
        nombre = get_ci(r, "name", "filename", "fileName", "displayName")
        if not nombre:
            raise ValueError("Un adjunto no tiene nombre reconocible.")
        nombre = str(nombre)
        tipo = tipo_por_nombre(nombre)
        if not tipo:
            raise ValueError(f"No se pudo identificar {nombre} como BDEP/SAP/EP/IP/RE.")
        if tipo in encontrados:
            raise ValueError(f"Hay más de un archivo para {tipo}.")

        contenido = get_ci(r, "content", "contentBytes", "fileContent", "data", "$content")
        data = decode_content(contenido, nombre)
        encontrados[tipo] = validar_excel(nombre, data, tipo)

    faltan = [t for t in ("BDEP", "SAP", "EP", "IP", "RE") if t not in encontrados]
    if faltan:
        raise ValueError("Faltan: " + ", ".join(faltan))

    orden = ["BDEP", "SAP", "EP", "IP", "RE"]
    df = pd.DataFrame([encontrados[t] for t in orden])
    resumen = pd.DataFrame([{
        "estado": "OK",
        "cantidad_archivos": 5,
        "mensaje": "Python recibió y abrió correctamente los cinco Excel."
    }])

    ruta_salida = Path(ruta_salida)
    with pd.ExcelWriter(ruta_salida, engine="xlsxwriter") as writer:
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        df.to_excel(writer, sheet_name="Validacion", index=False)

    return {"cantidad_archivos": 5}
