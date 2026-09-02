from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd
from fastapi import UploadFile


EXTENSIONES = (".xlsx", ".xlsm", ".xls")


async def _validar(
    archivo: UploadFile,
    tipo: str,
) -> dict:
    nombre = archivo.filename or ""

    if not nombre:
        raise ValueError(f"{tipo}: el archivo no tiene nombre.")

    if not nombre.lower().endswith(EXTENSIONES):
        raise ValueError(
            f"{tipo}: '{nombre}' no es un Excel permitido."
        )

    contenido = await archivo.read()

    if not contenido:
        raise ValueError(
            f"{tipo}: '{nombre}' está vacío."
        )

    try:
        libro = pd.ExcelFile(BytesIO(contenido))
        hojas = list(libro.sheet_names)

        if not hojas:
            raise ValueError("El libro no contiene hojas.")

        primera_hoja = hojas[0]

        muestra = pd.read_excel(
            BytesIO(contenido),
            sheet_name=primera_hoja,
            nrows=20,
        )

    except Exception as exc:
        raise ValueError(
            f"{tipo}: no se pudo abrir '{nombre}' como Excel: {exc}"
        ) from exc

    return {
        "tipo": tipo,
        "archivo": nombre,
        "primera_hoja": primera_hoja,
        "numero_hojas": len(hojas),
        "columnas_detectadas": len(muestra.columns),
        "filas_muestra": len(muestra),
    }


async def ejecutar_conciliacion(
    bdep: UploadFile,
    sap: UploadFile,
    ep: UploadFile,
    ip: UploadFile,
    re: UploadFile,
    ruta_salida: str | Path,
) -> dict:
    """
    PRUEBA MÍNIMA.

    Recibe los cinco Excel, los abre con pandas y genera un nuevo
    Excel llamado resultado_conciliacion.xlsx.

    Aquí, más adelante, se reemplaza esta validación por el código
    real de conciliación de más de 1000 líneas.
    """

    resultados = [
        await _validar(bdep, "BDEP"),
        await _validar(sap, "SAP"),
        await _validar(ep, "EP"),
        await _validar(ip, "IP"),
        await _validar(re, "RE"),
    ]

    df_resultado = pd.DataFrame(resultados)

    df_resumen = pd.DataFrame(
        [
            {
                "estado": "OK",
                "cantidad_archivos": 5,
                "mensaje": (
                    "Python recibió y abrió correctamente "
                    "los cinco archivos Excel."
                ),
            }
        ]
    )

    ruta_salida = Path(ruta_salida)
    ruta_salida.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(
        ruta_salida,
        engine="xlsxwriter",
    ) as writer:
        df_resumen.to_excel(
            writer,
            sheet_name="Resumen",
            index=False,
        )
        df_resultado.to_excel(
            writer,
            sheet_name="Validacion",
            index=False,
        )

        for nombre_hoja, dataframe in {
            "Resumen": df_resumen,
            "Validacion": df_resultado,
        }.items():
            worksheet = writer.sheets[nombre_hoja]
            worksheet.freeze_panes(1, 0)

            for indice, columna in enumerate(dataframe.columns):
                ancho = max(
                    len(str(columna)) + 2,
                    16,
                )
                worksheet.set_column(
                    indice,
                    indice,
                    min(ancho, 35),
                )

    return {
        "archivos_validados": 5,
        "bdep": resultados[0]["archivo"],
        "sap": resultados[1]["archivo"],
        "ep": resultados[2]["archivo"],
        "ip": resultados[3]["archivo"],
        "re": resultados[4]["archivo"],
    }
