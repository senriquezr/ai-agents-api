from __future__ import annotations

from io import BytesIO
from pathlib import Path
import openpyxl
import xlsxwriter
from fastapi import UploadFile

async def _validar_xlsx(archivo: UploadFile, tipo: str) -> dict:
    contenido = await archivo.read()

    if not contenido:
        raise ValueError(f"{tipo}: Copilot envió el parámetro File, pero llegó vacío.")

    nombre_recibido = archivo.filename or ""
    content_type = archivo.content_type or ""

    print(
        f"[ARCHIVO] tipo={tipo} filename={nombre_recibido!r} "
        f"content_type={content_type!r} bytes={len(contenido)}",
        flush=True,
    )

    try:
        libro = openpyxl.load_workbook(
            BytesIO(contenido),
            read_only=True,
            data_only=True,
        )
        hojas = libro.sheetnames
        if not hojas:
            raise ValueError("El libro no contiene hojas.")

        primera_hoja = hojas[0]
        ws = libro[primera_hoja]
        filas_muestra = 0
        columnas_detectadas = 0
        for fila in ws.iter_rows(max_row=20, values_only=True):
            filas_muestra += 1
            columnas_detectadas = max(columnas_detectadas, len(fila))
        libro.close()

    except Exception as exc:
        raise ValueError(
            f"{tipo}: llegaron {len(contenido)} bytes, pero Python no pudo "
            f"abrirlos como XLSX: {exc}"
        ) from exc

    nombre_resultado = (
        nombre_recibido
        if nombre_recibido.lower().endswith(".xlsx")
        else f"{tipo}.xlsx"
    )

    return {
        "tipo": tipo,
        "archivo": nombre_resultado,
        "nombre_recibido_connector": nombre_recibido or "(sin nombre)",
        "content_type": content_type or "(sin content-type)",
        "primera_hoja": primera_hoja,
        "numero_hojas": len(hojas),
        "columnas_detectadas": columnas_detectadas,
        "filas_muestra": filas_muestra,
        "tamano_bytes": len(contenido),
    }

async def ejecutar_conciliacion(
    bdep: UploadFile, sap: UploadFile, ep: UploadFile,
    ip: UploadFile, re: UploadFile, ruta_salida: str | Path
) -> dict:

    resultados = [
        await _validar_xlsx(bdep, "BDEP"),
        await _validar_xlsx(sap, "SAP"),
        await _validar_xlsx(ep, "EP"),
        await _validar_xlsx(ip, "IP"),
        await _validar_xlsx(re, "RE"),
    ]

    ruta_salida = Path(ruta_salida)
    ruta_salida.parent.mkdir(parents=True, exist_ok=True)

    workbook = xlsxwriter.Workbook(ruta_salida)
    try:
        ws1 = workbook.add_worksheet("Resumen")
        ws2 = workbook.add_worksheet("Validacion")
        header = workbook.add_format({"bold": True})

        ws1.write_row(0, 0, ["Campo", "Valor"], header)
        ws1.write_row(1, 0, ["estado", "OK"])
        ws1.write_row(2, 0, ["cantidad_archivos", 5])
        ws1.write_row(3, 0, ["mensaje", "Python recibió y abrió los cinco archivos."])

        columnas = [
            "tipo","archivo","nombre_recibido_connector","content_type",
            "primera_hoja","numero_hojas","columnas_detectadas",
            "filas_muestra","tamano_bytes"
        ]
        ws2.write_row(0, 0, columnas, header)

        for i, r in enumerate(resultados, start=1):
            ws2.write_row(i, 0, [r[c] for c in columnas])

        ws2.freeze_panes(1, 0)
    finally:
        workbook.close()

    return {"cantidad_archivos": 5}
