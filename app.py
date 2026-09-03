from __future__ import annotations

import os
import time
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from conciliacion import ejecutar_conciliacion


app = FastAPI(
    title="AI Agents API - Conciliación POC",
    version="final-1.0",
    description=(
        "Recibe cinco Excel (BDEP, SAP, EP, IP y RE) en una sola llamada, "
        "ejecuta Python y genera un Excel descargable."
    ),
)

RESULT_DIR = Path(os.getenv("RESULT_DIR", "/tmp/ai-agents-api-results"))
RESULT_DIR.mkdir(parents=True, exist_ok=True)

PUBLIC_BASE_URL = os.getenv(
    "PUBLIC_BASE_URL",
    "https://ai-agents-api-ww4v.onrender.com",
).rstrip("/")

RESULT_TTL_SECONDS = int(os.getenv("RESULT_TTL_SECONDS", "3600"))


def _limpiar_resultados_expirados() -> None:
    ahora = time.time()
    for archivo in RESULT_DIR.glob("*.xlsx"):
        try:
            if ahora - archivo.stat().st_mtime > RESULT_TTL_SECONDS:
                archivo.unlink(missing_ok=True)
        except OSError:
            pass


@app.get("/")
def home():
    return {
        "status": "ok",
        "service": "ai-agents-api",
        "version": "final-1.0",
    }


@app.get("/health")
def health():
    return {
        "ok": True,
        "service": "ai-agents-api",
        "version": "final-1.0",
    }


@app.post("/conciliacion")
async def conciliacion(
    bdep: UploadFile = File(..., description="Archivo BDEP.xlsx"),
    sap: UploadFile = File(..., description="Archivo SAP.xlsx"),
    ep: UploadFile = File(..., description="Archivo EP.xlsx"),
    ip: UploadFile = File(..., description="Archivo IP.xlsx"),
    re: UploadFile = File(..., description="Archivo RE.xlsx"),
):
    try:
        _limpiar_resultados_expirados()

        job_id = uuid.uuid4().hex
        ruta_salida = RESULT_DIR / f"resultado_conciliacion_{job_id}.xlsx"

        resumen = await ejecutar_conciliacion(
            bdep=bdep,
            sap=sap,
            ep=ep,
            ip=ip,
            re=re,
            ruta_salida=ruta_salida,
        )

        download_url = f"{PUBLIC_BASE_URL}/resultados/{job_id}"

        return {
            "estado": "OK",
            "mensaje": (
                "Python recibió y abrió correctamente los cinco Excel "
                "y generó el archivo de resultado."
            ),
            "cantidad_archivos": resumen["cantidad_archivos"],
            "archivo_resultado_nombre": "resultado_conciliacion.xlsx",
            "archivo_resultado_url": download_url,
            "respuesta_usuario": (
                "Procesamiento completado correctamente. "
                f"Descarga el resultado aquí: {download_url}"
            ),
        }

    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"Error inesperado: {exc}",
        ) from exc


@app.get("/resultados/{job_id}")
def descargar_resultado(job_id: str):
    _limpiar_resultados_expirados()

    if len(job_id) != 32 or not job_id.isalnum():
        raise HTTPException(status_code=404, detail="Resultado no encontrado.")

    ruta = RESULT_DIR / f"resultado_conciliacion_{job_id}.xlsx"

    if not ruta.exists():
        raise HTTPException(
            status_code=404,
            detail="El resultado no existe o ya expiró.",
        )

    return FileResponse(
        path=ruta,
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        filename="resultado_conciliacion.xlsx",
    )
