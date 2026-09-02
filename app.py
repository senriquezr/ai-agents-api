from __future__ import annotations

import os
import time
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse

from conciliacion import ejecutar_conciliacion


app = FastAPI(
    title="AI Agents API",
    version="1.1.0-test",
    description=(
        "Prueba end-to-end: recibe 5 Excel desde Copilot Studio, "
        "ejecuta Python y genera un Excel de resultado."
    ),
)

RESULT_DIR = Path(os.getenv("RESULT_DIR", "/tmp/ai-agents-api-results"))
RESULT_DIR.mkdir(parents=True, exist_ok=True)

RESULT_TTL_SECONDS = int(os.getenv("RESULT_TTL_SECONDS", "3600"))
PUBLIC_BASE_URL = os.getenv(
    "PUBLIC_BASE_URL",
    "https://ai-agents-api-ww4v.onrender.com",
).rstrip("/")


def _limpiar_resultados_expirados() -> None:
    ahora = time.time()
    for archivo in RESULT_DIR.glob("*.xlsx"):
        try:
            if ahora - archivo.stat().st_mtime > RESULT_TTL_SECONDS:
                archivo.unlink(missing_ok=True)
        except OSError:
            pass


@app.get("/health")
def health():
    return {
        "ok": True,
        "service": "ai-agents-api",
        "version": "1.1.0-test",
    }


@app.post("/conciliacion")
async def conciliacion(
    request: Request,
    bdep: UploadFile = File(..., description="Archivo Excel BDEP"),
    sap: UploadFile = File(..., description="Archivo Excel SAP"),
    ep: UploadFile = File(..., description="Archivo Excel EP"),
    ip: UploadFile = File(..., description="Archivo Excel IP"),
    re: UploadFile = File(..., description="Archivo Excel RE"),
):
    try:
        _limpiar_resultados_expirados()

        job_id = uuid.uuid4().hex
        nombre_salida = f"resultado_conciliacion_{job_id}.xlsx"
        ruta_salida = RESULT_DIR / nombre_salida

        resumen = await ejecutar_conciliacion(
            bdep=bdep,
            sap=sap,
            ep=ep,
            ip=ip,
            re=re,
            ruta_salida=ruta_salida,
        )

        base_url = PUBLIC_BASE_URL or str(request.base_url).rstrip("/")
        download_url = f"{base_url}/resultados/{job_id}"

        return {
            "estado": "OK",
            "mensaje": (
                "Python recibió los 5 archivos y generó un Excel de resultado."
            ),
            "cantidad_archivos": 5,
            "archivo_resultado_nombre": "resultado_conciliacion.xlsx",
            "archivo_resultado_url": download_url,
            "resumen": resumen,
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
        raise HTTPException(
            status_code=404,
            detail="Resultado no encontrado.",
        )

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
