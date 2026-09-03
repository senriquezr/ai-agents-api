from __future__ import annotations

import os
import time
import uuid
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel

from conciliacion import ejecutar_conciliacion


class SolicitudConciliacion(BaseModel):
    attachmentsJson: str


app = FastAPI(title="AI Agents API", version="poc-final-2.0")

RESULT_DIR = Path(os.getenv("RESULT_DIR", "/tmp/ai-agents-api-results"))
RESULT_DIR.mkdir(parents=True, exist_ok=True)
PUBLIC_BASE_URL = os.getenv(
    "PUBLIC_BASE_URL",
    "https://ai-agents-api-ww4v.onrender.com",
).rstrip("/")
TTL = int(os.getenv("RESULT_TTL_SECONDS", "3600"))


def limpiar():
    ahora = time.time()
    for f in RESULT_DIR.glob("*.xlsx"):
        try:
            if ahora - f.stat().st_mtime > TTL:
                f.unlink(missing_ok=True)
        except OSError:
            pass


@app.get("/health")
def health():
    return {"ok": True, "service": "ai-agents-api", "version": "poc-final-2.0"}


@app.post("/conciliacion")
def conciliacion(solicitud: SolicitudConciliacion):
    try:
        limpiar()
        job_id = uuid.uuid4().hex
        salida = RESULT_DIR / f"resultado_conciliacion_{job_id}.xlsx"

        resumen = ejecutar_conciliacion(
            attachments_json=solicitud.attachmentsJson,
            ruta_salida=salida,
        )

        return {
            "estado": "OK",
            "mensaje": "Python recibió los 5 adjuntos y generó un Excel.",
            "cantidad_archivos": resumen["cantidad_archivos"],
            "archivo_resultado_nombre": "resultado_conciliacion.xlsx",
            "archivo_resultado_url": f"{PUBLIC_BASE_URL}/resultados/{job_id}",
        }
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Error inesperado: {exc}") from exc


@app.get("/resultados/{job_id}")
def descargar_resultado(job_id: str):
    limpiar()
    if len(job_id) != 32 or not job_id.isalnum():
        raise HTTPException(status_code=404, detail="Resultado no encontrado.")

    ruta = RESULT_DIR / f"resultado_conciliacion_{job_id}.xlsx"
    if not ruta.exists():
        raise HTTPException(status_code=404, detail="Resultado no encontrado o expirado.")

    return FileResponse(
        ruta,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="resultado_conciliacion.xlsx",
    )
