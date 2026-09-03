from __future__ import annotations

import os, time, uuid
from pathlib import Path
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from conciliacion import ejecutar_conciliacion

app = FastAPI(title="AI Agents API - Conciliación POC", version="final-1.1")

RESULT_DIR = Path(os.getenv("RESULT_DIR", "/tmp/ai-agents-api-results"))
RESULT_DIR.mkdir(parents=True, exist_ok=True)

PUBLIC_BASE_URL = os.getenv(
    "PUBLIC_BASE_URL",
    "https://ai-agents-api-ww4v.onrender.com"
).rstrip("/")

RESULT_TTL_SECONDS = int(os.getenv("RESULT_TTL_SECONDS", "3600"))

def _limpiar_resultados_expirados():
    ahora = time.time()
    for archivo in RESULT_DIR.glob("*.xlsx"):
        try:
            if ahora - archivo.stat().st_mtime > RESULT_TTL_SECONDS:
                archivo.unlink(missing_ok=True)
        except OSError:
            pass

@app.get("/health")
def health():
    return {"ok": True, "service": "ai-agents-api", "version": "final-1.1"}

@app.post("/conciliacion")
async def conciliacion(
    bdep: UploadFile = File(...),
    sap: UploadFile = File(...),
    ep: UploadFile = File(...),
    ip: UploadFile = File(...),
    re: UploadFile = File(...),
):
    try:
        _limpiar_resultados_expirados()
        job_id = uuid.uuid4().hex
        ruta_salida = RESULT_DIR / f"resultado_conciliacion_{job_id}.xlsx"

        resumen = await ejecutar_conciliacion(
            bdep=bdep, sap=sap, ep=ep, ip=ip, re=re,
            ruta_salida=ruta_salida,
        )

        url = f"{PUBLIC_BASE_URL}/resultados/{job_id}"
        return {
            "estado": "OK",
            "mensaje": "Python recibió y abrió correctamente los cinco Excel.",
            "cantidad_archivos": resumen["cantidad_archivos"],
            "archivo_resultado_nombre": "resultado_conciliacion.xlsx",
            "archivo_resultado_url": url,
            "respuesta_usuario": f"Procesamiento completado. Descarga: {url}",
        }

    except ValueError as exc:
        print(f"[VALIDACION] {exc}", flush=True)
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        print(f"[ERROR] {type(exc).__name__}: {exc}", flush=True)
        raise HTTPException(status_code=500, detail=f"Error inesperado: {exc}") from exc

@app.get("/resultados/{job_id}")
def descargar_resultado(job_id: str):
    _limpiar_resultados_expirados()
    ruta = RESULT_DIR / f"resultado_conciliacion_{job_id}.xlsx"
    if len(job_id) != 32 or not job_id.isalnum() or not ruta.exists():
        raise HTTPException(status_code=404, detail="Resultado no encontrado o expirado.")

    return FileResponse(
        path=ruta,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="resultado_conciliacion.xlsx",
    )
