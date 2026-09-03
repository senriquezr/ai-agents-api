from __future__ import annotations

import os
import shutil
import time
import uuid
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Response, UploadFile
from fastapi.responses import FileResponse

from conciliacion import ejecutar_conciliacion

app = FastAPI(
    title="AI Agents API - Conciliación Contable",
    version="2.0.0",
    description=(
        "Recibe cinco archivos Excel desde Copilot Studio, ejecuta la "
        "conciliación contable completa y genera un Excel descargable."
    ),
)

RESULT_DIR = Path(os.getenv("RESULT_DIR", "/tmp/ai-agents-api-results"))
RESULT_DIR.mkdir(parents=True, exist_ok=True)
PUBLIC_BASE_URL = os.getenv(
    "PUBLIC_BASE_URL",
    "https://ai-agents-api-ww4v.onrender.com",
).rstrip("/")
RESULT_TTL_SECONDS = int(os.getenv("RESULT_TTL_SECONDS", "3600"))
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _limpiar_resultados_expirados() -> None:
    ahora = time.time()
    for archivo in RESULT_DIR.glob("*.xlsx"):
        try:
            if ahora - archivo.stat().st_mtime > RESULT_TTL_SECONDS:
                archivo.unlink(missing_ok=True)
        except OSError:
            pass


def _buscar_resultado(job_id: str) -> Path:
    if len(job_id) != 32 or not job_id.isalnum():
        raise HTTPException(status_code=404, detail="Resultado no encontrado.")

    candidatos = list(RESULT_DIR.glob(f"{job_id}__*.xlsx"))
    if not candidatos:
        raise HTTPException(
            status_code=404,
            detail="El resultado no existe o ya expiró.",
        )
    return candidatos[0]


@app.get("/")
def home():
    return {
        "status": "ok",
        "service": "ai-agents-api",
        "version": "2.0.0",
    }


@app.get("/health")
def health():
    return {
        "ok": True,
        "service": "ai-agents-api",
        "version": "2.0.0",
    }


@app.post("/conciliacion")
async def conciliacion(
    bdep: UploadFile = File(..., description="Contenido del archivo BDEP."),
    sap: UploadFile = File(..., description="Contenido del archivo SAP."),
    ep: UploadFile = File(..., description="Contenido del archivo EP."),
    ip: UploadFile = File(..., description="Contenido del archivo IP."),
    re: UploadFile = File(..., description="Contenido del archivo RE."),
    bdep_original_name: str | None = Form(None),
    sap_original_name: str | None = Form(None),
    ep_original_name: str | None = Form(None),
    ip_original_name: str | None = Form(None),
    re_original_name: str | None = Form(None),
):
    try:
        _limpiar_resultados_expirados()

        resultado = await ejecutar_conciliacion(
            bdep=bdep,
            sap=sap,
            ep=ep,
            ip=ip,
            re=re,
            nombres_originales={
                "bdep": bdep_original_name,
                "sap": sap_original_name,
                "ep": ep_original_name,
                "ip": ip_original_name,
                "re": re_original_name,
            },
        )

        job_id = uuid.uuid4().hex
        nombre_resultado = resultado["archivo"].name
        destino = RESULT_DIR / f"{job_id}__{nombre_resultado}"
        shutil.copy2(resultado["archivo"], destino)

        # Limpiar la copia temporal creada por conciliacion.py.
        try:
            origen_temporal = resultado["archivo"]
            origen_temporal.unlink(missing_ok=True)
            origen_temporal.parent.rmdir()
        except OSError:
            pass

        url = f"{PUBLIC_BASE_URL}/resultados/{job_id}"
        respuesta_usuario = (
            "Procesamiento completado correctamente.\n\n"
            f"[{nombre_resultado}]({url})"
        )

        return {
            "estado": "OK",
            "mensaje": "Conciliación contable completada correctamente.",
            "periodo": resultado["periodo"],
            "cantidad_archivos": 5,
            "cuentas_conciliadas": resultado["cuentas_conciliadas"],
            "cuentas_ok": resultado["cuentas_ok"],
            "cuentas_con_diferencia": resultado["cuentas_con_diferencia"],
            "archivo_resultado_nombre": nombre_resultado,
            "archivo_resultado_url": url,
            "respuesta_usuario": respuesta_usuario,
        }

    except ValueError as exc:
        print(f"[VALIDACION] {exc}", flush=True)
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        print(f"[ERROR] {type(exc).__name__}: {exc}", flush=True)
        raise HTTPException(
            status_code=500,
            detail=f"Error inesperado: {exc}",
        ) from exc


@app.get("/resultados/{job_id}")
def descargar_resultado(job_id: str):
    _limpiar_resultados_expirados()
    ruta = _buscar_resultado(job_id)
    nombre_visible = ruta.name.split("__", 1)[1]
    return FileResponse(
        path=ruta,
        media_type=XLSX_MIME,
        filename=nombre_visible,
    )


@app.head("/resultados/{job_id}")
def comprobar_resultado(job_id: str):
    _limpiar_resultados_expirados()
    ruta = _buscar_resultado(job_id)
    nombre_visible = ruta.name.split("__", 1)[1]
    return Response(
        status_code=200,
        headers={
            "Content-Type": XLSX_MIME,
            "Content-Length": str(ruta.stat().st_size),
            "Content-Disposition": f'attachment; filename="{nombre_visible}"',
        },
    )
