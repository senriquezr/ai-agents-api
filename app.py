from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field

from conciliacion import ejecutar_conciliacion


class ArchivoEntrada(BaseModel):
    name: str = Field(..., description="Nombre original del archivo, incluyendo extensión.")
    contentBytes: str = Field(..., description="Contenido del archivo codificado en Base64.")


class SolicitudConciliacion(BaseModel):
    bdep: ArchivoEntrada
    sap: ArchivoEntrada
    ep: ArchivoEntrada
    ip: ArchivoEntrada
    re: ArchivoEntrada


app = FastAPI(
    title="AI Agents API",
    version="2.0.0",
    description=(
        "API central para ejecutar conciliaciones desde agentes de Copilot Studio. "
        "Recibe cinco archivos Excel codificados en Base64."
    ),
    servers=[
        {
            "url": "https://ai-agents-api-ww4v.onrender.com"
        }
    ],
)


@app.get("/")
def home():
    return {
        "status": "ok",
        "mensaje": "API funcionando",
        "version": "2.0.0",
    }


@app.get("/ping")
def ping():
    return {
        "respuesta": "pong",
    }


@app.get("/health")
def health():
    return {
        "ok": True,
        "service": "ai-agents-api",
        "version": "2.0.0",
    }


@app.post("/conciliacion")
def conciliacion(solicitud: SolicitudConciliacion):
    try:
        return ejecutar_conciliacion(
            bdep=solicitud.bdep,
            sap=solicitud.sap,
            ep=solicitud.ep,
            ip=solicitud.ip,
            re=solicitud.re,
        )

    except ValueError as error:
        raise HTTPException(
            status_code=400,
            detail=str(error),
        ) from error

    except Exception as error:
        raise HTTPException(
            status_code=500,
            detail=f"Error inesperado: {error}",
        ) from error
