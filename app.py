from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field

from conciliacion import ejecutar_conciliacion


class SolicitudConciliacion(BaseModel):
    bdep_name: str = Field(..., description="Nombre original del archivo BDEP.")
    bdep_contentBytes: str = Field(..., description="Contenido Base64 del archivo BDEP.")
    sap_name: str = Field(..., description="Nombre original del archivo SAP.")
    sap_contentBytes: str = Field(..., description="Contenido Base64 del archivo SAP.")
    ep_name: str = Field(..., description="Nombre original del archivo EP.")
    ep_contentBytes: str = Field(..., description="Contenido Base64 del archivo EP.")
    ip_name: str = Field(..., description="Nombre original del archivo IP.")
    ip_contentBytes: str = Field(..., description="Contenido Base64 del archivo IP.")
    re_name: str = Field(..., description="Nombre original del archivo RE.")
    re_contentBytes: str = Field(..., description="Contenido Base64 del archivo RE.")


app = FastAPI(
    title="AI Agents API",
    version="3.0.0",
    description="API para recibir cinco archivos Excel desde Copilot Studio y ejecutar procesamiento Python.",
    servers=[{"url": "https://ai-agents-api-ww4v.onrender.com"}],
)


@app.get("/")
def home():
    return {"status": "ok", "mensaje": "API funcionando", "version": "3.0.0"}


@app.get("/ping")
def ping():
    return {"respuesta": "pong"}


@app.get("/health")
def health():
    return {"ok": True, "service": "ai-agents-api", "version": "3.0.0"}


@app.post("/conciliacion")
def conciliacion(solicitud: SolicitudConciliacion):
    try:
        return ejecutar_conciliacion(
            bdep_name=solicitud.bdep_name,
            bdep_contentBytes=solicitud.bdep_contentBytes,
            sap_name=solicitud.sap_name,
            sap_contentBytes=solicitud.sap_contentBytes,
            ep_name=solicitud.ep_name,
            ep_contentBytes=solicitud.ep_contentBytes,
            ip_name=solicitud.ip_name,
            ip_contentBytes=solicitud.ip_contentBytes,
            re_name=solicitud.re_name,
            re_contentBytes=solicitud.re_contentBytes,
        )
    except ValueError as error:
        raise HTTPException(status_code=400, detail=str(error)) from error
    except Exception as error:
        raise HTTPException(status_code=500, detail=f"Error inesperado: {error}") from error
