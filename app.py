from fastapi import FastAPI, HTTPException
from pydantic import BaseModel

from conciliacion import ejecutar_conciliacion


class ArchivoEntrada(BaseModel):
    name: str
    contentBytes: str


class SolicitudConciliacion(BaseModel):
    archivos: list[ArchivoEntrada]


app = FastAPI(
    title="AI Agents API",
    version="2.0.0",
    description=(
        "API central para ejecutar scripts utilizados "
        "por agentes de Copilot Studio."
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
    }


@app.get("/ping")
def ping():
    return {
        "respuesta": "pong",
    }


@app.post("/conciliacion")
def conciliacion(
    solicitud: SolicitudConciliacion,
):
    try:
        return ejecutar_conciliacion(
            archivos=solicitud.archivos,
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
