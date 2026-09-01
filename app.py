from fastapi import FastAPI, File, HTTPException, UploadFile

from conciliacion import ejecutar_conciliacion

app = FastAPI(
    title="AI Agents API",
    version="1.0.0",
    description=(
        "API central para ejecutar scripts utilizados "
        "por agentes de Copilot Studio."
    ),
    servers=[
        {
            "url": "https://ai-agents-api-ww4v.onrender.com"
        }
    ]
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
    archivos: list[UploadFile] = File(...)
):
    try:
        return ejecutar_conciliacion(
            archivos=archivos
        )

    except ValueError as error:
        raise HTTPException(
            status_code=400,
            detail=str(error),
        )

    except Exception as error:
        raise HTTPException(
            status_code=500,
            detail=(
                f"Error inesperado: {error}"
            ),
        )
