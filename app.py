from fastapi import FastAPI, File, HTTPException, UploadFile

from conciliacion import ejecutar_conciliacion


app = FastAPI(
    title="AI Agents API",
    version="1.0.0",
    description=(
        "API central para ejecutar scripts utilizados "
        "por agentes de Copilot Studio."
    ),
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
    archivo_bdep: UploadFile = File(...),
    archivo_sap: UploadFile = File(...),
    archivo_ep: UploadFile = File(...),
    archivo_ip: UploadFile = File(...),
    archivo_re: UploadFile = File(...),
):
    try:
        return ejecutar_conciliacion(
            archivo_bdep=archivo_bdep,
            archivo_sap=archivo_sap,
            archivo_ep=archivo_ep,
            archivo_ip=archivo_ip,
            archivo_re=archivo_re,
        )

    except ValueError as error:
        raise HTTPException(
            status_code=400,
            detail=str(error),
        ) from error

    except Exception as error:
        raise HTTPException(
            status_code=500,
            detail=(
                "Error inesperado al procesar la conciliación: "
                f"{error}"
            ),
        ) from error
