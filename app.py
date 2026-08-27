from fastapi import FastAPI

app = FastAPI()

@app.get("/")
def home():
    return {
        "status": "ok",
        "mensaje": "API funcionando"
    }

@app.get("/ping")
def ping():
    return {
        "respuesta": "pong"
    }
