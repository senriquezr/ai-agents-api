from fastapi import FastAPI
from conciliacion import ejecutar_conciliacion

app = FastAPI()

@app.get("/")
def home():
    return {
        "status": "ok"
    }

@app.get("/ping")
def ping():
    return {
        "respuesta": "pong"
    }

@app.get("/conciliacion")
def conciliacion():
    return ejecutar_conciliacion()
