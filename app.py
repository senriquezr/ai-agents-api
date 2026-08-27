from fastapi import FastAPI

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

@app.get("/saludo")
def saludo():
    return {
        "mensaje": "Hola Stephan, tu API está funcionando"
    }

@app.get("/sumar")
def sumar(a: int, b: int):
    return {
        "resultado": a + b
    }
