from fastapi import FastAPI
import Etl.Read as Read
import Etl.descargas.Limpieza as Limpieza
import Bot.AutoEasy as AutoEasy
import Envio as Envio

app = FastAPI()

@app.get("/Bot")
def home():
    """Ejecuta solo la AutoGestión (primer paso)"""
    return AutoEasy.main()

@app.get("/Bot/actualizar")
def actualizar():
    """Ejecuta solo la actualización del informe (segundo paso)"""
    return Read.actualizar_reporte_ventas()

@app.get("/Bot/limpiar")
def limpiar():
    """Ejecuta solo la limpieza de excels antiguos (tercer paso)"""
    return Limpieza.limpiar_excels_antiguos()

@app.get("/Bot/enviar")
def enviar():
    """Ejecuta solo el envío del informe (cuarto paso)"""
    return Envio.enviar_imagenes()

@app.get("/Bot/proceso_completo")
def proceso_completo():
    """
    Ejecuta los tres pasos en el orden correcto:
    1. AutoEasy.main()
    2. Read.actualizar_reporte_ventas()
    3. Limpieza.limpiar_excels_antiguos()
    """
    resultados = {}

    # Paso 1: AutoEasy
    try:
        resultado1 = AutoEasy.main()
        resultados["1. AutoEasy"] = resultado1
        # Si falla AutoEasy, detener el proceso (opcional)
        if resultado1.get("status") == "error":
            return resultados
    except Exception as e:
        resultados["1. AutoEasy"] = f"Error: {str(e)}"
        return resultados

    # Paso 2: Read (actualizar informe)
    try:
        resultado2 = Read.actualizar_reporte_ventas()
        resultados["2. Read.actualizar_reporte_ventas"] = resultado2
    except Exception as e:
        resultados["2. Read.actualizar_reporte_ventas"] = f"Error: {str(e)}"

    # Paso 3: Limpieza
    try:
        resultado3 = Limpieza.limpiar_excels_antiguos()
        resultados["3. Limpieza"] = resultado3
    except Exception as e:
        resultados["3. Limpieza"] = f"Error: {str(e)}"
    
    # Paso 4: Envio 
    '''try:
        resultado4 = Envio.enviar_imagenes()
        resultados["4. Envio"] = resultado4
    except Exception as e:
        resultados["4. Envio"] = f"Error: {str(e)}"
    '''
    return resultados