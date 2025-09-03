import io
import json
import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import JSONResponse

app = FastAPI()

def excel_to_json_bytes(excel_bytes: bytes, sheet_name: str = "Hoja 1"):
    df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=sheet_name, dtype=str)

    total_productos = len(df)
    total_si = (df.get('EXISTENTE', pd.Series(dtype=str)) == 'SI').sum()

    # Filtrar EXISTENTE=SI
    if 'EXISTENTE' in df.columns:
        df = df[df['EXISTENTE'] == 'SI']

    columnas = {
        "EXISTENTE": "EXISTENTE",
        "STOCK": "STOCK",
        "CODIGO": "CODIGO",
        "CATEGORIA": "CATEGORIA",
        "NOMBRE CONTIFICO": "NOMBRE CONTIFICO",
        "PRECIO": "PRECIO",
        "DESCRIPCION": "DESCRIPCION",
        "ENLACE WEB": "LINK IMAGEN",
    }

    # Validar columnas requeridas
    faltantes = [c for c in columnas.keys() if c not in df.columns]
    if faltantes:
        raise HTTPException(status_code=400, detail=f"Faltan columnas en el Excel: {', '.join(faltantes)}")

    df = df[list(columnas.keys())].copy()

    # Numéricos
    df["STOCK"] = pd.to_numeric(df["STOCK"], errors='coerce')
    df["PRECIO"] = pd.to_numeric(df["PRECIO"], errors='coerce')

    registros = df.to_dict(orient="records")
    meta = {
        "total_productos": int(total_productos),
        "productos_existente_si": int(total_si),
        "sheet_name": sheet_name,
        "registros_filtrados": len(registros),
    }
    return registros, meta

@app.get("/")
def root():
    return {"status": "ok", "usage": "POST /convert with file=excel and optional sheet_name"}

@app.post("/convert")
async def convert(file: UploadFile = File(...), sheet_name: str = Form("Hoja 1")):
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Sube un Excel .xlsx o .xls")

    content = await file.read()
    try:
        registros, meta = excel_to_json_bytes(content, sheet_name=sheet_name)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

    return JSONResponse({"meta": meta, "data": registros})
