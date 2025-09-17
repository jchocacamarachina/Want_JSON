import io
import json
import os
from datetime import datetime

import pandas as pd
from flask import Flask, request, render_template_string, send_file, redirect, url_for, flash

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-change-me")

# ------------------------------
# Núcleo de transformación
# ------------------------------

def excel_to_json_bytes(excel_bytes: bytes, sheet_name: str = "Hoja 1"):
    df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=sheet_name, dtype=str)

    # --- detectar columna de existencia de forma robusta ---
    candidates = ["EXISTENTE", "EXISTENCIAS", "EXISTENCIA", "EXISTE"]
    colmap = {c.strip().lower(): c for c in df.columns}
    existence_col = None
    for cand in candidates:
        key = cand.lower()
        if key in colmap:
            existence_col = colmap[key]
            break

    total_productos = len(df)

    # Normalizar valores a 'SI' y filtrar si existe la columna
    productos_existente_si = 0
    if existence_col:
        norm = (
            df[existence_col]
              .astype(str)
              .str.strip()
              .str.upper()
              .str.replace('Í', 'I', regex=False)
        )
        mask_si = norm.eq('SI')
        productos_existente_si = int(mask_si.sum())
        df = df[mask_si]

    columnas = {
        "EXISTENCIAS": "EXISTENTES",
        "STOCK": "STOCK",
        "CODIGO": "CODIGO",
        "CATEGORIA": "CATEGORIA",
        "NOMBRE CONTIFICO": "NOMBRE CONTIFICO",
        "PRECIO": "PRECIO",
        "DESCRIPCION": "DESCRIPCION",
        "ENLACE WEB": "LINK IMAGEN",
    }

    faltantes = [c for c in columnas.keys() if c not in df.columns]
    if faltantes:
        # No detenemos si solo falta EXISTENCIAS y ya detectamos otra variante
        if not ("EXISTENCIAS" in faltantes and existence_col and existence_col != "EXISTENCIAS"):
            raise Exception(f"Faltan columnas en el Excel: {', '.join(faltantes)}")

    # Si no existe EXACTAMENTE 'EXISTENCIAS' pero detectamos otra equivalente, créala para uniformar salida
    if "EXISTENCIAS" not in df.columns and existence_col:
        df = df.assign(**{"EXISTENCIAS": df[existence_col]})

    # Ordenar columnas esperadas disponibles
    cols_out = [c for c in columnas.keys() if c in df.columns]
    df = df[cols_out].copy()

    # Convertir numéricos
    if "STOCK" in df.columns:
        df["STOCK"] = pd.to_numeric(df["STOCK"], errors='coerce')
    if "PRECIO" in df.columns:
        df["PRECIO"] = pd.to_numeric(df["PRECIO"], errors='coerce')

    # Limpiezas ligeras
    if "NOMBRE CONTIFICO" in df.columns:
        df["NOMBRE CONTIFICO"] = df["NOMBRE CONTIFICO"].astype(str).str.strip()

    registros = df.to_dict(orient="records")
    meta = {
        "total_productos": int(total_productos),
        "productos_existente_si": int(productos_existente_si),
        "sheet_name": sheet_name,
        "registros_filtrados": len(registros),
        "columna_existencias": existence_col or "(no encontrada)",
    }

    return {"meta": meta, "data": registros}


# ------------------------------
# Utilidades
# ------------------------------

def list_sheets(excel_bytes: bytes):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), read_only=True)
    return wb.sheetnames


def safe_filename(base: str):
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    name = os.path.splitext(base)[0]
    return f"{name}-{ts}.json"


# ------------------------------
# Rutas (sin previsualización)
# ------------------------------

INDEX_HTML = """
<!doctype html>
<html lang=\"es\">
<head>
  <meta charset=\"utf-8\">
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
  <title>Excel → JSON</title>
  <link href=\"https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css\" rel=\"stylesheet\">
  <style>
    body{padding-bottom:4rem}
    .dropzone{border:2px dashed #ced4da;border-radius:.75rem;padding:2rem;text-align:center}
  </style>
</head>
<body class=\"bg-light\">
<div class=\"container py-4\">
  <h1 class=\"h3 mb-3\">Excel → JSON</h1>
  <p class=\"text-muted\">Sube tu archivo .xlsx, elige la hoja (opcional) y se <strong>descargará el JSON automáticamente</strong>.</code>.</p>

  {% with messages = get_flashed_messages(with_categories=true) %}
    {% if messages %}
      {% for category, msg in messages %}
        <div class=\"alert alert-{{category}}\">{{msg|safe}}</div>
      {% endfor %}
    {% endif %}
  {% endwith %}

  <form class=\"card p-3 mb-4\" action=\"{{ url_for('convert') }}\" method=\"post\" enctype=\"multipart/form-data\">
    <div class=\"mb-3\">
      <label for=\"file\" class=\"form-label\">Archivo Excel (.xlsx)</label>
      <input class=\"form-control\" type=\"file\" id=\"file\" name=\"file\" accept=\".xlsx\" required>
    </div>
    <div class=\"mb-3\">
      <label for=\"sheet_name\" class=\"form-label\">Nombre de hoja (opcional)</label>
      <input class=\"form-control\" type=\"text\" id=\"sheet_name\" name=\"sheet_name\" placeholder=\"Hoja 1\">
      <div class=\"form-text\">Si lo dejas vacío, tomaremos la <em>primera hoja</em> encontrada.</div>
    </div>
    <button class=\"btn btn-primary\" type=\"submit\">Convertir y descargar</button>
  </form>
</div>
<script src=\"https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js\"></script>
</body>
</html>
"""

SUCCESS_HTML = """
<!doctype html>
<html lang=\"es\">
<head>
  <meta charset=\"utf-8\">
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
  <title>Descarga lista</title>
  <link href=\"https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css\" rel=\"stylesheet\">
  <script>
    document.addEventListener('DOMContentLoaded', function(){
      const link = document.getElementById('download-link');
      if (link) link.click();
    });
  </script>
</head>
<body class=\"bg-light\">
<div class=\"container py-4\">
  <div class=\"alert alert-success\">
    <strong>¡Listo!</strong> Se encontraron <strong>{{ meta.total_productos }}</strong> productos en el Excel.\n    Con <strong>{{ meta.productos_existente_si }} articulos en EXISTENCIA = SI</strong>.
  </div>

  <a id=\"download-link\" class=\"btn btn-success me-2\" href=\"{{ url_for('download_by_name', fname=out_name) }}\">Descargar JSON</a>
  <a class=\"btn btn-outline-secondary\" href=\"{{ url_for('index') }}\">Subir otro archivo</a>

  <p class=\"text-muted small mt-3\">Hoja usada: <strong>{{ meta.sheet_name }}</strong>. Registros filtrados: <strong>{{ meta.registros_filtrados }}</strong>.</p>
</div>
<script src=\"https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js\"></script>
</body>
</html>
"""


@app.get("/")
def index():
    return render_template_string(INDEX_HTML)


@app.post("/convert")
def convert():
    file = request.files.get("file")
    sheet_name = (request.form.get("sheet_name") or "").strip()

    if not file:
        flash("Sube un archivo .xlsx", "warning")
        return redirect(url_for("index"))

    try:
        excel_bytes = file.read()
        if not sheet_name:
            sheets = list_sheets(excel_bytes)
            if not sheets:
                raise Exception("No se encontraron hojas en el archivo.")
            sheet_name = sheets[0]

        result = excel_to_json_bytes(excel_bytes, sheet_name=sheet_name)

        out_name = safe_filename(file.filename or "salida.xlsx")
        out_path = os.path.join("/tmp", out_name)
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(result, f, indent=2, ensure_ascii=False)

        # En lugar de enviar directamente el archivo, mostramos una página de éxito
        # que lanza la descarga automáticamente y muestra los conteos.
        return render_template_string(
            SUCCESS_HTML,
            meta=result["meta"],
            out_name=out_name,
        )

    except Exception as e:
        flash(f"Error: {e}", "danger")
        return redirect(url_for("index"))


@app.get("/download/<path:fname>")
def download_by_name(fname):
    # Evitar path traversal: solo nombre base
    safe_name = os.path.basename(fname)
    out_path = os.path.join("/tmp", safe_name)
    if not os.path.exists(out_path):
        flash("Archivo no encontrado. Vuelve a convertir.", "warning")
        return redirect(url_for("index"))
    return send_file(out_path, as_attachment=True, download_name=safe_name, mimetype="application/json")


if __name__ == "__main__":
    # Config rápido para desarrollo
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), debug=True)

    # Config rápido para desarrollo
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), debug=True)
