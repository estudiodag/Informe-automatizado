"""
app.py
Servidor web Flask para el generador de informes de tasacion.

Endpoints:
  GET  /            -> sirve la pagina web (index.html)
  POST /procesar    -> recibe plantilla + peritacion + cotizacion (opcional),
                       devuelve el informe completo
"""

import os
from io import BytesIO
from flask import Flask, request, send_file, render_template, jsonify

from procesador import procesar

app = Flask(__name__)

# Limite de tamano de subida: 10 MB
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024


# ============================================================
#  LECTURA DE ARCHIVOS DE COTIZACION
# ============================================================

def _texto_de_archivo(archivo):
    """
    Extrae el texto de un archivo de cotizacion subido.
    Soporta: .txt, .xlsx/.xls (Excel) y .docx (Word).
    Devuelve el texto plano, o "" si no se pudo leer.
    """
    if not archivo or archivo.filename == "":
        return ""

    nombre = archivo.filename.lower()
    datos = archivo.read()

    # --- Texto plano ---
    if nombre.endswith(".txt"):
        try:
            return datos.decode("utf-8", errors="ignore")
        except Exception:
            return ""

    # --- Excel ---
    if nombre.endswith((".xlsx", ".xls")):
        try:
            from openpyxl import load_workbook
            wb = load_workbook(BytesIO(datos), data_only=True)
            lineas = []
            for ws in wb.worksheets:
                for fila in ws.iter_rows(values_only=True):
                    celdas = [str(c) for c in fila if c not in (None, "")]
                    if celdas:
                        lineas.append("\t".join(celdas))
            return "\n".join(lineas)
        except Exception:
            return ""

    # --- Word ---
    if nombre.endswith(".docx"):
        try:
            from docx import Document
            doc = Document(BytesIO(datos))
            partes = [p.text for p in doc.paragraphs if p.text.strip()]
            # Tambien las tablas del documento
            for tabla in doc.tables:
                for fila in tabla.rows:
                    celdas = [c.text.strip() for c in fila.cells
                              if c.text.strip()]
                    if celdas:
                        partes.append("\t".join(celdas))
            return "\n".join(partes)
        except Exception:
            return ""

    return ""


@app.route("/")
def index():
    """Sirve la pagina principal."""
    return render_template("index.html")


@app.route("/procesar", methods=["POST"])
def procesar_informe():
    """
    Recibe:
      - plantilla    (archivo .xlsx, obligatorio)
      - texto        (texto libre del peritaje, obligatorio)
      - cotizacion   (texto de la cotizacion de repuestos, opcional)
      - cotiz_file   (archivo de cotizacion .txt/.xlsx/.docx, opcional)
    Devuelve:
      - el archivo .xlsx completado para descargar
    """
    try:
        # --- Validar plantilla ---
        if "plantilla" not in request.files:
            return jsonify({"error": "Falta la plantilla virgen"}), 400

        plantilla_file = request.files["plantilla"]
        if plantilla_file.filename == "":
            return jsonify({"error": "No se selecciono ninguna plantilla"}), 400

        if not plantilla_file.filename.lower().endswith(".xlsx"):
            return jsonify({"error": "La plantilla debe ser un archivo .xlsx"}), 400

        plantilla_bytes = plantilla_file.read()

        # --- Texto del peritaje ---
        texto = request.form.get("texto", "").strip()
        if not texto:
            return jsonify({"error": "Pega el texto de la peritacion"}), 400

        # --- Cotizacion: puede venir como texto pegado y/o como archivo ---
        cotizacion = request.form.get("cotizacion", "").strip()

        if "cotiz_file" in request.files:
            texto_archivo = _texto_de_archivo(request.files["cotiz_file"])
            if texto_archivo:
                # Si hay texto pegado Y archivo, se combinan ambos.
                cotizacion = (cotizacion + "\n\n" + texto_archivo).strip()

        # --- Procesar ---
        # La cotizacion se suma al texto del peritaje: el procesador la
        # entiende y empareja los precios con las piezas a CAMBIAR.
        texto_completo = texto
        if cotizacion:
            texto_completo += (
                "\n\n=== COTIZACION DE REPUESTOS ===\n" + cotizacion)

        informe_bytes, data, _ = procesar(
            plantilla_bytes,
            texto_peritacion=texto_completo,
        )

        # --- Nombre del archivo de salida ---
        nro = data.get("numeroSiniestro") or data.get("dominio") or "sin-numero"
        nombre = f"Informe Tasacion Agrosalta - COMPLETO - {nro}.xlsx"

        return send_file(
            BytesIO(informe_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=nombre,
        )

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"Error al procesar: {str(e)}"}), 500


@app.errorhandler(413)
def archivo_muy_grande(e):
    return jsonify({"error": "El archivo es demasiado grande (maximo 10 MB)"}), 413


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
