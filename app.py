"""
app.py
Servidor web Flask para el generador de informes de tasación.

Endpoints:
  GET  /            -> sirve la página web (index.html)
  POST /procesar    -> recibe plantilla + peritación, devuelve el informe completo
"""

import os
from io import BytesIO
from flask import Flask, request, send_file, render_template, jsonify

from procesador import procesar

app = Flask(__name__)

# Límite de tamaño de subida: 10 MB
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024


@app.route("/")
def index():
    """Sirve la página principal."""
    return render_template("index.html")


@app.route("/procesar", methods=["POST"])
def procesar_informe():
    """
    Recibe:
      - plantilla     (archivo .xlsx, obligatorio)
      - texto         (texto libre de la peritación, opcional)
      - excel_perit   (archivo .xlsx de peritación, opcional)
    Devuelve:
      - el archivo .xlsx completado para descargar
    """
    try:
        # --- Validar plantilla ---
        if "plantilla" not in request.files:
            return jsonify({"error": "Falta la plantilla virgen"}), 400

        plantilla_file = request.files["plantilla"]
        if plantilla_file.filename == "":
            return jsonify({"error": "No se seleccionó ninguna plantilla"}), 400

        if not plantilla_file.filename.lower().endswith(".xlsx"):
            return jsonify({"error": "La plantilla debe ser un archivo .xlsx"}), 400

        plantilla_bytes = plantilla_file.read()

        # --- Texto de la peritación (opcional) ---
        texto = request.form.get("texto", "").strip()

        # --- Excel de peritación (opcional) ---
        excel_perit_bytes = None
        if "excel_perit" in request.files:
            ef = request.files["excel_perit"]
            if ef.filename != "" and ef.filename.lower().endswith((".xlsx", ".xls")):
                excel_perit_bytes = ef.read()

        # --- Opción de cotizar en Mercado Libre ---
        # El checkbox del frontend envía "si" cuando está marcado.
        cotizar = request.form.get("cotizar", "").lower() in ("si", "true", "on", "1")

        # --- Validar que haya al menos una fuente de datos ---
        if not texto and not excel_perit_bytes:
            return jsonify({
                "error": "Cargá el texto de la peritación o subí un Excel con los datos"
            }), 400

        # --- Procesar ---
        informe_bytes, data, detalle_cotizacion = procesar(
            plantilla_bytes,
            texto_peritacion=texto,
            excel_peritacion_bytes=excel_perit_bytes,
            cotizar=cotizar,
        )

        # --- Nombre del archivo de salida ---
        nro = data.get("numeroSiniestro") or data.get("dominio") or "sin-numero"
        nombre = f"Informe Tasacion Agrosalta - COMPLETO - {nro}.xlsx"

        # --- Devolver el archivo ---
        return send_file(
            BytesIO(informe_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=nombre,
        )

    except Exception as e:
        # Log del error en consola del servidor
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"Error al procesar: {str(e)}"}), 500


@app.errorhandler(413)
def archivo_muy_grande(e):
    return jsonify({"error": "El archivo es demasiado grande (máximo 10 MB)"}), 413


if __name__ == "__main__":
    # Para desarrollo local. En producción Render usa gunicorn (ver Procfile).
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)