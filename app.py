"""
app.py
Servidor web Flask para el generador de informes de tasacion.

Endpoints:
  GET  /            -> sirve la pagina web (index.html)
  POST /procesar    -> recibe plantilla + peritacion, devuelve el informe completo
"""

import os
import zipfile
from io import BytesIO
from flask import Flask, request, send_file, render_template, jsonify

from procesador import procesar

app = Flask(__name__)

# Limite de tamano de subida: 10 MB
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024


def _diagnostico_plantilla(plantilla_bytes):
    """
    Imprime en el log del servidor cuantas imagenes tiene la plantilla
    que se acaba de recibir. Sirve para confirmar si el archivo subido
    realmente trae los logos o no.
    """
    try:
        z = zipfile.ZipFile(BytesIO(plantilla_bytes))
        media = [n for n in z.namelist() if n.startswith("xl/media/")]
        drawings = [n for n in z.namelist()
                    if n.startswith("xl/drawings/") and n.endswith(".xml")]
        print("=" * 50)
        print("DIAGNOSTICO PLANTILLA RECIBIDA:")
        print("  Tamano:", len(plantilla_bytes), "bytes")
        print("  Imagenes (xl/media):", len(media), media)
        print("  Drawings:", len(drawings))
        print("=" * 50)
    except Exception as e:
        print("DIAGNOSTICO: no se pudo leer la plantilla -", e)


@app.route("/")
def index():
    """Sirve la pagina principal."""
    return render_template("index.html")


@app.route("/procesar", methods=["POST"])
def procesar_informe():
    """
    Recibe:
      - plantilla     (archivo .xlsx, obligatorio)
      - texto         (texto libre de la peritacion, opcional)
      - excel_perit   (archivo .xlsx de peritacion, opcional)
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

        # --- DIAGNOSTICO: confirmar si la plantilla trae los logos ---
        _diagnostico_plantilla(plantilla_bytes)

        # --- Texto de la peritacion (opcional) ---
        texto = request.form.get("texto", "").strip()

        # --- Excel de peritacion (opcional) ---
        excel_perit_bytes = None
        if "excel_perit" in request.files:
            ef = request.files["excel_perit"]
            if ef.filename != "" and ef.filename.lower().endswith((".xlsx", ".xls")):
                excel_perit_bytes = ef.read()

        # --- Opcion de cotizar en Mercado Libre ---
        # El checkbox del frontend envia "si" cuando esta marcado.
        cotizar = request.form.get("cotizar", "").lower() in ("si", "true", "on", "1")

        # --- Validar que haya al menos una fuente de datos ---
        if not texto and not excel_perit_bytes:
            return jsonify({
                "error": "Carga el texto de la peritacion o subi un Excel con los datos"
            }), 400

        # --- Procesar ---
        informe_bytes, data, detalle_cotizacion = procesar(
            plantilla_bytes,
            texto_peritacion=texto,
            excel_peritacion_bytes=excel_perit_bytes,
            cotizar=cotizar,
        )

        # --- DIAGNOSTICO: confirmar si el informe generado tiene logos ---
        try:
            zo = zipfile.ZipFile(BytesIO(informe_bytes))
            media_out = [n for n in zo.namelist() if n.startswith("xl/media/")]
            print("DIAGNOSTICO INFORME GENERADO -> imagenes:", len(media_out))
        except Exception as e:
            print("DIAGNOSTICO INFORME: error -", e)

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
    return jsonify({"error": "El archivo es demasiado grande (maximo 10 MB)"}), 413


if __name__ == "__main__":
    # Para desarrollo local. En produccion Render usa gunicorn (ver Procfile).
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
