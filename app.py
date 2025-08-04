from flask import Flask, render_template, request, redirect, jsonify, send_from_directory
import os
import base64
from datetime import datetime
import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)
os.makedirs("static/fotos", exist_ok=True)

# === CREDENCIALES DE GOOGLE DESDE VARIABLE DE ENTORNO ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credenciales_dict = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT"])
credenciales = ServiceAccountCredentials.from_json_keyfile_dict(credenciales_dict, scope)
cliente = gspread.authorize(credenciales)

# === ABRIR ARCHIVO DE SHEETS ===
sheet_id = "1n1LBA7VkK05-8RbBBEaPZ3jwMCzZKTV9EcaqPlE3nfs"
hoja = cliente.open_by_key(sheet_id).sheet1

def guardar_archivo(nombre_base, base64_data):
    if base64_data:
        ext = ".jpg" if "image" in base64_data else ".pdf"
        nombre_archivo = f"{nombre_base}_{datetime.now().strftime('%Y%m%d%H%M%S')}{ext}"
        ruta = os.path.join("static/fotos", nombre_archivo)
        with open(ruta, "wb") as f:
            f.write(base64.b64decode(base64_data.split(",")[1]))
        return nombre_archivo
    return ""

def guardar_en_sheets(fila):
    try:
        hoja.append_row(fila)
    except Exception as e:
        print(f"❌ Error al guardar en Google Sheets: {e}")

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            if request.is_json:
                datos = request.get_json()

                archivos = {
                    "ine_frente": guardar_archivo("ine_frente", datos.get("ine_frente")),
                    "ine_reverso": guardar_archivo("ine_reverso", datos.get("ine_reverso")),
                    "curp_archivo": guardar_archivo("curp_archivo", datos.get("curp_archivo")),
                    "documentos": guardar_archivo("documentos", datos.get("documentos_base64")),
                    "foto": guardar_archivo("foto", datos.get("foto_base64"))
                }

                fila = [
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    datos.get("nombre", ""), datos.get("edad", ""), datos.get("curp", ""),
                    datos.get("rfc", ""), datos.get("nss", ""), datos.get("telefono", ""),
                    datos.get("direccion", ""), datos.get("leer_escribir", ""), datos.get("discapacidad", ""),
                    datos.get("experiencia", ""), datos.get("salud", ""), datos.get("origen", ""),
                    datos.get("observaciones", ""), datos.get("trabajo_previo", ""), datos.get("año_trabajo", ""),
                    datos.get("area_trabajo", ""), datos.get("contacto_emergencia", ""),
                    datos.get("telefono_emergencia", ""), archivos["ine_frente"], archivos["ine_reverso"],
                    archivos["curp_archivo"], archivos["documentos"], archivos["foto"]
                ]

                guardar_en_sheets(fila)
                return jsonify({"ok": True}), 200

        except Exception as e:
            print(f"❌ Error al procesar: {e}")
            return "Error interno", 500

    return render_template("index.html")

@app.route("/service-worker.js")
def sw():
    return send_from_directory("static", "service-worker.js")

# Render usará gunicorn app:app
