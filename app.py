import os
import base64
import json
import requests
from datetime import datetime
from flask import Flask, request, render_template, redirect, jsonify, send_from_directory
from openpyxl import load_workbook

app = Flask(__name__)
os.makedirs("static/fotos", exist_ok=True)

# Variables de entorno necesarias
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
TENANT_ID = os.environ["TENANT_ID"]
SHAREPOINT_SITE = os.environ["SHAREPOINT_SITE"]  # URL completa
SHAREPOINT_FOLDER = os.environ["SHAREPOINT_FOLDER"]
SHAREPOINT_DOC = os.environ["SHAREPOINT_DOC"]

# Obtener access token con endpoint moderno
def obtener_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "client_id": CLIENT_ID,
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials"
    }
    response = requests.post(url, headers=headers, data=data)
    if response.status_code == 200:
        return response.json()["access_token"]
    else:
        print("Error al obtener token:", response.text)
        raise Exception("No se pudo obtener access_token")

# Subir archivo a SharePoint (con Graph API)
def subir_a_sharepoint(excel_local):
    token = obtener_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }

    # Extraer sitio y nombre de carpeta desde la URL
    sitio = SHAREPOINT_SITE.replace("https://", "")
    carpeta = SHAREPOINT_FOLDER
    archivo = SHAREPOINT_DOC

    url = f"https://graph.microsoft.com/v1.0/sites/{sitio}/drive/root:/{carpeta}/{archivo}:/content"

    with open(excel_local, "rb") as f:
        r = requests.put(url, headers=headers, data=f.read())
    return r.status_code in [200, 201]

# Guardar imágenes/documentos
def guardar_archivo(nombre_base, base64_data):
    if base64_data:
        ext = ".jpg" if "image" in base64_data else ".pdf"
        nombre_archivo = f"{nombre_base}_{datetime.now().strftime('%Y%m%d%H%M%S')}{ext}"
        ruta = os.path.join("static/fotos", nombre_archivo)
        with open(ruta, "wb") as f:
            f.write(base64.b64decode(base64_data.split(",")[1]))
        return nombre_archivo
    return ""

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            datos = request.get_json() if request.is_json else request.form
            archivos = request.files if not request.is_json else None

            archivos_nombres = {
                "ine_frente": guardar_archivo("ine_frente", datos.get("ine_frente")) if request.is_json else archivos["ine_frente"].filename,
                "ine_reverso": guardar_archivo("ine_reverso", datos.get("ine_reverso")) if request.is_json else archivos["ine_reverso"].filename,
                "curp_archivo": guardar_archivo("curp_archivo", datos.get("curp_archivo")) if request.is_json else archivos["curp_archivo"].filename,
                "documentos": guardar_archivo("documentos", datos.get("documentos_base64")) if request.is_json else archivos["documentos"].filename,
                "foto": guardar_archivo("foto", datos.get("foto_base64")) if request.is_json else archivos["foto"].filename,
            }

            fila = [
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"), datos.get("nombre", ""), datos.get("edad", ""), datos.get("curp", ""),
                datos.get("rfc", ""), datos.get("nss", ""), datos.get("telefono", ""), datos.get("direccion", ""),
                datos.get("leer_escribir", ""), datos.get("discapacidad", ""), datos.get("experiencia", ""), datos.get("salud", ""),
                datos.get("origen", ""), datos.get("observaciones", ""), datos.get("trabajo_previo", ""), datos.get("año_trabajo", ""),
                datos.get("area_trabajo", ""), datos.get("contacto_emergencia", ""), datos.get("telefono_emergencia", ""),
                archivos_nombres["ine_frente"], archivos_nombres["ine_reverso"], archivos_nombres["curp_archivo"],
                archivos_nombres["documentos"], archivos_nombres["foto"]
            ]

            archivo_excel_local = "registro_empleados.xlsx"
            wb = load_workbook(archivo_excel_local)
            ws = wb.active
            ws.append(fila)
            wb.save(archivo_excel_local)

            if subir_a_sharepoint(archivo_excel_local):
                return jsonify({"ok": True}), 200
            else:
                return jsonify({"error": "No se pudo subir a SharePoint"}), 500

        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"error": str(e)}), 500

    return render_template("index.html")

@app.route("/service-worker.js")
def sw():
    return send_from_directory("static", "service-worker.js")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
