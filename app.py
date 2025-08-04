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
TENANT_ID = os.environ["TENANT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
SHAREPOINT_SITE = "agricactus2.sharepoint.com"
SHAREPOINT_SITE_NAME = "CALIDAD"
SHAREPOINT_DOC = "registro_empleados.xlsx"

# Función para obtener token de acceso
def obtener_token():
    url = f"https://accounts.accesscontrol.windows.net/{TENANT_ID}/tokens/OAuth/2"
    payload = {
        'grant_type': 'client_credentials',
        'client_id': f'{CLIENT_ID}@{TENANT_ID}',
        'client_secret': CLIENT_SECRET,
        'resource': f'spo.azure.com'
    }
    r = requests.post(url, data=payload)
    return r.json()["access_token"]

# Función para subir archivo a SharePoint
def subir_a_sharepoint(excel_local):
    access_token = obtener_token()
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    upload_url = f"https://{SHAREPOINT_SITE}/sites/{SHAREPOINT_SITE_NAME}/_api/web/GetFolderByServerRelativeUrl('Shared Documents')/Files/add(url='{SHAREPOINT_DOC}',overwrite=true)"

    with open(excel_local, 'rb') as f:
        response = requests.post(upload_url, headers=headers, data=f)
    return response.status_code == 200

# Guardar archivos individuales

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

@app.route('/service-worker.js')
def sw():
    return send_from_directory('static', 'service-worker.js')

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)

