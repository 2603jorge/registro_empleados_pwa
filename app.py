from flask import Flask, render_template, request, redirect, jsonify, send_from_directory
import os
import base64
from datetime import datetime
import openpyxl
import requests

app = Flask(__name__)
os.makedirs("static/fotos", exist_ok=True)

# Datos SharePoint
SHAREPOINT_SITE = "https://agricactus2.sharepoint.com/sites/CALIDAD"
SHAREPOINT_FOLDER = "RegistrosEmpleados"
EXCEL_FILE_NAME = "registro_empleados.xlsx"
CLIENT_ID = os.environ.get("SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = os.environ.get("SHAREPOINT_CLIENT_SECRET")
TENANT_ID = os.environ.get("SHAREPOINT_TENANT_ID")

# Obtener token
def obtener_token():
    url = f"https://accounts.accesscontrol.windows.net/{TENANT_ID}/tokens/OAuth/2"
    datos = {
        "grant_type": "client_credentials",
        "client_id": f"{CLIENT_ID}@{TENANT_ID}",
        "client_secret": CLIENT_SECRET,
        "resource": f"00000003-0000-0ff1-ce00-000000000000/{'agricactus2.sharepoint.com'}@{TENANT_ID}"
    }
    respuesta = requests.post(url, data=datos)
    return respuesta.json().get("access_token")

# Subir archivo Excel a SharePoint
def subir_a_sharepoint(ruta_local):
    token = obtener_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json;odata=verbose"
    }
    # Leer archivo
    with open(ruta_local, "rb") as f:
        contenido = f.read()
    # Subirlo
    url = f"{SHAREPOINT_SITE}/_api/web/GetFolderByServerRelativeUrl('{SHAREPOINT_FOLDER}')/Files/add(url='{EXCEL_FILE_NAME}',overwrite=true)"
    requests.post(url, headers=headers, data=contenido)

# Guardar foto
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
            archivos = request.files if not request.is_json else {}

            nombres_archivos = {
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
                datos.get("telefono_emergencia", ""), nombres_archivos["ine_frente"], nombres_archivos["ine_reverso"],
                nombres_archivos["curp_archivo"], nombres_archivos["documentos"], nombres_archivos["foto"]
            ]

            ruta_local = os.path.abspath(EXCEL_FILE_NAME)
            if not os.path.exists(ruta_local):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.append([
                    "Fecha", "Nombre", "Edad", "CURP", "RFC", "NSS", "Teléfono",
                    "Dirección", "Leer/Escribir", "Discapacidad", "Experiencia",
                    "Salud", "Origen", "Observaciones", "Trabajo previo",
                    "Año trabajo", "Área trabajo", "Contacto emergencia",
                    "Teléfono emergencia", "INE Frente", "INE Reverso", "CURP Archivo",
                    "Acta/Comprobante", "Foto facial"
                ])
            else:
                wb = openpyxl.load_workbook(ruta_local)
                ws = wb.active

            ws.append(fila)
            wb.save(ruta_local)

            subir_a_sharepoint(ruta_local)

            return jsonify({"ok": True}), 200
        except Exception as e:
            print(f"Error: {e}")
            return "Error interno", 500
    return render_template("index.html")

@app.route('/service-worker.js')
def sw():
    return send_from_directory('static', 'service-worker.js')

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)

