from flask import Flask, render_template, request, redirect, jsonify, send_from_directory
import os
import base64
import openpyxl
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)
archivo_excel = os.path.abspath("registro_empleados.xlsx")
os.makedirs("static/fotos", exist_ok=True)

# Configuración Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credenciales = ServiceAccountCredentials.from_json_keyfile_name(
    "endless-phoenix-308700-3040c62984f0.json", scope
)
cliente = gspread.authorize(credenciales)
sheet_id = "1n1LBA7VkK05-8RbBBEaPZ3jwMCzZKTV9EcaqPlE3nfs"
hoja = cliente.open_by_key(sheet_id).sheet1

# Función para guardar fila en Google Sheets
def guardar_en_sheets(fila):
    try:
        hoja.append_row(fila)
    except Exception as e:
        print(f"❌ Error al guardar en Google Sheets: {e}")

# Inicializa el Excel si no existe
if not os.path.exists(archivo_excel):
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
    wb.save(archivo_excel)

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
                    datos.get("nombre", ""),
                    datos.get("edad", ""),
                    datos.get("curp", ""),
                    datos.get("rfc", ""),
                    datos.get("nss", ""),
                    datos.get("telefono", ""),
                    datos.get("direccion", ""),
                    datos.get("leer_escribir", ""),
                    datos.get("discapacidad", ""),
                    datos.get("experiencia", ""),
                    datos.get("salud", ""),
                    datos.get("origen", ""),
                    datos.get("observaciones", ""),
                    datos.get("trabajo_previo", ""),
                    datos.get("año_trabajo", ""),
                    datos.get("area_trabajo", ""),
                    datos.get("contacto_emergencia", ""),
                    datos.get("telefono_emergencia", ""),
                    archivos["ine_frente"],
                    archivos["ine_reverso"],
                    archivos["curp_archivo"],
                    archivos["documentos"],
                    archivos["foto"]
                ]

                wb = openpyxl.load_workbook(archivo_excel)
                ws = wb.active
                ws.append(fila)
                wb.save(archivo_excel)

                guardar_en_sheets(fila)

                return jsonify({"ok": True}), 200

            # HTML convencional
            form = request.form
            archivos = request.files

            nombres_archivos = {}
            for key in ["ine_frente", "ine_reverso", "curp_archivo", "documentos", "foto"]:
                file = archivos.get(key)
                if file and file.filename:
                    nombre_archivo = f"{key}_{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}"
                    file.save(os.path.join("static/fotos", nombre_archivo))
                    nombres_archivos[key] = nombre_archivo
                else:
                    nombres_archivos[key] = ""

            fila = [
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                form.get("nombre", ""),
                form.get("edad", ""),
                form.get("curp", ""),
                form.get("rfc", ""),
                form.get("nss", ""),
                form.get("telefono", ""),
                form.get("direccion", ""),
                form.get("leer_escribir", ""),
                form.get("discapacidad", ""),
                form.get("experiencia", ""),
                form.get("salud", ""),
                form.get("origen", ""),
                form.get("observaciones", ""),
                form.get("trabajo_previo", ""),
                form.get("año_trabajo", ""),
                form.get("area_trabajo", ""),
                form.get("contacto_emergencia", ""),
                form.get("telefono_emergencia", ""),
                nombres_archivos["ine_frente"],
                nombres_archivos["ine_reverso"],
                nombres_archivos["curp_archivo"],
                nombres_archivos["documentos"],
                nombres_archivos["foto"]
            ]

            wb = openpyxl.load_workbook(archivo_excel)
            ws = wb.active
            ws.append(fila)
            wb.save(archivo_excel)

            guardar_en_sheets(fila)

            return redirect("/")

        except Exception as e:
            print(f"❌ Error al guardar: {e}")
            return "Error interno", 500

    return render_template("index.html")

@app.route('/service-worker.js')
def sw():
    return send_from_directory('static', 'service-worker.js')

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
