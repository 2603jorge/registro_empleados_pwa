function estaConectado() {
  return navigator.onLine;
}

let db;
const request = indexedDB.open("RegistroTrabajadoresDB", 1);

request.onupgradeneeded = function (event) {
  db = event.target.result;
  if (!db.objectStoreNames.contains("trabajadores")) {
    db.createObjectStore("trabajadores", { autoIncrement: true });
  }
};

request.onsuccess = function (event) {
  db = event.target.result;
};

function guardarOffline(datos) {
  const tx = db.transaction("trabajadores", "readwrite");
  const store = tx.objectStore("trabajadores");
  store.add(datos);
  tx.oncomplete = () => {
    alert("🗃️ Guardado offline. Se sincronizará cuando haya internet.");
  };
}

// Detectar envío del formulario
document.getElementById("registro-form").addEventListener("submit", function (e) {
  if (!estaConectado()) {
    e.preventDefault();

    const formData = new FormData(e.target);
    const datos = {};

    formData.forEach((valor, clave) => {
      datos[clave] = valor;
    });

    // Convertir imágenes a base64
    const docFile = formData.get("foto_documentos");
    const rostroFile = formData.get("foto_rostro");

    const leerImagen = archivo => new Promise((res, rej) => {
      if (!archivo) return res(null);
      const lector = new FileReader();
      lector.onload = () => res(lector.result);
      lector.onerror = err => rej(err);
      lector.readAsDataURL(archivo);
    });

    Promise.all([
      leerImagen(docFile),
      leerImagen(rostroFile)
    ]).then(([docBase64, rostroBase64]) => {
      datos["foto_documentos_base64"] = docBase64;
      datos["foto_rostro_base64"] = rostroBase64;
      guardarOffline(datos);
    });
  }
});

// Sincronizar al volver la conexión
window.addEventListener("online", sincronizarDatos);

function sincronizarDatos() {
  if (!db) return;
  const tx = db.transaction("trabajadores", "readonly");
  const store = tx.objectStore("trabajadores");
  const getAll = store.getAll();

  getAll.onsuccess = function () {
    const registros = getAll.result;
    if (registros.length === 0) return;

    registros.forEach((registro, index) => {
      fetch("/", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(registro)
      })
        .then(res => {
          if (res.ok) {
            const txDel = db.transaction("trabajadores", "readwrite");
            const storeDel = txDel.objectStore("trabajadores");
            storeDel.delete(index + 1);
            console.log("✅ Sincronizado y eliminado:", registro.nombre);
          }
        })
        .catch(err => console.error("❌ Error al sincronizar:", err));
    });

    alert("🔄 ¡Datos sincronizados correctamente!");
  };
}
