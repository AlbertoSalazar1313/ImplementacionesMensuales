import os
import re
import mysql.connector
import requests
import warnings
from PyPDF2 import PdfReader
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
from tkinter import PhotoImage
from tkinter import Label
from tkcalendar import DateEntry
from openpyxl import Workbook
from datetime import date
import time

warnings.filterwarnings("ignore", category=UserWarning)

CONFIG_BD = {
    "host": "",
    "port": ,
    "user": "",
    "password": "",
    "charset": "utf8"    
}

DOWNLOAD_DIR = ""
OUTPUT_PATH = ""
OUTPUT_FILE = ""

def nombreArchivo(strFechaInicio, strFechaFin):
    nombreBase = f"resumen del {strFechaInicio} al {strFechaFin}"
    extension = ".xlsx"
    archivo = os.path.join(OUTPUT_PATH, nombreBase + extension)
    i = 1    
    while os.path.exists(archivo):
        archivo = os.path.join(OUTPUT_PATH, f"{nombreBase} ({i}){extension}")
        i += 1
    global OUTPUT_FILE
    OUTPUT_FILE = archivo
    return archivo

def validarInputUsuario(input):
    if input == "":
        return True
    return input.isdigit() and len(input) <= 10

def seleccionarCarpeta(entry):
    carpeta = filedialog.askdirectory()
    if carpeta:
        entry.delete(0, tk.END)
        entry.insert(0, carpeta)

def reiniciaBarra():
    barraProgreso["value"]=0
    root.update_idletasks()
    etiquetaProgreso.config(text=f"Esperando...")

def descargaPDF(url, ubicacionDescarga):
    archivo = os.path.join(ubicacionDescarga, os.path.basename(url))
    if not os.path.exists(archivo):
        r = requests.get(url, stream=True)
        if r.status_code == 200:
            with open(archivo, "wb") as f:
                f.write(r.content)
    return archivo

def extraerTextoEntre(texto, textoInicio, textoFin):
    try:
        inicia = texto.index(textoInicio) + len(textoInicio)
        fin = texto.index(textoFin, inicia)
        return texto[inicia:fin].strip()
    except ValueError:
        return ""

def extraerTextoPDF(ruta_pdf):
    reader = PdfReader(ruta_pdf)
    text = ""
    for page in reader.pages:
        if page.extract_text():
            text += page.extract_text() + "\n"
    if not text.strip():
        return {
            "Nombre del Sistema" : "",
            "Versión" : "",
            "Solicitante" : "",
            "Fecha y Hora de cambio": "",
            "Desarrollador": "",
            "Descripción del cambio": "",
            "Sucursales donde se efectuará el cambio": ""
        }
    nombre_sistema = extraerTextoEntre(
        text,
        "Nombre del Sistema",
        "Versión"
        )
    version = extraerTextoEntre(
        text,
        "Versión:",
        "# PR"
        )
    solicitante = extraerTextoEntre(
        text,
        "Solicitante:",
        "Teléfono:"
        )
    descripcion_cambio = extraerTextoEntre(
        text,
        "Descripción del cambio (Escrito como historia de usuario)",
        "Especificaciones:"
    )
    sucursales = extraerTextoEntre(
        text,
        "Sucursales donde se efectuará el cambio",
        "Información de pruebas (QA)"
    )
    campos = {
        "Fecha y Hora de cambio": re.search(r"Fecha y Hora de cambio[:\- ]+(.*)", text),
        "Desarrollador": re.search(r"Desarrollador[:\- ]+(.*)", text)
    }
    camposExtraidos = {
        "Nombre sistema" : nombre_sistema.lstrip(":"),
        "Versión" : version.lstrip(":"),
        "Solicitante" : solicitante.lstrip(":"),
        "Fecha y Hora de cambio": campos["Fecha y Hora de cambio"].group(1).lstrip(":").strip() if campos["Fecha y Hora de cambio"] else "",
        "Desarrollador": campos["Desarrollador"].group(1).lstrip(":").strip() if campos["Desarrollador"] else "",
        "Descripción del cambio": descripcion_cambio.lstrip(":\n\r ").strip(),
        "Sucursales donde se efectuará el cambio": sucursales.lstrip(":\n\r ").strip()
    }
    return camposExtraidos

def procesaImplementaciones(id_usuario, strFechaInicio, strFechaFin, barraProgreso, etiquetaProgreso, root):
    
    etiquetaProgreso.config(text="Conectabdo a BD...")
    root.update_idletasks()
    
    conn = mysql.connector.connect(**CONFIG_BD)
    cursor = conn.cursor(dictionary=True)

    etiquetaProgreso.config(text="Consultando idpersonal...")
    root.update_idletasks()
    
    query = f'''SELECT idpersonal FROM personal.personal WHERE idusuario = "{id_usuario}";''';
    cursor.execute(query)
    rowp = cursor.fetchone()

    if rowp is None:
        messagebox.showerror("Error", "El usuario no existe.")
        return
    idpersonal = rowp["idpersonal"]

    etiquetaProgreso.config(text="Consultando idoficinapuesto...")
    root.update_idletasks()
    
    query = f'''SELECT idoficinapuesto FROM personal.oficinapuesto WHERE idpersonal = "{idpersonal}";''';
    cursor.execute(query)
    rowi = cursor.fetchone()

    if rowi is None:
        messagebox.showerror("Error", "No hay registro en oficinapuesto.")
        return
    idoficinapuesto = rowi["idoficinapuesto"]

    etiquetaProgreso.config(text="Consultando id_correo_oficina_puesto...")
    root.update_idletasks()
    
    query = f'''SELECT id_correo_oficina_puesto FROM documentos.correo_oficinapuesto WHERE idoficinapuesto = "{idoficinapuesto}";''';
    cursor.execute(query)
    rowc = cursor.fetchone()

    if rowc is None:
        messagebox.showerror("Error", "No hay registro en correo_oficinapuesto.")
        return
    idcorreooficinapuesto = rowc["id_correo_oficina_puesto"]

    etiquetaProgreso.config(text="Consultando implementaciones...")
    root.update_idletasks()
    query = f"""SELECT
    IF(ds.version_solicitud = 0, ds.folio, CONCAT(ds.folio, ' - v', ds.version_solicitud)) folioSolicitud, CASE WHEN a.nombre LIKE "%QA%" THEN "SI" ELSE "NO" END AS QA, a.ruta
    FROM documentos.archivo_solicitud a
    JOIN documentos.solicitud ds  ON a.id_solicitud=ds.id_solicitud AND ds.id_tipo_solicitud_proceso = 26 
    JOIN documentos.personal_solicitud dps ON ds.id_solicitud = dps.id_solicitud AND dps.id_correo_oficina_puesto_destinatario = "{idcorreooficinapuesto}" 
    JOIN documentos.correo_oficinapuesto dco ON dps.id_correo_oficina_puesto_destinatario = dco.id_correo_oficina_puesto  
    JOIN personal.oficinapuesto pop ON dco.idoficinapuesto = pop.idoficinapuesto AND pop.idpersonal = {idpersonal}
    WHERE dps.id_estatus_solicitud NOT IN (6,11) AND ds.fecha_alta BETWEEN "{strFechaInicio}" AND "{strFechaFin}";"""
    cursor.execute(query)
    results = cursor.fetchall()

    total = len(results)
    barraProgreso["maximum"] = total

    etiquetaProgreso.config(text=f"Procesando {total} registros...")
    root.update_idletasks()
    

    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"

    cabeceras = [
        "FOLIO",
        "NOMBRE DEL SISTEMA",
        "VERSIÓN",
        "FECHA SOLICITUD",
        "FECHA SOL IMPLEMENTACIÓN",
        "% IMPLEMENTACIÓN",
        "GRADO DE IMPACTO(ALTO, MEDIO, BAJO)",
        "SOLICITANTE",
        "DESARROLLADOR",
        "SUCURSALES SOLICITADAS",
        "TIPO DE ARCHIVO (EXE, WAR, JAR, ETC.)",
        "DESCRIPCIÓN CAMBIOS APLICADOS",
        "QA (SI,NO)",
        "TUVO ROLLBACK (SI, NO)",
        "REPORTO INCIDENCIA (SI, NO)",
        "RESPONSABLE INCIDENCIA (DESARROLLO, USUARIO, SAP, ETC.)",
        "Nombre Proyecto / No. Ticket / Nombre Actividad","Observaciones"
    ]
    ws.append(cabeceras)
    for i, row in enumerate(results, start=1):
        
        etiquetaProgreso.config(text=f"Procesando {i} de {total}...")
        root.update_idletasks()
        url_pdf = row["ruta"]
        ruta_pdf = descargaPDF(url_pdf, DOWNLOAD_DIR)
        textoExtraido = extraerTextoPDF(ruta_pdf)

        ws.append([
            row["folioSolicitud"],
            textoExtraido["Nombre sistema"],
            textoExtraido["Versión"], 
            textoExtraido["Fecha y Hora de cambio"],
            textoExtraido["Fecha y Hora de cambio"],
            "0", #% IMPLEMENTACIÓN
            "TBD", #GRADO IMPACTO
            textoExtraido["Solicitante"],
            textoExtraido["Desarrollador"],
            textoExtraido["Sucursales donde se efectuará el cambio"],
            "TBD", #TIPO ARCHIVO
            textoExtraido["Descripción del cambio"],
            row["QA"],
            "TBD", #ROLLBACK
            "TBD", #INCIDENCIA
            "TBD", #RESPONSABLE INCIDENCIA
            "TBD", #NOMBRE PROYECTO
            "TBD" #OBSERVACIONES
        ])
        barraProgreso["value"] = i
        root.update_idletasks()
        

    etiquetaProgreso.config(text="Exportando archivo...")
    root.update_idletasks()
    wb.save(nombreArchivo(strFechaInicio, strFechaFin))
    
    etiquetaProgreso.config(text="Cerrando Conexiones...")
    root.update_idletasks()
    
    cursor.close()
    conn.close()

    etiquetaProgreso.config(text="Proceso Terminado!")
    root.update_idletasks()
    

    messagebox.showinfo("Éxito", f"Proceso completado.\nArchivo generado en:\n{OUTPUT_FILE}")

def btnProcesar(inputUsuario, inputFechaInicio, inputFechaFin, inputDownload, inputOutput, barraProgreso, etiquetaProgreso, root):   
    usuario = inputUsuario.get()
    rutaDescargas = inputDownload.get()
    rutaSalida = inputOutput.get()
    id_usuario = inputUsuario.get();
    strFechaInicio = inputFechaInicio.get_date().strftime("%Y-%m-%d")
    strFechaFin = inputFechaFin.get_date().strftime("%Y-%m-%d")

    if not (len(usuario) == 10):
        messagebox.showerror("Error", "El usuario debe ser de 10 dígitos.")
        return

    if strFechaInicio > strFechaFin:
        messagebox.showerror("Error", f"La fecha inicio no puede ser mayor a la fecha fin.")
        return

    if (inputFechaFin.get_date()-inputFechaInicio.get_date()).days > 31:
        messagebox.showerror("Error", f"El rango de fechas no puede ser mayor a 31 días.")
        return

    if not os.path.exists(rutaDescargas):
        messagebox.showerror("Error", f"La carpeta de descarga no existe:\n{rutaDescargas}")
        return
    if not os.path.exists(rutaSalida):
        messagebox.showerror("Error", f"La carpeta de salida no existe:\n{rutaSalida}")
        return

    global DOWNLOAD_DIR, OUTPUT_PATH
    DOWNLOAD_DIR = rutaDescargas
    OUTPUT_PATH = rutaSalida

    try:
        procesaImplementaciones(id_usuario, strFechaInicio, strFechaFin, barraProgreso, etiquetaProgreso, root)
        root.after(1500,reiniciaBarra)
    except Exception as e:
        messagebox.showerror("Error", str(e))

root = tk.Tk()
root.title("Implementaciones Mensuales DO")
root.geometry("370x275")
root.resizable(False,False)

validarUsuario = (root.register(validarInputUsuario), "%P")
tk.Label(root, text="Usuario:").grid(row=0, column=0, padx=10, pady=5)
inputUsuario = tk.Entry(root, width=15, validate="key", validatecommand=validarUsuario)
inputUsuario.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="Fecha Inicio:").grid(row=1, column=0, padx=10, pady=5)
inputFechaInicio = DateEntry(root, width=12, background='steel blue', foreground='white', borderwidth=2, state="readonly", date_pattern='yyyy-mm-dd')
inputFechaInicio.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Fecha Fin:").grid(row=2, column=0, padx=10, pady=5)
inputFechaFin = DateEntry(root, width=12, background='steel blue', foreground='white', borderwidth=2, state="readonly", date_pattern='yyyy-mm-dd')
inputFechaFin.grid(row=2, column=1, padx=10, pady=5)


tk.Label(root, text="Carpeta Descarga:").grid(row=3, column=0, padx=10, pady=5)
inputDownload = tk.Entry(root, width=15)
inputDownload.grid(row=3, column=1, padx=10, pady=5)
tk.Button(root, text="Seleccionar", command=lambda: seleccionarCarpeta(inputDownload)).grid(row=3, column=2, padx=5, pady=5)

tk.Label(root, text="Carpeta Salida:").grid(row=4, column=0, padx=10, pady=5)
inputOutput = tk.Entry(root, width=15)
inputOutput.grid(row=4, column=1, padx=10, pady=5)
tk.Button(root, text="Seleccionar", command=lambda: seleccionarCarpeta(inputOutput)).grid(row=4, column=2, padx=5, pady=5)

btn = tk.Button(root, text="Procesar")
btn.grid(row=5, column=0, columnspan=3, pady=10)
btn.config(command=lambda: btnProcesar(inputUsuario, inputFechaInicio, inputFechaFin, inputDownload, inputOutput, barraProgreso, etiquetaProgreso, root))

barraProgreso = ttk.Progressbar(root, length=350, mode="determinate")
barraProgreso.grid(row=6, column=0, columnspan=3, padx=10, pady=5)

etiquetaProgreso = tk.Label(root, text="Esperando...", anchor="w")
etiquetaProgreso.grid(row=7, column=0, columnspan=3, padx=10, pady=5)

root.mainloop()
