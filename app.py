import pdfplumber
import pandas as pd
import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
 
def procesar_pdfs(archivos_pdf):
    """
    Procesa una lista de archivos PDF, extrae datos y los devuelve como una lista de diccionarios.
    """
    resultados = []
 
    for archivo in archivos_pdf:
        try:
            print(f"Procesando: {archivo}")
 
            # Abrir PDF y extraer texto
            with pdfplumber.open(archivo) as pdf:
                page = pdf.pages[0]  # Procesar solo la primera página
                texto = page.extract_text()
 
                # Crear un diccionario para los datos extraídos
                fila = {"Archivo": os.path.basename(archivo)}
 
                # Buscar el valor de "VALOR DEL PAGO"
                lineas = texto.split("\n")
                for linea in lineas:
                    if "VALOR DEL PAGO" in linea:
                        valores = linea.strip().split()
                        for i, valor in enumerate(valores):
                            fila[f"Columna_{i+1}"] = valor
                        break
 
                # Buscar el valor de "CODIGO ESTABLECIMIENTO:"
                for linea in lineas:
                    if "CODIGO ESTABLECIMIENTO:" in linea:
                        partes = linea.split("CODIGO ESTABLECIMIENTO:")
                        if len(partes) > 1:
                            codigo = partes[1].strip().split()[0]
                            fila["Codigo_Establecimiento"] = codigo
                        break
 
                # Buscar el valor de "NOTA DE CRÉDITO:"
                for linea in lineas:
                    if "NOTA DE CRÉDITO:" in linea:
                        partes = linea.split("NOTA DE CRÉDITO:")
                        if len(partes) > 1:
                            nota_credito = partes[1].strip().split()[0]
                            fila["Nota_Credito"] = nota_credito
                        break
 
                resultados.append(fila)
 
        except Exception as e:
            print(f"Error procesando {archivo}: {str(e)}")
 
    return resultados
 
def guardar_excel(resultados):
    """
    Guarda los resultados en un archivo Excel.
    """
    if resultados:
        df = pd.DataFrame(resultados)
        columnas_a_eliminar = [f"Columna_{i}" for i in range(1, 7)]
        df = df.drop(columns=[col for col in columnas_a_eliminar if col in df.columns])
 
        columnas = ["Archivo", "Codigo_Establecimiento", "Nota_Credito"] + [
            col for col in df.columns if col not in ["Archivo", "Codigo_Establecimiento", "Nota_Credito"]
        ]
        df = df[columnas]
 
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_excel = f"valores_separados_{timestamp}.xlsx"
        df.to_excel(nombre_excel, index=False)
 
        messagebox.showinfo("Éxito", f"Archivo Excel creado como '{nombre_excel}'")
 
def mostrar_resultados_en_tabla(resultados, tabla):
    """
    Muestra los resultados extraídos en la tabla de la interfaz.
    """
    # Limpiar tabla antes de insertar datos nuevos
    for item in tabla.get_children():
        tabla.delete(item)
 
    if resultados:
        df = pd.DataFrame(resultados)
        columnas_a_eliminar = [f"Columna_{i}" for i in range(1, 7)]
        df = df.drop(columns=[col for col in columnas_a_eliminar if col in df.columns])
 
        columnas = ["Archivo", "Codigo_Establecimiento", "Nota_Credito"] + [
            col for col in df.columns if col not in ["Archivo", "Codigo_Establecimiento", "Nota_Credito"]
        ]
        df = df[columnas]
 
        # Configurar columnas de la tabla
        tabla["columns"] = list(df.columns)
        for col in df.columns:
            tabla.heading(col, text=col)
            tabla.column(col, anchor=tk.CENTER, width=150)
 
        # Insertar datos en la tabla
        for _, row in df.iterrows():
            tabla.insert("", tk.END, values=list(row))
 
def seleccionar_archivos():
    """
    Abre un cuadro de diálogo para seleccionar archivos PDF y muestra los resultados.
    """
    archivos_pdf = filedialog.askopenfilenames(
        title="Seleccionar archivos PDF",
        filetypes=[("Archivos PDF", "*.pdf")]
    )
    if archivos_pdf:
        resultados = procesar_pdfs(archivos_pdf)
        if resultados:
            mostrar_resultados_en_tabla(resultados, tabla)
            boton_guardar.config(state=tk.NORMAL, command=lambda: guardar_excel(resultados))
        else:
            messagebox.showwarning("Sin datos", "No se encontraron datos en los archivos seleccionados.")
    else:
        messagebox.showinfo("Cancelado", "No se seleccionaron archivos.")
 
# Crear la interfaz gráfica
ventana = tk.Tk()
ventana.title("Extractor de Datos de PDF")
ventana.geometry("1000x600")
ventana.resizable(False, False)
 
# Etiqueta principal
etiqueta = tk.Label(ventana, text="Selecciona tus archivos PDF para procesar", font=("Arial", 14))
etiqueta.pack(pady=10)
 
# Botón para seleccionar archivos PDF
boton_seleccionar = tk.Button(ventana, text="Seleccionar Archivos PDF", command=seleccionar_archivos, font=("Arial", 12), bg="blue", fg="white")
boton_seleccionar.pack(pady=10)
 
# Tabla para mostrar resultados
tabla = ttk.Treeview(ventana, show="headings", height=20)
tabla.pack(pady=10, fill=tk.BOTH, expand=True)
 
# Botón para guardar los resultados en Excel
boton_guardar = tk.Button(ventana, text="Guardar en Excel", state=tk.DISABLED, font=("Arial", 12), bg="green", fg="white")
boton_guardar.pack(pady=10)
 
# Iniciar la aplicación
ventana.mainloop()
