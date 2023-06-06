import tkinter as tk
from tkinter import messagebox
import pandas as pd
import os
import glob
import tkinter.ttk as ttk

def buscar_archivo(nombre_archivo, directorio):
    patron = f"{directorio}/**/{nombre_archivo}"
    archivos_encontrados = glob.glob(patron, recursive=True)
    return archivos_encontrados

def leer_archivos_excel(ruta):
    archivos_excel = []
    for archivo in os.listdir(ruta):
        if archivo.endswith(".xlsx"):
            archivos_excel.append(archivo)
    return archivos_excel

def cargar_valores_columnas():
    nombre_archivo = "VR. Académica.xlsx"  # Aquí se define el nombre del archivo a buscar
    archivos_excel = buscar_archivo(nombre_archivo, ruta_archivos)
    if archivos_excel:
        archivo = archivos_excel[0]  # Tomamos el primer archivo encontrado
        df = pd.read_excel(archivo, sheet_name='1. Plan de Acción')
        columnas = df.columns.tolist()
        return columnas
    else:
        return []

def buscar_datos(programa, proyecto, componente, periodo):
    nombre_archivo = "VR. Académica.xlsx"  # Aquí se define el nombre del archivo a buscar
    archivos_excel = buscar_archivo(nombre_archivo, ruta_archivos)
    resultados = []
    for archivo in archivos_excel:
        df = pd.read_excel(archivo, sheet_name='1. Plan de Acción')
        coincidencias = df[(df['Programa'] == programa) & (df['Proyecto'] == proyecto) & (df['Componente'] == componente) & (df['Periodo'] == periodo)]
        if not coincidencias.empty:
            resultado = {
                'Archivo': archivo,
                'Coincidencias': coincidencias.to_dict('records')
            }
            resultados.append(resultado)
    return resultados

def buscar():
    programa = combo_programa.get()
    proyecto = combo_proyecto.get()
    componente = combo_componente.get()
    periodo = combo_periodo.get()
    
    resultados = buscar_datos(programa, proyecto, componente, periodo)
    if resultados:
        # Crear una nueva ventana
        resultados_window = tk.Toplevel(window)
        resultados_window.title("Resultados de búsqueda")

        # Crear un Treeview para mostrar los resultados como tabla
        treeview = ttk.Treeview(resultados_window)

        # Configurar las columnas
        treeview["columns"] = ("Archivo", "Coincidencias")
        treeview.heading("Archivo", text="Archivo")
        treeview.heading("Coincidencias", text="Coincidencias")

        # Agregar los datos al Treeview
        for resultado in resultados:
            archivo = resultado["Archivo"]
            coincidencias = resultado["Coincidencias"]
            # Insertar una nueva fila por cada coincidencia
            for coincidencia in coincidencias:
                treeview.insert("", "end", values=(archivo, coincidencia))

        # Ajustar las columnas al contenido
        for column in treeview["columns"]:
            treeview.column(column, width=100, anchor="center")

        # Mostrar el Treeview
        treeview.pack(fill="both", expand=True)
    else:
        messagebox.showinfo("Resultados de búsqueda", "No se encontraron coincidencias")

# Configuración de la ventana
window = tk.Tk()
window.title("Buscador de datos en archivos Excel")
window.geometry("400x300")

# Etiquetas y listas desplegables
label_programa = tk.Label(window, text="Programa:")
label_programa.pack()
combo_programa = ttk.Combobox(window, state="readonly", values=cargar_valores_columnas())
combo_programa.pack()

label_proyecto = tk.Label(window, text="Proyecto:")
label_proyecto.pack()
combo_proyecto = ttk.Combobox(window, state="readonly", values=cargar_valores_columnas())
combo_proyecto.pack()

label_componente = tk.Label(window, text="Componente:")
label_componente.pack()
combo_componente = ttk.Combobox(window, state="readonly", values=cargar_valores_columnas())
combo_componente.pack()

label_periodo = tk.Label(window, text="Periodo:")
label_periodo.pack()
combo_periodo = ttk.Combobox(window, state="readonly", values=cargar_valores_columnas())
combo_periodo.pack()

# Botón de búsqueda
search_button = tk.Button(window, text="Buscar", command=buscar)
search_button.pack()

# Ruta de los archivos de Excel
ruta_archivos = r"\\RIDER\DocInterno\PEI - PDI 2022 - 2031\PLANES DE ACCIÓN 2023\1. VR. Académica"

# Iniciar la aplicación
window.mainloop()
