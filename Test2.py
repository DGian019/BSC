'''
import tkinter as tk
from tkinter import ttk
import pandas as pd

# Ruta del archivo Excel
archivo_excel = r"\\RIDER\DocInterno\PEI - PDI 2022 - 2031\PLANES DE ACCIÓN 2023\1. VR. Académica\VR. Académica.xlsm"
hoja = "2. Seguimiento"

# Leer la hoja "2. Seguimiento" del archivo Excel
df = pd.read_excel(archivo_excel, sheet_name=hoja)

# Obtener los valores únicos de la columna "G" (Componente)
proyectos = df.iloc[3:, 6].dropna().unique().tolist()

def obtener_valores_unicos_columnas_combinadas():
    columnas_combinadas = ['G']
    valores_columnas = []

    for columna in columnas_combinadas:
        valores = df[columna].dropna().unique().tolist()
        valores_columnas.append(valores)

    return valores_columnas

def mostrar_datos(event):
    # Obtener el valor seleccionado de la lista desplegable
    proyecto_seleccionado = combo_proyecto.get()

    # Filtrar el DataFrame por el valor seleccionado en la columna "G"
    datos_filtrados = df[df.iloc[:, 6] == proyecto_seleccionado]

    # Limpiar el Treeview
    treeview.delete(*treeview.get_children())

    # Contador de filas
    contador_filas = 0

    # Agregar los datos filtrados al Treeview
    for _, row in datos_filtrados.iterrows():
        valores = row.tolist()[1:]
        # Insertar una nueva fila en el Treeview con los valores de la fila
        treeview.insert("", "end", text=str(contador_filas), values=valores, tags=("datos",))
        contador_filas += 1

# Crear la ventana
window = tk.Tk()
window.title("Tabla de datos Prueba")

# Crear la etiqueta y la lista desplegable
label_proyecto = tk.Label(window, text="Seleccion de Componentes:")
label_proyecto.pack()
combo_proyecto = ttk.Combobox(window, state="readonly", values=proyectos)
combo_proyecto.pack()
combo_proyecto.bind("<<ComboboxSelected>>", mostrar_datos)

# Crear el Treeview para mostrar la tabla
treeview = ttk.Treeview(window)

# Configurar las columnas con nombres personalizados
nombres_columnas = ["Perpectiva del BSC", "Objetivo Estrategico", "Politica", "Programa", "Proyecto", "Codigo de Componente", "Componentes", "Indicador", "META", "Macro Actvidades", "Actividad", "Porcentaje de aporte al componente", "Entregables", "Codigo del Componente", "Periodo de Ejecucion", "Porcentaje de avance de aporte de la actividad al compenete en el Año"]  # Reemplaza con los nombres deseados
treeview["columns"] = nombres_columnas
treeview.heading("#0", text="Índice")
for columna, nombre_columna in zip(range(len(nombres_columnas)), nombres_columnas):
    treeview.heading(columna, text=nombre_columna)

# Mostrar el Treeview vacío inicialmente
treeview.pack(fill="both", expand=True)

# Iniciar la aplicación
window.mainloop()

'''

'''import tkinter as tk
from tkinter import ttk
import pandas as pd

# Ruta del archivo Excel
archivo_excel = r"\\RIDER\DocInterno\PEI - PDI 2022 - 2031\PLANES DE ACCIÓN 2023\1. VR. Académica\VR. Académica.xlsm"
hoja = "2. Seguimiento"

# Leer la hoja "2. Seguimiento" del archivo Excel
df = pd.read_excel(archivo_excel, sheet_name=hoja)

# Obtener los valores únicos de la columna "F" (Proyecto)
proyectos = df.iloc[3:, 5].dropna().unique().tolist()

def mostrar_datos(event):
    # Obtener el valor seleccionado de la lista desplegable
    proyecto_seleccionado = combo_proyecto.get()

    # Filtrar el DataFrame por el valor seleccionado en la columna "F"
    datos_filtrados = df[df.iloc[:, 5] == proyecto_seleccionado]

    # Limpiar el Treeview
    treeview.delete(*treeview.get_children())

    # Agregar los datos filtrados al Treeview
    for i, row in datos_filtrados.iterrows():
        treeview.insert("", "end", text=str(i), values=row.tolist()[1:])

# Crear la ventana
window = tk.Tk()
window.title("Tabla de datos Prueba")

# Crear la etiqueta y la lista desplegable
label_proyecto = tk.Label(window, text="Seleccion de Proyecto:")
label_proyecto.pack()
combo_proyecto = ttk.Combobox(window, state="readonly", values=proyectos)
combo_proyecto.pack()
combo_proyecto.bind("<<ComboboxSelected>>", mostrar_datos)

# Crear el Treeview para mostrar la tabla
treeview = ttk.Treeview(window)

# Configurar las columnas
columnas = df.columns.tolist()[1:]
treeview["columns"] = columnas
treeview.heading("#0", text="Índice")
for columna in columnas:
    treeview.heading(columna, text=columna)

# Mostrar el Treeview vacío inicialmente
treeview.pack(fill="both", expand=True)

# Iniciar la aplicación
window.mainloop()'''
