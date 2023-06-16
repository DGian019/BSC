import tkinter as tk
from tkinter import ttk
import pandas as pd

# Ruta del archivo Excel
archivo_excel = r"\\RIDER\DocInterno\Planeacion\1. 2023\1. PDI\2. POAI 2023.xlsx"
hoja = "POAI 2023"

# Leer la hoja "POAI 2023" del archivo Excel
df = pd.read_excel(archivo_excel, sheet_name=hoja)

# Obtener los valores únicos de la columna "G" (Componente)
proyectos = df.iloc[3:, 6].dropna().unique().tolist()

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
        valores = [row[i] for i in indices_columnas_seleccionadas]  # Filtrar los valores de las columnas seleccionadas
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

# Definir las columnas seleccionadas y sus nombres
nombres_columnas = ["PERSPECTIVA DEL BSC","EJE","OBJETIVO ESTRATÉGICO", "POLITICA", "PROGRAMA ", "RESPONSABLE DE PROGRAMA", "PROYECTO", "RESPONSABLE DE PROYECTO", "CÓDIGO COMPONENTE", "COMPONENTES", "COMPONENTES Concatenados", "INDICADOR", "META", "MEGAS","MACROACTIVIDADES", "ACTIVIDAD", "RESPONSABLE COMPONENTE"]
indices_columnas_seleccionadas = [0, 2, 3, 4, 6, 7, 8, 9, 11, 12, 14, 15, 16]  # Índices de las columnas seleccionadas en el DataFrame

# Configurar las columnas del Treeview
treeview["columns"] = indices_columnas_seleccionadas
treeview.heading("#0", text="Índice")

# Configurar los nombres de las columnas seleccionadas
for indice, indice_columna in enumerate(indices_columnas_seleccionadas):
    nombre_columna = nombres_columnas[indice_columna]
    treeview.heading(indice_columna, text=f"{indice_columna}: {nombre_columna}")

# Mostrar el Treeview vacío inicialmente
treeview.pack(fill="both", expand=True)

# Iniciar la aplicación
window.mainloop()

