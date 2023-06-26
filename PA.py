import tkinter as tk
from ttkthemes import ThemedTk
from tkinter import ttk
import pandas as pd

# Ruta del archivo Excel
archivo_excel = r"\\RIDER\DocInterno\Planeacion\1. 2023\1. PDI\2. POAI 2023.xlsx"
hoja = "POAI 2023"

# Leer la hoja "POAI 2023" del archivo Excel
df = pd.read_excel(archivo_excel, sheet_name=hoja)

# Obtener los valores únicos de la columna "G" (Componente)
proyectos = df.iloc[3:, 6].dropna().unique().tolist()

# Obtener los códigos de componente únicos
codigos_componente = df.iloc[3:, 8].dropna().unique().tolist()

# Ruta del nuevo archivo Excel
nuevo_archivo_excel = r"\\RIDER\DocInterno\PEI - PDI 2022 - 2031\PLANES DE ACCIÓN 2023\12. Dir. Rec Físicos e Infr\12. Dir. Rec Físicos e Infr.xlsm"
nueva_hoja = "2. Seguimiento"

# Leer la hoja especificada del nuevo archivo Excel
df_nuevo = pd.read_excel(nuevo_archivo_excel, sheet_name=nueva_hoja)

# Obtener las columnas necesarias del nuevo archivo
columna1_nuevo = df_nuevo.iloc[:, 15].unique()  # Columna "P" (índice 15)
columna2_nuevo = df_nuevo.iloc[:, 16].unique()  # Columna "Q" (índice 16)
columna3_nuevo = df_nuevo.iloc[:, 17].unique()  # Columna "R" (índice 17)

proyectos_nuevo = df_nuevo.iloc[:, 5].dropna().unique().tolist()

def mostrar_datos(event):
    # Obtener el valor seleccionado de la lista desplegable (proyecto seleccionado)
    proyecto_seleccionado = combo_proyecto.get()

    # Filtrar el DataFrame original por el valor seleccionado en la columna "G" (proyecto seleccionado)
    datos_filtrados = df[df.iloc[:, 6] == proyecto_seleccionado]

    # Obtener los códigos de componente únicos correspondientes al proyecto seleccionado
    codigos_componente_seleccionados = datos_filtrados.iloc[:, 8].dropna().unique().tolist()

    # Actualizar los valores de la lista desplegable de código de componente
    combo_componente["values"] = codigos_componente_seleccionados

    # Limpiar el Treeview
    treeview.delete(*treeview.get_children())

    # Contador de filas
    contador_filas = 0

    # Agregar los datos filtrados del DataFrame POAI 2023 al Treeview
    for _, row in datos_filtrados.iterrows():
        valores = [row[i] for i in indices_columnas_seleccionadas]

        if proyecto_seleccionado in proyectos_nuevo:
            # Filtrar el DataFrame nuevo por el proyecto seleccionado en la columna "F" (POAI2023)
            datos_nuevo_filtrados = df_nuevo[df_nuevo.iloc[:, 5] == proyecto_seleccionado]

            for _, nuevo_row in datos_nuevo_filtrados.iterrows():
                valores_nuevo = [nuevo_row[i] for i in indices_columnas_nuevo]

                # Agregar los valores de las columnas seleccionadas del archivo nuevo
                valores.extend(valores_nuevo)

        # Insertar la fila en el Treeview
        treeview.insert("", "end", text=str(contador_filas), values=valores, tags=("datos",))
        contador_filas += 1

# Crear la ventana
window = ThemedTk()
window.title("Tabla de datos Prueba")

# Crear el frame principal
frame = ttk.Frame(window)
frame.pack(padx=30, pady=30)

# Crear la etiqueta y la lista desplegable de proyectos
label_proyecto = tk.Label(frame, text="Seleccion de Componentes:")
label_proyecto.pack(side="left")
combo_proyecto = ttk.Combobox(frame, state="readonly", values=proyectos)
combo_proyecto.pack(side="left")
combo_proyecto.bind("<<ComboboxSelected>>", mostrar_datos)

# Crear la etiqueta y la lista desplegable de códigos de componente
label_componente = tk.Label(frame, text="Código de Componente:")
label_componente.pack(side="left")
combo_componente = ttk.Combobox(frame, state="readonly", values=codigos_componente)
combo_componente.pack(side="left")

# Crear el Treeview para mostrar la tabla
treeview = ttk.Treeview(window)

# Definir las columnas seleccionadas y sus nombres
nombres_columnas = ["PERSPECTIVA DEL BSC", "EJE", "OBJETIVO ESTRATÉGICO", "POLITICA", "PROGRAMA ", "RESPONSABLE DE PROGRAMA", "PROYECTO", "RESPONSABLE DE PROYECTO", "CÓDIGO COMPONENTE", "COMPONENTES", "COMPONENTES Concatenados", "INDICADOR", "META", "MEGAS", "MACROACTIVIDADES", "ACTIVIDAD", "RESPONSABLE COMPONENTE"]
indices_columnas_seleccionadas = [0, 2, 3, 4, 6, 7, 8, 9, 11, 12, 14, 15, 16]  # Índices de las columnas seleccionadas en el DataFrame

# Definir los nombres de las nuevas columnas seleccionadas y sus índices
nombres_columnas_nuevo = ["PERIODO DE EJECUCIÓN","% AVANCE EN EL APORTE DE LA ACTIVIDAD AL COMPONENTE EN EL AÑO","% DE CUMPLIMIENTO DEL INDICADOR DE GESTIÓN"]
indices_columnas_nuevo = [15, 16, 17]  # Índices de las columnas seleccionadas en el DataFrame nuevo

# Configurar las columnas del Treeview
treeview["columns"] = indices_columnas_seleccionadas + indices_columnas_nuevo
treeview.heading("#0", text="Índice")

# Configurar los nombres de las columnas seleccionadas del archivo original
for indice, indice_columna in enumerate(indices_columnas_seleccionadas):
    nombre_columna = nombres_columnas[indice_columna]
    treeview.heading(indice_columna, text=f"{indice_columna}: {nombre_columna}")

# Configurar los nombres de las nuevas columnas
for indice, indice_columna in enumerate(indices_columnas_nuevo):
    nombre_columna = nombres_columnas_nuevo[indice]
    treeview.heading(indice_columna, text=f"{indice_columna}: {nombre_columna}")

# Mostrar el Treeview vacío inicialmente
treeview.pack(fill="both", expand=True)

# Iniciar la aplicación
window.mainloop()
