import tkinter as tk
from tkinter import ttk
import pandas as pd

# Ruta del archivo Excel
archivo_excel = r"\\RIDER\DocInterno\PEI - PDI 2022 - 2031\PLANES DE ACCIÓN 2023\1. VR. Académica\VR. Académica.xlsm"
hoja = "2. Seguimiento"

# Leer la hoja "2. Seguimiento" del archivo Excel y obtener los datos de las columnas B3 a Q174
df = pd.read_excel(archivo_excel, sheet_name=hoja, header=None, skiprows=2, usecols="B:Q", nrows=172)

# Crear la ventana
window = tk.Tk()
window.title("Tabla de datos")

# Crear el Treeview para mostrar la tabla
treeview = ttk.Treeview(window)

# Configurar las columnas
columnas = df.columns.tolist()
treeview["columns"] = columnas
treeview.heading("#0", text="Índice")
for columna in columnas:
    treeview.heading(columna, text=columna)

# Agregar los datos al Treeview
for i, row in df.iterrows():
    treeview.insert("", "end", text=str(i), values=row.tolist())

# Ajustar las columnas al contenido
for columna in columnas:
    treeview.column(columna, width=100, anchor="center")

# Mostrar el Treeview
treeview.pack(fill="both", expand=True)

# Iniciar la aplicación
window.mainloop()
