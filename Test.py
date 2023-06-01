import glob

def buscar_archivo(nombre_archivo, directorio):
    patron = f"{directorio}/**/{nombre_archivo}"
    archivos_encontrados = glob.glob(patron, recursive=True)
    return archivos_encontrados

nombre_archivo = "Líderes PDI.xlsx"
directorio = "C:\\Users\\mercadeo_ventas1\\Desktop"

archivos_encontrados = buscar_archivo(nombre_archivo, directorio)
if archivos_encontrados:
    print("Archivos encontrados:")
    for archivo in archivos_encontrados:
        print(archivo)
else:
    print("No se encontraron archivos con ese nombre.")
