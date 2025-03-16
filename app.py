import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# Iniciar y ocultar la ventana principal de tkinter
root = tk.Tk()
root.withdraw()

# Abrir diálogo para seleccionar el archivo Excel
file_path = filedialog.askopenfilename(
    title="Selecciona el archivo BlueStars",
    filetypes=[("Excel files", "*.xlsm *.xlsx")]
)

# Verificar si se seleccionó un archivo
if not file_path:
    print("No se seleccionó ningún archivo.")
    exit()

# 1. Leer el archivo Excel y la hoja ARMADRE
df = pd.read_excel(file_path, sheet_name="ARMADRE", engine="openpyxl")

# 2. Extraer CASO y NUNC de la columna Q (antes y después del guion "-")
df["CASO"] = df.iloc[:, 16].astype(str).str.split("-", n=1).str[0].str.strip()
df["NUNC"] = df.iloc[:, 16].astype(str).str.split("-", n=1).str[1].str.strip()

# 3. Extraer columnas NOMBRE (K), ID EMP (E), Nro. ID (F), TIPO EMP (H)
df["NOMBRE"]   = df.iloc[:, 10]
df["ID EMP"]   = df.iloc[:, 4]
df["Nro. ID"]  = df.iloc[:, 5]
df["TIPO EMP"] = df.iloc[:, 7]

# 4. Crear un nuevo DataFrame con las columnas requeridas
columnas_interes = ["CASO", "NUNC", "NOMBRE", "ID EMP", "Nro. ID", "TIPO EMP"]
df_procesado = df[columnas_interes].copy()

# 5. Guardar en un nuevo archivo Excel sin el índice
nuevo_archivo = "BlueStars_Procesado.xlsx"
with pd.ExcelWriter(nuevo_archivo, engine="openpyxl") as writer:
    df_procesado.to_excel(writer, sheet_name="Procesado", index=False)

# 6. Aplicar formato con openpyxl
workbook = load_workbook(nuevo_archivo)
worksheet = workbook["Procesado"]

# Definir estilos: fondo azul celeste y letras negras
fill_color = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
font_color = Font(color="000000")

# Aplicar formato a cada celda con datos
for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, max_col=worksheet.max_column):
    for cell in row:
        cell.fill = fill_color
        cell.font = font_color

# Ajustar el ancho de la columna "Nro. ID" a 30 (asumiendo que es la columna E en el nuevo archivo)
worksheet.column_dimensions['E'].width = 30

# Guardar los cambios en el archivo Excel
workbook.save(nuevo_archivo)
print("Archivo procesado y guardado como", nuevo_archivo)
