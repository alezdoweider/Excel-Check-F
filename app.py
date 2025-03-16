import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
from io import BytesIO

st.title("Procesamiento de BlueStars")

# Widget para cargar el archivo
uploaded_file = st.file_uploader("Selecciona el archivo BlueStars (xlsm o xlsx)", type=["xlsm", "xlsx"])

if uploaded_file is not None:
    try:
        # 1. Leer el archivo Excel y la hoja ARMADRE
        df = pd.read_excel(uploaded_file, sheet_name="ARMADRE", engine="openpyxl")
        
        # 2. Extraer CASO y NUNC de la columna Q (índice 16, ya que A=0,..., Q=16)
        df["CASO"] = df.iloc[:, 16].astype(str).str.split("-", n=1).str[0].str.strip()
        df["NUNC"] = df.iloc[:, 16].astype(str).str.split("-", n=1).str[1].str.strip()
        
        # 3. Extraer columnas NOMBRE (columna K, índice 10), ID EMP (columna E, índice 4),
        # Nro. ID (columna F, índice 5) y TIPO EMP (columna H, índice 7)
        df["NOMBRE"]   = df.iloc[:, 10]
        df["ID EMP"]   = df.iloc[:, 4]
        df["Nro. ID"]  = df.iloc[:, 5]
        df["TIPO EMP"] = df.iloc[:, 7]
        
        # 4. Crear un nuevo DataFrame con las columnas requeridas
        columnas_interes = ["CASO", "NUNC", "NOMBRE", "ID EMP", "Nro. ID", "TIPO EMP"]
        df_procesado = df[columnas_interes].copy()
        
        # 5. Guardar en un archivo Excel en memoria
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_procesado.to_excel(writer, sheet_name="Procesado", index=False)
        output.seek(0)
        
        # 6. Aplicar formato con openpyxl
        workbook = openpyxl.load_workbook(output)
        worksheet = workbook["Procesado"]
        
        # Definir estilo: fondo azul celeste (ADD8E6) y texto negro (000000)
        fill_color = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
        font_color = Font(color="000000")
        
        # Aplicar el formato a todas las celdas
        for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, max_col=worksheet.max_column):
            for cell in row:
                cell.fill = fill_color
                cell.font = font_color
        
        # Ajustar el ancho de la columna "Nro. ID" a 30 (en el nuevo archivo, "Nro. ID" es la columna E)
        worksheet.column_dimensions['E'].width = 30
        
        # Guardar el libro formateado en un nuevo BytesIO
        output_formatted = BytesIO()
        workbook.save(output_formatted)
        output_formatted.seek(0)
        
        # Botón para descargar el archivo procesado
        st.download_button(
            label="Descargar archivo procesado",
            data=output_formatted,
            file_name="BlueStars_Procesado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.success("El archivo ha sido procesado exitosamente.")
    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
else:
    st.info("Por favor, sube el archivo para procesarlo.")
