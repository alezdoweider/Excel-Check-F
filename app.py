import streamlit as st
import pandas as pd

st.title("Consulta de BlueStars - Filtrar por CASO")

# Subir archivo Excel (xlsm o xlsx)
uploaded_file = st.file_uploader("Selecciona el archivo BlueStars", type=["xlsm", "xlsx"])

if uploaded_file is not None:
    try:
        # 1. Leer el archivo Excel y la hoja "ARMADRE"
        df = pd.read_excel(uploaded_file, sheet_name="ARMADRE", engine="openpyxl")
        
        # 2. Extraer CASO y NUNC de la columna Q (índice 16: A=0,..., Q=16)
        df["CASO"] = df.iloc[:, 16].astype(str).str.split("-", n=1).str[0].str.strip()
        df["NUNC"] = df.iloc[:, 16].astype(str).str.split("-", n=1).str[1].str.strip()
        
        # 3. Extraer otras columnas:
        # - NOMBRE de la columna K (índice 10)
        # - ID EMP de la columna E (índice 4)
        # - Nro. ID de la columna F (índice 5)
        # - TIPO EMP de la columna H (índice 7)
        df["NOMBRE"]   = df.iloc[:, 10]
        df["ID EMP"]   = df.iloc[:, 4]
        df["Nro. ID"]  = df.iloc[:, 5]
        df["TIPO EMP"] = df.iloc[:, 7]
        
        # 4. Crear el DataFrame de consulta con las columnas requeridas
        columnas_interes = ["CASO", "NUNC", "NOMBRE", "ID EMP", "Nro. ID", "TIPO EMP"]
        df_procesado = df[columnas_interes].copy()
        
        # 5. Filtrar por CASO
        filtro_caso = st.text_input("Filtrar por CASO (contiene):")
        if filtro_caso:
            df_filtrado = df_procesado[df_procesado["CASO"].astype(str).str.contains(filtro_caso, case=False, na=False)]
            st.write("Resultados filtrados:")
            st.dataframe(df_filtrado)
        else:
            st.write("Datos procesados:")
            st.dataframe(df_procesado)
            
    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
else:
    st.info("Por favor, sube el archivo para visualizar la consulta.")
