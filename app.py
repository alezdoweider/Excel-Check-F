import streamlit as st
import pandas as pd
import subprocess
import sys
from io import BytesIO

# Verificar e instalar dependencias automáticamente
def instalar_dependencias():
    paquetes = ["pandas", "openpyxl", "streamlit"]
    for paquete in paquetes:
        try:
            __import__(paquete)
        except ImportError:
            subprocess.check_call([sys.executable, "-m", "pip", "install", paquete])

instalar_dependencias()

def cargar_archivo(archivo):
    global df, casos_disponibles
    df = pd.read_excel(archivo, sheet_name="MATRIZ")
    df.iloc[:, 16] = df.iloc[:, 16].astype(str).str.split(' - ').str[0]
    df.iloc[:, 16] = pd.to_numeric(df.iloc[:, 16], errors='coerce')
    df = df.dropna(subset=[df.columns[16]])
    casos_disponibles = sorted(df.iloc[:, 16].astype(int).unique().tolist())
    return casos_disponibles

def buscar_caso(caso):
    try:
        caso = int(caso)
        fila = df[df.iloc[:, 16] == caso]
        
        if fila.empty:
            return "Caso no encontrado"
        
        valor_k = fila.iloc[0, 10]
        valor_h = fila.iloc[0, 7]
        valor_e = fila.iloc[0, 4]
        
        return f"K: {valor_k}\nH: {valor_h}\nE: {valor_e}"
    except ValueError:
        return "Seleccione un número válido"

st.title("Buscar Caso en Excel")
archivo = st.file_uploader("Cargar archivo Excel", type=["xlsm", "xlsx", "xls"])

if archivo is not None:
    casos_disponibles = cargar_archivo(archivo)
    caso_seleccionado = st.selectbox("Seleccione un número de caso", casos_disponibles)
    if st.button("Buscar"):
        resultado = buscar_caso(caso_seleccionado)
        st.text_area("Resultado", resultado, height=100)
