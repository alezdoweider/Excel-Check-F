import streamlit as st
import pandas as pd
from datetime import datetime
from login import create_table, register_user, login_user
from io import BytesIO

# Crear tabla de usuarios al inicio
create_table()

# Interfaz de Login y Registro
st.title("Login / Registro")
menu = ["Login", "Registro"]
choice = st.sidebar.selectbox("Menú", menu)

if choice == "Login":
    st.subheader("Iniciar Sesión")
    usuario = st.text_input("Usuario")
    password = st.text_input("Contraseña", type='password')

    if st.button("Entrar"):
        result = login_user(usuario, password)
        if result:
            st.success(f"Bienvenido {usuario}")
            
            # Subir archivo Excel
            st.header("Buscar Caso en Excel")
            uploaded_file = st.file_uploader("Cargar archivo Excel (.xlsm, .xlsx)", type=["xlsm", "xlsx"])

            if uploaded_file:
                df = pd.read_excel(uploaded_file, sheet_name='ARMADARE')
                st.success("Archivo cargado exitosamente")
                
                st.write("Vista previa de los datos:")
                st.dataframe(df.head())

                # Procesamiento
                st.header("Procesamiento de Datos")
                fecha_actual = datetime.now().strftime("%d/%m/%Y")
                st.write(f"Fecha actual: {fecha_actual}")

                df['Numero_Caso'] = df['Q'].astype(str).str.split('-').str[0]
                df['NUNC'] = df['Q'].astype(str).str.split('-').str[1]
                df['ID'] = df['E'].astype(str).str.split('/').str[0]
                df['Numero_ID'] = df['E'].astype(str).str.extract(r'(\d+)')[0]
                df['Numero_Antes_Slash'] = df['F']
                df['Tipo_EMP'] = df['H']
                df['EMPS'] = df['K']

                tipo_envase = st.selectbox("Seleccione tipo de envase", ["TTG", "TTR", "TTL", "TTV", "FP", "BP"])
                df['Tipo_Envase'] = tipo_envase

                st.subheader("Resultados Procesados")
                columnas_mostrar = ['Numero_Caso', 'NUNC', 'ID', 'Numero_ID', 'Numero_Antes_Slash', 'Tipo_EMP', 'EMPS', 'Tipo_Envase']
                st.dataframe(df[columnas_mostrar].head())

                # Descargar archivo procesado
                output = BytesIO()
                df.to_excel(output, index=False, engine='openpyxl')
                st.download_button(
                    label="Descargar Excel Procesado",
                    data=output.getvalue(),
                    file_name="procesado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error("Usuario o contraseña incorrectos")

elif choice == "Registro":
    st.subheader("Crear Cuenta")
    usuario = st.text_input("Nuevo Usuario")
    cedula = st.text_input("Cédula")
    password = st.text_input("Nueva Contraseña", type='password')

    if st.button("Registrar"):
        if register_user(usuario, cedula, password):
            st.success("Usuario registrado exitosamente. Ahora puedes iniciar sesión.")
        else:
            st.error("Usuario ya existe. Intenta con otro.")
