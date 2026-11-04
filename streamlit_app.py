import streamlit as st

st.title("🎈 Combinar Excel")
st.write(
    "Let's start building! For help and inspiration, head over to [docs.streamlit.io](https://docs.streamlit.io/)."
)

import io
import os
import re
import pandas as pd
import streamlit as st
import openpyxl


# Configuración inicial de la página
st.set_page_config(
    page_title="Combinador de Archivos Excel", page_icon="📊", layout="centered"
)

# Título y descripción principal
st.title("📊 Combinador de Hojas Excel")
st.write(
    "Carga múltiples archivos Excel para consolidar sus hojas en un único archivo de salida."
)


# Función para limpiar y adaptar los nombres de las pestañas a los límites de Excel
def sanitizar_nombre_hoja(nombre: str) -> str:
    # 1. Eliminar caracteres no permitidos por Excel en los nombres de pestaña (\ / ? * : [ ])
    nombre_limpio = re.sub(r"[\\/*?:\[\]]", "_", nombre)
    # 2. Recortar el nombre a un máximo de 31 caracteres (límite estricto de Excel)
    return nombre_limpio[:31]


# 1. SECCIÓN DE CARGA DE ARCHIVOS
uploaded_files = st.file_uploader(
    "Selecciona o arrastra los archivos Excel (.xlsx, .xls)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,  # Permite seleccionar múltiples archivos simultáneamente
)

# 2. CONFIRMACIÓN Y PREVISUALIZACIÓN DE ARCHIVOS CARGADOS
if uploaded_files:
    # Mensaje de confirmación de carga correcta
    st.success(
        f"✅ Se han cargado **{len(uploaded_files)}** archivo(s) correctamente."
    )

    # Mostrar la lista detallada de los archivos listos para procesar
    with st.expander("📄 Ver lista de archivos cargados"):
        for file in uploaded_files:
            st.write(f"- `{file.name}`")

    # 3. BOTÓN DE PROCESAMIENTO Y LÓGICA DE COMBINACIÓN
    if st.button("🚀 Procesar y combinar archivos", type="primary"):
        with st.spinner("Procesando los archivos Excel..."):

            # Crear un buffer en memoria RAM para construir el archivo de salida sin escribir en disco
            output_buffer = io.BytesIO()

            # Usar pd.ExcelWriter junto con openpyxl para escribir en el buffer en memoria
            with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:

                # Conjunto para rastrear nombres de hojas y evitar duplicados entre archivos distintos
                hojas_usadas = set()

                for file in uploaded_files:
                    # Extraer el nombre base del archivo omitiendo la extensión (.xlsx o .xls)
                    nombre_base, _ = os.path.splitext(file.name)

                    # Leer todas las hojas del archivo (sheet_name=None devuelve un diccionario: {nombre_hoja: DataFrame})
                    hojas_dict = pd.read_excel(file, sheet_name=None)
                    total_hojas = len(hojas_dict)

                    # Recorrer cada una de las hojas leídas del archivo actual
                    for idx, (nombre_hoja_orig, df) in enumerate(
                        hojas_dict.items(), start=1
                    ):

                        # Definir el nombre según la cantidad de hojas presentes
                        if total_hojas == 1:
                            # Si solo tiene una hoja, mantiene el nombre del archivo
                            nombre_propuesto = nombre_base
                        else:
                            # Si tiene varias hojas, asigna NombreArchivo_1, NombreArchivo_2, etc.
                            nombre_propuesto = f"{nombre_base}_{idx}"

                        # Limpiar caracteres prohibidos y ajustar a 31 caracteres máximo
                        nombre_final = sanitizar_nombre_hoja(nombre_propuesto)

                        # Controlar duplicados en caso de que existan dos archivos con nombres idénticos
                        contador = 1
                        nombre_unico = nombre_final
                        while nombre_unico in hojas_usadas:
                            sufijo = f"_{contador}"
                            limite = 31 - len(sufijo)
                            nombre_unico = f"{nombre_final[:limite]}{sufijo}"
                            contador += 1

                        # Registrar la hoja como procesada
                        hojas_usadas.add(nombre_unico)

                        # Escribir el contenido del DataFrame en la hoja correspondiente del archivo consolidado
                        df.to_excel(
                            writer, sheet_name=nombre_unico, index=False
                        )

            # Reposicionar el puntero de lectura del buffer al inicio
            output_buffer.seek(0)

            # 4. SALIDA Y BOTÓN DE DESCARGA
            st.success("🎉 ¡Procesamiento completado con éxito!")

            st.download_button(
                label="📥 Descargar Excel Combinado",
                data=output_buffer.getvalue(),
                file_name="Excel_Combinado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

else:
    st.info(
        "Por favor, sube uno o más archivos Excel para habilitar el procesamiento."
    )
