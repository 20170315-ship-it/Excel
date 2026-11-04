
import io
import os
import re
import pandas as pd
import streamlit as st

# Configuración inicial de la página
st.set_page_config(
    page_title="Combinador de Archivos Excel", page_icon="📊", layout="centered"
)

# Título y descripción principal
st.title("📊 Combinador de Hojas Excel")
st.write(
    "Carga múltiples archivos Excel para consolidar sus hojas en un único archivo de salida."
)


# Función para sanitizar y ajustar nombres de pestañas al estándar de Excel
def sanitizar_nombre_hoja(nombre: str) -> str:
    # Eliminar caracteres no permitidos por Excel (\ / ? * : [ ])
    nombre_limpio = re.sub(r"[\\/*?:\[\]]", "_", nombre)
    # Recortar al límite estricto de 31 caracteres de Excel
    return nombre_limpio[:31]


# 1. SECCIÓN DE CARGA DE ARCHIVOS
uploaded_files = st.file_uploader(
    "Selecciona o arrastra los archivos Excel (.xlsx, .xls)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

# 2. CONFIRMACIÓN Y PREVISUALIZACIÓN DE ARCHIVOS CARGADOS
if uploaded_files:
    st.success(
        f"✅ Se han cargado **{len(uploaded_files)}** archivo(s) correctamente."
    )

    with st.expander("📄 Ver lista de archivos cargados"):
        for file in uploaded_files:
            st.write(f"- `{file.name}`")

    # 3. BOTÓN DE PROCESAMIENTO Y COMBINACIÓN DE HOJAS
    if st.button("🚀 Procesar y combinar archivos", type="primary"):
        with st.spinner("Procesando los archivos Excel..."):

            # Buffer en memoria RAM para construir el archivo final
            output_buffer = io.BytesIO()

            # Escritor de pandas especificando el motor openpyxl
            with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:

                # Control para prevenir nombres de hojas duplicados
                hojas_usadas = set()

                for file in uploaded_files:
                    # Extraer el nombre base del archivo sin extensión
                    nombre_base, _ = os.path.splitext(file.name)

                    # Leer todas las hojas del archivo actual
                    hojas_dict = pd.read_excel(file, sheet_name=None)
                    total_hojas = len(hojas_dict)

                    for idx, (nombre_hoja_orig, df) in enumerate(
                        hojas_dict.items(), start=1
                    ):

                        # Regla de renombrado:
                        # Si tiene 1 sola hoja: NombreArchivo
                        # Si tiene varias hojas: NombreArchivo_1, NombreArchivo_2, etc.
                        if total_hojas == 1:
                            nombre_propuesto = nombre_base
                        else:
                            nombre_propuesto = f"{nombre_base}_{idx}"

                        # Aplicar limpieza de caracteres y corte de 31 caracteres
                        nombre_final = sanitizar_nombre_hoja(nombre_propuesto)

                        # Evitar colisión si dos archivos tienen el mismo nombre
                        contador = 1
                        nombre_unico = nombre_final
                        while nombre_unico in hojas_usadas:
                            sufijo = f"_{contador}"
                            limite = 31 - len(sufijo)
                            nombre_unico = f"{nombre_final[:limite]}{sufijo}"
                            contador += 1

                        hojas_usadas.add(nombre_unico)

                        # Escribir la hoja procesada en el archivo final
                        df.to_excel(
                            writer, sheet_name=nombre_unico, index=False
                        )

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
