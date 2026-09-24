import pandas as pd
import numpy as np
import openpyxl
import os
from pathlib import Path
from openpyxl.utils.dataframe import dataframe_to_rows
import streamlit as st

st.title("Separacion de Inventarios Cotrapsa")

inventario_plataforma = st.file_uploader("Seleccionar Inventario de la Plataforma", type=["xlsx","xls"])
inventario_sae = st.file_uploader("Seleccionar Inventario de SAE", type=["xlsx","xls"])

if inventario_plataforma is not None and inventario_sae is not None:
    # Leer el archivo Excel del inventario de la plataforma e Inventario SAE
    df_inventario_plataforma = pd.read_excel(inventario_plataforma)
    df_inventario_sae = pd.read_excel(inventario_sae)

    st.write("### Inventario de la plataforma")
    st.dataframe(df_inventario_plataforma)

    st.write("### Inventario SAE")
    st.dataframe(df_inventario_sae)

    # Procesamiento para separacion de Inventario por Base
    if st.button("Separacion de inventarios"):
        # Quitamos Numeros de parte de servicios del inventario de la plataforma
        df_inventario_plataforma = df_inventario_plataforma[~df_inventario_plataforma['NUMERO DE PARTE'].isin(['001','002','003','004','009','010','012','013','015','016','018'])].reset_index(drop=True).copy()
        
        # Creamos un diccionario con las tablas de cada una de las bases
        bases = {}
        for i in df_inventario_plataforma['ALMACEN'].value_counts().keys()[:-1]:
            bases[f'base_{i}'] = df_inventario_plataforma[df_inventario_plataforma['ALMACEN']==f'{i}'].copy()

        #Creamos una lista con los nombres de cada base
        nombres_bases = list(df_inventario_plataforma['ALMACEN'].value_counts().index)

        # Creamos un diccionario para agregar la linea a cada producto
        dictio_linea = dict(zip([str(x) for x in df_inventario_sae['Clave ']], df_inventario_sae['Línea ']))

        # Creamos un archivo para cada BASE
        for i in nombres_bases[:-1]:
            df_base = bases[f'base_{i}']
            df_base['LINEA'] = [dictio_linea[str(x)] for x in list(df_base['NUMERO DE PARTE'])]
            df_base['NO. DE PARTE STRING'] = [str(x) for x in list(df_base['NUMERO DE PARTE'])]
            df_base.sort_values(['LINEA', 'NUMERO DE PARTE'], inplace=True)
            df_base = df_base[['NUMERO DE PARTE', 'DESCRIPCION', 'EXISTENCIA']].copy()

            st.success("¡Procesamiento completado!")
            st.write(f"### Inventario de la Base {i}:")
            st.dataframe(df_base)

            # Permitir descargar el resultado
            output_name = f"Inventario_{i}.xlsx"
            df_base.to_excel(output_name, index=False)
            with open(output_name, "rb") as file:
                st.download_button(
                    label=f"Descargar Inventario_{i}",
                    data=file,
                    file_name=f"inventario_{i}.xlsx",
                    mime=f"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
