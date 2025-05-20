import streamlit as st
from excel_budget import process_budget_excel
import os

st.set_page_config(page_title="Import Budget Excel", layout="centered")

st.title("Import & Traitement Budget Excel")

uploaded_file = st.file_uploader("Téléversez un fichier Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with open("temp_uploaded.xlsx", "wb") as f:
        f.write(uploaded_file.read())

    with st.spinner("Traitement en cours..."):
        output_file = process_budget_excel("temp_uploaded.xlsx")
        st.success("Traitement terminé !")

        with open(output_file, "rb") as f:
            st.download_button(
                label="📥 Télécharger le fichier traité",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    os.remove("temp_uploaded.xlsx")
    os.remove(output_file)
