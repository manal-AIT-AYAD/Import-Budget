import streamlit as st
from excel_budget import process_budget_excel
from transform_append import transform_budget_data_append_sheet
import os

st.set_page_config(page_title="🧮 Budget Tools", layout="centered")
st.title("🧾 Outils de traitement des fichiers Budget")

option = st.radio("Choisissez une opération :", ["🧠 Traitement classique", "📎 Ajout d'une feuille dans un Excel existant"])

if option == "🧠 Traitement classique":
    st.header("📊 Import & Traitement Budget Excel")
    uploaded_file = st.file_uploader("Téléversez un fichier Excel (.xlsx)", type=["xlsx"], key="classique")

    if uploaded_file is not None:
        with open("temp_uploaded.xlsx", "wb") as f:
            f.write(uploaded_file.read())

        with st.spinner("📈 Traitement en cours..."):
            output_file = process_budget_excel("temp_uploaded.xlsx")
            st.success("✅ Traitement terminé !")

            with open(output_file, "rb") as f:
                st.download_button(
                    label="📥 Télécharger le fichier traité",
                    data=f,
                    file_name=output_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        os.remove("temp_uploaded.xlsx")
        os.remove(output_file)

elif option == "📎 Ajout d'une feuille dans un Excel existant":
    st.header("📎 Ajouter une feuille 'Import Odoo' dans un fichier Excel")

    uploaded_files = st.file_uploader(
        "Sélectionnez un ou plusieurs fichiers Excel contenant les données à transformer",
        type=["xlsx"],
        accept_multiple_files=True,
        key="multi_files"
    )

    existing_file = st.file_uploader(
        "Sélectionnez le fichier Excel existant où ajouter la feuille",
        type=["xlsx"],
        key="existing"
    )

    if uploaded_files and existing_file:
        temp_inputs = []
        for i, file in enumerate(uploaded_files):
            temp_name = f"temp_input_{i}.xlsx"
            with open(temp_name, "wb") as f:
                f.write(file.read())
            temp_inputs.append(temp_name)

        temp_existing = "temp_existing.xlsx"
        with open(temp_existing, "wb") as f:
            f.write(existing_file.read())

        with st.spinner("🔧 Traitement en cours..."):
            transform_budget_data_append_sheet(temp_inputs, temp_existing)
            st.success("✅ Feuille ajoutée avec succès !")

            with open(temp_existing, "rb") as f:
                st.download_button(
                    label="📥 Télécharger le fichier modifié",
                    data=f,
                    file_name="fichier_modifié.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        for f in temp_inputs:
            os.remove(f)
        os.remove(temp_existing)
