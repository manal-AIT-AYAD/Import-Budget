import streamlit as st
import os
from excel_budget import process_budget_excel
from excel_importOddo import transform_budget_data_append_sheet

# Configuration de la page
st.set_page_config(page_title="Budget Excel App", layout="centered")
st.title(" Application de Traitement Budget Excel")

# Choix de l'action
option = st.radio("Sélectionnez une action :", [
    " Traiter un fichier Budget Excel",
    "Ajouter une feuille 'Import Odoo' à un fichier existant"
])

# === OPTION 1 : Traitement simple ===
if option == " Traiter un fichier Budget Excel":
    st.subheader("Téléversez le fichier à traiter")
    uploaded_file = st.file_uploader("Fichier Excel (.xlsx)", type=["xlsx"], key="traitement_file")

    if uploaded_file:
        with open("temp_uploaded.xlsx", "wb") as f:
            f.write(uploaded_file.read())

        with st.spinner("Traitement en cours..."):
            output_file = process_budget_excel("temp_uploaded.xlsx")

        st.success(" Traitement terminé.")
        with open(output_file, "rb") as f:
            st.download_button(
                label=" Télécharger le fichier traité",
                data=f,
                file_name="compte_de_resultats_budget_travaille.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        os.remove("temp_uploaded.xlsx")
        os.remove(output_file)

# === OPTION 2 : Ajout feuille "Import Odoo" ===
elif option == " Ajouter une feuille 'Import Odoo' à un fichier existant":
    st.subheader("Téléversez les fichiers Excel contenant les données sources")
    uploaded_sources = st.file_uploader("Un ou plusieurs fichiers de données budget", type=["xlsx"], accept_multiple_files=True)

    st.subheader("Téléversez le fichier Excel existant dans lequel ajouter la feuille")
    uploaded_existing = st.file_uploader("Fichier cible existant", type=["xlsx"], key="existing_file")

    if uploaded_sources and uploaded_existing:
        # Sauvegarder les fichiers temporairement
        source_paths = []
        for i, file in enumerate(uploaded_sources):
            path = f"source_{i}.xlsx"
            with open(path, "wb") as f:
                f.write(file.read())
            source_paths.append(path)

        existing_path = "existing_file.xlsx"
        with open(existing_path, "wb") as f:
            f.write(uploaded_existing.read())

        with st.spinner("Ajout de la feuille 'Import Odoo' en cours..."):
            transform_budget_data_append_sheet(source_paths, existing_path)
        
        st.success("Feuille 'Import Odoo' ajoutée avec succès.")
        with open(existing_path, "rb") as f:
            st.download_button(
                label="📥 Télécharger le fichier modifié",
                data=f,
                file_name="budget_avec_import_odoo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # Nettoyage
        os.remove(existing_path)
        for path in source_paths:
            os.remove(path)
