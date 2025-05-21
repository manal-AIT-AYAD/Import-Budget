import streamlit as st
import os
from excel_budget import process_budget_excel
from excel_importOddo import transform_budget_data_append_sheet

# Configuration de la page
st.set_page_config(page_title="Budget Excel App", layout="centered")
st.title("📊 Application de Traitement & Import Budget Excel")

# Choix de l'action
option = st.radio("Que voulez-vous faire ?", [
    "1️⃣ Traiter un fichier Budget Excel",
    "2️⃣ Ajouter une feuille 'Import Odoo' à un fichier existant"
])

# 🔹 Option 1 : Traitement uniquement
if option == "1️⃣ Traiter un fichier Budget Excel":
    st.subheader("📥 Téléversez un fichier Excel à traiter")
    uploaded_file = st.file_uploader("Fichier Excel (.xlsx)", type=["xlsx"], key="upload_traitement")

    if uploaded_file:
        with open("temp_to_process.xlsx", "wb") as f:
            f.write(uploaded_file.read())

        with st.spinner("Traitement du fichier..."):
            output_file = process_budget_excel("temp_to_process.xlsx")

        st.success("✅ Traitement terminé.")
        with open(output_file, "rb") as f:
            st.download_button(
                label="📥 Télécharger le fichier traité",
                data=f,
                file_name="fichier_traite.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # Nettoyage
        os.remove("temp_to_process.xlsx")
        os.remove(output_file)

# 🔹 Option 2 : Ajout de la feuille "Import Odoo"
elif option == "2️⃣ Ajouter une feuille 'Import Odoo' à un fichier existant":
    st.subheader("📤 Téléversez le fichier dans lequel ajouter la feuille")
    existing_file = st.file_uploader("Fichier Excel existant", type=["xlsx"], key="upload_odoo_target")

    st.subheader("📂 Téléversez les fichiers de données budget à utiliser")
    source_files = st.file_uploader("Fichiers sources (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="upload_odoo_sources")

    if existing_file and source_files:
        # Sauvegarde du fichier existant
        existing_path = "uploaded_existing.xlsx"
        with open(existing_path, "wb") as f:
            f.write(existing_file.read())

        # Sauvegarde des fichiers source
        source_paths = []
        for i, file in enumerate(source_files):
            path = f"source_{i}.xlsx"
            with open(path, "wb") as f:
                f.write(file.read())
            source_paths.append(path)

        # Ajout de la feuille
        with st.spinner("Ajout de la feuille 'Import Odoo'..."):
            transform_budget_data_append_sheet(source_paths, existing_path)

        st.success("✅ Feuille ajoutée avec succès.")
        with open(existing_path, "rb") as f:
            st.download_button(
                label="📥 Télécharger le fichier modifié",
                data=f,
                file_name="fichier_avec_import_odoo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # Nettoyage
        os.remove(existing_path)
        for path in source_paths:
            os.remove(path)
