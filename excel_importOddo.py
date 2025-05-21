import streamlit as st
import os
from tempfile import NamedTemporaryFile
from traitement_budget import transform_budget_data_append_sheet

st.set_page_config(page_title="Traitement Budget Odoo", layout="centered")

st.title("Importation & Traitement du Budget pour Odoo")

# Upload du fichier contenant les données de budget
uploaded_source_file = st.file_uploader("📥 Fichier source (budget à transformer)", type=["xlsx"], key="source")

# Upload du fichier existant dans lequel ajouter la nouvelle feuille
uploaded_existing_file = st.file_uploader("📤 Fichier Excel existant (recevra les données)", type=["xlsx"], key="target")

if uploaded_source_file and uploaded_existing_file:
    with NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_source:
        tmp_source.write(uploaded_source_file.read())
        tmp_source_path = tmp_source.name

    with NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_existing:
        tmp_existing.write(uploaded_existing_file.read())
        tmp_existing_path = tmp_existing.name

    if st.button("🚀 Lancer le traitement"):
        with st.spinner("Traitement en cours..."):
            try:
                transform_budget_data_append_sheet(
                    input_files=[tmp_source_path],
                    existing_file=tmp_existing_path,
                    new_sheet_name="Import Odoo"
                )
                with open(tmp_existing_path, "rb") as f:
                    st.success("✅ Traitement terminé ! Vous pouvez télécharger le fichier :")
                    st.download_button(
                        label="📥 Télécharger le fichier traité",
                        data=f,
                        file_name="budget_import_odoo.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"❌ Une erreur s'est produite : {e}")
            finally:
                os.remove(tmp_source_path)
                os.remove(tmp_existing_path)
