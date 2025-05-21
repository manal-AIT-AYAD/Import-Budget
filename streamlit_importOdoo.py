import streamlit as st
import os
import tempfile
from excel_importOddo import transform_budget_data_append_sheet  # Assure-toi que ce nom est correct

st.set_page_config(
    page_title="Ajout Feuille Import Odoo",
    layout="centered"
)

st.title("🔄 Ajout de la feuille 'Import Odoo'")
st.markdown("Cet outil ajoute une feuille **'Import Odoo'** dans un ou plusieurs fichiers Excel contenant un budget.")

uploaded_files = st.file_uploader(
    "📤 Téléversez un ou plusieurs fichiers Excel (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    temp_paths = []
    try:
        # Enregistrement temporaire des fichiers téléversés
        for uploaded_file in uploaded_files:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
                temp_file.write(uploaded_file.getvalue())
                temp_paths.append(temp_file.name)

        # Le premier fichier est utilisé comme fichier de base pour l'ajout
        output_filename = f"odoo_import_{uploaded_files[0].name}"
        output_path = os.path.join(tempfile.gettempdir(), output_filename)

        # Copie du premier fichier pour édition
        with open(temp_paths[0], "rb") as src, open(output_path, "wb") as dst:
            dst.write(src.read())

        with st.spinner("⏳ Création de la feuille d'import Odoo en cours..."):
            transform_budget_data_append_sheet(temp_paths, output_path)

        st.success("✅ Traitement terminé avec succès !")

        # Bouton de téléchargement
        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Télécharger le fichier modifié",
                data=f,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Une erreur est survenue : {str(e)}")

    finally:
        try:
            for path in temp_paths:
                os.remove(path)
            if 'output_path' in locals():
                os.remove(output_path)
        except:
            pass
