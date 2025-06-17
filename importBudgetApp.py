import streamlit as st
import os
import tempfile

from excel_importOddo import transform_budget_data_append_sheet
from excel_budget import process_budget_excel


st.set_page_config(page_title="Traitement Budget Excel", layout="centered")
st.title("📊 Outils de traitement de fichiers Excel Budget")

option = st.sidebar.selectbox(
    "Sélectionnez une fonctionnalité",
    ["➡️ Ajout de la feuille 'Import Odoo'", "➡️ Nettoyage & Traitement Budget"]
)


if option == "➡️ Ajout de la feuille 'Import Odoo'":
    st.header("🔄 Ajout de la feuille 'Import Odoo'")
    st.markdown("Ajoute une feuille **'Import Odoo'** dans un ou plusieurs fichiers Excel contenant un budget.")

    uploaded_files = st.file_uploader(
        "📤 Téléversez un ou plusieurs fichiers Excel (.xlsx)", 
        type=["xlsx"], 
        accept_multiple_files=True
    )

    if uploaded_files:
        temp_paths = []
        try:
            for uploaded_file in uploaded_files:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
                    temp_file.write(uploaded_file.getvalue())
                    temp_paths.append(temp_file.name)

            output_filename = f"odoo_import_{uploaded_files[0].name}"
            output_path = os.path.join(tempfile.gettempdir(), output_filename)

            # Copier le premier fichier pour modification
            with open(temp_paths[0], "rb") as src, open(output_path, "wb") as dst:
                dst.write(src.read())

            with st.spinner("⏳ Création de la feuille d'import Odoo..."):
                transform_budget_data_append_sheet(temp_paths, output_path)

            st.success("✅ Feuille ajoutée avec succès !")

            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Télécharger le fichier modifié",
                    data=f,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"❌ Erreur : {str(e)}")

        finally:
            try:
                for path in temp_paths:
                    os.remove(path)
                if 'output_path' in locals():
                    os.remove(output_path)
            except:
                pass

elif option == "➡️ Nettoyage & Traitement Budget":
    st.header("🧹 Nettoyage & Traitement Budget Excel")
    st.markdown("Nettoie et transforme un fichier budget Excel avec des feuilles structurées.")

    uploaded_file = st.file_uploader("📤 Téléversez un fichier Excel (.xlsx)", type=["xlsx"])
    date_input = st.date_input("🗓️ Sélectionnez le mois et l'année pour le traitement (le jour est ignoré)")
    if uploaded_file is not None:
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
                temp_file.write(uploaded_file.read())
                temp_input_path = temp_file.name

            with st.spinner("📈 Traitement du budget en cours..."):
                output_file_path = process_budget_excel(temp_input_path)

            st.success("✅ Traitement terminé !")

            with open(output_file_path, "rb") as f:
                st.download_button(
                    label="📥 Télécharger le fichier traité",
                    data=f,
                    file_name=os.path.basename(output_file_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"❌ Erreur : {str(e)}")
        finally:
            try:
                os.remove(temp_input_path)
                os.remove(output_file_path)
            except:
                pass
