import streamlit as st
import os
from excel_budget import process_budget_excel
from excel_importOddo import transform_budget_data_append_sheet

# Configuration page
st.set_page_config(page_title="Budget Excel App", layout="centered")
st.title("📊 Application de Traitement & Import Budget Excel")

# Gestion choix utilisateur
if "selected_option" not in st.session_state:
    st.session_state.selected_option = "1️⃣ Traiter un fichier Budget Excel"

option = st.radio(
    "Que voulez-vous faire ?",
    ["1️⃣ Traiter un fichier Budget Excel", "2️⃣ Ajouter une feuille 'Import Odoo' à un fichier existant"],
    index=0 if st.session_state.selected_option == "1️⃣ Traiter un fichier Budget Excel" else 1,
    key="selected_option_radio"
)

if option != st.session_state.selected_option:
    st.session_state.selected_option = option
    st.experimental_rerun()

# Option 1 : traitement fichier budget
if st.session_state.selected_option == "1️⃣ Traiter un fichier Budget Excel":
    st.subheader("📥 Téléversez un fichier Excel à traiter")
    uploaded_file = st.file_uploader("Fichier Excel (.xlsx)", type=["xlsx"], key="upload_traitement")

    if uploaded_file:
        temp_input = "temp_to_process.xlsx"
        with open(temp_input, "wb") as f:
            f.write(uploaded_file.read())

        with st.spinner("Traitement du fichier..."):
            output_file = process_budget_excel(temp_input)

        st.success("✅ Traitement terminé.")
        with open(output_file, "rb") as f:
            st.download_button(
                label="📥 Télécharger le fichier traité",
                data=f,
                file_name="fichier_traite.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        os.remove(temp_input)
        os.remove(output_file)

# Option 2 : ajout feuille "Import Odoo" dans fichier existant
elif st.session_state.selected_option == "2️⃣ Ajouter une feuille 'Import Odoo' à un fichier existant":
    st.subheader("📤 Téléversez le fichier Excel existant dans lequel ajouter la feuille")
    existing_file = st.file_uploader("Fichier Excel existant", type=["xlsx"], key="upload_odoo_target")

    if existing_file:
        existing_path = "uploaded_existing.xlsx"
        with open(existing_path, "wb") as f:
            f.write(existing_file.read())

        with st.spinner("Ajout de la feuille 'Import Odoo'..."):
            # Ici on suppose que ta fonction sait récupérer les sources internes ou est adaptée pour cet usage
            transform_budget_data_append_sheet([], existing_path)

        st.success("✅ Feuille ajoutée avec succès.")
        with open(existing_path, "rb") as f:
            st.download_button(
                label="📥 Télécharger le fichier modifié",
                data=f,
                file_name="fichier_avec_import_odoo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        os.remove(existing_path)

# Bouton reset
st.markdown("---")
if st.button("🔄 Réinitialiser l'application"):
    st.session_state.clear()
    st.experimental_rerun()
