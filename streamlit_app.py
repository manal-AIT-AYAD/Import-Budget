import streamlit as st

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

st.write("Option choisie :", st.session_state.selected_option)

if st.session_state.selected_option == "1️⃣ Traiter un fichier Budget Excel":
    st.file_uploader("Uploader pour traiter fichier", type=["xlsx"], key="upload_traitement")
elif st.session_state.selected_option == "2️⃣ Ajouter une feuille 'Import Odoo' à un fichier existant":
    st.file_uploader("Uploader pour fichier existant", type=["xlsx"], key="upload_odoo_target")
