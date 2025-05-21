import streamlit as st
from excel_budget import process_budget_excel
from excel_importOddo import transform_budget_data_append_sheet
import os
import tempfile

st.set_page_config(
    page_title="Outils Excel Budget",
    layout="wide",
    initial_sidebar_state="expanded"
)

def main():
    st.sidebar.title("🛠️ Outils Excel Budget")
    options = ["Créer modèle de budget", "Ajouter feuille d'import Odoo"]
    choice = st.sidebar.radio("Choisissez une fonctionnalité:", options)
    
    if choice == "Créer modèle de budget":
        budget_template_page()
    else:
        odoo_import_page()

def budget_template_page():
    st.title("📊 Création du modèle de budget")
    st.write("Cet outil transforme un compte de résultats en modèle de budget avec colonnes mensuelles")
    
    uploaded_file = st.file_uploader("Téléversez un fichier Excel (.xlsx)", type=["xlsx"], key="budget_upload")
    
    if uploaded_file is not None:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
            temp_input_path = temp_file.name
            temp_file.write(uploaded_file.getvalue())
        
        output_filename = f"budget_{uploaded_file.name}"
        
        with st.spinner("📈 Création du modèle de budget en cours..."):
            try:
                # Appel à la fonction de traitement
                output_path = process_budget_excel(temp_input_path, output_filename)
                st.success("✅ Traitement terminé !")
                
                # Bouton de téléchargement
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 Télécharger le modèle de budget",
                        data=f,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"❌ Erreur lors du traitement: {str(e)}")
            finally:
                # Nettoyage
                try:
                    os.remove(temp_input_path)
                    if 'output_path' in locals():
                        os.remove(output_path)
                except:
                    pass

def odoo_import_page():
    st.title("🔄 Création de la feuille d'import Odoo")
    st.write("Cet outil ajoute une feuille 'Import Odoo' à votre fichier budget")
    
    uploaded_files = st.file_uploader(
        "Téléversez un ou plusieurs fichiers budget (.xlsx)", 
        type=["xlsx"], 
        accept_multiple_files=True,
        key="odoo_upload"
    )
    
    if uploaded_files:
        temp_paths = []
        try:
            # Sauvegarde temporaire des fichiers uploadés
            for uploaded_file in uploaded_files:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
                    temp_path = temp_file.name
                    temp_file.write(uploaded_file.getvalue())
                    temp_paths.append(temp_path)
            
            # Le premier fichier sera utilisé comme fichier de sortie
            output_filename = f"odoo_import_{uploaded_files[0].name}"
            output_path = f"temp_{output_filename}"
            
            # Copier le premier fichier vers la sortie
            with open(temp_paths[0], "rb") as src, open(output_path, "wb") as dst:
                dst.write(src.read())
            
            with st.spinner("🔄 Création de la feuille d'import Odoo en cours..."):
                # Appel à la fonction de traitement
                transform_budget_data_append_sheet(temp_paths, output_path)
                st.success("✅ Traitement terminé !")
                
                # Bouton de téléchargement
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 Télécharger le fichier avec feuille d'import Odoo",
                        data=f,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"❌ Erreur lors du traitement: {str(e)}")
        finally:
            # Nettoyage
            try:
                for temp_path in temp_paths:
                    os.remove(temp_path)
                if 'output_path' in locals():
                    os.remove(output_path)
            except:
                pass

if __name__ == "__main__":
    main()