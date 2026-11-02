import streamlit as st
import pandas as pd
import os

st.set_page_config(
    page_title="OKPLAST - Adaptation Moules & Machines", 
    page_icon="🏭", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Cacher le menu Streamlit et le footer
hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# Affichage du logo et du titre côte à côte
col_logo, col_titre = st.columns([1, 5])
with col_logo:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=160)
with col_titre:
    st.title("Adaptation Moules & Machines")

EXCEL_FILE = "liste changement2025-2026.xlsx"

# --- ESPACE ADMINISTRATEUR DIRECTEMENT SUR LA PAGE PRINCIPALE ---
st.markdown("---")
with st.expander("🔐 Espace Administrateur (Mise à jour du fichier Excel)", expanded=True):
    mot_de_passe_saisi = st.text_input("Entrez le mot de passe administrateur", type="password")
    
    if mot_de_passe_saisi == "102030Tn":
        st.success("Mode Administrateur activé ✅")
        uploaded_file = st.file_uploader("Choisir un nouveau fichier Excel à téléverser", type=["xlsx"])
        if uploaded_file is not None:
            with open(EXCEL_FILE, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.success("Fichier mis à jour avec succès !")
            st.rerun()
    elif mot_de_passe_saisi != "":
        st.error("Mot de passe incorrect ❌")
st.markdown("---")

# Chargement des données
@st.cache_data(ttl=10)
def load_data():
    if os.path.exists(EXCEL_FILE):
        return pd.read_excel(EXCEL_FILE)
    return None

df = load_data()

if df is not None:
    st.success(f"Base chargée : {len(df)} enregistrements")
    
    # Recherche
    option = st.radio("Rechercher par :", ["Code Moule", "Nom Produit", "Machine"])
    query = st.text_input("Saisie de recherche :")
    
    if query:
        filtered_df = df[df.astype(str).apply(lambda row: query.lower() in row.str.lower().values, axis=1)]
        st.dataframe(filtered_df, use_container_width=True)
    else:
        st.dataframe(df.head(100), use_container_width=True)
else:
    st.error("Aucun fichier Excel trouvé.")
