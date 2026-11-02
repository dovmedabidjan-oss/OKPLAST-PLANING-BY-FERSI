import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Adaptation Moules & Machines", page_icon="🏭", layout="wide")

st.title("🏭 Adaptation Moules & Machines")

EXCEL_FILE = "liste changement2025-2026.xlsx"

# Section protégée par mot de passe dans la barre latérale
with st.sidebar.expander("🔄 Mettre à jour la base Excel"):
    mot_de_passe_admin = "102030Tn"
    
    saisie_mdp = st.text_input("Entrez le mot de passe administrateur", type="password")
    
    if saisie_mdp == mot_de_passe_admin:
        st.success("Accès autorisé")
        uploaded_file = st.file_uploader("Choisir un nouveau fichier Excel", type=["xlsx"])
        if uploaded_file is not None:
            with open(EXCEL_FILE, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.success("Fichier mis à jour avec succès !")
            st.rerun()
    elif saisie_mdp != "":
        st.error("Mot de passe incorrect")

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
