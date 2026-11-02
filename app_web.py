import streamlit as st
import pandas as pd
import openpyxl
import os

st.set_page_config(
    page_title="OKPLAST - Adaptations Moules & Machines", 
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

# En-tête avec un grand logo et le titre corrigé
col_logo, col_titre = st.columns([1, 4])
with col_logo:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=250)
with col_titre:
    st.markdown("<h1 style='padding-top: 20px;'>Adaptations Moules & Machines</h1>", unsafe_allow_html=True)

EXCEL_FILE = "liste changement2025-2026.xlsx"

# --- ESPACE ADMINISTRATEUR DISCRET ---
with st.expander("⚙️ Paramètres / Espace Administrateur"):
    mot_de_passe_saisi = st.text_input("Mot de passe administrateur", type="password")
    
    if mot_de_passe_saisi == "102030Tn":
        st.success("Mode Administrateur activé ✅")
        uploaded_file = st.file_uploader("Mettre à jour le fichier Excel", type=["xlsx"])
        if uploaded_file is not None:
            with open(EXCEL_FILE, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.success("Fichier mis à jour avec succès !")
            st.rerun()
    elif mot_de_passe_saisi != "":
        st.error("Mot de passe incorrect ❌")

st.markdown("---")

# Chargement des données et détection des lignes rouges
@st.cache_data(ttl=10)
def load_data_with_colors():
    if not os.path.exists(EXCEL_FILE):
        return None, []
    
    # Chargement normal avec pandas
    df = pd.read_excel(EXCEL_FILE)
    
    # Lecture des couleurs avec openpyxl pour identifier les lignes en surbrillance (ex: rouge)
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    ws = wb.active
    
    red_rows = []
    # On parcourt les lignes du fichier Excel (à partir de la ligne 2 pour sauter l'en-tête)
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=0):
        for cell in row:
            # Vérifie si le fond de la cellule est rouge (code couleur hexadécimal contenant "FF0000" ou similaire)
            if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb:
                color_val = str(cell.fill.start_color.rgb)
                if "FF0000" in color_val or "FFFF0000" in color_val:
                    red_rows.append(row_idx)
                    break
                    
    return df, red_rows

df, red_rows = load_data_with_colors()

if df is not None:
    st.success(f"Base chargée : {len(df)} enregistrements")
    
    # Fonction pour colorer les lignes en rouge dans l'affichage Streamlit
    def highlight_rows(row):
        if row.name in red_rows:
            return ['background-color: #ff9999'] * len(row)
        return [''] * len(row)

    # Recherche
    option = st.radio("Rechercher par :", ["Code Moule", "Nom Produit", "Machine"])
    query = st.text_input("Saisie de recherche :")
    
    if query:
        filtered_df = df[df.astype(str).apply(lambda row: query.lower() in row.str.lower().values, axis=1)]
        st.dataframe(filtered_df.style.apply(highlight_rows, axis=1), use_container_width=True)
    else:
        st.dataframe(df.head(100).style.apply(highlight_rows, axis=1), use_container_width=True)
else:
    st.error("Aucun fichier Excel trouvé.")
