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

# Chargement des données et détection des lignes colorées (anomalies / changements de machine)
@st.cache_data(ttl=10)
def load_data_with_colors():
    if not os.path.exists(EXCEL_FILE):
        return None, []
    
    df = pd.read_excel(EXCEL_FILE)
    
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    ws = wb.active
    
    colored_rows = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=0):
        for cell in row:
            if cell.fill and cell.fill.fill_type:
                if cell.fill.start_color and cell.fill.start_color.rgb:
                    color_val = str(cell.fill.start_color.rgb)
                    if color_val not in ["00000000", "FFFFFFFF", "None"]:
                        colored_rows.append(row_idx)
                        break
                        
    return df, colored_rows

df, colored_rows = load_data_with_colors()

if df is not None:
    st.success(f"Base chargée : {len(df)} enregistrements")
    
    def highlight_rows(row):
        if row.name in colored_rows:
            return ['background-color: #ff9999; color: black; font-weight: bold;'] * len(row)
        return [''] * len(row)

    # Options de filtrage et recherche
    col_f1, col_f2 = st.columns([2, 2])
    with col_f1:
        option = st.selectbox("Rechercher par :", ["Code Moule", "Nom Produit", "Machine"])
    with col_f2:
        # Option pour afficher uniquement les anomalies (lignes en rouge)
        seule_anomalies = st.checkbox("⚠️ Afficher uniquement les changements de machine non prévus (Lignes rouges)")

    query = st.text_input("Saisie de recherche :")
    
    # Filtrage des données
    display_df = df.copy()
    
    if seule_anomalies:
        # On ne garde que les indices des lignes colorées
        display_df = display_df.iloc[colored_rows]

    if query:
        display_df = display_df[display_df.astype(str).apply(lambda row: query.lower() in row.str.lower().values, axis=1)]
    
    # Affichage du tableau stylisé
    if not display_df.empty:
        st.dataframe(display_df.style.apply(highlight_rows, axis=1), use_container_width=True)
    else:
        st.info("Aucun résultat ne correspond à votre recherche ou filtre.")
else:
    st.error("Aucun fichier Excel trouvé.")
