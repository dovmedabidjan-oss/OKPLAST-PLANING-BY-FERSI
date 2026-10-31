import streamlit as st
import pandas as pd
import os

# Configuration de la page mobile
st.set_page_config(page_title="Adaptation Moules & Machines", page_icon="🏭", layout="wide")

st.title("🏭 Adaptation Moules & Machines")

# Chargement des données
@st.cache_data
def load_data(file_path):
    if not os.path.exists(file_path):
        return None
    df_raw = pd.read_excel(file_path, header=None)
    df_raw.columns = ['DATE', 'OKPLT', 'PROG', 'DEPROG', 'MACH']
    df_raw['DATE_STR'] = pd.to_datetime(df_raw['DATE'], errors='coerce').dt.strftime('%d/%m/%Y').fillna(str(df_raw['DATE']))
    df_raw['DESIGNATION'] = df_raw['PROG'].combine_first(df_raw['DEPROG']).fillna('N/A')
    df_raw['OKPLT'] = df_raw['OKPLT'].astype(str).str.strip()
    df_raw['MACH'] = df_raw['MACH'].astype(str).str.strip()
    df_raw['DESIGNATION'] = df_raw['DESIGNATION'].astype(str).str.strip()
    return df_raw[~df_raw['OKPLT'].str.contains('AUCUN|JAMAIS|DATE|PAS!', case=False, na=False)].copy()

file_path = 'liste changement2025-2026.xlsx'
df = load_data(file_path)

if df is None:
    st.error("⚠️ Fichier 'liste changement2025-2026.xlsx' introuvable dans le dossier.")
else:
    st.success(f"Base chargée : {len(df)} enregistrements")

    # Recherche
    critere = st.radio("Rechercher par :", ["Code Moule", "Nom Produit", "Machine"], horizontal=True)
    recherche = st.text_input("Saisie de recherche :", "").strip().upper()

    if recherche:
        if critere == "Code Moule":
            filtered = df[df['OKPLT'].str.upper().str.contains(recherche, na=False)]
            grouped = filtered.groupby(['OKPLT', 'DESIGNATION', 'MACH']).size().reset_index(name='raw_count')
            grouped['Nb de fois'] = (grouped['raw_count'] // 2).clip(lower=1)
            st.dataframe(grouped[['OKPLT', 'DESIGNATION', 'MACH', 'Nb de fois']], use_container_width=True)

        elif critere == "Nom Produit":
            filtered = df[df['DESIGNATION'].str.upper().str.contains(recherche, na=False)]
            grouped = filtered.groupby(['OKPLT', 'DESIGNATION', 'MACH']).size().reset_index(name='raw_count')
            grouped['Nb de fois'] = (grouped['raw_count'] // 2).clip(lower=1)
            st.dataframe(grouped[['OKPLT', 'DESIGNATION', 'MACH', 'Nb de fois']], use_container_width=True)

        elif critere == "Machine":
            filtered = df[df['MACH'].str.upper().str.contains(recherche, na=False)]
            grouped = filtered.groupby(['MACH', 'OKPLT', 'DESIGNATION']).size().reset_index(name='raw_count')
            grouped['Nb de fois'] = (grouped['raw_count'] // 2).clip(lower=1)
            st.dataframe(grouped[['MACH', 'OKPLT', 'DESIGNATION', 'Nb de fois']], use_container_width=True)

        # Détails des dates
        st.subheader("📅 Historique des dates de montage / démontage")
        st.dataframe(filtered[['DATE_STR', 'PROG', 'DEPROG', 'OKPLT', 'MACH']], use_container_width=True)