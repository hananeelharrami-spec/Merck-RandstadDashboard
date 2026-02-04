import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# Configuration de la page
st.set_page_config(page_title="Dashboard Pilotage Randstad", layout="wide")

# ==========================================
# ZONE DE DIAGNOSTIC (S'affiche en haut)
# ==========================================
st.title("📊 Dashboard de Pilotage")

FILE_NAME = "data.xlsx"

if not os.path.exists(FILE_NAME):
    st.error(f"⛔ ERREUR CRITIQUE : Le fichier '{FILE_NAME}' est introuvable sur le serveur.")
    st.info(f"Contenu du dossier actuel : {os.listdir('.')}")
    st.stop()

# Lecture du fichier Excel
try:
    xls = pd.ExcelFile(FILE_NAME)
    onglets_trouves = xls.sheet_names
    # st.success(f"✅ Fichier '{FILE_NAME}' chargé. Onglets détectés : {onglets_trouves}")
except Exception as e:
    st.error(f"⛔ Le fichier est présent mais illisible (Format corrompu ?). Erreur : {e}")
    st.stop()

# ==========================================
# CONFIGURATION & NETTOYAGE
# ==========================================

# En-tête KPI Manuels
col_kpis = st.expander("⚙️ Modifier les paramètres manuels (Effectifs / NPS)", expanded=True)
with col_kpis:
    c1, c2 = st.columns(2)
    effectif = c1.number_input("Intérimaires en poste", value=133, step=1)
    nps = c2.number_input("NPS Intérimaire (/10)", value=9.1, step=0.1, format="%.1f")

def clean_and_scale_data(df):
    df.columns = df.columns.str.strip() # Nettoyage titres colonnes
    
    # Conversion Texte -> Nombre
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                s = df[col].astype(str).str.strip().str.replace('"', '').str.replace('\u202f', '').str.replace('\xa0', '')
                s = s.str.replace('%', '').str.replace(' ', '').str.replace(',', '.')
                df[col] = pd.to_numeric(s, errors='coerce')
            except: pass

    # Sécurisation Année
    if 'Année' in df.columns:
        df['Année'] = pd.to_numeric(df['Année'], errors='coerce').fillna(2025).astype(int)

    # Mise à l'échelle %
    for col in df.columns:
        col_lower = col.lower()
        if any(x in col_lower for x in ['taux', '%', 'atteinte', 'validation', 'rendement', 'impact']):
            if pd.api.types.is_numeric_dtype(df[col]):
                max_val = df[col].max()
                if pd.notna(max_val) and -1.5 <= max_val <= 1.5 and max_val != 0:
                    df[col] = df[col] * 100
    return df

# ==========================================
# CHARGEMENT DES DONNÉES
# ==========================================
data = {}
# Mapping : Nom voulu dans le code -> Nom réel dans ton Excel
# VERIFIE BIEN QUE LES NOMS A DROITE SONT DANS TON EXCEL
mapping = {
    "YTD": "CONSOLIDATION_YTD", 
    "RECRUT": "Recrutement_Mensuel",
    "ABS": "Absentéisme_Global_Mois", 
    "ABS_MOTIF": "Absentéisme_Par_Motif",
    "ABS_SERVICE": "Absentéisme_Par_Service", 
    "SOURCE": "KPI_Sourcing_Rendement",
    "PLAN": "Suivi_Plan_Action"
}

for key, sheet_name in mapping.items():
    if sheet_name in onglets_trouves:
        try:
            df_raw = pd.read_excel(xls, sheet_name=sheet_name)
            data[key] = clean_and_scale_data(df_raw)
        except Exception as e:
            st.warning(f"Erreur lecture onglet {sheet_name}: {e}")
    else:
        # Message discret si onglet manquant
        # st.warning(f"⚠️ Onglet manquant : '{sheet_name}'")
        pass

if not data:
    st.error("Aucun onglet valide n'a été trouvé. Vérifiez les noms des feuilles dans Excel.")
    st.stop()

# ==========================================
# FILTRES
# ==========================================
st.sidebar.header("Filtres")
annees_dispo = set()
for key, df in data.items():
    if 'Année' in df.columns:
        valid_years = [int(y) for y in df['Année'].dropna().unique() if y > 2020]
        annees_dispo.update(valid_years)

annees_dispo = sorted(list(annees_dispo))
# Selection de la dernière année par défaut
default_idx = len(annees_dispo) - 1 if annees_dispo else 0
options_annee = [str(a) for a in annees_dispo] + ["Toutes"]
annee_select = st.sidebar.selectbox("Année :", options_annee, index=max(0, default_idx))

def filter_year(df):
    if annee_select == "Toutes": return df
    if 'Année' in df.columns: return df[df['Année'] == int(annee_select)]
    return df

# ==========================================
# DASHBOARD
# ==========================================
tab1, tab2, tab3, tab4, tab5 = st.tabs(["Vue Globale", "Recrutement", "Absentéisme", "Sourcing", "Plan d'Action"])

# --- 1. VUE GLOBALE ---
with tab1:
    st.subheader(f"Performance {annee_select}")
    k1, k2 = st.columns(2)
    k1.metric("Effectifs", f"{effectif}")
    k2.metric("NPS", f"{nps}/10")
    st.markdown("---")
    
    if "YTD" in data:
        df_ytd = filter_year(data["YTD"])
        if not df_ytd.empty and 'Valeur YTD' in df_ytd.columns:
            cols = st.columns(4)
            for i, (idx, row) in enumerate(df_ytd.iterrows()):
                val = row['Valeur YTD']
                val_str = f"{val:.2f}%" if isinstance(val, (int, float)) else str(val)
                cols[i % 4].metric(row['Indicateur'], val_str)

# --- 2. RECRUTEMENT ---
with tab2:
    if "RECRUT" in data:
        df_rec = filter_year(data["RECRUT"])
        if not df_rec.empty:
            df_rec = df_rec.sort_values(['Année', 'Mois'])
            df_rec['Période'] = df_rec['Mois'].astype(str) + "/" + df_rec['Année'].astype(str)
            
            c1, c2 = st.columns(2)
            fig1 = px.line(df_rec, x='Période', y=['Taux Service', 'Taux Transfo'], markers=True)
            fig1.update_layout(yaxis_ticksuffix="%")
            c1.plotly_chart(fig1, use_container_width=True)
            
            fig2 = px.bar(df_rec, x='Période', y=['Nb Requisitions', 'Nb Hired'], barmode='group')
            c2.plotly_chart(fig2, use_container_width=True)
            
            with st.expander("Données"): st.dataframe(df_rec)

# --- 3. ABSENTÉISME ---
with tab3:
    if "ABS" in data:
        df_abs = filter_year(data["ABS"])
        if not df_abs.empty:
            df_abs = df_abs.sort_values(['Année', 'Mois'])
            df_abs['Période'] = df_abs['Mois'].astype(str) + "/" + df_abs['Année'].astype(str)
            
            fig = px.area(df_abs, x='Période', y='Taux Absentéisme', markers=True, color_discrete_sequence=['#FF5733'])
            fig.update_layout(yaxis_ticksuffix="%")
            st.plotly_chart(fig, use_container_width=True)
            
            c1, c2 = st.columns(2)
            if "ABS_MOTIF" in data:
                df_m = filter_year(data["ABS_MOTIF"])
                if not df_m.empty:
                    df_m['Période'] = df_m['Mois'].astype(str) + "/" + df_m['Année'].astype(str)
                    fig_m = px.bar(df_m, x='Période', y='Impact Motif (%)', color='Motif')
                    fig_m.update_layout(yaxis_ticksuffix="%")
                    c1.plotly_chart(fig_m, use_container_width=True)
            
            if "ABS_SERVICE" in data:
                df_s = filter_year(data["ABS_SERVICE"])
                if not df_s.empty:
                    df_s['Période'] = df_s['Mois'].astype(str) + "/" + df_s['Année'].astype(str)
                    fig_s = px.bar(df_s, x='Période', y='Taux Absentéisme', color='Service', barmode='group')
                    fig_s.update_layout(yaxis_ticksuffix="%")
                    c2.plotly_chart(fig_s, use_container_width=True)

# --- 4. SOURCING ---
with tab4:
    if "SOURCE" in data:
        df_src = filter_year(data["SOURCE"])
        if not df_src.empty and 'Source' in df_src.columns:
            df_src['Source_Clean'] = df_src['Source'].astype(str).str.upper().str.strip()
            df_agg = df_src.groupby('Source_Clean', as_index=False)[['1. Appels Reçus', '2. Validés (Sél.)', '3. Intégrés (Délégués)']].sum()
            
            # FOCUS TALENT CENTER
            st.subheader("🔥 Focus Talent Center")
            mask_tc = df_agg['Source_Clean'].str.contains("TALENT")
            df_tc = df_agg[mask_tc]
            
            if not df_tc.empty:
                v = df_tc['1. Appels Reçus'].sum()
                val = df_tc['2. Validés (Sél.)'].sum()
                i = df_tc['3. Intégrés (Délégués)'].sum()
                taux = (i/v*100) if v>0 else 0
                cols = st.columns(4)
                cols[0].metric("Appels", int(v))
                cols[1].metric("Validés", int(val))
                cols[2].metric("Intégrés", int(i))
                cols[3].metric("Rendement", f"{taux:.2f}%")
            else:
                st.warning("Aucune source 'TALENT' trouvée.")

            st.markdown("---")
            # TOP 5
            st.subheader("🏆 Top 5 Sources")
            df_top = df_agg.sort_values(['3. Intégrés (Délégués)', '1. Appels Reçus'], ascending=False).head(5)
            df_top['Color'] = df_top['Source_Clean'].apply(lambda x: "Talent Center" if "TALENT" in x else "Autres")
            
            fig = px.bar(df_top, x='Source_Clean', y='3. Intégrés (Délégués)', color='Color',
                         text='3. Intégrés (Délégués)',
                         color_discrete_map={"Talent Center": "#FF4500", "Autres": "#1f77b4"})
            fig.update_traces(textposition='outside')
            st.plotly_chart(fig, use_container_width=True)

# --- 5. PLAN D'ACTION ---
with tab5:
    if "PLAN" in data:
        df_p = data["PLAN"]
        row_g = df_p[df_p['Catégorie / Section'].astype(str).str.contains('GLOBAL', case=False, na=False)]
        if not row_g.empty:
            val = row_g.iloc[0]['% Atteinte']
            fig = go.Figure(go.Indicator(mode="gauge+number", value=val, number={'suffix':"%"}, 
                            title={'text':"Avancement Global"}, gauge={'axis':{'range':[None,100]}, 'bar':{'color':"green"}}))
            st.plotly_chart(fig, use_container_width=True)
        st.dataframe(df_p)
