import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import glob
import os

# Configuration de la page
st.set_page_config(page_title="Dashboard Pilotage Randstad", layout="wide")

# --- EN-TÊTE ---
col_t, col_k = st.columns([2, 3])
with col_t:
    st.title("📊 Pilotage")
    st.caption("Randstad / Merck")

with col_k:
    c1, c2 = st.columns(2)
    effectif = c1.number_input("👥 Intérimaires en poste", value=133, step=1)
    nps = c2.number_input("⭐ NPS Intérimaire", value=9.1, step=0.1, format="%.1f")

# --- FONCTION DE NETTOYAGE ---
def clean_df(df):
    # Nettoyage noms colonnes
    df.columns = df.columns.str.strip()
    
    # Nettoyage valeurs (virgules, %, espaces)
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                s = df[col].astype(str).str.strip().str.replace('"', '').str.replace('\u202f', '').str.replace('\xa0', '')
                s = s.str.replace('%', '').str.replace(' ', '').str.replace(',', '.')
                df[col] = pd.to_numeric(s, errors='coerce')
            except: pass
            
    # Mise à l'échelle des % (0.88 -> 88.0)
    for col in df.columns:
        # Mots clés pour identifier les taux
        if any(x in col.lower() for x in ['taux', '%', 'atteinte', 'rendement', 'validation', 'service', 'transfo']):
            if pd.api.types.is_numeric_dtype(df[col]):
                max_val = df[col].max()
                # Si le max est petit (ex: 1 ou 0.9), c'est un format décimal Excel -> x100
                if pd.notna(max_val) and -1.5 <= max_val <= 1.5 and max_val != 0:
                    df[col] = df[col] * 100
                    
    # Gestion Année (Conversion float -> int)
    if 'Année' in df.columns:
        df['Année'] = df['Année'].fillna(0).astype(int)
        
    return df

# --- CHARGEMENT UNIQUEMENT EXCEL ---
@st.cache_data
def load_data():
    data = {}
    
    # On cherche uniquement les fichiers .xlsx
    excel_files = glob.glob("*.xlsx")
    
    if not excel_files:
        return None, "Aucun fichier .xlsx"

    # On prend le premier fichier Excel trouvé (ex: data.xlsx)
    file_path = excel_files[0]
    
    try:
        xls = pd.ExcelFile(file_path)
        
        # Mapping : Clé Code -> Nom de l'onglet dans Excel
        sheet_map = {
            "YTD": "CONSOLIDATION_YTD",
            "RECRUT": "Recrutement_Mensuel",
            "ABS": "Absentéisme_Global_Mois",
            "ABS_MOTIF": "Absentéisme_Par_Motif",
            "ABS_SERVICE": "Absentéisme_Par_Service",
            "SOURCE": "KPI_Sourcing_Rendement",
            "PLAN": "Suivi_Plan_Action"
        }
        
        # On charge chaque onglet s'il existe
        for key, sheet_name in sheet_map.items():
            if sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                data[key] = clean_df(df)
            else:
                # Optionnel : Essayer de trouver l'onglet même si la casse diffère
                for s in xls.sheet_names:
                    if sheet_name.lower() in s.lower():
                        df = pd.read_excel(xls, sheet_name=s)
                        data[key] = clean_df(df)
                        break
                        
        return data, file_path

    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier Excel : {e}")
        return None, str(e)

data, source_name = load_data()

if not data:
    st.error("❌ Aucun fichier Excel (.xlsx) valide trouvé.")
    st.info("Veuillez rassembler vos données dans un fichier Excel unique (ex: data.xlsx) avec les bons onglets.")
    st.stop()

# st.toast(f"Données chargées depuis : {source_name}", icon="✅")
st.markdown("---")

# --- FILTRES ---
st.sidebar.header("Filtres")

# Années disponibles
annees = set()
for k, df in data.items():
    if 'Année' in df.columns:
        annees.update(df['Année'].unique())

# On filtre les années bizarres (0) et on trie décroissant (2026 en premier)
annees = sorted([int(a) for a in annees if a > 2020], reverse=True)
opts = [str(a) for a in annees] + ["Vue Globale"]

# Par défaut, sélectionne l'année la plus récente
annee_sel = st.sidebar.selectbox("Année", opts, index=0)

def filter(df):
    if annee_sel == "Vue Globale": return df
    if 'Année' in df.columns: return df[df['Année'] == int(annee_sel)]
    return df

# --- DASHBOARD ---
t1, t2, t3, t4, t5 = st.tabs(["Global", "Recrutement", "Absentéisme", "Sourcing", "Plan d'Action"])

# 1. GLOBAL YTD
with t1:
    st.subheader(f"Performance YTD - {annee_sel}")
    if "YTD" in data:
        df = filter(data["YTD"])
        if not df.empty and 'Valeur YTD' in df.columns:
            # Tri par Année pour lisibilité
            df = df.sort_values('Année', ascending=False)
            cols = st.columns(4)
            for i, r in df.iterrows():
                val = r['Valeur YTD']
                v_str = f"{val:.2f}%" if isinstance(val, (int, float)) else str(val)
                label = r['Indicateur']
                if annee_sel == "Vue Globale": label += f" ({r['Année']})"
                
                cols[i % 4].metric(label, v_str)
        else:
            st.info(f"Pas de données YTD pour {annee_sel}")

# 2. RECRUTEMENT
with t2:
    if "RECRUT" in data:
        df = filter(data["RECRUT"])
        if not df.empty:
            if 'Mois' in df.columns:
                df = df.sort_values(['Année', 'Mois'])
                df['Période'] = df['Mois'].astype(str) + "/" + df['Année'].astype(str)
                
                c1, c2 = st.columns(2)
                # Graphique Taux
                fig = px.line(df, x='Période', y=['Taux Service', 'Taux Transfo'], markers=True, title="Qualité")
                fig.update_layout(yaxis_ticksuffix="%")
                c1.plotly_chart(fig, use_container_width=True)
                
                # Graphique Volume
                fig2 = px.bar(df, x='Période', y=['Nb Requisitions', 'Nb Hired'], barmode='group', title="Volumes")
                c2.plotly_chart(fig2, use_container_width=True)
                
                with st.expander("Détail"): st.dataframe(df)

# 3. ABSENTEISME
with t3:
    if "ABS" in data:
        df = filter(data["ABS"])
        if not df.empty:
            df = df.sort_values(['Année', 'Mois'])
            df['Période'] = df['Mois'].astype(str) + "/" + df['Année'].astype(str)
            
            st.subheader("Taux Global")
            fig = px.area(df, x='Période', y='Taux Absentéisme', markers=True, color_discrete_sequence=['#FF5733'])
            fig.update_layout(yaxis_ticksuffix="%")
            st.plotly_chart(fig, use_container_width=True)
            
            c1, c2 = st.columns(2)
            if "ABS_MOTIF" in data:
                dfm = filter(data["ABS_MOTIF"])
                if not dfm.empty:
                    dfm['Période'] = dfm['Mois'].astype(str) + "/" + dfm['Année'].astype(str)
                    c1.subheader("Par Motif")
                    figm = px.bar(dfm, x='Période', y='Impact Motif (%)', color='Motif')
                    figm.update_layout(yaxis_ticksuffix="%")
                    c1.plotly_chart(figm, use_container_width=True)
            
            if "ABS_SERVICE" in data:
                dfs = filter(data["ABS_SERVICE"])
                if not dfs.empty:
                    dfs['Période'] = dfs['Mois'].astype(str) + "/" + dfs['Année'].astype(str)
                    c2.subheader("Par Service")
                    figs = px.bar(dfs, x='Période', y='Taux Absentéisme', color='Service', barmode='group')
                    figs.update_layout(yaxis_ticksuffix="%")
                    c2.plotly_chart(figs, use_container_width=True)

# 4. SOURCING
with t4:
    if "SOURCE" in data:
        df = filter(data["SOURCE"])
        
        # On vérifie que les colonnes existent (format KPI_Sourcing)
        if '1. Appels Reçus' in df.columns:
             # Normalisation Source pour le tri et les couleurs
             if 'Source' in df.columns:
                 df['Source_Clean'] = df['Source'].astype(str).str.upper().str.strip()
             
             # Aggrégation sur la période (Année choisie)
             # On recalcule les sommes au cas où il y aurait plusieurs lignes par source (mois différents)
             df_agg = df.groupby('Source_Clean')[['1. Appels Reçus', '2. Validés (Sél.)', '3. Intégrés (Délégués)']].sum().reset_index()

             if not df_agg.empty:
                # --- FOCUS TALENT CENTER ---
                st.subheader("🔥 Focus Talent Center")
                # Recherche du mot clé "TALENT"
                mask_tc = df_agg['Source_Clean'].str.contains("TALENT")
                dftc = df_agg[mask_tc]
                
                if not dftc.empty:
                    v = dftc['1. Appels Reçus'].sum()
                    val = dftc['2. Validés (Sél.)'].sum()
                    i = dftc['3. Intégrés (Délégués)'].sum()
                    # Calcul rendement en direct
                    r = (i/v*100) if v>0 else 0
                    
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Appels", int(v))
                    c2.metric("Validés", int(val))
                    c3.metric("Intégrés", int(i))
                    c4.metric("Rendement", f"{r:.2f}%")
                else:
                    st.warning("Pas de source 'Talent Center' trouvée sur cette période.")
                
                st.markdown("---")
                
                # --- TOP 5 SOURCES ---
                st.subheader("🏆 Top 5 Sources (Intégrations)")
                # Tri par Intégrés puis Volume
                df_top = df_agg.sort_values(['3. Intégrés (Délégués)', '1. Appels Reçus'], ascending=False).head(5)
                
                # Création couleur spéciale
                df_top['Type'] = df_top['Source_Clean'].apply(lambda x: "Talent Center" if "TALENT" in x else "Autres")
                
                fig = px.bar(df_top, x='Source_Clean', y='3. Intégrés (Délégués)', color='Type',
                             text='3. Intégrés (Délégués)',
                             color_discrete_map={"Talent Center": "#FF4500", "Autres": "#1f77b4"})
                fig.update_traces(textposition='outside')
                st.plotly_chart(fig, use_container_width=True)
                
                with st.expander("Données complètes"):
                    st.dataframe(df_agg)

# 5. PLAN D'ACTION
with t5:
    if "PLAN" in data:
        df = data["PLAN"]
        # Jauge Globale
        row = df[df['Catégorie / Section'].astype(str).str.contains('GLOBAL', case=False, na=False)]
        if not row.empty:
            val = row.iloc[0]['% Atteinte']
            fig = go.Figure(go.Indicator(mode="gauge+number", value=val, number={'suffix':"%"}, 
                            title={'text':"Avancement Global"}, gauge={'axis':{'range':[None,100]}, 'bar':{'color':"green"}}))
            st.plotly_chart(fig, use_container_width=True)
        st.dataframe(df)
