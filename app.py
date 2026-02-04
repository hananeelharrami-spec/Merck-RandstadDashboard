import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import glob

# Configuration de la page
st.set_page_config(page_title="Dashboard Pilotage Randstad", layout="wide")

# --- EN-TÊTE AVEC KPI PLANNING ---
col_title, col_plan = st.columns([3, 1])
with col_title:
    st.title("📊 Dashboard de Pilotage - Randstad / Merck")

with col_plan:
    # Zone Planning modifiable
    st.markdown("### 👥 Planning / Effectifs")
    # On met 133 par défaut, mais tu peux le changer à la volée
    effectif = st.number_input("Intérimaires en poste (Semaine en cours)", value=133, step=1)
    st.caption("Donnée temps réel")

# --- FONCTION DE NETTOYAGE RENFORCÉE ---
def clean_and_scale_data(df):
    # 1. CONVERSION TEXTE -> NOMBRE
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                series = df[col].astype(str).str.replace('"', '').str.strip()
                series = series.str.replace('%', '').str.replace(' ', '').str.replace('\u202f', '')
                series = series.str.replace(',', '.')
                df[col] = pd.to_numeric(series, errors='ignore')
            except Exception:
                pass

    # 2. MISE A L'ECHELLE DES POURCENTAGES (0.88 -> 88.0)
    for col in df.columns:
        col_lower = col.lower()
        if any(x in col_lower for x in ['taux', '%', 'atteinte', 'validation', 'rendement', 'impact']):
            if pd.api.types.is_numeric_dtype(df[col]):
                max_val = df[col].max()
                if pd.notna(max_val) and -1.5 <= max_val <= 1.5 and max_val != 0:
                    df[col] = df[col] * 100
    return df

# --- CHARGEMENT INTELLIGENT ---
@st.cache_data
def load_data():
    excel_files = glob.glob("*.xlsx")
    if not excel_files:
        return None, None

    found_file = excel_files[0]
    data = {}
    try:
        xls = pd.ExcelFile(found_file)
        all_sheets = xls.sheet_names
        
        # Mapping complet avec les nouveaux onglets Absentéisme
        expected = {
            "YTD": "CONSOLIDATION_YTD",
            "RECRUT": "Recrutement_Mensuel",
            "ABS": "Absentéisme_Global_Mois",
            "ABS_MOTIF": "Absentéisme_Par_Motif",   # NOUVEAU
            "ABS_SERVICE": "Absentéisme_Par_Service", # NOUVEAU
            "SOURCE": "KPI_Sourcing_Rendement",
            "PLAN": "Suivi_Plan_Action"
        }
        
        for key, sheet_name in expected.items():
            if sheet_name in all_sheets:
                df_raw = pd.read_excel(found_file, sheet_name=sheet_name)
                data[key] = clean_and_scale_data(df_raw)
        return data, found_file
        
    except Exception as e:
        st.error(f"Erreur de lecture : {e}")
        return None, found_file

# Exécution du chargement
data, filename = load_data()

if data is None:
    st.error("❌ Aucun fichier Excel (.xlsx) trouvé sur le serveur.")
    st.info("Veuillez uploader votre fichier de données (ex: data.xlsx) sur GitHub.")
    st.stop()
else:
    st.markdown("---")

# --- DASHBOARD ---

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📈 Vue Globale", "🤝 Recrutement", "🏥 Absentéisme (Détails)", "🔍 Sourcing", "✅ Plan d'Action"])

# --- 1. VUE GLOBALE (YTD) ---
with tab1:
    st.header("Performance Annuelle (Year-To-Date)")
    
    # On affiche aussi le planning ici pour la vue d'ensemble
    st.metric(label="📍 Effectif Intérimaire Actuel", value=effectif, delta="En poste cette semaine")
    st.markdown("---")
    
    if "YTD" in data:
        df_ytd = data["YTD"]
        col_val = 'Valeur YTD'
        if col_val in df_ytd.columns:
            cols = st.columns(4)
            for index, row in df_ytd.iterrows():
                indic = row['Indicateur']
                val = row[col_val]
                val_str = f"{val:.2f}%" if isinstance(val, (int, float)) else str(val)
                cols[index % 4].metric(label=indic, value=val_str)

# --- 2. RECRUTEMENT ---
with tab2:
    st.header("Recrutement Mensuel")
    if "RECRUT" in data:
        df_rec = data["RECRUT"]
        if 'Mois' in df_rec.columns:
            if 'Année' in df_rec.columns:
                df_rec = df_rec.sort_values(['Année', 'Mois'])
                df_rec['Période'] = df_rec['Mois'].astype(str) + "/" + df_rec['Année'].astype(str)
            else:
                df_rec['Période'] = df_rec['Mois'].astype(str)

            c1, c2 = st.columns(2)
            with c1:
                st.subheader("Taux de Service & Transformation")
                fig = px.line(df_rec, x='Période', y=['Taux Service', 'Taux Transfo'], markers=True)
                fig.update_layout(yaxis_ticksuffix="%")
                st.plotly_chart(fig, use_container_width=True)

            with c2:
                st.subheader("Volume Commandes vs Hired")
                fig_bar = px.bar(df_rec, x='Période', y=['Nb Requisitions', 'Nb Hired'], barmode='group')
                st.plotly_chart(fig_bar, use_container_width=True)
            
            with st.expander("Voir le détail chiffré"):
                df_show = df_rec.copy()
                for col in df_show.columns:
                    if 'Taux' in col and pd.api.types.is_numeric_dtype(df_show[col]):
                        df_show[col] = df_show[col].apply(lambda x: f"{x:.2f}%")
                st.dataframe(df_show)

# --- 3. ABSENTÉISME (NOUVEAUX GRAPHIQUES) ---
with tab3:
    st.header("Analyse de l'Absentéisme")
    
    # 1. KPI GLOBAL (Graphique existant)
    if "ABS" in data:
        df_abs = data["ABS"]
        if 'Mois' in df_abs.columns:
            if 'Année' in df_abs.columns:
                 df_abs = df_abs.sort_values(['Année', 'Mois'])
                 df_abs['Période'] = df_abs['Mois'].astype(str) + "/" + df_abs['Année'].astype(str)
            else:
                 df_abs['Période'] = df_abs['Mois'].astype(str)
        
        st.subheader("Taux Global")
        fig_abs = px.area(df_abs, x='Période', y='Taux Absentéisme', markers=True, color_discrete_sequence=['#FF5733'])
        fig_abs.update_layout(yaxis_ticksuffix="%")
        st.plotly_chart(fig_abs, use_container_width=True)
    
    st.markdown("---")

    # 2. DETAILS PAR MOTIF ET SERVICE (NOUVEAU)
    c_motif, c_service = st.columns(2)
    
    with c_motif:
        st.subheader("Par Motif (Impact %)")
        if "ABS_MOTIF" in data:
            df_motif = data["ABS_MOTIF"]
            # Création Période
            if 'Année' in df_motif.columns and 'Mois' in df_motif.columns:
                df_motif = df_motif.sort_values(['Année', 'Mois'])
                df_motif['Période'] = df_motif['Mois'].astype(str) + "/" + df_motif['Année'].astype(str)
            
            # Graphique Empilé
            if 'Impact Motif (%)' in df_motif.columns:
                fig_mot = px.bar(df_motif, x='Période', y='Impact Motif (%)', color='Motif',
                                 title="Répartition des Motifs", barmode='stack')
                fig_mot.update_layout(yaxis_ticksuffix="%")
                st.plotly_chart(fig_mot, use_container_width=True)
            else:
                st.warning("Données Motif incomplètes")
        else:
            st.info("Onglet 'Absentéisme_Par_Motif' manquant.")

    with c_service:
        st.subheader("Par Service (Taux %)")
        if "ABS_SERVICE" in data:
            df_service = data["ABS_SERVICE"]
            # Création Période
            if 'Année' in df_service.columns and 'Mois' in df_service.columns:
                df_service = df_service.sort_values(['Année', 'Mois'])
                df_service['Période'] = df_service['Mois'].astype(str) + "/" + df_service['Année'].astype(str)
            
            # Graphique Groupé
            if 'Taux Absentéisme' in df_service.columns:
                fig_serv = px.bar(df_service, x='Période', y='Taux Absentéisme', color='Service',
                                  title="Comparatif Services", barmode='group')
                fig_serv.update_layout(yaxis_ticksuffix="%")
                st.plotly_chart(fig_serv, use_container_width=True)
            else:
                st.warning("Données Service incomplètes")
        else:
            st.info("Onglet 'Absentéisme_Par_Service' manquant.")

# --- 4. SOURCING ---
with tab4:
    st.header("Performance Sourcing")
    if "SOURCE" in data:
        df_src = data["SOURCE"]
        if 'Source' in df_src.columns:
            df_agg = df_src.groupby('Source', as_index=False)[['1. Appels Reçus', '2. Validés (Sél.)', '3. Intégrés (Délégués)']].sum()
            
            st.subheader("🔥 Focus : Efficience Talent Center")
            mask_tc = df_agg['Source'].astype(str).str.upper().str.contains("TALENT CENTER")
            df_tc = df_agg[mask_tc]
            
            if not df_tc.empty:
                vol_tc = df_tc['1. Appels Reçus'].sum()
                val_tc = df_tc['2. Validés (Sél.)'].sum()
                int_tc = df_tc['3. Intégrés (Délégués)'].sum()
                taux_transfo_tc = (int_tc / vol_tc * 100) if vol_tc > 0 else 0
                
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("Volume Appels (TC)", int(vol_tc))
                k2.metric("Validés (TC)", int(val_tc))
                k3.metric("Intégrés (TC)", int(int_tc))
                k4.metric("Rendement Final (TC)", f"{taux_transfo_tc:.2f}%", delta_color="normal")
            
            st.markdown("---")

            st.subheader("🏆 Top 5 des Meilleures Sources")
            df_top5 = df_agg.sort_values(by=['3. Intégrés (Délégués)', '1. Appels Reçus'], ascending=[False, False]).head(5)
            
            def categorize_source(source_name):
                return "Talent Center" if "TALENT CENTER" in str(source_name).upper() else "Autres Sources"

            df_top5['Catégorie'] = df_top5['Source'].apply(categorize_source)
            
            fig_best = px.bar(
                df_top5,
                x='Source',
                y='3. Intégrés (Délégués)',
                color='Catégorie',
                text='3. Intégrés (Délégués)',
                color_discrete_map={"Talent Center": "#FF4500", "Autres Sources": "#1f77b4"}
            )
            fig_best.update_traces(textposition='outside')
            st.plotly_chart(fig_best, use_container_width=True)

# --- 5. PLAN D'ACTION ---
with tab5:
    st.header("Plan d'Action")
    if "PLAN" in data:
        df_plan = data["PLAN"]
        row_global = df_plan[df_plan['Catégorie / Section'].astype(str).str.contains('GLOBAL', case=False, na=False)]
        if not row_global.empty:
            val = row_global.iloc[0]['% Atteinte']
            fig_gauge = go.Figure(go.Indicator(
                mode = "gauge+number",
                value = val,
                number = {'suffix': "%"},
                title = {'text': "Avancement Global"},
                gauge = {'axis': {'range': [None, 100]}, 'bar': {'color': "green"}}
            ))
            st.plotly_chart(fig_gauge, use_container_width=True)
        
        df_plan_show = df_plan.copy()
        if '% Atteinte' in df_plan_show.columns:
            df_plan_show['% Atteinte'] = df_plan_show['% Atteinte'].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float)) else x)
        st.dataframe(df_plan_show)
