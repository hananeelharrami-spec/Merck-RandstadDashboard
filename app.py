import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import glob

# Configuration de la page
st.set_page_config(page_title="Dashboard Pilotage Randstad", layout="wide", initial_sidebar_state="expanded")

# --- EN-TÊTE ET SAISIE MANUELLE ---
col_logo, col_kpis = st.columns([1, 3])

with col_logo:
    st.title("📊 Pilotage")
    st.caption("Randstad / Merck")

with col_kpis:
    # Zone de saisie manuelle pour les données "Terrain"
    st.markdown("##### 📝 Données Hebdomadaires (Saisie Manuelle)")
    c1, c2 = st.columns(2)
    with c1:
        effectif = st.number_input("👥 Intérimaires en poste", value=133, step=1, help="Effectif actif cette semaine")
    with c2:
        nps = st.number_input("⭐ NPS Intérimaire", value=9.1, step=0.1, format="%.1f", help="Dernière note de satisfaction")

# --- FONCTION DE NETTOYAGE INTELLIGENTE ---
def clean_and_scale_data(df):
    # 0. Nettoyage des noms de colonnes (supprime les espaces invisibles)
    df.columns = df.columns.str.strip()

    # 1. Conversion des formats français (ex: "12,50%" -> 12.50)
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                # Nettoyage brutal des caractères invisibles et formatage
                s = df[col].astype(str).str.strip().str.replace('"', '').str.replace('\u202f', '').str.replace('\xa0', '')
                s = s.str.replace('%', '').str.replace(' ', '').str.replace(',', '.')
                # Conversion
                df[col] = pd.to_numeric(s, errors='coerce')
            except:
                pass

    # 2. Gestion de l'Année (Crucial pour 2026)
    if 'Année' in df.columns:
        # On convertit en nombre, on remplace les erreurs par 0
        df['Année'] = pd.to_numeric(df['Année'], errors='coerce').fillna(0).astype(int)

    # 3. Mise à l'échelle des pourcentages (0.88 -> 88.0)
    # On détecte les colonnes de taux qui sont restées en décimales
    for col in df.columns:
        col_lower = col.lower()
        keywords = ['taux', '%', 'atteinte', 'validation', 'rendement', 'impact']
        if any(x in col_lower for x in keywords):
            if pd.api.types.is_numeric_dtype(df[col]):
                max_val = df[col].max()
                # Si le max est petit (<= 1.5), c'est un ratio (ex: 0.88). On multiplie par 100.
                if pd.notna(max_val) and -1.5 <= max_val <= 1.5 and max_val != 0:
                    df[col] = df[col] * 100
    return df

# --- CHARGEMENT DES DONNÉES (UNIVERSEL : EXCEL OU CSV) ---
@st.cache_data
def load_data():
    data = {}
    source_info = "Aucune"
    
    # Stratégie 1 : Chercher un fichier Excel (.xlsx)
    excel_files = glob.glob("*.xlsx")
    if excel_files:
        file_path = excel_files[0]
        source_info = f"Excel ({file_path})"
        try:
            xls = pd.ExcelFile(file_path)
            # Mapping des onglets attendus
            mapping = {
                "YTD": "CONSOLIDATION_YTD",
                "RECRUT": "Recrutement_Mensuel",
                "ABS": "Absentéisme_Global_Mois",
                "ABS_MOTIF": "Absentéisme_Par_Motif",
                "ABS_SERVICE": "Absentéisme_Par_Service",
                "SOURCE": "KPI_Sourcing_Rendement",
                "PLAN": "Suivi_Plan_Action"
            }
            for key, sheet in mapping.items():
                if sheet in xls.sheet_names:
                    data[key] = clean_and_scale_data(pd.read_excel(xls, sheet_name=sheet))
        except Exception as e:
            st.error(f"Erreur Excel: {e}")

    # Stratégie 2 : Si pas d'Excel, chercher tous les CSV
    if not data:
        csv_files = glob.glob("*.csv")
        if csv_files:
            source_info = "CSV Multiples"
            for f in csv_files:
                try:
                    # Lecture tolérante (séparateur , ou ;)
                    df_temp = pd.read_csv(f, sep=None, engine='python')
                    df_clean = clean_and_scale_data(df_temp)
                    
                    # Détection du type de fichier par son nom
                    fname = f.lower()
                    if "consolidation" in fname: data["YTD"] = df_clean
                    elif "recrutement" in fname: data["RECRUT"] = df_clean
                    elif "global_mois" in fname: data["ABS"] = df_clean
                    elif "par_motif" in fname: data["ABS_MOTIF"] = df_clean
                    elif "par_service" in fname: data["ABS_SERVICE"] = df_clean
                    elif "sourcing" in fname: data["SOURCE"] = df_clean
                    elif "plan_action" in fname: data["PLAN"] = df_clean
                except:
                    pass

    return data, source_info

data, source_msg = load_data()

if not data:
    st.warning("⚠️ En attente de données.")
    st.info("Veuillez déposer le fichier 'data.xlsx' ou vos CSV exportés dans le dossier.")
    st.stop()
else:
    # st.toast(f"Source : {source_msg}", icon="✅") # Debug
    pass

st.markdown("---")

# --- BARRE LATÉRALE : SÉLECTEUR D'ANNÉE ---
st.sidebar.header("📅 Période")

# Récupération automatique des années disponibles (ex: 2025, 2026)
annees_dispo = set()
for key, df in data.items():
    if 'Année' in df.columns:
        # On filtre les valeurs aberrantes (0, NaN)
        valid_years = [int(y) for y in df['Année'].unique() if y > 2020]
        annees_dispo.update(valid_years)

annees_dispo = sorted(list(annees_dispo), reverse=True) # Descendant (2026 en premier)
options_annee = [str(a) for a in annees_dispo] + ["Vue Globale (Toutes)"]

# Par défaut : La plus récente (2026)
annee_select = st.sidebar.selectbox("Sélectionner l'exercice :", options_annee, index=0)

# Fonction de filtrage global
def filter_year(df):
    if "Vue Globale" in annee_select:
        return df
    if 'Année' in df.columns:
        return df[df['Année'] == int(annee_select)]
    return df

# --- DASHBOARD (ONGLETS) ---
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📈 Consolidation YTD", "🤝 Recrutement", "🏥 Absentéisme", "🔍 Sourcing & Talent Center", "✅ Plan d'Action"])

# ==========================================
# 1. CONSOLIDATION YTD
# ==========================================
with tab1:
    st.subheader(f"Performance Cumulée - {annee_select}")
    
    # Rappel des KPIs Manuels
    k1, k2 = st.columns(2)
    k1.metric("Effectifs Semaine", f"{effectif}", delta="Intérimaires")
    k2.metric("NPS Satisfaction", f"{nps}/10", delta="Note")
    st.markdown("---")

    if "YTD" in data:
        df_ytd = filter_year(data["YTD"])
        
        # Affichage en grille
        if not df_ytd.empty and 'Valeur YTD' in df_ytd.columns:
            # S'il y a plusieurs années (Vue Globale), on trie par Année DESC
            df_ytd = df_ytd.sort_values(by=['Année'], ascending=False)
            
            cols = st.columns(4)
            for i, (idx, row) in enumerate(df_ytd.iterrows()):
                indic = row['Indicateur']
                val = row['Valeur YTD']
                annee_row = row['Année']
                
                # Formatage
                val_str = f"{val:.2f}%" if isinstance(val, (int, float)) else str(val)
                label = f"{indic} ({annee_row})" if "Vue Globale" in annee_select else indic
                
                cols[i % 4].metric(label, val_str)
        else:
            st.info(f"Pas de données consolidées disponibles pour {annee_select}")

# ==========================================
# 2. RECRUTEMENT
# ==========================================
with tab2:
    st.header("Recrutement")
    if "RECRUT" in data:
        df_rec = filter_year(data["RECRUT"])
        if not df_rec.empty:
            if 'Mois' in df_rec.columns:
                df_rec = df_rec.sort_values(['Année', 'Mois'])
                df_rec['Période'] = df_rec['Mois'].astype(str) + "/" + df_rec['Année'].astype(str)

                col_g1, col_g2 = st.columns(2)
                
                with col_g1:
                    st.subheader("Qualité (Service & Transformation)")
                    fig = px.line(df_rec, x='Période', y=['Taux Service', 'Taux Transfo'], markers=True, 
                                  color_discrete_sequence=["#00CC96", "#EF553B"])
                    fig.update_layout(yaxis_ticksuffix="%")
                    st.plotly_chart(fig, use_container_width=True)

                with col_g2:
                    st.subheader("Volumes (Commandes vs Embauches)")
                    fig_bar = px.bar(df_rec, x='Période', y=['Nb Requisitions', 'Nb Hired'], barmode='group')
                    st.plotly_chart(fig_bar, use_container_width=True)
                
                with st.expander("Voir le tableau détaillé"):
                    st.dataframe(df_rec)
        else:
            st.warning("Aucune donnée pour la période sélectionnée.")

# ==========================================
# 3. ABSENTÉISME
# ==========================================
with tab3:
    st.header("Absentéisme")
    if "ABS" in data:
        df_abs = filter_year(data["ABS"])
        if not df_abs.empty:
            df_abs = df_abs.sort_values(['Année', 'Mois'])
            df_abs['Période'] = df_abs['Mois'].astype(str) + "/" + df_abs['Année'].astype(str)

            st.subheader("Taux Global Mensuel")
            fig_abs = px.area(df_abs, x='Période', y='Taux Absentéisme', markers=True, color_discrete_sequence=['#FF5733'])
            fig_abs.update_layout(yaxis_ticksuffix="%")
            st.plotly_chart(fig_abs, use_container_width=True)
            
            col_m, col_s = st.columns(2)
            
            with col_m:
                st.subheader("Répartition par Motif")
                if "ABS_MOTIF" in data:
                    df_mot = filter_year(data["ABS_MOTIF"])
                    if not df_mot.empty:
                        df_mot['Période'] = df_mot['Mois'].astype(str) + "/" + df_mot['Année'].astype(str)
                        fig_m = px.bar(df_mot, x='Période', y='Impact Motif (%)', color='Motif')
                        fig_m.update_layout(yaxis_ticksuffix="%")
                        st.plotly_chart(fig_m, use_container_width=True)
            
            with col_s:
                st.subheader("Comparatif Services")
                if "ABS_SERVICE" in data:
                    df_serv = filter_year(data["ABS_SERVICE"])
                    if not df_serv.empty:
                        df_serv['Période'] = df_serv['Mois'].astype(str) + "/" + df_serv['Année'].astype(str)
                        fig_s = px.bar(df_serv, x='Période', y='Taux Absentéisme', color='Service', barmode='group')
                        fig_s.update_layout(yaxis_ticksuffix="%")
                        st.plotly_chart(fig_s, use_container_width=True)

# ==========================================
# 4. SOURCING & TALENT CENTER
# ==========================================
with tab4:
    st.header("Sourcing")
    if "SOURCE" in data:
        df_src = filter_year(data["SOURCE"])
        
        if not df_src.empty and 'Source' in df_src.columns:
            # Normalisation du nom de la source
            df_src['Source_Clean'] = df_src['Source'].astype(str).str.upper().str.strip()
            
            # Aggrégation par Source sur la période
            df_agg = df_src.groupby('Source_Clean', as_index=False)[['1. Appels Reçus', '2. Validés (Sél.)', '3. Intégrés (Délégués)']].sum()
            
            # --- FOCUS TALENT CENTER ---
            st.subheader("🔥 Focus : Talent Center")
            # Recherche large du mot clé
            mask_tc = df_agg['Source_Clean'].str.contains("TALENT")
            df_tc = df_agg[mask_tc]
            
            if not df_tc.empty:
                # Somme des résultats (si plusieurs lignes Talent Center)
                v_tc = df_tc['1. Appels Reçus'].sum()
                val_tc = df_tc['2. Validés (Sél.)'].sum()
                int_tc = df_tc['3. Intégrés (Délégués)'].sum()
                rdt_tc = (int_tc / v_tc * 100) if v_tc > 0 else 0
                
                kc1, kc2, kc3, kc4 = st.columns(4)
                kc1.metric("1. Volume Appels", int(v_tc))
                kc2.metric("2. Validés", int(val_tc))
                kc3.metric("3. Intégrés", int(int_tc))
                kc4.metric("Rendement Final", f"{rdt_tc:.2f}%", delta="Talent Center")
            else:
                st.info("Aucune donnée 'Talent Center' trouvée sur cette période.")

            st.markdown("---")
            
            # --- TOP 5 SOURCES ---
            st.subheader("🏆 Top 5 Sources (par Intégrations)")
            
            # Tri par Intégrés puis Volume
            df_top = df_agg.sort_values(['3. Intégrés (Délégués)', '1. Appels Reçus'], ascending=False).head(5)
            
            # Coloriage spécial
            def color_logic(name):
                return "Talent Center" if "TALENT" in name else "Autres Sources"
            
            df_top['Type'] = df_top['Source_Clean'].apply(color_logic)
            
            fig_src = px.bar(
                df_top, 
                x='Source_Clean', 
                y='3. Intégrés (Délégués)', 
                color='Type',
                title="Nombre de recrutements réussis par source",
                text='3. Intégrés (Délégués)',
                color_discrete_map={
                    "Talent Center": "#FF4500",  # Orange Vif
                    "Autres Sources": "#1F77B4"  # Bleu Standard
                }
            )
            fig_src.update_traces(textposition='outside')
            st.plotly_chart(fig_src, use_container_width=True)
            
            with st.expander("Voir toutes les sources"):
                # Calcul du rendement pour le tableau
                df_agg['Rendement (%)'] = (df_agg['3. Intégrés (Délégués)'] / df_agg['1. Appels Reçus'] * 100).fillna(0).map('{:.2f}%'.format)
                st.dataframe(df_agg.sort_values('3. Intégrés (Délégués)', ascending=False))

# ==========================================
# 5. PLAN D'ACTION
# ==========================================
with tab5:
    st.header("Suivi du Plan d'Action")
    if "PLAN" in data:
        df_plan = data["PLAN"]
        
        # Jauge Globale
        row_global = df_plan[df_plan['Catégorie / Section'].astype(str).str.contains('GLOBAL', case=False, na=False)]
        
        if not row_global.empty:
            val = row_global.iloc[0]['% Atteinte']
            fig_gauge = go.Figure(go.Indicator(
                mode = "gauge+number",
                value = val,
                number = {'suffix': "%"},
                title = {'text': "Avancement Global du Projet"},
                gauge = {'axis': {'range': [None, 100]}, 'bar': {'color': "green"}}
            ))
            st.plotly_chart(fig_gauge, use_container_width=True)
        
        st.dataframe(df_plan)
