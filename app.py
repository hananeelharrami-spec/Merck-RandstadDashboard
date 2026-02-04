import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import glob

# Configuration de la page
st.set_page_config(page_title="Dashboard Pilotage Randstad", layout="wide")

# --- EN-TÊTE AVEC KPI MANUELS ---
col_title, col_kpis = st.columns([2, 2])
with col_title:
    st.title("📊 Dashboard de Pilotage")
    st.caption("Randstad / Merck")

with col_kpis:
    # Zone de saisie manuelle (côte à côte)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### 👥 Effectifs")
        effectif = st.number_input("Intérimaires en poste", value=133, step=1)
    with c2:
        st.markdown("### ⭐ Satisfaction")
        nps = st.number_input("NPS Intérimaire (/10)", value=9.1, step=0.1, format="%.1f")

# --- FONCTION DE NETTOYAGE ROBUSTE ---
def clean_and_scale_data(df):
    # 0. Nettoyage des Noms de Colonnes
    df.columns = df.columns.str.strip()

    # 1. CONVERSION TEXTE -> NOMBRE (Sécurisée)
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                # On travaille sur une copie
                series = df[col].astype(str).str.strip()
                # Enlève les guillemets et espaces insécables
                series = series.str.replace('"', '').str.replace('\u202f', '').str.replace('\xa0', '')
                # Enlève % et remplace virgule
                series = series.str.replace('%', '').str.replace(' ', '').str.replace(',', '.')
                
                # Tente la conversion, remplace les erreurs par NaN
                df[col] = pd.to_numeric(series, errors='coerce')
            except Exception:
                pass

    # 2. SECURISATION DES ANNEES (Anti-Crash)
    if 'Année' in df.columns:
        # On force en numérique
        df['Année'] = pd.to_numeric(df['Année'], errors='coerce')
        # On remplit les NaN par une année par défaut (ex: 2025) pour ne pas perdre la donnée
        df['Année'] = df['Année'].fillna(2025).astype(int)

    # 3. MISE A L'ECHELLE DES POURCENTAGES
    for col in df.columns:
        col_lower = col.lower()
        keywords = ['taux', '%', 'atteinte', 'validation', 'rendement', 'impact']
        if any(x in col_lower for x in keywords):
            if pd.api.types.is_numeric_dtype(df[col]):
                max_val = df[col].max()
                # Si max est petit (ex: 0.88), on x100
                if pd.notna(max_val) and -1.5 <= max_val <= 1.5 and max_val != 0:
                    df[col] = df[col] * 100
    return df

# --- CHARGEMENT ---
@st.cache_data
def load_data():
    excel_files = glob.glob("*.xlsx")
    if not excel_files:
        return None, None

    found_file = excel_files[0]
    data = {}
    
    expected = {
        "YTD": "CONSOLIDATION_YTD",
        "RECRUT": "Recrutement_Mensuel",
        "ABS": "Absentéisme_Global_Mois",
        "ABS_MOTIF": "Absentéisme_Par_Motif",
        "ABS_SERVICE": "Absentéisme_Par_Service",
        "SOURCE": "KPI_Sourcing_Rendement",
        "PLAN": "Suivi_Plan_Action"
    }
    
    try:
        xls = pd.ExcelFile(found_file)
        all_sheets = xls.sheet_names
        
        for key, sheet_name in expected.items():
            if sheet_name in all_sheets:
                # Lecture brute puis nettoyage
                df_raw = pd.read_excel(found_file, sheet_name=sheet_name)
                data[key] = clean_and_scale_data(df_raw)
        return data, found_file
        
    except Exception as e:
        st.error(f"Erreur technique lors de la lecture : {e}")
        return None, found_file

data, filename = load_data()

if data is None:
    st.error("❌ Aucun fichier Excel trouvé. Uploadez data.xlsx sur GitHub.")
    st.stop()
else:
    st.markdown("---")

# --- BARRE LATÉRALE : FILTRES ---
st.sidebar.header("Filtres")

# Récupération dynamique des années
annees_dispo = set()
for key, df in data.items():
    if 'Année' in df.columns:
        try:
            unique_years = df['Année'].dropna().unique()
            # On ne garde que ce qui ressemble à une année (ex: > 2020)
            valid_years = [int(y) for y in unique_years if y > 2020]
            annees_dispo.update(valid_years)
        except:
            pass

annees_dispo = sorted(list(annees_dispo))
# Par défaut, on met l'année la plus récente si disponible, sinon Toutes
index_default = len(annees_dispo) # Par défaut "Toutes" (qui sera à la fin de la liste d'options)

options_annee = [str(a) for a in annees_dispo] + ["Toutes"]
annee_select = st.sidebar.selectbox("Choisir l'année :", options_annee, index=len(options_annee)-1)

def filter_year(df):
    if annee_select == "Toutes":
        return df
    if 'Année' in df.columns:
        return df[df['Année'] == int(annee_select)]
    return df

# --- DASHBOARD ---

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📈 Vue Globale", "🤝 Recrutement", "🏥 Absentéisme", "🔍 Sourcing", "✅ Plan d'Action"])

# --- 1. VUE GLOBALE ---
with tab1:
    st.subheader(f"Performance Consolidée ({annee_select})")
    
    # Affichage des KPIs manuels en gros ici aussi
    k1, k2 = st.columns(2)
    k1.metric("Effectifs Semaine", f"{effectif}", delta="Intérimaires")
    k2.metric("NPS Intérimaire", f"{nps}/10", delta="Satisfaction")
    st.markdown("---")

    if "YTD" in data:
        df_ytd = filter_year(data["YTD"])
        if not df_ytd.empty and 'Valeur YTD' in df_ytd.columns:
            cols = st.columns(4)
            for index, row in df_ytd.iterrows():
                indic = row['Indicateur']
                val = row['Valeur YTD']
                val_str = f"{val:.2f}%" if isinstance(val, (int, float)) else str(val)
                cols[index % 4].metric(label=indic, value=val_str)
        else:
            st.info(f"Pas de données YTD pour {annee_select}")

# --- 2. RECRUTEMENT ---
with tab2:
    st.header("Recrutement")
    if "RECRUT" in data:
        df_rec = filter_year(data["RECRUT"])
        if not df_rec.empty:
            if 'Mois' in df_rec.columns:
                df_rec = df_rec.sort_values(['Année', 'Mois'])
                df_rec['Période'] = df_rec['Mois'].astype(str) + "/" + df_rec['Année'].astype(str)

                c1, c2 = st.columns(2)
                with c1:
                    st.subheader("Taux Qualitatifs")
                    fig = px.line(df_rec, x='Période', y=['Taux Service', 'Taux Transfo'], markers=True)
                    fig.update_layout(yaxis_ticksuffix="%")
                    st.plotly_chart(fig, use_container_width=True)

                with c2:
                    st.subheader("Volumes")
                    fig_bar = px.bar(df_rec, x='Période', y=['Nb Requisitions', 'Nb Hired'], barmode='group')
                    st.plotly_chart(fig_bar, use_container_width=True)
                
                with st.expander("Détail"):
                    st.dataframe(df_rec)
        else:
            st.warning(f"Aucune donnée Recrutement pour {annee_select}")

# --- 3. ABSENTÉISME ---
with tab3:
    st.header("Absentéisme")
    if "ABS" in data:
        df_abs = filter_year(data["ABS"])
        if not df_abs.empty:
            df_abs = df_abs.sort_values(['Année', 'Mois'])
            df_abs['Période'] = df_abs['Mois'].astype(str) + "/" + df_abs['Année'].astype(str)

            st.subheader("Taux Global")
            fig_abs = px.area(df_abs, x='Période', y='Taux Absentéisme', markers=True, color_discrete_sequence=['#FF5733'])
            fig_abs.update_layout(yaxis_ticksuffix="%")
            st.plotly_chart(fig_abs, use_container_width=True)
            
            c_mot, c_serv = st.columns(2)
            with c_mot:
                st.subheader("Par Motif")
                if "ABS_MOTIF" in data:
                    df_mot = filter_year(data["ABS_MOTIF"])
                    if not df_mot.empty:
                         df_mot['Période'] = df_mot['Mois'].astype(str) + "/" + df_mot['Année'].astype(str)
                         fig_m = px.bar(df_mot, x='Période', y='Impact Motif (%)', color='Motif')
                         fig_m.update_layout(yaxis_ticksuffix="%")
                         st.plotly_chart(fig_m, use_container_width=True)
            
            with c_serv:
                st.subheader("Par Service")
                if "ABS_SERVICE" in data:
                    df_serv = filter_year(data["ABS_SERVICE"])
                    if not df_serv.empty:
                         df_serv['Période'] = df_serv['Mois'].astype(str) + "/" + df_serv['Année'].astype(str)
                         fig_s = px.bar(df_serv, x='Période', y='Taux Absentéisme', color='Service', barmode='group')
                         fig_s.update_layout(yaxis_ticksuffix="%")
                         st.plotly_chart(fig_s, use_container_width=True)

# --- 4. SOURCING ---
with tab4:
    st.header("Sourcing")
    if "SOURCE" in data:
        df_src = filter_year(data["SOURCE"])
        if not df_src.empty and 'Source' in df_src.columns:
            # Nettoyage colonne Source
            df_src['Source_Clean'] = df_src['Source'].astype(str).str.upper().str.strip()
            
            # Aggrégation
            df_agg = df_src.groupby('Source_Clean', as_index=False)[['1. Appels Reçus', '2. Validés (Sél.)', '3. Intégrés (Délégués)']].sum()
            
            # FOCUS TC (Recherche large)
            st.subheader("🔥 Focus : Talent Center")
            mask_tc = df_agg['Source_Clean'].str.contains("TALENT")
            df_tc = df_agg[mask_tc]
            
            if not df_tc.empty:
                # Somme au cas où plusieurs lignes matchent
                vol_tc = df_tc['1. Appels Reçus'].sum()
                val_tc = df_tc['2. Validés (Sél.)'].sum()
                int_tc = df_tc['3. Intégrés (Délégués)'].sum()
                taux_tc = (int_tc / vol_tc * 100) if vol_tc > 0 else 0
                
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("Volume", int(vol_tc))
                k2.metric("Validés", int(val_tc))
                k3.metric("Intégrés", int(int_tc))
                k4.metric("Rendement", f"{taux_tc:.2f}%")
            else:
                st.info("Source 'Talent Center' non détectée sur cette période.")

            st.markdown("---")
            
            # TOP 5
            st.subheader("🏆 Top 5 Sources")
            df_top5 = df_agg.sort_values(['3. Intégrés (Délégués)', '1. Appels Reçus'], ascending=False).head(5)
            
            def get_color(name):
                return "Talent Center" if "TALENT" in name else "Autres"
            
            df_top5['Type'] = df_top5['Source_Clean'].apply(get_color)
            
            fig_best = px.bar(
                df_top5, 
                x='Source_Clean', 
                y='3. Intégrés (Délégués)', 
                color='Type',
                title="Intégrations par Source",
                text='3. Intégrés (Délégués)',
                color_discrete_map={"Talent Center": "#FF4500", "Autres": "#1f77b4"}
            )
            fig_best.update_traces(textposition='outside')
            st.plotly_chart(fig_best, use_container_width=True)

# --- 5. PLAN D'ACTION ---
with tab5:
    st.header("Plan d'Action")
    if "PLAN" in data:
        df_plan = data["PLAN"]
        # Pas de filtre année ici car le plan est souvent transverse
        
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
            
        st.dataframe(df_plan)
