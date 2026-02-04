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

# --- FONCTION DE NETTOYAGE INTELLIGENTE ---
def clean_df(df):
    df.columns = df.columns.str.strip()
    
    # LISTE DES COLONNES A NE PAS CONVERTIR EN NOMBRE (Textes)
    # C'est ici que se situait le problème précédent
    protected_cols = ['Indicateur', 'Source', 'Service', 'Motif', 'Poste', 'Catégorie', 'Section']
    
    for col in df.columns:
        # Si le nom de la colonne contient un mot protégé, on ne touche pas
        if any(p.upper() in col.upper() for p in protected_cols):
            continue

        # Sinon, on tente de convertir en nombre
        if df[col].dtype == 'object':
            try:
                s = df[col].astype(str).str.strip().str.replace('"', '').str.replace('\u202f', '').str.replace('\xa0', '')
                s = s.str.replace('%', '').str.replace(' ', '').str.replace(',', '.')
                df[col] = pd.to_numeric(s, errors='coerce')
            except: pass
            
    # Mise à l'échelle des % (0.88 -> 88.0)
    for col in df.columns:
        if any(x in col.lower() for x in ['taux', '%', 'atteinte', 'rendement', 'validation']):
            if pd.api.types.is_numeric_dtype(df[col]):
                max_val = df[col].max()
                if pd.notna(max_val) and -1.5 <= max_val <= 1.5 and max_val != 0:
                    df[col] = df[col] * 100
                    
    # Gestion Année
    if 'Année' in df.columns:
        df['Année'] = df['Année'].fillna(0).astype(int)
        
    return df

# --- CHARGEMENT DES FICHIERS (CSV) ---
@st.cache_data
def load_data():
    data = {}
    
    # Mots-clés pour identifier les fichiers
    keywords_map = {
        "YTD": "CONSOLIDATION_YTD",
        "RECRUT": "Recrutement_Mensuel",
        "ABS": "Absentéisme_Global_Mois",
        "ABS_MOTIF": "Absentéisme_Par_Motif",
        "ABS_SERVICE": "Absentéisme_Par_Service",
        "SOURCE": "KPI_Sourcing", 
        "PLAN": "Suivi_Plan_Action"
    }
    
    # On prend tous les CSV du dossier
    all_csvs = glob.glob("*.csv")
    
    if not all_csvs:
        return None

    for key, keyword in keywords_map.items():
        for filename in all_csvs:
            if keyword in filename:
                try:
                    # Lecture robuste
                    df = pd.read_csv(filename, sep=None, engine='python')
                    data[key] = clean_df(df)
                except: pass
                # On ne break pas forcément ici pour laisser une chance aux autres fichiers si doublons
    return data

data = load_data()

if not data:
    st.error("❌ Aucun fichier CSV trouvé.")
    st.info("Les fichiers CSV (ex: 'Dashboard Merck... - CONSOLIDATION_YTD.csv') doivent être présents.")
    st.stop()

st.markdown("---")

# --- FILTRES ---
st.sidebar.header("Filtres")
annees = set()
for k, df in data.items():
    if 'Année' in df.columns:
        annees.update(df['Année'].unique())

annees = sorted([int(a) for a in annees if a > 2020], reverse=True)
opts = [str(a) for a in annees] + ["Vue Globale"]
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
        # On ne garde que les lignes où l'indicateur existe
        df = df.dropna(subset=['Indicateur'])
        
        if not df.empty and 'Valeur YTD' in df.columns:
            df = df.sort_values('Année', ascending=False)
            cols = st.columns(4)
            for i, r in df.iterrows():
                val = r['Valeur YTD']
                v_str = f"{val:.2f}%" if isinstance(val, (int, float)) else str(val)
                
                # --- PROTECTION ULTIME CONTRE LE CRASH ---
                lbl = str(r['Indicateur'])
                if lbl.lower() == "nan" or lbl.strip() == "": continue
                
                if annee_sel == "Vue Globale": lbl += f" ({r['Année']})"
                cols[i % 4].metric(lbl, v_str)
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
                fig = px.line(df, x='Période', y=['Taux Service', 'Taux Transfo'], markers=True)
                fig.update_layout(yaxis_ticksuffix="%")
                c1.plotly_chart(fig, use_container_width=True)
                fig2 = px.bar(df, x='Période', y=['Nb Requisitions', 'Nb Hired'], barmode='group')
                c2.plotly_chart(fig2, use_container_width=True)
                with st.expander("Détail"): st.dataframe(df)

# 3. ABSENTEISME
with t3:
    if "ABS" in data:
        df = filter(data["ABS"])
        if not df.empty:
            df = df.sort_values(['Année', 'Mois'])
            df['Période'] = df['Mois'].astype(str) + "/" + df['Année'].astype(str)
            fig = px.area(df, x='Période', y='Taux Absentéisme', markers=True, color_discrete_sequence=['#FF5733'])
            fig.update_layout(yaxis_ticksuffix="%")
            st.plotly_chart(fig, use_container_width=True)
            c1, c2 = st.columns(2)
            if "ABS_MOTIF" in data:
                dfm = filter(data["ABS_MOTIF"])
                if not dfm.empty:
                    dfm['Période'] = dfm['Mois'].astype(str) + "/" + dfm['Année'].astype(str)
                    figm = px.bar(dfm, x='Période', y='Impact Motif (%)', color='Motif')
                    figm.update_layout(yaxis_ticksuffix="%")
                    c1.plotly_chart(figm, use_container_width=True)
            if "ABS_SERVICE" in data:
                dfs = filter(data["ABS_SERVICE"])
                if not dfs.empty:
                    dfs['Période'] = dfs['Mois'].astype(str) + "/" + dfs['Année'].astype(str)
                    figs = px.bar(dfs, x='Période', y='Taux Absentéisme', color='Service', barmode='group')
                    figs.update_layout(yaxis_ticksuffix="%")
                    c2.plotly_chart(figs, use_container_width=True)

# 4. SOURCING
with t4:
    if "SOURCE" in data:
        df = filter(data["SOURCE"])
        # Normalisation Source
        if 'Source' in df.columns:
             df['Source_Clean'] = df['Source'].astype(str).str.upper().str.strip()
        
        # Aggrégation sécurisée
        if '1. Appels Reçus' in df.columns:
             df_agg = df.groupby('Source_Clean')[['1. Appels Reçus', '2. Validés (Sél.)', '3. Intégrés (Délégués)']].sum().reset_index()
        elif 'Retenu Présel.' in df.columns:
             # Fallback si fichier brut
             df_agg = df.groupby('Source_Clean').agg(
                 Appels=('Source_Clean', 'count'),
                 Valides=('Retenu Sél.', lambda x: x.astype(str).str.contains('OUI', case=False).sum()),
                 Integres=('Délégué', 'sum')
             ).reset_index()
             df_agg.columns = ['Source_Clean', '1. Appels Reçus', '2. Validés (Sél.)', '3. Intégrés (Délégués)']
        else:
             df_agg = pd.DataFrame()

        if not df_agg.empty:
            st.subheader("🔥 Focus Talent Center")
            mask_tc = df_agg['Source_Clean'].str.contains("TALENT")
            dftc = df_agg[mask_tc]
            if not dftc.empty:
                v = dftc['1. Appels Reçus'].sum()
                val = dftc['2. Validés (Sél.)'].sum()
                i = dftc['3. Intégrés (Délégués)'].sum()
                r = (i/v*100) if v>0 else 0
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Appels", int(v))
                c2.metric("Validés", int(val))
                c3.metric("Intégrés", int(i))
                c4.metric("Rendement", f"{r:.2f}%")
            else:
                st.warning("Pas de Talent Center.")
            
            st.markdown("---")
            st.subheader("🏆 Top 5 Sources")
            df_top = df_agg.sort_values(['3. Intégrés (Délégués)', '1. Appels Reçus'], ascending=False).head(5)
            df_top['Type'] = df_top['Source_Clean'].apply(lambda x: "Talent Center" if "TALENT" in x else "Autres")
            fig = px.bar(df_top, x='Source_Clean', y='3. Intégrés (Délégués)', color='Type',
                         text='3. Intégrés (Délégués)',
                         color_discrete_map={"Talent Center": "#FF4500", "Autres": "#1f77b4"})
            fig.update_traces(textposition='outside')
            st.plotly_chart(fig, use_container_width=True)
            with st.expander("Données"): st.dataframe(df_agg)

# 5. PLAN
with t5:
    if "PLAN" in data:
        df = data["PLAN"]
        row = df[df['Catégorie / Section'].astype(str).str.contains('GLOBAL', case=False, na=False)]
        if not row.empty:
            val = row.iloc[0]['% Atteinte']
            fig = go.Figure(go.Indicator(mode="gauge+number", value=val, number={'suffix':"%"}, 
                            title={'text':"Avancement"}, gauge={'axis':{'range':[None,100]}, 'bar':{'color':"green"}}))
            st.plotly_chart(fig, use_container_width=True)
        st.dataframe(df)
