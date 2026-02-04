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

# --- NETTOYAGE ---
def clean_df(df):
    df.columns = df.columns.str.strip()
    
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                s = df[col].astype(str).str.strip().str.replace('"', '').str.replace('\u202f', '').str.replace('\xa0', '')
                s = s.str.replace('%', '').str.replace(' ', '').str.replace(',', '.')
                df[col] = pd.to_numeric(s, errors='coerce')
            except: pass
            
    for col in df.columns:
        if any(x in col.lower() for x in ['taux', '%', 'atteinte', 'rendement']):
            if pd.api.types.is_numeric_dtype(df[col]):
                if -1.5 <= df[col].max() <= 1.5 and df[col].max() != 0:
                    df[col] = df[col] * 100
                    
    if 'Année' in df.columns:
        df['Année'] = df['Année'].fillna(0).astype(int)
    
    # Sécurisation des colonnes textes importantes
    if 'Indicateur' in df.columns:
        df['Indicateur'] = df['Indicateur'].astype(str)
        
    return df

# --- CHARGEMENT (COMPATIBLE CSV) ---
@st.cache_data
def load_data():
    data = {}
    
    # Mapping des fichiers par mot-clé
    keywords = {
        "YTD": "CONSOLIDATION_YTD",
        "RECRUT": "Recrutement_Mensuel",
        "ABS": "Absentéisme_Global_Mois",
        "ABS_MOTIF": "Absentéisme_Par_Motif",
        "ABS_SERVICE": "Absentéisme_Par_Service",
        "SOURCE": "KPI_Sourcing", # ou MerckPresel
        "PLAN": "Suivi_Plan_Action"
    }
    
    # On prend tous les CSV
    all_files = glob.glob("*.csv")
    
    if not all_files:
        return None

    for key, keyword in keywords.items():
        for f in all_files:
            if keyword in f:
                try:
                    df = pd.read_csv(f, sep=None, engine='python')
                    data[key] = clean_df(df)
                    
                    # Traitement spécial Sourcing si c'est le fichier brut
                    if key == "SOURCE" and "Date Appel" in df.columns:
                        df["Date_Obj"] = pd.to_datetime(df["Date Appel"], dayfirst=True, errors='coerce')
                        df["Année"] = df["Date_Obj"].dt.year.fillna(2025).astype(int)
                        if "Source" in df.columns:
                            df["Source_Clean"] = df["Source"].astype(str).str.upper().str.strip()
                    
                except: pass
                break 
    return data

data = load_data()

if not data:
    st.error("❌ Aucun fichier CSV trouvé.")
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

# 1. GLOBAL YTD (CORRECTIF TYPEERROR ICI)
with t1:
    st.subheader(f"Performance YTD - {annee_sel}")
    if "YTD" in data:
        df = filter(data["YTD"])
        if not df.empty and 'Valeur YTD' in df.columns:
            df = df.sort_values('Année', ascending=False)
            cols = st.columns(4)
            for i, r in df.iterrows():
                val = r['Valeur YTD']
                v_str = f"{val:.2f}%" if isinstance(val, (int, float)) else str(val)
                
                # --- LE FIX EST ICI : str() ---
                lbl = str(r['Indicateur']) 
                if lbl == "nan": lbl = "Indicateur"
                
                if annee_sel == "Vue Globale": lbl += f" ({r['Année']})"
                cols[i % 4].metric(lbl, v_str)
        else:
            st.info("Pas de données YTD.")

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
        # Gestion hybride : soit fichier brut (Presel), soit déjà calculé (KPI)
        # On normalise la colonne source si elle existe
        if 'Source' in df.columns:
             df['Source_Clean'] = df['Source'].astype(str).str.upper().str.strip()
        
        # Cas fichier déjà calculé (KPI_Sourcing_Rendement.csv)
        if '1. Appels Reçus' in df.columns:
             df_agg = df.groupby('Source_Clean')[['1. Appels Reçus', '2. Validés (Sél.)', '3. Intégrés (Délégués)']].sum().reset_index()
        # Cas fichier brut (MerckPresel)
        elif 'Retenu Présel.' in df.columns:
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
