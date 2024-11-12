import streamlit as st
import pandas as pd
from datetime import datetime
import io

# Fonction pour traiter les présences
def process_attendance_data(df):
    try:
        if not {'Nom', 'Heure'}.issubset(df.columns):
            st.error("Le fichier ne contient pas les colonnes 'Nom' et 'Heure' nécessaires.")
            return None
        df['Heure'] = pd.to_datetime(df['Heure'])
        df['Date'] = df['Heure'].dt.date
        df = df.sort_values(by=['Nom', 'Date', 'Heure'])
        df_filtered = df.groupby(['Nom', 'Date']).agg(
            Heure_Arrive=('Heure', 'first'),
            Heure_Sortie=('Heure', 'last')
        ).reset_index()
        df_filtered['Heure d\'arrive et de sortie'] = (
            df_filtered['Heure_Arrive'].dt.strftime('%H:%M:%S') + ' - ' + df_filtered['Heure_Sortie'].dt.strftime('%H:%M:%S')
        )
        return df_filtered[['Date', 'Nom', 'Heure d\'arrive et de sortie']]
    except Exception as e:
        st.error(f"Erreur dans le traitement du fichier: {str(e)}")
        return None

# Fonction pour traiter les absences
def process_absence_data(df):
    try:
        if not {'Nom', 'Heure'}.issubset(df.columns):
            st.error("Le fichier ne contient pas les colonnes 'Nom' et 'Heure' nécessaires.")
            return None, None, None
        df['Date'] = pd.to_datetime(df['Heure']).dt.date
        all_dates = pd.date_range(df['Date'].min(), df['Date'].max()).date
        unique_names = df['Nom'].unique()
        absence_data = []
        for name in unique_names:
            dates_present = set(df[df['Nom'] == name]['Date'].unique())
            dates_absent = [date for date in all_dates if date not in dates_present]
            for date in dates_absent:
                absence_data.append({'Nom': name, 'Date': date})
        absence_df = pd.DataFrame(absence_data)
        absence_df['Semaine'] = pd.to_datetime(absence_df['Date']).dt.isocalendar().week
        absence_df['Mois'] = pd.to_datetime(absence_df['Date']).dt.month
        absence_summary = absence_df.groupby(['Nom', 'Semaine']).size().reset_index(name='Absences_Semaine')
        absence_summary_month = absence_df.groupby(['Nom', 'Mois']).size().reset_index(name='Absences_Mois')
        return absence_df, absence_summary, absence_summary_month
    except Exception as e:
        st.error(f"Erreur dans le traitement des absences: {str(e)}")
        return None, None, None

# Fonction pour générer le rapport
def generate_report(df, period):
    try:
        df['Date'] = pd.to_datetime(df['Date'])
        if period == 'Jour':
            report = df.groupby(df['Date'].dt.date).size().reset_index(name="Nombre de présences")
        elif period == 'Semaine':
            report = df.groupby(df['Date'].dt.isocalendar().week).size().reset_index(name="Nombre de présences")
        elif period == 'Mois':
            report = df.groupby(df['Date'].dt.to_period('M')).size().reset_index(name="Nombre de présences")
        elif period == 'Trimestre':
            report = df.groupby(df['Date'].dt.to_period('Q')).size().reset_index(name="Nombre de présences")
        elif period == 'Année':
            report = df.groupby(df['Date'].dt.year).size().reset_index(name="Nombre de présences")
        else:
            return None
        report.columns = [period, "Nombre de présences"]
        return report
    except Exception as e:
        st.error(f"Erreur dans la génération du rapport: {str(e)}")
        return None

# Interface utilisateur Streamlit
st.title("Système de Gestion de Présences et d'Absences - CHU")
st.markdown("Cette application vous permet de gérer les données de présence et d'absence du personnel, ainsi que de générer des rapports basés sur les périodes.")

uploaded_file = st.file_uploader("📁 Importer un fichier Excel", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        if not sheet_names:
            st.error("Le fichier Excel doit contenir au moins une feuille visible.")
        else:
            df = pd.read_excel(xls, sheet_name=sheet_names[0])
            st.success("Fichier importé avec succès.")
            
            # Initialisation des variables
            presence_data = None
            absence_data = None
            absence_summary = None
            absence_summary_month = None
            report_data = None
            
            # Affichage des présences
            if st.button("Afficher les présences"):
                presence_data = process_attendance_data(df)
                if presence_data is not None:
                    st.subheader("Données de Présence")
                    st.dataframe(presence_data)

            # Affichage des absences
            if st.button("Afficher les absences"):
                absence_data, absence_summary, absence_summary_month = process_absence_data(df)
                if absence_data is not None:
                    st.subheader("Données d'Absence")
                    st.dataframe(absence_data)
                    st.subheader("Résumé des Absences par Semaine")
                    st.dataframe(absence_summary)
                    st.subheader("Résumé des Absences par Mois")
                    st.dataframe(absence_summary_month)

            # Génération du rapport
            period = st.selectbox("Choisissez la période pour le rapport:", ["Jour", "Semaine", "Mois", "Trimestre", "Année"])

            if st.button("Générer le rapport"):
                report_data = generate_report(df, period)
                if report_data is not None:
                    st.subheader(f"Rapport de Présences - Période: {period}")
                    st.dataframe(report_data)

            # Génération et téléchargement du fichier Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                if presence_data is not None:
                    presence_data.to_excel(writer, sheet_name="Présences", index=False)
                if absence_data is not None:
                    absence_data.to_excel(writer, sheet_name="Absences", index=False)
                if absence_summary is not None:
                    absence_summary.to_excel(writer, sheet_name="Absences par Semaine", index=False)
                if absence_summary_month is not None:
                    absence_summary_month.to_excel(writer, sheet_name="Absences par Mois", index=False)
                if report_data is not None:
                    report_data.to_excel(writer, sheet_name="Rapport", index=False)
            output.seek(0)

            st.download_button(
                label="📥 Télécharger le fichier Excel traité",
                data=output,
                file_name="Rapport_Presences.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Erreur lors de l'importation du fichier: {str(e)}")
else:
    st.info("Veuillez importer un fichier pour commencer.")
