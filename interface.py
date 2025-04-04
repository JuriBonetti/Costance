import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from datetime import datetime

# Funzione per caricare i dati esistenti dal file Excel
def carica_dati(file):
    try:
        df = pd.read_excel(file)
    except Exception as e:
        st.error(f"Errore nel caricare il file: {e}")
        return pd.DataFrame()  # Restituisce un DataFrame vuoto se c'e un errore
    return df

# Funzione per calcolare la media del mese per il componente scelto
def calcola_media(file, mese_target, componente):
    try:
        df = carica_dati(file)
        
        # Converte la colonna 'data' in formato datetime
        df['Data Prelievo'] = pd.to_datetime(df['Data Prelievo'], dayfirst=True)
        
        # Converti mese_target in pandas Timestamp
        mese_target = pd.to_datetime(mese_target).to_period('M')  # Converti in periodo mensile
        
        # Filtra per mese e componente
        df_filtro = df[(df['Nome parametro'] == componente) & (df['Data Prelievo'].dt.to_period('M') == mese_target)]
        
        # Calcola la media
        media = df_filtro['Risultato numerico'].mean() if not df_filtro.empty else "Nessun dato"
        
        return media
    except Exception as e:
        st.error(f"Errore nel calcolare la media: {e}")
        return None

# Funzione per scrivere la media nel file Excel "File_DaPopolare"
def scrivi_media_su_file(file, cella, media):
    try:
        # Carica il file di destinazione
        wb = load_workbook(file)
        ws = wb.active
        
        # Scrivi la media nella cella specificata
        ws[cella] = media
        
        # Salva il file
        wb.save(file)
        wb.close()
    except Exception as e:
        st.error(f"Errore nell'aggiungere la media nel file di destinazione: {e}")

# Titolo dell'app
st.title("Costance_ENEA")

# Sezione per caricare il file Excel
st.header("Carica il File Excel")

# File Excel di riferimento
file_dati = "data_8.xlsx"
file_risultati = "File_DaPopolare.xlsx"

# Carica il file Excel tramite uploader
uploaded_file = st.file_uploader("Carica il tuo file Excel", type=["xlsx"])

# Se il file e stato caricato, carica e mostra il contenuto
if uploaded_file:
    df_caricato = pd.read_excel(uploaded_file)
    st.write("Contenuto del file Excel caricato:")
    st.dataframe(df_caricato)
    
    # Salva il file caricato in una variabile globale
    file_dati = uploaded_file

    # Estrai i componenti unici dal file caricato
    componenti_disponibili = df_caricato['Nome parametro'].unique().tolist()

    # Inizializza la tabella (se non e stata ancora creata in sessione)
    if "tabella_dati" not in st.session_state:
        st.session_state.tabella_dati = pd.DataFrame(columns=["Nome parametro", "Mese", "Cella Excel"])

    # Crea una form per aggiungere nuove righe alla tabella
    col1, col2, col3 = st.columns(3)

    with col1:
        # Seleziona il componente dalla lista
        componente_input = st.selectbox("Nome parametro", componenti_disponibili)

    with col2:
        mese_input = st.date_input("Mese", value=datetime.today())

    with col3:
        cella_input = st.text_input("Cella Excel")

    # Bottone per aggiungere una riga alla tabella
    if st.button("Aggiungi Dati"):
        if componente_input and cella_input:
            nuova_riga = pd.DataFrame([[componente_input, mese_input, cella_input]], columns=["Nome parametro", "Mese", "Cella Excel"])
            
            # Aggiungi la nuova riga alla tabella memorizzata in session_state
            st.session_state.tabella_dati = pd.concat([st.session_state.tabella_dati, nuova_riga], ignore_index=True)
            
            st.success("Riga aggiunta alla tabella.")
        else:
            st.error("Completa tutti i campi prima di aggiungere i dati.")

    # Visualizza la tabella aggiornata con opzioni per eliminare le righe
    st.dataframe(st.session_state.tabella_dati)

    # Seleziona quale riga eliminare dalla tabella
    st.subheader("Elimina una riga dalla tabella")
    riga_da_eliminare = st.selectbox("Seleziona la riga da eliminare", options=st.session_state.tabella_dati.index.tolist())
    if st.button("Elimina Riga"):
        if riga_da_eliminare is not None:
            st.session_state.tabella_dati = st.session_state.tabella_dati.drop(riga_da_eliminare).reset_index(drop=True)
            st.success(f"Riga {riga_da_eliminare} eliminata.")
        else:
            st.error("Seleziona una riga da eliminare.")

    # Sezione per calcolare la media del mese e scrivere la media nel file Excel
    st.header("Calcola la Media e Scrivi nel File Excel")

    # Bottone per calcolare la media e scrivere i risultati nel file Excel
    if st.button("Calcola e Scrivi Media"):
        for index, row in st.session_state.tabella_dati.iterrows():
            componente = row["Nome parametro"]
            mese = row["Mese"]
            cella = row["Cella Excel"]
            
            media_componente = calcola_media(file_dati, mese, componente)
            
            if media_componente != "Nessun dato":
                scrivi_media_su_file(file_risultati, cella, media_componente)
                st.success(f"Media di {componente} per il mese {mese.strftime('%B %Y')} scritta nella cella {cella}.")
            else:
                st.write(f"Nessun dato per {componente} in {mese.strftime('%B %Y')}.")

# Visualizza i dati attuali (opzionale)
st.header("Visualizza Dati Attuali")
df = carica_dati(file_dati)
st.write(df)
