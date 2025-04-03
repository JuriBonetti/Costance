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

# Funzione per aggiungere i dati al file Excel
def aggiungi_dati(file, nuova_riga):
    try:
        # Carica i dati
        df = carica_dati(file)
        
        # Converte nuova_riga in DataFrame
        nuova_riga_df = pd.DataFrame([nuova_riga])
        
        # Aggiungi i dati al DataFrame
        df = pd.concat([df, nuova_riga_df], ignore_index=True)
        
        # Usa 'openpyxl' per aprire il file Excel in modalita append senza creare nuovi fogli
        with pd.ExcelWriter(file, engine='openpyxl', mode='w') as writer:
            # Scrivi i dati nel foglio esistente (presumiamo che il foglio si chiami 'Sheet1')
            df.to_excel(writer, index=False, sheet_name='Sheet1')
    except Exception as e:
        st.error(f"Errore nell'aggiungere i dati al file: {e}")

# Funzione per calcolare la media del mese per il componente scelto
def calcola_media(file, mese_target, componente):
    try:
        df = carica_dati(file)
        
        # Converte la colonna 'data' in formato datetime
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        
        # Converti mese_target in pandas Timestamp
        mese_target = pd.to_datetime(mese_target).to_period('M')  # Converti in periodo mensile
        
        # Filtra per mese e componente
        df_filtro = df[(df['componente'] == componente) & (df['data'].dt.to_period('M') == mese_target)]
        
        # Calcola la media
        media = df_filtro['quantita'].mean() if not df_filtro.empty else "Nessun dato"
        
        return media
    except Exception as e:
        st.error(f"Errore nel calcolare la media: {e}")
        return None

    # Funzione per scrivere la media nel file Excel "File_DaPopolare"
def scrivi_media_su_file(file, componente, media):
    try:
        # Carica il file di destinazione
        wb = load_workbook(file)
        ws = wb.active
        
        # Scrivi la media nelle celle specifiche
        if componente == "Azoto":
            ws["B1"] = media
        elif componente == "COD":
            ws["B2"] = media
        elif componente == "Fosforo":
            ws["B3"] = media
        
        # Salva il file
        wb.save(file)
        wb.close()
    except Exception as e:
        st.error(f"Errore nell'aggiungere la media nel file di destinazione: {e}")

# Titolo dell'app
st.title("Costance_ENEA")

# File Excel di riferimento
file_dati = "Dati_Ingresso.xlsx"
file_risultati = "File_DaPopolare.xlsx"

# Sezione per inserire i nuovi dati
st.header("Inserisci i dati")
data_input = st.date_input("Data", min_value=pd.to_datetime("2020-01-01"))
df = carica_dati(file_dati)
componenti_unici = df['componente'].unique().tolist()  # Lista dei componenti gia presenti nel file Excel

# Crea un elenco a discesa per i componenti esistenti o inserisci un nuovo componente
componente_input = st.selectbox("Seleziona un componente", ["--Nuovo Componente--"] + componenti_unici)

# Se l'utente seleziona '--Nuovo Componente--', permette di inserire un nuovo componente
if componente_input == "--Nuovo Componente--":
    componente_input = st.text_input("Inserisci nuovo componente")
    if not componente_input:
        st.warning("Puoi aggiungere un nuovo componente, ma il campo non deve essere vuoto.")
else:
    st.write(f"Componente selezionato: {componente_input}")

# Aggiungi la quantita
quantita_input = st.number_input("Quantita", min_value=0.0, format="%.2f")

# Bottone per aggiungere i dati
if st.button("Aggiungi Dati"):
    if componente_input:
        nuova_riga = {"data": data_input, "componente": componente_input, "quantita": quantita_input}
        aggiungi_dati(file_dati, nuova_riga)
        st.success("Dati aggiunti con successo")
        df = carica_dati(file_dati)  # Ricarica i dati dal file per aggiornare il menu a discesa
    else:
        st.error("Il campo 'Nuovo componente' non deve essere vuoto.")

# Sezione per calcolare la media del mese
st.header("Calcola la Media del Mese")

# Seleziona mese per la media
mese_target = st.date_input("Seleziona il mese", value=pd.to_datetime("2025-04-03"))

# Ottieni la lista dei componenti unici dalla colonna "componente"
componenti_unici = df['componente'].unique().tolist()

# Seleziona il componente per cui calcolare la media
componente_media = st.selectbox("Seleziona il componente", componenti_unici)

# Bottone per calcolare la media e scrivere nei risultati
if st.button("Calcola Media Componente"):
    media_componente = calcola_media(file_dati, mese_target, componente=componente_media)
    if media_componente != "Nessun dato":
        st.write(f"La media di {componente_media} per {mese_target.strftime('%B %Y')} e': {media_componente}")
        scrivi_media_su_file(file_risultati, componente_media, media_componente)
        st.success(f"Media di {componente_media} scritta in '{file_risultati}'.")
    else:
        st.write(f"Nessun dato per {componente_media} in questo mese.")

# Visualizza i dati attuali
st.header("Visualizza Dati Attuali")
st.write(df)
