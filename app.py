import streamlit as st
from st_aggrid import AgGrid
import pandas as pd
from datetime import datetime
import io
import sys

date = datetime.now().strftime("%d-%m-%Y")

st.set_page_config(layout='wide', initial_sidebar_state='expanded')

st.title('GeoVet Mailing List Cleaner ')
st.markdown("""
    L'app esegue le seguenti operazioni per produrre l'output:
    1) crea campo "da_sito" in *statgis_latest.xlsx* (da_sito = 'NO') e *udanet_latest.csv* (da_sito = 'YES')
    2) elimina i doppioni dai due file citati nel punto precedente 
    3) concatena e pulisce il dato concatenato dai doppioni, tenendo l'occorrenza presente in *statgis_latest.xlsx* (quindi quella che contiene tutte le info oltre all'indirizzo email)
    3) rimuove dal dato concatenato tutti gli indirizzi presenti nel file *unsubscribe_latest.xls*
    4) mostra a schermo il dato pulito sotto forma di tabella e permette di scaricarlo in formato excel 
            """)

col1, col2, col3 = st.columns([1,1,1])

statgis_file = col1.file_uploader("Carica il file **statgis_latest.xlsx**", type=["xls","xlsx"], accept_multiple_files=False)
udanet_file = col2.file_uploader("Carica il file **udanet_latest.csv**", type="csv", accept_multiple_files=False)
unsubscribe_file = col3.file_uploader("Carica il file **unsubscribe_latest.xls**", type=["xls","xlsx"], accept_multiple_files=False)
    
def create_final_df(statgis_file, udanet_file, unsubscribe_file):
    # StatGIS
    statgis_df = pd.read_excel(statgis_file)
    statgis_df.drop(["da_sito"], axis=1, errors="ignore") # Cancella il campo da_sito se esiste già
    statgis_df["da_sito"] = "NO" # Aggiunge campo "da_sito"
    statgis_df["email"] = statgis_df["email"].str.strip()
    statgis_df["Aggiornata"]= statgis_df["Aggiornata"].dt.strftime("%d-%m-%Y")
    statgis_df.drop_duplicates(subset=["email"], inplace=True)
    col1.metric('StatGIS', len(statgis_df))
    # UdaNet
    udanet_df = pd.read_csv(udanet_file)
    udanet_df_mail = udanet_df[["Email"]]
    udanet_df_mail.rename(columns={'Email': 'email'}, inplace=True)
    udanet_df_mail["da_sito"] = "YES" # Aggiunge campo "da_sito"
    udanet_df_mail["email"] = udanet_df_mail["email"].str.strip()
    udanet_df_mail["Aggiornata"] = date
    udanet_df_mail.drop_duplicates(subset=["email"], inplace=True)
    col2.metric('UdaNet', len(udanet_df_mail))
    # Unsubscribers
    try:
        unsub_df = pd.read_excel(unsubscribe_file)
        unsub_df.rename(columns={'EMAIL': 'email'}, inplace=True)
        unsub_df.columns = ["email"]
        col3.metric('Disiscritti', len(unsub_df))
    except:
        st.error("**ERRORE**: Il file delle disiscrizioni fornito non è un vero Excel. Aprilo con MS Excel e salvalo nuovamente con estensione .xls o .xlsx per risolvere il problema")
        sys.exit(1)
    # Concatenazione dei dataframe Udanet e StatGIS
    df_concat = pd.concat([udanet_df_mail, statgis_df], ignore_index=True)
    df_concat.sort_values(by="da_sito", ascending=True, inplace=True)
    # Pulizia eventuali duplicati tenendo il primo (che in base all'ordinamento ascendente sul campo "da_sito", sarà quello presente nel file StatGIS)
    df_unique = df_concat.drop_duplicates(subset=['email'], keep="first")
    # Pulizia indirizzi presenti nel dataframe dei disicritti
    unsubscribers = unsub_df["email"].to_list()
    df_clean = df_unique.query('email != @unsubscribers')
    
    return df_clean

if statgis_file is not None and udanet_file is not None and unsubscribe_file is not None:
    df = create_final_df(statgis_file, udanet_file, unsubscribe_file)
    
    # st.dataframe(df.sort_values(by="email", ascending=True), height=400)
    AgGrid(df.sort_values(by="email", ascending=True), height=250, theme='alpine')
    
    st.success('Il risultato contiene **{}** indirizzi'.format(len(df)))
    
    # Aggiunta della funzionalità di download
    file_name = 'ml_final_output_'+date+'.xlsx'
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
  
        df.to_excel(writer, sheet_name='ML', index=False)
        # Close the Pandas Excel writer and output the Excel file to the buffer
        writer.save()

        download = st.download_button(
            label="Scarica Excel",
            data=buffer,
            file_name=file_name,
            mime='application/vnd.ms-excel'
        )

else:
    st.markdown("In attesa dei file di input...")
    
    
