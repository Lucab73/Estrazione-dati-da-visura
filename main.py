import streamlit as st
import pandas as pd
import openpyxl
from PyPDF2 import PdfReader
import re

# Funzione per estrarre i dati
def estrai_dati(filepath):
    # Caricamento del PDF
    reader = PdfReader (filepath)
    text = ""
    for page in reader.pages:
        text += page.extract_text ()

    righe = text.splitlines ()

    # Debug: Stampa le prime 500 righe del testo
    print("\n".join(text.splitlines()[:500]))  # Stampa le prime 500 righe del testo

    # Lista delle sezioni che determinano la fine della ricerca
    sezioni_fine = [
        "Trasferimenti d'azienda, fusioni, scissioni, subentri",
        "Attivita', albi ruoli e licenze",
        "Storia delle modifiche"
    ]

    # Trova la seconda occorrenza di una qualsiasi delle sezioni di fine
    occorrenze_sezioni = {}
    for sezione in sezioni_fine:
        occorrenze = [i for i, riga in enumerate (righe) if sezione in riga]
        if len (occorrenze) >= 2:
            occorrenze_sezioni[sezione] = occorrenze[1]  # Prendi la seconda occorrenza

    if not occorrenze_sezioni:
        print ("Non è stata trovata una seconda occorrenza di nessuna delle sezioni di fine.")
        return []

    # Prendi la prima seconda occorrenza tra tutte le sezioni trovate
    riga_fine = min (occorrenze_sezioni.values ())

    # Limita le righe fino alla seconda occorrenza della prima sezione di fine trovata
    righe = righe[:riga_fine]
    testo_completo = "\n".join (righe)

    # Lista delle possibili sezioni da cercare
    sezioni_da_cercare = [
        "Soci e titolari di diritti su azioni e quote",
        "Amministratori",
        "Sindaci, membri organi di controllo",
        "Titolari di altre cariche o qualifiche",
        "Titolari di cariche o qualifiche"

        # Aggiungi qui altre sezioni che potrebbero essere presenti
    ]

    # Trova tutte le sezioni presenti nel testo
    sezioni_trovate = []
    testo_sezioni = {}

    for i, sezione in enumerate (sezioni_da_cercare):
        indici = [m.start () for m in re.finditer (re.escape (sezione), testo_completo)]
        for indice in indici:
            sezioni_trovate.append ((indice, sezione))

    # Ordina le sezioni per posizione nel testo
    sezioni_trovate.sort ()

    # Estrai il testo per ogni sezione
    for i, (pos, sezione) in enumerate (sezioni_trovate):
        inizio = pos + len (sezione)
        if i < len (sezioni_trovate) - 1:
            fine = sezioni_trovate[i + 1][0]
        else:
            fine = len (testo_completo)

        testo_sezioni[sezione] = testo_completo[inizio:fine].strip ()

    # Regex e funzioni di supporto
    pattern_cf = r"\b[A-Z]{6}[0-9]{2}[A-Z][0-9]{2}[A-Z][0-9]{3}[A-Z]\b"
    codici_trovati = set ()
    dati = []

    def verifica_cognome(nome, codice_fiscale):
        prime_3_lettere = codice_fiscale[:3]
        for parola in nome.split ():
            if all (lettera in parola for lettera in prime_3_lettere):
                return True
        return False

    def rimuovi_numeri(riga):
        return re.sub (r"\d+", "", riga).strip ()

    def is_valid_word(parola):
        return all (lettera in "ABCDEFGHIJKLMNOPQRSTUVWXYZÀÈÌÒÙàèìòù''\"-" for lettera in parola)

    def elabora_sezione(testo_sezione, tipo_sezione):
        righe = testo_sezione.splitlines ()

        for i, riga in enumerate (righe):
            match_cf = re.search (pattern_cf, riga)
            if match_cf:
                codice_fiscale = match_cf.group ()
                if codice_fiscale in codici_trovati:
                    continue
                codici_trovati.add (codice_fiscale)

                # Analisi progressiva delle righe, partendo dalla riga corrente
                nome_completo = []
                nome_trovato = False

                # Lista degli offset da provare in ordine: 0 (riga corrente), -1, -2, -3
                offsets_da_provare = [0, -1, -2, -3]

                for offset in offsets_da_provare:
                    index = i + offset
                    if 0 <= index < len (righe):
                        riga_corrente = rimuovi_numeri (righe[index].strip ())
                        if riga_corrente and not riga_corrente.split ()[0].isupper ():
                            continue

                        parole = riga_corrente.split ()
                        parole_valide = []
                        for parola in parole:
                            if is_valid_word (parola):
                                parole_valide.append (parola)
                            elif parola.islower ():
                                break

                        # Se troviamo parole valide in questa riga
                        if parole_valide:
                            nome_completo = parole_valide
                            nome_trovato = True
                            break  # Usciamo dal ciclo non appena troviamo un nome valido

                if nome_trovato:
                    cognome_candidato = nome_completo[0]
                    if not verifica_cognome (cognome_candidato, codice_fiscale):
                        cognome = " ".join (nome_completo[:2])
                        nomi = " ".join (nome_completo[2:]) if len (nome_completo) > 2 else ""
                    else:
                        cognome = cognome_candidato
                        nomi = " ".join (nome_completo[1:])

                    dati.append ({
                        "Cognome": cognome,
                        "Nomi": nomi,
                        "Codice Fiscale": codice_fiscale,
                        "Sezione": tipo_sezione
                    })

    # Elabora tutte le sezioni trovate
    for sezione, testo in testo_sezioni.items ():
        elabora_sezione (testo, sezione)

    return dati

# Interfaccia Streamlit
st.set_page_config(page_title="Estrazione Nominativi", page_icon="📜", layout="centered")

# Header con titolo personalizzato
st.markdown(
    """
    <h1 style="color:darkblue; text-align:center;">Estrazione Nominativi da Visura Camerale TELEMACO</h1>
    <h3 style="color:gray; text-align:center;">per verifiche presso il Casellario</h3>
    """,
    unsafe_allow_html=True,
)

# Sezione di caricamento file
st.write("**Carica un file PDF di una visura camerale Telemaco per estrarre i nominativi e scaricare i dati in formato Excel.**")

# Caricamento del file PDF
uploaded_file = st.file_uploader("Seleziona un file PDF", type=["pdf"])

if uploaded_file is not None:
    # Salva il file caricato
    with open("uploaded_file.pdf", "wb") as f:
        f.write(uploaded_file.read())

    # Estrai i dati
    st.info("Elaborazione in corso, attendere qualche secondo...")
    dati = estrai_dati("uploaded_file.pdf")

    # Mostra i dati estratti
    if dati:
        df = pd.DataFrame(dati)
        st.success("Dati estratti con successo! Visualizza o scarica il file Excel.")
        st.dataframe(df)

        # Consenti il download del file Excel
        output_path = "Elenco per casellario.xlsx"

        # Esporta il DataFrame in Excel usando pandas
        df.to_excel (output_path, index=False, engine='openpyxl')

        # Adatta le colonne al contenuto
        wb = openpyxl.load_workbook (output_path)
        ws = wb.active

        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter  # Ottieni la lettera della colonna
            for cell in col:
                try:  # Calcola la lunghezza massima del contenuto
                    if cell.value:
                        max_length = max (max_length, len (str (cell.value)))
                except:
                    pass
            adjusted_width = max_length + 2  # Aggiungi un margine
            ws.column_dimensions[column].width = adjusted_width

        # Salva il file con le colonne adattate
        wb.save (output_path)

        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Scarica il file Excel",
                data=f,
                file_name="Elenco per casellario.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ Nessun dato trovato nel file PDF.")

# Barra laterale (opzionale)
with st.sidebar:
    st.markdown("### Informazioni sull'app:")
    st.write("Questa applicazione permette di estrarre i nominativi e i codici fiscali dai file PDF delle visure camerali di Telemaco, consentendo successivamente di effettuare i controlli presso il Casellario Giudiziale.")
    st.write("Versione: 1.0")
    st.write("Sviluppata da Luca Bruzzi")