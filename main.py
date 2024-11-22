import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from PyPDF2 import PdfReader
import re

def estrai_dati(filepath):
    # Caricamento del PDF
    reader = PdfReader(filepath)
    text = ""
    for page in reader.pages:
        text += page.extract_text()

    righe = text.splitlines()

    # Estrarre la ragione sociale
    ragione_sociale = ""
    for i, riga in enumerate(righe):
        if "VISURA" in riga:
            # Ragione sociale inizia due righe dopo "VISURA"
            inizio = i + 2
            fine = inizio + 3  # Include 3 righe dopo
            ragione_sociale = " ".join(righe[inizio:fine]).strip()
            break  # Interrompiamo la ricerca dopo aver trovato la prima occorrenza

    # Se non troviamo "VISURA", avvisiamo
    if not ragione_sociale:
        print("Non è stato possibile trovare la ragione sociale.")
    else:
        print(f"Ragione Sociale estratta: {ragione_sociale}")

    # Estrarre l'indirizzo (Comune e Via)
    comune = ""
    via = ""

    for i, riga in enumerate(righe):
        if "Indirizzo Sede" in riga:
            # Trova il Comune e la Via
            parti = riga.split()
            comune_parole = []
            via_parole = []
            trovato_comune = False

            # Analizza la prima riga per estrarre il Comune e la Via
            for parola in parti[2:]:  # Ignora "Indirizzo Sede"
                print(f"Parola analizzata: {parola}")  # Debug
                if not trovato_comune:
                    # Aggiungi al Comune solo parole che iniziano con una maiuscola
                    if parola[0].isupper():
                        comune_parole.append(parola)
                    # Se trovi una parentesi chiusa, il Comune è completo
                    if ")" in parola:
                        comune_parole.append(parola)  # Aggiungi la sigla del Comune
                        trovato_comune = True
                elif trovato_comune:
                    # Aggiungi la parola alla via, ma non includere numeri o CAP
                    if "CAP" in parola:
                        break  # Interrompi l'analisi delle parole dopo "CAP"
                    via_parole.append(parola)

            # Se la riga successiva contiene il CAP, aggiungi la parte della via senza il CAP
            if i + 1 < len(righe):  # Controlla se esiste una riga successiva
                riga_successiva = righe[i + 1]
                parti_successiva = riga_successiva.split()
                for parola in parti_successiva:
                    if "CAP" in parola:
                        break  # Interrompi se trovi "CAP" nella riga successiva
                    via_parole.append(parola)

            # Risultato
            comune = " ".join(comune_parole).strip()
            via = " ".join(via_parole).strip()

            # Rimuovi tutte le parole dopo il CAP, incluso CAP stesso
            if "CAP" in via:
                via = via.split("CAP")[0].strip()

            break  # Esci dal ciclo dopo aver trovato la prima occorrenza

    # Se non troviamo "Indirizzo Sede", avvisiamo
    if not comune or not via:
        print("Non è stato possibile trovare il Comune o la Via.")
    else:
        print(f"Comune estratto: {comune}")
        print(f"Via estratta: {via}")

    # Lista delle sezioni che determinano la fine della ricerca
    sezioni_fine = [
        "Trasferimenti d'azienda, fusioni, scissioni, subentri",
        "Trasferimenti d'azienda, subentri",
        "Attivita', albi ruoli e licenze",
        "Storia delle modifiche"
    ]

    # Trova la seconda occorrenza di una qualsiasi delle sezioni di fine
    occorrenze_sezioni = {}
    for sezione in sezioni_fine:
        occorrenze = [i for i, riga in enumerate(righe) if sezione in riga]
        if len(occorrenze) >= 2:
            occorrenze_sezioni[sezione] = occorrenze[1]  # Prendi la seconda occorrenza

    if not occorrenze_sezioni:
        print("Non è stata trovata una seconda occorrenza di nessuna delle sezioni di fine.")
        return [], ragione_sociale, comune, via  # Restituisci valori vuoti per i dati

    # Prendi la prima seconda occorrenza tra tutte le sezioni trovate
    riga_fine = min(occorrenze_sezioni.values())

    # Limita le righe fino alla seconda occorrenza della prima sezione di fine trovata
    righe = righe[:riga_fine]
    testo_completo = "\n".join(righe)

    # Lista delle possibili sezioni da cercare
    sezioni_da_cercare = [
        "Soci e titolari di diritti su azioni e quote",
        "Soci e titolari di cariche o qualifiche",
        "Amministratori",
        "Sindaci, membri organi di controllo",
        "Titolari di altre cariche o qualifiche",
        "Titolari di cariche o qualifiche"
    ]

    # Trova tutte le sezioni presenti nel testo
    sezioni_trovate = []
    testo_sezioni = {}

    for i, sezione in enumerate(sezioni_da_cercare):
        indici = [m.start() for m in re.finditer(re.escape(sezione), testo_completo)]
        for indice in indici:
            sezioni_trovate.append((indice, sezione))

    # Ordina le sezioni per posizione nel testo
    sezioni_trovate.sort()

    # Estrai il testo per ogni sezione
    for i, (pos, sezione) in enumerate(sezioni_trovate):
        inizio = pos + len(sezione)
        if i < len(sezioni_trovate) - 1:
            fine = sezioni_trovate[i + 1][0]
        else:
            fine = len(testo_completo)

        testo_sezioni[sezione] = testo_completo[inizio:fine].strip()

    # Regex e funzioni di supporto
    pattern_cf = r"\b[A-Z]{6}[0-9]{2}[A-Z][0-9]{2}[A-Z][0-9]{3}[A-Z]\b"
    codici_trovati = {}  # Dizionario invece di set
    dati = []

    def verifica_cognome(nome, codice_fiscale):
        """
        Verifica se le prime 3 lettere del codice fiscale sono presenti nel cognome
        e se la seconda parola è parte del cognome o del nome.
        """
        prime_3_lettere = codice_fiscale[:3]
        successive_3_lettere = codice_fiscale[3:6]
        quinto_sesto_carattere = codice_fiscale[4:6]

        parole = nome.split()

        if len(parole) >= 2:
            prima_parola = parole[0]
            seconda_parola = parole[1]

            # Verifica se le prime 3 lettere del codice fiscale sono nella prima parola
            if all(lettera in prima_parola for lettera in prime_3_lettere):
                # Verifica se le successive 3 lettere sono nella seconda parola
                if len(parole) == 2:
                    if all(lettera in seconda_parola for lettera in successive_3_lettere):
                        return True  # Il nome è già separato correttamente
                    else:
                        return False  # Non trovate nella seconda parola, quindi non è separato correttamente
                elif len(parole) >= 3:  # Caso con più di due parole
                    terza_parola = parole[2]
                    if all(lettera in seconda_parola for lettera in successive_3_lettere) or \
                            all(lettera in terza_parola for lettera in quinto_sesto_carattere):
                        return True  # Nome correttamente separato
                    else:
                        return False  # Se non trovato in seconda o terza parola
        else:
            # Caso base: una sola parola o non corrisponde
            return all(lettera in parole[0] for lettera in prime_3_lettere)

    def rimuovi_numeri(riga):
        return re.sub(r"\d+", "", riga).strip()

    def is_valid_word(parola):
        return all(lettera in "ABCDEFGHIJKLMNOPQRSTUVWXYZÀÈÌÒÙàèìòù''\"-" for lettera in parola)

    def elabora_sezione(testo_sezione, tipo_sezione):
        righe = testo_sezione.splitlines()

        for i, riga in enumerate(righe):
            match_cf = re.search(pattern_cf, riga)
            if match_cf:
                codice_fiscale = match_cf.group()

                # Analisi progressiva delle righe, partendo dalla riga successiva a quella del CF
                for j in range(i + 1, len(righe)):
                    riga_info = rimuovi_numeri(righe[j])
                    if is_valid_word(riga_info):
                        nome = riga_info.split(" ")[0]
                        cognome = riga_info.split(" ")[-1]

                        if verifica_cognome(nome, codice_fiscale):
                            dati.append({
                                "Cognome": cognome,
                                "Nome": nome,
                                "Codice Fiscale": codice_fiscale,
                                "Tipo Sezione": tipo_sezione
                            })
                            break  # Esci dalla ricerca quando hai trovato una corrispondenza

    # Estrai i dati per ciascuna sezione
    for sezione, testo_sezione in testo_sezioni.items():
        elabora_sezione(testo_sezione, sezione)

    return dati, ragione_sociale, comune, via

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

    # Mostra i dati estratti (Ragione Sociale, Comune, Via)
    if dati:
        # Assicurati che ci siano dati
        ragione_sociale = ragione_sociale if ragione_sociale else 'Non disponibile'
        comune = comune if comune else 'Non disponibile'
        via = via if via else 'Non disponibile'

        # Visualizza i dati estratti sulla pagina
        st.success ("Dati estratti con successo!")

        st.subheader ("Dettagli Sede Legale")
        st.write (f"**Ragione Sociale:** {ragione_sociale}")
        st.write (f"**Comune:** {comune}")
        st.write (f"**Via:** {via}")

        # Mostra i dati in un dataframe
        df = pd.DataFrame (dati)
        st.dataframe (df)

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

        with open (output_path, "rb") as f:
            st.download_button (
                label="📥 Scarica il file Excel",
                data=f,
                file_name="Elenco per casellario.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning ("⚠️ Nessun dato trovato nel file PDF.")

# Barra laterale (opzionale)
with st.sidebar:
    st.markdown ("### Informazioni sull'app:")
    st.write (
        "Questa applicazione permette di estrarre i nominativi e i codici fiscali dai file PDF delle visure camerali di Telemaco, consentendo successivamente di effettuare i controlli presso il Casellario Giudiziale.")
    st.write ("Versione: 1.0")
    st.write ("Sviluppata da Luca Bruzzi")