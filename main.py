import streamlit as st
import pandas as pd
import openpyxl
from PyPDF2 import PdfReader
import re

# Configurazione iniziale della pagina con tema personalizzato
st.set_page_config(
    page_title="Estrazione Nominativi",
    page_icon="📜",
    layout="centered"
)

# Custom CSS per migliorare l'aspetto

st.markdown("""
    <style>
    
    .stApp {
        background: linear-gradient(to bottom right, #f5f7fa, #e3e6e8);
    }
    .main {
        padding: 2rem;
        border-radius: 10px;
        background: rgba(255, 255, 255, 0.95);
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    h1 {
        color: #1e3799;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
        font-family: 'Helvetica Neue', sans-serif;
    }
    .stButton>button {
        background-color: #1e3799;
        color: white;
        border-radius: 5px;
        padding: 0.5rem 1rem;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #0c2461;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }
     /* Personalizzazione area upload */
    .uploadfile {
        border: 2px dashed #1e3799 !important;
        border-radius: 10px !important;
        padding: 2rem !important;
        background-color: #f8f9fa !important;
    }
    .uploadfile:hover {
        background-color: #e9ecef !important;
        border-color: #0c2461 !important;
    }
    /* Nascondi il testo predefinito in inglese */
    .uploadfile > div > div > p {
        display: none !important;
    }
    /* Personalizza il pulsante Browse files */
    .uploadfile > div > div > button {
        background-color: #1e3799 !important;
        color: white !important;
    }
    .css-1vq4p4l {
        padding: 1rem !important;
    }
    </style>
    """, unsafe_allow_html=True)

# Funzione per estrarre i dati
def estrai_dati(filepath):
    # Caricamento del PDF
    reader = PdfReader (filepath)
    text = ""
    for page in reader.pages:
        text += page.extract_text ()

    righe = text.splitlines ()

    # Estrarre il numero degli addetti
    numero_addetti = "Non trovato"
    for i, riga in enumerate (righe):
        if "Addetti" in riga:
            # Cerca un numero che viene dopo una data (se presente) o dopo la parola Addetti
            match = re.search (r'Addetti.*?(?:\d{2}/\d{2}/\d{4})?\s*(\d+)\s*$', riga)
            if match:
                numero_addetti = match.group (1)
                print (f"Riga trovata: {riga}")  # Debug: stampa la riga per verifica
                break

    # Estrarre la ragione sociale
    ragione_sociale = ""
    for i, riga in enumerate (righe):
        if "VISURA" in riga:
            # Ragione sociale inizia due righe dopo "VISURA"
            inizio = i + 2
            fine = inizio + 3  # Include 3 righe dopo
            ragione_sociale = " ".join (righe[inizio:fine]).strip ()
            break  # Interrompiamo la ricerca dopo aver trovato la prima occorrenza

    # Se non troviamo "VISURA", avvisiamo
    if not ragione_sociale:
        print ("Non è stato possibile trovare la ragione sociale.")
    else:
        print (f"Ragione Sociale estratta: {ragione_sociale}")

    # Estrarre l'indirizzo (Comune e Via)
    comune = ""
    via = ""

    for i, riga in enumerate (righe):
        if "Indirizzo Sede" in riga:
            # Trova il Comune e la Via
            parti = riga.split ()
            comune_parole = []
            via_parole = []
            trovato_comune = False

            # Analizza la prima riga per estrarre il Comune e la Via
            for parola in parti[2:]:  # Ignora "Indirizzo Sede"
                if not trovato_comune:
                    # Aggiungi al Comune solo parole che iniziano con una maiuscola
                    if parola[0].isupper ():
                        comune_parole.append (parola)
                    # Se trovi una parentesi chiusa, il Comune è completo
                    if ")" in parola:
                        comune_parole.append (parola)  # Aggiungi la sigla del Comune
                        trovato_comune = True
                elif trovato_comune:
                    # Aggiungi la parola alla via, ma non includere numeri o CAP
                    if "CAP" in parola:
                        break  # Interrompi l'analisi delle parole dopo "CAP"
                    via_parole.append (parola)

            # Se la riga successiva contiene il CAP, aggiungi la parte della via senza il CAP
            if i + 1 < len (righe):  # Controlla se esiste una riga successiva
                riga_successiva = righe[i + 1]
                parti_successiva = riga_successiva.split ()
                for parola in parti_successiva:
                    if "CAP" in parola:
                        break  # Interrompi se trovi "CAP" nella riga successiva
                    via_parole.append (parola)

            # Risultato
            comune = " ".join (comune_parole).strip ()
            via = " ".join (via_parole).strip ()

            # Rimuovi tutte le parole dopo il CAP, incluso CAP stesso
            if "CAP" in via:
                via = via.split ("CAP")[0].strip ()

            break  # Esci dal ciclo dopo aver trovato la prima occorrenza

    # Se non troviamo "Indirizzo Sede", avvisiamo
    if not comune or not via:
        print ("Non è stato possibile trovare il Comune o la Via.")
    else:
        print (f"Comune estratto: {comune}")
        print (f"Via estratta: {via}")
    # Debug: Stampa le prime 500 righe del testo
    print("\n".join(text.splitlines()[:500]))  # Stampa le prime 500 righe del testo

    # Lista delle sezioni che determinano la fine della ricerca
    sezioni_fine = [
        "Trasferimenti d'azienda, fusioni, scissioni, subentri",
        "Trasferimenti d'azienda, subentri"
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
        "Soci e titolari di cariche o qualifiche",
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
    codici_trovati = {} # Dizionario invece di set
    dati = []

    def verifica_cognome(nome, codice_fiscale):
        """
        Verifica se le prime 3 lettere del codice fiscale sono presenti nel cognome
        e se la seconda parola è parte del cognome o del nome. Restituisce TRUE se il cognome
        è dato solo dalla prima parola e FALSE se il cognome è composto anche dalla seconda parola
        """
        prime_3_lettere = codice_fiscale[:3]
        successive_3_lettere = codice_fiscale[3:6]
        quinto_sesto_carattere = codice_fiscale[4:6]

        parole = nome.split ()

        if len (parole) >= 2:
            prima_parola = parole[0]
            seconda_parola = parole[1]

            # Verifica se le prime 3 lettere del codice fiscale sono nella prima parola
            if all (lettera in prima_parola for lettera in prime_3_lettere):
                # Verifica se le successive 3 lettere sono nella seconda parola
                if len (parole) == 2:
                    if all (lettera in seconda_parola for lettera in successive_3_lettere):
                        return True  # Il nome è già separato correttamente
                    else:
                        return False  # Non trovate nella seconda parola, quindi non è separato correttamente
                elif len (parole) >= 3:  # Caso con più di due parole
                    terza_parola = parole[2]
                    if all (lettera in seconda_parola for lettera in successive_3_lettere) or \
                            all (lettera in terza_parola for lettera in quinto_sesto_carattere):
                        return True  # Nome correttamente separato
                    else:
                        return False  # Se non trovato in seconda o terza parola
        else:
            # Caso base: una sola parola o non corrisponde
            return all (lettera in parole[0] for lettera in prime_3_lettere)

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

                    cognome_candidato = " ".join (nome_completo)
                    if not verifica_cognome (cognome_candidato, codice_fiscale):
                        cognome = " ".join (nome_completo[:2])  # Cognome = prime due parole
                        nomi = " ".join (nome_completo[2:]) if len (nome_completo) > 2 else ""
                    else:
                        cognome = nome_completo[0]  # Cognome = prima parola
                        nomi = " ".join (nome_completo[1:])

                    # Se il codice fiscale è già presente
                    if codice_fiscale in codici_trovati:
                        # Aggiorna il record esistente aggiungendo la nuova sezione
                        for record in dati:
                            if record["Codice Fiscale"] == codice_fiscale:
                                if tipo_sezione not in record["Sezione"].split (", "):
                                    record["Sezione"] = f"{record['Sezione']}, {tipo_sezione}"
                    else:
                        # Crea un nuovo record
                        codici_trovati[codice_fiscale] = True
                        dati.append ({
                            "Cognome": cognome,
                            "Nomi": nomi,
                            "Codice Fiscale": codice_fiscale,
                            "Sezione": tipo_sezione
                        })

    # Elabora tutte le sezioni trovate
    for sezione, testo in testo_sezioni.items ():
        elabora_sezione (testo, sezione)

    return dati, ragione_sociale, comune, via, numero_addetti

# Interfaccia Streamlit
#st.set_page_config(page_title="Estrazione Nominativi", page_icon="📜", layout="centered")

# Contenitore principale con stile migliorato
st.markdown(
    """
    <div style="text-align: center; padding: 2rem 0;">
        <h1 style="color: #1e3799; margin-bottom: 0.5rem;">
            Estrazione Nominativi da Visura Camerale TELEMACO
        </h1>
        <h3 style="color: #576574; font-weight: normal;">
            (per verifiche presso il Casellario)
        </h3>
    </div>
    """,
    unsafe_allow_html=True
)

# Area di upload con testo personalizzato
uploaded_file = st.file_uploader(
    label="Carica un file PDF di una visura camerale Telemaco",
    type=["pdf"],
    key="pdf_uploader"
)

if uploaded_file is not None:
        # Salva il file caricato
        with open ("uploaded_file.pdf", "wb") as f:
            f.write (uploaded_file.read ())

        # Mostra un loader durante l'elaborazione
        with st.spinner ('Elaborazione in corso...'):
            dati, ragione_sociale, comune, via, numero_addetti = estrai_dati ("uploaded_file.pdf")

        # Mostra i dati estratti
        if dati:
            df = pd.DataFrame (dati)
            st.success ("✅ Dati estratti con successo!")

            # Card per i dati societari
            st.markdown ("""
                <div style="background: #f8f9fa; padding: 1.5rem; border-radius: 8px; margin: 1rem 0;">
                    <h3 style="color: #1e3799; margin-bottom: 1rem;">📊 Dati Societari</h3>
                </div>
                """, unsafe_allow_html=True)

            col1, col2 = st.columns (2)
            with col1:
                st.markdown (f"**🏢 Ragione Sociale:**\n{ragione_sociale}")
                st.markdown (f"**📍 Sede legale:**\n{comune}")
            with col2:
                st.markdown (f"**🏠 Indirizzo:**\n{via}")
                st.markdown (f"**👥 Numero Addetti:**\n{numero_addetti}")

            # Visualizzazione della tabella con stile
            st.markdown ("### 📋 Elenco Nominativi")
            st.dataframe (
                df,
                use_container_width=True,
                hide_index=True
            )

            # Preparazione e download del file Excel
            output_path = "Elenco per casellario.xlsx"
            df.to_excel (output_path, index=False, engine='openpyxl')

            # Formattazione Excel
            wb = openpyxl.load_workbook (output_path)
            ws = wb.active
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if cell.value:
                            max_length = max (max_length, len (str (cell.value)))
                    except:
                        pass
                adjusted_width = max_length + 2
                ws.column_dimensions[column].width = adjusted_width
            wb.save (output_path)

            # Pulsante di download stilizzato
            with open (output_path, "rb") as f:
                st.download_button (
                    label="📥 Scarica il file Excel",
                    data=f,
                    file_name="Elenco per casellario.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error ("❌ Nessun dato trovato nel file PDF.")

# Barra laterale migliorata
with st.sidebar:
    st.markdown ("""
        <div style="background: #f8f9fa; padding: 1.5rem; border-radius: 8px;">
            <h3 style="color: #1e3799;">ℹ️ Informazioni sull'app</h3>
        </div>
        """, unsafe_allow_html=True)
    st.write ("""
        Questa applicazione permette di estrarre i nominativi e i codici fiscali dai file PDF 
        delle visure camerali di Telemaco, consentendo successivamente di effettuare i controlli 
        presso il Casellario Giudiziale.
    """)
    st.divider ()
    st.markdown ("**🔄 Versione:** 1.1")
    st.markdown ("**👨‍💻 Sviluppato da:** Luca Bruzzi")