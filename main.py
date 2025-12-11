import streamlit as st
import pandas as pd
import openpyxl
from PyPDF2 import PdfReader
import re
from datetime import datetime

# Configurazione iniziale della pagina con tema personalizzato
st.set_page_config(
    page_title="Estrazione Nominativi",
    page_icon="📜",
    layout="centered"
)

# Custom CSS per migliorare l'aspetto
st.markdown("""
    <style>
    /* Stile per l'area di upload */
    [data-testid="stFileUploader"] {
        border: 2px dashed #1e3799 !important;
        border-radius: 10px !important;
        padding: 20px !important;
    }

    [data-testid="stFileUploader"]:hover {
        border-color: #4a69bd !important;
        background-color: #f8f9fa !important;
    }

    /* Stile per la card dei dati societari */
    .societary-data-card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 20px;
        margin: 20px 0;
        border-left: 5px solid #1e3799;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }

    .societary-data-card h3 {
        color: #1e3799;
        margin-bottom: 20px;
        padding-bottom: 10px;
        border-bottom: 1px solid #dee2e6;
    }

    /* Stile per i singoli campi dei dati societari */
    .data-field {
        background-color: white;
        padding: 10px 15px;
        border-radius: 5px;
        margin-bottom: 10px;
        border: 1px solid #e9ecef;
    }

    .data-field strong {
        color: #1e3799;
    }

    /* Separatore visivo */
    .section-divider {
        height: 2px;
        background-color: #e9ecef;
        margin: 30px 0;
    }
    </style>
""", unsafe_allow_html=True)


# Funzione per decodificare la data di nascita dal codice fiscale
def decodifica_data_nascita(codice_fiscale):
    """
    Estrae la data di nascita dal codice fiscale italiano
    Formato: RRSSAAMMGGCCCC
    Posizioni 7-12: AAMMGG (Anno, Mese, Giorno)
    """
    try:
        if len(codice_fiscale) != 16:
            return "N/A"

        # Estrai anno, mese, giorno
        anno_cf = codice_fiscale[6:8]
        mese_cf = codice_fiscale[8:9]
        giorno_cf = codice_fiscale[9:11]

        # Decodifica dell'anno (assumiamo che anni 00-30 siano 2000-2030, 31-99 siano 1931-1999)
        anno = int(anno_cf)
        if anno <= 30:
            anno += 2000
        else:
            anno += 1900

        # Decodifica del mese
        mesi = {
            'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5, 'H': 6,
            'L': 7, 'M': 8, 'P': 9, 'R': 10, 'S': 11, 'T': 12
        }

        if mese_cf not in mesi:
            return "N/A"

        mese = mesi[mese_cf]

        # Decodifica del giorno (per le donne si aggiunge 40)
        giorno = int(giorno_cf)
        if giorno > 31:
            giorno -= 40

        # Verifica validità della data
        try:
            data = datetime(anno, mese, giorno)
            return data.strftime("%d/%m/%Y")
        except ValueError:
            return "N/A"

    except (ValueError, IndexError, KeyError):
        return "N/A"


# Funzione per estrarre il codice catastale dal codice fiscale
def estrai_codice_catastale(codice_fiscale):
    """
    Estrae il codice catastale del comune di nascita dal codice fiscale
    Si trova negli ultimi 4 caratteri del codice fiscale
    """
    try:
        if len(codice_fiscale) != 16:
            return "N/A"

        codice_catastale = codice_fiscale[11:15]
        return codice_catastale

    except (ValueError, IndexError):
        return "N/A"


# Funzione per estrarre i dati
def estrai_dati(filepath):
    # Caricamento del PDF
    reader = PdfReader(filepath)
    text = ""
    for page in reader.pages:
        text += page.extract_text()

    righe = text.splitlines()

    # Ricerca della "Forma giuridica" con controllo sicurezza
    forma_giuridica = "NON TROVATO"
    for i, riga in enumerate(righe):
        if "Forma giuridica" in riga:
            # Trova tutte le parole successive alla "Forma giuridica"
            forma_giuridica_parole = []
            parti = riga.split()
            trovato_forma = False

            # Partiamo dalla parola successiva a "Forma giuridica"
            if len(parti) > 2:  # Controllo che ci siano parole dopo "Forma giuridica"
                for parola in parti[2:]:
                    if parola and len(parola) > 0 and parola[0].isupper():
                        trovato_forma = True
                        break
                    forma_giuridica_parole.append(parola)

            # Se non è completa, continua con la riga successiva (CON CONTROLLO)
            if not trovato_forma and i + 1 < len(righe):
                riga_successiva = righe[i + 1]
                parti_successiva = riga_successiva.split()
                for parola in parti_successiva:
                    if parola and len(parola) > 0 and parola[0].isupper():
                        trovato_forma = True
                        break
                    forma_giuridica_parole.append(parola)

            # Unisci le parole per ottenere la forma giuridica
            if forma_giuridica_parole:
                forma_giuridica = " ".join(forma_giuridica_parole).strip()
            break

    # Estrarre il numero degli addetti
    numero_addetti = "NON TROVATO"
    for i, riga in enumerate(righe):
        if "Addetti" in riga:
            # Cerca un numero che viene dopo una data (se presente) o dopo la parola Addetti
            match = re.search(r'Addetti.*?(?:\d{2}/\d{2}/\d{4})?\s*(\d+)\s*$', riga)
            if match:
                numero_addetti = match.group(1)
                break

    # Estrarre la ragione sociale
    ragione_sociale = "NON TROVATO"
    for i, riga in enumerate(righe):
        if "VISURA" in riga or "FASCICOLO" in riga:
            # Ragione sociale inizia due o tre righe dopo "VISURA" o "FASCICOLO"
            inizio = i + 2

            # Controllo che l'indice sia valido
            if inizio >= len(righe):
                break

            # Verifica se la riga iniziale è vuota
            if inizio < len(righe) and righe[inizio].strip() == "":
                inizio += 1

            # Controllo che l'indice aggiornato sia ancora valido
            if inizio >= len(righe):
                break

            # Concatenare righe fino a incontrare una riga vuota
            for j in range(inizio, len(righe)):
                if righe[j].strip() == "":  # Interrompe se la riga è vuota
                    break
                ragione_sociale += righe[j].strip() + " "

            ragione_sociale = ragione_sociale.strip()  # Rimuove spazi superflui
            break  # Interrompiamo la ricerca dopo aver trovato la prima occorrenza

    # Estrarre l'indirizzo (Comune e Via)
    comune = "NON TROVATO"
    via = "NON TROVATO"

    for i, riga in enumerate(righe):
        if "Indirizzo Sede" in riga:
            # Aggiungiamo uno spazio dopo "Sede" per separare "Sede" da "BOLOGNA" o altre parole
            riga = riga.replace("Sede", "Sede ")
            # Trova il Comune e la Via
            parti = riga.split()
            comune_parole = []
            via_parole = []
            trovato_comune = False

            # Analizza la prima riga per estrarre il Comune e la Via
            if len(parti) > 2:  # Controllo che ci siano parole dopo "Indirizzo Sede"
                for parola in parti[2:]:  # Ignora "Indirizzo Sede"
                    if not trovato_comune:
                        # Aggiungi al Comune solo parole che iniziano con una maiuscola
                        if parola and len(parola) > 0 and parola[0].isupper():
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

            # Se la riga successiva contiene il CAP, aggiungi la parte della via senza il CAP (CON CONTROLLO)
            if i + 1 < len(righe):  # Controlla se esiste una riga successiva
                riga_successiva = righe[i + 1]
                parti_successiva = riga_successiva.split()
                for parola in parti_successiva:
                    if "CAP" in parola:
                        break  # Interrompi se trovi "CAP" nella riga successiva
                    via_parole.append(parola)

            # Risultato
            if comune_parole:
                comune = " ".join(comune_parole).strip()
            if via_parole:
                via = " ".join(via_parole).strip()

            # Rimuovi tutte le parole dopo il CAP, incluso CAP stesso
            if "CAP" in via:
                via = via.split("CAP")[0].strip()

            break  # Esci dal ciclo dopo aver trovato la prima occorrenza

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

    if occorrenze_sezioni:
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
        e se la seconda parola è parte del cognome o del nome. Restituisce TRUE se il cognome
        è dato solo dalla prima parola e FALSE se il cognome è composto anche dalla seconda parola
        """
        prime_3_lettere = codice_fiscale[:3]
        successive_3_lettere = codice_fiscale[3:6]
        quarto_carattere = codice_fiscale[3:4]
        quinto_sesto_carattere = codice_fiscale[4:6]

        parole = nome.split()

        if len(parole) < 2:
            return False  # Caso con una sola parola

        prima_parola = parole[0]
        seconda_parola = parole[1]
        terza_parola = parole[2] if len(parole) > 2 else ""

        # Verifica che le prime 3 lettere siano nella prima parola
        if not all(lettera in prima_parola for lettera in prime_3_lettere):
            return False

        # Se successive_3_lettere sono nella terza parola, escludi la seconda parola dal cognome
        if all(lettera in terza_parola for lettera in successive_3_lettere):
            return False

        # Verifica se successive_3_lettere sono nella seconda parola, quinto_sesto_carattere nella terza, o quarto_carattere nella seconda
        if all(lettera in seconda_parola for lettera in successive_3_lettere) or \
                all(lettera in terza_parola for lettera in quinto_sesto_carattere) or \
                all(lettera in seconda_parola for lettera in quarto_carattere):
            return True  # Nome correttamente separato
        else:
            return False  # Se non trovato in seconda o terza parola

    def rimuovi_numeri(riga):
        return re.sub(r"\d+", "", riga).strip()

    def is_valid_word(parola):
        return all(lettera in "ABCDEFGHIJKLMNOPQRSTUVWXYZÀÈÌÒÙ àèìòù''\"- " for lettera in parola)

    def elabora_sezione(testo_sezione, tipo_sezione):
        righe = testo_sezione.splitlines()

        for i, riga in enumerate(righe):
            match_cf = re.search(pattern_cf, riga)
            if match_cf:
                codice_fiscale = match_cf.group()

                # Analisi progressiva delle righe, partendo dalla riga corrente
                nome_completo = []
                nome_trovato = False

                # Lista degli offset da provare in ordine: 0 (riga corrente), -1, -2, -3
                offsets_da_provare = [0, -1, -2, -3]
                parole_valide_totali = []  # Accumula le parole valide trovate finora
                ultima_parola = None  # Variabile per conservare l'ultima parola trovata

                for offset in offsets_da_provare:
                    index = i + offset
                    if 0 <= index < len(righe):  # CONTROLLO SICUREZZA
                        riga_corrente = rimuovi_numeri(righe[index].strip())
                        if not riga_corrente:
                            continue

                        # La prima parola deve iniziare con una maiuscola, altrimenti ignoriamo la riga
                        parole_riga_split = riga_corrente.split()
                        if not parole_riga_split or not parole_riga_split[0] or not parole_riga_split[0][0].isupper():
                            continue

                        # Estraiamo le parole valide dalla riga corrente
                        parole = riga_corrente.split()
                        parole_valide_riga = []
                        for parola in parole:
                            if is_valid_word(parola):
                                parole_valide_riga.append(parola)
                            elif parola.islower():
                                break  # Se troviamo una parola minuscola, ci fermiamo su questa riga

                        # Caso: una sola parola valida sulla riga
                        if len(parole_valide_riga) == 1:
                            # Memorizza temporaneamente come ultima parola
                            if not ultima_parola:
                                ultima_parola = parole_valide_riga[0]
                            continue  # Continua a cercare altre parole valide in righe precedenti

                        # Caso: più parole valide sulla riga
                        elif len(parole_valide_riga) > 1:
                            parole_valide_totali.extend(parole_valide_riga)

                        # Se abbiamo trovato almeno due parole valide, interrompiamo la ricerca
                        if len(parole_valide_totali) >= 2:
                            break

                # Aggiunge l'ultima parola solo dopo aver completato il nome
                if ultima_parola:
                    parole_valide_totali.append(ultima_parola)

                # Assegna il nome completo solo se abbiamo trovato almeno due parole
                if len(parole_valide_totali) >= 2:
                    nome_completo = parole_valide_totali
                    nome_trovato = True

                if nome_trovato:
                    cognome_candidato = " ".join(nome_completo)
                    if not verifica_cognome(cognome_candidato, codice_fiscale):
                        cognome = " ".join(nome_completo[:2])  # Cognome = prime due parole
                        nomi = " ".join(nome_completo[2:]) if len(nome_completo) > 2 else ""
                    else:
                        cognome = nome_completo[0]  # Cognome = prima parola
                        nomi = " ".join(nome_completo[1:])

                    # Estrai data di nascita e codice catastale dal codice fiscale
                    data_nascita = decodifica_data_nascita(codice_fiscale)
                    codice_catastale = estrai_codice_catastale(codice_fiscale)

                    # Se il codice fiscale è già presente
                    if codice_fiscale in codici_trovati:
                        # Aggiorna il record esistente aggiungendo la nuova sezione
                        for record in dati:
                            if record["Codice Fiscale"] == codice_fiscale:
                                if tipo_sezione not in record["Sezione"].split(", "):
                                    record["Sezione"] = f"{record['Sezione']}, {tipo_sezione}"
                    else:
                        # Crea un nuovo record
                        codici_trovati[codice_fiscale] = True
                        dati.append({
                            "Cognome": cognome,
                            "Nomi": nomi,
                            "Codice Fiscale": codice_fiscale,
                            "Data di nascita": data_nascita,
                            "Codice catastale": codice_catastale,
                            "Sezione": tipo_sezione
                        })

    # Elabora tutte le sezioni trovate
    for sezione, testo in testo_sezioni.items():
        elabora_sezione(testo, sezione)

    return dati, ragione_sociale, comune, via, numero_addetti, forma_giuridica


# Interfaccia Streamlit
# Contenitore principale con stile migliorato
st.markdown(
    """
    <div style="text-align: center; padding: 2rem 0;">
        <h1 style="color: #1e3799; margin-bottom: 0.5rem;">
            Analisi di Visure Camerali TELEMACO
        </h1>
        <h3 style="color: #576574; font-weight: normal;">
            (per Controlli ai sensi del D.Lgs. 36/2023)
        </h3>
    </div>
    """,
    unsafe_allow_html=True
)

# Area di upload con testo personalizzato
uploaded_file = st.file_uploader(
    label="Carica un file PDF di una visura camerale Telemaco",
    type=["pdf"],
    key="pdf_uploader",
    help="Trascina o carica un file PDF da elaborare.",
    label_visibility="collapsed"
)

if uploaded_file is not None:
    # Salva il file caricato
    with open("uploaded_file.pdf", "wb") as f:
        f.write(uploaded_file.read())

    # Mostra un loader durante l'elaborazione
    with st.spinner('Elaborazione in corso...'):
        dati, ragione_sociale, comune, via, numero_addetti, forma_giuridica = estrai_dati("uploaded_file.pdf")

    # Mostra i dati estratti
    if dati:
        df = pd.DataFrame(dati)
        st.success("✅ Dati estratti con successo!")

        # Card per i dati societari con nuovo stile
        st.markdown(f"""
            <div class="societary-data-card">
                <h3>📊 Dati Societari</h3>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
                        <div>
                            <div class="data-field" style="word-wrap: break-word; white-space: normal;">
                            <strong>🏢 Ragione Sociale</strong><br>
                            {ragione_sociale}
                        </div>
                        <div class="data-field">
                            <strong>⚖️ Forma giuridica</strong><br>
                            {forma_giuridica.upper()}
                        </div>
                        <div class="data-field">
                            <strong>📍 Sede legale</strong><br>
                            {comune}
                        </div>
                    </div>
                    <div>
                        <div class="data-field">
                            <strong>🏠 Indirizzo</strong><br>
                            {via}
                        </div>
                        <div class="data-field">
                            <strong>👥 Numero Addetti</strong><br>
                            {"&lt;" + numero_addetti + "&gt;" if numero_addetti == "NON TROVATO" else numero_addetti}
                        </div>
                    </div>
                </div>
            </div>
            <div class="section-divider"></div>
        """, unsafe_allow_html=True)
        
        # Visualizzazione della tabella con stile
        st.markdown("### 📋 Elenco Nominativi")
        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True
        )
        
        # Preparazione e download del file Excel
        output_path = "Elenco per casellario.xlsx"
        df.to_excel(output_path, index=False, engine='openpyxl')

        # Formattazione Excel
        wb = openpyxl.load_workbook(output_path)
        ws = wb.active
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = max_length + 2
            ws.column_dimensions[column].width = adjusted_width
        wb.save(output_path)

        # Pulsante di download stilizzato
        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Scarica il file Excel",
                data=f,
                file_name="Elenco per casellario.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("❌ Nessun dato trovato nel file PDF.")

with st.sidebar:
    st.markdown("""
        <div style="background: #f8f9fa; padding: 1.5rem; border-radius: 6px;">
            <h3 style="color: #1e3799; font-size: 22px;">ℹ️ Informazioni sull'app</h3>
        </div>
    """, unsafe_allow_html=True)
    st.divider()

    st.markdown("""
        <div style="font-size: 18px;">
            Carica il PDF di una visura camerale di Telemaco e ottieni:<br><br>
        </div>
        <div style="font-size: 18px;">
            • <strong>Dati societari principali</strong> (ragione sociale, sede, forma giuridica, numero addetti).<br>
        </div>
        <div style="font-size: 18px;">
            • <strong>Elenco delle cariche aziendali</strong> (nome, cognome, codice fiscale, data di nascita, codice catastale).<br><br>
        </div>
        <div style="font-size: 18px;">
            Puoi esportare i risultati in formato Excel per effettuare i controlli previsti dal <strong>D.Lgs. 36/2023</strong>.<br><br>
        </div>
        <div style="font-size: 18px;">
            Una soluzione semplice e veloce per chi deve gestire verifiche aziendali.
        </div>
    """, unsafe_allow_html=True)

    st.divider()

    st.markdown("""
        <div style="font-size: 20px;">
            <strong>📄 Versione:</strong> 1.4 (Beta)
        </div>
    """, unsafe_allow_html=True)
    st.markdown("""
        <div style="font-size: 20px;">
            <strong>👨‍💻 Sviluppato da:</strong> Luca Bruzzi
        </div>
    """, unsafe_allow_html=True)