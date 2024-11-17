import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader
import re

# Funzione per estrarre i dati
def estrai_dati(filepath):
    # Caricamento del PDF
    reader = PdfReader(filepath)
    text = ""
    for page in reader.pages:
        text += page.extract_text()

    righe = text.splitlines ()
    # Trova la seconda occorrenza della sezione target
    target_sezione = "Trasferimenti d'azienda, fusioni, scissioni, subentri"
    righe_target = [i for i, riga in enumerate (righe) if target_sezione in riga]

    if len (righe_target) < 2:
        print (f"Non è stata trovata una seconda occorrenza della sezione '{target_sezione}'.")
        return []

    riga_fine = righe_target[1]  # Ottieni l'indice della seconda occorrenza
    # Limita le righe fino alla seconda occorrenza
    righe = righe[:riga_fine]

    # Debug: Stampa le prime 500 righe del testo
    #print("\n".join(text.splitlines()[:500]))  # Stampa le prime 500 righe del testo

    # Prosegui con l'estrazione dei dati
    sezioni = "\n".join (righe).split ("Soci e titolari di diritti su azioni e quote", 2)
    if len (sezioni) < 2:
        print ("La sezione 'Soci e titolari di diritti su azioni e quote' non è stata trovata.")
        return []

    fine_sezione = sezioni[-1]
    sezioni_amministratori = fine_sezione.split("Amministratori", 1)
    soci_e_titolari_text = sezioni_amministratori[0]
    amministratori_text = sezioni_amministratori[1] if len(sezioni_amministratori) > 1 else ""

    # Regex per identificare i codici fiscali
    pattern_cf = r"\b[A-Z]{6}[0-9]{2}[A-Z][0-9]{2}[A-Z][0-9]{3}[A-Z]\b"

    dati = []
    codici_trovati = set()  # Set per tenere traccia dei codici fiscali già trovati

    def verifica_cognome(nome, codice_fiscale):
        prime_3_lettere = codice_fiscale[:3]
        for parola in nome.split ():
            if all (lettera in parola for lettera in prime_3_lettere):
                return True
        return False

    # Funzione di utilità: rimuovere numeri dalla riga per analisi nomi e cognomi
    def rimuovi_numeri(riga):
        return re.sub(r"\d+", "", riga).strip()

    def is_valid_word(parola):
        """Verifica se una parola è composta solo da maiuscole, accenti e trattini."""
        for lettera in parola:
            if lettera not in "ABCDEFGHIJKLMNOPQRSTUVWXYZÀÈÌÒÙàèìòù’'\"-":
                return False
        return True

    # Elaborazione sezione "Soci e titolari di diritti su azioni e quote"
    righe_soci = soci_e_titolari_text.splitlines()
    for riga in righe_soci:
        match_cf = re.search(pattern_cf, riga)  # Cerca il codice fiscale nella riga originale
        if match_cf:
            codice_fiscale = match_cf.group()
            if codice_fiscale in codici_trovati:
                continue  # Salta se il codice fiscale è già stato trovato
            codici_trovati.add(codice_fiscale)

            # Analizza nomi e cognomi rimuovendo i numeri
            riga_pulita = rimuovi_numeri(riga)
            parole = riga_pulita.split()
            nome_completo = []

            for parola in parole:
                if is_valid_word (parola):
                    nome_completo.append (parola)
                elif parola.islower ():
                    break

            if nome_completo:
                cognome_candidato = nome_completo[0]
                # Verifica se il cognome è composto
                if not verifica_cognome (cognome_candidato, codice_fiscale):
                    cognome = " ".join (nome_completo[:2])  # Considera le prime due parole come cognome
                    nomi = " ".join (nome_completo[2:]) if len (nome_completo) > 2 else ""
                else:
                    cognome = cognome_candidato
                    nomi = " ".join (nome_completo[1:])
                dati.append ({
                    "Cognome": cognome,
                    "Nomi": nomi,
                    "Codice Fiscale": codice_fiscale
                })

    # Elaborazione sezione "Amministratori"
    righe_amministratori = amministratori_text.splitlines()
    for i, riga in enumerate(righe_amministratori):
        match_cf = re.search(pattern_cf, riga)  # Cerca il codice fiscale nella riga originale
        if match_cf:
            codice_fiscale = match_cf.group()
            if codice_fiscale in codici_trovati:
                continue  # Salta se il codice fiscale è già stato trovato
            codici_trovati.add(codice_fiscale)

            # Analizza le righe -2, -1 e quella corrente
            nome_completo = []
            for offset in range(-2, 1):  # Da riga -2 alla riga corrente inclusa
                index = i + offset
                if 0 <= index < len(righe_amministratori):
                    riga_corrente = rimuovi_numeri(righe_amministratori[index].strip())
                    # Ignora la riga se inizia con una parola non tutta maiuscola
                    if riga_corrente and not riga_corrente.split()[0].isupper():
                        continue

                    parole = riga_corrente.split()
                    for parola in parole:
                        if is_valid_word (parola):
                            nome_completo.append (parola)
                        elif parola.islower ():
                            break

            if nome_completo:
                cognome_candidato = nome_completo[0]
                # Verifica se il cognome è composto
                if not verifica_cognome (cognome_candidato, codice_fiscale):
                    cognome = " ".join (nome_completo[:2])  # Considera le prime due parole come cognome
                    nomi = " ".join (nome_completo[2:]) if len (nome_completo) > 2 else ""
                else:
                    cognome = cognome_candidato
                    nomi = " ".join (nome_completo[1:])
                dati.append ({
                    "Cognome": cognome,
                    "Nomi": nomi,
                    "Codice Fiscale": codice_fiscale
                })

    return dati

# Interfaccia Streamlit
st.title("Estrazione Dati da PDF")
st.write("Carica un file PDF per estrarre i dati e scaricare un file Excel.")

uploaded_file = st.file_uploader("Carica un file PDF", type=["pdf"])

if uploaded_file is not None:
    # Salva il file caricato
    with open("uploaded_file.pdf", "wb") as f:
        f.write(uploaded_file.read())

    # Estrai i dati
    st.write("Elaborazione in corso...")
    dati = estrai_dati("uploaded_file.pdf")

    # Mostra i dati estratti
    if dati:
        df = pd.DataFrame(dati)
        st.write("Dati estratti:")
        st.dataframe(df)

        # Consenti il download del file Excel
        output_path = "dati_estratti.xlsx"
        df.to_excel(output_path, index=False)
        with open(output_path, "rb") as f:
            st.download_button(
                label="Scarica il file Excel",
                data=f,
                file_name="dati_estratti.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.write("Nessun dato trovato nel file PDF.")