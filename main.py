import streamlit as st
import pandas as pd
import openpyxl
from PyPDF2 import PdfReader
import re
from datetime import datetime

# Configurazione pagina
st.set_page_config(page_title="Estrazione Visure", page_icon="📂", layout="centered")

# CSS Custom
st.markdown("""
    <style>
    [data-testid="stFileUploader"] {border: 2px dashed #1e3799; padding: 20px;}
    .data-card {background-color: #f8f9fa; padding: 15px; border-radius: 10px; border-left: 5px solid #1e3799; margin-bottom: 20px;}
    .data-card h3 {color: #1e3799;}
    </style>
""", unsafe_allow_html=True)

# --- FUNZIONI DI SUPPORTO ---
def decodifica_data_nascita(cf):
    """Estrae data di nascita dal CF in modo robusto."""
    if len(cf) != 16: return "N/A"
    try:
        # Decodifica Anno
        anno_cf = int(cf[6:8])
        anno = anno_cf + (2000 if anno_cf <= 30 else 1900)
        
        # Decodifica Mese
        mesi = {'A':1,'B':2,'C':3,'D':4,'E':5,'H':6,'L':7,'M':8,'P':9,'R':10,'S':11,'T':12}
        mese = mesi.get(cf[8], 0)
        if mese == 0: return "N/A" # Mese non valido

        # Decodifica Giorno (sottrae 40 per le donne)
        giorno_cf = int(cf[9:11])
        giorno = giorno_cf - 40 if giorno_cf > 40 else giorno_cf

        # Codice Catastale
        codice_catastale = cf[11:15]

        # Verifica e formatta la data
        data = datetime(anno, mese, giorno)
        return data.strftime("%d/%m/%Y"), codice_catastale
    except (ValueError, KeyError, IndexError): 
        # Cattura qualsiasi errore di conversione o indice
        return "N/A", "N/A"

def estrai_dati(filepath):
    # Caricamento e pulizia testo
    reader = PdfReader(filepath)
    text = "".join(page.extract_text() + "\n" for page in reader.pages)
    righe = [r.strip() for r in text.splitlines() if r.strip()]

    # Variabili output
    dati = []
    ragione_sociale = "NON TROVATA"
    forma_giuridica = "NON TROVATA"
    comune = "NON TROVATO"
    via = "NON TROVATO"
    addetti = "NON TROVATO"

    # --- PARSING RIGHE (Dati Societari) ---
    for i, riga in enumerate(righe):
        
        # 1. Ragione Sociale
        if ("VISURA" in riga or "FASCICOLO" in riga) and ragione_sociale == "NON TROVATA":
            try:
                # Cerca la prima riga sensata dopo l'intestazione
                for offset in range(1, 5): 
                    if i + offset < len(righe):
                        linea = righe[i + offset]
                        if not any(word in linea for word in ["Camera di Commercio", "Registro Imprese", "Documento n."]):
                            ragione_sociale = linea.split("Codice Fiscale")[0].strip() # Pulisce se il CF è sulla stessa riga
                            break
            except: pass

        # 2. Forma Giuridica
        if "Forma giuridica" in riga:
            try:
                # Prende la parte successiva alla keyword, sia sulla stessa riga che a capo
                match = re.search(r'Forma giuridica\s*(.*)', riga)
                if match and match.group(1).strip():
                    forma_giuridica = match.group(1).strip()
                elif i + 1 < len(righe):
                    # Prova a prendere il valore dalla riga successiva (se è vuota sulla corrente)
                    forma_giuridica = righe[i+1]
            except: pass
            
        # 3. Sede Legale (Comune e Via)
        if "Indirizzo Sede" in riga:
            try:
                full_text = riga
                if i + 1 < len(righe) and "CAP" not in riga: 
                    full_text += " " + righe[i+1]
                
                # Cerca Comune (prima di (PROV))
                match_comune = re.search(r'Sede legale\s*(.*?)\s*\(', full_text)
                if match_comune:
                    comune = match_comune.group(1).strip()
                
                # Cerca Via (dopo la sigla della provincia e prima di CAP)
                match_via = re.search(r'\([A-Z]{2}\)\s*(.*?)\s*CAP', full_text)
                if match_via:
                    via = match_via.group(1).strip()
                elif "CAP" in full_text: # Fallback se manca la sigla provincia ma c'è CAP
                     match_via_fallback = re.search(r'Sede legale.*?CAP', full_text, re.DOTALL)
                     if match_via_fallback:
                         # Logica complessa, ma l'obiettivo è evitare il crash. Usiamo i dati trovati.
                         pass
            except: pass

        # 4. Addetti
        if "Addetti" in riga:
            match = re.search(r'(\d+)$', riga)
            if match: addetti = match.group(1)

    # --- ESTRAZIONE PERSONE (Nominativi) ---
    keywords_sezioni = [
        "Soci e titolari", "Amministratori", "Soci accomandatari", 
        "Soci accomandanti", "Titolari di cariche", "Sindaci"
    ]
    pattern_cf = r"[A-Z]{6}\d{2}[A-Z]\d{2}[A-Z]\d{3}[A-Z]"
    
    current_section = "Generico/Da verificare"
    codici_visti = set()

    for riga in righe:
        # Aggiorna la sezione corrente
        for key in keywords_sezioni:
            if key in riga:
                # Se è una sezione di "chiusura" o una sezione superiore, resettiamo
                if "Soci e titolari" in key or "Amministratori" in key:
                     current_section = key
                break
        
        # Estrai CF
        match = re.search(pattern_cf, riga)
        if match:
            cf = match.group()
            if cf not in codici_visti:
                
                # Estrazione Nominativo: prende le parole in maiuscolo prima del CF (euristica)
                parts = riga.split(cf)[0].strip().split()
                valid_names = [p for p in parts if len(p) > 2 and p.isalpha() and p.isupper()]
                nome_completo = " ".join(valid_names)
                
                if not nome_completo: 
                    # Fallback per nominativo a capo: prova la riga precedente
                    if i > 0:
                        prev_parts = righe[i-1].strip().split()
                        prev_valid_names = [p for p in prev_parts if len(p) > 2 and p.isalpha() and p.isupper()]
                        nome_completo = " ".join(prev_valid_names)

                nome_completo = nome_completo or "NOMINATIVO DA VERIFICARE"

                # Decodifica CF
                data_nascita, codice_catastale = decodifica_data_nascita(cf)
                
                # Suddivisione nome/cognome (euristica)
                parole = nome_completo.split()
                if len(parole) >= 2:
                    # Assumiamo Cognome è la prima parola, il resto è Nome (semplificazione)
                    cognome = parole[0]
                    nome = " ".join(parole[1:])
                else:
                    cognome = nome_completo
                    nome = ""

                dati.append({
                    "Cognome": cognome,
                    "Nomi": nome,
                    "Codice Fiscale": cf,
                    "Data di nascita": data_nascita,
                    "Codice catastale": codice_catastale,
                    "Ruolo/Sezione": current_section
                })
                codici_visti.add(cf)

    return dati, ragione_sociale, forma_giuridica, comune, via, addetti

# --- INTERFACCIA STREAMLIT ---
st.title("📄 Analisi Visura Telemaco")
uploaded = st.file_uploader("Carica PDF", type="pdf")

if uploaded:
    # Gestione temporanea del file
    temp_file_path = "temp.pdf"
    with open(temp_file_path, "wb") as f: f.write(uploaded.getbuffer())
    
    try:
        with st.spinner('Elaborazione dati in corso...'):
            dati, rag_soc, forma, comune, via, addetti = estrai_dati(temp_file_path)
        
        # Visualizza Card Dati
        st.markdown(f"""
        <div class="data-card">
            <h3>🏢 Ragione Sociale: {rag_soc}</h3>
            <p><strong>⚖️ Forma:</strong> {forma} | <strong>👥 Addetti:</strong> {addetti}</p>
            <p><strong>📍 Sede:</strong> {via}, {comune}</p>
        </div>
        """, unsafe_allow_html=True)

        if dati:
            # Creazione del DataFrame in modo sicuro
            df = pd.DataFrame(dati)
            
            st.success(f"✅ Trovati {len(df)} nominativi.")
            st.dataframe(df, use_container_width=True)
            
            # Export Excel
            file_excel = "Elenco_per_casellario.xlsx"
            
            # Formattazione e salvataggio Excel
            with pd.ExcelWriter(file_excel, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
                
                # Auto-fit colonne (migliorato)
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = max_length + 2
                    worksheet.column_dimensions[column].width = adjusted_width

            # Pulsante di download
            with open(file_excel, "rb") as f:
                st.download_button("📥 Scarica Excel", f, file_name=file_excel)
        else:
            st.warning("⚠️ Nessun nominativo trovato. (Codice Fiscale non riconosciuto).")
            
    except Exception as e:
        # Mostra un errore generico se il PDF è illeggibile o il parser fallisce
        st.error(f"❌ Errore critico durante l'elaborazione del file: {type(e).__name__} - {e}")
        st.info("Verifica che il file sia una visura camerale ufficiale in formato testo (non scansione immagine).")