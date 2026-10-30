import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import os
import datetime
import time
import base64

st.set_page_config(page_title="Test Magazzino RFID", page_icon="📦", layout="wide")

FILE_CATALOGO = "MagazzinoUfficiale2026(categorie).xlsx"
FILE_LIVE = "Magazzino_Live.xlsx"
FILE_STORICO = "Storico_Movimenti.xlsx"
# IL PERCORSO CORRETTO: Cerca il file dentro la cartella Registro_ISO
FILE_ISO = "Registro_ISO/Registro_Strumenti_Finale_2.xlsx" 
NOME_LOGO = "logo.png"

# --- MOSTRA IL LOGO NELLA BARRA LATERALE ---
if os.path.exists(NOME_LOGO):
    st.sidebar.image(NOME_LOGO, use_container_width=True)

# --- GESTIONE MEMORIA SESSIONE ---
if 'ultimo_cantiere' not in st.session_state:
    st.session_state['ultimo_cantiere'] = ""

# --- 1. INIZIALIZZAZIONE DATI E PONTE ISO CORAZZATO ---
@st.cache_data(show_spinner=False)
def inizializza_database_live():
    if os.path.exists(FILE_LIVE):
        df_live = pd.read_excel(FILE_LIVE, dtype=str)
        return df_live.fillna("")
    
    st.info("Prima esecuzione: Importazione dal Catalogo Master in corso...")
    fogli = pd.ExcelFile(FILE_CATALOGO).sheet_names
    tutti_i_dati = []
    
    for foglio in fogli:
        df_temp = pd.read_excel(FILE_CATALOGO, sheet_name=foglio, dtype=str)
        try:
            indice_header = df_temp[df_temp.iloc[:, 0] == 'CODICE RFID'].index[0]
            df_pulito = pd.read_excel(FILE_CATALOGO, sheet_name=foglio, header=indice_header + 1, dtype=str)
            df_pulito['Categoria_Foglio'] = foglio.replace("Magazzino_", "")
            tutti_i_dati.append(df_pulito)
        except IndexError:
            pass 

    df_completo = pd.concat(tutti_i_dati, ignore_index=True)
    df_completo = df_completo.dropna(subset=['CODICE RFID'])
    df_completo['CODICE RFID'] = df_completo['CODICE RFID'].astype(str).str.strip().apply(lambda x: x.zfill(24))
    
    df_completo['Stato'] = "🟢 In Sede"
    df_completo['Assegnatario / Cantiere'] = ""
    df_completo['Ultima Modifica'] = ""
    
    df_completo.to_excel(FILE_LIVE, index=False)
    return df_completo.fillna("")

def carica_storico():
    if os.path.exists(FILE_STORICO):
        df_s = pd.read_excel(FILE_STORICO, dtype=str)
        return df_s.fillna("")
    return pd.DataFrame(columns=["Data", "Ora", "Operazione", "CODICE RFID", "Strumento", "Modello", "Cantiere / Assegnatario"])

def non_valido(valore):
    val_str = str(valore).strip().lower()
    return val_str == "" or val_str == "nan"

def verifica_taratura_iso(matricola):
    """Ponte Corazzato: gestisce percorsi, spazi, maiuscole e date native di Excel."""
    if not os.path.exists(FILE_ISO):
        return "FILE_NON_TROVATO"
        
    if non_valido(matricola):
        return "MATRICOLA_MANCANTE"
    
    try:
        # Leggiamo il file permettendo a Pandas di riconoscere le date native
        df_iso = pd.read_excel(FILE_ISO)
        
        # Rendiamo i nomi delle colonne tutti MAIUSCOLI e senza spazi laterali
        df_iso.columns = df_iso.columns.str.strip().str.upper()
        
        if 'MATRICOLA' in df_iso.columns and 'SCADENZA' in df_iso.columns:
            # Puliamo la colonna Matricola da eventuali spazi o minuscole
            df_iso['MATRICOLA_CLEAN'] = df_iso['MATRICOLA'].astype(str).str.strip().str.upper()
            mat_cercata = str(matricola).strip().upper()
            
            strumento_iso = df_iso[df_iso['MATRICOLA_CLEAN'] == mat_cercata]
            
            if not strumento_iso.empty:
                scadenza_val = strumento_iso.iloc[0]['SCADENZA']
                if pd.notnull(scadenza_val) and str(scadenza_val).strip() != "":
                    # Convertiamo in data capendo sia le date testuali (es 15/05) sia quelle di excel
                    scadenza_data = pd.to_datetime(scadenza_val, errors='coerce', dayfirst=True)
                    if pd.notnull(scadenza_data):
                        if scadenza_data < datetime.datetime.now():
                            return "SCADUTO"
                        else:
                            return "REGOLARE"
            return "NON_IN_REGISTRO"
        else:
            return "COLONNE_MANCANTI"
    except Exception as e:
        return f"ERRORE: {e}"

# --- 2. FUNZIONI DI SUPPORTO ---
def salva_live(df):
    try:
        df.to_excel(FILE_LIVE, index=False)
        return True
    except:
        return False

def salva_in_storico(rfid, nome, mod, op, cant):
    df_s = carica_storico()
    now = datetime.datetime.now()
    nuova_riga = pd.DataFrame([{
        "Data": now.strftime("%d/%m/%Y"),
        "Ora": now.strftime("%H:%M:%S"),
        "Operazione": op,
        "CODICE RFID": rfid,
        "Strumento": nome,
        "Modello": mod,
        "Cantiere / Assegnatario": cant
    }])
    df_s = pd.concat([df_s, nuova_riga], ignore_index=True)
    try:
        with pd.ExcelWriter(FILE_STORICO, engine="openpyxl", mode="w") as writer:
            df_s.to_excel(writer, index=False)
    except:
        pass

def ottieni_logo_base64(path_immagine):
    if os.path.exists(path_immagine):
        with open(path_immagine, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    return None

def genera_bolla_html(cantiere, df_kit, tipo_bolla="USCITA"):
    righe = ""
    for _, row in df_kit.iterrows():
        righe += f"<tr><td>{row['CODICE RFID']}</td><td>{row['Strumento']}</td><td>{row['Modello']}</td><td></td></tr>"
    
    data_oggi = datetime.datetime.now().strftime("%d/%m/%Y")
    titolo_doc = "DISTINTA USCITA MATERIALI" if tipo_bolla == "USCITA" else "DISTINTA RIENTRO MATERIALI"
    colore_tema = "#0A2540" if tipo_bolla == "USCITA" else "#1A4D2E" 
    
    logo_b64 = ottieni_logo_base64(NOME_LOGO)
    estensione = "jpeg" if "jpg" in NOME_LOGO.lower() else "png"
    tag_logo = f'<img src="data:image/{estensione};base64,{logo_b64}" class="logo">' if logo_b64 else ""
    
    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <title>Bolla_{tipo_bolla}_{cantiere}</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 40px; color: #333; }}
            .logo {{ max-height: 80px; margin-bottom: 15px; }}
            h1 {{ color: {colore_tema}; margin-bottom: 5px; font-size: 24px; }}
            .header {{ display: flex; justify-content: space-between; margin-bottom: 40px; border-bottom: 2px solid {colore_tema}; padding-bottom: 10px; }}
            table {{ width: 100%; border-collapse: collapse; margin-bottom: 50px; font-size: 14px; }}
            th, td {{ border: 1px solid #ddd; padding: 12px; text-align: left; }}
            th {{ background-color: #f4f7fb; color: {colore_tema}; }}
            .signatures {{ display: flex; justify-content: space-between; margin-top: 70px; }}
            .sig-box {{ width: 40%; border-top: 1px solid #000; padding-top: 10px; text-align: center; font-weight: bold; font-size: 14px; }}
            @media print {{ body {{ margin: 0; }} @page {{ margin: 1cm; }} }}
        </style>
    </head>
    <body onload="window.print()">
        <div class="header">
            <div>
                {tag_logo}
                <h1>{titolo_doc}</h1>
                <h2>Progetto/Cantiere: {cantiere.upper()}</h2>
            </div>
            <div style="text-align: right; margin-top: 20px;">
                <p><strong>Data Stampa:</strong> {data_oggi}</p>
                <p><strong>Totale Strumenti:</strong> {len(df_kit)}</p>
            </div>
        </div>
        <table>
            <thead>
                <tr>
                    <th style="width:25%">Codice RFID</th>
                    <th style="width:35%">Strumento</th>
                    <th style="width:25%">Modello</th>
                    <th style="width:15%">Note / Check</th>
                </tr>
            </thead>
            <tbody>{righe}</tbody>
        </table>
        <div class="signatures">
            <div class="sig-box">Firma Operatore / Tecnico</div>
            <div class="sig-box">Firma Responsabile Magazzino</div>
        </div>
    </body>
    </html>
    """
    return html

df_magazzino = inizializza_database_live()

# ==========================================
# --- MENU LATERALE ---
# ==========================================
with st.sidebar:
    st.markdown("### 🧭 Menu Magazzino")
    modulo = st.radio("Scegli operazione:", [
        "📡 Scanner RFID (In/Out)", 
        "📊 Situazione Attuale", 
        "🏗️ Cantieri e Bolle", 
        "📜 Storico Movimenti"
    ])

# ==========================================
# --- MODULO 1: SCANNER RFID ---
# ==========================================
if modulo == "📡 Scanner RFID (In/Out)":
    st.title("📡 Movimentazione Magazzino (RFID)")
    st.markdown("Usa la **Zebra RFD40** per scansionare lo strumento.")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        codice_scansionato = st.text_input("Codice RFID:", key="scanner", placeholder="Spara il codice qui...")
        components.html(
            """<script>
            var input = window.parent.document.querySelector("input[aria-label='Codice RFID:']");
            if (input) { input.focus(); }
            </script>""", height=0
        )
    
    if codice_scansionato:
        codice_grezzo = str(codice_scansionato).strip()
        strumenti_trovati = []
        for idx, row in df_magazzino.iterrows():
            cod_rfid_valido = str(row['CODICE RFID']).strip()
            if cod_rfid_valido and cod_rfid_valido in codice_grezzo:
                strumenti_trovati.append(row)
        
        if not strumenti_trovati:
            st.error("❌ Nessun codice riconosciuto.")
        else:
            strumento = strumenti_trovati[0]
            idx = strumento.name
            nome = strumento['STRUMENTAZIONE']
            modello = strumento['MODELLO/TIPOLOGIA']
            stato_attuale = strumento['Stato']
            codice_pulito = strumento['CODICE RFID']
            matricola_strumento = strumento['MATRICOLA']
            
            with col2:
                st.success(f"**Strumento Identificato:** {nome} ({modello}) - Matr: {matricola_strumento}")
                
                if stato_attuale == "🟢 In Sede":
                    st.info("📦 Lo strumento è attualmente IN SEDE.")
                    
                    # --- CONTROLLO ISO TRAMITE MATRICOLA ---
                    stato_iso = verifica_taratura_iso(matricola_strumento)
                    
                    # FEEDBACK VISIVO PER IL DEBUG
                    if stato_iso == "FILE_NON_TROVATO":
                        st.caption(f"ℹ️ File non trovato al percorso: {FILE_ISO}. Verifica che la cartella esista.")
                    elif stato_iso == "MATRICOLA_MANCANTE":
                        st.caption("ℹ️ Strumento senza matricola, impossibile verificare certificazione.")
                    elif stato_iso == "NON_IN_REGISTRO":
                        st.caption(f"ℹ️ Nessuna corrispondenza ISO per la matricola: {matricola_strumento}")
                    elif stato_iso == "COLONNE_MANCANTI":
                        st.caption("⚠️ Il Registro ISO deve avere le colonne chiamate esattamente 'Matricola' e 'Scadenza'.")
                    elif stato_iso == "REGOLARE":
                        st.success("✅ Certificazione ISO in regola.")
                    elif stato_iso.startswith("ERRORE"):
                        st.error(f"Errore lettura ISO: {stato_iso}")
                    
                    assegnatario = st.text_input("Assegna a (Cantiere o Persona):", value=st.session_state['ultimo_cantiere'])
                    
                    if stato_iso == "SCADUTO":
                        st.error("⚠️ ATTENZIONE: Questo strumento risulta con TARATURA SCADUTA nel registro ISO!")
                        bottone_testo = "⚠️ FORZA USCITA (Sotto mia responsabilità)"
                        bottone_tipo = "secondary"
                    else:
                        bottone_testo = "🔴 Registra USCITA"
                        bottone_tipo = "primary"

                    if st.button(bottone_testo, type=bottone_tipo):
                        if assegnatario:
                            st.session_state['ultimo_cantiere'] = assegnatario
                            df_magazzino.at[idx, 'Stato'] = "🔴 In Cantiere"
                            df_magazzino.at[idx, 'Assegnatario / Cantiere'] = assegnatario
                            df_magazzino.at[idx, 'Ultima Modifica'] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
                            
                            nota_operazione = "🔴 USCITA (FORZATA - TARATURA SCADUTA)" if stato_iso == "SCADUTO" else "🔴 USCITA"
                            
                            if salva_live(df_magazzino):
                                salva_in_storico(codice_pulito, nome, modello, nota_operazione, assegnatario)
                                st.success("✅ Uscita registrata! Pronto per il prossimo.")
                                st.cache_data.clear()
                                st.rerun()
                        else:
                            st.warning("Inserisci il nome del cantiere!")
                
                else:
                    vecchio_assegnatario = strumento['Assegnatario / Cantiere']
                    st.warning(f"🏗️ Lo strumento è fuori! (Assegnato a: **{vecchio_assegnatario}**)")
                    st.session_state['ultimo_cantiere'] = vecchio_assegnatario
                    
                    if st.button("🟢 Registra RIENTRO", type="primary"):
                        df_magazzino.at[idx, 'Stato'] = "🟢 In Sede"
                        df_magazzino.at[idx, 'Assegnatario / Cantiere'] = ""
                        df_magazzino.at[idx, 'Ultima Modifica'] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
                        
                        if salva_live(df_magazzino):
                            salva_in_storico(codice_pulito, nome, modello, "🟢 RIENTRO", vecchio_assegnatario)
                            st.success("✅ Rientro registrato!")
                            st.cache_data.clear()
                            st.rerun()

# ==========================================
# GLI ALTRI MODULI RIMANGONO INVARIATI
# ==========================================
elif modulo == "📊 Situazione Attuale":
    st.title("📦 Inventario e Giacenze")
    filtro_stato = st.radio("Filtra per stato:", ["Tutti", "🟢 In Sede", "🔴 In Cantiere"], horizontal=True)
    df_view = df_magazzino.copy()
    if filtro_stato != "Tutti":
        df_view = df_view[df_view['Stato'] == filtro_stato]
    st.metric(f"Totale Strumenti visualizzati", len(df_view))
    colonne_da_mostrare = ['Stato', 'CODICE RFID', 'Categoria_Foglio', 'STRUMENTAZIONE', 'MODELLO/TIPOLOGIA', 'MATRICOLA', 'Assegnatario / Cantiere', 'Ultima Modifica']
    st.dataframe(df_view[[c for c in colonne_da_mostrare if c in df_view.columns]], use_container_width=True, hide_index=True, height=600)

elif modulo == "🏗️ Cantieri e Bolle":
    st.title("🏗️ Riepilogo Cantieri e Distinte")
    df_storico = carica_storico()
    if df_storico.empty:
        st.info("Nessun movimento registrato finora.")
    else:
        lista_cantieri = sorted([str(c) for c in df_storico['Cantiere / Assegnatario'].unique() if str(c).strip() != ""])
        if not lista_cantieri:
            st.info("Nessun cantiere rilevato.")
        else:
            cantiere_selezionato = st.selectbox("📂 Scegli Cantiere:", ["-- Seleziona --"] + lista_cantieri)
            if cantiere_selezionato != "-- Seleziona --":
                tipo_operazione = st.radio("Seleziona la fase:", ["🔴 Materiale in Uscita", "🟢 Materiale Rientrato"], horizontal=True)
                filtro_testo = "USCITA" if "Uscita" in tipo_operazione else "RIENTRO"
                titolo_bolla = "USCITA" if "Uscita" in tipo_operazione else "RIENTRO"
                df_storico_view = df_storico[df_storico['Cantiere / Assegnatario'] == cantiere_selezionato]
                df_kit = df_storico_view[df_storico_view['Operazione'].str.contains(filtro_testo)].drop_duplicates(subset=['CODICE RFID'])
                if not df_kit.empty:
                    st.dataframe(df_kit[['Strumento', 'Modello', 'CODICE RFID', 'Data']], use_container_width=True, hide_index=True)
                    st.download_button(label=f"🖨️ Stampa Bolla", data=genera_bolla_html(cantiere_selezionato, df_kit, titolo_bolla), file_name=f"Bolla_{titolo_bolla}_{cantiere_selezionato}.html", mime="text/html", type="primary")

elif modulo == "📜 Storico Movimenti":
    st.title("📜 Registro Completo Movimentazioni")
    df_storico = carica_storico()
    if df_storico.empty:
        st.info("Nessun movimento registrato finora.")
    else:
        df_storico_editato = st.data_editor(df_storico.iloc[::-1].copy(), use_container_width=True, hide_index=True, num_rows="dynamic", height=600)
        if st.button("💾 Salva Modifiche Storico", type="secondary"):
            try:
                with pd.ExcelWriter(FILE_STORICO, engine="openpyxl", mode="w") as writer:
                    df_storico_editato.iloc[::-1].to_excel(writer, index=False)
                st.success("✅ Storico aggiornato!")
                time.sleep(0.5)
                st.rerun()
            except Exception as e:
                st.error(f"⚠️ Errore salvataggio: {e}")