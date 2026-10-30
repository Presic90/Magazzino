import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import os
import datetime
import time
import base64

st.set_page_config(page_title="Test Magazzino RFID", page_icon="📦", layout="wide")

st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .stButton>button {
        width: 100%;
        height: 65px !important;
        border-radius: 12px !important;
        font-size: 18px !important;
        font-weight: bold !important;
    }
    
    .stTextInput>div>div>input {
        font-size: 16px !important;
        padding: 14px !important;
        border-radius: 8px !important;
    }
    
    .block-container {
        padding-top: 2rem;
        padding-left: 1rem;
        padding-right: 1rem;
        padding-bottom: 2rem;
    }
</style>
""", unsafe_allow_html=True)

FILE_CATALOGO = "MagazzinoUfficiale2026(categorie).xlsx"
FILE_LIVE = "Magazzino_Live.xlsx"
FILE_STORICO = "Storico_Movimenti.xlsx"
FILE_ISO = "Registro_ISO/Registro_Strumenti_Finale_2.xlsx" 
NOME_LOGO = "logo.png"

if 'ultimo_cantiere' not in st.session_state:
    st.session_state['ultimo_cantiere'] = ""

@st.cache_data(show_spinner=False)
def inizializza_database_live():
    if os.path.exists(FILE_LIVE):
        df_live = pd.read_excel(FILE_LIVE, dtype=str)
        return df_live.fillna("")
    
    st.info("Importazione dal Catalogo Master in corso...")
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
        return pd.read_excel(FILE_STORICO, dtype=str).fillna("")
    return pd.DataFrame(columns=["Data", "Ora", "Operazione", "CODICE RFID", "Strumento", "Modello", "Cantiere / Assegnatario"])

def non_valido(valore):
    val_str = str(valore).strip().lower()
    return val_str == "" or val_str == "nan"

def verifica_taratura_iso(matricola):
    if not os.path.exists(FILE_ISO):
        return "FILE_NON_TROVATO"
    if non_valido(matricola):
        return "MATRICOLA_MANCANTE"
    try:
        df_iso = pd.read_excel(FILE_ISO)
        df_iso.columns = df_iso.columns.str.strip().str.upper()
        if 'MATRICOLA' in df_iso.columns and 'SCADENZA' in df_iso.columns:
            df_iso['MATRICOLA_CLEAN'] = df_iso['MATRICOLA'].astype(str).str.strip().str.upper()
            mat_cercata = str(matricola).strip().upper()
            strumento_iso = df_iso[df_iso['MATRICOLA_CLEAN'] == mat_cercata]
            
            if not strumento_iso.empty:
                scadenza_val = strumento_iso.iloc[0]['SCADENZA']
                if pd.notnull(scadenza_val) and str(scadenza_val).strip() != "":
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
    <head><meta charset="utf-8"><title>Bolla_{tipo_bolla}_{cantiere}</title>
    <style>body {{ font-family: Arial, sans-serif; margin: 40px; }} table {{ width: 100%; border-collapse: collapse; }} th, td {{ border: 1px solid #ddd; padding: 12px; }}</style>
    </head>
    <body onload="window.print()">
        {tag_logo}
        <h1 style="color:{colore_tema}">{titolo_doc} - Cantiere: {cantiere.upper()}</h1>
        <table><thead><tr><th>Codice RFID</th><th>Strumento</th><th>Modello</th><th>Note</th></tr></thead><tbody>{righe}</tbody></table>
    </body></html>
    """
    return html

df_magazzino = inizializza_database_live()

col_logo, col_menu = st.columns([1, 2], vertical_alignment="center")
with col_logo:
    if os.path.exists(NOME_LOGO):
        st.image(NOME_LOGO, width=120)
with col_menu:
    modulo = st.selectbox("Menu di Navigazione:", [
        "📡 Scanner RFID (In/Out)", 
        "📝 Uscita/Rientro Manuale",
        "📊 Situazione Attuale", 
        "🏗️ Cantieri e Bolle", 
        "📜 Storico Movimenti"
    ], label_visibility="collapsed")
st.divider()

if modulo == "📡 Scanner RFID (In/Out)":
    st.markdown("### 📡 Movimentazione (RFID)")
    codice_scansionato = st.text_input("Codice RFID:", key="scanner", placeholder="Spara il codice qui...")
    components.html("""<script>var i=window.parent.document.querySelector("input[aria-label='Codice RFID:']"); if(i)i.focus();</script>""", height=0)
    
    if codice_scansionato:
        codice_grezzo = str(codice_scansionato).strip()
        strumenti_trovati = [row for _, row in df_magazzino.iterrows() if str(row['CODICE RFID']).strip() in codice_grezzo and str(row['CODICE RFID']).strip()]
        
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
            
            st.success(f"**{nome}** ({modello}) - Matr: {matricola_strumento}")
            
            if stato_attuale == "🟢 In Sede":
                stato_iso = verifica_taratura_iso(matricola_strumento)
                if stato_iso == "REGOLARE": st.success("✅ Taratura Regolare")
                elif stato_iso == "SCADUTO": st.error("⚠️ TARATURA SCADUTA!")
                else: st.caption(f"ℹ️ Info ISO: {stato_iso}")
                
                assegnatario = st.text_input("Assegna a Cantiere/Persona:", value=st.session_state['ultimo_cantiere'])
                bottone_testo = "⚠️ FORZA USCITA" if stato_iso == "SCADUTO" else "🔴 Registra USCITA"
                bottone_tipo = "secondary" if stato_iso == "SCADUTO" else "primary"

                if st.button(bottone_testo, type=bottone_tipo):
                    if assegnatario:
                        st.session_state['ultimo_cantiere'] = assegnatario
                        df_magazzino.at[idx, 'Stato'] = "🔴 In Cantiere"
                        df_magazzino.at[idx, 'Assegnatario / Cantiere'] = assegnatario
                        df_magazzino.at[idx, 'Ultima Modifica'] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
                        nota = "🔴 USCITA (FORZATA)" if stato_iso == "SCADUTO" else "🔴 USCITA"
                        if salva_live(df_magazzino):
                            salva_in_storico(codice_pulito, nome, modello, nota, assegnatario)
                            st.success("✅ Registrato!")
                            st.cache_data.clear()
                            time.sleep(0.5)
                            st.rerun()
                    else:
                        st.warning("Inserisci il cantiere!")
            else:
                vecchio_assegnatario = strumento['Assegnatario / Cantiere']
                st.warning(f"Fuori in: **{vecchio_assegnatario}**")
                st.session_state['ultimo_cantiere'] = vecchio_assegnatario
                if st.button("🟢 Registra RIENTRO", type="primary"):
                    df_magazzino.at[idx, 'Stato'] = "🟢 In Sede"
                    df_magazzino.at[idx, 'Assegnatario / Cantiere'] = ""
                    df_magazzino.at[idx, 'Ultima Modifica'] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
                    if salva_live(df_magazzino):
                        salva_in_storico(codice_pulito, nome, modello, "🟢 RIENTRO", vecchio_assegnatario)
                        st.success("✅ Rientro registrato!")
                        st.cache_data.clear()
                        time.sleep(0.5)
                        st.rerun()

elif modulo == "📝 Uscita/Rientro Manuale":
    st.markdown("### 📝 Assegnazione Manuale Multipla")
    operazione = st.radio("Seleziona operazione:", ["🔴 Uscita verso Cantiere", "🟢 Rientro in Sede"], horizontal=True)
    
    if "Uscita" in operazione:
        disponibili = df_magazzino[df_magazzino['Stato'] == "🟢 In Sede"]
        opzioni_dict = {f"{row['STRUMENTAZIONE']} - {row['MODELLO/TIPOLOGIA']} (Matr: {row['MATRICOLA']}) - ID:{idx}": idx for idx, row in disponibili.iterrows()}
        
        selezionati_chiavi = st.multiselect("Cerca e seleziona strumenti da far uscire:", list(opzioni_dict.keys()))
        assegnatario = st.text_input("Assegna a Cantiere/Persona:", value=st.session_state['ultimo_cantiere'])
        
        if st.button("🔴 Registra Uscita di Massa", type="primary"):
            if selezionati_chiavi and assegnatario:
                st.session_state['ultimo_cantiere'] = assegnatario
                for chiave in selezionati_chiavi:
                    idx = opzioni_dict[chiave]
                    strumento = df_magazzino.loc[idx]
                    df_magazzino.at[idx, 'Stato'] = "🔴 In Cantiere"
                    df_magazzino.at[idx, 'Assegnatario / Cantiere'] = assegnatario
                    df_magazzino.at[idx, 'Ultima Modifica'] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
                    salva_in_storico(strumento['CODICE RFID'], strumento['STRUMENTAZIONE'], strumento['MODELLO/TIPOLOGIA'], "🔴 USCITA (MANUALE)", assegnatario)
                
                salva_live(df_magazzino)
                st.success(f"✅ {len(selezionati_chiavi)} strumenti registrati in uscita!")
                st.cache_data.clear()
                time.sleep(1)
                st.rerun()
            else:
                st.warning("Seleziona almeno uno strumento e inserisci il cantiere.")
                
    else:
        fuori = df_magazzino[df_magazzino['Stato'] != "🟢 In Sede"]
        opzioni_dict = {f"{row['STRUMENTAZIONE']} - {row['MODELLO/TIPOLOGIA']} (Da: {row['Assegnatario / Cantiere']}) - ID:{idx}": idx for idx, row in fuori.iterrows()}
        
        selezionati_chiavi = st.multiselect("Cerca e seleziona strumenti da far rientrare:", list(opzioni_dict.keys()))
        
        if st.button("🟢 Registra Rientro di Massa", type="primary"):
            if selezionati_chiavi:
                for chiave in selezionati_chiavi:
                    idx = opzioni_dict[chiave]
                    strumento = df_magazzino.loc[idx]
                    vecchio_cantiere = strumento['Assegnatario / Cantiere']
                    df_magazzino.at[idx, 'Stato'] = "🟢 In Sede"
                    df_magazzino.at[idx, 'Assegnatario / Cantiere'] = ""
                    df_magazzino.at[idx, 'Ultima Modifica'] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
                    salva_in_storico(strumento['CODICE RFID'], strumento['STRUMENTAZIONE'], strumento['MODELLO/TIPOLOGIA'], "🟢 RIENTRO (MANUALE)", vecchio_cantiere)
                
                salva_live(df_magazzino)
                st.success(f"✅ {len(selezionati_chiavi)} strumenti rientrati!")
                st.cache_data.clear()
                time.sleep(1)
                st.rerun()
            else:
                st.warning("Seleziona almeno uno strumento.")

elif modulo == "📊 Situazione Attuale":
    st.markdown("### 📦 Inventario e Giacenze")
    filtro = st.radio("Filtra:", ["Tutti", "🟢 In Sede", "🔴 In Cantiere"], horizontal=True)
    df_v = df_magazzino.copy() if filtro == "Tutti" else df_magazzino[df_magazzino['Stato'] == filtro]
    st.metric("Totale Strumenti", len(df_v))
    st.dataframe(df_v[['Stato', 'STRUMENTAZIONE', 'MODELLO/TIPOLOGIA', 'Assegnatario / Cantiere']], use_container_width=True, hide_index=True)

elif modulo == "🏗️ Cantieri e Bolle":
    st.markdown("### 🏗️ Distinte")
    df_st = carica_storico()
    cantieri = sorted([str(c) for c in df_st['Cantiere / Assegnatario'].unique() if str(c).strip() != ""])
    if cantieri:
        cant = st.selectbox("Scegli Cantiere:", ["-- Seleziona --"] + cantieri)
        if cant != "-- Seleziona --":
            tipo = st.radio("Fase:", ["🔴 Uscita", "🟢 Rientro"], horizontal=True)
            filtro = "USCITA" if "Uscita" in tipo else "RIENTRO"
            df_kit = df_st[(df_st['Cantiere / Assegnatario'] == cant) & (df_st['Operazione'].str.contains(filtro))].drop_duplicates(subset=['CODICE RFID'])
            st.dataframe(df_kit[['Strumento', 'Data']], use_container_width=True, hide_index=True)
            if not df_kit.empty:
                st.download_button("🖨️ Stampa", data=genera_bolla_html(cant, df_kit, filtro), file_name=f"Bolla_{cant}.html", mime="text/html", type="primary")

elif modulo == "📜 Storico Movimenti":
    st.markdown("### 📜 Movimentazioni")
    df_st = carica_storico()
    df_edit = st.data_editor(df_st.iloc[::-1].copy(), use_container_width=True, hide_index=True)
    if st.button("💾 Salva Modifiche", type="secondary"):
        with pd.ExcelWriter(FILE_STORICO, engine="openpyxl", mode="w") as w:
            df_edit.iloc[::-1].to_excel(w, index=False)
        st.success("✅ Aggiornato!")
        st.rerun()