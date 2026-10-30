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
    .stTextInput>div>div>input, .stTextArea>div>div>textarea {
        font-size: 16px !important;
        padding: 14px !important;
        border-radius: 8px !important;
    }
    .block-container {
        padding-top: 2rem; padding-left: 1rem; padding-right: 1rem; padding-bottom: 2rem;
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
        return pd.read_excel(FILE_LIVE, dtype=str).fillna("")
    
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
        righe += f"<tr><td style='font-family: monospace;'>{row['CODICE RFID']}</td><td>{row['Strumento']}</td><td>{row['Modello']}</td><td></td></tr>"
    
    data_oggi = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    titolo_doc = "DISTINTA DI CONSEGNA MATERIALI" if tipo_bolla == "USCITA" else "DISTINTA DI RIENTRO MATERIALI"
    colore_tema = "#0A2540" if tipo_bolla == "USCITA" else "#1A4D2E" 
    
    logo_b64 = ottieni_logo_base64(NOME_LOGO)
    estensione = "jpeg" if "jpg" in NOME_LOGO.lower() else "png"
    tag_logo = f'<img src="data:image/{estensione};base64,{logo_b64}" class="logo">' if logo_b64 else ""
    
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>Bolla_{tipo_bolla}_{cantiere}</title>
        <style>
            body {{ font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; margin: 40px; color: #333; }}
            .logo {{ max-height: 80px; margin-bottom: 10px; }}
            .header-container {{ display: flex; justify-content: space-between; align-items: flex-end; border-bottom: 3px solid {colore_tema}; padding-bottom: 15px; margin-bottom: 30px; }}
            .doc-title {{ margin: 0; font-size: 26px; color: {colore_tema}; text-transform: uppercase; letter-spacing: 1px; }}
            .doc-subtitle {{ margin: 5px 0 0 0; font-size: 18px; color: #555; }}
            .doc-meta {{ text-align: right; font-size: 14px; line-height: 1.5; }}
            table {{ width: 100%; border-collapse: collapse; margin-bottom: 40px; font-size: 13px; }}
            th, td {{ border: 1px solid #dcdcdc; padding: 12px 15px; text-align: left; }}
            th {{ background-color: #f8f9fa; color: {colore_tema}; font-weight: bold; text-transform: uppercase; font-size: 12px; }}
            tr:nth-child(even) {{ background-color: #fcfcfc; }}
            .signatures {{ display: flex; justify-content: space-between; margin-top: 60px; }}
            .sig-box {{ width: 45%; border-top: 1px solid #000; padding-top: 8px; text-align: center; font-size: 14px; font-weight: bold; }}
            @media print {{
                body {{ margin: 0; padding: 20px; }}
                .header-container {{ border-bottom: 2px solid #000; }}
                th {{ background-color: #eee !important; color: #000 !important; -webkit-print-color-adjust: exact; }}
            }}
        </style>
    </head>
    <body onload="window.print()">
        <div class="header-container">
            <div>
                {tag_logo}
                <h1 class="doc-title">{titolo_doc}</h1>
                <h2 class="doc-subtitle">Cantiere / Assegnatario: <strong>{cantiere.upper()}</strong></h2>
            </div>
            <div class="doc-meta">
                <div><strong>Data Generazione:</strong> {data_oggi}</div>
                <div><strong>Totale Strumenti:</strong> {len(df_kit)}</div>
            </div>
        </div>
        <table>
            <thead>
                <tr>
                    <th style="width: 25%;">Codice RFID</th>
                    <th style="width: 35%;">Strumento</th>
                    <th style="width: 25%;">Modello / Tipologia</th>
                    <th style="width: 15%;">Note / Stato</th>
                </tr>
            </thead>
            <tbody>{righe}</tbody>
        </table>
        <div class="signatures">
            <div class="sig-box">Firma Operatore / Ricevente</div>
            <div class="sig-box">Firma Responsabile Magazzino</div>
        </div>
    </body>
    </html>
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
    st.markdown("### 📡 Scansione di Massa (RFID)")
    st.info("💡 **Istruzioni:** Clicca nel riquadro, tieni premuto il grilletto della Zebra per leggere tutti i tag dell'area. Quando hai finito, **clicca fuori dal riquadro**.")
    
    testo_scansioni = st.text_area("Tag Scansionati:", height=150, placeholder="Spara i codici a raffica qui...")
    components.html("""<script>var i=window.parent.document.querySelector("textarea[aria-label='Tag Scansionati:']"); if(i)i.focus();</script>""", height=0)
    
    if testo_scansioni:
        codici_letti = list(set([str(c).strip() for c in testo_scansioni.split('\n') if str(c).strip()]))
        
        if codici_letti:
            df_magazzino['Codice_Pulito'] = df_magazzino['CODICE RFID'].astype(str).str.strip()
            df_trovati = df_magazzino[df_magazzino['Codice_Pulito'].isin(codici_letti)].copy()
            
            codici_non_trovati = set(codici_letti) - set(df_trovati['Codice_Pulito'].tolist())
            if codici_non_trovati:
                st.warning(f"⚠️ {len(codici_non_trovati)} tag letti non sono registrati nel gestionale.")
            
            if not df_trovati.empty:
                st.write(f"🔍 Trovati **{len(df_trovati)}** strumenti validi. Deseleziona eventuali errori:")
                
                df_trovati.insert(0, 'Seleziona', True)
                df_editato = st.data_editor(
                    df_trovati[['Seleziona', 'Stato', 'STRUMENTAZIONE', 'MODELLO/TIPOLOGIA', 'Assegnatario / Cantiere']],
                    column_config={"Seleziona": st.column_config.CheckboxColumn("Conferma", default=True)},
                    disabled=['Stato', 'STRUMENTAZIONE', 'MODELLO/TIPOLOGIA', 'Assegnatario / Cantiere'],
                    hide_index=True, use_container_width=True
                )
                
                idx_selezionati = df_editato[df_editato['Seleziona'] == True].index
                
                in_sede = df_trovati.loc[idx_selezionati][df_trovati.loc[idx_selezionati, 'Stato'] == "🟢 In Sede"]
                fuori = df_trovati.loc[idx_selezionati][df_trovati.loc[idx_selezionati, 'Stato'] != "🟢 In Sede"]
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown(f"#### 🔴 Invia in Cantiere ({len(in_sede)})")
                    if not in_sede.empty:
                        scaduti = []
                        for idx in in_sede.index:
                            if verifica_taratura_iso(in_sede.loc[idx, 'MATRICOLA']) == "SCADUTO":
                                scaduti.append(in_sede.loc[idx, 'STRUMENTAZIONE'])
                        if scaduti:
                            st.error(f"⚠️ Taratura SCADUTA per: {', '.join(scaduti)}")
                        
                        assegnatario = st.text_input("Assegna tutti a:", value=st.session_state['ultimo_cantiere'])
                        bottone_testo = "⚠️ FORZA USCITA GRUPPO" if scaduti else "🔴 Registra USCITA di massa"
                        bottone_tipo = "secondary" if scaduti else "primary"
                        
                        if st.button(bottone_testo, type=bottone_tipo, key="btn_out_mass"):
                            if assegnatario:
                                st.session_state['ultimo_cantiere'] = assegnatario
                                for idx in in_sede.index:
                                    strumento = df_magazzino.loc[idx]
                                    df_magazzino.at[idx, 'Stato'] = "🔴 In Cantiere"
                                    df_magazzino.at[idx, 'Assegnatario / Cantiere'] = assegnatario
                                    df_magazzino.at[idx, 'Ultima Modifica'] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
                                    nota = "🔴 USCITA (FORZATA)" if strumento['STRUMENTAZIONE'] in scaduti else "🔴 USCITA"
                                    salva_in_storico(strumento['CODICE RFID'], strumento['STRUMENTAZIONE'], strumento['MODELLO/TIPOLOGIA'], nota, assegnatario)
                                salva_live(df_magazzino)
                                st.success(f"✅ {len(in_sede)} strumenti registrati in uscita!")
                                st.cache_data.clear()
                                time.sleep(1.5)
                                st.rerun()
                            else:
                                st.warning("Inserisci il cantiere di destinazione!")
                    else:
                        st.info("Nessuno strumento selezionato per l'uscita.")
                        
                with col2:
                    st.markdown(f"#### 🟢 Fai Rientrare ({len(fuori)})")
                    if not fuori.empty:
                        if st.button("🟢 Registra RIENTRO di massa", type="primary", key="btn_in_mass"):
                            for idx in fuori.index:
                                strumento = df_magazzino.loc[idx]
                                vecchio_ass = strumento['Assegnatario / Cantiere']
                                df_magazzino.at[idx, 'Stato'] = "🟢 In Sede"
                                df_magazzino.at[idx, 'Assegnatario / Cantiere'] = ""
                                df_magazzino.at[idx, 'Ultima Modifica'] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
                                salva_in_storico(strumento['CODICE RFID'], strumento['STRUMENTAZIONE'], strumento['MODELLO/TIPOLOGIA'], "🟢 RIENTRO", vecchio_ass)
                            salva_live(df_magazzino)
                            st.success(f"✅ {len(fuori)} strumenti rientrati in sede!")
                            st.cache_data.clear()
                            time.sleep(1.5)
                            st.rerun()
                    else:
                        st.info("Nessuno strumento selezionato per il rientro.")

elif modulo == "📝 Uscita/Rientro Manuale":
    st.markdown("### 📝 Selezione Multipla per Reparto")
    operazione = st.radio("Seleziona operazione:", ["🔴 Uscita verso Cantiere", "🟢 Rientro in Sede"], horizontal=True)
    
    if "Uscita" in operazione:
        disponibili = df_magazzino[df_magazzino['Stato'] == "🟢 In Sede"].copy()
        if disponibili.empty:
            st.info("Nessuno strumento in sede.")
        else:
            categorie = sorted([c for c in disponibili['Categoria_Foglio'].unique() if str(c).strip() != "nan" and str(c).strip() != ""])
            if not categorie: categorie = ["Generale"]
            
            tabs = st.tabs([str(c).upper() for c in categorie])
            tutti_selezionati = []
            
            for i, cat in enumerate(categorie):
                with tabs[i]:
                    df_cat = disponibili[disponibili['Categoria_Foglio'] == cat].copy()
                    df_cat.insert(0, 'Seleziona', False)
                    df_editato = st.data_editor(
                        df_cat[['Seleziona', 'STRUMENTAZIONE', 'MODELLO/TIPOLOGIA', 'MATRICOLA']],
                        column_config={"Seleziona": st.column_config.CheckboxColumn("Spunta", default=False)},
                        disabled=['STRUMENTAZIONE', 'MODELLO/TIPOLOGIA', 'MATRICOLA'],
                        hide_index=True, use_container_width=True, key=f"uscita_{cat}"
                    )
                    selezionati_cat = df_editato[df_editato['Seleziona'] == True]
                    if not selezionati_cat.empty:
                        tutti_selezionati.append(selezionati_cat)

            st.divider()
            assegnatario = st.text_input("Assegna a Cantiere/Persona:", value=st.session_state['ultimo_cantiere'])
            
            if st.button("🔴 Registra Uscita Selezionati", type="primary"):
                if tutti_selezionati and assegnatario:
                    st.session_state['ultimo_cantiere'] = assegnatario
                    df_selezionati_completo = pd.concat(tutti_selezionati)
                    for idx in df_selezionati_completo.index:
                        strumento = df_magazzino.loc[idx]
                        df_magazzino.at[idx, 'Stato'] = "🔴 In Cantiere"
                        df_magazzino.at[idx, 'Assegnatario / Cantiere'] = assegnatario
                        df_magazzino.at[idx, 'Ultima Modifica'] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
                        salva_in_storico(strumento['CODICE RFID'], strumento['STRUMENTAZIONE'], strumento['MODELLO/TIPOLOGIA'], "🔴 USCITA (MANUALE)", assegnatario)
                    salva_live(df_magazzino)
                    st.success(f"✅ {len(df_selezionati_completo)} strumenti in uscita verso {assegnatario}!")
                    st.cache_data.clear()
                    time.sleep(1)
                    st.rerun()
                else:
                    st.warning("Spunta almeno uno strumento e inserisci il cantiere.")
    else:
        fuori = df_magazzino[df_magazzino['Stato'] != "🟢 In Sede"].copy()
        if fuori.empty:
            st.info("Nessuno strumento fuori sede.")
        else:
            categorie = sorted([c for c in fuori['Categoria_Foglio'].unique() if str(c).strip() != "nan" and str(c).strip() != ""])
            if not categorie: categorie = ["Generale"]
            
            tabs = st.tabs([str(c).upper() for c in categorie])
            tutti_selezionati = []
            
            for i, cat in enumerate(categorie):
                with tabs[i]:
                    df_cat = fuori[fuori['Categoria_Foglio'] == cat].copy()
                    df_cat.insert(0, 'Seleziona', False)
                    df_editato = st.data_editor(
                        df_cat[['Seleziona', 'STRUMENTAZIONE', 'MODELLO/TIPOLOGIA', 'Assegnatario / Cantiere']],
                        column_config={"Seleziona": st.column_config.CheckboxColumn("Spunta", default=False)},
                        disabled=['STRUMENTAZIONE', 'MODELLO/TIPOLOGIA', 'Assegnatario / Cantiere'],
                        hide_index=True, use_container_width=True, key=f"rientro_{cat}"
                    )
                    selezionati_cat = df_editato[df_editato['Seleziona'] == True]
                    if not selezionati_cat.empty:
                        tutti_selezionati.append(selezionati_cat)
            
            st.divider()
            if st.button("🟢 Registra Rientro Selezionati", type="primary"):
                if tutti_selezionati:
                    df_selezionati_completo = pd.concat(tutti_selezionati)
                    for idx in df_selezionati_completo.index:
                        strumento = df_magazzino.loc[idx]
                        vecchio_cantiere = strumento['Assegnatario / Cantiere']
                        df_magazzino.at[idx, 'Stato'] = "🟢 In Sede"
                        df_magazzino.at[idx, 'Assegnatario / Cantiere'] = ""
                        df_magazzino.at[idx, 'Ultima Modifica'] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
                        salva_in_storico(strumento['CODICE RFID'], strumento['STRUMENTAZIONE'], strumento['MODELLO/TIPOLOGIA'], "🟢 RIENTRO (MANUALE)", vecchio_cantiere)
                    salva_live(df_magazzino)
                    st.success(f"✅ {len(df_selezionati_completo)} strumenti rientrati!")
                    st.cache_data.clear()
                    time.sleep(1)
                    st.rerun()
                else:
                    st.warning("Spunta almeno uno strumento.")

elif modulo == "📊 Situazione Attuale":
    st.markdown("### 📦 Inventario e Giacenze")
    filtro = st.radio("Filtra:", ["Tutti", "🟢 In Sede", "🔴 In Cantiere"], horizontal=True)
    df_v = df_magazzino.copy() if filtro == "Tutti" else df_magazzino[df_magazzino['Stato'] == filtro]
    st.metric("Totale Strumenti", len(df_v))
    st.dataframe(df_v[['Stato', 'STRUMENTAZIONE', 'MODELLO/TIPOLOGIA', 'Assegnatario / Cantiere']], use_container_width=True, hide_index=True)

elif modulo == "🏗️ Cantieri e Bolle":
    st.markdown("### 🏗️ Gestione Distinte")
    df_st = carica_storico()
    cantieri = sorted([str(c) for c in df_st['Cantiere / Assegnatario'].unique() if str(c).strip() != ""])
    
    if cantieri:
        cant = st.selectbox("Scegli Cantiere:", ["-- Seleziona --"] + cantieri)
        if cant != "-- Seleziona --":
            tipo = st.radio("Fase:", ["🔴 Uscita", "🟢 Rientro"], horizontal=True)
            filtro = "USCITA" if "Uscita" in tipo else "RIENTRO"
            
            maschera = (df_st['Cantiere / Assegnatario'] == cant) & (df_st['Operazione'].str.contains(filtro))
            df_kit = df_st[maschera].drop_duplicates(subset=['CODICE RFID'])
            
            if not df_kit.empty:
                st.markdown("Modifica la bolla (puoi cancellare righe con l'icona cestino prima di stampare):")
                df_kit_edit = st.data_editor(
                    df_kit[['CODICE RFID', 'Strumento', 'Modello', 'Data']], 
                    use_container_width=True, hide_index=True, num_rows="dynamic"
                )
                
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button("🖨️ Stampa Bolla", data=genera_bolla_html(cant, df_kit_edit, filtro), file_name=f"Bolla_{cant}.html", mime="text/html", type="primary")
                with col2:
                    if st.button("🗑️ Elimina intera Bolla", type="secondary"):
                        df_st_pulito = df_st[~maschera]
                        with pd.ExcelWriter(FILE_STORICO, engine="openpyxl", mode="w") as w:
                            df_st_pulito.to_excel(w, index=False)
                        st.success("✅ Bolla eliminata!")
                        time.sleep(1)
                        st.rerun()
            else:
                st.info("Nessuna movimentazione trovata.")

elif modulo == "📜 Storico Movimenti":
    st.markdown("### 📜 Movimentazioni (Cancella Riga o Modifica)")
    df_st = carica_storico()
    df_edit = st.data_editor(df_st.iloc[::-1].copy(), use_container_width=True, hide_index=True, num_rows="dynamic")
    if st.button("💾 Salva Modifiche", type="secondary"):
        with pd.ExcelWriter(FILE_STORICO, engine="openpyxl", mode="w") as w:
            df_edit.iloc[::-1].to_excel(w, index=False)
        st.success("✅ Aggiornato!")
        st.rerun()