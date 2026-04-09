"""
BAP Processor - Consolidamento automatico inventari + export SAP
Progetti supportati: ECO, Sirius, DCT300
Uso:
    python bap_processor.py
    -> legge i file inventario e genera dashboard.html + dati_bap.json
"""

import pandas as pd
import numpy as np
import json
import os
import warnings
from datetime import datetime

# Sopprime solo i warning rumorosi di openpyxl (stili, date, ecc.)
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Modalità DEBUG per verificare le celle lette (imposta a True per i dettagli)
DEBUG = False

# ─────────────────────────────────────────────
# CONFIGURAZIONE PERCORSI FILE
# ─────────────────────────────────────────────
CONFIG = {
    "master_data":       "bap_master.json",
    "inventory_baseline": "bap_inventory_baseline.json",
    "sap_zpp093":        "CONFERMESAP.XLSX",
    "sap_mb51":          "mb51.xlsx",
    "sap_mapping":       "BAP1.xlsx",
    "output_dashboard":  "dashboard.html",
    "output_json":       "dati_bap.json",
}

FAMILY_NAMES = ['SG1','SG2','SG3','SG4','SG5','SG6','SG7','SG8',
                'DG2','SGR','SGRW','RG','FG5','FG57','PIGNON','PG',
                'SG 1','SG 2','SG 3','SG 4','SG 5','SG 6','SG 7','SG 8']

# Mappe SAP globali (popolate all'avvio)
SAP_MATERIALS_MAP = {} # { 'hard_code': { 'soft': '...', 'inter': '...', 'hard': '...' } }
SAP_STATIONS_MAP  = {} # { 'inv_station_name': ['op1', 'op2'] }

def load_sap_mapping(bap1_base_path):
    global SAP_MATERIALS_MAP, SAP_STATIONS_MAP
    PERMANENT_CACHE = 'bap_mapping_permanent.json'
    
    # 1. Carica prima la cache permanente (JSON) se esiste - QUESTA HA LA PRIORITÀ
    if os.path.exists(PERMANENT_CACHE):
        try:
            with open(PERMANENT_CACHE, 'r', encoding='utf-8') as f:
                cache = json.load(f)
                SAP_MATERIALS_MAP = cache.get("materiali", {})
                SAP_STATIONS_MAP  = cache.get("stazioni", {})
            print(f"  -> Caricate {len(SAP_MATERIALS_MAP)} mappature materiali e {len(SAP_STATIONS_MAP)} stazioni da JSON (Priorità).")
        except Exception as e:
            print(f"  [ERRORE] Caricamento cache permanente fallito: {e}")

    # 2. Trova e integra dal file Excel (BAP1)
    excel_path = None
    for ext in ['.xlsx', '.xls', '.XLSX', '.XLS']:
        p = bap1_base_path.replace('.xlsx', '').replace('.XLSX', '') + ext
        if os.path.exists(p):
            excel_path = p
            break

    if excel_path:
        # Caricamento flessibile del file Excel
        df = _read_excel_flexible(excel_path)
        
        if df is not None:
            try:
                print(f"  -> Integrazione mappatura da Excel: {excel_path}")
                
                # Se è un DataFrame .xlsx, forziamo lo sheet 'nuovo flusso' se presente
                if excel_path.lower().endswith('.xlsx'):
                    try:
                        df_spec = pd.read_excel(excel_path, sheet_name='nuovo flusso', header=None)
                        df = df_spec
                    except:
                        pass # Se lo sheet non esiste, usiamo quello letto in precedenza
                else:
                    # Per .xls non-standard, azzeriamo l'header se necessario per mappare indici fissi
                    df.columns = range(df.shape[1])

                # Mappa Stazioni (Righe 1, 2, 3 - indici 0-based: 1, 2, 3)
                for c in range(4, df.shape[1]):
                    st_val = df.iloc[3, c]
                    st_name = str(st_val).strip() if pd.notna(st_val) else ""
                    
                    if st_name and st_name != 'nan' and st_name != 'None':
                        # NON sovrascrivere se già presente nel JSON (manual override)
                        if st_name in SAP_STATIONS_MAP: continue
                        
                        ops = []
                        def _clean_op(val):
                            if pd.isna(val): return None
                            s = str(val).strip()
                            if not s or s.lower() == 'nan': return None
                            if s.endswith('.0'): s = s[:-2]
                            return s

                        op1 = _clean_op(df.iloc[1, c]); op2 = _clean_op(df.iloc[2, c])
                        if op1: ops.append(op1)
                        if op2: ops.append(op2)
                        if ops: SAP_STATIONS_MAP[st_name] = ops

                # Mappa Materiali (da riga 5 in poi)
                mats_df = df.iloc[4:, [0,1,2]].dropna(how='all')
                for _, row in mats_df.iterrows():
                    soft  = str(row[0]).strip() if pd.notna(row[0]) else ""
                    inter = str(row[1]).strip() if pd.notna(row[1]) else ""
                    hard  = str(row[2]).strip() if pd.notna(row[2]) else ""
                    if hard:
                        key = str(hard).strip().upper()
                        if key not in SAP_MATERIALS_MAP: # Non sovrascrivere
                            SAP_MATERIALS_MAP[key] = {"soft": soft, "inter": inter, "hard": hard}
                
                # Aggiorna la cache con i dati integrati
                with open(PERMANENT_CACHE, 'w', encoding='utf-8') as f:
                    json.dump({"materiali": SAP_MATERIALS_MAP, "stazioni": SAP_STATIONS_MAP}, f, ensure_ascii=False, indent=2)
                
            except Exception as e:
                print(f"  [ERRORE] Integrazione Excel fallita: {e}")

    if not SAP_MATERIALS_MAP:
        print(f"  [ATTENZIONE] Nessuna mappatura SAP caricata! Mappatura mancante sia in Excel che in JSON.")
    else:
        print(f"  -> Totale mappature attive: {len(SAP_MATERIALS_MAP)} materiali, {len(SAP_STATIONS_MAP)} stazioni.")

def get_sap_data_for_comp(comp_sap_codes, st_name, sap_zpp):
    """
    Ritorna i dati SAP per un componente e una stazione specifica.
    Utilizza una mappatura euristica di fallback se non definita nel JSON.
    """
    ops = SAP_STATIONS_MAP.get(st_name, [])
    
    # Heuristic mapping if empty (common operations based on pnumb.xlsx)
    if not ops:
        n = st_name.upper()
        if 'DENTATUR' in n or 'PFAUTER' in n or 'HOBBING' in n or 'SHAPING' in n or 'FRW' in n: ops = ["90"]
        elif 'TORNITUR' in n or 'EMAG' in n or 'DRA' in n or 'MZA' in n or 'UT ' in n: ops = ["10", "20", "60"]
        elif 'SCA' in n or 'LASER' in n: ops = ["60", "151"]
        elif 'RETTIFIC' in n or 'GRINDING' in n or 'RZ' in n or 'SLA' in n or 'LAPPAT' in n: ops = ["101", "119", "120", "230"]
        elif 'WASH' in n or 'LAVAGGIO' in n or 'LAVARE' in n or 'FINALE' in n: ops = ["220", "250"]
        elif 'SBV' in n or 'SBAVATUR' in n or 'DEBURRING' in n: ops = ["90"]
        elif 'TERMICO' in n or 'HT ' in n or 'TT' in n or 'CEMENT' in n: ops = ["100", "240"]

    if not ops or not sap_zpp:
        return {"qty": 0, "ops": ops}
    
    codes = [v for v in comp_sap_codes.values() if v]
    total_sap_wip = 0
    
    for code in codes:
        code_upper = str(code).upper()
        # Se il codice ha già un suffisso (es. /S), agiamo sul base per sicurezza o cerchiamo esatto
        base_code = code_upper.split('/')[0]
        
        def norm_pad(s):
            if str(s).startswith('M01'): return 'M1' + s[3:]
            return s
        
        norm_base = norm_pad(base_code)
        
        # Cerchiamo tutti i materiali in SAP che iniziano con il base_code
        # (es. M0162645, M0162645/S, M0162645/T)
        for sap_mat in sap_zpp.keys():
            s_base = sap_mat.split('/')[0].strip()
            if s_base == base_code or norm_pad(s_base) == norm_base:
                # Somma WIP per tutte le operazioni SAP mappate a questa stazione
                for op in ops:
                    if op in sap_zpp[sap_mat]:
                        data = sap_zpp[sap_mat][op]
                        # Sommiamo sia il WIP (rimanente) che l'Ottenuta (confermata) 
                        # per riflettere i 'pezzi' che l'utente vede in SAP
                        total_sap_wip += (data.get('wip', 0) + data.get('ottenuta', 0))
                    
    return {"qty": int(total_sap_wip), "ops": ops}


# ═══════════════════════════════════════════════
# PARSER ECO
# ═══════════════════════════════════════════════
# ═══════════════════════════════════════════════
# CARICAMENTO BASELINE (JSON)
# ═══════════════════════════════════════════════
def load_baseline_data(master_path, baseline_path):
    """
    Carica l'anagrafica e le quantità fisiche di partenza dai file JSON.
    Sostituisce la vecchia estrazione dai file Excel.
    """
    if not os.path.exists(master_path) or not os.path.exists(baseline_path):
        print(f"  [ERRORE] File baseline non trovati: {master_path} o {baseline_path}")
        return []

    try:
        with open(master_path, 'r', encoding='utf-8') as f:
            master = json.load(f)
        with open(baseline_path, 'r', encoding='utf-8') as f:
            baseline = json.load(f)
            
        components = []
        for m in master:
            key = f"{m['progetto']}||{m['label']}"
            base = baseline.get(key, {})
            
            # Ricostruiamo la struttura del componente
            c = {
                "progetto":    m.get('progetto'),
                "famiglia":    m.get('famiglia'),
                "label":       m.get('label'),
                "codice_hard": m.get('codice_hard'),
                "codice_s":    m.get('codice_s'),
                "demand_fd1":  m.get('demand_fd1', 0),
                "finiti":      base.get('finiti', 0),
                "data_inv":    base.get('data_inv', '2026-03-27'),
                "stazioni":    {}
            }
            
            # Popoliamo le stazioni con le quantità dalla baseline
            st_baseline = base.get('stazioni', {})
            for st in m.get('stazioni_list', []):
                c['stazioni'][st] = {
                    "qty": _to_int(st_baseline.get(st, 0)),
                    "sap_ops": SAP_STATIONS_MAP.get(st, [])
                }
            components.append(c)
            
        print(f"  -> Caricati {len(components)} componenti dalla baseline JSON.")
        return components
    except Exception as e:
        print(f"  [ERRORE] Caricamento baseline fallito: {e}")
        return []


# ═══════════════════════════════════════════════
# PARSER SAP ZPP_093
# ═══════════════════════════════════════════════
def parse_sap_zpp093(filepath):
    if not os.path.exists(filepath):
        return {}
    try:
        df = _read_excel_flexible(filepath)
        if df is None: return {}
        print(f"  SAP zpp093: {len(df)} righe")
        result = {}
        mat_col  = _find_col(df, ['Materiale','Material','MATNR'])
        qty_col  = _find_col(df, ['Qtà totale','Qty totale','Quantity','GAMNG','Mg.ds'])
        cons_col = _find_col(df, ['Qtà consegnata','Delivered','WEMNG','Quantità ottenuta'])
        op_col   = _find_col(df, ['Fase','Operazione','Phase','Operation','VORNR','Fino'])
        sosp_col = _find_col(df, ['SOSPESI'])
        
        if mat_col:
            for _, row in df.iterrows():
                mat = str(row[mat_col]).strip().upper()
                qty = _to_num(row[qty_col]) if qty_col else 0
                cons = _to_num(row[cons_col]) if cons_col else 0
                sosp = _to_num(row[sosp_col]) if sosp_col else 0
                op_raw = str(row[op_col]).strip() if op_col else ""
                
                # WIP per Overview (priorità SOSPESI)
                if sosp_col and sosp > 0:
                    wip = sosp
                else:
                    wip = max(0, qty - cons)

                # Quantità ottenuta/consegnata (per P-NUM Flow)
                ottenuta = cons

                # Normalizza operazione
                try:
                    op = str(int(float(op_raw))) if op_raw and op_raw != 'nan' else op_raw
                except (ValueError, TypeError):
                    op = op_raw

                if mat and mat != 'nan':
                    # Creiamo anche una versione 'base' del codice materiale (senza /S, /T)
                    # per i match dell'Overview
                    base_mat = mat.split('/')[0].strip()
                    
                    if mat not in result: result[mat] = {}
                    if op not in result[mat]:
                        result[mat][op] = {'qty_totale': 0, 'qty_consegnata': 0, 'wip': 0, 'ordini': 0}
                    
                    data = result[mat][op]
                    data['qty_totale']     += qty
                    data['qty_consegnata'] += cons
                    data['ottenuta']       = data['qty_consegnata'] # Alias per chiarezza
                    data['wip']            += wip
                    data['ordini']         += 1
        return result
    except Exception as e:
        print(f"  [ERRORE] Lettura zpp093: {e}")
        return {}


# ═══════════════════════════════════════════════
# PARSER SAP MB51
# ═══════════════════════════════════════════════
def parse_sap_mb51(filepath):
    if not os.path.exists(filepath):
        return {}
    try:
        df = _read_excel_flexible(filepath)
        if df is None: return {}
        print(f"  SAP mb51: {len(df)} righe")
        result = {}
        mat_col  = _find_col(df, ['Materiale', 'Material', 'MATNR'])
        qty_col  = _find_col(df, ['Quantità', 'Quantity', 'MENGE'])
        mvt_col  = _find_col(df, ['Tipo mov.', 'Movement', 'BWART'])
        date_col = _find_col(df, ['Data reg', 'Posting Date', 'BUDAT', 'Data doc'])
        
        if mat_col and qty_col:
            for _, row in df.iterrows():
                mat = str(row[mat_col]).strip().upper()
                qty = abs(_to_num(row[qty_col]))
                mvt = str(row[mvt_col]).strip() if mvt_col else ''
                
                # Parsing data per filtraggio temporale
                m_date = None
                if date_col:
                    try:
                        m_date = pd.to_datetime(row[date_col])
                    except:
                        pass

                if mat and mat != 'nan':
                    if mat not in result:
                        result[mat] = []
                    
                    # Salviamo i movimenti come lista per poter filtrare per data in calcola_metriche
                    result[mat].append({
                        'qty': qty,
                        'mvt': mvt,
                        'date': m_date,
                        'tipo': 'entrate' if mvt.startswith('1') else ('uscite' if (mvt.startswith('2') or mvt.startswith('5')) else 'altro')
                    })
        return result
    except Exception as e:
        print(f"  [ERRORE] Lettura mb51: {e}")
        return {}
# ═══════════════════════════════════════════════
# PARSER P-NUM MATRIX (pnumb.xlsx)
# ═══════════════════════════════════════════════
def load_pnum_matrix(filepath, sap_zpp):
    """
    Legge pnumb.xlsx e crea una matrice sincronizzata con SAP.
    Ogni riga rappresenta un Part Number con le sue fasi.
    """
    if not os.path.exists(filepath):
        return []
    
    try:
        # Leggiamo senza header per gestire le righe di intestazione multiple (0 e 1)
        df = pd.read_excel(filepath, header=None)
        
        # Saltiamo le prime 2 righe (intestazioni)
        data_rows = df.iloc[2:].values
        matrix = []
        
        for row in data_rows:
            # Estrarre le info base (colonne 0, 1, 2, 3)
            # 0: SI/NO, 1: Progetto, 2: Famiglia, 3: P.Number
            item = {
                "si_no":    str(row[0]) if pd.notna(row[0]) else "",
                "progetto": str(row[1]) if pd.notna(row[1]) else "",
                "famiglia": str(row[2]) if pd.notna(row[2]) else "",
                "p_number": str(row[3]) if pd.notna(row[3]) else "",
                "fasi": []
            }
            
            # Fasi (Ogni fase ha 4 colonne + 1 spaziometrica, partendo dalla colonna 5)
            # Indici: 5,6,7,8 | 10,11,12,13 | 15,16,17,18 | 20,21,22,23 | 25,26,27,28 | 30,31,32,33
            phase_starts = [5, 10, 15, 20, 25, 30]
            
            # Funzione di normalizzazione padding (M01 <=> M1)
            def norm_pad(s):
                s = str(s).strip().upper()
                if s.startswith('M01'): return 'M1' + s[3:]
                return s

            for start in phase_starts:
                if start >= len(row): break
                
                pn   = str(row[start]).strip().upper() if pd.notna(row[start]) else ""
                op   = str(row[start+1]).strip() if pd.notna(row[start+1]) else ""
                name = str(row[start+2]).strip() if pd.notna(row[start+2]) else ""
                
                # Normalizza operazione per SAP
                try:
                    op_sap = str(int(float(op))) if op and op != 'nan' else op
                except:
                    op_sap = op

                # Recupero dati da SAP (con supporto flessibile a suffissi e padding)
                valore_fase = 0
                if pn and op_sap and sap_zpp:
                    base_pn = pn.split('/')[0].strip()
                    norm_base = norm_pad(base_pn)
                    
                    for s_mat in sap_zpp.keys():
                        s_base = s_mat.split('/')[0].strip()
                        if s_base == base_pn or norm_pad(s_base) == norm_base:
                            if op_sap in sap_zpp[s_mat]:
                                # Mostriamo la QUANTITÀ OTTENUTA (confermata)
                                valore_fase += sap_zpp[s_mat][op_sap].get('ottenuta', 0)
                
                item["fasi"].append({
                    "pn": pn,
                    "op": op,
                    "name": name,
                    "wip": int(valore_fase) # 'wip' nel JSON ma contiene la qta ottenuta
                })
            
            matrix.append(item)
            
        print(f"  -> Matrice P-NUM caricata: {len(matrix)} righe")
        return matrix
    except Exception as e:
        print(f"  [ERRORE] Lettura pnumb.xlsx: {e}")
        return []



# ═══════════════════════════════════════════════
# CALCOLO METRICHE
# ═══════════════════════════════════════════════
def calcola_metriche(components, sap_zpp=None, sap_mb51=None):
    sap_zpp  = sap_zpp  or {}
    sap_mb51 = sap_mb51 or {}

    for c in components:
        # Recupera codici SAP associati (soft, inter, hard)
        hard_key = c.get('codice_hard', '').strip().upper()
        sap_codes = SAP_MATERIALS_MAP.get(hard_key, {})
        
        # Se non trovato in mappa, prova con codice_s pulito (per DCT300/Sirius)
        fallback = c.get('codice_s', '').replace('/S','').replace('/s','').strip().upper()
        if not sap_codes:
            real_sap_codes = SAP_MATERIALS_MAP.get(fallback)
            if real_sap_codes:
                sap_codes = real_sap_codes
            else:
                sap_codes = { "hard": fallback }

        c['sap_codes'] = sap_codes
        
        # Foolproof backfill
        if c.get('codice_s'):
            c['codice_s'] = str(c['codice_s']).upper()
        if c.get('codice_hard'):
            c['codice_hard'] = str(c['codice_hard']).upper()
            
        if not c.get('codice_hard') or str(c['codice_hard']).strip() == "":
            if fallback in SAP_MATERIALS_MAP:
                c['codice_hard'] = str(SAP_MATERIALS_MAP[fallback].get('hard', '')).upper()
        if not c.get('codice_s') or str(c['codice_s']).strip() == "":
            hard_upper = c.get('codice_hard', '').strip().upper()
            if hard_upper in SAP_MATERIALS_MAP:
                c['codice_s'] = str(SAP_MATERIALS_MAP[hard_upper].get('soft', '')).upper()
        
        # Inizializza KPI SAP
        c['sap_ordini'] = 0
        c['sap_qty_aperta'] = 0
        c['sap_entrate'] = 0
        c['sap_uscite'] = 0

        # ─── ATTUALIZZAZIONE INVENTARIO (WIP & FINITI) ───
        inv_date_str = c.get('data_inv', 'N/D')
        try:
            target_date = pd.to_datetime(inv_date_str) if inv_date_str != 'N/D' else None
        except:
            target_date = None

        # 1. Aggrega dati SAP per Finiti (MB51)
        for role, code in sap_codes.items():
            if not code: continue
            code_upper = str(code).upper()
            if code_upper in sap_mb51:
                movements = sap_mb51[code_upper]
                for m in movements:
                    if target_date and m['date'] and m['date'] <= target_date:
                        continue
                    if m['tipo'] == 'entrate':
                        c['sap_entrate'] += m['qty']
                        c['finiti'] += m['qty']
                    elif m['tipo'] == 'uscite':
                        c['sap_uscite'] += m['qty']
                        c['finiti'] = max(0, c['finiti'] - m['qty'])

        # 2. Logica Avanzata WIP (ZPP_093) con Movimento tra Tecnologie
        # SOFT: Tornitura, Dentatura, Lavaggio, Rasatura, ecc.
        # TERMICO: Trattamento Termico, TT, IND, Cementazione, Tempra, Bonifica, Pallinatura, Brunitura
        # HARD: Rettifica, BAA, Lappatura, Finale, Collaudo, Spedizione
        TECH_ORDER = {'SOFT': 0, 'TERMICO': 1, 'HARD': 2, 'FINALE': 3, 'ALTRO': 99}
        
        def get_tech(st):
            s = st.lower()
            if 'spedizion' in s or 'magazzino' in s or 'fine' in s: return 'FINALE'
            if 'hard' in s or 'rettific' in s or 'baa' in s or 'lappatur' in s: return 'HARD'
            if 'termico' in s or 'trattamento' in s or 'pallinatur' in s or 'ind' in s or 'cementazion' in s or 'tempra' in s or 'brunitur' in s: return 'TERMICO'
            if 'soft' in s or 'tornitur' in s or 'lavaggio' in s or 'dentatur' in s or 'saldatur' in s or 'rasatur' in s or 'laser' in s: return 'SOFT'
            return 'ALTRO'

        # Identifichiamo la tecnologia SAP più avanzata in cui c'è materiale
        max_sap_tech_idx = -1
        # Prepariamo i dati SAP per ogni stazione per evitare ricalcoli
        st_sap_info = {}
        for st in c.get('stazioni', {}).keys():
            info = get_sap_data_for_comp(sap_codes, st, sap_zpp)
            st_sap_info[st] = info
            if info['qty'] > 0:
                idx = TECH_ORDER.get(get_tech(st), 99)
                if idx < 99 and idx > max_sap_tech_idx:
                    max_sap_tech_idx = idx

        # Aggiornamento stazioni con logica di movimento
        total_wip_updated = 0
        new_stazioni = {}
        for st, data in c.get('stazioni', {}).items():
            info = st_sap_info[st]
            tech = get_tech(st)
            tech_idx = TECH_ORDER.get(tech, 99)
            
            # Valore di base (fisico) - Forza conversione in dict se numero
            if not isinstance(data, dict):
                data = {'qty': _to_int(data)}
            
            physical_qty = _to_int(data.get('qty', 0))
            
            # Logica WIP Sommata (Inventario + SAP)
            actual_qty = physical_qty + info['qty']
            orig = "Somma (Inv+SAP)" if info['qty'] > 0 else "Manuale"
            
            # Logica di Sincronizzazione (Avanzamento):
            # Se SAP dice che siamo in una fase successiva (es. HARD), 
            # azzeriamo le somme delle fasi precedenti (es. SOFT)
            if tech_idx < max_sap_tech_idx and tech_idx != 99:
                actual_qty = 0
                orig = "Sync (Avanzamento)"

            data['qty'] = actual_qty
            data['sap_qty'] = info['qty']
            data['sap_ops'] = info['ops']
            data['origin']  = orig # Per debug / visibilità
            data['tech']    = tech
            
            new_stazioni[st] = data
            total_wip_updated += actual_qty

        c['stazioni'] = new_stazioni
        c['wip_totale'] = total_wip_updated
        c['tot_wip'] = total_wip_updated

        # 3. Aggrega Ordini/Aperto SAP (per info generale)
        for role, code in sap_codes.items():
            if not code: continue
            code_upper = str(code).upper()
            if code_upper in sap_zpp:
                for op, data in sap_zpp[code_upper].items():
                    c['sap_ordini']     += data.get('ordini', 0)
                    c['sap_qty_aperta'] += (data.get('qty_totale', 0) - data.get('qty_consegnata', 0))

        # Calcolo Dinamico Copertura
        target = c.get('demand_fd1', 0)
        if target > 0:
            c['gg_copertura_finiti'] = c.get('finiti', 0) / target
            # NB: tot_wip in alcuni fogli è già la somma, ma ricalcoliamolo come WIP + Finiti se ha senso, 
            # oppure consideriamolo solo WIP. Nel dubbio manteniamo (tot_wip + finiti) / target.
            # Se la cella tot_wip contiene già i finiti, la formula corretta è tot_wip / target.
            # Assumiamo tot_wip / target come standard per WIP totale = linea + finiti.
            c['gg_copertura_wip_fin'] = c.get('tot_wip', 0) / target

        # Semaforo basato su copertura
        # Usiamo i finiti come metrica principale per il semaforo, poichè richiesto dall'utente
        gg = c.get('gg_copertura_finiti', 0) or 0
        if gg <= 1:
            c['semaforo'] = 'rosso'
        elif gg <= 3:
            c['semaforo'] = 'giallo'
        else:
            c['semaforo'] = 'verde'
    return components


# ═══════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════
def _read_excel_flexible(filepath):
    """
    Tenta di leggere un file Excel con diversi motori.
    Supporta .xlsx (openpyxl), .xls (xlrd) e tabelle HTML (mascherate da .xls).
    """
    if not os.path.exists(filepath):
        return None
        
    # Controllo preliminare per file corrotti o segnaposto (es. 11 byte "TESTCONTENT")
    if os.path.getsize(filepath) < 100:
        try:
            with open(filepath, 'r', errors='ignore') as f:
                content = f.read(20).strip()
                if content and not content.startswith(('\x50\x4B', '\xD0\xCF')): # PK... o DOC...
                    return None
        except:
            pass

    # 1. Tenta con il motore di default (solitamente openpyxl per .xlsx)
    try:
        return pd.read_excel(filepath)
    except Exception:
        pass
        
    # 2. Tenta specificando xlrd (per veri .xls binari)
    try:
        return pd.read_excel(filepath, engine='xlrd')
    except Exception:
        pass
        
    # 3. Tenta come tabella HTML (molti export SAP .xls sono in realtà HTML)
    try:
        # Silenzia i warning di lxml se presente
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            dfs = pd.read_html(filepath)
            if dfs:
                return dfs[0]
    except Exception:
        pass
        
    # 4. Tenta come TSV (Tab Separated Values) codificato UTF-16 (comune in SAP)
    try:
        # Molti export SAP hanno righe di intestazione report prima della tabella vera e propria.
        with open(filepath, 'rb') as f:
            raw_content = f.read()
            content = raw_content.decode('utf-16')
            lines   = content.split('\n')
            skip    = 0
            # Cerchiamo l'intestazione nelle prime 100 righe (alcune Query SAP sono lunghe)
            for i, line in enumerate(lines[:100]):
                tokens = [t.strip() for t in line.split('\t')]
                if any(x in tokens for x in ['Materiale', 'Material', 'MATNR', 'Componente']):
                    skip = i
                    break
        
        # Riavviamo la lettura CSV con lo skip rilevato
        import io
        df = pd.read_csv(io.StringIO(content), sep='\t', skiprows=skip)
        if not df.empty and len(df.columns) > 1:
            return df
    except Exception:
        pass
        
    print(f"  [AVVISO] Impossibile leggere il file {filepath} (formato non riconosciuto o corrotto).")
    return None

def _to_num(val):
    try:
        v = float(val)
        return 0.0 if np.isnan(v) else v
    except (TypeError, ValueError):
        return 0.0

def _to_int(val):
    """Versione sicura di int(_to_num(val)): non crasha su NaN/inf."""
    return int(_to_num(val))

def _find_col(df, candidates):
    for c in candidates:
        for col in df.columns:
            if c.lower() in str(col).lower():
                return col
    return None


# ═══════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════
def run(base_dir="."):
    print("\n" + "="*50)
    print("  BAP PROCESSOR")
    print("="*50)

    def path(key):
        fname = CONFIG[key]
        p = os.path.join(base_dir, fname)
        # Se il file non esiste, prova con l'estensione opposta (.xlsx <-> .xls)
        if not os.path.exists(p):
            if fname.upper().endswith('.XLSX'):
                alt = fname[:-5] + '.xls'
            elif fname.upper().endswith('.XLS'):
                alt = fname[:-4] + '.xlsx'
            else:
                alt = fname
            p_alt = os.path.join(base_dir, alt)
            if os.path.exists(p_alt):
                return p_alt
        return p

    print("\n[1] Lettura baseline...")
    load_sap_mapping(path("sap_mapping"))
    raw_c = load_baseline_data(path("master_data"), path("inventory_baseline"))

    print("\n[2] Lettura SAP (opzionale)...")
    zpp = _read_excel_flexible(path("sap_zpp093"))
    mb51 = _read_excel_flexible(path("sap_mb51"))

    # Elaborazione
    all_c = process_components(baseline, zpp, mb51, overrides, targets)

    # Ordinamento (ECO prima, poi il resto)
    ECO_FAM_ORDER = ['DG2','SG1','SG2','SG3','SG4','SG5','SG6','SG7','SG8','SGR','SGRW','RG','FG5','FG57','PIGNON']
    def eco_sort_key(c):
        if c.get('progetto') != 'DCT Eco':
            return (99, 0, '')
        fam = c.get('famiglia', '')
        try:
            fam_idx = ECO_FAM_ORDER.index(fam)
        except ValueError:
            fam_idx = 98
        cod = c.get('codice_s', '') or c.get('codice_hard', '') or ''
        return (fam_idx, 0, cod)

    eco_sorted = sorted([c for c in all_c if c.get('progetto') == 'DCT Eco'],  key=eco_sort_key)
    others = [c for c in all_c if c.get('progetto') != 'DCT Eco']
    all_c = eco_sorted + others

    # Aggiungi indice display_order per preservarlo nel JS
    for i, c in enumerate(all_c):
        c['display_order'] = i

    print("\n[3] Estrazione Matrice P-NUM...")
    pnum_matrix = load_pnum_matrix(os.path.join(base_dir, "pnumb.xlsx"), zpp)

    print(f"\n[4] Totale componenti: {len(all_c)}")

    final_data = {
        "generato_il": datetime.now().isoformat(),
        "componenti": all_c,
        "pnum_matrix": pnum_matrix, # Nuova chiave per la pagina dedicata
        "sap_disponibile": bool(zpp is not None or mb51 is not None),
        "sap_mapping": {
            "materiali": SAP_MATERIALS_MAP,
            "stazioni": SAP_STATIONS_MAP
        }
    }

    out_json = path("output_json")
    try:
        with open(out_json, 'w', encoding='utf-8') as f:
            json.dump(final_data, f, ensure_ascii=False, indent=2)
        print(f"  -> JSON: {out_json}")
    except Exception as e:
        print(f"  [AVVISO] Impossibile scrivere JSON locale: {e}")

    print("\n[5] Generazione dashboard (legacy)...")
    try:
        from bap_dashboard import genera_dashboard
        genera_dashboard(all_c, path("output_dashboard"), mapping={
            "materiali": SAP_MATERIALS_MAP,
            "stazioni": SAP_STATIONS_MAP
        }, pnum_matrix=pnum_matrix)
        print(f"  -> HTML: {path('output_dashboard')}")
    except Exception as e:
        print(f"  [AVVISO] Dashboard legacy non generata: {e}")

    print("\n✅ Completato!\n")
    return final_data


if __name__ == "__main__":
    base = os.path.dirname(os.path.abspath(__file__))
    run(base_dir=base)
