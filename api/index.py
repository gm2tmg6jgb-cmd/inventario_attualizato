from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import json
import os
import re
import urllib.request
import urllib.error
from datetime import datetime
import pandas as pd
import numpy as np

app = Flask(__name__)
CORS(app)

# ─── Config & Constants ──────────────────────────────────────────────────────

ARCHIVE_TABLE = 'bap_archivio'
STATE_TABLE   = 'bap_state' # Tabella per lo stato "live" (targets, overrides)
FIELD_TO_FILENAME = {
    'sap_zpp':      'CONFERMESAP',
    'sap_mb51':     'mb51',
    'sap_mapping':  'BAP1',
}

def get_config():
    # Priority to environment variables (for Vercel)
    url = os.environ.get('SUPABASE_URL')
    key = os.environ.get('SUPABASE_ANON_KEY')
    
    if not url or not key:
        # Fallback to local config if it exists (for local testing)
        if os.path.exists('bap_config.json'):
            with open('bap_config.json', 'r') as f:
                cfg = json.load(f).get('archive', {})
                url = cfg.get('supabase_url')
                key = cfg.get('supabase_anon_key')
    
    return {
        'url': url,
        'key': key,
        'table': os.environ.get('SUPABASE_TABLE', ARCHIVE_TABLE)
    }

# ─── Supabase Helper ─────────────────────────────────────────────────────────

def supabase_req(method, suffix='', body=None):
    cfg = get_config()
    if not cfg['url'] or not cfg['key']:
        raise Exception("Supabase not configured. Set SUPABASE_URL and SUPABASE_ANON_KEY.")
    
    url = cfg['url'].rstrip('/') + '/rest/v1/' + cfg['table'] + suffix
    headers = {
        'apikey':        cfg['key'],
        'Authorization': 'Bearer ' + cfg['key'],
        'Content-Type':  'application/json',
        'Prefer':        'return=representation',
    }
    
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    
    with urllib.request.urlopen(req, timeout=10) as r:
        return json.loads(r.read())

# ─── API Routes ──────────────────────────────────────────────────────────────

@app.route('/api/archive-list', methods=['GET'])
def archive_list():
    try:
        rows = supabase_req('GET', '?select=id,created_at,label,kpi&order=created_at.desc')
        result = [{'id': r['id'], 'timestamp': r['created_at'], 
                   'label': r.get('label', ''), 'kpi': r.get('kpi', {})} for r in rows]
        return jsonify(result)
    except Exception as e:
        return jsonify({'message': str(e)}), 500

@app.route('/api/archive-load/<id_>', methods=['GET'])
def archive_load(id_):
    try:
        rows = supabase_req('GET', f'?select=*&id=eq.{id_}')
        if not rows:
            return jsonify({'message': 'Non trovato.'}), 404
        r = rows[0]
        return jsonify({
            'timestamp': r.get('created_at', ''),
            'label': r.get('label', ''),
            'kpi': r.get('kpi', {}),
            'componenti': r.get('componenti', [])
        })
    except Exception as e:
        return jsonify({'message': str(e)}), 500

@app.route('/api/archive-config', methods=['GET'])
def archive_config():
    cfg = get_config()
    return jsonify({
        'mode': 'supabase',
        'supabase_url': cfg['url'] or '',
        'supabase_table': cfg['table'],
        'configured': bool(cfg['url'] and cfg['key']),
    })

@app.route('/api/archive-save', methods=['POST'])
def archive_save():
    data = request.json
    label = data.get('label', '').strip()[:80]
    # In Vercel, we can't read 'dati_bap.json'. We assume 'componenti' is provided in the body 
    # OR we need a way to build it. Since the frontend has the data, 
    # it might be better if the frontend sends the whole state.
    # But for compatibility, let's see if we can read current state from Supabase?
    
    # NEW LOGIC: Frontend should ideally send the data to save.
    # If not present, we can't save what we don't have.
    comps = data.get('componenti', [])
    if not comps:
        return jsonify({'message': 'Dati mancanti per il salvataggio.'}), 400
    
    # Calculate KPI
    kpi = {}
    for c in comps:
        p = c.get('progetto', '?')
        if p not in kpi:
            kpi[p] = {'rossi': 0, 'gialli': 0, 'verdi': 0, 'finiti': 0, 'wip': 0}
        kpi[p]['finiti'] += c.get('finiti', 0)
        kpi[p]['wip']    += c.get('tot_wip', 0)
        sem = c.get('semaforo', 'verde')
        if sem == 'rosso':    kpi[p]['rossi']  += 1
        elif sem == 'giallo': kpi[p]['gialli'] += 1
        else:                 kpi[p]['verdi']  += 1

    try:
        rows = supabase_req('POST', body={'label': label, 'kpi': kpi, 'componenti': comps})
        return jsonify({'message': f"Archiviato con successo", 'id': rows[0]['id']})
    except Exception as e:
        return jsonify({'message': str(e)}), 500

import sys
import os

# Aggiungi la directory corrente al path per importare i moduli locali
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import bap_processor

# Cartella temporanea per Vercel
TEMP_DIR = '/tmp' if os.environ.get('VERCEL') else os.path.dirname(os.path.abspath(__file__))

def sync_from_supabase():
    """Scarica targets e overrides da Supabase in /tmp se presenti."""
    try:
        rows = supabase_req('GET', f'?select=data&id=eq.live', table=STATE_TABLE)
        if rows:
            state = rows[0].get('data', {})
            # Aggiunto bap_master.json alla sincronizzazione
            for key, filename in [('targets', 'bap_targets.json'), ('overrides', 'bap_overrides.json'), ('master', 'bap_master.json')]:
                if key in state:
                    with open(os.path.join(TEMP_DIR, filename), 'w', encoding='utf-8') as f:
                        json.dump(state[key], f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Errore sync Supabase: {e}")

def sync_to_supabase():
    """Salva targets, overrides e master correnti su Supabase."""
    try:
        state = {}
        for key, filename in [('targets', 'bap_targets.json'), ('overrides', 'bap_overrides.json'), ('master', 'bap_master.json')]:
            p = os.path.join(TEMP_DIR, filename)
            if not os.path.exists(p) and os.path.exists(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', filename)):
                # Fallback per il master se non ancora in /tmp
                p = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', filename)
                
            if os.path.exists(p):
                with open(p, 'r', encoding='utf-8') as f:
                    state[key] = json.load(f)
        
        # Upsert su Supabase
        supabase_req('POST', body={'id': 'live', 'data': state}, table=STATE_TABLE, prefer='resolution=merge-duplicates')
    except Exception as e:
        print(f"Errore save Supabase: {e}")

def supabase_req(method, suffix='', body=None, table=None, prefer=None):
    cfg = get_config()
    target_table = table or cfg['table']
    if not cfg['url'] or not cfg['key']:
        return [] # Silent fail if not configured
    
    url = cfg['url'].rstrip('/') + '/rest/v1/' + target_table + suffix
    headers = {
        'apikey':        cfg['key'],
        'Authorization': 'Bearer ' + cfg['key'],
        'Content-Type':  'application/json',
    }
    if prefer:
        headers['Prefer'] = prefer
    elif method == 'POST':
        headers['Prefer'] = 'return=representation'
    
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    
    try:
        with urllib.request.urlopen(req, timeout=10) as r:
            res = r.read()
            return json.loads(res) if res else []
    except Exception as e:
        print(f"Supabase error: {e}")
        return []

@app.route('/api/data', methods=['GET'])
def get_data():
    """Restituisce i dati elaborati correnti, sincronizzando prima da Supabase."""
    sync_from_supabase()
    try:
        result = bap_processor.run(base_dir=TEMP_DIR)
        return jsonify({'data': result})
    except Exception as e:
        return jsonify({'message': f'Errore: {str(e)}'}), 500

@app.route('/api/upload', methods=['POST'])
def upload():
    # Salvataggio dei file caricati in /tmp
    files = request.files
    uploaded = []
    
    for field, filename_base in FIELD_TO_FILENAME.items():
        if field in files:
            file = files[field]
            if file.filename:
                ext = os.path.splitext(file.filename.lower())[1]
                target = os.path.join(TEMP_DIR, filename_base + ext)
                file.save(target)
                uploaded.append(target)
                
    if uploaded:
        # Esegui il processore usando /tmp come base
        try:
            result = bap_processor.run(base_dir=TEMP_DIR)
            # In Vercel, non possiamo salvare il risultato su disco permanentemente,
            # lo restituiamo direttamente o lo salviamo su Supabase come "latest".
            return jsonify({
                'message': 'Elaborazione completata.',
                'data': result
            })
        except Exception as e:
            return jsonify({'message': f'Errore elaborazione: {str(e)}'}), 500
            
    return jsonify({'message': 'Nessun file caricato.'}), 400

@app.route('/api/save-targets', methods=['POST'])
def save_targets():
    data = request.json
    target_path = os.path.join(TEMP_DIR, 'bap_targets.json')
    with open(target_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    sync_to_supabase()
    
    # Rielabora
    try:
        result = bap_processor.run(base_dir=TEMP_DIR)
        return jsonify({'message': 'Target salvati.', 'data': result})
    except Exception as e:
        return jsonify({'message': str(e)}), 500

@app.route('/api/save-inventory', methods=['POST'])
def save_inventory():
    data = request.json
    override_path = os.path.join(TEMP_DIR, 'bap_overrides.json')
    
    existing = {}
    if os.path.exists(override_path):
        with open(override_path, 'r', encoding='utf-8') as f:
            existing = json.load(f)
    
    existing.update(data)
    with open(override_path, 'w', encoding='utf-8') as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
        
    sync_to_supabase()
    
    # Rielabora
    try:
        result = bap_processor.run(base_dir=TEMP_DIR)
        return jsonify({'message': 'Inventario salvato.', 'data': result})
    except Exception as e:
        return jsonify({'message': str(e)}), 500

@app.route('/api/get-master', methods=['GET'])
def get_master():
    sync_from_supabase()
    master_path = os.path.join(TEMP_DIR, 'bap_master.json')
    if not os.path.exists(master_path):
        # Fallback al file statico iniziale
        master_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'bap_master.json')
    
    try:
        with open(master_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return jsonify(data)
    except Exception as e:
        return jsonify({'message': str(e)}), 500

@app.route('/api/save-master', methods=['POST'])
def save_master():
    data = request.json
    master_path = os.path.join(TEMP_DIR, 'bap_master.json')
    with open(master_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    sync_to_supabase()
    
    try:
        result = bap_processor.run(base_dir=TEMP_DIR)
        return jsonify({'message': 'Master aggiornato.', 'data': result})
    except Exception as e:
        return jsonify({'message': str(e)}), 500

@app.route('/api/clear-data', methods=['POST'])
def clear_data():
    # Pulisce /tmp
    for field, filename_base in FIELD_TO_FILENAME.items():
        if field == 'sap_mapping': continue
        for ext in ('.xlsx', '.XLSX', '.xls', '.XLS'):
            p = os.path.join(TEMP_DIR, filename_base + ext)
            if os.path.exists(p): os.remove(p)
            
    for extra in ('bap_overrides.json', 'bap_targets.json'):
        p = os.path.join(TEMP_DIR, extra)
        if os.path.exists(p): os.remove(p)
        
    try:
        result = bap_processor.run(base_dir=TEMP_DIR)
        return jsonify({'message': 'Dati resettati.', 'data': result})
    except Exception as e:
        return jsonify({'message': str(e)}), 500

# Error handler
@app.errorhandler(Exception)
def handle_exception(e):
    return jsonify({'message': str(e)}), 500

# Vercel needs the app object
if __name__ == "__main__":
    app.run()
