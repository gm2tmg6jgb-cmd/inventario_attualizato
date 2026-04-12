import http.server
import email.parser
import json
import os
import re
import sys
import glob
import subprocess
import webbrowser
import urllib.request
import urllib.error
from datetime import datetime

PORT         = 8000
ARCHIVE_DIR  = 'archivio'
CONFIG_FILE  = 'bap_config.json'
MAX_JSON_MB  = 10          # limite body JSON (MB)
MAX_UPLOAD_MB = 50         # limite upload file (MB)

FIELD_TO_FILENAME = {
    'sap_zpp':      'CONFERMESAP', # L'estensione verrà gestita dal parser e dallo script di upload
    'sap_mb51':     'mb51',
    'sap_mapping':  'BAP1',
}

# Regex per validare id archivio locale  es. 20250330_143022.json
_LOCAL_ID_RE    = re.compile(r'^\d{8}_\d{6}\.json$')
# Regex per validare UUID Supabase
_UUID_RE        = re.compile(r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$', re.I)

# ─── Config ──────────────────────────────────────────────────────────────────

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_config(cfg):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

# ─── Archive backends ─────────────────────────────────────────────────────────

def _build_kpi(comps):
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
    return kpi


class LocalArchive:
    def _validate(self, id_):
        """Ritorna il basename validato o None se non sicuro."""
        fname = os.path.basename(id_)
        return fname if _LOCAL_ID_RE.match(fname) else None

    def save(self, label, comps, kpi):
        os.makedirs(ARCHIVE_DIR, exist_ok=True)
        ts    = datetime.now()
        fname = ts.strftime('%Y%m%d_%H%M%S') + '.json'
        tmp   = os.path.join(ARCHIVE_DIR, fname + '.tmp')
        dest  = os.path.join(ARCHIVE_DIR, fname)
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump({'timestamp': ts.isoformat(), 'label': label,
                       'kpi': kpi, 'componenti': comps}, f, ensure_ascii=False, indent=2)
        os.replace(tmp, dest)   # scrittura atomica
        return fname

    def list(self):
        os.makedirs(ARCHIVE_DIR, exist_ok=True)
        result = []
        for fp in sorted(glob.glob(os.path.join(ARCHIVE_DIR, '*.json')), reverse=True):
            try:
                with open(fp, 'r', encoding='utf-8') as f:
                    d = json.load(f)
                result.append({'id': os.path.basename(fp),
                                'timestamp': d.get('timestamp', ''),
                                'label': d.get('label', ''),
                                'kpi': d.get('kpi', {})})
            except Exception:
                pass
        return result

    def load(self, id_):
        fname = self._validate(id_)
        if not fname:
            return None
        fp = os.path.join(ARCHIVE_DIR, fname)
        if not os.path.exists(fp):
            return None
        with open(fp, 'r', encoding='utf-8') as f:
            return json.load(f)

    def delete(self, id_):
        fname = self._validate(id_)
        if not fname:
            return False
        fp = os.path.join(ARCHIVE_DIR, fname)
        if os.path.exists(fp):
            os.remove(fp)
            return True
        return False


class SupabaseArchive:
    def __init__(self, url, key, table='bap_archivio'):
        self.base    = url.rstrip('/') + '/rest/v1/' + table
        self.headers = {
            'apikey':        key,
            'Authorization': 'Bearer ' + key,
            'Content-Type':  'application/json',
            'Prefer':        'return=representation',
        }

    def _validate_uuid(self, id_):
        return id_ if _UUID_RE.match(str(id_)) else None

    def _req(self, method, suffix='', body=None):
        url  = self.base + suffix
        data = json.dumps(body).encode() if body is not None else None
        req  = urllib.request.Request(url, data=data, headers=self.headers, method=method)
        with urllib.request.urlopen(req, timeout=10) as r:
            return json.loads(r.read())

    def save(self, label, comps, kpi):
        rows = self._req('POST', body={'label': label, 'kpi': kpi, 'componenti': comps})
        return rows[0]['id'] if rows else None

    def list(self):
        rows = self._req('GET', '?select=id,created_at,label,kpi&order=created_at.desc')
        return [{'id': r['id'], 'timestamp': r['created_at'],
                 'label': r.get('label', ''), 'kpi': r.get('kpi', {})} for r in rows]

    def load(self, id_):
        uid = self._validate_uuid(id_)
        if not uid:
            return None
        rows = self._req('GET', f'?select=*&id=eq.{uid}')
        if not rows:
            return None
        r = rows[0]
        return {'timestamp': r.get('created_at', ''), 'label': r.get('label', ''),
                'kpi': r.get('kpi', {}), 'componenti': r.get('componenti', [])}

    def delete(self, id_):
        uid = self._validate_uuid(id_)
        if not uid:
            return False
        self._req('DELETE', f'?id=eq.{uid}')
        return True


def get_archive_backend():
    cfg = load_config().get('archive', {})
    if (cfg.get('mode') == 'supabase'
            and cfg.get('supabase_url')
            and cfg.get('supabase_anon_key')):
        return SupabaseArchive(cfg['supabase_url'], cfg['supabase_anon_key'],
                               cfg.get('supabase_table', 'bap_archivio'))
    return LocalArchive()


# ─── HTTP Handler ─────────────────────────────────────────────────────────────

class UploadHandler(http.server.SimpleHTTPRequestHandler):

    def _json_response(self, code, obj):
        body = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self):
        """Legge il body JSON con limite MAX_JSON_MB."""
        limit  = MAX_JSON_MB * 1024 * 1024
        length = int(self.headers.get('Content-Length', 0))
        if length > limit:
            raise ValueError(f'Body troppo grande ({length} byte, max {limit})')
        if length == 0:
            return {}   # nessun body → dizionario vuoto, non errore
        return json.loads(self.rfile.read(length))

    def _parse_upload(self):
        """
        Parsing manuale di multipart/form-data per massima compatibilità.
        """
        content_type = self.headers.get('Content-Type', '')
        if 'boundary=' not in content_type:
            return {}
        
        boundary = b'--' + content_type.split('boundary=')[1].encode()
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)
        
        parts = body.split(boundary)
        fields = {}
        
        for part in parts:
            if not part or part == b'--\r\n' or part == b'--':
                continue
            
            # Separa header e contenuto della parte
            try:
                head_raw, content = part.split(b'\r\n\r\n', 1)
                content = content.rsplit(b'\r\n', 1)[0] # rimuove \r\n finale
                head = head_raw.decode('utf-8', errors='ignore')
            except ValueError:
                continue
                
            # Estrai name e filename
            name_match = re.search(r'name="([^"]+)"', head)
            file_match = re.search(r'filename="([^"]+)"', head)
            
            if name_match:
                name = name_match.group(1)
                filename = file_match.group(1) if file_match else ""
                fields.setdefault(name, []).append({'filename': filename, 'data': content})
                
        return fields

    def _run_processor(self):
        """Esegue bap_processor.py con lo stesso interprete del server."""
        try:
            res = subprocess.run(
                [sys.executable, 'bap_processor.py'],
                capture_output=True, text=True, timeout=300
            )
            if res.returncode == 0:
                return True, 'Salvato e ricalcolato con successo.'
            err = [l for l in res.stderr.splitlines() if l.strip()][-3:]
            return False, 'Salvataggio OK, errore ricalcolo: ' + ' | '.join(err)
        except subprocess.TimeoutExpired:
            return False, 'Salvataggio OK, timeout ricalcolo.'
        except Exception as e:
            return False, f'Salvataggio OK, errore: {type(e).__name__}'

    # ── GET ──────────────────────────────────────────────────────────────────

    def do_GET(self):
        if self.path == '/archive-list':
            try:
                self._json_response(200, get_archive_backend().list())
            except Exception as e:
                self._json_response(500, {'message': str(e)})

        elif self.path.startswith('/archive-load/'):
            id_ = self.path.split('/')[-1]
            try:
                data = get_archive_backend().load(id_)
                if data:
                    self._json_response(200, data)
                else:
                    self._json_response(404, {'message': 'Non trovato.'})
            except Exception as e:
                self._json_response(500, {'message': str(e)})

        elif self.path == '/api/data':
            fname = 'dati_bap.json'
            if os.path.exists(fname):
                with open(fname, 'r', encoding='utf-8') as f:
                    self._json_response(200, json.load(f))
            else:
                self._json_response(404, {'message': 'File dati non trovato.'})

        elif self.path == '/archive-config':
            cfg = load_config().get('archive', {})
            self._json_response(200, {
                'mode':           cfg.get('mode', 'local'),
                'supabase_url':   cfg.get('supabase_url', ''),
                'supabase_table': cfg.get('supabase_table', 'bap_archivio'),
                'configured':     bool(cfg.get('supabase_url') and cfg.get('supabase_anon_key')),
            })

        elif self.path == '/api/get-master':
            fname = 'bap_master.json'
            if os.path.exists(fname):
                with open(fname, 'r', encoding='utf-8') as f:
                    self._json_response(200, json.load(f))
            else:
                self._json_response(200, [])
            return

        elif self.path in ('/dashboard.html', '/', '/index.html'):
            self.path = '/index.html'
            return super().do_GET()
        else:
            super().do_GET()

    # ── POST ─────────────────────────────────────────────────────────────────

    def do_POST(self):
        ts_str = datetime.now().strftime('%H:%M:%S')

        if self.path == '/upload':
            print(f"[{ts_str}] Ricevuta richiesta di upload...")
            parsed         = self._parse_upload()
            uploaded_count = 0
            
            print(f"[{ts_str}] Campi ricevuti dal form: {list(parsed.keys())}")
            
            for field, items in parsed.items():
                for item in items:
                    fname = item['filename']
                    if not fname: continue
                    
                    ext = os.path.splitext(fname.lower())[1]
                    if ext not in ('.xlsx', '.xls'):
                        print(f"[{ts_str}] File ignorato (estensione non valida): {fname}")
                        continue
                    
                    # Nome target standard basato sul campo del form
                    target_base = FIELD_TO_FILENAME.get(field)
                    if not target_base:
                        # Se il campo non è mappato, usa il nome file originale pulito
                        target_base = os.path.splitext(os.path.basename(fname))[0]
                    
                    target = target_base + ext
                    print(f"[{ts_str}] Tenta salvataggio: {fname} (campo: {field}) -> {target}")
                    
                    try:
                        # scrittura atomica tramite file temporaneo
                        tmp = target + '.tmp'
                        with open(tmp, 'wb') as f:
                            f.write(item['data'])
                        os.replace(tmp, target)
                        print(f"[{ts_str}] ✅ Salvato correttamente: {target} ({len(item['data'])} byte)")
                        uploaded_count += 1
                    except Exception as e:
                        print(f"[{ts_str}] ❌ Errore durante il salvataggio di {target}: {e}")

            if uploaded_count > 0:
                print(f"[{ts_str}] Avvio elaborazione processor...")
                ok, msg = self._run_processor()
                print(f"[{ts_str}] Finita elaborazione: {'✅' if ok else '❌'} {msg}")

            self.send_response(303)
            self.send_header('Location', '/dashboard.html')
            self.end_headers()
            return

        if self.path == '/clear-data':
            deleted = []
            # 1. Cancella i file SAP caricati (SOLO dati produzione, NO mapping)
            for key, fname_base in FIELD_TO_FILENAME.items():
                if key == 'sap_mapping': continue # NON cancellare mai la mappatura BAP1
                for ext in ('.xlsx', '.XLSX', '.xls', '.XLS'):
                    p = fname_base + ext
                    if os.path.exists(p):
                        os.remove(p)
                        deleted.append(p)
            # 2. Cancella le modifiche manuali (override inventario)
            for extra in ('bap_overrides.json',):
                if os.path.exists(extra):
                    os.remove(extra)
                    deleted.append(extra)
            msg = f'Eliminati {len(deleted)} file (SAP + override manuali).' if deleted else 'Nessun file da eliminare.'
            print(f'[{ts_str}] {msg} - Avvio ricalcolo dalla baseline...')
            self._run_processor()
            print(f'[{ts_str}] ✅ Dashboard ripristinata alla baseline.')
            self._json_response(200, {'message': msg})
            return

        # tutti gli altri POST leggono JSON
        try:
            data = self._read_json()
        except ValueError as e:
            self._json_response(413, {'message': str(e)})
            return
        except Exception:
            self._json_response(400, {'message': 'JSON non valido'})
            return

        if self.path in ('/api/save-targets', '/api/save-inventory'):
            if self.path == '/api/save-targets':
                tmp = 'bap_targets.json.tmp'
                with open(tmp, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                os.replace(tmp, 'bap_targets.json')
                print(f'[{ts_str}] Target salvati: {data}')
            else:
                # Gestione salvataggio combinato: Ovverrides (qty) + Master (metadata)
                ov_data = data.get('overrides', {})
                ma_data = data.get('master', {})

                # 1. Salva Overrides
                if ov_data:
                    fname_ov = 'bap_overrides.json'
                    existing_ov = {}
                    if os.path.exists(fname_ov):
                        with open(fname_ov, 'r', encoding='utf-8') as f:
                            existing_ov = json.load(f)
                    for k, v in ov_data.items():
                        if k not in existing_ov:
                            existing_ov[k] = v
                        else:
                            # Merge profondo (stazioni + altri campi)
                            if 'stazioni' in v:
                                if 'stazioni' not in existing_ov[k]:
                                    existing_ov[k]['stazioni'] = {}
                                existing_ov[k]['stazioni'].update(v['stazioni'])
                            if 'finiti' in v:
                                existing_ov[k]['finiti'] = v['finiti']
                            if 'tot_wip' in v:
                                existing_ov[k]['tot_wip'] = v['tot_wip']
                    
                    tmp_ov = fname_ov + '.tmp'
                    with open(tmp_ov, 'w', encoding='utf-8') as f:
                        json.dump(existing_ov, f, ensure_ascii=False, indent=2)
                    os.replace(tmp_ov, fname_ov)
                    print(f'[{ts_str}] Override salvati (merge): {len(ov_data)} componenti')

                # 2. Salva modifiche Master (struttura)
                if ma_data:
                    fname_ma = 'bap_master.json'
                    if os.path.exists(fname_ma):
                        with open(fname_ma, 'r', encoding='utf-8') as f:
                            master_list = json.load(f)
                        
                        updated_count = 0
                        for key, mods in ma_data.items():
                            # key = progetto||label
                            for item in master_list:
                                if (item.get('progetto','') + '||' + item.get('label','')) == key:
                                    item.update(mods)
                                    
                                    # SINCRONIZZA stazioni_list con le chiavi di stazioni (se presente)
                                    if 'stazioni' in mods:
                                        if 'stazioni_list' not in item:
                                            item['stazioni_list'] = []
                                        # Aggiungi solo le stazioni mancanti preservando l'ordine
                                        for st_name in mods['stazioni'].keys():
                                            if st_name not in item['stazioni_list']:
                                                item['stazioni_list'].append(st_name)
                                                
                                    updated_count += 1
                                    break
                        
                        if updated_count > 0:
                            tmp_ma = fname_ma + '.tmp'
                            with open(tmp_ma, 'w', encoding='utf-8') as f:
                                json.dump(master_list, f, ensure_ascii=False, indent=2)
                            os.replace(tmp_ma, fname_ma)
                            print(f'[{ts_str}] Master data aggiornata: {updated_count} componenti')

            ok, msg = self._run_processor()
            print(f'[{ts_str}] {"✅" if ok else "❌"} {msg}')
            
            # Restituisci anche i dati freschi per permettere al frontend di aggiornarsi
            res_data = {}
            if os.path.exists('dati_bap.json'):
                with open('dati_bap.json', 'r', encoding='utf-8') as f:
                    res_data = json.load(f)
            
            self._json_response(200, {'message': msg, 'data': res_data})

        elif self.path == '/api/archive-save':
            label     = data.get('label', '').strip()[:80]
            comp_key  = data.get('component_key', '').strip()[:120]
            if not os.path.exists('dati_bap.json'):
                self._json_response(400, {'message': 'Nessun dato. Importa prima i file.'})
                return
            with open('dati_bap.json', 'r', encoding='utf-8') as f:
                source = json.load(f)
            comps = source.get('componenti', [])
            if comp_key:
                comps = [c for c in comps
                         if (c.get('progetto','') + '||' + c.get('label','')) == comp_key]
            kpi = _build_kpi(comps)
            try:
                backend = get_archive_backend()
                id_     = backend.save(label, comps, kpi)
                ts      = datetime.now()
                scope   = 'componente' if comp_key else 'completo'
                print(f'[{ts_str}] Archivio salvato ({type(backend).__name__}, {scope}): {id_}')
                self._json_response(200, {'message': f"Archiviato: {ts.strftime('%d/%m/%Y %H:%M')}", 'id': id_})
            except Exception as e:
                self._json_response(500, {'message': f'Errore archivio: {e}'})

        elif self.path == '/api/archive-delete':
            id_ = data.get('id', '')
            try:
                ok = get_archive_backend().delete(id_)
                print(f'[{ts_str}] Archivio eliminato: {id_}')
                self._json_response(200 if ok else 404,
                                    {'message': 'Eliminato.' if ok else 'Non trovato o id non valido.'})
            except Exception as e:
                self._json_response(500, {'message': str(e)})

        elif self.path == '/archive-config':
            cfg = load_config()
            cfg['archive'] = {
                'mode':              data.get('mode', 'local'),
                'supabase_url':      data.get('supabase_url', '').rstrip('/'),
                'supabase_anon_key': data.get('supabase_anon_key', ''),
                'supabase_table':    data.get('supabase_table', 'bap_archivio'),
            }
            save_config(cfg)
            mode = cfg['archive']['mode']
            print(f'[{ts_str}] Configurazione archivio: {mode}')
            if mode == 'supabase':
                try:
                    get_archive_backend().list()
                    self._json_response(200, {'message': 'Supabase connesso correttamente.'})
                except Exception as e:
                    self._json_response(200, {'message': f'Config salvata. Errore test: {e}'})
            else:
                self._json_response(200, {'message': 'Modalita locale attivata.'})

        elif self.path == '/api/save-station-mapping':
            # Salva la mappatura manuale Stazioni -> Operazioni SAP
            fname = 'bap_mapping_permanent.json'
            mapping = {}
            if os.path.exists(fname):
                with open(fname, 'r', encoding='utf-8') as f:
                    mapping = json.load(f)
            
            # Assicurati che le sezioni esistano
            if 'materiali' not in mapping: mapping['materiali'] = {}
            mapping['stazioni'] = data # Il corpo della POST è il nuovo dizionario stazioni
            
            # Salvataggio atomico
            tmp = fname + '.tmp'
            with open(tmp, 'w', encoding='utf-8') as f:
                json.dump(mapping, f, ensure_ascii=False, indent=2)
            os.replace(tmp, fname)
            print(f'[{ts_str}] Mappatura stazioni salvata: {len(data)} stazioni')
            
            # Ricalcola
            ok, msg = self._run_processor()
            print(f'[{ts_str}] {"✅" if ok else "❌"} {msg}')
            self._json_response(200, {'message': 'Mappature salvate e dati ricalcolati.'})

        elif self.path == '/api/clear-data':
            # La logica principale è gestita all'inizio del metodo per efficienza.
            self._json_response(200, {'message': 'Dashboard resettata.'})
        elif self.path == '/api/reset-baseline':
            # Azzera tutte le quantità nella baseline (punto zero)
            BASELINE_FILE = 'bap_inventory_baseline.json'
            if not os.path.exists(BASELINE_FILE):
                self._json_response(404, {'message': 'File baseline non trovato.'})
                return
            try:
                with open(BASELINE_FILE, 'r', encoding='utf-8') as f:
                    baseline = json.load(f)
                
                # Backup della baseline corrente prima di azzerarla
                backup = BASELINE_FILE.replace('.json', f'_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json')
                import shutil
                shutil.copy2(BASELINE_FILE, backup)
                print(f'[{ts_str}] Backup baseline creato prima del reset: {backup}')

                # ... (rest of reset logic remains same)

                today = datetime.now().strftime('%Y-%m-%d')
                for key in baseline:
                    baseline[key]['finiti'] = 0
                    baseline[key]['data_inv'] = today
                    sts = baseline[key].get('stazioni', {})
                    for st in sts:
                        sts[st] = 0
                tmp = BASELINE_FILE + '.tmp'
                with open(tmp, 'w', encoding='utf-8') as f:
                    json.dump(baseline, f, ensure_ascii=False, indent=2)
                os.replace(tmp, BASELINE_FILE)
                print(f'[{ts_str}] Baseline azzerata ({len(baseline)} componenti). Avvio ricalcolo...')
                self._run_processor()
                print(f'[{ts_str}] ✅ Dashboard a zero.')
                self._json_response(200, {'message': f'Baseline azzerata: {len(baseline)} componenti portati a zero.'})
            except Exception as e:
                self._json_response(500, {'message': f'Errore azzeramento: {e}'})

        elif self.path == '/api/delete-component':
            key = data.get('key', '')
            if not key:
                self._json_response(400, {'message': 'Key mancante.'})
                return
            
            fname_ma = 'bap_master.json'
            if os.path.exists(fname_ma):
                with open(fname_ma, 'r', encoding='utf-8') as f:
                    master_list = json.load(f)
                
                # Filtra via il componente
                new_list = [item for item in master_list 
                            if (item.get('progetto','') + '||' + item.get('label','')) != key]
                
                if len(new_list) < len(master_list):
                    tmp = fname_ma + '.tmp'
                    with open(tmp, 'w', encoding='utf-8') as f:
                        json.dump(new_list, f, ensure_ascii=False, indent=2)
                    os.replace(tmp, fname_ma)
                    
                    # Rimuoviamo anche eventuali override per pulizia
                    fname_ov = 'bap_overrides.json'
                    if os.path.exists(fname_ov):
                        with open(fname_ov, 'r', encoding='utf-8') as f:
                            ov_data = json.load(f)
                        if key in ov_data:
                            del ov_data[key]
                            with open(fname_ov, 'w', encoding='utf-8') as f:
                                json.dump(ov_data, f, ensure_ascii=False, indent=2)

                    print(f'[{ts_str}] Componente eliminato: {key}')
                    self._run_processor()
                    self._json_response(200, {'ok': True, 'message': 'Componente eliminato con successo.'})
                else:
                    self._json_response(404, {'message': 'Componente non trovato.'})
            else:
                self._json_response(500, {'message': 'File master non trovato.'})

        elif self.path == '/api/save-baseline':
            # Salva lo stato attuale come nuova baseline
            BASELINE_FILE = 'bap_inventory_baseline.json'
            DATI_FILE     = 'dati_bap.json'
            label = data.get('label', datetime.now().strftime('%Y-%m-%d'))
            if not os.path.exists(DATI_FILE):
                self._json_response(400, {'message': 'Nessun dato elaborato. Importa prima i file SAP.'})
                return
            try:
                with open(DATI_FILE, 'r', encoding='utf-8') as f:
                    dati = json.load(f)
                # Backup della baseline corrente
                if os.path.exists(BASELINE_FILE):
                    backup = BASELINE_FILE.replace('.json', f'_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json')
                    import shutil
                    shutil.copy2(BASELINE_FILE, backup)
                    print(f'[{ts_str}] Backup baseline: {backup}')
                # Costruisci nuova baseline dai dati elaborati correnti
                new_baseline = {}
                for c in dati.get('componenti', []):
                    key = f"{c.get('progetto', '')}||{c.get('label', '')}"
                    stazioni_new = {}
                    for st_name, st_data in c.get('stazioni', {}).items():
                        qty = st_data.get('qty', 0) if isinstance(st_data, dict) else int(st_data)
                        stazioni_new[st_name] = qty
                    new_baseline[key] = {
                        'finiti':   int(c.get('finiti', 0)),
                        'stazioni': stazioni_new,
                        'data_inv': label,
                    }
                tmp = BASELINE_FILE + '.tmp'
                with open(tmp, 'w', encoding='utf-8') as f:
                    json.dump(new_baseline, f, ensure_ascii=False, indent=2)
                os.replace(tmp, BASELINE_FILE)
                print(f'[{ts_str}] Nuova baseline salvata: {len(new_baseline)} componenti, data={label}')
                self._run_processor()
                print(f'[{ts_str}] ✅ Nuova baseline attiva.')
                self._json_response(200, {'message': f'Nuova baseline salvata ({len(new_baseline)} componenti, data: {label}).'})
            except Exception as e:
                self._json_response(500, {'message': f'Errore salvataggio baseline: {e}'})

        elif self.path == '/api/save-master':
            fname = 'bap_master.json'
            tmp = fname + '.tmp'
            with open(tmp, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            os.replace(tmp, fname)
            print(f'[{ts_str}] Register Master salvato: {len(data)} componenti')
            
            # Ricalcola
            ok, msg = self._run_processor()
            print(f'[{ts_str}] {"✅" if ok else "❌"} {msg}')
            
            # Restituisci i dati aggiornati
            with open('dati_bap.json', 'r', encoding='utf-8') as f:
                res_data = json.load(f)
            self._json_response(200, {'message': 'Master salvato con successo', 'data': res_data})

        else:
            self.send_error(404, 'Not Found')

    def log_message(self, fmt, *args):
        # Silenzia i log HTTP di default (già stampiamo i nostri)
        pass


def start_server():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    try:
        server = http.server.HTTPServer(('localhost', PORT), UploadHandler)
        url    = f'http://localhost:{PORT}/'
        print('\n' + '='*50)
        print(' BAP SERVER - Dashboard Interattiva')
        print('='*50)
        print(f'\nDashboard: {url}')
        cfg  = load_config().get('archive', {})
        mode = cfg.get('mode', 'local')
        print(f'[Archivio] modalita: {mode}')
        print(f'[Python]   {sys.executable}')
        print('[INFO] Premi CTRL+C per fermare.\n')
        
        # webbrowser.open rimosso per stabilità backend
            
        print('[INFO] Server in ascolto. Premi CTRL+C per fermare.\n')
        server.serve_forever()
    except OSError as e:
        if e.errno == 48:
            print(f"\n[ERRORE] La porta {PORT} è già occupata.")
            print(f"Suggerimento: Chiudi altre istanze del server o usa un'altra porta.")
            print(f"Comando per trovare il processo: lsof -i :{PORT}\n")
        else:
            raise e
    except KeyboardInterrupt:
        print('\nServer arrestato.')
        server.server_close()


if __name__ == '__main__':
    start_server()
