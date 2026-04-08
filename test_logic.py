SAP_MATERIALS_MAP = {'M0162587': {'soft': 'm0162587/s', 'inter': 'm0162587/t', 'hard': 'm0162587'}}
c = {'progetto': '8Fe', 'codice_s': 'm0162587/s'}

hard_key = c.get('codice_hard', '').strip().upper()
sap_codes = SAP_MATERIALS_MAP.get(hard_key, {})

if not sap_codes:
    fallback = c.get('codice_s', '').replace('/S','').replace('/s','').strip().upper()
    real_sap_codes = SAP_MATERIALS_MAP.get(fallback)
    if real_sap_codes:
        sap_codes = real_sap_codes
    else:
        sap_codes = { "hard": fallback }

if sap_codes.get('soft') and sap_codes.get('hard'):
    if not c.get('codice_hard') or c['codice_hard'].strip() == "":
        c['codice_hard'] = sap_codes['hard']
    if not c.get('codice_s') or c['codice_s'].strip() == "":
        c['codice_s'] = sap_codes['soft']

print("Finale c['codice_hard']:", c.get('codice_hard'))
