"""
update_dashboard.py  –  Genera data.json con el cruce IMSS-Nómina.
El HTML lo carga con fetch('./data.json') en cada visita.
NO modifica el HTML directamente (evita errores de sintaxis JS).
"""
import io, json, re, sys, zipfile, requests, openpyxl

FOLDER_ID = '1QoOP4jnPmp7_x9tz1Q7l2Yy2JmOx7MXc'
API_KEY   = 'AIzaSyAId7gthv7EEzmaTrfbt07FK4Kf-ii51uM'
BASE_URL  = 'https://www.googleapis.com/drive/v3'
OUT_FILE  = 'data.json'

RFC_RE = re.compile(r'^[A-ZÑ&]{3,4}[0-9]{6}[A-Z0-9]{3}$')

def get_type(name):
    c = re.sub(r'\d{6,8}', '', name).lower()
    if 'imss'      in c: return 'imss'
    if re.search(r'n[oó]mina', c): return 'nomina'
    if 'cedular emp' in c or 'cedular empresarial' in c: return 'cedular_emp'
    if 'cedular'   in c: return 'cedular'
    if 'hospedaj'  in c: return 'hospedaje'   # cubre "hospedaje" y typos
    if 'profesional' in c: return 'profesional'
    if re.search(r'gases?', c): return 'gases'
    if 'agua'      in c: return 'agua'
    return None

def list_files():
    url = (f"{BASE_URL}/files?q='{FOLDER_ID}'+in+parents+and+trashed=false"
           f"&fields=files(id,name)&key={API_KEY}&pageSize=50")
    r = requests.get(url, timeout=30); r.raise_for_status()
    return [f for f in r.json().get('files', []) if get_type(f['name'])]

def download(fid):
    r = requests.get(f"{BASE_URL}/files/{fid}?alt=media&key={API_KEY}", timeout=120)
    r.raise_for_status(); return r.content

def extract_rfcs(content):
    rfcs = set()
    try:
        wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        ws = wb.active
        header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None) or []
        col = next((i for i, h in enumerate(header) if h and 'rfc' in str(h).lower()), 0)
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and col < len(row) and row[col]:
                v = str(row[col]).strip().upper()
                if RFC_RE.match(v): rfcs.add(v)
        wb.close()
    except Exception as e:
        # Fallback: regex sobre sharedStrings.xml del ZIP
        print(f'  openpyxl falló ({e}), usando regex ZIP')
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                ss = next((n for n in z.namelist() if 'sharedstrings' in n.lower()), None)
                if ss:
                    text = z.read(ss).decode('utf-8', errors='ignore')
                    rfcs = set(re.findall(r'[A-ZÑ&]{3,4}[0-9]{6}[A-Z0-9]{3}', text))
        except Exception as e2:
            print(f'  Fallback también falló: {e2}')
    return rfcs

def compute(rfc_sets):
    imss = rfc_sets.get('imss', set())
    nom  = rfc_sets.get('nomina', set())
    ced  = rfc_sets.get('cedular', set())
    hos  = rfc_sets.get('hospedaje', set())
    pro  = rfc_sets.get('profesional', set())
    gas  = rfc_sets.get('gases', set())
    agu  = rfc_sets.get('agua', set())
    cemp = rfc_sets.get('cedular_emp', set())

    sin = [r for r in imss if r not in nom]
    n   = len(sin)
    rows = [{'rfc': r,
             'cedular':     r in ced,
             'hospedaje':   r in hos,
             'profesional': r in pro,
             'gases':       r in gas,
             'agua':        r in agu,
             'cedular_emp': r in cemp} for r in sin]

    def pct(s): return round(len(s & set(sin)) / n * 100, 2) if n else 0.0
    stats = {'total_imss': len(imss), 'total_nomina': len(nom), 'sin_nomina': n,
             'pct_cedular': pct(ced), 'pct_hospedaje': pct(hos),
             'pct_profesional': pct(pro), 'pct_gases': pct(gas),
             'pct_agua': pct(agu), 'pct_cedular_emp': pct(cemp)}
    return rows, stats

def main():
    from datetime import datetime, timezone
    print('── Auditoría Fiscal ──')
    print('▸ Listando archivos…')
    files = list_files()
    if not files:
        print('✗ Carpeta vacía o no pública'); sys.exit(1)
    for f in files: print(f"  • {f['name']} → {get_type(f['name'])}")

    rfc_sets = {}
    for f in files:
        t = get_type(f['name'])
        print(f"▸ Descargando {f['name']}…")
        try:
            rfcs = extract_rfcs(download(f['id']))
            rfc_sets[t] = rfcs
            print(f"  {len(rfcs):,} RFCs")
        except Exception as e:
            print(f'  ⚠ Error: {e}'); rfc_sets[t] = set()

    print('▸ Calculando cruce…')
    rows, stats = compute(rfc_sets)
    print(f"  IMSS:{stats['total_imss']:,} | Nómina:{stats['total_nomina']:,} | Sin Nómina:{stats['sin_nomina']:,}")

    out = {'updated': datetime.now(timezone.utc).isoformat(), 'stats': stats, 'rows': rows}
    with open(OUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, separators=(',', ':'))
    print(f'✓ {OUT_FILE} escrito ({len(rows):,} registros)')

if __name__ == '__main__': main()
