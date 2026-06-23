"""
update_dashboard.py – Genera data.json con cruce IMSS-Nómina.
Extrae RFC, nombre, dirección y trabajadores del archivo IMSS.
"""
import io, json, re, sys, zipfile, requests, openpyxl
from datetime import datetime, timezone

FOLDER_ID = '1QoOP4jnPmp7_x9tz1Q7l2Yy2JmOx7MXc'
API_KEY   = 'AIzaSyAId7gthv7EEzmaTrfbt07FK4Kf-ii51uM'
BASE_URL  = 'https://www.googleapis.com/drive/v3'
OUT_FILE  = 'data.json'

RFC_RE = re.compile(r'^[A-ZÑ&]{3,4}[0-9]{6}[A-Z0-9]{3}$')

def get_type(name):
    c = re.sub(r'\d{6,8}', '', name).lower()
    if 'imss'        in c: return 'imss'
    if re.search(r'n[oó]mina', c): return 'nomina'
    if 'cedular emp' in c or 'cedular empresarial' in c: return 'cedular_emp'
    if 'cedular'     in c: return 'cedular'
    if 'hospedaj'    in c: return 'hospedaje'
    if 'profesional' in c: return 'profesional'
    if re.search(r'gases?', c): return 'gases'
    if 'agua'        in c: return 'agua'
    return None

def list_files():
    url = (f"{BASE_URL}/files?q='{FOLDER_ID}'+in+parents+and+trashed=false"
           f"&fields=files(id,name)&key={API_KEY}&pageSize=50")
    r = requests.get(url, timeout=30); r.raise_for_status()
    return [f for f in r.json().get('files', []) if get_type(f['name'])]

def download(fid):
    r = requests.get(f"{BASE_URL}/files/{fid}?alt=media&key={API_KEY}", timeout=180)
    r.raise_for_status(); return r.content

def find_col(headers, *keywords):
    h = [str(x).lower().strip() if x else '' for x in headers]
    for kw in keywords:
        for i, v in enumerate(h):
            if kw in v:
                return i
    return None

def extract_imss(content):
    records = []
    try:
        wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        ws = wb.active
        header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None) or []
        print(f'  Columnas IMSS: {[str(h) for h in header[:15]]}')

        rfc_col  = find_col(header, 'rfc') or 0
        nom_col  = find_col(header, 'razón social', 'razon social', 'nombre', 'denominaci')
        dir_col  = find_col(header, 'domicilio', 'direcci', 'calle', 'municipio')
        trab_col = find_col(header, 'trabajador', 'empleado', 'num. trab', 'número de trab')

        print(f'  rfc_col={rfc_col} nom_col={nom_col} dir_col={dir_col} trab_col={trab_col}')

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or rfc_col >= len(row) or not row[rfc_col]:
                continue
            v = str(row[rfc_col]).strip().upper()
            if not RFC_RE.match(v):
                continue
            def cell(col):
                if col is None or col >= len(row) or row[col] is None:
                    return ''
                return str(row[col]).strip()
            trab = None
            if trab_col is not None and trab_col < len(row) and row[trab_col] is not None:
                try: trab = int(float(str(row[trab_col])))
                except: pass
            records.append({'rfc': v, 'nombre': cell(nom_col),
                            'direccion': cell(dir_col), 'trabajadores': trab})
        wb.close()
    except Exception as e:
        print(f'  openpyxl IMSS fallo: {e}')
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                ss = next((n for n in z.namelist() if 'sharedstrings' in n.lower()), None)
                if ss:
                    text = z.read(ss).decode('utf-8', errors='ignore')
                    for r in re.findall(r'[A-ZN&]{3,4}[0-9]{6}[A-Z0-9]{3}', text):
                        records.append({'rfc': r, 'nombre': '', 'direccion': '', 'trabajadores': None})
        except Exception as e2:
            print(f'  Fallback ZIP fallo: {e2}')
    return records

def extract_rfcs(content):
    rfcs = set()
    try:
        wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        ws = wb.active
        header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None) or []
        col = find_col(header, 'rfc') or 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and col < len(row) and row[col]:
                v = str(row[col]).strip().upper()
                if RFC_RE.match(v): rfcs.add(v)
        wb.close()
    except Exception as e:
        print(f'  openpyxl fallo ({e}), usando regex ZIP')
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                ss = next((n for n in z.namelist() if 'sharedstrings' in n.lower()), None)
                if ss:
                    text = z.read(ss).decode('utf-8', errors='ignore')
                    rfcs = set(re.findall(r'[A-ZN&]{3,4}[0-9]{6}[A-Z0-9]{3}', text))
        except Exception as e2:
            print(f'  Fallback fallo: {e2}')
    return rfcs

def compute(imss_records, rfc_sets):
    nom  = rfc_sets.get('nomina',      set())
    ced  = rfc_sets.get('cedular',     set())
    hos  = rfc_sets.get('hospedaje',   set())
    pro  = rfc_sets.get('profesional', set())
    gas  = rfc_sets.get('gases',       set())
    agu  = rfc_sets.get('agua',        set())
    cemp = rfc_sets.get('cedular_emp', set())

    sin = [r for r in imss_records if r['rfc'] not in nom]
    n   = len(sin)
    sin_rfcs = {r['rfc'] for r in sin}

    rows = [{
        'rfc':         r['rfc'],
        'nombre':      r['nombre'],
        'direccion':   r['direccion'],
        'trabajadores': r['trabajadores'],
        'cedular_emp': r['rfc'] in cemp,
        'profesional': r['rfc'] in pro,
        'hospedaje':   r['rfc'] in hos,
        'gases':       r['rfc'] in gas,
        'agua':        r['rfc'] in agu,
    } for r in sin]

    def pct(s): return round(len(s & sin_rfcs) / n * 100, 2) if n else 0.0
    stats = {
        'total_imss':      len(imss_records),
        'total_nomina':    len(nom),
        'sin_nomina':      n,
        'pct_cedular_emp': pct(cemp),
        'pct_profesional': pct(pro),
        'pct_hospedaje':   pct(hos),
        'pct_gases':       pct(gas),
        'pct_agua':        pct(agu),
    }
    return rows, stats

def main():
    print('-- Auditoria Fiscal --')
    print('> Listando archivos...')
    files = list_files()
    if not files:
        print('X Carpeta vacia o no publica'); sys.exit(1)
    for f in files: print(f"  . {f['name']} -> {get_type(f['name'])}")

    imss_records = []
    rfc_sets = {}

    for f in files:
        t = get_type(f['name'])
        print(f"> Descargando {f['name']}...")
        try:
            content = download(f['id'])
            if t == 'imss':
                imss_records = extract_imss(content)
                print(f"  {len(imss_records):,} registros IMSS")
            else:
                rfcs = extract_rfcs(content)
                rfc_sets[t] = rfcs
                print(f"  {len(rfcs):,} RFCs")
        except Exception as e:
            print(f'  Error: {e}')
            if t == 'imss': imss_records = []
            else: rfc_sets[t] = set()

    print('> Calculando cruce...')
    rows, stats = compute(imss_records, rfc_sets)
    print(f"  IMSS:{stats['total_imss']:,} | Nomina:{stats['total_nomina']:,} | Sin Nomina:{stats['sin_nomina']:,}")

    out = {'updated': datetime.now(timezone.utc).isoformat(), 'stats': stats, 'rows': rows}
    with open(OUT_FILE, 'w', encoding='utf-8') as fh:
        json.dump(out, fh, ensure_ascii=False, separators=(',', ':'))
    print(f'OK {OUT_FILE} escrito ({len(rows):,} registros)')

if __name__ == '__main__': main()
