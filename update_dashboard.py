"""
update_dashboard.py – Genera data.json con cruce IMSS-Nómina multi-mes.

Detecta automáticamente el mes de cada archivo IMSS por nombre:
  · Nombre en español:  "IMSS Mayo 2026.xlsx"   → 2026-05
  · Fecha DDMMYYYY:     "imss16062026.xlsx"      → 2026-06
  · Fecha YYYYMM:       "imss_202601.xlsx"       → 2026-01

Todos los meses se cruzan contra el mismo padrón de nómina vigente.
Se calcula reincidencia: en cuántos meses aparece cada RFC sin nómina.
"""
import io, json, re, sys, zipfile, requests, openpyxl
from collections import defaultdict
from datetime import datetime, timezone

FOLDER_ID = '1QoOP4jnPmp7_x9tz1Q7l2Yy2JmOx7MXc'
API_KEY   = 'AIzaSyAId7gthv7EEzmaTrfbt07FK4Kf-ii51uM'
BASE_URL  = 'https://www.googleapis.com/drive/v3'
OUT_FILE  = 'data.json'

RFC_RE = re.compile(r'^[A-ZÑ&]{3,4}[0-9]{6}[A-Z0-9]{3}$')

MESES_ES = {
    'enero':1,'febrero':2,'marzo':3,'abril':4,'mayo':5,'junio':6,
    'julio':7,'agosto':8,'septiembre':9,'octubre':10,'noviembre':11,'diciembre':12
}
MESES_NOMBRE = {
    1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',
    7:'Julio',8:'Agosto',9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre'
}

# ── Helpers de mes ────────────────────────────────────────────────────────────

def mes_label(ym):
    """'2026-06' → 'Junio 2026'"""
    try:
        y, m = ym.split('-')
        return f"{MESES_NOMBRE[int(m)]} {y}"
    except Exception:
        return ym

def get_imss_month(name):
    """Extrae YYYY-MM del nombre de archivo IMSS. Devuelve None si no detecta."""
    n = name.lower()
    # Año base (buscar primero para reutilizar)
    ym = re.search(r'(20\d{2})', name)
    year = int(ym.group(1)) if ym else datetime.now().year

    # 1) Nombre de mes en español  → "mayo 2026"
    for mes_name, mes_num in MESES_ES.items():
        if mes_name in n:
            return f"{year}-{mes_num:02d}"

    # 2) Patrón DDMMYYYY: 16062026 → mes=06
    m = re.search(r'\d{2}(\d{2})(20\d{2})', name)
    if m:
        mes = int(m.group(1))
        if 1 <= mes <= 12:
            return f"{m.group(2)}-{m.group(1)}"

    # 3) Patrón YYYYMM o YYYY-MM
    m = re.search(r'(20\d{2})[-_]?(\d{2})(?!\d)', name)
    if m:
        mes = int(m.group(2))
        if 1 <= mes <= 12:
            return f"{m.group(1)}-{m.group(2)}"

    return None

# ── Detección de tipo de archivo ──────────────────────────────────────────────

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

# ── Drive API ─────────────────────────────────────────────────────────────────

def list_files():
    url = (f"{BASE_URL}/files?q='{FOLDER_ID}'+in+parents+and+trashed=false"
           f"&fields=files(id,name)&key={API_KEY}&pageSize=100")
    r = requests.get(url, timeout=30); r.raise_for_status()
    return [f for f in r.json().get('files', []) if get_type(f['name'])]

def download(fid):
    r = requests.get(f"{BASE_URL}/files/{fid}?alt=media&key={API_KEY}", timeout=180)
    r.raise_for_status(); return r.content

# ── Parseo de XLSX ────────────────────────────────────────────────────────────

def find_col(headers, *keywords):
    h = [str(x).lower().strip() if x else '' for x in headers]
    for kw in keywords:
        for i, v in enumerate(h):
            if kw in v:
                return i
    return None

def extract_imss(content):
    """Extrae registros completos (rfc, nombre, direccion, trabajadores) del IMSS."""
    records = []
    try:
        wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        ws = wb.active
        header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None) or []
        print(f'  Columnas IMSS: {[str(h) for h in header[:15]]}')

        rfc_col  = find_col(header, 'rfc') or 0
        nom_col  = find_col(header, 'razon', 'razón', 'razón social', 'razon social', 'nombre',
                            'denominaci', 'patron', 'patrón', 'empresa', 'social', 'entidad', 'contribuyente')
        dir_col  = find_col(header, 'domicilio', 'direcci', 'calle', 'municipio', 'colonia', 'estado')
        trab_col = find_col(header, 'trabajador', 'empleado', 'num. trab', 'número de trab',
                            'num trab', 'ntrab', 'asegurado', 'personal')
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
    """Extrae solo el conjunto de RFC de un padrón."""
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

# ── Fuzzy matching ────────────────────────────────────────────────────────────

def build_fuzzy_index(rfcs):
    """Índice pigeonhole: ≤2 errores → al menos 1 bloque de 4 chars idéntico."""
    idx = defaultdict(set)
    for rfc in rfcs:
        n = len(rfc)
        if n < 9: continue
        mid = n // 3
        idx[(0, rfc[:mid])].add(rfc)
        idx[(1, rfc[mid:2*mid])].add(rfc)
        idx[(2, rfc[2*mid:])].add(rfc)
    return idx

def fuzzy_find(rfc, idx, max_err=2):
    n = len(rfc)
    if n < 9: return None, 99
    mid = n // 3
    candidates = (idx.get((0, rfc[:mid]),     set()) |
                  idx.get((1, rfc[mid:2*mid]), set()) |
                  idx.get((2, rfc[2*mid:]),    set()))
    candidates.discard(rfc)
    best, best_d = None, max_err + 1
    for c in candidates:
        if len(c) != n: continue
        d = sum(a != b for a, b in zip(rfc, c))
        if d < best_d:
            best_d, best = d, c
            if d == 1: break
    return best, best_d

# ── Cruce principal ───────────────────────────────────────────────────────────

def compute(imss_records, rfc_sets):
    """Cruza registros IMSS contra padrones; devuelve (rows, stats)."""
    nom  = rfc_sets.get('nomina',      set())
    cemp = rfc_sets.get('cedular_emp', set())
    ced  = rfc_sets.get('cedular',     set())
    hos  = rfc_sets.get('hospedaje',   set())
    pro  = rfc_sets.get('profesional', set())
    gas  = rfc_sets.get('gases',       set())
    agu  = rfc_sets.get('agua',        set())

    # 1. Deduplicar por RFC (conservar el de más trabajadores)
    dedup = {}
    for r in imss_records:
        rfc = r['rfc']
        if rfc not in dedup:
            dedup[rfc] = dict(r)
        else:
            if (r['trabajadores'] or 0) > (dedup[rfc]['trabajadores'] or 0):
                dedup[rfc] = dict(r)
    imss_uniq = list(dedup.values())
    print(f'  IMSS únicos: {len(imss_uniq):,} (de {len(imss_records):,} totales)')

    # 2. Excluir exactos en nómina
    sin_exacto = [r for r in imss_uniq if r['rfc'] not in nom]

    # 3. Fuzzy: excluir RFC con ≤2 errores tipográficos vs nómina
    print('  Construyendo índice fuzzy...')
    nom_idx = build_fuzzy_index(nom)
    sin, fuzzy_excl = [], 0
    for r in sin_exacto:
        if fuzzy_find(r['rfc'], nom_idx, max_err=2)[0]:
            fuzzy_excl += 1
        else:
            sin.append(r)
    print(f'  Excluidos fuzzy: {fuzzy_excl:,}')

    n        = len(sin)
    sin_rfcs = {r['rfc'] for r in sin}

    rows = [{
        'rfc':          r['rfc'],
        'nombre':       r['nombre'],
        'direccion':    r['direccion'],
        'trabajadores': r['trabajadores'],
        'cedular_emp':  r['rfc'] in cemp,
        'profesional':  r['rfc'] in pro,
        'hospedaje':    r['rfc'] in hos,
        'gases':        r['rfc'] in gas,
        'agua':         r['rfc'] in agu,
        # Reincidencia — se rellena en main() después de procesar todos los meses
        'meses_detectado': [],
        'meses_count':     0,
        'es_nuevo':        True,
    } for r in sin]

    def pct(s): return round(len(s & sin_rfcs) / n * 100, 2) if n else 0.0
    stats = {
        'total_imss':      len(imss_uniq),
        'total_nomina':    len(nom),
        'sin_nomina':      n,
        'excluidos_fuzzy': fuzzy_excl,
        'pct_cedular_emp': pct(cemp),
        'pct_profesional': pct(pro),
        'pct_hospedaje':   pct(hos),
        'pct_gases':       pct(gas),
        'pct_agua':        pct(agu),
        'nuevos':          0,   # se calcula en main()
        'reincidentes':    0,   # se calcula en main()
    }
    return rows, stats

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print('-- Auditoría Fiscal Multi-Mes --')
    print('> Listando archivos en Drive...')
    files = list_files()
    if not files:
        print('X Carpeta vacía o no pública'); sys.exit(1)

    # Separar archivos IMSS (por mes) de los demás padrones
    imss_by_month = {}   # "2026-05" → [file, ...]
    other_files   = []   # [(tipo, file), ...]

    for f in files:
        t = get_type(f['name'])
        if t == 'imss':
            month = get_imss_month(f['name'])
            if month:
                imss_by_month.setdefault(month, []).append(f)
                print(f"  IMSS {mes_label(month)}: {f['name']}")
            else:
                print(f"  WARN sin mes detectado: {f['name']} — ignorado")
        elif t:
            other_files.append((t, f))
            print(f"  Padrón {t}: {f['name']}")

    if not imss_by_month:
        print('X No se encontraron archivos IMSS con mes identificable'); sys.exit(1)

    # Descargar padrones (comunes a todos los meses)
    rfc_sets = {}
    for t, f in other_files:
        print(f"\n> Descargando padrón {t}: {f['name']}...")
        try:
            rfc_sets[t] = extract_rfcs(download(f['id']))
            print(f"  {len(rfc_sets[t]):,} RFCs")
        except Exception as e:
            print(f"  Error: {e}"); rfc_sets[t] = set()

    # Procesar cada mes en orden cronológico
    por_mes = {}
    for month in sorted(imss_by_month.keys()):
        print(f"\n> Procesando {mes_label(month)}...")
        imss_records = []
        for f in imss_by_month[month]:
            try:
                records = extract_imss(download(f['id']))
                imss_records.extend(records)
                print(f"  {len(records):,} registros de «{f['name']}»")
            except Exception as e:
                print(f"  Error en {f['name']}: {e}")

        if not imss_records:
            print(f"  SKIP: sin registros para {mes_label(month)}")
            continue

        rows, stats = compute(imss_records, rfc_sets)
        por_mes[month] = {'label': mes_label(month), 'stats': stats, 'rows': rows}
        print(f"  → Sin nómina: {stats['sin_nomina']:,}")

    if not por_mes:
        print('X No se pudo procesar ningún mes'); sys.exit(1)

    # ── Calcular reincidencia entre todos los meses ──
    # rfc_meses[rfc] = lista ordenada de meses donde aparece sin nómina
    rfc_meses = defaultdict(list)
    for month in sorted(por_mes.keys()):
        for row in por_mes[month]['rows']:
            rfc_meses[row['rfc']].append(month)

    # Enriquecer cada fila con su historial y actualizar stats
    for month, data in por_mes.items():
        nuevos = reincidentes = 0
        for row in data['rows']:
            meses = sorted(rfc_meses[row['rfc']])
            row['meses_detectado'] = meses
            row['meses_count']     = len(meses)
            # es_nuevo = este mes es la primera vez que aparece
            row['es_nuevo'] = (meses[0] == month)
            if row['es_nuevo']:
                nuevos += 1
            else:
                reincidentes += 1
        data['stats']['nuevos']       = nuevos
        data['stats']['reincidentes'] = reincidentes

    meses_list  = sorted(por_mes.keys())
    mes_vigente = meses_list[-1]

    out = {
        'updated':     datetime.now(timezone.utc).isoformat(),
        'meses':       meses_list,
        'mes_vigente': mes_vigente,
        'por_mes':     por_mes,
        # Compatibilidad con versión anterior (mes vigente)
        'stats': por_mes[mes_vigente]['stats'],
        'rows':  por_mes[mes_vigente]['rows'],
    }

    with open(OUT_FILE, 'w', encoding='utf-8') as fh:
        json.dump(out, fh, ensure_ascii=False, separators=(',', ':'))

    labels = ', '.join(mes_label(m) for m in meses_list)
    print(f'\nOK — {OUT_FILE} escrito | {len(meses_list)} meses: {labels}')

if __name__ == '__main__': main()
