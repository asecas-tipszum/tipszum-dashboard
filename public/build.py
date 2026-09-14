import pandas as pd, json, re, numpy as np
from datetime import datetime

# ============================================================
#  build.py  v2 ·  Meta Ads (USD) x PremiumPay (EUR) -> data.json
#  Uso:  python build.py [--fx 0.8673]
#  Pon los archivos en la carpeta ./entrada :
#     - cualquier .xlsx  = export de Meta Ads
#     - cualquier .json  = extraccion de PremiumPay (extractor_FINAL.js)
#  Salida: data.json en esta misma carpeta (UTF-8)
#
#  CAMBIOS v2 respecto a v1:
#   1. data.json se escribe en UTF-8 (antes cp1252: rompia €, — y acentos)
#   2. Paises normalizados (LAT y LATAM eran dos mercados distintos)
#   3. Nombres de tipster desambiguados (habia dos "Tipster Verde Free")
#   4. pais se rellena desde la clave del conjunto cuando Meta no aporta fila
#   5. estado ACTIVO/PAUSADO leido de la columna 'Ad set delivery' (opcional)
#   6. nombres reales de conjunto y campana por grupo (opcional)
#   7. totales: gasto atribuible vs gasto total real (incl. tipsters sin PP)
#   8. ultimo dia con gasto por grupo -> deteccion de conjuntos dormidos
# ============================================================
import sys, glob, os
from datetime import date

FX = 0.8673                    # USD -> EUR (Wise). Editable tambien en el dashboard.
FX_FECHA = date.today().strftime('%d/%m/%Y')
if '--fx' in sys.argv: FX = float(sys.argv[sys.argv.index('--fx') + 1])

# ---------- normalizacion de paises ----------
# Anade aqui cualquier variante nueva que aparezca. La clave es en MAYUSCULAS.
PAIS_ALIAS = {
    'LAT': 'LATAM', 'LATM': 'LATAM', 'LATINOAMERICA': 'LATAM',
    'ES': 'ESP', 'ESPANA': 'ESP', 'ESPAÑA': 'ESP',
    'US': 'USA', 'EEUU': 'USA',
    'UK': 'GRB', 'GB': 'GRB',
    'CA': 'CAN', 'CAN.': 'CAN',
}
PAIS_VACIO = '—'

# ---------- nombres de tipster ----------
# Sobrescribe aqui el nombre visible de cualquier prefijo. Evita ambiguedades.
TIPSTER_ALIAS = {
    'T.VERDE':  'Tipster Verde (ES)',
    'T.GREEN':  'Tipster Verde (EN)',
    'SB':       'Surebet',
}

# ---------- corte de fuente por tipster ----------
# Estos tipsters migraron de PremiumPay a Quanty. El historico de las dos
# fuentes convive en la carpeta ./entrada, asi que hay que quedarse con UNA por
# cada tramo de fechas o las entradas se cuentan dos veces.
#   antes de la fecha -> manda PremiumPay   |   desde la fecha -> manda Quanty
# ---------- prefijo forzado por canal ----------
# Cuando los enlaces de un canal NO llevan el nombre del conjunto de Meta, no se
# puede deducir el prefijo del tipster: el canal "Surebet" de Quanty, por
# ejemplo, tiene los 6 enlaces con nombre libre ("Surebet Free ESP", "GRUPO VIP
# SEMANAL"). Sin esto el script inventaria un tipster llamado "SUREBET FREE ESP"
# y el gasto real de SB quedaria como no atribuible.
# Sus entradas cuentan para el tipster, pero caen en "Generico (sin conjunto)":
# no hay CPL por conjunto hasta que Quanty exponga meta_adset_id.
CANAL_PREF = {
    'Surebet':       'SB',
    'Tipster Verde': 'T.VERDE',   # respaldo: sus enlaces se resuelven uno a uno
}

CORTE = {
    'T.VERDE': ('2026-08-22', 'quanty'),
    'T.GREEN': ('2026-08-22', 'quanty'),
    'SB':      ('2026-08-22', 'quanty'),
}

# ---------- alias de PREFIJO (une Meta <-> PremiumPay) ----------
# Cuando el mismo tipster usa un prefijo en los conjuntos de Meta y otro distinto
# en los nombres de enlace de PremiumPay, mapealo aqui a UNA clave canonica.
# Clave = como aparece escrito (MAYUSCULAS) | Valor = clave canonica (la de config.json).
PREF_ALIAS = {
    'POWER':       'POWER',
    'POWERBETS':   'POWER',
    'POWER BETS':  'POWER',
    'ENLACE TKTK': 'POWER',
    # --- Quanty (sep 2026) ---------------------------------------------
    # Quanty nombra los conjuntos con prefijo corto; el config.json y el export
    # de Meta usan el largo. Se unifican aqui a la clave canonica del config.
    'VERDE':       'T.VERDE',
    'GREEN':       'T.GREEN',
    'SBFREE':      'SB',
    'SUREBET':     'SB',
}
def norm_pref(p):
    p = str(p or '').strip().upper()
    return PREF_ALIAS.get(p, p)

def _sid_norm(v):
    """ID de Meta como texto limpio. Nunca via float: 18 digitos no caben exactos."""
    t = str(v).strip()
    if t in ('', 'nan', 'None', 'NaT', '<NA>'): return ''
    if t.endswith('.0'): t = t[:-2]
    return t

ENT = 'entrada'
xl = sorted(glob.glob(os.path.join(ENT, '*.xlsx'))) + sorted(glob.glob(os.path.join(ENT, '*.xls')))
js = sorted(glob.glob(os.path.join(ENT, '*.json')))
if not xl or not js:
    sys.exit(f'Faltan archivos en ./{ENT}/  (encontrados: {len(xl)} Excel, {len(js)} JSON)')
print(f'Meta: {len(xl)} archivo(s) | PremiumPay: {len(js)} archivo(s) | FX {FX}')

# ── Excel de Meta: cada cuenta factura en su divisa ───────────────────────
# La columna viene como "Amount spent (USD)" o "Amount spent (EUR)". Se
# normaliza a `gasto` + `moneda` y la conversion se hace al final, fila a fila:
# asi conviven una cuenta en dolares y otra en euros sin inventar un FX medio.
SPEND = re.compile(r'^Amount spent \(([A-Za-z]{3})\)$')
_partes = []
for f in xl:
    # Los IDs de Meta tienen 18 digitos: pandas los leeria como float y perderia
    # los ultimos digitos (un float64 solo guarda 15-16 cifras exactas). Hay que
    # forzarlos a texto ANTES de leer el archivo entero.
    _cab = pd.read_excel(f, nrows=0).columns
    _idc = [c for c in _cab if str(c).strip().lower() in
            ('ad set id', 'adset id', 'ad id', 'campaign id',
             'id del conjunto de anuncios', 'id de la campaña', 'id del anuncio')]
    _d = pd.read_excel(f, dtype={c: str for c in _idc})
    _col = next((c for c in _d.columns if SPEND.match(str(c).strip())), None)
    if not _col:
        sys.exit(f'{os.path.basename(f)}: no encuentro la columna de gasto '
                 f'("Amount spent (USD)" o "(EUR)"). Columnas: {list(_d.columns)[:12]}')
    _mon = SPEND.match(str(_col).strip()).group(1).upper()
    _d = _d.rename(columns={_col: 'gasto'}); _d['moneda'] = _mon
    print(f'  {os.path.basename(f)}: {len(_d)} filas en {_mon}')
    _partes.append(_d)
df = pd.concat(_partes, ignore_index=True)
print('monedas en el export:', sorted(df.moneda.unique()), '- el FX solo se aplica a USD')

def eur(sub):
    """Suma en EUR de filas de Meta, respetando la moneda de cada cuenta."""
    if not len(sub): return 0.0
    f = np.where(sub['moneda'].astype(str).str.upper() == 'EUR', 1.0, FX)
    return float((pd.to_numeric(sub['gasto'], errors='coerce').fillna(0) * f).sum())

# ── Subtotales del modo "Pivot table" ──────────────────────────────────────
# Ads Reporting agrupado mete filas de subtotal marcadas como "All". Si llegan
# hasta aqui, el gasto y las impresiones se cuentan dos veces. Va antes de
# convertir 'Day' a fecha, porque pd.to_datetime('All') reventaria.
TODO={'All','Todos','Todas','Total','Totals'}
_n0=len(df)
COL_PLAT=next((c for c in ['Platform','Publisher platform','Plataforma'] if c in df.columns),None)
for _c in [x for x in ['Day','Placement',COL_PLAT] if x and x in df.columns]:
    df=df[~df[_c].astype(str).str.strip().isin(TODO)]
df=df.reset_index(drop=True)
if len(df)<_n0: print(f'subtotales descartados: {_n0-len(df)} filas "All" (el export venia agrupado)')
subs = []; canales = []
pend=[]
for f in js:
    d = json.load(open(f, encoding='utf-8'))
    _sus = d.get('suscriptores', []) or []
    _can = d.get('canales', []) or []
    # Quanty exporta ademas su propio 'daily' ya construido. Solo usamos su
    # parte cruda (canales + suscriptores): el cruce con Meta, el FX, el embudo
    # y el placement los hace este script, que es la unica fuente de verdad.
    _e0 = (_can[0].get('enlaces') or [{}])[0] if _can else {}
    es_q = ('daily' in d) or (d.get('version') == 3) or ('meta_campaign_id' in _e0)
    fuente = 'quanty' if es_q else 'premiumpay'
    for r in _sus: r['fuente'] = fuente
    subs += _sus; canales += _can
    pend += d.get('pendientes', []) or []
    print(f'  {os.path.basename(f)}: {fuente} - {len(_can)} canales, {len(_sus)} suscriptores'
          + (' (ignoro su daily/avisos/totales)' if es_q else ''))
s = pd.DataFrame(subs)
if 'fuente' not in s.columns: s['fuente'] = 'premiumpay'

# ---------- clave canonica ----------
RUIDO = {'PRO', 'INTERESES', 'PROINTERESES'}
MES = re.compile(r'^(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\d{4}$')
def toks(n):
    t = [x.strip().upper() for x in str(n).split('_') if x.strip()]
    t = [x for x in t if x not in RUIDO and not MES.match(x)]
    if t: t[0] = norm_pref(t[0])
    return t
def clave(n):
    return '_'.join(toks(n))
def campos(n):
    t = toks(n)
    return dict(tipster=t[0] if t else '?', pais=t[1] if len(t) > 1 else '?',
                segmento='_'.join(t[2:-1]) if len(t) > 3 else (t[2] if len(t) > 2 else '?'))
def norm_pais(p):
    p = str(p or '').strip().upper()
    if not p or p in {'?', 'NAN', 'NONE', '—', '-'}: return PAIS_VACIO
    return PAIS_ALIAS.get(p, p)

# ---------- Meta ----------
NUM = ['Results', 'Reach', 'gasto', 'Impressions', 'Link clicks', 'Clicks (all)', 'Landing page views']
falta = [c for c in NUM if c not in df.columns]
if falta:
    sys.exit('Faltan columnas obligatorias en el export de Meta: ' + ', '.join(falta))
for c in NUM: df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

# columnas opcionales: si estan, se usan; si no, el pipeline sigue
# 'Delivery status' es como lo llama el Raw Data Report; 'Ad set delivery' el informe
# clasico de Ads Reporting. Los valores (active, not_delivering, rejected...) son los
# mismos y el panel ya los entiende.
COL_EST = next((c for c in ['Ad set delivery', 'Ad Set Delivery', 'Delivery status',
                            'Delivery', 'Entrega'] if c in df.columns), None)
COL_CAM = next((c for c in ['Campaign name', 'Campaign Name', 'Nombre de la campaña'] if c in df.columns), None)
COL_FRQ = next((c for c in ['Frequency', 'Frecuencia'] if c in df.columns), None)
# ID del conjunto: permite cruzar por identificador en vez de por nombre (Quanty v3).
COL_ASID = next((c for c in ['Ad set ID', 'Ad Set ID', 'Ad set id', 'Adset ID',
                             'ID del conjunto de anuncios'] if c in df.columns), None)
# ID del anuncio: hace falta para repartir el gasto por plantilla creativa, porque
# la plantilla es del ANUNCIO, no del conjunto. Sin el, la pestana de Creatividad
# muestra entradas, pagos e ingresos por plantilla, pero no CPL ni ROAS.
COL_ADID = next((c for c in ['Ad ID', 'Ad Id', 'Ad id', 'ID del anuncio']
                 if c in df.columns), None)
if COL_FRQ: df[COL_FRQ] = pd.to_numeric(df[COL_FRQ], errors='coerce').fillna(0)
print('Columnas opcionales -> estado:', COL_EST or 'NO', '| campaña:', COL_CAM or 'NO',
      '| frecuencia:', COL_FRQ or 'NO', '| id de conjunto:', COL_ASID or 'NO',
      '| id de anuncio:', COL_ADID or 'NO')

df['dia'] = pd.to_datetime(df['Day']).dt.strftime('%Y-%m-%d')
df['pref'] = df['Ad set name'].str.split('_').str[0].apply(norm_pref)
df['clave'] = df['Ad set name'].apply(clave)
df['pais'] = df['Ad set name'].apply(lambda x: norm_pais(campos(x)['pais']))
PLAC = {'Feed': 'Feed', 'Instagram Reels': 'IG Reels', 'Facebook Reels': 'FB Reels', 'Instagram Stories': 'IG Stories',
        'Facebook Stories': 'FB Stories', 'Facebook profile feed': 'FB perfil', 'Marketplace': 'Marketplace',
        'Instagram search results': 'IG búsqueda', 'Search results': 'Búsqueda', 'In-stream reels': 'In-stream',
        'Rewarded video': 'Vídeo recompensado', 'Native, banner & interstitial': 'Audience Network', 'Unknown': 'Desconocido'}
PLAT={'facebook':'FB','instagram':'IG','messenger':'MSG','fb':'FB','ig':'IG'}
def etiqueta(r):
    p=PLAC.get(r['Placement'],r['Placement'])
    p='Desconocido' if p is None or (isinstance(p,float) and pd.isna(p)) or not str(p).strip() else str(p)
    p=re.sub(r'^Facebook ','FB ',p); p=re.sub(r'^Instagram ','IG ',p)
    if not COL_PLAT: return p
    pre=PLAT.get(str(r[COL_PLAT]).strip().lower())
    if not pre or p.startswith(('FB ','IG ','MSG ')) or p in ('Audience Network','Desconocido'): return p
    return pre+' '+p
df['placement']=df.apply(etiqueta,axis=1)
print('ubicaciones:',sorted(df.placement.unique()))
if not COL_PLAT:
    print('AVISO: el export no trae columna de plataforma - "Feed" queda sin separar entre Facebook e Instagram.')
    print('       -> Reexporta desde Ads Reporting con el desglose Platform.')
elif not any(str(v).strip().lower() in PLAT for v in df[COL_PLAT].unique()):
    _vals=sorted({str(v).strip() for v in df[COL_PLAT].dropna().unique()})[:8]
    print(f'AVISO: la columna "{COL_PLAT}" existe pero no trae plataformas reconocibles: {_vals}')
    print('       -> "Feed" queda sin separar. El export es de antes de activar el desglose Platform.')

# ---------- PremiumPay ----------
s['imp'] = s.pagos_importe.apply(lambda x: float(str(x).replace('.', '').replace(',', '.')) if str(x).strip() else 0.0)
s['pag'] = pd.to_numeric(s.pagos_num, errors='coerce').fillna(0).astype(int)
s['fe'] = pd.to_datetime(s.fecha_entrada); s['fs'] = pd.to_datetime(s.fecha_salida, errors='coerce')
s['dia'] = s.fe.dt.strftime('%Y-%m-%d')
s['perm'] = (s.fs - s.fe).dt.total_seconds() / 86400
s['clave'] = s.nombre_enlace.apply(clave)

# ── Cruce por ID de conjunto (Quanty v3) ──────────────────────────────────
# Quanty expone meta_adset_id por suscriptor. Si el export de Meta trae la
# columna "Ad set ID", el cruce deja de depender del nombre: se acabaron los
# renombrados y los fallos por un espacio o un sufijo de fecha distinto.
# Solo pisa la clave cuando el ID existe en Meta; si no, se queda la del nombre.
N_ID = 0
if COL_ASID and 'meta_adset_id' in s.columns:
    _map = (df.assign(_id=df[COL_ASID].map(_sid_norm))
              .query('_id != ""').groupby('_id')['clave'].first().to_dict())
    _sid = s.meta_adset_id.map(_sid_norm)
    _hay = _sid != ''
    _nueva = _sid.where(_hay).map(_map)
    N_ID = int(_nueva.notna().sum())
    _dif = int((_nueva.notna() & (_nueva != s.clave)).sum())
    s['clave'] = _nueva.fillna(s.clave)
    print(f'cruce por ID de conjunto: {N_ID} entradas casadas por meta_adset_id '
          f'({_dif} que el nombre no habria emparejado)')
elif 'meta_adset_id' in s.columns:
    _n = int(s.meta_adset_id.notna().sum())
    print(f'cruce por ID: el export de Meta no trae la columna "Ad set ID" - {_n} entradas '
          f'de Quanty traen meta_adset_id y no se puede aprovechar. Anadela al informe guardado.')

# ── Mercado: el pais que declara Quanty ───────────────────────────────────
# Verificado contra 772 entradas: coincide al 100% con el pais del nombre del
# conjunto y ya viene en codigos canonicos (ESP, USA, CAN, GRB, SLV, AUS).
# Su valor esta en las entradas SIN conjunto: hoy caen en "—" y desaparecen del
# filtro de pais; con esto Surebet pasa a tener CPL por mercado.
if 'mercado' in s.columns:
    def _mkt(x):
        v = norm_pais(x)
        return '' if v == PAIS_VACIO else v
    s['mkt'] = s.mercado.apply(_mkt)
    print('mercados declarados por Quanty:', dict(s[s.mkt != ''].mkt.value_counts()))
else:
    s['mkt'] = ''

PREF_META = set(df.pref.dropna().unique())   # prefijos que existen en el export de Meta

def pref_de(sub, validos=None):
    """Prefijo del tipster = primer campo del nombre de enlace.

    Si se pasan `validos` (los prefijos presentes en Meta), se elige el mas
    frecuente DE ESOS. Asi un enlace interno muy usado que no tiene campana
    (p.ej. 'IKER POWER') no le roba el puesto al prefijo real ('GONZALOAST').
    Solo si ninguno aparece en Meta se cae al mas frecuente global."""
    c = {}
    for n in sub.nombre_enlace:
        p = norm_pref(str(n).split('_')[0])
        # Un prefijo real nunca lleva espacios: "Surebet Free ESP" es el nombre
        # de un canal, no un prefijo. (PREF_ALIAS ya normalizo "POWER BETS".)
        # Se acepta un prefijo corto (2 letras, "SB") solo si Meta lo conoce:
        # asi no se cuela ruido pero tampoco se pierde un tipster real.
        corto_ok = len(p) > 2 or (validos and p in validos)
        if corto_ok and ' ' not in p and not p.startswith(('GEN', 'FB', 'ESPA', 'MEX', 'MÉX')): c[p] = c.get(p, 0) + 1
    if not c: return '?'
    en_meta = {k: v for k, v in c.items() if validos and k in validos}
    fuente = en_meta or c
    elegido = max(fuente, key=fuente.get)
    top_global = max(c, key=c.get)
    if elegido != top_global:
        print(f'  prefijo por Meta: "{elegido}" ({c[elegido]} enlaces) en vez de "{top_global}" ({c[top_global]}) - sin campanas en Meta')
    return elegido

NOMBRE = {}; pf_map = {}
_dueno = {}
for canal, sub in s.groupby('tipster'):
    p = CANAL_PREF.get(str(canal).strip()) or pref_de(sub, PREF_META)
    pf_map[canal] = p
    if p == '?':
        print(f'AVISO: no se puede deducir el prefijo del canal "{canal}" ({len(sub)} entradas). '
              f'Anadelo a CANAL_PREF o sus entradas no contaran para ningun tipster.')
    if p in _dueno and p != '?':
        if p in CORTE:
            print(f'  "{p}": convive en "{_dueno[p]}" y "{canal}" - lo resuelve el corte de {CORTE[p][0]}')
        else:
            print(f'AVISO: colision de prefijo "{p}": los canales "{_dueno[p]}" y "{canal}" comparten prefijo - uno pisa al otro.')
    _dueno[p] = canal
    limpio = str(canal).replace('Publicidad Tipszum - ', '').replace('Publicidad Tipzum - ', '') \
                       .replace('-tipszum', '').replace(' - Tipszum', '').replace(' - TIPSZUM', '').strip()
    NOMBRE[p] = TIPSTER_ALIAS.get(p, limpio)

# --- desambiguar nombres visibles repetidos (case-insensitive) ---
vistos = {}
for p, n in list(NOMBRE.items()):
    k = n.strip().lower()
    vistos.setdefault(k, []).append(p)
for k, ps in vistos.items():
    if len(ps) > 1:
        for p in ps:
            if p not in TIPSTER_ALIAS:
                NOMBRE[p] = f'{NOMBRE[p]} [{p}]'
        print('AVISO: nombre de tipster duplicado, desambiguado ->', [NOMBRE[p] for p in ps])

s['pref'] = s.tipster.map(pf_map)

# ── Quanty: un solo canal puede contener varios tipsters ──────────────────
# En PremiumPay cada tipster tiene su canal (T.VERDE en el 698, T.GREEN en el
# 955). Quanty los junta en un unico canal "Tipster Verde", asi que asignar el
# prefijo por canal -como se hace arriba- los fundiria en uno. Para las filas de
# Quanty el prefijo se decide enlace a enlace, y solo entre prefijos que existen
# en Meta: un enlace con nombre libre ("Surebet Free ESP") se queda en el
# dominante de su canal, exactamente como antes.
# Se aplica SOLO a Quanty: el comportamiento con PremiumPay no cambia.
if (s.fuente == 'quanty').any():
    _cand = set(PREF_META) | set(NOMBRE)
    _m = s.fuente == 'quanty'
    def _pref_q(nombre, canal):
        p = norm_pref(str(nombre).split('_')[0])
        return p if p in _cand else pf_map.get(canal, '?')
    s.loc[_m, 'pref'] = [_pref_q(n, t) for n, t in zip(s.loc[_m].nombre_enlace, s.loc[_m].tipster)]
    for p in sorted(set(s.loc[_m, 'pref']) - set(NOMBRE)):
        if p == '?': continue
        NOMBRE[p] = TIPSTER_ALIAS.get(p, p)
        print(f'  quanty: alta de tipster "{p}" -> {NOMBRE[p]}')
    print('  quanty: entradas por prefijo ->', dict(s.loc[_m, 'pref'].value_counts()))

# ── Corte de fuente por tipster ───────────────────────────────────────────
# Sin esto, un tipster con historico en las dos plataformas cuenta sus entradas
# dos veces en el tramo solapado.
if CORTE:
    _n0 = len(s); _drop = pd.Series(False, index=s.index)
    for _p, (_fec, _nueva) in CORTE.items():
        _en = s.pref == _p
        if not _en.any(): continue
        _mal_n = _en & (s.dia >= _fec) & (s.fuente != _nueva)
        _mal_v = _en & (s.dia <  _fec) & (s.fuente == _nueva)
        if _mal_n.any() or _mal_v.any():
            print(f'  corte {_p} desde {_fec}: fuera {int(_mal_n.sum())} entradas de la fuente '
                  f'antigua posteriores al corte y {int(_mal_v.sum())} de {_nueva} anteriores')
        _drop |= _mal_n | _mal_v
    if _drop.any():
        s = s[~_drop].reset_index(drop=True)
        print(f'corte de fuentes: {_n0} -> {len(s)} entradas (evita duplicar el tramo solapado)')

s['tname'] = s.pref.map(NOMBRE)
print('Tipsters detectados:', NOMBRE)

# ---------- emparejar POR TIPSTER ----------
PREFS = list(NOMBRE.keys())
metaT = df[df.pref.isin(PREFS)].copy()
pares = {}; avisos = []
for p in PREFS:
    mk = set(metaT[metaT.pref == p].clave); pk = set(s[s.pref == p].clave)
    for k in mk & pk: pares[(p, k)] = 'exacto'
    for k in mk - pk:
        sub = metaT[(metaT.pref == p) & (metaT.clave == k)]
        avisos.append(dict(tipo='Conjunto de Meta sin enlace en PremiumPay', tipster=NOMBRE[p],
                           valor=sub['Ad set name'].iloc[0],
                           detalle=f"{eur(sub):.2f} € invertidos sin entradas atribuibles"))
    for k in pk - mk:
        sub = s[(s.pref == p) & (s.clave == k)]
        avisos.append(dict(tipo='Enlace de PremiumPay sin conjunto en Meta', tipster=NOMBRE[p],
                           valor=sub.nombre_enlace.iloc[0],
                           detalle=f"{len(sub)} entradas, {sub.pag.sum()} pagos, sin inversión asociada"))
allk = {}
for p in PREFS:
    for k in set(metaT[metaT.pref == p].clave): allk.setdefault(k, []).append(p)
COLIS = sum(1 for k, v in allk.items() if len(v) > 1)

SIN = 'Genérico (sin conjunto)'
metaT['ok'] = [(r.pref, r.clave) in pares for r in metaT.itertuples()]
s['ok'] = [(r.pref, r.clave) in pares for r in s.itertuples()]
s['grupo'] = np.where(s.ok, s.clave, SIN)
metaT['grupo'] = metaT.clave   # meta sin pareja mantiene su clave (gasto sin entradas)

# ---------- metadatos por grupo: estado, campana, nombres reales, ultimo gasto ----------
EST_ACT = {'active', 'activo', 'delivering', 'entregando', 'learning', 'aprendizaje',
           'limited', 'active (learning)', 'recently completed'}
# OJO: la comprobacion es por subcadena, y 'not_delivering' CONTIENE 'delivering'.
# Sin esta lista, un conjunto que no entrega se marcaba ACTIVO. Los negativos mandan.
EST_NO = {'not_delivering', 'inactive', 'paused', 'adset_paused', 'campaign_paused',
          'archived', 'deleted', 'rejected', 'recently_rejected', 'pausado', 'no_entrega'}
def _activo(v):
    v = str(v).strip().lower()
    if v in EST_NO or v.startswith('not_') or 'reject' in v or 'paus' in v: return False
    return any(a in v for a in EST_ACT)
meta_grupo = {}
for gk, gg in metaT.groupby('grupo'):
    con_gasto = gg[gg['gasto'] > 0]
    est = None
    if COL_EST:
        vals = {str(v).strip().lower() for v in gg[COL_EST].dropna().unique()}
        est = 'ACTIVO' if any(_activo(v) for v in vals) else 'PAUSADO'
    meta_grupo[gk] = dict(
        estado=est,
        estado_meta=sorted({str(v).strip() for v in gg[COL_EST].dropna().unique()}) if COL_EST else [],
        campana=sorted({str(v).strip() for v in gg[COL_CAM].dropna().unique()})[:3] if COL_CAM else [],
        adsets=sorted({str(v).strip() for v in gg['Ad set name'].dropna().unique()}),
        ultimo_gasto=(con_gasto.dia.max() if len(con_gasto) else None),
        dias_con_gasto=int(con_gasto.dia.nunique()),
        frecuencia=(float(gg[COL_FRQ].max()) if COL_FRQ else None),
    )

# ---------- diario por tipster x grupo ----------
AGG = dict(gasto_usd=('gasto', 'sum'), impresiones=('Impressions', 'sum'), alcance=('Reach', 'sum'),
           clics=('Link clicks', 'sum'), clics_all=('Clicks (all)', 'sum'), lpv=('Landing page views', 'sum'),
           leads=('Results', 'sum'))
gas = metaT.groupby(['dia', 'pref', 'grupo', 'pais', 'moneda']).agg(**AGG).reset_index()
# El mercado solo se usa para partir las filas SIN conjunto: las que si tienen
# conjunto ya sacan el pais del nombre, y asi el agrupado de siempre no cambia.
s['mkt_g'] = np.where(s.grupo == SIN, s.mkt, '')
gas['mkt_g'] = ''
_K = ['dia', 'pref', 'grupo', 'mkt_g']
ent = s.groupby(_K).agg(entradas=('pag', 'size'), pagadores=('pag', lambda x: (x > 0).sum()),
                        pagos=('pag', 'sum'), ingresos=('imp', 'sum')).reset_index()
baj = (s.dropna(subset=['fs']).assign(dd=lambda x: x.fs.dt.strftime('%Y-%m-%d'))
       .groupby(['dd', 'pref', 'grupo', 'mkt_g']).size().rename('bajas').reset_index().rename(columns={'dd': 'dia'}))
D = gas.merge(ent, on=_K, how='outer').merge(baj, on=_K, how='outer')
for c in ['gasto_usd', 'impresiones', 'alcance', 'clics', 'clics_all', 'lpv', 'leads',
          'entradas', 'pagadores', 'pagos', 'ingresos', 'bajas']:
    D[c] = pd.to_numeric(D[c], errors='coerce').fillna(0)
# Filas que solo vienen de PremiumPay/Quanty no traen moneda: no tienen gasto,
# asi que el valor es indiferente, pero el panel espera el campo siempre.
D['moneda'] = D['moneda'].fillna('USD') if 'moneda' in D.columns else 'USD'

# FIX: el pais se deduce de la clave del conjunto cuando la fila viene solo de PremiumPay.
# Antes quedaba a NaN -> '—' y esas entradas desaparecian al filtrar por pais.
def _pais_fila(r):
    if pd.notna(r['pais']) and str(r['pais']).strip():
        return norm_pais(r['pais'])
    if r['grupo'] != SIN:
        return norm_pais(campos(r['grupo'])['pais'])
    # Sin conjunto: manda el mercado que declara Quanty. Sin el, queda sin pais.
    return norm_pais(r['mkt_g']) if str(r.get('mkt_g') or '').strip() else PAIS_VACIO
D['pais'] = D.apply(_pais_fila, axis=1)
D = D.drop(columns=['mkt_g'])
D['tipster'] = D.pref.map(NOMBRE)
D = D.sort_values(['dia', 'pref', 'grupo'])

# ── Creatividad: plantilla, audiencia y segmentacion (Quanty v3) ──────────
# Quanty declara por suscriptor con que PLANTILLA creativa se hizo el anuncio que
# lo trajo, mas el tipo de audiencia y la segmentacion. Eso permite medir el
# creativo con entradas reales en vez de con clics.
# Ojo con el gasto: la plantilla es del ANUNCIO, no del conjunto. El export de
# Meta a nivel de conjunto no se puede repartir entre plantillas sin inventar,
# asi que el gasto solo se calcula si viene la columna "Ad ID".
CRE = []
TIENE_PLANT = 'plantilla_nombre' in s.columns and s.plantilla_nombre.notna().any()
CRE_GASTO = False
if TIENE_PLANT:
    def _txt(col):
        if col not in s.columns: return pd.Series([''] * len(s), index=s.index)
        return (s[col].fillna('').astype(str).str.strip()
                .replace({'None': '', 'nan': ''}))
    s['_pl'] = _txt('plantilla_nombre'); s['_au'] = _txt('tipo_audiencia'); s['_sg'] = _txt('tipo_segmentacion')
    _KC = ['dia', 'pref', '_pl', '_au', '_sg', 'mkt']
    _ce = s[s._pl != ''].groupby(_KC).agg(
        entradas=('pag', 'size'), pagadores=('pag', lambda x: (x > 0).sum()),
        pagos=('pag', 'sum'), ingresos=('imp', 'sum')).reset_index()
    _cb = (s[(s._pl != '')].dropna(subset=['fs'])
           .assign(dia=lambda x: x.fs.dt.strftime('%Y-%m-%d'))
           .groupby(_KC).size().rename('bajas').reset_index())
    _C = _ce.merge(_cb, on=_KC, how='outer')

    # Gasto por plantilla, solo con ID de anuncio
    if COL_ADID and 'meta_ad_id' in s.columns:
        _ad2 = {}
        # zip y no itertuples: los namedtuple de pandas renombran las columnas que
        # empiezan por "_" y se pierde el acceso por nombre.
        _sp = s[s._pl != '']
        for _a, _p, _u, _g, _m in zip(_sp.meta_ad_id, _sp._pl, _sp._au, _sp._sg, _sp.mkt):
            _a = _sid_norm(_a)
            if _a: _ad2[_a] = (_p, _u, _g, _m)
        if _ad2:
            _mt = metaT.assign(_aid=metaT[COL_ADID].map(_sid_norm))
            _mt = _mt[_mt._aid.isin(_ad2)].copy()
            if len(_mt):
                _mt['_pl'] = _mt._aid.map(lambda a: _ad2[a][0])
                _mt['_au'] = _mt._aid.map(lambda a: _ad2[a][1])
                _mt['_sg'] = _mt._aid.map(lambda a: _ad2[a][2])
                _mt['mkt'] = _mt._aid.map(lambda a: _ad2[a][3])
                _cg = _mt.groupby(_KC + ['moneda']).agg(
                    gasto_usd=('gasto', 'sum'), impresiones=('Impressions', 'sum'),
                    clics=('Link clicks', 'sum'), lpv=('Landing page views', 'sum'),
                    leads=('Results', 'sum')).reset_index()
                _C = _cg.merge(_C, on=_KC, how='outer')
                CRE_GASTO = True
                print(f'creatividad: gasto repartido por plantilla con {len(_ad2)} anuncios')
    if not CRE_GASTO:
        print('creatividad: sin columna "Ad ID" en el export de Meta - habra entradas, '
              'pagos e ingresos por plantilla, pero no CPL ni ROAS.')
    for _c in ['gasto_usd', 'impresiones', 'clics', 'lpv', 'leads',
               'entradas', 'pagadores', 'pagos', 'ingresos', 'bajas']:
        if _c not in _C.columns: _C[_c] = 0
        _C[_c] = pd.to_numeric(_C[_c], errors='coerce').fillna(0)
    if 'moneda' not in _C.columns: _C['moneda'] = 'USD'
    _C['moneda'] = _C['moneda'].fillna('USD')
    _C = _C.rename(columns={'_pl': 'plantilla', '_au': 'audiencia',
                            '_sg': 'segmentacion', 'mkt': 'pais'})
    _C['pais'] = _C.pais.apply(lambda x: norm_pais(x) if str(x).strip() else PAIS_VACIO)
    _C['tipster'] = _C.pref.map(NOMBRE)
    CRE = _C.to_dict('records')
    _tp = _C.groupby('plantilla').entradas.sum().sort_values(ascending=False)
    print('creatividad: entradas por plantilla ->', {k: int(v) for k, v in _tp.items()})

# placement (solo Meta)
PL = metaT.groupby(['dia', 'pref', 'placement', 'moneda']).agg(
    gasto_usd=('gasto', 'sum'), impresiones=('Impressions', 'sum'),
    clics=('Link clicks', 'sum'), lpv=('Landing page views', 'sum'), leads=('Results', 'sum')).reset_index()
PL['tipster'] = PL.pref.map(NOMBRE)

# retención
ret = []
for p in PREFS:
    ss = s[s.pref == p]; n = len(ss)
    for h in [0, 0.25, 0.5, 1, 2, 3, 5, 7]:
        vivos = n - ((ss.perm < h) & ss.perm.notna()).sum()
        ret.append(dict(tipster=NOMBRE[p], dias=h, pct=round(vivos / n * 100, 2) if n else 0))

# otros tipsters en el export sin datos de PremiumPay
otros = []
for p, gg in df[~df.pref.isin(PREFS)].groupby('pref'):
    otros.append(dict(tipster=p, gasto_eur=round(eur(gg), 2),
                      adsets=int(gg['Ad set name'].nunique()), leads=int(gg['Results'].sum())))

# serie diaria de esos tipsters: no se les puede medir el CPL, pero si controlar el presupuesto
otros_daily = []
if (~df.pref.isin(PREFS)).any():
    od = (df[~df.pref.isin(PREFS)]
          .groupby(['dia', 'pref', 'pais', 'moneda'])
          .agg(gasto_usd=('gasto', 'sum'), impresiones=('Impressions', 'sum'),
               clics=('Link clicks', 'sum'), leads=('Results', 'sum')).reset_index())
    otros_daily = od.to_dict('records')

gasto_atrib = eur(metaT)
gasto_otros = float(sum(o['gasto_eur'] for o in otros))
totales = dict(gasto_atribuible_eur=round(gasto_atrib, 2), gasto_otros_eur=round(gasto_otros, 2),
               gasto_total_eur=round(gasto_atrib + gasto_otros, 2),
               pct_atribuible=round(100 * gasto_atrib / (gasto_atrib + gasto_otros), 1) if (gasto_atrib + gasto_otros) else 0)


# ── Cola de aprobacion ─────────────────────────────────────────────────────
# pendientes = solicitudes - entradas_total, ambas DE POR VIDA del enlace: es una
# foto de la cola de ahora, no una serie por dia. Por eso viaja en su propia lista
# y nunca se suma a `daily`: el panel la usa para el CPL ajustado, no para el CPL.
PEND=[]
if pend:
    _pp=pd.DataFrame(pend)
    for _c in ['solicitudes','entradas_vida','pendientes']:
        if _c not in _pp.columns: _pp[_c]=0
        _pp[_c]=pd.to_numeric(_pp[_c],errors='coerce').fillna(0)
    _known=set(NOMBRE)
    def _prefp(nombre,canal):
        p=norm_pref(str(nombre).split('_')[0])
        if p in _known: return p
        m=s[s.tipster==canal]
        return m.pref.mode().iat[0] if len(m) and m.pref.notna().any() else '?'
    _pp['pref']=[_prefp(n,t) for n,t in zip(_pp.nombre,_pp.tipster)]
    _pp['clave']=_pp.nombre.apply(clave)
    _pp['ok']=[(r.pref,r.clave) in pares for r in _pp.itertuples()]
    _pp['grupo']=np.where(_pp.ok,_pp.clave,SIN)
    _g=_pp.groupby(['pref','grupo']).agg(pendientes=('pendientes','sum'),
        solicitudes=('solicitudes','sum'),entradas_vida=('entradas_vida','sum'),
        enlaces=('nombre','nunique')).reset_index()
    _g['tipster']=_g.pref.map(NOMBRE)
    _g=_g[_g.solicitudes>0]
    PEND=_g.to_dict('records')
    _tp=int(_g.pendientes.sum()); _ts=int(_g.solicitudes.sum())
    print(f'cola de aprobacion: {_tp} pendientes sobre {_ts} solicitudes de por vida ({_tp/_ts*100:.1f}% sin aprobar)' if _ts else 'cola de aprobacion: sin datos')
else:
    print('cola de aprobacion: el .json no trae "pendientes" (extractor antiguo) - el panel lo dira')

out = dict(version=2, fx=FX, fx_fecha=FX_FECHA, fx_fuente='Wise (mid-market)',
           generado=datetime.now().strftime('%Y-%m-%d %H:%M'),
           daily=D.to_dict('records'), placement=PL.to_dict('records'), retencion=ret,
           avisos=avisos, otros=otros, otros_daily=otros_daily, colisiones=COLIS, sin_conjunto=SIN, pendientes=PEND, pendientes_fecha=pd.Timestamp.now().strftime('%Y-%m-%d %H:%M'),
           nombres=NOMBRE, meta_grupo=meta_grupo, totales=totales,
           tiene_estado=bool(COL_EST), tiene_campana=bool(COL_CAM),
           rango=[D.dia.min(), D.dia.max()],
           rango_meta=[metaT.dia.min(), metaT.dia.max()], rango_pp=[s.dia.min(), s.dia.max()],
           monedas=sorted(df.moneda.unique()),
           cruce_por_id=int(N_ID), tiene_adset_id=bool(COL_ASID),
           creativos=CRE, tiene_plantilla=bool(TIENE_PLANT), tiene_gasto_creativo=bool(CRE_GASTO),
           corte={k: list(v) for k, v in CORTE.items()},
           cuadre=dict(entradas=int(len(s)), pagos=int(s.pag.sum()), ingresos=float(s.imp.sum()),
                       gasto_usd=float(metaT.loc[metaT.moneda != 'EUR', 'gasto'].sum()),
                       gasto_eur_nativo=float(metaT.loc[metaT.moneda == 'EUR', 'gasto'].sum()),
                       gasto_total_eur=round(eur(metaT), 2)))

# FIX CRITICO: sin encoding='utf-8' Windows escribe cp1252 y rompe €, — y acentos.
with open('data.json', 'w', encoding='utf-8') as fh:
    json.dump(out, fh, ensure_ascii=False, default=str, separators=(',', ':'))

print(f"data.json escrito ({os.path.getsize('data.json')/1024:.0f} KB, UTF-8)")
print('emparejados:', len(pares), '| avisos:', len(avisos), '| colisiones:', COLIS)
print('cuadre:', out['cuadre'])
print('totales €:', totales)
print('paises:', sorted(D.pais.unique()))
if COL_EST:
    act = [k for k, v in meta_grupo.items() if v['estado'] == 'ACTIVO']
    print(f'estado: {len(act)} conjuntos ACTIVOS de {len(meta_grupo)}')
else:
    print('estado: columna "Ad set delivery" ausente -> el dashboard usara la regla de inactividad')
print('filas D:', len(D), '| placement:', len(PL))
print('otros tipsters en export sin PP:', [o['tipster'] for o in otros])
