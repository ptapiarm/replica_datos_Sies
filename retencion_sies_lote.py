"""
Retención de 1er año de pregrado por lote — réplica de la metodología SIES
(Informe de Retención de 1er Año de Pregrado, cohortes 2021-2025, agosto 2026).

Recibe una carpeta con las bases públicas de Matrícula SIES con MRUN ya descomprimidas
(archivos .csv, también en subcarpetas), detecta el año de cada una y calcula la
retención de 1er año para toda cohorte t que tenga la matrícula t+1 disponible.

Uso:
    python retencion_sies_lote.py --carpeta ./bases --out retencion.xlsx
    python retencion_sies_lote.py --carpeta ./bases --inst 69 86 --out retencion.xlsx
    python retencion_sies_lote.py --carpeta ./bases --acreditacion cna.xls --out retencion.xlsx

    --inst      Instituciones con detalle (por carrera, sexo, tipo de programa y por MRUN).
                Las tablas por institución y del sistema se calculan siempre para todas.
    --excluir   Instituciones a excluir por cohorte, formato AÑO:COD[,COD] (repetible).
                Por defecto 2025:430 (CFT Inacap, como en el informe 2026).
    --acreditacion  (opcional) Excel del Buscador Avanzado de Procesos de Acreditación
                Institucional de la CNA (.xls o .xlsx). Del archivo se usa solo la columna
                "Años" (columna R): se agrega como "Años acreditación (CNA)" por institución.
                Es la acreditación del último proceso informado, no la de cada cohorte.
    --cache     Carpeta donde guardar las bases ya leídas (acelera corridas siguientes).
                Por defecto <carpeta>/_cache_retencion.

Requisitos: pandas y openpyxl.

Reglas  [doc] = explícito en el informe | [sup] = supuesto operativo
  Cohorte t (denominador)
    1. NIVEL_GLOBAL = Pregrado                                              [doc]
    2. TIPO_PLAN_CARR = Plan Regular (excluye Especial y Regular de
       Continuidad)                                                         [doc/sup]
    3. Duración mínima con DUR_ESTUDIO_CARR: 4 sem técnicas/bachillerato/
       plan común; 6 prof. sin lic.; 8 prof. con lic. y lic. no conducente  [doc/sup: campo]
    4. ANIO_ING_CARR_ORI = t, ambos semestres                     [calibrado con Excel SIES]
    5. Unidad = registro (sin deduplicar); los registros sin MRUN quedan en N
       como no retenidos (S)                                      [calibrado con Excel SIES]
    6. Exclusiones puntuales por cohorte (--excluir)                         [doc]
  Situación en t+1 (prioridad A > B > X > C)                                 [doc defs / sup prioridad]
    A Retenido: misma COD_INST y ANIO_ING_CARR_ORI = t, o un código especial
      (1900, 9998, 9999) en vez de año (cualquier carrera)       [códigos: calibrado]
    B Persiste misma inst.: misma COD_INST y ANIO_ING_CARR_ORI = t+1
    C Persiste en ES: otra COD_INST (cualquier nivel)
    D No matrícula: no aparece en t+1
    S Sin MRUN en t: no se puede seguir; cuenta en N como no retenido
    X Inconsistencia: misma inst. con otro año de origen (p. ej. t-2) -> se excluye del
      cálculo; se informa el rango de retención si X contara como no
      retenido (mín) o como retenido (máx)                                   [doc: se eliminan; sup: criterio]
  Publicación: categorías con N >= 10                                        [doc]
"""
import argparse
import unicodedata
import xml.etree.ElementTree as ET
import hashlib
import re
import sys
from pathlib import Path

import pandas as pd

COLS_T = ['mrun', 'gen_alu', 'anio_ing_carr_ori', 'sem_ing_carr_ori', 'anio_ing_carr_act',
          'tipo_inst_1', 'cod_inst', 'nomb_inst', 'nomb_sede', 'cod_carrera', 'nomb_carrera',
          'codigo_unico', 'modalidad', 'jornada', 'tipo_plan_carr', 'dur_estudio_carr',
          'nivel_global', 'nivel_carrera_1', 'nivel_carrera_2', 'area_conocimiento',
          'area_carrera_generica', 'forma_ingreso', 'cat_periodo']
DUR_MIN = {
    'Técnico de Nivel Superior': 4,
    'Bachillerato, Ciclo Inicial o Plan Común': 4,
    'Profesional Sin Licenciatura': 6,
    'Profesional Con Licenciatura': 8,
    'Licenciatura No Conducente a Título': 8,
}
OBLIGATORIAS = ['mrun', 'cod_inst', 'anio_ing_carr_ori', 'sem_ing_carr_ori', 'nivel_global',
                'tipo_plan_carr', 'nivel_carrera_1', 'dur_estudio_carr']
ORDEN = {'A': 0, 'B': 1, 'X': 2, 'C': 3}
SEXO = {'1': 'Hombre', '2': 'Mujer'}
CAT_COLS = ['tipo_inst_1', 'cod_inst', 'nomb_inst', 'nomb_sede', 'nomb_carrera', 'modalidad',
            'jornada', 'tipo_plan_carr', 'nivel_global', 'nivel_carrera_1', 'nivel_carrera_2',
            'area_conocimiento', 'area_carrera_generica', 'forma_ingreso', 'gen_alu']


# ----------------------------------------------------------------------------- lectura
def log(msg):
    print(msg, flush=True)


def leer_csv(path: Path) -> pd.DataFrame:
    with open(path, 'rb') as f:
        cab = f.readline()
    enc = 'utf-8'
    try:
        cab.decode('utf-8')
    except UnicodeDecodeError:
        enc = 'latin-1'
    texto = cab.decode(enc)
    sep = ';' if texto.count(';') >= texto.count(',') else ','
    df = pd.read_csv(path, sep=sep, dtype=str, encoding=enc, encoding_errors='replace',
                     usecols=lambda c: c.strip().lower().lstrip('\ufeff') in COLS_T)
    df.columns = df.columns.str.strip().str.lower().str.lstrip('\ufeff')
    faltan = [c for c in OBLIGATORIAS if c not in df]
    if faltan:
        raise ValueError(f'no parece una base de Matrícula SIES con MRUN (faltan columnas: {", ".join(faltan)})')
    for c in COLS_T:
        if c not in df:
            df[c] = pd.NA
    df = df[COLS_T]
    # normalizar códigos numéricos guardados como texto ("2025.0", " 69 ")
    for c in ['mrun', 'anio_ing_carr_ori', 'sem_ing_carr_ori', 'anio_ing_carr_act',
              'cod_inst', 'cat_periodo', 'gen_alu']:
        df[c] = df[c].str.strip().str.replace(r'\.0$', '', regex=True)
    for c in CAT_COLS:
        df[c] = df[c].astype('category')
    return df


def cargar_bases(carpeta: Path, cache: Path):
    cache.mkdir(parents=True, exist_ok=True)
    bases, archivos = {}, []
    csvs = sorted(p for p in carpeta.rglob('*') if p.is_file() and p.suffix.lower() == '.csv'
                  and cache not in p.parents)
    if not csvs:
        sys.exit(f'No hay archivos .csv en {carpeta}')
    for f in csvs:
        firma = hashlib.md5(f'{f.name}{f.stat().st_size}{f.stat().st_mtime}'.encode()).hexdigest()[:10]
        pkl = cache / f'{f.stem}_{firma}.pkl'
        try:
            if pkl.exists():
                df = pd.read_pickle(pkl)
                origen = 'cache'
            else:
                log(f'  Leyendo {f.name} ...')
                df = leer_csv(f)
                df.to_pickle(pkl)
                origen = 'leído'
            anios = df.cat_periodo.dropna().astype(str)
            if anios.empty:
                m = re.search(r'(20\d{2})', f.name)
                if not m:
                    raise ValueError('no se pudo determinar el año (sin CAT_PERIODO ni año en el nombre)')
                anio = int(m.group(1))
            else:
                anio = int(anios.mode()[0])
            if anio in bases:
                raise ValueError(f'año {anio} repetido (ya cargado desde {bases[anio][1]})')
            bases[anio] = (pkl, f.name)
            archivos.append({'Archivo': f.name, 'Año': anio, 'Registros': len(df), 'Estado': origen})
            log(f'  {f.name}: año {anio}, {len(df):,} registros ({origen})')
            del df
        except Exception as e:  # un archivo malo no detiene el lote
            archivos.append({'Archivo': f.name, 'Año': None, 'Registros': None, 'Estado': f'ERROR: {e}'})
            log(f'  ERROR en {f.name}: {e}')
    return bases, pd.DataFrame(archivos)


# ----------------------------------------------------------------------------- cálculo
def cohorte(df, t, excluir):
    pasos = []
    def paso(d, nombre):
        pasos.append((nombre, len(d)))
        return d
    d = paso(df, f'Registros base {t}')
    d = paso(d[d.nivel_global == 'Pregrado'], 'Pregrado')
    d = paso(d[d.tipo_plan_carr == 'Plan Regular'], 'Plan Regular')
    d = paso(d[d.anio_ing_carr_ori == str(t)], f'1er año: año origen {t} (ambos semestres)')
    dur = pd.to_numeric(d.dur_estudio_carr, errors='coerce')
    minimo = d.nivel_carrera_1.astype(str).map(DUR_MIN)
    d = paso(d[dur >= minimo], 'Duración mínima')
    sin = d.mrun.isna() | (d.mrun == '')
    pasos.append(('  de ellos sin MRUN (quedan en N como no retenidos)', int(sin.sum())))
    if excluir:
        d = paso(d[~d.cod_inst.astype(str).isin(excluir)], f'Excluye inst. {",".join(excluir)}')
    d = d.copy()
    d['gen_alu'] = d.gen_alu.astype(str).map(SEXO).fillna('Sin información')
    return d, pasos


def clasificar(coh, t1df, t):
    """Una fila por registro de la cohorte (unidad: registro, como SIES)."""
    m = t1df[['mrun', 'cod_inst', 'anio_ing_carr_ori', 'nomb_carrera']].copy()
    m.columns = ['mrun', 'inst_t1', 'ori_t1', 'carrera_t1']
    m['inst_t1'] = m.inst_t1.astype(str)
    m = m[m.mrun.notna() & m.mrun.isin(set(coh.mrun.dropna()))]
    claves = coh[['mrun', 'cod_inst']].dropna().assign(cod_inst=lambda d: d.cod_inst.astype(str)).drop_duplicates()
    x = claves.merge(m, on='mrun', how='left')
    misma = x.inst_t1 == x.cod_inst
    ori = pd.to_numeric(x.ori_t1, errors='coerce')
    codigo = misma & ori.notna() & ((ori < 1990) | (ori >= 9000))   # 1900, 9998, 9999: códigos, no años
    x['sit'] = 'D'
    x.loc[x.inst_t1.notna() & ~misma, 'sit'] = 'C'
    x.loc[misma & ~codigo & ~x.ori_t1.isin([str(t), str(t + 1)]), 'sit'] = 'X'
    x.loc[misma & (x.ori_t1 == str(t + 1)), 'sit'] = 'B'
    x.loc[misma & ((x.ori_t1 == str(t)) | codigo), 'sit'] = 'A'
    x['prio'] = x.sit.map(ORDEN).fillna(9)
    x = x.sort_values('prio').drop_duplicates(['mrun', 'cod_inst'])
    out = coh.assign(cod_inst=coh.cod_inst.astype(str)).merge(
        x[['mrun', 'cod_inst', 'sit', 'inst_t1', 'ori_t1', 'carrera_t1']],
        on=['mrun', 'cod_inst'], how='left')
    out.loc[out.mrun.isna() | (out.mrun == ''), 'sit'] = 'S'
    return out


def tabla(df, by):
    by = list(by)
    g = df.groupby(by, observed=True).sit.value_counts().unstack(fill_value=0)
    for s in 'ABCDSX':
        if s not in g:
            g[s] = 0
    g = g[list('ABCDSX')]
    g['N'] = g[list('ABCDS')].sum(axis=1)
    n = g.N.where(g.N > 0)
    g['Retención'] = g.A / n
    g['Persist. misma inst.'] = (g.A + g.B) / n
    g['Persist. ES'] = (g.A + g.B + g.C) / n
    g['No matrícula'] = g.D / n
    tot = g.N + g.X
    g['Retención mín (X no retenidos)'] = g.A / tot
    g['Retención máx (X retenidos)'] = (g.A + g.X) / tot
    g['Publicable (N>=10)'] = g.N >= 10
    g = g[['N', 'A', 'B', 'C', 'D', 'S', 'X', 'Retención', 'Persist. misma inst.', 'Persist. ES',
           'No matrícula', 'Retención mín (X no retenidos)', 'Retención máx (X retenidos)',
           'Publicable (N>=10)']]
    return g.reset_index()



# ----------------------------------------------------------------------------- acreditación CNA
# Nombres que no coinciden entre SIES (COD_INST) y la CNA. Clave: COD_INST; valor: nombre CNA normalizado.
EQUIV_CNA = {
    '42': 'UNIVERSIDAD CATOLICA SILVA HENRIQUEZ',
    '22': 'UNIVERSIDAD VINA DEL MAR',
    '218': 'CFT ENAC',
    '435': 'CFT DEL MEDIO AMBIENTE IDMA',
    '629': 'CORPORACION CENTRO DE FORMACION TECNICA DE LA PONTIFICIA UNIVERSIDAD CATOLICA DE VALPARAISO O CFT PUCV',
    '113': 'IP GALDAMEZ IPG',
    '117': 'INSTITUTO SUPERIOR DE ARTES Y CIENCIAS DE LA COMUNICACION IACC',
    '139': 'IP VIRGINIO GOMEZ',
    '193': 'INSTITUTO NACIONAL DEL FUTBOL DEPORTE Y ACTIVIDAD FISICA INAF',
    '676': 'INSTITUTO INTERNACIONAL DE ARTES CULINARIAS Y SERVICIOS',
}
COL_ACRED = 'Años acreditación (CNA)'


def norm_nombre(s):
    s = unicodedata.normalize('NFD', str(s))
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn').upper()
    s = ' '.join(re.sub(r'[^A-Z0-9 ]', ' ', s).split())
    s = re.sub(r'^CENTRO DE FORMACION TECNICA ', 'CFT ', s)
    s = re.sub(r'^INSTITUTO PROFESIONAL ', 'IP ', s)
    return s


def _leer_xml_2003(path):
    ns = '{urn:schemas-microsoft-com:office:spreadsheet}'
    filas = []
    ws = ET.parse(path).getroot().find(f'{ns}Worksheet')
    for r in ws.iter(f'{ns}Row'):
        fila = []
        for c in r.findall(f'{ns}Cell'):
            idx = c.get(f'{ns}Index')
            if idx:
                fila += [None] * (int(idx) - 1 - len(fila))
            d = c.find(f'{ns}Data')
            fila.append(d.text if d is not None else None)
        filas.append(fila)
    ancho = max(map(len, filas))
    filas = [f + [None] * (ancho - len(f)) for f in filas]
    return pd.DataFrame(filas[1:], columns=filas[0])


def leer_acreditacion(path):
    """Devuelve {nombre normalizado: años} con la columna 'Años' (R) del último proceso."""
    try:
        df = pd.read_excel(path, dtype=str)
    except Exception:
        df = _leer_xml_2003(path)
    cols = [str(c).strip() for c in df.columns]
    i_anios = cols.index('Años') if 'Años' in cols else 17          # columna R
    i_inst = cols.index('Institución') if 'Institución' in cols else 2
    i_fecha = next((i for i, c in enumerate(cols) if 'Fecha Sesi' in c), None)
    a = pd.DataFrame({
        'k': df.iloc[:, i_inst].map(norm_nombre),
        'anios': pd.to_numeric(df.iloc[:, i_anios], errors='coerce'),
        'fecha': pd.to_datetime(df.iloc[:, i_fecha], dayfirst=True, errors='coerce') if i_fecha is not None else pd.NaT,
    })
    a = a.sort_values('fecha').drop_duplicates('k', keep='last')   # si hay varios procesos, el más reciente
    return dict(zip(a.k, a.anios))


def cruzar_acreditacion(pi, mapa):
    insts = pi.sort_values('cohorte').groupby('cod_inst', observed=True).agg(
        nomb_inst=('nomb_inst', 'last'), tipo_inst_1=('tipo_inst_1', 'last')).reset_index()
    insts['cod_inst'] = insts.cod_inst.astype(str)
    insts['Nombre CNA (normalizado)'] = [EQUIV_CNA.get(c, norm_nombre(n)) for c, n in zip(insts.cod_inst, insts.nomb_inst)]
    insts['Encontrada en CNA'] = insts['Nombre CNA (normalizado)'].isin(mapa.keys())
    insts[COL_ACRED] = pd.to_numeric(insts['Nombre CNA (normalizado)'].map(mapa)).astype('Int64')
    return insts


# ----------------------------------------------------------------------------- salida
LEEME = [
    ('Qué es', 'Réplica de la Tasa de Retención de 1er año de pregrado (metodología SIES), calculada '
               'con las bases públicas de Matrícula con MRUN.'),
    ('Cohorte', 'Registros de 1er año del año t: pregrado, Plan Regular, duración mínima, año de '
                'origen = t (ambos semestres). Se cuentan registros, incluidos los sin MRUN.'),
    ('A', 'Retenido: en t+1 sigue en la misma institución y en la misma cohorte (cualquier carrera); '
          'incluye año de origen informado con código especial (1900, 9998, 9999).'),
    ('S', 'Sin MRUN en la base pública: no se puede seguir y cuenta como no retenido. SIES sí los '
          'sigue con el RUN real, por lo que aquí la retención queda levemente subestimada.'),
    ('B', 'Persiste en la misma institución, pero reingresó como 1er año (cohorte t+1).'),
    ('C', 'Persiste en educación superior en otra institución.'),
    ('D', 'No matrícula: no aparece en ninguna institución en t+1.'),
    ('X', 'Inconsistencia: sigue en la misma institución con otro año de origen. Se excluye de N; '
          'las columnas mín/máx muestran la retención si se contara como no retenido o retenido.'),
    ('N', 'A + B + C + D + S (denominador de todas las tasas).'),
    ('Tasas', 'Retención = A/N; Persist. misma inst. = (A+B)/N; Persist. ES = (A+B+C)/N; '
              'No matrícula = D/N.'),
    ('Publicable', 'SIES publica categorías con N >= 10.'),
    ('Supuestos', 'Excluye Plan Regular de Continuidad; duración con DUR_ESTUDIO_CARR; C incluye '
                  'cualquier nivel; prioridad A > B > X > C ante doble matrícula; atributos del año t.'),
    ('Validación', 'Cohorte 2025 contra Excel SIES por IES: 80 de 105 instituciones exactas; el resto '
                   'se explica casi por completo por registros sin MRUN. Sistema 77,8% vs 77,9% oficial.'),
]

PCT_COLS = {'Retención', 'Persist. misma inst.', 'Persist. ES', 'No matrícula',
            'Retención mín (X no retenidos)', 'Retención máx (X retenidos)'}


def escribir_excel(out, hojas):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre, index=False)
            ws = w.sheets[nombre]
            ws.freeze_panes = 'A2'
            for j, col in enumerate(df.columns, start=1):
                letra = get_column_letter(j)
                c = ws.cell(row=1, column=j)
                c.font = Font(name='Arial', bold=True, color='FFFFFF')
                c.fill = PatternFill('solid', fgColor='1F4E78')
                c.alignment = Alignment(wrap_text=True, vertical='center')
                es_pct = str(col) in PCT_COLS or (nombre == 'Retención por inst.' and str(col).isdigit())
                ancho = min(max(len(str(col)), *(len(str(v)) for v in df[col].head(200))) + 2, 60) if len(df) else 14
                ws.column_dimensions[letra].width = max(ancho, 10)
                for i in range(2, len(df) + 2):
                    cell = ws.cell(row=i, column=j)
                    cell.font = Font(name='Arial')
                    if es_pct:
                        cell.number_format = '0.0%'


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--carpeta', required=True)
    ap.add_argument('--out', default='retencion_1er_anio.xlsx')
    ap.add_argument('--inst', nargs='*', default=[], help='COD_INST con detalle (ej: 69 86)')
    ap.add_argument('--excluir', action='append', default=None,
                    help='AÑO:COD[,COD] instituciones a excluir de la cohorte de ese año')
    ap.add_argument('--cache', default=None)
    ap.add_argument('--acreditacion', default=None, help='Excel de acreditación institucional CNA (usa la columna "Años")')
    a = ap.parse_args()

    carpeta = Path(a.carpeta).resolve()
    if not carpeta.is_dir():
        sys.exit(f'No existe la carpeta {carpeta}')
    cache = (Path(a.cache) if a.cache else carpeta / '_cache_retencion').resolve()
    excl = {}
    for e in (a.excluir if a.excluir is not None else ['2025:430']):
        anio, cods = e.split(':')
        excl[int(anio)] = [c.strip() for c in cods.split(',') if c.strip()]
    inst_det = [str(i) for i in a.inst]

    log('1) Leyendo bases')
    bases, archivos = cargar_bases(carpeta, cache)
    anios = sorted(bases)
    cohortes = [t for t in anios if t + 1 in bases]
    sin_t1 = [t for t in anios if t + 1 not in bases]
    log(f'   Años disponibles: {anios}')
    log(f'   Cohortes calculables: {cohortes}  (sin año siguiente: {sin_t1})')
    if not cohortes:
        sys.exit('No hay pares de años consecutivos (t y t+1) para calcular.')

    sistema, tipo_inst, por_inst, filtros = [], [], [], []
    det = {'carrera': [], 'sexo': [], 'programa': [], 'detalle': []}
    actual = None
    for t in cohortes:
        log(f'2) Cohorte {t}')
        dt = actual if actual is not None and actual[0] == t else (t, pd.read_pickle(bases[t][0]))
        dt1 = (t + 1, pd.read_pickle(bases[t + 1][0]))
        coh, pasos = cohorte(dt[1], t, excl.get(t, []))
        res = clasificar(coh, dt1[1], t)
        res.insert(0, 'cohorte', t)
        filtros += [{'Cohorte': t, 'Paso': p, 'Registros': n} for p, n in pasos]
        sistema.append(tabla(res.assign(Nivel='Total sistema'), ['cohorte', 'Nivel']))
        tipo_inst.append(tabla(res, ['cohorte', 'tipo_inst_1']))
        por_inst.append(tabla(res, ['cohorte', 'cod_inst', 'nomb_inst', 'tipo_inst_1']))
        if inst_det:
            r = res[res.cod_inst.isin(inst_det)]
            if len(r):
                det['carrera'].append(tabla(r, ['cohorte', 'cod_inst', 'nomb_inst', 'nomb_carrera', 'jornada']))
                det['sexo'].append(tabla(r, ['cohorte', 'cod_inst', 'nomb_inst', 'gen_alu']))
                det['programa'].append(tabla(r, ['cohorte', 'cod_inst', 'nomb_inst', 'nivel_carrera_1']))
                det['detalle'].append(r)
        s = sistema[-1].iloc[0]
        log(f'   N={int(s.N):,}  Retención={s["Retención"]:.1%}  Persist.ES={s["Persist. ES"]:.1%}')
        actual = dt1  # t+1 se reutiliza como cohorte siguiente
        del dt, coh, res

    pi = pd.concat(por_inst, ignore_index=True)
    nombres = pi.sort_values('cohorte').groupby('cod_inst').agg(
        nomb_inst=('nomb_inst', 'last'), tipo_inst_1=('tipo_inst_1', 'last'))
    piv = (pi[pi['Publicable (N>=10)']]
           .pivot_table(index='cod_inst', columns='cohorte', values='Retención', aggfunc='first'))
    piv.columns = [str(c) for c in piv.columns]
    piv = nombres.join(piv, how='right').reset_index().sort_values(['tipo_inst_1', 'nomb_inst'])

    cruce = None
    if a.acreditacion:
        log('   Cruzando con acreditación CNA (columna "Años")')
        cruce = cruzar_acreditacion(pi, leer_acreditacion(a.acreditacion))
        ac = dict(zip(cruce.cod_inst, cruce[COL_ACRED]))
        pi[COL_ACRED] = pd.to_numeric(pi.cod_inst.astype(str).map(ac)).astype('Int64')
        piv.insert(3, COL_ACRED, pd.to_numeric(piv.cod_inst.astype(str).map(ac)).astype('Int64'))
        sin = cruce[~cruce['Encontrada en CNA']]
        log(f'   {int(cruce["Encontrada en CNA"].sum())} de {len(cruce)} instituciones encontradas en el archivo CNA')

    leeme = pd.DataFrame(LEEME, columns=['Concepto', 'Descripción'])
    leeme = pd.concat([leeme, pd.DataFrame([
        ('Cohortes calculadas', ', '.join(map(str, cohortes))),
        ('Años sin t+1 (no calculables)', ', '.join(map(str, sin_t1)) or '—'),
        ('Exclusiones aplicadas', '; '.join(f'{k}: {",".join(v)}' for k, v in excl.items()) or '—'),
        ('Acreditación', (f'Columna "Años" del archivo CNA {Path(a.acreditacion).name}: último proceso de cada '
                          'institución, igual para todas las cohortes. Vacío = no acreditada o sin cruce '
                          '(ver hoja "Cruce CNA").') if a.acreditacion else 'No se cruzó (sin --acreditacion).'),
    ], columns=['Concepto', 'Descripción'])])

    hojas = {
        'Léeme': leeme,
        'Archivos': archivos,
        'Sistema': pd.concat(sistema, ignore_index=True),
        'Por tipo inst.': pd.concat(tipo_inst, ignore_index=True),
        'Retención por inst.': piv,
        'Por institución': pi.sort_values(['cod_inst', 'cohorte']),
        'Filtros': pd.DataFrame(filtros),
    }
    if cruce is not None:
        hojas['Cruce CNA'] = cruce.sort_values(['Encontrada en CNA', 'tipo_inst_1', 'nomb_inst'])
    if inst_det and det['carrera']:
        hojas['Det. por carrera'] = pd.concat(det['carrera'], ignore_index=True)
        hojas['Det. por sexo'] = pd.concat(det['sexo'], ignore_index=True)
        hojas['Det. por programa'] = pd.concat(det['programa'], ignore_index=True)
        detalle = pd.concat(det['detalle'], ignore_index=True)
        if len(detalle) < 1_000_000:
            hojas['Det. por MRUN'] = detalle
        else:
            csv = Path(a.out).with_suffix('.detalle.csv')
            detalle.to_csv(csv, index=False, sep=';')
            log(f'   Detalle por MRUN supera el límite de Excel; guardado en {csv}')

    log('3) Escribiendo Excel')
    escribir_excel(a.out, hojas)
    log(f'Listo: {a.out}')


if __name__ == '__main__':
    main()
