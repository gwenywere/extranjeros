"""Build data.json for the 'Foreigners in Spain' site from the two official-statistics workbooks.

Usage (from the project root):
    python site_foreigners/scripts/build_data.py \
        "Extranjeros en España empleo, afiliación, prestaciones y pobreza (datos oficiales).xlsx" \
        Fuentes_estadisticas_inmigracion_Espana_Euskadi.xlsx \
        site_foreigners/data.json

Optional: --eustat <csv> (default data/eustat_tbl0021467_2025-01-01.csv), Eustat's table of the Basque population 16+
by continent of birth, activity, province and sex, downloaded from https://www.eustat.eus/elementos/xls0021467_e.csv

Optional: --census <px> (default data/ine_67129_censo_actividad_pais_nacimiento.px), INE table 67129 from the annual
census: population 16+ by region/province, sex, country of birth and activity, 2021-2024. Written to census.json
(next to data.json) because it is large; the page loads it separately. https://www.ine.es/jaxiT3/files/t/es/px/67129.px

Every time series is stored as {"freq": Q|M|Y|D, "t": [ISO dates], <key>: [values]}; missing values are null.
EPA figures are converted from thousands to persons; shares are stored as percentages (0-100).
"""
from __future__ import annotations

import argparse
import csv
import json
import math
import re
from datetime import date
from pathlib import Path

import pandas as pd

COUNTRY_EN = {
    'Marruecos': 'Morocco', 'Rumania': 'Romania', 'Italia': 'Italy', 'Perú': 'Peru', 'Ucrania': 'Ukraine',
    'Pakistán': 'Pakistan',
}
EPA_GROUPS = {'Total': 'total', 'Española': 'es', 'Extranjera': 'fx', 'Extranjera UE': 'eu', 'Extranjera no UE': 'neu'}
EPA_MEASURES = {  # column prefix -> (key, multiplier)
    'Tasa de paro (%)': ('unempRate', 1), 'Tasa de actividad (%)': ('actRate', 1),
    'Ocupados (miles)': ('employed', 1000), 'Activos (miles)': ('active', 1000),
    'Población 16+ (miles)': ('pop16', 1000), 'Parados (miles, calculado)': ('unemployed', 1000),
}
ECV_MEASURES = {
    'AROPE Europa 2030 (%)': 'arope', 'Riesgo de pobreza (%)': 'poverty',
    'Carencia material y social severa (%)': 'deprivation', 'Baja intensidad en el empleo (%)': 'lowWork',
    'AROPE Europa 2020 (%)': 'arope2020',
}
ECV_GROUPS = {'Total': 'total', 'Española': 'es', 'Extranjera UE': 'eu', 'Extranjera resto del mundo': 'neu'}


def clean(v, digits=4):
    """JSON-safe number: NaN -> None, floats rounded."""
    if v is None:
        return None
    if isinstance(v, float):
        if math.isnan(v):
            return None
        r = round(v, digits)
        return int(r) if r == int(r) and abs(r) > 1000 else r
    if hasattr(v, 'item'):
        return clean(v.item(), digits)
    return v


def col(s: pd.Series, mult=1, digits=4) -> list:
    return [clean(v * mult if v is not None and not (isinstance(v, float) and math.isnan(v)) else None, digits) for v in s]


def quarter_dates(periods) -> list[str]:
    out = []
    for p in periods:
        y, q = str(p).split('T')
        out.append(date(int(y), 3 * int(q) - 1, 15).isoformat())  # mid-quarter
    return out


def epa_block(df: pd.DataFrame) -> dict:
    block = {'freq': 'Q', 't': quarter_dates(df['Periodo'])}
    for prefix, (key, mult) in EPA_MEASURES.items():
        groups = {}
        for g, gk in EPA_GROUPS.items():
            c = f'{prefix} · {g}'
            if c in df:
                groups[gk] = col(df[c], mult, 2)
        block[key] = groups
    return block


def snapshot(path: Path) -> dict:
    x = pd.read_excel(path, sheet_name='INE_Poblacion', header=3, nrows=4)
    cross = {f"{r['Nacionalidad']}|{r['Lugar de nacimiento']}": int(r['Personas']) for _, r in x.iterrows()}
    eus = pd.read_excel(path, sheet_name='Eustat_Actividad', header=3, nrows=4).set_index('Categoría')
    sepe = pd.read_excel(path, sheet_name='SEPE_Extranjeros', header=3, nrows=10).set_index('Indicador')['Valor']
    pr = pd.read_excel(path, sheet_name='Prestaciones', header=3, nrows=9)
    val = lambda prog, cat: float(pr.loc[pr['Programa'].str.startswith(prog) & (pr['Categoría'] == cat), 'Personas / porcentaje'].iloc[0])
    return {
        'nationalityBirth': {
            'date': '2026-10-01',
            'esBornEs': cross['Española|España'], 'esBornAbroad': cross['Española|Extranjero'],
            'fxBornAbroad': cross['Extranjera|Extranjero'], 'fxBornEs': cross['Extranjera|España'],
        },
        'eustat': {
            'date': '2025-01-01',
            **{k: {'pop16': int(eus.loc['Población 16+', c]), 'employed': int(eus.loc['Ocupada', c]),
                   'unemployed': int(eus.loc['Parada', c]), 'inactive': int(eus.loc['Inactiva', c])}
               for k, c in [('bornEs', 'Nacidos en España'), ('bornAbroad', 'Nacidos en el extranjero')]},
        },
        'sepe2025': {
            'population': int(sepe['Población extranjera']), 'active': int(sepe['Población activa extranjera (EPA)']),
            'employed': int(sepe['Población ocupada extranjera (EPA)']), 'unemployedEpa': int(sepe['Población parada extranjera (EPA)']),
            'affiliations': int(sepe['Afiliaciones extranjeras a la Seguridad Social']),
            'registeredUnemployed': int(sepe['Personas extranjeras paradas registradas']),
            'benefitRecipients': int(sepe['Beneficiarias de prestaciones por desempleo']),
            'actRate': round(sepe['Tasa de actividad'] * 100, 2), 'empRate': round(sepe['Tasa de empleo'] * 100, 2),
            'unempRate': round(sepe['Tasa de paro'] * 100, 2),
        },
        'benefitsDec2025': {
            'contributory': int(val('Prestaciones', 'Contributiva')), 'subsidy': int(val('Prestaciones', 'Subsidio')),
            'rai': int(val('Prestaciones', 'Renta Activa de Inserción')),
            'agrarian': int(val('Prestaciones', 'Subsidios eventuales agrarios')),
        },
        'imvDec2025': {'fx': int(val('Ingreso Mínimo Vital', 'Titulares nacionalidad extranjera')),
                       'es': int(val('Ingreso Mínimo Vital', 'Titulares nacionalidad española'))},
        'rgi2024': {
            'shareForeignOfRecipients': round(val('Rentas mínimas', 'Proporción de perceptores con nacionalidad extranjera') * 100, 1),
            'coverageForeign': round(val('Rentas mínimas', 'Cobertura entre población extranjera') * 100, 1),
            'coverageSpanish': round(val('Rentas mínimas', 'Cobertura entre población española') * 100, 1),
        },
    }


EUSTAT_GROUPS = {'Guztira': 'total', 'Espainia': 'es', 'Europa': 'europe', 'Afrika': 'africa',
                 'Amerika': 'americas', 'Asia/Ozeania': 'asia'}
EUSTAT_STATUS = {'Biztanleria landuna': 'employed', 'Biztanleria langabea': 'unemployed', 'Biztanleria ez-aktiboa': 'inactive'}
EUSTAT_TERR = ['pv', 'araba', 'bizkaia', 'gipuzkoa']  # column blocks, in file order
EUSTAT_SEX = ['all', 'men', 'women']


def eustat_continents(path: Path) -> dict:
    """Eustat tbl0021467 CSV (Basque labels, ';'-separated, '.' thousands) ->
    {date, groups: {group: {territory: {sex: {employed, unemployed, inactive}}}}}."""
    rows = list(csv.reader(path.read_text(encoding='utf-8-sig').splitlines(), delimiter=';'))
    m = re.search(r'(\d{4})/(\d{2})/(\d{2})', rows[0][0])
    out: dict = {'date': f'{m[1]}-{m[2]}-{m[3]}' if m else None, 'groups': {}}
    group = None
    for r in rows:
        if not r or not r[0]:
            continue
        name = r[0].strip()
        if name in EUSTAT_GROUPS:
            group = EUSTAT_GROUPS[name]
            out['groups'][group] = {t: {s: {} for s in EUSTAT_SEX} for t in EUSTAT_TERR}
        elif name in EUSTAT_STATUS and group:
            nums = [int(v.replace('.', '')) for v in r[1:13]]
            for i, v in enumerate(nums):
                out['groups'][group][EUSTAT_TERR[i // 3]][EUSTAT_SEX[i % 3]][EUSTAT_STATUS[name]] = v
    # check: the five birth groups add up to the total in every cell
    g = out['groups']
    if set(g) != set(EUSTAT_GROUPS.values()):
        raise SystemExit(f'Unexpected Eustat file: groups found {sorted(g)}')
    for t in EUSTAT_TERR:
        for s in EUSTAT_SEX:
            for k in EUSTAT_STATUS.values():
                parts = sum(g[x][t][s][k] for x in g if x != 'total')
                if parts != g['total'][t][s][k]:
                    raise SystemExit(f'Eustat groups do not add up for {t}/{s}/{k}: {parts} vs {g["total"][t][s][k]}')
    return out


# ---------- INE table 67129 (annual census): population 16+ by sex, country of birth and activity ----------
CENSUS_COUNTRIES = {  # INE name -> (English name, continent)
    'España': ('Spain', 'spain'),
    'Francia': ('France', 'europe'), 'Reino Unido': ('United Kingdom', 'europe'), 'Rumanía': ('Romania', 'europe'),
    'Ucrania': ('Ukraine', 'europe'), 'Otros países de Europa': ('Other Europe', 'europe'),
    'Marruecos': ('Morocco', 'africa'), 'Otros países de África': ('Other Africa', 'africa'),
    'Cuba': ('Cuba', 'americas'), 'República Dominicana': ('Dominican Republic', 'americas'),
    'Argentina': ('Argentina', 'americas'), 'Bolivia': ('Bolivia', 'americas'), 'Colombia': ('Colombia', 'americas'),
    'Ecuador': ('Ecuador', 'americas'), 'Perú': ('Peru', 'americas'), 'Venezuela': ('Venezuela', 'americas'),
    'Otros países de América': ('Other Americas', 'americas'),
    'China': ('China', 'asia'), 'Otros países de Asia': ('Other Asia', 'asia'), 'Oceanía': ('Oceania', 'asia'),
}
CENSUS_ACTS = {  # INE name -> key; 'Total' is checked, not stored
    'Ocupado/a': 'employed', 'Parado/a': 'unemployed',
    'Perceptor/a pensión de incapacidad, jubilación, prejubilación': 'pension',
    'Otra situación de inactividad': 'otherInactive', 'Estudiante': 'student',
}
CENSUS_SEX = {'Total': 'all', 'Hombres': 'men', 'Mujeres': 'women'}


def read_px(path: Path) -> tuple[list[tuple[str, list[str]]], list[float | None]]:
    """Minimal PC-Axis reader: dimensions (STUB then HEADING order) and the flat DATA list."""
    txt = path.read_text(encoding='latin-1')
    def values(name):
        m = re.search(r'VALUES\("%s"\)=((?:\s*"[^"]*"\s*,?)+)\s*;' % re.escape(name), txt, re.S)
        return re.findall(r'"([^"]*)"', m.group(1))
    def names(key):
        m = re.search(r'^%s=((?:\s*"[^"]*"\s*,?)+)\s*;' % key, txt, re.S | re.M)
        return re.findall(r'"([^"]*)"', m.group(1)) if m else []
    dims = [(n, values(n)) for n in names('STUB') + names('HEADING')]
    raw = txt.split('DATA=', 1)[1].rsplit(';', 1)[0].split()
    data = [None if x.strip('"') in ('', '.', '..', '...', '-') else float(x.strip('"')) for x in raw]
    if len(data) != math.prod(len(v) for _, v in dims):
        raise SystemExit(f'{path}: DATA has {len(data)} values, dimensions need {math.prod(len(v) for _, v in dims)}')
    return dims, data


def census_birth(path: Path) -> dict:
    """INE 67129 -> compact JSON: per region, a flat list ordered [sex][country][year][activity]."""
    dims, data = read_px(path)
    names = [n for n, _ in dims]
    expected = ['Comunidad Autónoma y Provincia', 'Sexo', 'País de nacimiento', 'Relación con la actividad', 'Periodo']
    if names != expected:
        raise SystemExit(f'Unexpected dimensions in {path}: {names}')
    regions, sexes, countries, acts, years = (v for _, v in dims)
    unknown = set(countries) - set(CENSUS_COUNTRIES) - {'Total'}
    if unknown or set(acts) - set(CENSUS_ACTS) - {'Total'}:
        raise SystemExit(f'Unexpected categories in {path}: {unknown or set(acts) - set(CENSUS_ACTS)}')
    strides = [math.prod(len(v) for _, v in dims[i + 1:]) for i in range(len(dims))]
    at = lambda *ix: data[sum(i * s for i, s in zip(ix, strides))]

    keep_c = [c for c in countries if c != 'Total']
    keep_a = [a for a in acts if a != 'Total']
    yrs = sorted(years)  # oldest first
    out_regions = []
    for ri, r in enumerate(regions):
        flat = []
        for si, _ in enumerate(sexes):
            for c in keep_c:
                ci = countries.index(c)
                for y in yrs:
                    yi = years.index(y)
                    vals = [at(ri, si, ci, acts.index(a), yi) for a in keep_a]
                    total = at(ri, si, ci, acts.index('Total'), yi)
                    if None not in vals and total is not None and abs(sum(vals) - total) > 0.5:
                        raise SystemExit(f'67129: activities do not add up for {r}/{c}/{y}')
                    flat += [None if v is None else int(v) for v in vals]
        name = 'Spain' if r == 'Total Nacional' else re.sub(r'^\d\d\s+', '', r)
        out_regions.append({'name': name, 'v': flat})
    # the PX lists each comunidad autónoma ('01 Andalucía', code CA01) right before its provinces ('04 Almería', code 04);
    # CODES tell them apart
    txt = path.read_text(encoding='latin-1')
    codes = re.findall(r'"([^"]*)"', re.search(r'CODES\("Comunidad Autónoma y Provincia"\)=((?:\s*"[^"]*"\s*,?)+)\s*;', txt, re.S).group(1))
    if len(codes) != len(out_regions):
        raise SystemExit(f'67129: {len(codes)} region codes for {len(out_regions)} regions')
    for reg, code in zip(out_regions, codes):
        reg['level'] = 'country' if code == 'CA00' else 'region' if code.startswith('CA') else 'province'
        reg['code'] = code  # provinces: INE province code ('04'), the map's feature id; regions: 'CA01'…
    return {
        'source': 'INE 67129', 'years': [int(y) for y in yrs], 'sexes': [CENSUS_SEX[s] for s in sexes],
        'countries': [{'name': CENSUS_COUNTRIES[c][0], 'cont': CENSUS_COUNTRIES[c][1]} for c in keep_c],
        'acts': [CENSUS_ACTS[a] for a in keep_a], 'regions': out_regions,
    }


class PX:
    """A PC-Axis cube: px.get(dim=value, …) for one cell; px.values(dim) for a dimension's categories."""
    def __init__(self, path: Path):
        self.dims, self.data = read_px(path)
        self.names = [n for n, _ in self.dims]
        self.index = {n: {v: i for i, v in enumerate(vals)} for n, vals in self.dims}
        self.strides = [math.prod(len(v) for _, v in self.dims[i + 1:]) for i in range(len(self.dims))]
        txt = path.read_text(encoding='latin-1')
        self.codes = {n: re.findall(r'"([^"]*)"', m.group(1))
                      for n in self.names if (m := re.search(r'CODES\("%s"\)=((?:\s*"[^"]*"\s*,?)+)\s*;' % re.escape(n), txt, re.S))}

    def values(self, name):
        return dict(self.dims)[name]

    def get(self, **cells):
        cells = {k.replace('_', ' '): v for k, v in cells.items()}
        return self.data[sum(self.index[n][cells[n]] * s for n, s in zip(self.names, self.strides))]


EPA_NAT = {'Total': 'total', 'Española': 'es', 'Extranjera: Total': 'fx', 'Extranjera: Unión Europea': 'eu',
           'Extranjera: No pertenecientes a la Unión Europea': 'neu'}
SEX3 = {'Ambos sexos': 'all', 'Total': 'all', 'Hombres': 'men', 'Mujeres': 'women'}


def epa_regions(pop_px: Path, act_px: Path, emp_px: Path) -> dict:
    """EPA by region × nationality × sex: annual averages (complete years only) of population 16+, active and employed,
    in thousands. Rates are computed from these averages on the page."""
    cubes = {'pop16': PX(pop_px), 'active': PX(act_px), 'employed': PX(emp_px)}
    ref = cubes['pop16']
    reg_dim, nat_dim = 'Comunidades y Ciudades Autónomas', 'Nacionalidad'
    for k, c in cubes.items():
        if c.values(reg_dim) != ref.values(reg_dim) or c.values(nat_dim) != ref.values(nat_dim):
            raise SystemExit(f'EPA tables do not share regions/nationalities ({k})')
    periods = ref.values('Periodo')
    years = sorted({p[:4] for p in periods if all(f'{p[:4]}T{q}' in periods for q in range(1, 5))})
    regions = ref.values(reg_dim)
    out = {'source': 'INE EPA 65290 65299 65309', 'years': [int(y) for y in years], 'measures': list(cubes),
           'nats': [EPA_NAT[n] for n in ref.values(nat_dim)], 'sexes': [SEX3[s] for s in ref.values('Sexo')], 'regions': []}
    for r in regions:
        flat = []  # [sex][nat][year][measure]
        for s in ref.values('Sexo'):
            for n in ref.values(nat_dim):
                for y in years:
                    for k, c in cubes.items():
                        q = [c.get(Sexo=s, **{reg_dim: r, nat_dim: n}, Periodo=f'{y}T{i}') for i in range(1, 5)]
                        flat.append(None if None in q else round(sum(q) / 4, 2))
        code = 'CA00' if r == 'Total Nacional' else f'CA{r[:2]}'
        out['regions'].append({'code': code, 'name': 'Spain' if r == 'Total Nacional' else re.sub(r'^\d\d\s+', '', r), 'v': flat})
    return out


def census_nationality(path: Path) -> dict:
    """INE 67130 reduced to Spanish vs foreign nationals: per region/province, flat [sex][es, fx][year][activity]."""
    c = PX(path)
    reg_dim, nat_dim, act_dim = 'Comunidad Autónoma y Provincia', 'País de nacionalidad', 'Relación con la actividad'
    years = sorted(c.values('Periodo'))
    acts = [a for a in c.values(act_dim) if a != 'Total']
    codes = c.codes[reg_dim]
    out = {'source': 'INE 67130', 'years': [int(y) for y in years], 'acts': [CENSUS_ACTS[a] for a in acts], 'regions': []}
    for r, code in zip(c.values(reg_dim), codes):
        flat = []
        for s in c.values('Sexo'):
            for which in ('es', 'fx'):
                for y in years:
                    for a in acts:
                        tot = c.get(**{reg_dim: r, nat_dim: 'Total', act_dim: a}, Sexo=s, Periodo=y)
                        es = c.get(**{reg_dim: r, nat_dim: 'España', act_dim: a}, Sexo=s, Periodo=y)
                        flat.append(int(es if which == 'es' else tot - es))
        out['regions'].append({'code': code, 'v': flat})
    return out


SECTORS_EN = {
    'A': 'Agriculture, forestry, fishing', 'B': 'Mining and quarrying', 'C': 'Manufacturing', 'D': 'Electricity and gas supply',
    'E': 'Water supply and waste', 'F': 'Construction', 'G': 'Wholesale and retail trade', 'H': 'Transport and storage',
    'I': 'Hospitality (hotels, restaurants)', 'J': 'Information and communication', 'K': 'Finance and insurance',
    'L': 'Real estate', 'M': 'Professional, scientific, technical', 'N': 'Administrative and support services',
    'O': 'Public administration and defence', 'P': 'Education', 'Q': 'Health and social work', 'R': 'Arts and recreation',
    'S': 'Other services', 'T': 'Households as employers (domestic work)', 'U': 'Extraterritorial organisations',
    'X': 'Not recorded',
}


def census_sectors(path: Path) -> dict:
    """INE 69964: employed 16+ by region/province, economic section (CNAE letter), sex and nationality, 2021-2024.
    Per region a flat list [sex][nationality: total, es, fx][year][section]; section 0 is the total."""
    c = PX(path)
    reg_dim, act_dim = 'Comunidad Autónoma y Provincia', 'Actividad económica'
    acts = c.values(act_dim)
    letters = ['total' if a.startswith('Total') else 'X' if a == 'No consta' else a.split(' ', 1)[0] for a in acts]
    unknown = [l for l in letters if l != 'total' and l not in SECTORS_EN]
    if unknown:
        raise SystemExit(f'69964: unexpected sections {unknown}')
    years = sorted(c.values('Periodo'))
    nats = c.values('Nacionalidad')
    out = {'source': 'INE 69964', 'years': [int(y) for y in years], 'nats': ['total' if n == 'Total' else 'es' if n == 'Española' else 'fx' for n in nats],
           'sections': [{'code': l, 'name': 'All sectors' if l == 'total' else SECTORS_EN[l]} for l in letters], 'regions': []}
    for r, code in zip(c.values(reg_dim), c.codes[reg_dim]):
        flat = []
        for s in c.values('Sexo'):
            for n in nats:
                for y in years:
                    vals = [c.get(**{reg_dim: r, act_dim: a}, Sexo=s, Nacionalidad=n, Periodo=y) for a in acts]
                    if None not in vals and abs(sum(vals[1:]) - vals[0]) > 0.5:
                        raise SystemExit(f'69964: sections do not add up for {r}/{s}/{n}/{y}')
                    flat += [None if v is None else int(v) for v in vals]
        out['regions'].append({'code': code, 'v': flat})
    return out


def px_pop16(path: Path, region_prefix: str) -> dict[str, dict[str, float]]:
    """Population 16+ (activity 'Total', both sexes) for one region of INE 67129/67130: {year: {country: n}}."""
    dims, data = read_px(path)
    regions, sexes, countries, acts, years = (v for _, v in dims)
    ri = next(i for i, r in enumerate(regions) if r.startswith(region_prefix))
    strides = [math.prod(len(v) for _, v in dims[i + 1:]) for i in range(len(dims))]
    at = lambda *ix: data[sum(i * s for i, s in zip(ix, strides))]
    si, ai = sexes.index('Total'), acts.index('Total')
    return {y: {c: at(ri, si, ci, ai, yi) for ci, c in enumerate(countries)} for yi, y in enumerate(years)}


RGI_BIRTH_TO_CONT = {'spain': 'spain', 'eu': 'europe', 'rest_of_europe': 'europe', 'maghreb': 'africa',
                     'subsaharan_africa': 'africa', 'south_america': 'americas', 'north_america': 'americas',
                     'asia': 'asia', 'oceania': 'asia', 'stateless': None}


def rgi_context(rgi_csv: Path, census_birth_px: Path, census_nat_px: Path, region='16 País Vasco') -> dict:
    """Basque RGI holders in 2023 by nationality and continent of birth, with INE census adults (16+) as denominators;
    and the foreign-national share of Basque adults, 2021-2024."""
    holders = {(r['variable'], r['group']): int(r['holders']) for r in csv.DictReader(rgi_csv.open(encoding='utf-8'))}
    nat = px_pop16(census_nat_px, region)
    birth = px_pop16(census_birth_px, region)
    for var in ('nationality', 'birth'):
        tot = sum(v for (k, _), v in holders.items() if k == var)
        if var == 'birth' and tot != holders[('nationality', 'es')] + holders[('nationality', 'fx')]:
            raise SystemExit(f'{rgi_csv}: birth groups add up to {tot}, nationality to a different total')
    y = '2023'
    pop_cont: dict[str, float] = {}
    for c, n in birth[y].items():
        if c != 'Total':
            k = CENSUS_COUNTRIES[c][1]
            pop_cont[k] = pop_cont.get(k, 0) + n
    hold_cont: dict[str, int] = {}
    for (var, g), n in holders.items():
        if var == 'birth' and RGI_BIRTH_TO_CONT[g]:
            hold_cont[RGI_BIRTH_TO_CONT[g]] = hold_cont.get(RGI_BIRTH_TO_CONT[g], 0) + n
    return {
        'year': 2023,
        'foreignShareAdults': {int(yr): round(100 * (v['Total'] - v['España']) / v['Total'], 4) for yr, v in sorted(nat.items())},
        'nationality': {
            'es': {'holders': holders[('nationality', 'es')], 'pop16': int(nat[y]['España'])},
            'fx': {'holders': holders[('nationality', 'fx')], 'pop16': int(nat[y]['Total'] - nat[y]['España'])},
        },
        'birth': {k: {'holders': hold_cont[k], 'pop16': int(pop_cont[k])} for k in ['spain', 'europe', 'africa', 'americas', 'asia']},
        'birthDetail': {g: n for (var, g), n in holders.items() if var == 'birth'},
    }


def sepe_report(path: Path) -> dict:
    """Figures typed in from SEPE's 2025 report on foreign workers (see the .README.txt next to the CSV):
    {section: {key: {group: value}}}, after checking they add up to the report's own totals."""
    out: dict = {}
    for r in csv.DictReader(path.open(encoding='utf-8')):
        v = float(r['value'])
        out.setdefault(r['section'], {}).setdefault(r['key'], {})[r['group']] = int(v) if v.is_integer() and v > 100 else v
    def check(name, got, want, tol=0.6):
        if abs(got - want) > tol:
            raise SystemExit(f'{path}: {name} adds up to {got}, expected {want}')
    a = out['affil_scheme']
    men, women = sum(x['men'] for x in a.values()), sum(x['women'] for x in a.values())
    check('affiliation by scheme', men + women + out['affil_total']['foreign']['sex_unknown'], out['affil_total']['foreign']['count'])
    for sex, total in (('men', men), ('women', women)):
        check(f'affiliation by nationality, {sex}', sum(out['affil_nat_sex'][sex].values()), total)
        check(f'employees + self-employed, {sex}', sum(out['affil_dependency'][sex].values()), total)
    ct = out['contracts_total']['all']['foreign']
    check('contracts by sex and nationality', sum(sum(x.values()) for x in out['contracts_sex_nat'].values()), ct)
    check('contracts by type', sum(x['full_time'] + x['part_time'] + x['fixed_discontinuous'] for x in out['contracts_type'].values()), ct)
    check('contracts by sector', sum(x['foreign'] for x in out['contracts_sector'].values()), ct)
    for g in ('foreign', 'spanish'):
        check(f'occupational groups, {g}', sum(x[g] for x in out['occ_group_pct'].values()), 100, tol=0.1)
    fl = out['occ_top']['farm_labourers']
    check('foreign share of farm labourers', round(100 * fl['foreign'] / (fl['foreign'] + fl['spanish']), 2), 51.77, tol=0.01)
    return out


def build(series_file: Path, snapshot_file: Path, eustat_csv: Path | None = None) -> dict:
    rd = lambda sheet: pd.read_excel(series_file, sheet_name=sheet)
    d: dict = {'generated': date.today().isoformat()}

    d['epa'] = epa_block(rd('EPA España'))
    d['epaPV'] = epa_block(rd('EPA País Vasco'))

    a = rd('Afiliación SS')
    d['affiliation'] = {
        'freq': 'M', 't': [f'{m}-15' for m in a['Mes'].astype(str)],
        'fx': col(a['Afiliados extranjeros']), 'total': col(a['Afiliados total sistema']),
        'share': col(a['% extranjeros sobre total (calculado)'], 100, 3),
        'eu': col(a['Extranjeros UE']), 'neu': col(a['Extranjeros no UE']), 'women': col(a['Extranjeros mujeres']),
        'general': col(a['Extranjeros · Régimen General']), 'agrarian': col(a['Extranjeros · S.E. Agrario']),
        'domestic': col(a['Extranjeros · S.E. Hogar']), 'selfEmployed': col(a['Extranjeros · Autónomos']),
        'pv': col(a['Afiliados extranjeros en País Vasco']),
    }

    p = rd('Afiliados por país')
    d['byCountry'] = {
        'freq': 'D', 't': [pd.Timestamp(x).date().isoformat() for x in p['Fecha']],
        'countries': {COUNTRY_EN.get(c, c): col(p[c]) for c in p.columns if c != 'Fecha'},
    }

    b = rd('Prestaciones desempleo')
    d['benefits'] = {
        'freq': 'Y', 't': [f'{y}-07-01' for y in b['Año']],
        'fx': col(b['Beneficiarios extranjeros (media anual)']), 'share': col(b['% sobre el total de beneficiarios'], 100, 2),
        'contributory': col(b['Nivel contributivo']), 'assistance': col(b['Nivel asistencial (subsidios)']),
        'rai': col(b['Renta activa de inserción']), 'pae': col(b['Programa de activación para el empleo']),
    }

    i = rd('IMV titulares')
    d['imv'] = {
        'freq': 'M', 't': [f'{m}-15' for m in i['Mes'].astype(str)],
        'es': col(i['Titulares españoles']), 'fx': col(i['Titulares extranjeros']),
        'share': col(i['% extranjeros (calculado)'], 100, 3),
    }

    e = rd('Pobreza (ECV)')
    d['poverty'] = {'freq': 'Y', 't': [f'{y}-07-01' for y in e['Año de la encuesta']]}
    for prefix, key in ECV_MEASURES.items():
        d['poverty'][key] = {gk: col(e[f'{prefix} · {g}'], 1, 2) for g, gk in ECV_GROUPS.items()}

    pob = rd('Población')
    d['population'] = {
        'freq': 'D', 't': [pd.Timestamp(x).date().isoformat() for x in pob['Fecha']],
        'total': col(pob['Población total']), 'esNat': col(pob['Nacionalidad española']),
        'fxNat': col(pob['Nacionalidad extranjera']), 'fxNatShare': col(pob['% nacionalidad extranjera (calculado)'], 100, 3),
        'bornEs': col(pob['Nacidos en España']), 'bornAbroad': col(pob['Nacidos en el extranjero']),
        'bornAbroadShare': col(pob['% nacidos en el extranjero (calculado)'], 100, 3),
    }

    d['snapshots'] = snapshot(snapshot_file)
    if eustat_csv is not None:
        d['snapshots']['eustatContinents'] = eustat_continents(eustat_csv)
    return d


def benefit_rates(path: Path, data: dict) -> dict:
    """All unemployment-benefit recipients (yearbook PRD-1) minus foreign ones (PEX-1) = Spanish nationals,
    each per 100 people in work, in the labour force, unemployed and aged 16+ (EPA annual averages)."""
    prd = pd.read_csv(path)
    parts = prd[['contributory', 'assistance', 'rai', 'other']].sum(axis=1)
    bad = prd[(parts - prd['total']).abs() > 5]
    if len(bad):
        raise SystemExit(f'{path}: benefit levels do not add up to the total in {bad["year"].tolist()}')
    b, e = data['benefits'], data['epa']
    years = [int(t[:4]) for t in b['t']]
    prd = prd.set_index('year')
    if set(years) - set(prd.index):
        raise SystemExit(f'{path}: no total for {sorted(set(years) - set(prd.index))}')
    out = {'freq': 'Y', 't': b['t'], 'recipients': {'fx': [], 'es': []}, 'share': []}
    measures = {'employed': 'employed', 'active': 'active', 'unemployed': 'unemployed', 'pop16': 'pop16'}
    for m in measures:
        out[m] = {'fx': [], 'es': []}
    for y, fx, sh in zip(years, b['fx'], b['share']):
        total = int(prd.loc[y, 'total'])
        share = 100 * fx / total
        # the yearbook's own foreign share (PEX-1, one decimal) must agree with total and foreign count
        if abs(share - sh) > 0.2:
            raise SystemExit(f'{y}: foreign share {share:.2f}% from PRD-1 total vs {sh}% in PEX-1')
        out['recipients']['fx'].append(fx)
        out['recipients']['es'].append(total - fx)
        out['share'].append(round(share, 2))
        idx = [i for i, t in enumerate(e['t']) if t.startswith(str(y))]
        if len(idx) != 4:
            raise SystemExit(f'EPA: {len(idx)} quarters in {y}, need 4 for an annual average')
        for m in measures:
            for g in ('fx', 'es'):
                out[m][g].append(round(sum(e[m][g][i] for i in idx) / 4))
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('series_file', type=Path)
    ap.add_argument('snapshot_file', type=Path)
    ap.add_argument('out', type=Path)
    ap.add_argument('--eustat', type=Path, default=Path('data/eustat_tbl0021467_2025-01-01.csv'),
                    help='Eustat tbl0021467 CSV (continent of birth); skipped with a warning if missing')
    ap.add_argument('--census', type=Path, default=Path('data/ine_67129_censo_actividad_pais_nacimiento.px'),
                    help='INE table 67129 (.px); written to census.json next to OUT; skipped with a warning if missing')
    ap.add_argument('--census-nat', type=Path, default=Path('data/ine_67130_censo_actividad_nacionalidad.px'),
                    help='INE table 67130 (.px, by nationality): Basque adults by nationality for the RGI card')
    ap.add_argument('--rgi-origin', type=Path, default=Path('data/rgi_titulares_2023_origen.csv'),
                    help='RGI holders 2023 by nationality and area of birth (see the .README.txt next to it)')
    ap.add_argument('--sepe-report', type=Path, default=Path('data/sepe_informe_extranjeros_2025.csv'),
                    help="figures typed in from SEPE's 2025 report on foreign workers (see the .README.txt next to it)")
    ap.add_argument('--epa-pop', type=Path, default=Path('data/ine_65290_epa_poblacion16_nacionalidad_ccaa.px'))
    ap.add_argument('--epa-active', type=Path, default=Path('data/ine_65299_epa_activos_nacionalidad_ccaa.px'))
    ap.add_argument('--epa-employed', type=Path, default=Path('data/ine_65309_epa_ocupados_nacionalidad_ccaa.px'))
    ap.add_argument('--prd', type=Path, default=Path('data/mites_prd1_beneficiarios_2013_2024.csv'),
                    help='All unemployment-benefit recipients by year (Anuario de Estadísticas, table PRD-1)')
    ap.add_argument('--sectors', type=Path, default=Path('data/ine_69964_censo_ocupados_actividad_nacionalidad.px'),
                    help='INE census table 69964 (employed by economic section and nationality)')
    args = ap.parse_args()
    # regional.json: maps and sectors, loaded by the page after the rest
    regional = {}
    if all(p.exists() for p in (args.epa_pop, args.epa_active, args.epa_employed)):
        regional['epa'] = epa_regions(args.epa_pop, args.epa_active, args.epa_employed)
    if args.census_nat.exists():
        regional['censusNat'] = census_nationality(args.census_nat)
    if args.sectors.exists():
        regional['sectors'] = census_sectors(args.sectors)
    missing = {'epa', 'censusNat', 'sectors'} - set(regional)
    if missing:
        print(f'Warning: regional.json without {sorted(missing)} (input files missing)')
    rpath = args.out.with_name('regional.json')
    rpath.write_text(json.dumps(regional, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    print(f'Wrote {rpath} ({rpath.stat().st_size / 1024:.0f} KB)')
    if args.census.exists():
        cpath = args.out.with_name('census.json')
        cpath.write_text(json.dumps(census_birth(args.census), ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
        print(f'Wrote {cpath} ({cpath.stat().st_size / 1024:.0f} KB)')
    else:
        print(f'Warning: {args.census} not found; no census.json. Download https://www.ine.es/jaxiT3/files/t/es/px/67129.px')
    eustat = args.eustat if args.eustat.exists() else None
    if eustat is None:
        print(f'Warning: {args.eustat} not found; the Basque chart will show born in Spain / born abroad only. '
              'Download it from https://www.eustat.eus/elementos/xls0021467_e.csv')
    out = args.out
    data = build(args.series_file, args.snapshot_file, eustat)
    if args.sepe_report.exists():
        data['snapshots']['sepeReport'] = sepe_report(args.sepe_report)
    else:
        print(f'Warning: {args.sepe_report} not found; no contracts/occupations or scheme-by-sex charts')
    if all(p.exists() for p in (args.rgi_origin, args.census, args.census_nat)):
        data['snapshots']['rgiContext'] = rgi_context(args.rgi_origin, args.census, args.census_nat)
    else:
        print('Warning: RGI context skipped (needs --rgi-origin, --census and --census-nat files)')
    if args.prd.exists():
        data['benefitRates'] = benefit_rates(args.prd, data)
    else:
        print(f'Warning: {args.prd} not found; no benefits-per-100 chart')
    out.write_text(json.dumps(data, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    print(f'Wrote {out} ({out.stat().st_size / 1024:.0f} KB)')


if __name__ == '__main__':
    main()
