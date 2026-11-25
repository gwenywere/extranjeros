# Foreigners in Spain · work, unemployment, benefits and poverty

A one-page interactive site on foreign nationals in Spain and the Basque Country, built from official statistics: population by nationality and place of birth, labour-market rates, Social Security affiliation, unemployment benefits, the Minimum Living Income (IMV), the Basque minimum income (RGI) and poverty.

Same design as the Nacimientos pages in `../site/` (plain HTML + [D3](https://d3js.org) from a CDN, light and dark mode). Every chart has a tooltip and a data table.

## Sections

- **Who counts as foreign**: nationality × place of birth (1 Oct 2026), and the share born abroad vs the share with foreign nationality since 2002
- **Share of the total**: foreign nationals' share of the adult population, the unemployed, Social Security affiliates, unemployment-benefit recipients and IMV recipients
- **Work and unemployment** (EPA): unemployment, activity and employment rates (Spanish vs foreign, or EU / non-EU); foreigners in work and unemployed; three ways to count unemployed foreigners in 2025; work status by continent or country of birth from the annual census, with switches for sex, year (2021–2024) and region or province
- **Regions and provinces**: a province map from the census (share of adults, unemployment rate or share in work, for foreign nationals, people born abroad, each continent or each of 17 birth countries; click a province to open it in the census chart) and a region map from the EPA (activity, employment and unemployment rates of foreign, EU, non-EU and Spanish nationals, or the foreign − Spanish gap; annual averages 2002–2025)
- **Sectors**: employed foreign nationals by economic section (CNAE) and their share of each sector's workforce, by region or province, sex and year (census 2021–2024)
- **Contracts and occupations** (SEPE's 2025 report on foreign workers and its monthly contract statistics): the 10 occupations with most contracts for foreign nationals and their share of each occupation's contracts; contracts by broad occupational group, foreign vs Spanish; contract type and working time, foreign vs Spanish nationals, with the temporary share by sex or age (SEPE contract statistics, current year to date); temporary and part-time contracts 2016 → 2025
- **Social Security**: foreign affiliates by scheme, EU / non-EU or sex; the 15 largest nationalities, 2012 → 2026
- **Benefits**: foreign unemployment-benefit recipients by type 2013–2024 and on 31 Dec 2025; recipients for every 100 people in work, in the labour force, unemployed or aged 16+, Spanish vs foreign nationals; IMV recipients by nationality, as counts or per 100 households or residents
- **Poverty** (ECV): AROPE, poverty risk, severe deprivation and low work intensity by nationality
- **Basque Country**: unemployment rate (4-quarter average or quarterly), foreign affiliates, work status by continent of birth with sex and province switches (Eustat), RGI coverage

## Files

| File | Contents |
|---|---|
| `index.html` | The page |
| `data.json` | Every series and snapshot the page uses (~70 KB) |
| `regional.json` | EPA by region and nationality (annual averages), census by nationality per province, and census employment by sector (~450 KB, loaded after the rest of the page) |
| `provinces.geo.json` | Simplified province shapes (same file as `../site/`), feature id = INE province code |
| `census.json` | INE table 67129: population 16+ by region/province, sex, country of birth and activity, 2021–2024 (~330 KB, loaded after the rest of the page) |
| `scripts/build_data.py` | Builds `data.json` from the two workbooks |

## Rebuilding the data

Needs `pandas`, `openpyxl` and `python-calamine` (for SEPE's `.xls` workbook). From the project root:

```bash
python site_foreigners/scripts/build_data.py \
  "Extranjeros en España empleo, afiliación, prestaciones y pobreza (datos oficiales).xlsx" \
  Fuentes_estadisticas_inmigracion_Espana_Euskadi.xlsx \
  site_foreigners/data.json
```

The script also reads Eustat's table of the Basque population by continent of birth from `data/eustat_tbl0021467_2025-01-01.csv` (change it with `--eustat <csv>`). Download a newer one from <https://www.eustat.eus/elementos/xls0021467_e.csv>. If the file is missing, the Basque chart falls back to born in Spain / born abroad.

It also writes `census.json` from INE table 67129, read from `data/ine_67129_censo_actividad_pais_nacimiento.px` (change it with `--census <px>`; download a newer one from <https://www.ine.es/jaxiT3/files/t/es/px/67129.px>). The script stops if any group's activity categories don't add up to its total.

The RGI card also uses INE census table 67130 (adults by nationality, `data/ine_67130_censo_actividad_nacionalidad.px`, option `--census-nat`) and RGI holders in 2023 by nationality and area of birth (`data/rgi_titulares_2023_origen.csv`, option `--rgi-origin`). Lanbide's open tables don't include nationality, so those counts are as reported by the press; `data/rgi_titulares_2023_origen.README.txt` gives the source and the checks against the official tables.

The contracts card and the "who the foreign affiliates are" card use figures from SEPE's 2025 report on foreign workers (`4832-1.pdf`), which publishes them only in the PDF. They are typed into `data/sepe_informe_extranjeros_2025.csv` with the page of each figure (option `--sepe-report`); the build stops if they don't add up to the report's own totals. `data/sepe_informe_extranjeros_2025.README.txt` lists the checks.

The benefits-per-100 card needs all recipients each year (Spanish = all − foreign). They come from table PRD-1 of the Ministry of Labour's Anuario de Estadísticas (<https://www.mites.gob.es/es/estadisticas/anuarios/index.htm>, chapter PRD), typed into `data/mites_prd1_beneficiarios_2013_2024.csv` (option `--prd`). The build stops if a year's levels don't add up to its total, or if the total and the foreign count don't give the foreign share in table PEX-1 (to within 0.2 points). Each yearbook has the current and the previous year, so add a row when a new one comes out. The denominators are EPA annual averages (mean of the four quarters).

The IMV card's per-100-households view needs households by nationality, which INE no longer publishes. They are estimated each month from 2021: residents of each nationality ÷ the average size of all-Spanish and all-foreign households in ECEPOV 2021 (`data/ine_56556_ecepov2021_hogares_nacionalidad.csv`, option `--ecepov`), scaled so both add up to INE's total households (`data/ine_60135_hogares_total.csv`, option `--households`; see the README.txt next to them). The build stops if the estimate is more than 10% off INE's total before scaling.

The contract-type card reads SEPE's contract statistics, year-to-date workbook, straight from `data/ESTADISTICA_DE_CONTRATOS_ACUM.xls` (option `--contracts`). Download a newer one from <https://www.sepe.es/HomeSepe/que-es-el-sepe/estadisticas/contratos/estadisticas-nuevas.html> (the month's "acumulado" Excel file) and rebuild; the page takes the period from the file. It uses tables 13.3A, 13.3.1A and 13.3.2A (all contracts by working time, duration and age; men; women), 13.3.3A (the same for foreign nationals) and 43.1A and 43.3A (foreign nationals by EU / non-EU and sex). Spanish = all − foreign. The build stops if the tables are for different months, if working times or ages don't add up to the totals, if men + women don't add up to all, or if the foreign totals differ between tables. SEPE publishes foreign nationals' working time by age but not by sex, so the sex breakdown is permanent / temporary only.

`regional.json` is built from INE tables 65290, 65299, 65309 (EPA by region), 67130 (census by nationality) and 69964 (census employment by sector), all in `data/` (options `--epa-pop`, `--epa-active`, `--epa-employed`, `--census-nat`, `--sectors`). Download newer versions from `https://www.ine.es/jaxiT3/files/t/es/px/<table>.px`.

To view locally, serve the folder (the page loads `data.json`, so opening `index.html` directly won't work):

```bash
cd site_foreigners && python -m http.server 8000   # then open http://localhost:8000
```

## Sources

- INE, Encuesta de Población Activa (tables 65290, 65299, 65300, 65309, 65336)
- Seguridad Social, PX-Web: average affiliates; affiliates by nationality (last day of the month)
- Ministerio de Trabajo, Anuario de Estadísticas, tables PEX-1 and PRD-1 (SEPE data)
- Seguridad Social (INSS), IMV recipients by nationality, published by Maldita.es
- INE, Encuesta de Condiciones de Vida (tables 60258, 10009)
- INE, Estadística Continua de Población
- INE, Censo anual de población, table 67129 (work status by country of birth)
- SEPE, Informe del Mercado de Trabajo de las Personas Extranjeras 2026
- Eustat, table tbl0021467
- Ministerio de Inclusión, indicadores del Marco Estratégico II

Links and caveats are in the page footer. Statistics classify by nationality, not origin: naturalised immigrants count as Spanish.
