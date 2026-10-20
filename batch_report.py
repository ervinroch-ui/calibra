"""Conjunto de ensayos para el informe del Excel G-16 CORASCO–LLANSA.

Los límites son los que aparecen en las diapositivas de referencia del Grupo 16,
excepto temperatura (indicada por el propio Excel como 145–165 °C).
No se asignan límites a los ensayos que no tienen una especificación confirmada.
"""
from excel_reader import DataError, load_data
from spc import analyze
from date_filter import filter_table_dates, attach_date_filter, parse_period

PRESETS = [
    ('VOLUMETRÍA', 'Va (%)', 3, 5),
    ('VOLUMETRÍA', 'VMA (%)', 14, None),
    ('VOLUMETRÍA', 'VFA (%)', 65, 75),
    ('VOLUMETRÍA', 'Polvo/Asfalto', .6, 1.2),
    ('COMPACTACION', '% Compactación', 92, 96),
    ('TEMPERATURA', 'Temperatura Medida (°C)', 145, 165),
    ('GRANULOMETRIA', 'Tamiz 1/2"', None, None),
    ('RUEDA DE HAMBURGO', 'Promedio huella mm', None, None),
    ('TENSION DIAM Y RETENIDA', '% Relación TSR', None, None),
]


def analyze_g16(content, filename, date_from="", date_to="", date_column=""):
    parse_period(date_from, date_to)
    selected = []
    omitted = []
    cache = {}
    for sheet, metric, lsl, usl in PRESETS:
        try:
            if sheet not in cache:
                cache[sheet] = load_data(content, filename, sheet)
            table = cache[sheet]
            if metric not in table['headers']:
                omitted.append(f'{sheet}: no se encontró {metric}')
                continue
            table, date_meta = filter_table_dates(table, date_from, date_to, date_column)
            label = 'Fecha' if 'Fecha' in table['headers'] else None
            r = analyze(table, metric, label, lsl, usl)
            attach_date_filter(r, date_meta)
            selected.append(r)
        except (DataError, ValueError) as exc:
            omitted.append(f'{sheet} / {metric}: {exc}')
    if not selected:
        raise DataError('No pude reconocer ensayos de la plantilla Grupo 16. Para otros Excel utilizá el análisis individual.')
    return selected, omitted
