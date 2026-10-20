"""Filtro inclusivo de períodos para hojas CSV/XLSX, sin editar el documento fuente.

Soporta fechas seriales de Excel (sistema 1900/1904), ISO y formatos habituales
DD/MM/AAAA (o DD-MM-AAAA). No interpreta números arbitrarios como fechas.
"""
from __future__ import annotations
from datetime import date, datetime, timedelta
import re
from excel_reader import DataError

_BASE_1900 = datetime(1899, 12, 30)
_BASE_1904 = datetime(1904, 1, 1)


def parse_excel_date(raw, date1904=False):
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw
    value = str(raw or '').strip()
    if not value:
        return None
    number = value.replace(',', '.')
    if re.fullmatch(r'\d{4,5}(?:\.\d+)?', number):
        serial = float(number)
        if 20000 <= serial <= 90000:
            return ((_BASE_1904 if date1904 else _BASE_1900) + timedelta(days=serial)).date()
    # ISO, con horas optativas.
    try:
        return date.fromisoformat(value[:10])
    except ValueError:
        pass
    # Formato español, común en registros de laboratorio.
    for pattern in ('%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y',
                    '%d/%m/%y', '%d-%m-%y', '%d.%m.%y'):
        try:
            return datetime.strptime(value[:10], pattern).date()
        except ValueError:
            pass
    return None


def parse_period(start='', end=''):
    start, end = str(start or '').strip(), str(end or '').strip()
    dates = []
    for key, value in (('inicio', start), ('fin', end)):
        if not value:
            dates.append(None)
            continue
        try:
            parsed = date.fromisoformat(value)
        except ValueError:
            raise DataError(f'La fecha de {key} no es válida (formato AAAA-MM-DD).')
        if not 1900 <= parsed.year <= 2100:
            raise DataError(f'El año de la fecha de {key} está fuera del rango permitido.')
        dates.append(parsed)
    if dates[0] and dates[1] and dates[0] > dates[1]:
        raise DataError('La fecha inicial no puede ser posterior a la fecha final.')
    return dates[0], dates[1]


def filter_table_dates(table, start='', end='', date_column=''):
    """Devuelve (tabla separada, metadata). No modifica table ni su Excel."""
    first, last = parse_period(start, end)
    if not first and not last:
        return table, None
    headers = table['headers']
    column = str(date_column or '').strip() or next(
        (h for h in headers if h.casefold().strip() in ('fecha', 'date', 'fecha de ensayo', 'fecha ensayo')), None)
    if column not in headers:
        raise DataError('Para filtrar por período elegí una columna de fecha válida (por ejemplo, Fecha).')
    filtered = []
    invalid_dates = 0
    other_dates = 0
    year_1904 = bool(table.get('date1904', False))
    for row in table['rows']:
        day = parse_excel_date(row.get(column), date1904=year_1904)
        if day is None:
            invalid_dates += 1
            continue
        if (first and day < first) or (last and day > last):
            other_dates += 1
            continue
        copy = dict(row)
        # Solo se convierte la etiqueta de fecha de la tabla en memoria.
        copy[column] = day.strftime('%d/%m/%Y')
        filtered.append(copy)
    if not filtered:
        raise DataError('No hay mediciones con fechas válidas dentro del período seleccionado en esta hoja.')
    result = dict(table, rows=filtered)
    metadata = {
        'start': first.isoformat() if first else None,
        'end': last.isoformat() if last else None,
        'date_column': column,
        'matching_rows': len(filtered),
        'outside_period': other_dates,
        'invalid_dates': invalid_dates,
        'source_rows': len(table['rows']),
    }
    return result, metadata


def attach_date_filter(analysis, metadata):
    if metadata:
        analysis['settings']['date_filter'] = metadata
        if metadata['invalid_dates']:
            analysis['warnings'].append(
                f'{metadata["invalid_dates"]} fila(s) sin fecha legible fueron excluidas del filtro; no se consideran conformes.')
    return analysis
