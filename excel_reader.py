"""Lector ligero de XLSX/CSV con bibliotecas estándar de Python.

No ejecuta macros ni conserva archivos originales. Soporta libros XLSX comunes,
con varias hojas, cadenas compartidas y cadenas inline.
"""
from __future__ import annotations

import csv
import io
import posixpath
import re
import zipfile
from pathlib import PurePosixPath
from xml.etree import ElementTree as ET

MAX_ROWS = 10000
MAX_UNCOMPRESSED = 48 * 1024 * 1024
MAX_COLUMNS = 100
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CELL_RE = re.compile(r"^([A-Z]+)")


class DataError(ValueError):
    pass


def _xlsx_sheets(z):
    try:
        workbook = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    except (KeyError, ET.ParseError) as e:
        raise DataError("El archivo Excel no tiene una estructura XLSX válida.") from e
    by_id = {}
    for rel in rels.findall(f"{{{NS_PKG_REL}}}Relationship"):
        tgt = rel.get("Target", "")
        if not tgt or ".." in PurePosixPath(tgt).parts or re.match(r"^[a-z]+:", tgt, re.I):
            continue
        path = tgt.lstrip("/") if tgt.startswith("/") else posixpath.normpath(posixpath.join("xl", tgt))
        by_id[rel.get("Id")] = path
    sheets = []
    for sh in workbook.findall(f".//{{{NS_MAIN}}}sheet"):
        rid = sh.get(f"{{{NS_REL}}}id")
        path = by_id.get(rid)
        if path in z.namelist():
            sheets.append((sh.get("name", "Hoja"), path))
    if not sheets:
        raise DataError("No se encontraron hojas de datos legibles en el archivo.")
    return sheets


def _strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    result = []
    for si in root.findall(f"{{{NS_MAIN}}}si"):
        result.append("".join(t.text or "" for t in si.iter(f"{{{NS_MAIN}}}t")))
    return result


def _column_index(ref):
    m = CELL_RE.match(ref or "")
    if not m:
        return None
    idx = 0
    for c in m.group(1):
        idx = idx * 26 + ord(c) - ord("A") + 1
    return idx - 1


def _excel_rows(z, path, shared):
    root = ET.fromstring(z.read(path))
    data = root.find(f"{{{NS_MAIN}}}sheetData")
    if data is None:
        return []
    output = []
    for row in data.findall(f"{{{NS_MAIN}}}row"):
        entries = {}
        for cell in row.findall(f"{{{NS_MAIN}}}c"):
            index = _column_index(cell.get("r"))
            if index is None or index >= MAX_COLUMNS:
                continue
            typ = cell.get("t", "")
            value_element = cell.find(f"{{{NS_MAIN}}}v")
            raw = "" if value_element is None else (value_element.text or "")
            if typ == "s":
                try:
                    value = shared[int(raw)]
                except (ValueError, IndexError):
                    value = ""
            elif typ == "inlineStr":
                is_node = cell.find(f"{{{NS_MAIN}}}is")
                value = "" if is_node is None else "".join(t.text or "" for t in is_node.iter(f"{{{NS_MAIN}}}t"))
            elif typ == "b":
                value = "Sí" if raw == "1" else "No"
            elif typ in ("e", "str"):
                value = raw if typ == "str" else ""
            else:
                value = raw
            entries[index] = str(value).strip()
        if not entries:
            continue
        width = min(max(entries) + 1, MAX_COLUMNS)
        output.append([entries.get(i, "") for i in range(width)])
        if len(output) > MAX_ROWS + 22:
            raise DataError(f"El archivo supera el límite de {MAX_ROWS:,} filas de datos por análisis.")
    return output


def _csv_rows(content):
    text = None
    for enc in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            text = content.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise DataError("No pude leer la codificación del CSV.")
    try:
        dialect = csv.Sniffer().sniff(text[:8192], delimiters=",;\t|")
        delimiter = dialect.delimiter
    except csv.Error:
        delimiter = ";" if text.count(";") > text.count(",") else ","
    rdr = csv.reader(io.StringIO(text), delimiter=delimiter)
    rows = []
    for index, r in enumerate(rdr):
        if index == 0 and len(r) == 1 and r[0].lower().startswith("sep="):
            delimiter = r[0][4:5]
            rdr = csv.reader(io.StringIO("\n".join(text.splitlines()[1:])), delimiter=delimiter)
            rows = [r[:MAX_COLUMNS] for r in rdr if any(c.strip() for c in r)]
            break
        if any(str(c).strip() for c in r):
            rows.append([str(c).strip() for c in r[:MAX_COLUMNS]])
        if len(rows) > MAX_ROWS + 22:
            raise DataError(f"El archivo supera el límite de {MAX_ROWS:,} filas de datos por análisis.")
    return rows


def _normalize_table(rows, sheet_name=""):
    if not rows:
        raise DataError("Esta hoja no contiene mediciones. Seleccioná una hoja de ensayos.")
    # Se permite un título encima de la tabla: se toma la primera fila con dos celdas.
    # Los reportes de obra contienen portadas y encabezados múltiples.
    # Detectamos la fila real de encabezados usando los campos de ensayo.
    candidates = []
    for i, row in enumerate(rows[:55]):
        norm = [str(x).strip().casefold() for x in row]
        score = (5 if any(x == "fecha" for x in norm) else 0)
        score += (3 if any("laboratorio" in x for x in norm) else 0)
        score += (2 if any("muestra" in x for x in norm) else 0)
        score += (2 if any("proveedor" in x for x in norm) else 0)
        score += min(5, sum(bool(x) for x in norm))
        if score >= 9:
            candidates.append((score, i))
    if sheet_name.upper() in ("GRANULOMETRIA", "VOLUMETRÍA", "TEMPERATURA", "COMPACTACION", "PG, RECUP. ELÁSTICA Y PUNTO A.", "RUEDA DE HAMBURGO", "TENSION DIAM Y RETENIDA", "RECICLADO DEL PAV") and candidates:
        head_at = max(candidates, key=lambda x: (x[0], -x[1]))[1]
    else:
        head_at = next((i for i, r in enumerate(rows[:30]) if sum(bool(str(s).strip()) for s in r) >= 2), 0)
    raw_headers = list(rows[head_at])
    if sheet_name.upper() == "GRANULOMETRIA" and head_at + 1 < len(rows):
        sub = rows[head_at + 1]
        for j in range(5, min(18, len(sub))):
            if sub[j].strip():
                raw_headers += [""] * max(0, j + 1 - len(raw_headers))
                raw_headers[j] = f"Tamiz {sub[j].strip()}"
    if not raw_headers:
        raise DataError("No encontré encabezados de columnas.")
    headers = []
    used = set()
    for i, raw in enumerate(raw_headers):
        name = raw.strip() or f"Columna {i + 1}"
        original = name
        count = 2
        while name.casefold() in used:
            name = f"{original} ({count})"
            count += 1
        used.add(name.casefold())
        headers.append(name)
    table = []
    fecha_col = next((j for j, name in enumerate(headers) if name.casefold() == "fecha"), None)
    structured_lab = sheet_name.upper() in ("GRANULOMETRIA", "VOLUMETRÍA", "TEMPERATURA", "COMPACTACION", "PG, RECUP. ELÁSTICA Y PUNTO A.", "RUEDA DE HAMBURGO", "TENSION DIAM Y RETENIDA", "RECICLADO DEL PAV")
    for i, row in enumerate(rows[head_at + 1:], start=head_at + 2):
        if structured_lab and fecha_col is not None:
            date = str(row[fecha_col]).strip() if len(row) > fecha_col else ""
            # Las filas de especificaciones, totales y títulos no son ensayos.
            if not date or (not re.match(r"^\d+(?:\.\d+)?$", date) and not re.search(r"\d{1,4}[-/]\d{1,2}[-/]\d{1,4}", date)):
                continue
        if not any(x.strip() for x in row):
            continue
        table.append({"_excel_row": i, **{name: (row[j] if j < len(row) else "") for j, name in enumerate(headers)}})
    if not table:
        raise DataError("Encontré encabezados, pero no filas de datos debajo.")
    if len(table) > MAX_ROWS:
        raise DataError(f"Solo se permiten {MAX_ROWS:,} filas de datos por análisis.")
    return headers, table


def load_data(content: bytes, filename: str, sheet_name: str | None = None):
    if filename.lower().endswith(".csv"):
        headers, data = _normalize_table(_csv_rows(content), "CSV")
        return {"sheets": ["CSV"], "selected_sheet": "CSV", "headers": headers, "rows": data}
    if not filename.lower().endswith(".xlsx"):
        raise DataError("Admitimos .xlsx y .csv. Guardá los archivos .xls antiguos como .xlsx.")
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as z:
            info = z.infolist()
            if sum(x.file_size for x in info) > MAX_UNCOMPRESSED:
                raise DataError("El Excel descomprimido es demasiado grande (máximo 48 MB).")
            sheet_map = _xlsx_sheets(z)
            names = [n for n, _ in sheet_map]
            selected = sheet_name or ("GRANULOMETRIA" if "GRANULOMETRIA" in names else names[0])
            if selected not in names:
                raise DataError("La hoja seleccionada no existe en este archivo.")
            path = dict(sheet_map)[selected]
            rows = _excel_rows(z, path, _strings(z))
            headers, data = _normalize_table(rows, selected)
            return {"sheets": names, "selected_sheet": selected, "headers": headers, "rows": data}
    except (zipfile.BadZipFile, ET.ParseError) as e:
        raise DataError("El Excel está dañado o no es un XLSX válido.") from e


def as_number(raw):
    if raw is None:
        return None
    value = str(raw).strip().replace("\u00a0", "").replace(" ", "")
    if not value:
        return None
    # 1.234,56 y 1,234.56; 12,5; 12.5
    if "," in value and "." in value:
        if value.rfind(",") > value.rfind("."):
            value = value.replace(".", "").replace(",", ".")
        else:
            value = value.replace(",", "")
    elif "," in value:
        value = value.replace(",", ".")
    try:
        result = float(value)
        if abs(result) == float("inf") or result != result:
            return None
        return result
    except ValueError:
        return None


def inspect_data(table):
    sample = table["rows"][:min(300, len(table["rows"]))]
    columns = []
    for col in table["headers"]:
        numeric = sum(as_number(row.get(col)) is not None for row in sample)
        columns.append({"name": col, "numeric_count": numeric, "total_sample": len(sample), "numeric": numeric >= 2})
    return {
        "sheets": table["sheets"], "selected_sheet": table["selected_sheet"],
        "columns": columns, "row_count": len(table["rows"]),
        "preview": [{k: r.get(k, "") for k in table["headers"][:6]} for r in table["rows"][:5]],
    }
