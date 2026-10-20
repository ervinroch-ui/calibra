"""Cálculos para carta de individuales (I) y rango móvil (MR)."""
from statistics import mean


def analyze(table, metric, label_column=None, lsl=None, usl=None):
    if metric not in table["headers"]:
        raise ValueError("Seleccioná una columna de mediciones válida.")
    if label_column and label_column not in table["headers"]:
        raise ValueError("Seleccioná una columna de etiquetas válida.")
    from excel_reader import as_number

    points = []
    ignored = []
    for row in table["rows"]:
        raw = row.get(metric, "")
        value = as_number(raw)
        label = str(row.get(label_column, "")).strip() if label_column else ""
        label = label or f"Fila {row['_excel_row']}"
        if value is None:
            ignored.append({"label": label, "row": row["_excel_row"], "value": raw})
        else:
            points.append({"label": label[:85], "value": value, "row": row["_excel_row"]})
    if len(points) < 3:
        raise ValueError("Se necesitan al menos 3 mediciones numéricas para elaborar las cartas I-MR.")

    values = [p["value"] for p in points]
    xbar = mean(values)
    moving = [abs(values[i] - values[i-1]) for i in range(1, len(values))]
    mrbar = mean(moving)
    # Constantes para I-MR con rango móvil de 2 observaciones.
    d2 = 1.128
    d4 = 3.267
    i_ucl = xbar + 3 * mrbar / d2
    i_lcl = xbar - 3 * mrbar / d2
    mr_ucl = d4 * mrbar
    below = above = within = spc_violations = 0
    rows = []
    for idx, point in enumerate(points):
        val = point["value"]
        in_specs = (lsl is None or val >= lsl) and (usl is None or val <= usl)
        if lsl is not None and val < lsl:
            below += 1
            status = "Debajo del mínimo"
        elif usl is not None and val > usl:
            above += 1
            status = "Sobre el máximo"
        else:
            within += 1
            status = "Cumple"
        in_control = i_lcl <= val <= i_ucl
        if not in_control:
            spc_violations += 1
        rows.append({**point, "index": idx + 1, "in_specs": in_specs,
                     "in_control": in_control, "status": status,
                     "mr": None if idx == 0 else moving[idx - 1]})
    mr_violations = sum(1 for m in moving if m > mr_ucl)
    total = len(rows)
    spec_defined = lsl is not None or usl is not None
    if not spec_defined:
        verdict = "SIN_LIMITE"
    elif (below + above) > 0:
        verdict = "FUERA_DE_RANGO"
    elif ignored:
        verdict = "REVISAR_OMISIONES"
    else:
        verdict = "CUMPLE"
    return {
        "verdict": verdict, "spec_defined": spec_defined,
        "settings": {"metric": metric, "label_column": label_column,
                     "lsl": lsl, "usl": usl, "sheet": table["selected_sheet"]},
        "summary": {
            "total": total, "within": within, "below": below, "above": above,
            "omitted": len(ignored), "mean": xbar, "minimum": min(values),
            "maximum": max(values), "percent": round(within / total * 100, 2),
            "spc_violations": spc_violations, "mr_violations": mr_violations,
        },
        "limits": {"i_cl": xbar, "i_lcl": i_lcl, "i_ucl": i_ucl,
                   "mr_cl": mrbar, "mr_lcl": 0, "mr_ucl": mr_ucl},
        "points": rows,
        "omitted": ignored,
        "warnings": [
            *( ["Se recomienda reunir al menos 20–25 mediciones para estimar límites de control estables."] if total < 20 else [] ),
            *( [f"Hay {len(ignored)} fila(s) con mediciones vacías o no numéricas. Verificá esos datos antes de aprobar el lote."] if ignored else [] ),
            *( ["El rango móvil promedio es 0: las mediciones son idénticas o casi idénticas."] if mrbar == 0 else [] ),
        ],
    }
