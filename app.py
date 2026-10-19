"""Calibra | Sistema web de cartas de control estadístico (I-MR)."""
import csv
import hashlib
import hmac
import io
import os
import secrets
from datetime import datetime, timezone
from pathlib import Path

from flask import Flask, Response, jsonify, redirect, render_template, request, send_file, session, url_for
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy.orm import load_only
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.utils import secure_filename

from excel_reader import DataError, inspect_data, load_data
from spc import analyze
from ppt_report import create_pptx, operational_status
from batch_report import analyze_g16

BASE_DIR = Path(__file__).resolve().parent
(BASE_DIR / "data").mkdir(exist_ok=True)

app = Flask(__name__)
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY") or secrets.token_hex(32)
app.config["MAX_CONTENT_LENGTH"] = 12 * 1024 * 1024
app.config["JSON_SORT_KEYS"] = False
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.getenv("RENDER", "").lower() == "true"
db_uri = os.getenv("DATABASE_URL", f"sqlite:///{BASE_DIR / 'data' / 'analyses.db'}")
if db_uri.startswith("postgres://"):
    db_uri = "postgresql+psycopg2://" + db_uri[len("postgres://"):]
app.config["SQLALCHEMY_DATABASE_URI"] = db_uri
DB = SQLAlchemy(app)
PASSWORD = os.getenv("APP_PASSWORD", "").strip()


class Analysis(DB.Model):
    id = DB.Column(DB.Integer, primary_key=True)
    created_at = DB.Column(DB.DateTime(timezone=True), nullable=False, default=lambda: datetime.now(timezone.utc))
    filename = DB.Column(DB.String(250), nullable=False)
    metric = DB.Column(DB.String(140), nullable=False)
    sheet = DB.Column(DB.String(140), nullable=False)
    result = DB.Column(DB.JSON, nullable=False)


class FileArchive(DB.Model):
    """Cada análisis exitoso conserva su Excel y el grupo de resultados asociado.

    El BLOB reside en la misma base de datos para no depender del disco efímero
    de Render. El PPTX se regenera bajo demanda a partir de resultados guardados.
    """
    id = DB.Column(DB.Integer, primary_key=True)
    created_at = DB.Column(DB.DateTime(timezone=True), nullable=False, default=lambda: datetime.now(timezone.utc))
    filename = DB.Column(DB.String(250), nullable=False)
    file_size = DB.Column(DB.Integer, nullable=False)
    sha256 = DB.Column(DB.String(64), nullable=False)
    source_bytes = DB.Column(DB.LargeBinary, nullable=False)
    kind = DB.Column(DB.String(30), nullable=False)
    analysis_ids = DB.Column(DB.JSON, nullable=False)
    skipped = DB.Column(DB.JSON, nullable=False, default=list)
    manifest = DB.Column(DB.JSON, nullable=False)


def _archive(filename, content, kind, records, skipped=None):
    """Llamar antes de commit, después de flush para obtener los IDs."""
    counters = {'VERDE': 0, 'AMARILLO': 0, 'ROJO': 0, 'NO EVALUABLE': 0}
    for record in records:
        counters[operational_status(record.result)[0]] += 1
    manifest = dict(count=len(records),
                    sample_count=sum(a.result['summary']['total'] for a in records),
                    statuses=counters,
                    analyses=[_summary_record(a) for a in records])
    entry = FileArchive(filename=filename, file_size=len(content),
                        sha256=hashlib.sha256(content).hexdigest(),
                        source_bytes=content, kind=kind,
                        analysis_ids=[record.id for record in records],
                        skipped=skipped or [], manifest=manifest)
    DB.session.add(entry)
    DB.session.flush()
    return entry


def _get_file_analyses(entry):
    ids = entry.analysis_ids
    if not ids:
        return []
    found = DB.session.execute(DB.select(Analysis).where(Analysis.id.in_(ids))).scalars().all()
    mapping = {a.id: a for a in found}
    return [mapping[id_] for id_ in ids if id_ in mapping]


def _archive_summary(entry):
    return dict(id=entry.id, filename=entry.filename, created_at=entry.created_at.isoformat(),
                size=entry.file_size, sha256=entry.sha256, kind=entry.kind,
                skipped=entry.skipped or [], **entry.manifest)


with app.app_context():
    DB.create_all()


@app.before_request
def guard():
    if request.path.startswith("/static/") or request.path in ("/login", "/health"):
        return None
    if PASSWORD and not session.get("ok"):
        if request.path.startswith("/api/"):
            return jsonify(error="Tu sesión expiró. Iniciá sesión nuevamente."), 401
        return redirect(url_for("login"))
    # Evita solicitudes cross-site simples para acciones API autenticadas.
    if request.path.startswith("/api/") and request.method in ("POST", "DELETE"):
        if request.headers.get("X-Requested-With") != "XMLHttpRequest":
            return jsonify(error="Solicitud no autorizada."), 403


@app.get("/health")
def health():
    return jsonify(status="ok")


@app.route("/login", methods=["GET", "POST"])
def login():
    if not PASSWORD:
        return redirect(url_for("home"))
    error = None
    if request.method == "POST":
        candidate = request.form.get("password", "")
        if hmac.compare_digest(candidate, PASSWORD):
            session.clear()
            session["ok"] = True
            return redirect(url_for("home"))
        error = "Contraseña incorrecta. Intentá nuevamente."
    return render_template("login.html", error=error)


@app.post("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.get("/")
def home():
    return render_template("index.html", protected=bool(PASSWORD))


def _uploaded_data():
    upload = request.files.get("file")
    if not upload or not upload.filename:
        raise DataError("Seleccioná un archivo Excel o CSV para continuar.")
    filename = secure_filename(upload.filename)
    if not filename:
        raise DataError("El nombre del archivo no es válido.")
    if not filename.lower().endswith((".xlsx", ".csv")):
        raise DataError("Solo se permiten archivos .xlsx o .csv.")
    content = upload.read()
    if not content:
        raise DataError("El archivo está vacío.")
    return filename, load_data(content, filename, request.form.get("sheet") or None), content


def _limit(raw, label):
    value = str(raw or "").strip().replace(" ", "")
    if not value:
        return None
    value = value.replace(",", ".")
    try:
        number = float(value)
    except ValueError:
        raise DataError(f"{label} debe ser un número válido.")
    if abs(number) == float("inf") or number != number:
        raise DataError(f"{label} debe ser un número finito.")
    return number


@app.post("/api/inspect")
def inspect():
    try:
        filename, table, _ = _uploaded_data()
        return jsonify(filename=filename, **inspect_data(table))
    except (DataError, ValueError) as ex:
        return jsonify(error=str(ex)), 400


@app.post("/api/analyze")
def new_analysis():
    try:
        filename, table, content = _uploaded_data()
        lsl = _limit(request.form.get("lsl"), "El límite inferior")
        usl = _limit(request.form.get("usl"), "El límite superior")
        if lsl is not None and usl is not None and lsl >= usl:
            raise DataError("El límite superior debe ser mayor que el límite inferior.")
        result = analyze(table, request.form.get("metric", ""),
                         request.form.get("label_column") or None, lsl, usl)
        record = Analysis(filename=filename, metric=result["settings"]["metric"],
                          sheet=result["settings"]["sheet"], result=result)
        DB.session.add(record)
        DB.session.flush()
        archive = _archive(filename, content, 'Individual', [record])
        DB.session.commit()
        return jsonify(id=record.id, archive_id=archive.id, filename=filename,
                       created_at=record.created_at.isoformat(), result=result)
    except (DataError, ValueError) as ex:
        DB.session.rollback()
        return jsonify(error=str(ex)), 400


@app.post("/api/analyze-all")
def analyze_all_g16():
    """Analiza nueve ensayos de la estructura G16 y guarda un lote para PPTX."""
    try:
        upload = request.files.get("file")
        if not upload or not upload.filename:
            raise DataError("Seleccioná un Excel para generar el informe general.")
        filename = secure_filename(upload.filename)
        if not filename.lower().endswith('.xlsx'):
            raise DataError("Para el informe G16 automático necesitás un archivo XLSX.")
        raw = upload.read()
        if not raw:
            raise DataError("El archivo está vacío.")
        results, skipped = analyze_g16(raw, filename)
        saved = []
        for result in results:
            record = Analysis(filename=filename, metric=result['settings']['metric'],
                              sheet=result['settings']['sheet'], result=result)
            DB.session.add(record)
            saved.append(record)
        DB.session.flush()
        archive = _archive(filename, raw, 'G16 completo', saved, skipped)
        DB.session.commit()
        ids = [a.id for a in saved]
        return jsonify(ok=True, count=len(ids), skipped=skipped, ids=ids,
                       archive_id=archive.id,
                       pptx_url=url_for('archive_pptx', archive_id=archive.id),
                       first={'id':saved[0].id, 'result':saved[0].result, 'filename':filename})
    except (DataError, ValueError) as exc:
        DB.session.rollback()
        return jsonify(error=str(exc)), 400


def _summary_record(record):
    s = record.result["summary"]
    return {"id": record.id, "filename": record.filename, "metric": record.metric,
            "sheet": record.sheet, "created_at": record.created_at.isoformat(),
            "verdict": record.result["verdict"], "total": s["total"],
            "outside": s["below"] + s["above"], "omitted": s["omitted"]}


@app.get('/api/library')
def library():
    """Biblioteca ordenada por fecha; no mezcla cargas independientes del mismo nombre."""
    entries = DB.session.execute(
        DB.select(FileArchive).options(load_only(
            FileArchive.id, FileArchive.created_at, FileArchive.filename,
            FileArchive.file_size, FileArchive.sha256, FileArchive.kind,
            FileArchive.skipped, FileArchive.manifest,
        )).order_by(FileArchive.id.desc()).limit(150)
    ).scalars().all()
    return jsonify(items=[_archive_summary(a) for a in entries])


@app.get('/api/library/<int:archive_id>/pptx')
def archive_pptx(archive_id):
    entry = DB.get_or_404(FileArchive, archive_id)
    records = _get_file_analyses(entry)
    if not records:
        return jsonify(error='Los análisis de este archivo ya no están disponibles.'), 404
    pptx = create_pptx(records, title=f'Informe de calidad — {entry.filename}')
    return send_file(pptx,
                     mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                     as_attachment=True,
                     download_name=f'calibra-archivo-{archive_id}.pptx')


@app.get('/api/library/<int:archive_id>/excel')
def archive_excel(archive_id):
    entry = DB.get_or_404(FileArchive, archive_id)
    is_csv = entry.filename.lower().endswith('.csv')
    mimetype = ('text/csv' if is_csv else
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    return send_file(io.BytesIO(entry.source_bytes), mimetype=mimetype,
                     as_attachment=True, download_name=entry.filename)


@app.get('/api/library/compile.pptx')
def compile_library_pptx():
    """Apila varios archivos completos, con subtotales y semáforo por archivo."""
    raw = request.args.get('ids', '')
    parts = raw.split(',')
    if (not raw or len(parts) > 10 or
            any(not x.isdecimal() or int(x) <= 0 for x in parts)):
        return jsonify(error='Seleccioná entre 1 y 10 archivos de la biblioteca.'), 400
    ids = list(dict.fromkeys(int(x) for x in parts))
    entries = DB.session.execute(DB.select(FileArchive).where(FileArchive.id.in_(ids))).scalars().all()
    by_id = {a.id: a for a in entries}
    if len(by_id) != len(ids):
        return jsonify(error='Algún archivo seleccionado ya no existe.'), 404
    groups = []
    for archive_id in ids:
        entry = by_id[archive_id]
        analyses = _get_file_analyses(entry)
        if not analyses:
            return jsonify(error=f'El archivo {entry.filename} no tiene análisis recuperables.'), 409
        groups.append(({'filename': entry.filename, 'created_at': entry.created_at.isoformat(),
                        'kind': entry.kind, 'archive_id': entry.id}, analyses))
    all_records = [a for _, analyses in groups for a in analyses]
    if len(all_records) > 60:
        return jsonify(error='El compilado admite hasta 60 análisis. Seleccioná menos archivos.'), 400
    pptx = create_pptx(all_records, title=f'Compilado de calidad — {len(groups)} archivos',
                       groups=groups if len(groups) > 1 else None)
    return send_file(pptx,
                     mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                     as_attachment=True, download_name='calibra-compilado-archivos.pptx')


@app.delete('/api/library/<int:archive_id>')
def delete_archive(archive_id):
    """Elimina explícitamente la copia archivada y sus análisis (irreversible)."""
    entry = DB.get_or_404(FileArchive, archive_id)
    records = _get_file_analyses(entry)
    for analysis in records:
        DB.session.delete(analysis)
    DB.session.delete(entry)
    DB.session.commit()
    return jsonify(ok=True)


@app.get("/api/history")
def history():
    records = DB.session.execute(DB.select(Analysis).order_by(Analysis.id.desc()).limit(35)).scalars()
    return jsonify(items=[_summary_record(a) for a in records])


@app.get("/api/analysis/<int:analysis_id>")
def saved_analysis(analysis_id):
    record = DB.get_or_404(Analysis, analysis_id)
    return jsonify(**_summary_record(record), result=record.result)


@app.get("/api/analysis/<int:analysis_id>/csv")
def report_csv(analysis_id):
    record = DB.get_or_404(Analysis, analysis_id)
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Numero", "Etiqueta", "Fila origen", "Medicion", "Resultado especificacion", "Dentro control I", "Rango movil"])
    for row in record.result["points"]:
        writer.writerow([row["index"], row["label"], row["row"], row["value"], row["status"],
                         "Sí" if row["in_control"] else "No", "" if row["mr"] is None else row["mr"]])
    if record.result["omitted"]:
        writer.writerow([])
        writer.writerow(["FILAS OMITIDAS", "Etiqueta", "Fila origen", "Valor original"])
        for row in record.result["omitted"]:
            writer.writerow(["", row["label"], row["row"], row["value"]])
    payload = "\ufeff" + output.getvalue()
    return Response(payload, mimetype="text/csv; charset=utf-8",
                    headers={"Content-Disposition": f"attachment; filename=calibra-reporte-{analysis_id}.csv"})


@app.get("/api/analysis/<int:analysis_id>/pptx")
def report_pptx(analysis_id):
    """Diapositivas completas para un análisis (datos almacenados, no Excel original)."""
    record = DB.get_or_404(Analysis, analysis_id)
    pptx = create_pptx([record], title=f"Control de calidad — {record.metric}")
    return send_file(pptx,
                     mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                     as_attachment=True,
                     download_name=f"calibra-analisis-{analysis_id}.pptx")


@app.get("/api/reports/pptx")
def all_reports_pptx():
    """Consolida TODOS los análisis guardados, incluidos de diferentes archivos."""
    selected = request.args.get("ids", "").strip()
    if selected:
        ids = selected.split(',')
        if len(ids) > 80 or any(not i.isdecimal() for i in ids):
            return jsonify(error="La selección de análisis no es válida."), 400
        ids = list(dict.fromkeys(int(i) for i in ids))
        loaded = DB.session.execute(DB.select(Analysis).where(Analysis.id.in_(ids))).scalars().all()
        by_id = {a.id: a for a in loaded}
        if len(by_id) != len(ids):
            return jsonify(error="Uno o más análisis ya no existen."), 404
        records = [by_id[i] for i in ids]
    else:
        records = DB.session.execute(DB.select(Analysis).order_by(Analysis.id.desc())).scalars().all()
    if not records:
        return jsonify(error="Todavía no hay análisis guardados para exportar."), 404
    pptx = create_pptx(records, title="Informe consolidado — Control estadístico de calidad")
    return send_file(pptx,
                     mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                     as_attachment=True,
                     download_name="calibra-informe-consolidado.pptx")


@app.delete("/api/analysis/<int:analysis_id>")
def delete_analysis(analysis_id):
    record = DB.get_or_404(Analysis, analysis_id)
    # Un análisis archivado forma parte de un documento completo: no dejar grupos rotos.
    for group in DB.session.execute(DB.select(FileArchive.id, FileArchive.analysis_ids)).all():
        if analysis_id in group.analysis_ids:
            return jsonify(error='Este análisis pertenece a un Excel archivado. Eliminá el registro completo desde Biblioteca Excel.'), 409
    DB.session.delete(record)
    DB.session.commit()
    return jsonify(ok=True)


@app.errorhandler(RequestEntityTooLarge)
def too_big(_err):
    return jsonify(error="El archivo supera los 12 MB permitidos."), 413


@app.errorhandler(500)
def server_error(_err):
    DB.session.rollback()
    return jsonify(error="Hubo un error inesperado. Comprobá el archivo y reintentá."), 500


if __name__ == "__main__":
    app.run(debug=os.getenv("FLASK_DEBUG") == "1")
