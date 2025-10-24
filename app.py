# app.py (pandas-free)
import os, tempfile
from flask import Flask, request, render_template, send_file, abort
from werkzeug.utils import secure_filename
from parser_presupuesto import parse_pdf, build_xlsx_result

ALLOWED_EXT = {".pdf"}
MAX_CONTENT_LENGTH = 40 * 1024 * 1024  # 40MB per request

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/convertir", methods=["POST"])
def convertir():
    if "files" not in request.files:
        abort(400, "No se envió ningún archivo.")
    files = request.files.getlist("files")
    if not files:
        abort(400, "No se envió ningún archivo.")

    all_rows, all_unmatched = [], []
    with tempfile.TemporaryDirectory() as tmpdir:
        for f in files:
            filename = secure_filename(f.filename or "")
            ext = os.path.splitext(filename)[1].lower()
            if ext not in ALLOWED_EXT:
                abort(400, f"Extensión no permitida: {filename}")
            pdf_path = os.path.join(tmpdir, filename)
            f.save(pdf_path)

            rows, um = parse_pdf(pdf_path)
            all_rows.extend(rows)
            all_unmatched += [(os.path.basename(pdf_path), x) for x in um]

        if not all_rows:
            abort(400, "No se extrajo ninguna partida. Revisa el formato de tus PDFs.")

        xlsx_buffer = build_xlsx_result(all_rows, out_xlsx_name="presupuesto_bd.xlsx")
        return send_file(
            xlsx_buffer,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="presupuesto_bd.xlsx",
        )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=True)
