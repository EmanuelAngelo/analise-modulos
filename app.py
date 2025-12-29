from pathlib import Path
from uuid import uuid4

from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename

from analyze_core import run_analysis_file

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

ALLOWED_EXT = {".xlsx"}


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/analyze")
def analyze():
    if "file" not in request.files:
        return render_template("index.html", error="Nenhum arquivo enviado.")

    f = request.files["file"]
    if not f or f.filename == "":
        return render_template("index.html", error="Selecione um arquivo .xlsx.")

    filename = secure_filename(f.filename)
    ext = Path(filename).suffix.lower()
    if ext not in ALLOWED_EXT:
        return render_template("index.html", error="Formato inválido. Envie um arquivo .xlsx.")

    # salva upload
    token = uuid4().hex
    in_path = UPLOAD_DIR / f"{token}_{filename}"
    f.save(in_path)

    # gera saída
    out_path = OUTPUT_DIR / f"comparacao_{token}.xlsx"
    try:
        run_analysis_file(in_path, out_path)
    except Exception as e:
        return render_template("index.html", error=f"Erro ao processar: {e}")

    # devolve o arquivo para download imediato
    return send_file(
        out_path,
        as_attachment=True,
        download_name="comparacao.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
