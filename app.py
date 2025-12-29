from pathlib import Path
from uuid import uuid4

import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename

from analyze_core import run_analysis_file

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

ALLOWED_EXT = {".xlsx"}


def compute_error_counts(excel_path: Path) -> dict:
    """
    Lê a aba 'analysis' do comparacao.xlsx e calcula contagens de divergência
    por base e tipo (descrição/unidade).
    """
    df = pd.read_excel(excel_path, sheet_name="analysis", dtype=object)

    # colunas esperadas do seu pipeline
    cols_needed = [
        "match_desc_modulo_sap",
        "match_desc_modulo_orca",
        "match_desc_modulo_caderno",
        "match_un_modulo_sap",
        "match_un_modulo_orca",
        "match_un_modulo_caderno",
    ]
    for c in cols_needed:
        if c not in df.columns:
            raise ValueError(f"Coluna esperada não encontrada em 'analysis': {c}")

    def to_bool(s):
        # garante boolean; match_* já tende a vir como True/False, mas por segurança:
        return s.astype(str).str.lower().isin(["true", "1", "t", "yes", "sim"])

    desc_sap_ok = to_bool(df["match_desc_modulo_sap"])
    desc_orca_ok = to_bool(df["match_desc_modulo_orca"])
    desc_cad_ok = to_bool(df["match_desc_modulo_caderno"])

    un_sap_ok = to_bool(df["match_un_modulo_sap"])
    un_orca_ok = to_bool(df["match_un_modulo_orca"])
    un_cad_ok = to_bool(df["match_un_modulo_caderno"])

    counts = {
        "desc": {
            "SAP": int((~desc_sap_ok).sum()),
            "ORCA": int((~desc_orca_ok).sum()),
            "CADERNO": int((~desc_cad_ok).sum()),
        },
        "un": {
            "SAP": int((~un_sap_ok).sum()),
            "ORCA": int((~un_orca_ok).sum()),
            "CADERNO": int((~un_cad_ok).sum()),
        },
        "total_rows": int(len(df)),
    }
    return counts


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/analyze-json")
def analyze_json():
    """
    Recebe o Excel, processa, gera comparacao.xlsx e devolve:
      - métricas para o gráfico
      - URL para download do Excel
    """
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "Nenhum arquivo enviado."}), 400

    f = request.files["file"]
    if not f or f.filename == "":
        return jsonify({"ok": False, "error": "Selecione um arquivo .xlsx."}), 400

    filename = secure_filename(f.filename)
    ext = Path(filename).suffix.lower()
    if ext not in ALLOWED_EXT:
        return jsonify({"ok": False, "error": "Formato inválido. Envie um arquivo .xlsx."}), 400

    token = uuid4().hex
    in_path = UPLOAD_DIR / f"{token}_{filename}"
    f.save(in_path)

    out_path = OUTPUT_DIR / f"comparacao_{token}.xlsx"
    try:
        run_analysis_file(in_path, out_path)
        counts = compute_error_counts(out_path)
    except Exception as e:
        return jsonify({"ok": False, "error": f"Erro ao processar: {e}"}), 400

    return jsonify({
        "ok": True,
        "token": token,
        "counts": counts,
        "download_url": f"/download/{token}",
    })


@app.get("/download/<token>")
def download(token: str):
    out_path = OUTPUT_DIR / f"comparacao_{token}.xlsx"
    if not out_path.exists():
        return "Arquivo não encontrado ou expirado.", 404

    return send_file(
        out_path,
        as_attachment=True,
        download_name="comparacao.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        max_age=0,
    )


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
