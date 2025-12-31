from pathlib import Path
from uuid import uuid4

import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename

from analyze_core import run_analysis_file  # seu core que gera comparacao.xlsx

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

ALLOWED_EXT = {".xlsx"}


def _series_is_ok(s: pd.Series) -> pd.Series:
    """
    Converte uma série 'match_*' para booleano 'ok' de forma robusta.
    Compatível com:
      - bool (True/False)
      - strings: TRUE/FALSE, VERDADEIRO/FALSO
      - strings novas: CORRETO / A VERIFICAR
    """
    if s.dtype == bool:
        return s.fillna(False)

    x = s.fillna("").astype(str).str.strip().str.upper()
    ok_values = {"TRUE", "VERDADEIRO", "CORRETO", "OK", "SIM", "1", "T", "YES"}
    return x.isin(ok_values)


def compute_error_counts_and_scatter(excel_path: Path, max_points: int = 1200) -> dict:
    """
    Lê a aba 'analysis' do comparacao.xlsx e calcula:
      - contagem de divergências por base (desc/un)
      - dados para scatter: (x=desc_err, y=un_err) por COD_SAP

    max_points: limita quantidade de pontos enviados ao front (performance).
    """
    df = pd.read_excel(excel_path, sheet_name="analysis", dtype=object)

    cols_needed = [
        "COD_SAP",
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

    # OK / ERRO por base
    desc_sap_ok = _series_is_ok(df["match_desc_modulo_sap"])
    desc_orca_ok = _series_is_ok(df["match_desc_modulo_orca"])
    desc_cad_ok = _series_is_ok(df["match_desc_modulo_caderno"])

    un_sap_ok = _series_is_ok(df["match_un_modulo_sap"])
    un_orca_ok = _series_is_ok(df["match_un_modulo_orca"])
    un_cad_ok = _series_is_ok(df["match_un_modulo_caderno"])

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

    # Scatter: score por item (0..3)
    desc_err = (~desc_sap_ok).astype(int) + (~desc_orca_ok).astype(int) + (~desc_cad_ok).astype(int)
    un_err = (~un_sap_ok).astype(int) + (~un_orca_ok).astype(int) + (~un_cad_ok).astype(int)

    scatter_df = pd.DataFrame({
        "cod": df["COD_SAP"].fillna("").astype(str),
        "x": desc_err.astype(int),
        "y": un_err.astype(int),
    })

    # Prioriza pontos mais críticos (x+y maior) para visualização, e limita volume
    scatter_df["score"] = scatter_df["x"] + scatter_df["y"]
    scatter_df = scatter_df.sort_values(["score", "x", "y"], ascending=False)

    if len(scatter_df) > max_points:
        scatter_df = scatter_df.head(max_points)

    # Para não sobrecarregar o JSON, envia só o essencial
    scatter_points = scatter_df[["x", "y", "cod"]].to_dict(orient="records")

    # Top críticos (para legenda/tooltip e uso futuro)
    top_criticos = scatter_df.head(12)[["cod", "x", "y", "score"]].to_dict(orient="records")

    return {
        "counts": counts,
        "scatter": {
            "points": scatter_points,
            "top": top_criticos,
            "max_points": int(max_points),
        }
    }


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/analyze-json")
def analyze_json():
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
        payload = compute_error_counts_and_scatter(out_path, max_points=1200)
    except Exception as e:
        return jsonify({"ok": False, "error": f"Erro ao processar: {e}"}), 400

    return jsonify({
        "ok": True,
        "token": token,
        "counts": payload["counts"],
        "scatter": payload["scatter"],
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
