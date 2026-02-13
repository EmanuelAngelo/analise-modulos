from pathlib import Path
from uuid import uuid4
import os
from io import StringIO
from datetime import datetime
import re

import pandas as pd
import google.generativeai as genai
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename

# Importa o core existente (Módulos)
from analyze_core import run_analysis_file

app = Flask(__name__)

# --- CONFIGURAÇÃO GEMINI ---
GEMINI_KEY = None  # Inicialmente sem chave

def set_gemini_key(key):
    global GEMINI_KEY
    GEMINI_KEY = key
    try:
        genai.configure(api_key=GEMINI_KEY)
        return True
    except Exception as e:
        return False
from flask import session
from flask import request
# ------------------------------------------------------------------------------
# ROTA PARA SETAR CHAVE GEMINI
# ------------------------------------------------------------------------------
from flask import jsonify
@app.post("/set-gemini-key")
def set_gemini_key_route():
    data = request.get_json()
    key = data.get("key", "").strip()
    if not key:
        return jsonify({"ok": False, "error": "Chave não informada"}), 400
    ok = set_gemini_key(key)
    if ok:
        return jsonify({"ok": True})
    else:
        return jsonify({"ok": False, "error": "Chave inválida ou erro ao configurar Gemini"}), 400
# ---------------------------

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

ALLOWED_EXT = {".xlsx", ".csv"}

# ==============================================================================
# LÓGICA EXISTENTE (MÓDULOS)
# ==============================================================================

def _series_is_ok(s: pd.Series) -> pd.Series:
    if s.dtype == bool:
        return s.fillna(False)
    x = s.fillna("").astype(str).str.strip().str.upper()
    ok_values = {"TRUE", "VERDADEIRO", "CORRETO", "OK", "SIM", "1", "T", "YES"}
    return x.isin(ok_values)


def compute_error_counts_and_scatter(excel_path: Path, max_points: int = 1200) -> dict:
    df = pd.read_excel(excel_path, sheet_name="analysis", dtype=object)

    def get_ok_col(cname):
        if cname not in df.columns:
            return pd.Series([False]*len(df), index=df.index)
        return _series_is_ok(df[cname])

    desc_sap_ok = get_ok_col("match_desc_modulo_sap")
    desc_orca_ok = get_ok_col("match_desc_modulo_orca")
    desc_cad_ok = get_ok_col("match_desc_modulo_caderno")

    un_sap_ok = get_ok_col("match_un_modulo_sap")
    un_orca_ok = get_ok_col("match_un_modulo_orca")
    un_cad_ok = get_ok_col("match_un_modulo_caderno")

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

    desc_err = (~desc_sap_ok).astype(int) + (~desc_orca_ok).astype(int) + (~desc_cad_ok).astype(int)
    un_err = (~un_sap_ok).astype(int) + (~un_orca_ok).astype(int) + (~un_cad_ok).astype(int)

    scatter_df = pd.DataFrame({
        "cod": df["COD_SAP"].fillna("").astype(str),
        "x": desc_err.astype(int),
        "y": un_err.astype(int),
    })

    scatter_df["score"] = scatter_df["x"] + scatter_df["y"]
    scatter_df = scatter_df.sort_values(["score", "x", "y"], ascending=False)

    if len(scatter_df) > max_points:
        scatter_df = scatter_df.head(max_points)

    scatter_points = scatter_df[["x", "y", "cod"]].to_dict(orient="records")
    top_criticos = scatter_df.head(12)[["cod", "x", "y", "score"]].to_dict(orient="records")

    return {
        "counts": counts,
        "scatter": {
            "points": scatter_points,
            "top": top_criticos,
            "max_points": int(max_points),
        }
    }

# ==============================================================================
# LÓGICA DE IA (GERAL)
# ==============================================================================

def parse_csv_to_charts(csv_text: str) -> list:
    charts = {}
    try:
        # 1. Limpeza: Remove blocos ```csv e ``` e espaços extras
        cleaned_text = re.sub(r"```[a-zA-Z]*", "", csv_text).replace("```", "")
        # Remove linhas vazias
        lines = [line.strip() for line in cleaned_text.splitlines() if line.strip()]
        cleaned_text = "\n".join(lines)

        # 2. Tenta ler detectando separador automaticamente
        try:
            # Tenta com vírgula (padrão)
            df = pd.read_csv(StringIO(cleaned_text), sep=",", on_bad_lines='skip')
            # Se colunas forem insuficientes, tenta ponto e vírgula
            if len(df.columns) < 2:
                df = pd.read_csv(StringIO(cleaned_text), sep=";", on_bad_lines='skip')
        except:
            return [{"error": "Falha na leitura do CSV gerado pela IA."}]

        # 3. Normaliza Colunas
        df.columns = [str(c).strip().lower() for c in df.columns]

        # 4. Fallback de Colunas: Se não achar os nomes exatos, assume pela posição (4 primeiras)
        expected = ['charttitle', 'charttype', 'label', 'value']
        
        # Se não tem as colunas obrigatórias 'charttitle' e 'value', mas tem 4 colunas, renomeia
        if not all(k in df.columns for k in ['charttitle', 'value']):
            if len(df.columns) >= 4:
                df = df.iloc[:, :4] # Pega só as 4 primeiras
                df.columns = expected
            else:
                 return [{"error": f"Formato inválido. Colunas: {list(df.columns)}"}]

        # 5. Processa dados
        for title, group in df.groupby('charttitle'):
            ctype = 'bar'
            if 'charttype' in group.columns:
                val = str(group.iloc[0]['charttype']).lower().strip()
                if val in ['bar', 'pie', 'doughnut']: ctype = val
            
            # Garante que Value é numérico
            group['value'] = pd.to_numeric(group['value'], errors='coerce').fillna(0)
            
            charts[title] = {
                "title": str(title),
                "type": ctype,
                "labels": group['label'].astype(str).tolist(),
                "data": group['value'].tolist()
            }
            
        return list(charts.values())

    except Exception as e:
        return [{"error": f"Erro interno IA: {str(e)}"}]

def generate_ai_analysis_modules(analysis_df: pd.DataFrame, errors_df: pd.DataFrame) -> list:
    # (Lógica original de módulos)
    total_linhas = len(analysis_df)
    status_counts = ""
    if 'status_desc' in analysis_df.columns:
        status_counts = analysis_df[['status_desc', 'status_un']].value_counts().to_string()
    
    sample_errors = ""
    if not errors_df.empty:
        cols = [c for c in errors_df.columns if 'match' not in c]
        if not cols: cols = errors_df.columns
        sample_errors = errors_df[cols].head(50).to_csv(index=False)

    prompt = (
        "Analise cadastro de materiais. Gere gráficos CSV.\n"
        f"Dados: {total_linhas} linhas.\nErros: {sample_errors}\n"
        "Saída: ChartTitle,ChartType,Label,Value\n"
    )
    return call_gemini(prompt)

# ==============================================================================
# NOVA LÓGICA: CONTRATOS (COM TRATAMENTO DE MOEDA E DATA)
# ==============================================================================

def clean_currency(val):
    if pd.isna(val) or str(val).strip() == "": return 0.0
    if isinstance(val, (int, float)): return float(val)
    
    # Remove R$ e espaços
    s = re.sub(r'[R\$\s]', '', str(val))
    
    # Detecção simples: se tem vírgula no final, é decimal BR (1000,00)
    # Se tem ponto no final, é decimal US (1000.00)
    try:
        if ',' in s and '.' in s:
            if s.rfind(',') > s.rfind('.'): # BR: 1.000,00
                s = s.replace('.', '').replace(',', '.')
            else: # US: 1,000.00
                s = s.replace(',', '')
        elif ',' in s: # 1000,00
            s = s.replace(',', '.')
        return float(s)
    except:
        return 0.0

def generate_contract_ai_analysis(df: pd.DataFrame) -> list:
    # Normaliza headers
    df.columns = [str(c).strip() for c in df.columns]
    
    total = len(df)
    
    # Busca inteligente de colunas
    cols = df.columns
    c_val = next((c for c in cols if 'validade' in c.lower()), None)
    c_fix = next((c for c in cols if 'fixado' in c.lower()), None)
    c_med = next((c for c in cols if 'medido' in c.lower()), None)
    c_emp = next((c for c in cols if 'empresa' in c.lower()), None)
    
    # Vencidos
    vencidos_info = "N/A"
    if c_val:
        # Tenta converter para data
        df['dt_temp'] = pd.to_datetime(df[c_val], errors='coerce', dayfirst=True)
        hj = datetime.now()
        vencidos = df[df['dt_temp'] < hj]
        vencidos_info = len(vencidos)

    # Financeiro
    total_fix = 0
    total_med = 0
    sample_estourados = ""
    
    if c_fix and c_med:
        df['v_fix'] = df[c_fix].apply(clean_currency)
        df['v_med'] = df[c_med].apply(clean_currency)
        total_fix = df['v_fix'].sum()
        total_med = df['v_med'].sum()
        
        # Filtra estourados
        estourados = df[df['v_med'] > df['v_fix']].copy()
        if not estourados.empty:
            estourados['diff'] = estourados['v_med'] - estourados['v_fix']
            cols_show = [c for c in [c_emp, c_fix, c_med] if c]
            sample_estourados = estourados.nlargest(10, 'diff')[cols_show].to_csv(index=False)

    # Amostra
    sample_data = df.iloc[:30].to_csv(index=False)

    # Aliases para bater com o template do prompt solicitado
    total_contratos = total
    qtd_vencidos = vencidos_info
    total_fixado = total_fix
    total_medido = total_med
    estourados_info = sample_estourados

    # --- PROMPT ATUALIZADO ---
    user_prompt = (
        "Atue como um Cientista de Dados Sênior, especialista em análise financeira, gestão de contratos e "
        "controle orçamentário corporativo.\n\n"
    )
    
    tech_instructions = (
        f"\n\n---  ---\n"
        
    )

    return call_gemini(user_prompt + tech_instructions)

def call_gemini(prompt):
    if not GEMINI_KEY:
        return [{"error": "Chave Gemini não configurada. Informe a chave na tela."}]
    try:
        genai.configure(api_key=GEMINI_KEY)
    except Exception as e:
        return [{"error": f"Erro ao configurar Gemini: {e}"}]
    for m in ['gemini-2.5-flash']:
        try:
            model = genai.GenerativeModel(m)
            resp = model.generate_content(prompt)
            return parse_csv_to_charts(resp.text)
        except Exception as e:
            print(f"Erro modelo {m}: {e}")
            continue
    return [{"error": "IA indisponível."}]

# ==============================================================================
# ROTAS
# ==============================================================================

@app.get("/")
def index():
    return render_template("index.html")

@app.post("/analyze-json")
def analyze_modules():
    if "file" not in request.files: return jsonify({"error":"Sem arquivo"}), 400
    f = request.files["file"]
    if not f: return jsonify({"error":"Sem arquivo"}), 400
    
    fname = secure_filename(f.filename)
    token = uuid4().hex
    in_path = UPLOAD_DIR / f"{token}_{fname}"
    f.save(in_path)
    out_path = OUTPUT_DIR / f"comparacao_{token}.xlsx"

    try:
        run_analysis_file(in_path, out_path)
        payload = compute_error_counts_and_scatter(out_path)
        
        ai_charts = []
        try:
            dfa = pd.read_excel(out_path, sheet_name="analysis", dtype=object)
            try: dfe = pd.read_excel(out_path, sheet_name="erros", dtype=object)
            except: dfe = pd.DataFrame()
            ai_charts = generate_ai_analysis_modules(dfa, dfe)
        except: ai_charts = [{"error": "Erro IA"}]

        return jsonify({
            "ok": True, 
            "type": "modules",
            "counts": payload["counts"], 
            "download_url": f"/download/{token}",
            "ai_charts": ai_charts
        })
    except Exception as e:
        return jsonify({"ok":False, "error":str(e)}), 400

@app.post("/analyze-contracts")
def analyze_contracts():
    if "file" not in request.files: return jsonify({"error":"Sem arquivo"}), 400
    f = request.files["file"]
    if not f: return jsonify({"error":"Sem arquivo"}), 400
    
    fname = secure_filename(f.filename)
    token = uuid4().hex
    in_path = UPLOAD_DIR / f"contrato_{token}_{fname}"
    f.save(in_path)

    try:
        # Lê Excel ou CSV
        if fname.lower().endswith('.csv'):
            try: df = pd.read_csv(in_path, sep=',')
            except: df = pd.read_csv(in_path, sep=';', on_bad_lines='skip')
        else:
            df = pd.read_excel(in_path)
            
        ai_charts = generate_contract_ai_analysis(df)
        
        return jsonify({
            "ok": True,
            "type": "contracts",
            "ai_charts": ai_charts
        })
    except Exception as e:
        return jsonify({"ok":False, "error":f"Erro contratos: {e}"}), 400

@app.get("/download/<token>")
def download(token: str):
    p = OUTPUT_DIR / f"comparacao_{token}.xlsx"
    if not p.exists(): return "404", 404
    return send_file(p, as_attachment=True, download_name="comparacao.xlsx")

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)