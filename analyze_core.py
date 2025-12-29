import re
import unicodedata
from pathlib import Path
from typing import Dict, List

import pandas as pd


SHEETS_DEFAULT = {
    "modulo": "BASE ( SE AUT LD)",
    "sap": "BASE ( SAP )",
    "orca": "BASE ( ORÇAFASCIO)",
    "caderno": "BASE ( CADERNO)",
}

COL_COD = "COD_SAP"
COL_DESC = "DESCRICAO"
COL_UN = "UNIDADE"

STOPWORDS_PT = {
    "DE", "DA", "DO", "DAS", "DOS", "E", "EM", "PARA", "COM", "SEM", "NA", "NO", "NAS", "NOS",
    "A", "O", "AS", "OS", "UM", "UMA", "UNS", "UMAS",
}

UNIT_MAP = {
    "UNIDADE": "UN",
    "UNID": "UN",
    "UND": "UN",
    "UN": "UN",
    "PECA": "PC",
    "PEÇA": "PC",
    "PCA": "PC",
    "PC": "PC",
    "PÇ": "PC",
    "METRO": "M",
    "MT": "M",
    "M": "M",
    "M2": "M2",
    "M²": "M2",
    "M 2": "M2",
    "M ²": "M2",
    "M3": "M3",
    "M³": "M3",
    "M 3": "M3",
    "M ³": "M3",
    "LITRO": "L",
    "LT": "L",
    "L": "L",
    "KG": "KG",
    "KILO": "KG",
    "QUILO": "KG",
    "TON": "T",
    "TONELADA": "T",
    "T": "T",
}

ROLE_ALIASES: Dict[str, Dict[str, List[str]]] = {
    "modulo": {
        COL_COD: ["COD_SAP", "CÓD. SAP", "CÓD SAP", "COD SAP", "CÓDIGO SAP", "CODIGO SAP"],
        COL_DESC: [
            "DESCRIÇÃO", "DESCRICAO",
            "DESCRIÇÃO DO MATERIAL / SERVIÇO  MODULO",
            "DESCRIÇÃO DO MATERIAL / SERVIÇO MODULO",
            "DESCRIÇÃO DO MATERIAL / SERVICO MODULO",
            "DESCRICAO DO MATERIAL / SERVICO MODULO",
            "DESCRIÇÃO MODULO", "DESCRICAO MODULO",
        ],
        COL_UN: ["UNIDADE", "UN.", "UNIDADE MODULO", "UNIDADE DO MODULO"],
    },
    "sap": {
        COL_COD: ["COD_SAP", "COD SAP", "CÓD. SAP", "CÓD SAP", "CÓDIGO SAP", "CODIGO SAP"],
        COL_DESC: [
            "DESCRIÇÃO", "DESCRICAO",
            "DESCRIÇÃO DO MATERIAL / SERVIÇO",
            "DESCRIÇÃO DO MATERIAL / SERVICO",
            "DESCRICAO DO MATERIAL / SERVICO",
            "DESCRIÇÃO DO MATERIAL", "DESCRICAO DO MATERIAL",
            "DESCRIÇÃO SAP", "DESCRICAO SAP",
        ],
        COL_UN: ["UNIDADE", "UN.", "UNIDADE SAP", "UNIDADE SAP ", "UNIDADE DO MATERIAL"],
    },
    "orca": {
        COL_COD: ["COD_SAP", "COD SAP", "CÓD. SAP", "CÓD SAP", "CÓDIGO SAP", "CODIGO SAP"],
        COL_DESC: [
            "DESCRIÇÃO", "DESCRICAO",
            "DESCRIÇÃO DO MATERIAL / SERVIÇO",
            "DESCRIÇÃO DO MATERIAL / SERVICO",
            "DESCRICAO DO MATERIAL / SERVICO",
            "DESCRIÇÃO ORÇAFASCIO", "DESCRICAO ORCAFASCIO",
            "DESCRIÇÃO ORÇASFACIO", "DESCRICAO ORCASFACIO",
        ],
        COL_UN: [
            "UNIDADE", "UN.",
            "UNIDADE ORÇAFASCIO", "UNIDADE ORÇAFASCIO ",
            "UNIDADE ORÇASFACIO", "UNIDADE ORÇASFACIO ",
        ],
    },
    "caderno": {
        COL_COD: ["COD_SAP", "COD SAP", "CÓD. SAP", "CÓD SAP", "CÓDIGO SAP", "CODIGO SAP"],
        COL_DESC: [
            "DESCRIÇÃO", "DESCRICAO",
            "DESCRIÇÃO DO MATERIAL / SERVIÇO",
            "DESCRIÇÃO DO MATERIAL / SERVICO",
            "DESCRICAO DO MATERIAL / SERVICO",
            "DESCRIÇÃO CADERNO", "DESCRICAO CADERNO",
            "DESCRIÇÃO CARDERNO", "DESCRICAO CARDERNO",
        ],
        COL_UN: ["UNIDADE", "UN.", "UNIDADE CADERNO", "UNIDADE CARDERNO"],
    },
}


def _norm_header(h: object) -> str:
    s = str(h).strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s).strip()
    return s


def clean_cod(c: object) -> str:
    if c is None:
        return ""
    s = str(c).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    s = s.replace(" ", "")
    return s


def norm_text(s: object) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s).strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[^A-Z0-9\s/\-+]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    toks = [t for t in s.split() if t not in STOPWORDS_PT]
    return " ".join(toks)


def norm_unit(u: object) -> str:
    if u is None or (isinstance(u, float) and pd.isna(u)):
        return ""
    s = str(u).strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s).strip()
    s = s.replace(".", "").replace(",", "")
    return UNIT_MAP.get(s, s)


def ensure_cols_by_role(df: pd.DataFrame, sheet_name: str, role: str) -> pd.DataFrame:
    if role not in ROLE_ALIASES:
        raise ValueError(f"Role inválida: {role}")

    norm_to_original = {_norm_header(c): c for c in df.columns}
    rename_map = {}
    missing = []

    aliases = ROLE_ALIASES[role]
    for canonical, variants in aliases.items():
        found_original = None
        for v in variants:
            v_norm = _norm_header(v)
            if v_norm in norm_to_original:
                found_original = norm_to_original[v_norm]
                break
        if not found_original:
            missing.append(canonical)
        else:
            rename_map[found_original] = canonical

    if missing:
        raise ValueError(
            f"A aba '{sheet_name}' (role={role}) não contém colunas necessárias: {missing}. "
            f"Colunas encontradas: {list(df.columns)}"
        )

    return df.rename(columns=rename_map)


def load_sheet(excel_path: Path, sheet_name: str, role: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=object)
    df = ensure_cols_by_role(df, sheet_name=sheet_name, role=role).copy()

    df[COL_COD] = df[COL_COD].map(clean_cod)
    df[COL_DESC] = df[COL_DESC].astype(str).fillna("")
    df[COL_UN] = df[COL_UN].astype(str).fillna("")

    df = df[df[COL_COD].astype(str).str.len() > 0].copy()
    df = df.drop_duplicates(subset=[COL_COD], keep="first").reset_index(drop=True)
    return df


def build_analysis(modulo: pd.DataFrame, sap: pd.DataFrame, orca: pd.DataFrame, caderno: pd.DataFrame) -> pd.DataFrame:
    m = modulo[[COL_COD, COL_DESC, COL_UN]].rename(columns={
        COL_DESC: "desc_modulo_raw",
        COL_UN: "un_modulo_raw",
    }).copy()

    sap2 = sap[[COL_COD, COL_DESC, COL_UN]].rename(columns={
        COL_DESC: "sap_desc_raw",
        COL_UN: "sap_un_raw",
    })
    orca2 = orca[[COL_COD, COL_DESC, COL_UN]].rename(columns={
        COL_DESC: "orca_desc_raw",
        COL_UN: "orca_un_raw",
    })
    cad2 = caderno[[COL_COD, COL_DESC, COL_UN]].rename(columns={
        COL_DESC: "cad_desc_raw",
        COL_UN: "cad_un_raw",
    })

    out = (
        m.merge(sap2, on=COL_COD, how="left")
         .merge(orca2, on=COL_COD, how="left")
         .merge(cad2, on=COL_COD, how="left")
    )

    out["existe_no_sap"] = out["sap_desc_raw"].notna()
    out["existe_no_orcafascio"] = out["orca_desc_raw"].notna()
    out["existe_no_caderno"] = out["cad_desc_raw"].notna()

    out["desc_modulo_norm"] = out["desc_modulo_raw"].map(norm_text)
    out["sap_desc_norm"] = out["sap_desc_raw"].map(norm_text)
    out["orca_desc_norm"] = out["orca_desc_raw"].map(norm_text)
    out["cad_desc_norm"] = out["cad_desc_raw"].map(norm_text)

    out["un_modulo_norm"] = out["un_modulo_raw"].map(norm_unit)
    out["sap_un_norm"] = out["sap_un_raw"].map(norm_unit)
    out["orca_un_norm"] = out["orca_un_raw"].map(norm_unit)
    out["cad_un_norm"] = out["cad_un_raw"].map(norm_unit)

    out["match_desc_modulo_sap"] = (out["desc_modulo_norm"] != "") & (out["desc_modulo_norm"] == out["sap_desc_norm"])
    out["match_desc_modulo_orca"] = (out["desc_modulo_norm"] != "") & (out["desc_modulo_norm"] == out["orca_desc_norm"])
    out["match_desc_modulo_caderno"] = (out["desc_modulo_norm"] != "") & (out["desc_modulo_norm"] == out["cad_desc_norm"])

    out["match_un_modulo_sap"] = (out["un_modulo_norm"] != "") & (out["un_modulo_norm"] == out["sap_un_norm"])
    out["match_un_modulo_orca"] = (out["un_modulo_norm"] != "") & (out["un_modulo_norm"] == out["orca_un_norm"])
    out["match_un_modulo_caderno"] = (out["un_modulo_norm"] != "") & (out["un_modulo_norm"] == out["cad_un_norm"])

    def status_desc(row) -> str:
        if not (row["existe_no_sap"] or row["existe_no_orcafascio"] or row["existe_no_caderno"]):
            return "NAO_ENCONTRADO_EM_NENHUMA_BASE"
        ok_any = bool(row["match_desc_modulo_sap"] or row["match_desc_modulo_orca"] or row["match_desc_modulo_caderno"])
        return "OK" if ok_any else "DIVERGENTE"

    def status_un(row) -> str:
        if not (row["existe_no_sap"] or row["existe_no_orcafascio"] or row["existe_no_caderno"]):
            return "NAO_ENCONTRADO_EM_NENHUMA_BASE"
        ok_any = bool(row["match_un_modulo_sap"] or row["match_un_modulo_orca"] or row["match_un_modulo_caderno"])
        return "OK" if ok_any else "DIVERGENTE"

    out["status_desc"] = out.apply(status_desc, axis=1)
    out["status_un"] = out.apply(status_un, axis=1)

    return out.sort_values([COL_COD]).reset_index(drop=True)


def build_errors(analysis: pd.DataFrame) -> pd.DataFrame:
    mask = (analysis["status_desc"] != "OK") | (analysis["status_un"] != "OK")
    err = analysis.loc[mask].copy()

    def classify(row) -> str:
        if row["status_desc"] == "NAO_ENCONTRADO_EM_NENHUMA_BASE":
            return "NAO_ENCONTRADO"
        flags = []
        if row["status_desc"] == "DIVERGENTE":
            flags.append("DESC")
        if row["status_un"] == "DIVERGENTE":
            flags.append("UN")
        return "DIVERGENTE_" + "_".join(flags) if flags else "OUTRO"

    err["tipo_erro"] = err.apply(classify, axis=1)
    return err


def build_resumo(analysis: pd.DataFrame) -> pd.DataFrame:
    return (
        analysis.groupby(["status_desc", "status_un"], dropna=False)
        .size()
        .reset_index(name="qtd")
        .sort_values("qtd", ascending=False)
        .reset_index(drop=True)
    )


def export_excel(out_path: Path, analysis: pd.DataFrame, errors: pd.DataFrame, resumo: pd.DataFrame) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        resumo.to_excel(w, index=False, sheet_name="resumo")
        errors.to_excel(w, index=False, sheet_name="erros")
        analysis.to_excel(w, index=False, sheet_name="analysis")


def run_analysis_file(excel_path: Path, out_path: Path,
                      sheet_modulo: str = SHEETS_DEFAULT["modulo"],
                      sheet_sap: str = SHEETS_DEFAULT["sap"],
                      sheet_orca: str = SHEETS_DEFAULT["orca"],
                      sheet_caderno: str = SHEETS_DEFAULT["caderno"]) -> Path:
    modulo = load_sheet(excel_path, sheet_modulo, role="modulo")
    sap = load_sheet(excel_path, sheet_sap, role="sap")
    orca = load_sheet(excel_path, sheet_orca, role="orca")
    caderno = load_sheet(excel_path, sheet_caderno, role="caderno")

    analysis = build_analysis(modulo, sap, orca, caderno)
    errors = build_errors(analysis)
    resumo = build_resumo(analysis)

    export_excel(out_path, analysis=analysis, errors=errors, resumo=resumo)
    return out_path
