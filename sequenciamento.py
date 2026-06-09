from flask import Flask, request, jsonify, render_template
import csv
import os
from pathlib import Path
import pandas as pd
from datetime import date
import unicodedata
import re

app = Flask(__name__)

# =========================
# PATHS
# =========================
BASE = Path.home() / "Desktop" / "Scraping_Explosao"
CSV_DIR = BASE / "csv"

CSV_OPS = "Report.csv"
CSV_ATRIB = "atribuicoes.csv"
CSV_OPERADORES = "operadores.csv"
CSV_OPS_MANUAL = "ops_manual.csv"
CSV_LT_TABELA_GERAL = "lt_tabela_geral.csv"
CSV_ESTOQUE_COMPRAS = "estoque_itens_comprados.csv"
CSV_SEPARACAO = "separacao.csv"
CSV_ANALISE_MES_A_MES = "explosao mes a mes .csv"

ESTOQUE_COMPRAS_XLSX = Path(r"Y:\Produção\Estoque\Estoque itens comprados.xlsx")
ESTOQUE_COMPRAS_SHEET = "Estoque Geral"


@app.route('/')
def sequenciamento_analise_mes_a_mes():
    return render_template('analise_mes_a_mes.html')

def buscar_produto(produto):
    atribuicoes = carregar_atribuicoes()
    lista = []

    with open(CSV_DIR / CSV_OPS, newline='', encoding='utf-8') as f:
        for row in csv.reader(f):
            if len(row) < 9 or row[4].strip() != produto:
                continue

            qtd_total = safe_int(row[7])
            qtd_prod = safe_int(row[8])

            status = (
                "Inválida" if qtd_total <= 0 else
                "Finalizada" if qtd_prod >= qtd_total else
                "Em aberto"
            )

            op = row[0].strip()

            lista.append({
                "op": op,
                "produto": produto,
                "descricao": row[5].strip(),
                "quantidade_total": qtd_total,
                "quantidade_produzida": qtd_prod,
                "status": status,
                "operador": atribuicoes.get(op)
            })

    lista.sort(key=lambda x: x["quantidade_produzida"])
    return lista


# =========================
# OPS EM ABERTO
# =========================
@app.route('/ops_abertas')
def ops_abertas():
    atribuicoes = carregar_atribuicoes()
    lista = []

    with open(CSV_DIR / CSV_OPS, newline='', encoding='utf-8') as f:
        for row in csv.reader(f):
            if len(row) < 9:
                continue

            qtd_total = safe_int(row[7])
            qtd_prod = safe_int(row[8])

            if qtd_total <= 0 or qtd_prod >= qtd_total:
                continue

            op = row[0].strip()

            lista.append({
                "op": op,
                "produto": row[4].strip(),
                "descricao": row[5].strip(),
                "quantidade_total": qtd_total,
                "quantidade_produzida": qtd_prod,
                "status": "Em aberto",
                "operador": atribuicoes.get(op)
            })

    lista.sort(key=lambda x: x["quantidade_produzida"])
    return jsonify(lista)


# =========================
# OPERADORES POR LINHA
# =========================
@app.route('/operadores/<linha>')
def operadores_por_linha(linha):
    operadores = []

    with open(CSV_DIR / CSV_OPERADORES, newline='', encoding='utf-8') as f:
        for row in csv.reader(f):
            if len(row) >= 2 and row[1].strip().lower() == linha.lower():
                operadores.append(row[0].strip())

    return jsonify(operadores)


# =========================
# ATRIBUIÇÕES
# =========================
def carregar_atribuicoes():
    dados = {}

    if os.path.exists(CSV_DIR / CSV_ATRIB):
        with open(CSV_DIR / CSV_ATRIB, newline='', encoding='utf-8') as f:
            for row in csv.reader(f):
                if len(row) >= 3:
                    dados[row[0].strip()] = row[2].strip()

    return dados


# =========================
# SAFE INT
# =========================
def safe_int(v):
    try:
        return int(str(v).replace(',', '').strip())
    except:
        return 0


def _normalize_text(v):
    s = str(v or "").strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = s.replace("_", " ").replace("-", " ")
    s = " ".join(s.split())
    return s


def _coluna_por_alias(df, aliases):
    cols_norm = {c: _normalize_text(c) for c in df.columns}
    for alias in aliases:
        key = _normalize_text(alias)
        for col, norm in cols_norm.items():
            if key == norm or key in norm:
                return col
    return None


def _is_blank(v):
    if v is None:
        return True
    s = str(v).strip()
    return s == "" or s.lower() in {"nan", "none", "null"}


def _valor_texto(v):
    if _is_blank(v):
        return "-"
    return str(v).strip()


def _to_num_series(series):
    s = series.astype(str).str.strip()
    s = s.str.replace(" ", "", regex=False)
    s = s.str.replace(".", "", regex=False)
    s = s.str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0)


def _atualizar_csv_estoque_compras():
    src = ESTOQUE_COMPRAS_XLSX
    dst = CSV_DIR / CSV_ESTOQUE_COMPRAS

    if not src.exists():
        return None, f"Arquivo não encontrado: {src}"

    try:
        df = pd.read_excel(src, sheet_name=ESTOQUE_COMPRAS_SHEET, dtype=str)
    except ValueError:
        return None, f"Aba '{ESTOQUE_COMPRAS_SHEET}' não encontrada em {src}"
    except Exception as e:
        return None, f"Erro ao ler o Excel de compras: {e}"

    df = df.dropna(how="all")
    if df.empty:
        return None, "A planilha de compras está vazia."

    try:
        CSV_DIR.mkdir(parents=True, exist_ok=True)
        df.to_csv(dst, index=False, encoding="utf-8-sig")
    except Exception as e:
        return None, f"Erro ao gerar CSV de compras: {e}"

    return dst, None


def _carregar_compras_pendentes():
    csv_path, erro = _atualizar_csv_estoque_compras()
    if erro:
        return None, erro

    try:
        df = pd.read_csv(csv_path, dtype=str, sep=None, engine="python")
    except Exception as e:
        return None, f"Erro ao ler CSV de compras: {e}"

    df = df.dropna(how="all")
    if df.empty:
        return [], None

    col_codigo = _coluna_por_alias(df, ["codigo", "cod", "item", "produto", "codigo item"])
    col_desc = _coluna_por_alias(df, ["descricao", "descrição", "descricao item"])
    col_pr = _coluna_por_alias(df, ["pr", "ponto reposicao", "ponto de reposicao"])
    col_estoque = _coluna_por_alias(df, ["estoque", "estoque atual", "saldo"])
    col_brazip = _coluna_por_alias(df, ["nº brazip", "no brazip", "numero brazip", "brazip"])
    col_data_solic = _coluna_por_alias(
        df,
        ["data solic.", "data solic", "data da solicitacao", "data solicitacao", "solicitacao", "data solicitação"]
    )
    col_qtd_solic = _coluna_por_alias(
        df,
        ["qtd solic.", "qtd solic", "quantidade solicitada", "qtd solicitada", "qtde solicitada", "solicitado"]
    )
    col_status = _coluna_por_alias(df, ["status", "situacao", "situação"])
    col_dias = _coluna_por_alias(df, ["dias", "dias de estoque", "cobertura", "dias cobertura"])

    itens = []
    for _, row in df.iterrows():
        item = {
            "codigo": _valor_texto(row.get(col_codigo, "")) if col_codigo else "-",
            "descricao": _valor_texto(row.get(col_desc, "")) if col_desc else "-",
            "estoque": _valor_texto(row.get(col_estoque, "")) if col_estoque else "-",
            "pr": _valor_texto(row.get(col_pr, "")) if col_pr else "-",
            "brazip": _valor_texto(row.get(col_brazip, "")) if col_brazip else "-",
            "data_solicitacao": _valor_texto(row.get(col_data_solic, "")) if col_data_solic else "-",
            "qtd_solicitada": _valor_texto(row.get(col_qtd_solic, "")) if col_qtd_solic else "-",
            "status": _valor_texto(row.get(col_status, "")) if col_status else "-",
            "dias_estoque": _valor_texto(row.get(col_dias, "")) if col_dias else "-",
        }
        campos_controle = [
            "estoque",
            "pr",
            "dias_estoque",
            "brazip",
            "data_solicitacao",
            "qtd_solicitada",
        ]
        item["sem_dados"] = all(item[c] == "-" for c in campos_controle)
        itens.append(item)

    return itens, None


# =========================
# ANÁLISE MÊS A MÊS
# =========================
def _parse_num_analise(v):
    s = str(v).strip()
    if s == "" or s.lower() in {"nan", "none", "null", "-"}:
        return 0.0
    s = s.replace(" ", "")
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def _colunas_mes_analise(df):
    cols = [c for c in df.columns if re.fullmatch(r"\d{2}/\d{4}", str(c).strip())]
    cols = sorted(
        cols,
        key=lambda c: pd.to_datetime(f"01/{c}", format="%d/%m/%Y", errors="coerce")
    )
    return cols


def _carregar_analise_mes_a_mes():
    path = CSV_DIR / CSV_ANALISE_MES_A_MES
    if not path.exists():
        return None, "CSV de análise não encontrado. Gere 'explosao mes a mes .csv' na pasta csv."

    try:
        df = pd.read_csv(path, dtype=str, sep=None, engine="python")
    except Exception as e:
        return None, f"Erro ao ler CSV de análise mês a mês: {e}"

    df = df.dropna(how="all")
    if df.empty:
        return None, "CSV de análise mês a mês está vazio."

    codigo_col = _coluna_por_alias(df, ["codigo", "código", "cod"])
    desc_col = _coluna_por_alias(df, ["descricao", "descrição", "desc"])
    if not codigo_col:
        return None, "Coluna de código não encontrada no CSV de análise."
    if not desc_col:
        desc_col = codigo_col

    mes_cols = _colunas_mes_analise(df)
    if not mes_cols:
        return None, "Nenhuma coluna mensal no formato MM/AAAA foi encontrada no CSV."

    df = df[[codigo_col, desc_col] + mes_cols].copy()
    df.columns = ["codigo", "descricao"] + mes_cols
    df["codigo"] = df["codigo"].astype(str).str.strip()
    df["descricao"] = df["descricao"].astype(str).str.strip()
    df["codigo_norm"] = df["codigo"].apply(_norm_code)
    return df, None


@app.route('/api/analise-mes-a-mes/item')
def api_analise_mes_a_mes_item():
    codigo = str(request.args.get("codigo", "")).strip()
    if not codigo:
        return jsonify({"sucesso": False, "mensagem": "Informe um código para pesquisar."}), 400

    df, erro = _carregar_analise_mes_a_mes()
    if erro:
        return jsonify({"sucesso": False, "mensagem": erro}), 400

    codigo_norm = _norm_code(codigo)
    hit = df[df["codigo_norm"] == codigo_norm]
    if hit.empty:
        # fallback por "contains" quando o usuário cola formato com pontos/traços
        hit = df[df["codigo"].astype(str).str.contains(codigo, na=False)]

    if hit.empty:
        return jsonify({"sucesso": False, "mensagem": "Código não encontrado no CSV de análise."}), 404

    row = hit.iloc[0]
    mes_cols = [c for c in df.columns if re.fullmatch(r"\d{2}/\d{4}", str(c).strip())]
    serie = [{"mes": col, "valor": _parse_num_analise(row[col])} for col in mes_cols]

    return jsonify(
        {
            "sucesso": True,
            "item": {
                "codigo": str(row["codigo"]).strip(),
                "descricao": str(row["descricao"]).strip(),
                "serie": serie,
            },
            "meta": {
                "meses_disponiveis": mes_cols,
                "origem_csv": str(CSV_DIR / CSV_ANALISE_MES_A_MES),
            },
        }
    )


def _norm_code(v):
    s = str(v).strip()
    digits = "".join(ch for ch in s if ch.isdigit())
    return digits if digits else s



if __name__ == '__main__':
    app.run(debug=True)
