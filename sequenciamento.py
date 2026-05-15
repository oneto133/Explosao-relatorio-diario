from flask import Flask, request, jsonify, render_template
import csv
import os
from pathlib import Path
import pandas as pd
from datetime import date
import unicodedata

app = Flask(__name__)

# =========================
# PATHS
# =========================
BASE = Path.home() / "Desktop" / "Scraping_Explosão"
CSV_DIR = BASE / "csv"

CSV_OPS = "Report.csv"
CSV_ATRIB = "atribuicoes.csv"
CSV_OPERADORES = "operadores.csv"
CSV_OPS_MANUAL = "ops_manual.csv"
CSV_LT_TABELA_GERAL = "lt_tabela_geral.csv"
CSV_ESTOQUE_COMPRAS = "estoque_itens_comprados.csv"
CSV_SEPARACAO = "separacao.csv"

ESTOQUE_COMPRAS_XLSX = Path(r"Y:\Produção\Estoque\Estoque itens comprados.xlsx")
ESTOQUE_COMPRAS_SHEET = "Estoque Geral"


# =========================
# INDEX
# =========================
@app.route('/')
def index():
    return render_template('index.html')


# =========================
# SEQUENCIAMENTO (PÁGINA)
# =========================
@app.route('/sequenciamento')
def sequenciamento_page():
    return render_template('sequenciamento.html')

@app.route('/sequenciamento/imprimir')
def sequenciamento_impressao():
    return render_template('sequenciamento_impressao.html')

@app.route('/sequenciamento/compras')
def sequenciamento_compras():
    itens, erro = _carregar_compras_pendentes()
    return render_template(
        'compras.html',
        erro=erro,
        itens=itens or [],
        origem_excel=str(ESTOQUE_COMPRAS_XLSX),
        csv_cache=str(CSV_DIR / CSV_ESTOQUE_COMPRAS),
    )


# =========================
# EXPLOSÃO FUSOS (PÁGINA)
# =========================
@app.route('/explosao_fusos')
def explosao_fusos_page():
    return render_template('explosao_fusos.html')


# =========================
# RELATÓRIO PROGRAMAÇÃO
# =========================
@app.route('/relatorio_programacao')
def relatorio_programacao():
    vendas_path = CSV_DIR / "VENDAS__0402.csv"
    op_path = CSV_DIR / "OP__Report.csv"

    if not vendas_path.exists() or not op_path.exists():
        return render_template(
            "relatorio_programacao.html",
            erro="CSV não encontrado. Verifique se VENDAS__0402.csv e OP__Report.csv existem na pasta csv.",
            programado=[],
            nao_programado=[]
        )

    # ===== VENDAS =====
    vendas = pd.read_csv(vendas_path, header=2)
    vendas = vendas[vendas["Cód"].notna()].copy()
    vendas["Cód"] = vendas["Cód"].astype(str).str.strip()
    vendas["Descrição"] = vendas["Descrição"].astype(str).str.strip()
    vendas["Seção"] = vendas["Seção"].astype(str).str.strip()

    demanda_cols = [
        "Atraso",
        "Hoje",
        "1º Semana 08/02",
        "2º Semana 15/02",
        "3º Semana 22/02",
        "4º Semana 01/03",
        "MARÇO",
        "ABRIL",
    ]
    demanda_cols = [c for c in demanda_cols if c in vendas.columns]

    for c in demanda_cols:
        vendas[c] = pd.to_numeric(vendas[c], errors="coerce").fillna(0)

    # Vendas geralmente vêm negativas: soma apenas valores negativos
    vendas["vendido"] = -vendas[demanda_cols].clip(upper=0).sum(axis=1)
    vendas = vendas[vendas["vendido"] > 0]

    # ===== MAPA DE DATAS (BASE 04/02/2026) =====
    base_date = pd.Timestamp(year=2026, month=2, day=4)
    date_map = {
        "Atraso": base_date,
        "Hoje": base_date,
        "1º Semana 08/02": pd.Timestamp(year=2026, month=2, day=8),
        "2º Semana 15/02": pd.Timestamp(year=2026, month=2, day=15),
        "3º Semana 22/02": pd.Timestamp(year=2026, month=2, day=22),
        "4º Semana 01/03": pd.Timestamp(year=2026, month=3, day=1),
        "MARÇO": pd.Timestamp(year=2026, month=3, day=1),
        "ABRIL": pd.Timestamp(year=2026, month=4, day=1),
    }

    # ===== OPs =====
    ops = pd.read_csv(op_path, header=None)
    ops.columns = [
        "op",
        "data_emissao",
        "data_prevista",
        "status",
        "cod_prod",
        "descricao",
        "setor",
        "qtd_total",
        "qtd_produzida",
    ]
    ops["cod_prod"] = ops["cod_prod"].astype(str).str.strip()
    ops["qtd_total"] = pd.to_numeric(ops["qtd_total"], errors="coerce").fillna(0)
    ops["qtd_produzida"] = pd.to_numeric(ops["qtd_produzida"], errors="coerce").fillna(0)
    ops["pendente"] = (ops["qtd_total"] - ops["qtd_produzida"]).clip(lower=0)

    # ===== OPs MANUAIS =====
    manual_path = CSV_DIR / CSV_OPS_MANUAL
    manual_ops = pd.DataFrame()
    if manual_path.exists():
        manual_ops = pd.read_csv(manual_path)
        # alinhar colunas
        manual_ops = manual_ops.rename(columns={
            "qtd_total": "qtd_total",
            "qtd_produzida": "qtd_produzida"
        })

    # Remover manuais se a OP já existe no arquivo principal
    if not manual_ops.empty:
        ops_existentes = set(ops["op"].astype(str).str.strip())
        manual_ops = manual_ops[~manual_ops["op"].astype(str).str.strip().isin(ops_existentes)]

    # Remover manuais que já foram finalizadas no arquivo principal
    if not manual_ops.empty:
        ops_finalizadas = set(
            ops[ops["status"] == "Finalizada"]["op"].astype(str).str.strip()
        )
        manual_ops = manual_ops[~manual_ops["op"].astype(str).str.strip().isin(ops_finalizadas)]

    if not manual_ops.empty:
        manual_ops["qtd_total"] = pd.to_numeric(manual_ops["qtd_total"], errors="coerce").fillna(0)
        manual_ops["qtd_produzida"] = pd.to_numeric(manual_ops["qtd_produzida"], errors="coerce").fillna(0)
        manual_ops["status"] = manual_ops["status"].fillna("Aberta")
        manual_ops = manual_ops[[
            "op", "data_emissao", "data_prevista", "status",
            "cod_prod", "descricao", "setor", "qtd_total", "qtd_produzida"
        ]]
        ops = pd.concat([ops, manual_ops], ignore_index=True)

    ops_abertas = ops[ops["status"].isin(["Aberta", "Iniciada"])]
    programado = ops_abertas.groupby("cod_prod", as_index=False)["pendente"].sum()
    programado.rename(columns={"pendente": "programado"}, inplace=True)

    # Sequenciamento (atribuicoes)
    atribuicoes = carregar_atribuicoes()

    # ===== JUNÇÃO =====
    rel = vendas[["Cód", "Descrição", "Seção", "vendido"]].merge(
        programado, left_on="Cód", right_on="cod_prod", how="left"
    )
    rel["programado"] = rel["programado"].fillna(0)
    rel["saldo"] = rel["programado"] - rel["vendido"]

    rel = rel.sort_values(by="vendido", ascending=False)

    programado_ok = rel[(rel["programado"] >= rel["vendido"]) & (rel["programado"] > 0)]
    nao_programado = rel[(rel["programado"] < rel["vendido"]) | (rel["programado"] == 0)]

    def excel_date_to_str(value):
        if value is None or value == "":
            return ""
        try:
            num = float(str(value).replace(",", "."))
            dt = pd.to_datetime(num, unit="D", origin="1899-12-30", errors="coerce")
            if pd.isna(dt):
                return str(value).strip()
            return dt.strftime("%d/%m/%Y")
        except Exception:
            s = str(value).strip()
            try:
                dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
                if pd.isna(dt):
                    return s
                return dt.strftime("%d/%m/%Y")
            except Exception:
                return s

    def ops_por_produto(cod):
        subset = ops_abertas[ops_abertas["cod_prod"] == cod].copy()
        if subset.empty:
            return []
        itens = []
        for _, r in subset.iterrows():
            op_num = str(r["op"]).strip()
            itens.append({
                "op": op_num,
                "data_emissao": excel_date_to_str(r["data_emissao"]),
                "pendente": int(round(r["pendente"])),
                "sequenciado": op_num in atribuicoes
            })
        return itens

    def previsao_e_grade(venda_row):
        agenda = []
        atraso = False
        for col in demanda_cols:
            if col not in venda_row:
                continue
            qty = float(venda_row[col]) if pd.notna(venda_row[col]) else 0
            if qty >= 0:
                continue
            qty = int(round(abs(qty)))
            if qty <= 0:
                continue
            dt = date_map.get(col, base_date)
            if col == "Atraso":
                atraso = True
            agenda.append({
                "col": col,
                "data": dt.strftime("%d/%m/%Y"),
                "quantidade": qty
            })
        if not agenda:
            return ("", False, [])
        # Data mais próxima = primeira coluna com quantidade (ordem das colunas)
        previsao = agenda[0]["data"]
        return (previsao, atraso, agenda)

    def to_list(df):
        itens = []
        for _, r in df.iterrows():
            prev, atr, agenda = previsao_e_grade(r)
            itens.append({
                "codigo": r["Cód"],
                "descricao": r["Descrição"],
                "secao": r["Seção"],
                "vendido": int(round(r["vendido"])),
                "programado": int(round(r["programado"])),
                "saldo": int(round(r["saldo"])),
                "ops": ops_por_produto(r["Cód"]),
                "previsao": prev,
                "atraso": atr,
                "agenda": agenda,
            })
        return itens

    return render_template(
        "relatorio_programacao.html",
        erro=None,
        programado=to_list(programado_ok),
        nao_programado=to_list(nao_programado),
        secoes=sorted(vendas["Seção"].dropna().astype(str).str.strip().unique().tolist()),
    )


# =========================
# RELATÓRIO VENDAS (0402)
# =========================
@app.route('/relatorio_vendas')
def relatorio_vendas():
    vendas_path = CSV_DIR / "VENDAS__0402.csv"
    if not vendas_path.exists():
        return render_template(
            "relatorio_vendas.html",
            erro="CSV não encontrado. Verifique se VENDAS__0402.csv existe na pasta csv.",
            colunas=[],
            linhas=[]
        )

    df = pd.read_csv(vendas_path, header=2)
    df = df[df["Cód"].notna()].copy()
    df["Cód"] = df["Cód"].astype(str).str.strip()
    df["Descrição"] = df["Descrição"].astype(str).str.strip()
    df["Seção"] = df["Seção"].astype(str).str.strip()

    colunas = [
        "Cód",
        "Descrição",
        "Seção",
        "Atraso",
        "Hoje",
        "1º Semana 08/02",
        "2º Semana 15/02",
        "3º Semana 22/02",
        "4º Semana 01/03",
        "MARÇO",
        "ABRIL",
        "Estoque Disponível",
        "Situação",
    ]
    colunas = [c for c in colunas if c in df.columns]

    num_cols = {
        "Atraso",
        "Hoje",
        "1º Semana 08/02",
        "2º Semana 15/02",
        "3º Semana 22/02",
        "4º Semana 01/03",
        "MARÇO",
        "ABRIL",
    }

    linhas = []
    for _, r in df.iterrows():
        linha = {}
        for c in colunas:
            v = r.get(c, "")
            if pd.isna(v):
                v = ""
            is_neg = False
            if c in num_cols and v != "":
                try:
                    num = float(v)
                    if num.is_integer():
                        num = int(num)
                    is_neg = num < 0
                    v = num
                except Exception:
                    pass
            linha[c] = {"valor": v, "neg": is_neg} if c in num_cols else {"valor": v, "neg": False}
        linhas.append(linha)

    return render_template(
        "relatorio_vendas.html",
        erro=None,
        colunas=colunas,
        linhas=linhas
    )


# =========================
# PESQUISA
# =========================
@app.route('/pesquisar')
def pesquisar():
    termo = request.args.get('q', '').strip()

    if not termo.isdigit():
        return jsonify({"tipo": "erro", "mensagem": "Digite apenas números"})

    if len(termo) == 5:
        op = buscar_op(termo)
        return jsonify({"tipo": "op", "dados": op}) if op else jsonify({"tipo": "erro", "mensagem": "OP não encontrada"})

    if len(termo) in (6, 7):
        return jsonify({"tipo": "produto", "dados": buscar_produto(termo)})

    return jsonify({"tipo": "erro", "mensagem": "Digite uma OP (5) ou Produto (6/7)"})


# =========================
# BUSCAR OP
# =========================
def buscar_op(numero_op):
    atribuicoes = carregar_atribuicoes()

    with open(CSV_DIR / CSV_OPS, newline='', encoding='utf-8') as f:
        for row in csv.reader(f):
            if len(row) < 9 or row[0].strip() != numero_op:
                continue

            qtd_total = safe_int(row[7])
            qtd_prod = safe_int(row[8])

            status = (
                "Inválida" if qtd_total <= 0 else
                "Finalizada" if qtd_prod >= qtd_total else
                "Em aberto"
            )

            return {
                "op": numero_op,
                "produto": row[4].strip(),
                "descricao": row[5].strip(),
                "quantidade_total": qtd_total,
                "quantidade_produzida": qtd_prod,
                "status": status,
                "operador": atribuicoes.get(numero_op)
            }


# =========================
# BUSCAR PRODUTO
# =========================
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
# EXPLOSÃO FUSOS (LÓGICA)
# =========================
def _normalize_header(s):
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = s.replace("_", " ").replace("-", " ")
    s = " ".join(s.split())
    return s


def _find_col(cols_norm, keys):
    for key in keys:
        for col, norm in cols_norm.items():
            if key in norm:
                return col
    return None


def _parse_number(v):
    s = str(v).strip()
    if s == "":
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


def _norm_code(v):
    s = str(v).strip()
    digits = "".join(ch for ch in s if ch.isdigit())
    return digits if digits else s


def _carregar_tabela_geral():
    path = CSV_DIR / CSV_LT_TABELA_GERAL
    if not path.exists():
        return None, f"CSV não encontrado. Crie {CSV_LT_TABELA_GERAL} na pasta csv."
    try:
        df = pd.read_csv(path, dtype=str, sep=None, engine="python")
    except Exception as e:
        return None, f"Erro ao ler o CSV: {e}"
    df = df.dropna(how="all")
    if df.empty:
        return None, "CSV vazio."

    # Se o CSV veio sem cabeçalho (Column1..Column5), mapear por posição.
    if len(df.columns) >= 5 and all(str(c).lower().startswith("column") for c in df.columns[:5]):
        df = df.iloc[:, :5].copy()
        df.columns = ["produto_pai", "descricao_pai", "cod_mp", "descricao_mp", "qtd_base"]
        return df, None

    cols_norm = {c: _normalize_header(c) for c in df.columns}

    parent_col = _find_col(
        cols_norm,
        ["produto pai", "codigo pai", "cod pai", "produto_pai", "pai"]
    )
    desc_pai_col = _find_col(
        cols_norm,
        ["descricao pai", "descr pai", "descricao produto pai"]
    )
    mp_col = _find_col(
        cols_norm,
        ["codigo mp", "cod mp", "codigo materia prima", "materia prima", "materiaprima", "mp"]
    )
    desc_col = _find_col(
        cols_norm,
        ["descricao mp", "descricao materia prima", "descricao", "descr"]
    )
    qtd_col = _find_col(
        cols_norm,
        ["quantidade", "qtd", "qtde", "consumo", "qtd por", "quantidade por", "qtd_por"]
    )

    if not parent_col or not mp_col or not desc_col or not qtd_col:
        msg = (
            "Colunas obrigatórias não encontradas no CSV. "
            "Esperado algo como: produto_pai, cod_mp, descricao_mp, qtd_base "
            "(os nomes podem variar, mas precisam indicar essas informações)."
        )
        return None, msg

    keep = [parent_col]
    if desc_pai_col:
        keep.append(desc_pai_col)
    keep += [mp_col, desc_col, qtd_col]
    df = df[keep].copy()
    if desc_pai_col:
        df.columns = ["produto_pai", "descricao_pai", "cod_mp", "descricao_mp", "qtd_base"]
    else:
        df.columns = ["produto_pai", "cod_mp", "descricao_mp", "qtd_base"]
    return df, None


@app.route('/explosao_fusos/calcular', methods=['POST'])
def explosao_fusos_calcular():
    data = request.json or {}
    produto = str(data.get('produto', '')).strip()
    quantidade = str(data.get('quantidade', '')).strip()

    if not produto or not quantidade:
        return jsonify({
            "sucesso": False,
            "mensagem": "Informe produto pai e quantidade."
        }), 400

    df, erro = _carregar_tabela_geral()
    if erro:
        return jsonify({"sucesso": False, "mensagem": erro}), 400

    df["produto_pai"] = df["produto_pai"].astype(str).str.strip()
    df["cod_mp"] = df["cod_mp"].astype(str).str.strip()
    df["descricao_mp"] = df["descricao_mp"].astype(str).str.strip()
    df["qtd_base"] = df["qtd_base"].apply(_parse_number)

    produto_norm = _norm_code(produto)
    df["produto_pai_norm"] = df["produto_pai"].apply(_norm_code)
    subset = df[df["produto_pai_norm"] == produto_norm]

    # Se não encontrou, tentar detectar campos invertidos (produto/quantidade)
    if subset.empty:
        possivel_produto = _norm_code(quantidade)
        subset_alt = df[df["produto_pai_norm"] == possivel_produto]
        if not subset_alt.empty:
            produto, quantidade = quantidade, produto
            produto_norm = possivel_produto
            subset = subset_alt

    if subset.empty:
        return jsonify({
            "sucesso": False,
            "mensagem": "Produto pai não encontrado no CSV."
        }), 404

    qtd_pai = _parse_number(quantidade)
    if qtd_pai <= 0:
        return jsonify({
            "sucesso": False,
            "mensagem": "Quantidade inválida."
        }), 400

    # Somente materiais cujo texto começa com "FUSO"
    subset = subset[subset["descricao_mp"].str.startswith("FUSO", na=False)]
    if subset.empty:
        return jsonify({
            "sucesso": True,
            "mensagem": "Nenhum fuso encontrado para este produto.",
            "itens": []
        })

    subset = subset.copy()
    subset["qtd_total"] = subset["qtd_base"] * qtd_pai

    agrupado = (
        subset.groupby(["cod_mp", "descricao_mp"], as_index=False)["qtd_total"]
        .sum()
        .sort_values(by=["descricao_mp", "cod_mp"])
    )

    itens = []
    for _, r in agrupado.iterrows():
        qtd = float(r["qtd_total"])
        if qtd.is_integer():
            qtd = int(qtd)
        itens.append({
            "codigo": str(r["cod_mp"]).strip(),
            "descricao": str(r["descricao_mp"]).strip(),
            "quantidade": qtd
        })

    return jsonify({"sucesso": True, "itens": itens})


@app.route('/explosao_fusos/impressao', methods=['POST'])
def explosao_fusos_impressao():
    data = request.json or {}
    itens = data.get("itens", [])
    data_impressao = str(data.get("data_impressao", "")).strip()
    total_geral = _parse_number(data.get("total_geral", 0))

    if not isinstance(itens, list):
        return jsonify({
            "sucesso": False,
            "mensagem": "Formato invalido para itens."
        }), 400

    linhas = []
    for item in itens:
        if not isinstance(item, dict):
            continue

        codigo = str(item.get("codigo", "")).strip()
        descricao = str(item.get("descricao", "")).strip()
        tamanho = str(item.get("tamanho", "")).strip()
        procedimento = str(item.get("procedimento", "")).strip()
        quantidade = _parse_number(item.get("quantidade", 0))

        if quantidade <= 0 and not codigo and not descricao:
            continue

        qtd_saida = int(quantidade) if float(quantidade).is_integer() else quantidade
        linhas.append({
            "codigo": codigo,
            "quantidade": qtd_saida,
            "descricao": descricao,
            "tamanho": tamanho,
            "procedimento": procedimento,
        })

    if not linhas:
        return jsonify({
            "sucesso": False,
            "mensagem": "Nenhum item valido para salvar."
        }), 400

    total_saida = int(total_geral) if float(total_geral).is_integer() else total_geral
    path = CSV_DIR / CSV_SEPARACAO

    try:
        CSV_DIR.mkdir(parents=True, exist_ok=True)
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow([
                "codigo",
                "quantidade",
                "descricao",
                "tamanho",
                "procedimento",
                "data_impressao",
                "total_geral",
            ])
            for linha in linhas:
                w.writerow([
                    linha["codigo"],
                    linha["quantidade"],
                    linha["descricao"],
                    linha["tamanho"],
                    linha["procedimento"],
                    data_impressao,
                    total_saida,
                ])
    except Exception as e:
        return jsonify({
            "sucesso": False,
            "mensagem": f"Erro ao salvar CSV: {e}"
        }), 500

    return jsonify({
        "sucesso": True,
        "arquivo": str(path),
        "linhas": len(linhas)
    })


# =========================
# SEQUENCIAR
# =========================
@app.route('/sequenciar', methods=['POST'])
def sequenciar():
    data = request.json

    with open(CSV_DIR / CSV_ATRIB, 'a', newline='', encoding='utf-8') as f:
        csv.writer(f).writerow([
            data['op'],
            data['linha'],
            data['operador']
        ])

    return jsonify({'sucesso': True})


# =========================
# OP MANUAL
# =========================
@app.route('/ops_manual', methods=['POST'])
def ops_manual():
    data = request.json
    op = str(data.get('op', '')).strip()
    cod_prod = str(data.get('cod_prod', '')).strip()
    descricao = str(data.get('descricao', '')).strip()
    quantidade = str(data.get('quantidade', '')).strip()

    if not op or not cod_prod or not quantidade:
        return jsonify({'sucesso': False, 'mensagem': 'Dados incompletos'}), 400

    try:
        qtd = int(str(quantidade).replace(',', '').strip())
    except Exception:
        return jsonify({'sucesso': False, 'mensagem': 'Quantidade inválida'}), 400

    linha = str(data.get('linha', '')).strip()
    data_emissao = date.today().strftime("%d/%m/%Y")

    path = CSV_DIR / CSV_OPS_MANUAL
    novo = not path.exists()
    with open(path, 'a', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        if novo:
            w.writerow([
                "op", "data_emissao", "data_prevista", "status",
                "cod_prod", "descricao", "setor", "qtd_total",
                "qtd_produzida", "linha", "origem"
            ])
        w.writerow([
            op, data_emissao, "", "Aberta",
            cod_prod, descricao, "MOTORES",
            qtd, 0, linha, "manual"
        ])

    return jsonify({'sucesso': True})

@app.route('/ops')
def listar_ops():
    operador = request.args.get('operador')
    produto = request.args.get('produto')
    status = request.args.get('status')  # aberto | finalizada
    producao = request.args.get('producao')  # sim | nao

    atribuicoes = carregar_atribuicoes()
    lista = []

    with open(CSV_DIR / CSV_OPS, newline='', encoding='utf-8') as f:
        for row in csv.reader(f):
            if len(row) < 9:
                continue

            op = row[0].strip()
            cod_prod = row[4].strip()
            descricao = row[5].strip()
            qtd_total = safe_int(row[7])
            qtd_prod = safe_int(row[8])

            if qtd_total <= 0:
                status_op = "Inválida"
            elif qtd_prod >= qtd_total:
                status_op = "Finalizada"
            else:
                status_op = "Em aberto"

            operador_op = atribuicoes.get(op)

            # ===== FILTROS =====
            if operador and operador_op != operador:
                continue

            if produto and produto != cod_prod:
                continue

            if status == "aberto" and status_op != "Em aberto":
                continue

            if status == "finalizada" and status_op != "Finalizada":
                continue

            if producao == "sim" and qtd_prod == 0:
                continue

            if producao == "nao" and qtd_prod > 0:
                continue

            lista.append({
                "op": op,
                "produto": cod_prod,
                "descricao": descricao,
                "quantidade_total": qtd_total,
                "quantidade_produzida": qtd_prod,
                "status": status_op,
                "operador": operador_op
            })

    lista.sort(key=lambda x: x["quantidade_produzida"])
    return jsonify(lista)

@app.route('/operadores_lista')
def operadores_lista():
    operadores = set(carregar_atribuicoes().values())
    return jsonify(sorted(o for o in operadores if o))


if __name__ == '__main__':
    app.run(debug=True)
