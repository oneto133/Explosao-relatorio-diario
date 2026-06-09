import argparse
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

import pandas as pd

"""O intuito dessa aplicação é extrair o lead time de itens comprados
atualmente trabalhamos com um sistema onde os setores inserem os itens 
das quais estão necessitados, então um administrador que controla gastos
passa ao comprador após aprovação, o comprador vai mudando o status entre
em contação, aguardando liberação da diretoria, pedido de compra realizado
e concluído, então essas solicitações e logs ficam salvos, nessa aplicação
há como extrair relatórios de soicitações de diversas maneiras, então a ideia foi
pegar a data da solicitação pelo setor e a data de conclusão pelo comprador
o que significa que o item está disponível na fábrica, então eu extrai as
solicitações e códigos e fui cruzando os dados, após é subtraido a data de 
consclusão da data inicial e calculado a média com base no código, assim
se obtendo um lead time de itens, tendo dados detalhados disponíveis para
pegar prazo mínimo e máximo."""

RE_SOLIC = re.compile(r"Solicita(?:ç|c)ão\s*:\s*(\d+)", re.IGNORECASE)
RE_DT = re.compile(r"^\s*(\d{2}/\d{2}/\d{2}\s+\d{2}:\d{2}:\d{2})\s*$")
UNID_RE = r"(?:UN|UNID(?:ADES?)?|UNIDADES?|M|MT|MTS|METRO(?:S)?|L|LT|LTS|LITRO(?:S)?|KG|KGS|KILO(?:S)?|QUILO(?:S)?)"
# Ex.: 350 UN - 401225 - DESCRICAO
RE_QTD_ANTES = re.compile(
    rf"(?<!\d)(\d{{1,3}}(?:[\.,]\d{{3}})*|\d+)\s*({UNID_RE})?\s*-\s*(\d{{5,6}})(?!\d)\s*-\s*[^\n]+",
    re.IGNORECASE,
)
# Ex.: 105069 - FITA ... - 40 unidades
RE_QTD_DEPOIS = re.compile(
    rf"(?<!\d)(\d{{5,6}})(?!\d)\s*-\s*[^\n]+?\s*-\s*(\d{{1,3}}(?:[\.,]\d{{3}})*|\d+)\s*({UNID_RE})",
    re.IGNORECASE,
)
# Ex.: 401225 - CORRENTE ... (sem quantidade explícita)
RE_COD_DESC = re.compile(r"(?<!\d)(\d{5,6})(?!\d)\s*-\s*[A-ZÀ-ÿ]", re.IGNORECASE)
XLSX_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def parse_dt(texto):
    return datetime.strptime(texto, "%d/%m/%y %H:%M:%S")


def fmt_dt(dt):
    return dt.strftime("%Y-%m-%d %H:%M:%S") if dt else "-"


def _coluna_excel_para_indice(ref):
    letras = re.match(r"([A-Z]+)", ref or "")
    if not letras:
        return None
    valor = 0
    for char in letras.group(1):
        valor = valor * 26 + (ord(char) - ord("A") + 1)
    return valor - 1


def _local_name(tag):
    return tag.split("}", 1)[-1]


def _ler_xlsx_primeira_aba(path):
    with zipfile.ZipFile(path) as zf:
        shared_strings = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall("main:si", XLSX_NS):
                partes = []
                for node in si.iter():
                    if _local_name(node.tag) == "t" and node.text:
                        partes.append(node.text)
                shared_strings.append("".join(partes))

        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        sheets = workbook.find("main:sheets", XLSX_NS)
        if sheets is None:
            return pd.DataFrame()

        primeiro_sheet = next(iter(sheets), None)
        if primeiro_sheet is None:
            return pd.DataFrame()

        rel_id = primeiro_sheet.attrib.get(f"{{{XLSX_NS['rel']}}}id")
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        target = None
        for rel in rels.findall("pkgrel:Relationship", XLSX_NS):
            if rel.attrib.get("Id") == rel_id:
                target = rel.attrib.get("Target")
                break

        if not target:
            target = "worksheets/sheet1.xml"
        sheet_path = f"xl/{target.lstrip('/')}"
        if sheet_path not in zf.namelist():
            sheet_path = "xl/worksheets/sheet1.xml"

        sheet = ET.fromstring(zf.read(sheet_path))
        linhas = []
        for row in sheet.findall(".//main:sheetData/main:row", XLSX_NS):
            valores = []
            for cell in row.findall("main:c", XLSX_NS):
                ref = cell.attrib.get("r", "")
                idx = _coluna_excel_para_indice(ref)
                if idx is None:
                    continue

                while len(valores) <= idx:
                    valores.append("")

                tipo = cell.attrib.get("t", "")
                valor = ""
                if tipo == "s":
                    v = cell.find("main:v", XLSX_NS)
                    if v is not None and v.text is not None:
                        try:
                            valor = shared_strings[int(v.text)]
                        except (ValueError, IndexError):
                            valor = v.text
                elif tipo == "inlineStr":
                    is_node = cell.find("main:is", XLSX_NS)
                    if is_node is not None:
                        textos = []
                        for node in is_node.iter():
                            if _local_name(node.tag) == "t" and node.text:
                                textos.append(node.text)
                        valor = "".join(textos)
                else:
                    v = cell.find("main:v", XLSX_NS)
                    if v is not None and v.text is not None:
                        valor = v.text

                valores[idx] = valor

            linhas.append(valores)

    max_cols = max((len(linha) for linha in linhas), default=0)
    linhas = [linha + [""] * (max_cols - len(linha)) for linha in linhas]
    return pd.DataFrame(linhas)


def normalizar_qtd(valor):
    if not valor:
        return "-"
    q = str(valor).strip().replace(" ", "")
    if "," in q and "." in q:
        q = q.replace(".", "").replace(",", ".")
    elif "." in q and all(parte.isdigit() for parte in q.split(".")):
        q = q.replace(".", "")
    elif "," in q and all(parte.isdigit() for parte in q.split(",")):
        q = q.replace(",", "")
    return q


def normalizar_unidade(valor):
    if not valor:
        return "-"
    u = str(valor).strip().upper()
    mapa = {
        "UNIDADE": "UN",
        "UNIDADES": "UN",
        "UNID": "UN",
        "UNIDS": "UN",
        "UN": "UN",
        "M": "M",
        "MT": "M",
        "MTS": "M",
        "METRO": "M",
        "METROS": "M",
        "L": "L",
        "LT": "L",
        "LTS": "L",
        "LITRO": "L",
        "LITROS": "L",
        "KG": "KG",
        "KGS": "KG",
        "KILO": "KG",
        "KILOS": "KG",
        "QUILO": "KG",
        "QUILOS": "KG",
    }
    return mapa.get(u, u)


def extrair_dados(arquivo_xlsx):
    if arquivo_xlsx.suffix.lower() == ".xlsx":
        df = _ler_xlsx_primeira_aba(arquivo_xlsx).fillna("")
    else:
        df = pd.read_excel(arquivo_xlsx, header=None, dtype=str).fillna("")

    col_data = 0
    col_interacao = 3
    solicitacao_atual = None
    datas_por_solic = {}
    itens_extraidos = []

    for _, linha in df.iterrows():
        c0 = str(linha.get(col_data, "") or "").strip()
        c3 = str(linha.get(col_interacao, "") or "").strip()

        msol = RE_SOLIC.search(c0)
        if msol:
            solicitacao_atual = msol.group(1)
            datas_por_solic.setdefault(solicitacao_atual, [])
            continue

        if not solicitacao_atual:
            continue

        mdt = RE_DT.match(c0)
        if mdt:
            datas_por_solic.setdefault(solicitacao_atual, []).append(parse_dt(mdt.group(1)))

        if not c3:
            continue

        texto = c3.replace("\r", "\n").replace(" / ", "\n")
        partes = [p.strip() for p in texto.split("\n") if p.strip()]

        for parte in partes:
            m1 = RE_QTD_ANTES.search(parte)
            if m1:
                qtd, unidade, cod = m1.group(1), m1.group(2), m1.group(3)
                itens_extraidos.append(
                    {
                        "solicitacao": solicitacao_atual,
                        "item": cod,
                        "quantidade": normalizar_qtd(qtd),
                        "unidade": normalizar_unidade(unidade),
                    }
                )
                continue

            m2 = RE_QTD_DEPOIS.search(parte)
            if m2:
                cod, qtd, unidade = m2.group(1), m2.group(2), m2.group(3)
                itens_extraidos.append(
                    {
                        "solicitacao": solicitacao_atual,
                        "item": cod,
                        "quantidade": normalizar_qtd(qtd),
                        "unidade": normalizar_unidade(unidade),
                    }
                )
                continue

            m3 = RE_COD_DESC.search(parte)
            if m3:
                itens_extraidos.append(
                    {"solicitacao": solicitacao_atual, "item": m3.group(1), "quantidade": "-", "unidade": "-"}
                )

    periodo = {}
    for solic, datas in datas_por_solic.items():
        periodo[solic] = (min(datas), max(datas)) if datas else (None, None)

    saida = []
    for rec in itens_extraidos:
        ini, fim = periodo.get(rec["solicitacao"], (None, None))
        saida.append(
            {
                "Solicitação": rec["solicitacao"],
                "data inicio": fmt_dt(ini),
                "data fim": fmt_dt(fim),
                "item": rec["item"],
                "quantidade": rec["quantidade"],
                "unidade": rec.get("unidade", "-"),
            }
        )

    return pd.DataFrame(saida, columns=["Solicitação", "data inicio", "data fim", "item", "quantidade", "unidade"])


def main():
    parser = argparse.ArgumentParser(
        description="Extrai itens e quantidades de interações de solicitações detalhadas."
    )
    parser.add_argument(
        "--input",
        default=r"Z:\PCP\05 - Motores\Dados\solicitacao.solicitacoesdetalhadas.xlsx",
        help="Caminho do arquivo XLSX de origem.",
    )
    parser.add_argument(
        "--output",
        default=r"c:\Users\pc\Desktop\Scraping_Explosao\csv\solicitacoes_itens2.csv",
        help="Caminho do arquivo CSV de saída.",
    )
    args = parser.parse_args()

    origem = Path(args.input)
    destino = Path(args.output)

    if not origem.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {origem}")

    df_saida = extrair_dados(origem)
    destino.parent.mkdir(parents=True, exist_ok=True)
    df_saida.to_csv(destino, index=False, encoding="utf-8-sig")

    print(f"Arquivo gerado: {destino}")
    print(f"Linhas: {len(df_saida)}")
    print(f"Solicitações: {df_saida['Solicitação'].nunique() if len(df_saida) else 0}")


if __name__ == "__main__":
    main()
