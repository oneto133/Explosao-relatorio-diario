import json
import os
import re
import shutil
import sys
import threading
import traceback
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd


MODULE_DIR = Path(__file__).resolve().parent
BUNDLE_DIR = Path(getattr(sys, "_MEIPASS", MODULE_DIR))
APP_BASE_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else MODULE_DIR
APP_DATA_DIR = APP_BASE_DIR / "explosao_op_dados"

CACHE_DIR = APP_DATA_DIR / "cache"
CACHE_FILE = CACHE_DIR / "arquivos_enderecos.json"
LEGACY_CACHE_FILES = [
    MODULE_DIR.parent / "cache" / "arquivos_enderecos.json",
    MODULE_DIR / "cache" / "arquivos_enderecos.json",
]
TEMP_DIR = APP_DATA_DIR / "temp"
OUTPUT_DIR = APP_DATA_DIR / "saida"
OUTPUT_FILE = OUTPUT_DIR / "explosao_materia_prima.csv"
SETOR_SEM_INFO = "SEM_SETOR"
SETORES_FILE = APP_DATA_DIR / "setores.csv"
BUNDLED_SETORES_FILE = BUNDLE_DIR / "app" / "explosao_op" / "setores.csv"


def _normalizar_xlsx(caminho_xlsx: Path) -> None:
    caminho_tmp = caminho_xlsx.with_suffix(caminho_xlsx.suffix + ".tmp")
    houve_ajuste = False

    with zipfile.ZipFile(caminho_xlsx, "r") as origem_zip, zipfile.ZipFile(
        caminho_tmp, "w", compression=zipfile.ZIP_DEFLATED
    ) as destino_zip:
        for item in origem_zip.infolist():
            conteudo = origem_zip.read(item.filename)

            if item.filename == "xl/workbook.xml":
                texto = conteudo.decode("utf-8", errors="ignore")
                texto_normalizado = (
                    texto.replace(" WindowWidth=", " windowWidth=")
                    .replace(" WindowHeight=", " windowHeight=")
                )
                if texto_normalizado != texto:
                    houve_ajuste = True
                    conteudo = texto_normalizado.encode("utf-8")

            elif item.filename == "docProps/core.xml":
                texto = conteudo.decode("utf-8", errors="ignore")
                texto_normalizado = re.sub(
                    r"<cp:lastPrinted>.*?</cp:lastPrinted>",
                    "",
                    texto,
                    flags=re.IGNORECASE | re.DOTALL,
                )
                texto_normalizado = re.sub(
                    r"<dcterms:created[^>]*>.*?</dcterms:created\s*>",
                    "",
                    texto_normalizado,
                    flags=re.IGNORECASE | re.DOTALL,
                )
                texto_normalizado = re.sub(
                    r"<dcterms:modified[^>]*>.*?</dcterms:modified\s*>",
                    "",
                    texto_normalizado,
                    flags=re.IGNORECASE | re.DOTALL,
                )
                texto_normalizado = texto_normalizado.replace(" \u00edndexed=", " indexed=")
                if texto_normalizado != texto:
                    houve_ajuste = True
                    conteudo = texto_normalizado.encode("utf-8")

            elif item.filename == "xl/styles.xml":
                texto = conteudo.decode("utf-8", errors="ignore")
                texto_normalizado = texto.replace(" \u00edndexed=", " indexed=")
                if texto_normalizado != texto:
                    houve_ajuste = True
                    conteudo = texto_normalizado.encode("utf-8")

            destino_zip.writestr(item, conteudo)

    if houve_ajuste:
        caminho_tmp.replace(caminho_xlsx)
    else:
        caminho_tmp.unlink(missing_ok=True)


def _regravar_xlsx_com_excel(caminho_xlsx: Path) -> bool:
    try:
        import win32com.client as win32
    except Exception:
        return False

    excel = None
    workbook = None
    caminho_regravado = caminho_xlsx.with_name(f"{caminho_xlsx.stem}__excel{caminho_xlsx.suffix}")

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(str(caminho_xlsx.resolve()))
        workbook.SaveAs(str(caminho_regravado.resolve()), FileFormat=51)
        workbook.Close(SaveChanges=False)
        workbook = None

        excel.Quit()
        excel = None

        caminho_regravado.replace(caminho_xlsx)
        return True
    except Exception:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        caminho_regravado.unlink(missing_ok=True)
        return False


def _ler_tabular(caminho: Path) -> pd.DataFrame:
    sufixo = caminho.suffix.lower()

    if sufixo in {".csv", ".txt"}:
        try:
            return pd.read_csv(caminho, header=None, dtype=object, sep=None, engine="python")
        except Exception:
            return pd.read_csv(caminho, header=None, dtype=object, sep=";", engine="python")

    if sufixo in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        _normalizar_xlsx(caminho)
        try:
            return pd.read_excel(caminho, header=None, dtype=object)
        except Exception:
            if _regravar_xlsx_com_excel(caminho):
                return pd.read_excel(caminho, header=None, dtype=object)
            raise

    if sufixo == ".xlsb":
        try:
            return pd.read_excel(caminho, header=None, dtype=object, engine="pyxlsb")
        except Exception:
            if _regravar_xlsx_com_excel(caminho):
                return pd.read_excel(caminho, header=None, dtype=object)
            raise

    if sufixo == ".xls":
        if _regravar_xlsx_com_excel(caminho):
            return pd.read_excel(caminho, header=None, dtype=object)
        raise ValueError(f"Nao foi possivel ler '{caminho.name}'. Salve como XLSX e tente novamente.")

    raise ValueError(f"Formato de arquivo nao suportado: {caminho.suffix}")


def _to_numeric_series(serie: pd.Series) -> pd.Series:
    def _converter_texto_numero(valor: object) -> float | None:
        if pd.isna(valor):
            return None

        texto = str(valor).strip()
        if texto in {"", "nan", "None", "NaT"}:
            return None

        texto = texto.replace("\u00A0", "").replace(" ", "").replace("−", "-")
        if not texto:
            return None

        tem_virgula = "," in texto
        tem_ponto = "." in texto

        if tem_virgula and tem_ponto:
            # Quando os dois separadores existem, assume o ultimo como decimal.
            if texto.rfind(",") > texto.rfind("."):
                texto = texto.replace(".", "").replace(",", ".")
            else:
                texto = texto.replace(",", "")
        elif tem_virgula:
            # Com apenas virgula, privilegia decimal; remove milhar so se houver 2+ blocos.
            if re.fullmatch(r"[+-]?\d{1,3}(,\d{3}){2,}", texto):
                texto = texto.replace(",", "")
            else:
                texto = texto.replace(",", ".")
        elif tem_ponto:
            # Com apenas ponto, privilegia decimal; remove milhar so se houver 2+ blocos.
            if re.fullmatch(r"[+-]?\d{1,3}(\.\d{3}){2,}", texto):
                texto = texto.replace(".", "")

        try:
            return float(texto)
        except ValueError:
            return None

    return pd.to_numeric(serie.map(_converter_texto_numero), errors="coerce")


def _normalizar_codigo(serie: pd.Series) -> pd.Series:
    texto = serie.astype(str).str.strip()
    numeros = _to_numeric_series(serie)
    codigo = texto.copy()
    mascara = numeros.notna()
    codigo.loc[mascara] = numeros.loc[mascara].round().astype("Int64").astype(str)
    codigo = codigo.replace({"nan": "", "None": ""})
    codigo = codigo.str.replace(r"\.0+$", "", regex=True).str.strip()
    return codigo


def _to_datetime_series(serie: pd.Series) -> pd.Series:
    datas = pd.to_datetime(serie, errors="coerce", dayfirst=True)
    numeros = _to_numeric_series(serie)
    mascara_serial_excel = datas.isna() & numeros.notna()
    if mascara_serial_excel.any():
        datas.loc[mascara_serial_excel] = pd.to_datetime(
            numeros.loc[mascara_serial_excel],
            errors="coerce",
            unit="D",
            origin="1899-12-30",
        )
    return datas


def _preparar_lista_tecnica(df_bruto: pd.DataFrame) -> pd.DataFrame:
    if df_bruto.shape[1] < 5:
        raise ValueError("Lista tecnica invalida: esperado minimo de 5 colunas.")

    df = pd.DataFrame(
        {
            "cod_produto_pai": _normalizar_codigo(df_bruto.iloc[:, 0]),
            "cod_materia_prima": _normalizar_codigo(df_bruto.iloc[:, 2]),
            "desc_materia_prima": df_bruto.iloc[:, 3].astype(str).str.strip(),
            "qtd_mp": _to_numeric_series(df_bruto.iloc[:, 4]),
        }
    )
    df["desc_materia_prima"] = df["desc_materia_prima"].replace({"nan": "", "None": ""})

    mascara_valida = (
        df["cod_produto_pai"].str.fullmatch(r"\d+")
        & df["cod_materia_prima"].str.fullmatch(r"\d+")
        & df["qtd_mp"].notna()
        & (df["qtd_mp"] != 0)
    )
    df = df.loc[mascara_valida].copy()

    if df.empty:
        raise ValueError("Lista tecnica sem linhas validas apos limpeza.")

    return df


def _preparar_op(df_bruto: pd.DataFrame) -> pd.DataFrame:
    if df_bruto.shape[1] < 9:
        raise ValueError("Arquivo de OP invalido: esperado minimo de 9 colunas.")

    setor = df_bruto.iloc[:, 6].astype(str).str.strip()
    setor = setor.replace({"nan": "", "None": ""})
    setor = setor.mask(setor == "", SETOR_SEM_INFO)

    df = pd.DataFrame(
        {
            "data_referencia": _to_datetime_series(df_bruto.iloc[:, 1]),
            "cod_produto_pai": _normalizar_codigo(df_bruto.iloc[:, 4]),
            "setor": setor,
            "qtd_produzida": _to_numeric_series(df_bruto.iloc[:, 8]),
        }
    )

    mascara_valida = (
        df["cod_produto_pai"].str.fullmatch(r"\d+")
        & df["data_referencia"].notna()
        & df["qtd_produzida"].notna()
        & (df["qtd_produzida"] > 0)
    )
    df = df.loc[mascara_valida].copy()

    if df.empty:
        raise ValueError("OP sem linhas validas apos limpeza.")

    return df


def _normalizar_setor_chave(valor: object) -> str:
    texto = str(valor).strip().upper()
    if not texto:
        return ""
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    # Remove pontuacao e espacos para comparar variacoes como "A / B" e "A-B".
    texto = re.sub(r"[^A-Z0-9]+", "", texto)
    return texto


def _carregar_setores_referencia(caminho_setores: Path = SETORES_FILE) -> list[str]:
    if not caminho_setores.exists():
        return []

    setores: list[str] = []
    vistos: set[str] = set()

    texto_arquivo = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            texto_arquivo = caminho_setores.read_text(encoding=encoding)
            break
        except Exception:
            continue
    if texto_arquivo is None:
        return []

    for linha in texto_arquivo.splitlines():
        linha = linha.strip()
        if not linha:
            continue

        # Aceita arquivo de uma coluna ou com delimitadores simples.
        partes = [p.strip() for p in re.split(r"[;,\t|]", linha) if p.strip()]
        if not partes:
            continue

        for setor in partes:
            chave = _normalizar_setor_chave(setor)
            if chave in {"SETOR", "SETORES", SETOR_SEM_INFO}:
                continue
            if chave and chave not in vistos:
                vistos.add(chave)
                setores.append(setor)
    return setores


def _obter_pasta_downloads() -> Path:
    # Windows: usa a pasta de Downloads configurada no perfil do usuario.
    if os.name == "nt":
        try:
            import winreg

            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
            ) as chave:
                caminho, _ = winreg.QueryValueEx(chave, "{374DE290-123F-4565-9164-39C4925E467B}")
            caminho_expandido = os.path.expandvars(str(caminho))
            if caminho_expandido:
                return Path(caminho_expandido)
        except Exception:
            pass

    return Path.home() / "Downloads"


def _caminho_unico(destino: Path) -> Path:
    if not destino.exists():
        return destino

    base = destino.stem
    sufixo = destino.suffix
    pasta = destino.parent
    contador = 1
    while True:
        candidato = pasta / f"{base} ({contador}){sufixo}"
        if not candidato.exists():
            return candidato
        contador += 1


def _garantir_ambiente_execucao() -> None:
    for pasta in (APP_DATA_DIR, CACHE_DIR, TEMP_DIR, OUTPUT_DIR):
        pasta.mkdir(parents=True, exist_ok=True)

    if SETORES_FILE.exists():
        return

    candidatos = [
        BUNDLED_SETORES_FILE,
        MODULE_DIR / "setores.csv",
        APP_BASE_DIR / "setores.csv",
    ]
    for origem in candidatos:
        if origem.exists():
            shutil.copy2(origem, SETORES_FILE)
            return

    SETORES_FILE.write_text(
        "FECHADURAS\nMOTORES\nUSINAGEM / LAMINACAO\n",
        encoding="utf-8-sig",
    )


def _extrair_setores(df_op: pd.DataFrame) -> list[str]:
    return sorted(s for s in df_op["setor"].dropna().astype(str).str.strip().unique() if s)


def _filtrar_setores_op_por_referencia(df_op: pd.DataFrame, setores_referencia: list[str]) -> list[str]:
    setores_op = _extrair_setores(df_op)
    if not setores_referencia:
        return setores_op

    chaves_ref = {_normalizar_setor_chave(s) for s in setores_referencia if str(s).strip()}
    return [setor for setor in setores_op if _normalizar_setor_chave(setor) in chaves_ref]


def _producao_por_periodo(df_op: pd.DataFrame) -> tuple[pd.DataFrame, pd.Timestamp]:
    data_base = df_op["data_referencia"].max()
    cortes = {
        "producao_12_meses": data_base - pd.DateOffset(months=12),
        "producao_6_meses": data_base - pd.DateOffset(months=6),
        "producao_3_meses": data_base - pd.DateOffset(months=3),
    }

    producao = pd.DataFrame({"cod_produto_pai": sorted(df_op["cod_produto_pai"].unique())})
    for coluna_saida, data_corte in cortes.items():
        parcial = (
            df_op.loc[df_op["data_referencia"] >= data_corte]
            .groupby("cod_produto_pai", as_index=False)["qtd_produzida"]
            .sum()
            .rename(columns={"qtd_produzida": coluna_saida})
        )
        producao = producao.merge(parcial, on="cod_produto_pai", how="left")

    for coluna in ["producao_12_meses", "producao_6_meses", "producao_3_meses"]:
        producao[coluna] = producao[coluna].fillna(0.0)

    return producao, data_base


def _gerar_resultado_explosao(df_lt: pd.DataFrame, df_producao: pd.DataFrame) -> pd.DataFrame:
    base = df_lt.merge(df_producao, on="cod_produto_pai", how="left")
    for coluna in ["producao_12_meses", "producao_6_meses", "producao_3_meses"]:
        base[coluna] = base[coluna].fillna(0.0)

    base["consumo_mp_12_meses"] = base["qtd_mp"] * base["producao_12_meses"]
    base["consumo_mp_6_meses"] = base["qtd_mp"] * base["producao_6_meses"]
    base["consumo_mp_3_meses"] = base["qtd_mp"] * base["producao_3_meses"]

    resultado = (
        base.groupby(["cod_materia_prima", "desc_materia_prima"], as_index=False)[
            ["consumo_mp_12_meses", "consumo_mp_6_meses", "consumo_mp_3_meses"]
        ]
        .sum()
        .rename(
            columns={
                "cod_materia_prima": "codigo_materia_prima",
                "desc_materia_prima": "descricao_materia_prima",
            }
        )
    )

    resultado["media_12_meses"] = resultado["consumo_mp_12_meses"] / 12.0
    resultado["media_6_meses"] = resultado["consumo_mp_6_meses"] / 6.0
    resultado["media_3_meses"] = resultado["consumo_mp_3_meses"] / 3.0
    resultado["maior_media"] = resultado[["media_12_meses", "media_6_meses", "media_3_meses"]].max(axis=1)
    resultado = resultado.loc[resultado["maior_media"] > 0].copy()

    resultado = resultado[
        [
            "codigo_materia_prima",
            "descricao_materia_prima",
            "consumo_mp_12_meses",
            "consumo_mp_6_meses",
            "consumo_mp_3_meses",
            "media_12_meses",
            "media_6_meses",
            "media_3_meses",
            "maior_media",
        ]
    ]
    resultado = resultado.sort_values("maior_media", ascending=False).reset_index(drop=True)

    colunas_numericas = [
        "consumo_mp_12_meses",
        "consumo_mp_6_meses",
        "consumo_mp_3_meses",
        "media_12_meses",
        "media_6_meses",
        "media_3_meses",
        "maior_media",
    ]
    resultado[colunas_numericas] = resultado[colunas_numericas].round(4)

    return resultado


class CacheEnderecosApp:
    def __init__(self) -> None:
        _garantir_ambiente_execucao()
        self.root = tk.Tk()
        self.root.title("Cache e explosao de arquivos")
        self.root.geometry("980x660")
        self.root.minsize(800, 560)

        self.enderecos_listas: list[str] = []
        self.endereco_ops: str = ""
        self.setores_referencia: list[str] = _carregar_setores_referencia()
        self.setores_ops: list[str] = []
        self.setor_vars: dict[str, tk.BooleanVar] = {}
        self.animacao_ativa = False
        self.animacao_after_id: str | None = None
        self.animacao_indice = 0
        self.animacao_contexto = "Setores: todos"
        self.animacao_frames = [
            "[=      ]",
            "[==     ]",
            "[===    ]",
            "[ ===   ]",
            "[  ===  ]",
            "[   === ]",
            "[    ===]",
            "[     ==]",
            "[      =]",
            "[     < ]",
            "[    << ]",
            "[   <<< ]",
            "[  <<<< ]",
            "[ <<<   ]",
            "[ <<    ]",
            "[ <     ]",
        ]
        self.animacao_botao_frames = [
            "Fazendo explosao",
            "Fazendo explosao.",
            "Fazendo explosao..",
            "Fazendo explosao...",
        ]

        self._montar_tela()
        self._carregar_cache_inicial()

    def _montar_tela(self) -> None:
        topo = tk.Frame(self.root, padx=12, pady=10)
        topo.pack(fill="x")

        titulo = tk.Label(
            topo,
            text="Pagina inicial: carregamento por categoria",
            font=("Segoe UI", 12, "bold"),
            anchor="w",
        )
        titulo.pack(fill="x")

        subtitulo = tk.Label(
            topo,
            text="Arquivos nao sao copiados para cache. Somente enderecos sao salvos.",
            font=("Segoe UI", 10),
            anchor="w",
        )
        subtitulo.pack(fill="x", pady=(2, 0))

        botoes_iniciais = tk.Frame(self.root, padx=12, pady=8)
        botoes_iniciais.pack(fill="x")

        tk.Button(
            botoes_iniciais,
            text="Carregar listas",
            width=20,
            command=self.carregar_listas,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            botoes_iniciais,
            text="Carregar ops",
            width=20,
            command=self.carregar_ops,
        ).pack(side="left", padx=(0, 8))

        self.botao_explosao = tk.Button(
            botoes_iniciais,
            text="Fazer explosao",
            width=20,
            command=self.fazer_explosao,
            bg="#d9ead3",
        )
        self.botao_explosao.pack(side="left", padx=(0, 8))

        conteudo = tk.Frame(self.root, padx=12, pady=8)
        conteudo.pack(fill="both", expand=True)

        secao_listas = tk.LabelFrame(
            conteudo,
            text="Listas tecnicas (aceita varios arquivos)",
            padx=8,
            pady=8,
        )
        secao_listas.pack(fill="both", expand=True)

        lista_frame = tk.Frame(secao_listas)
        lista_frame.pack(fill="both", expand=True)

        self.lista = tk.Listbox(lista_frame, selectmode=tk.EXTENDED, font=("Consolas", 10))
        self.lista.pack(side="left", fill="both", expand=True)

        scroll = tk.Scrollbar(lista_frame, orient="vertical", command=self.lista.yview)
        scroll.pack(side="right", fill="y")
        self.lista.config(yscrollcommand=scroll.set)

        acoes_listas = tk.Frame(secao_listas)
        acoes_listas.pack(fill="x", pady=(8, 0))

        tk.Button(
            acoes_listas,
            text="Remover selecionados",
            width=20,
            command=self.remover_listas_selecionadas,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            acoes_listas,
            text="Limpar listas",
            width=16,
            command=self.limpar_listas,
        ).pack(side="left")

        secao_ops = tk.LabelFrame(
            conteudo,
            text="OP (aceita apenas 1 arquivo)",
            padx=8,
            pady=8,
        )
        secao_ops.pack(fill="x", pady=(10, 0))

        ops_top = tk.Frame(secao_ops)
        ops_top.pack(fill="x")

        self.ops_var = tk.StringVar(value="")
        ops_entry = tk.Entry(ops_top, textvariable=self.ops_var, state="readonly")
        ops_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        tk.Button(
            ops_top,
            text="Limpar ops",
            width=14,
            command=self.limpar_ops,
        ).pack(side="left")

        setor_frame = tk.Frame(secao_ops)
        setor_frame.pack(fill="x", pady=(8, 0))

        tk.Label(
            setor_frame,
            text="Setores para explosao (multiplos):",
            width=30,
            anchor="w",
        ).pack(anchor="w")

        setores_lista_frame = tk.Frame(setor_frame)
        setores_lista_frame.pack(fill="x", pady=(4, 0))

        self.setores_checks_frame = tk.Frame(
            setores_lista_frame,
            bd=1,
            relief="sunken",
            padx=4,
            pady=4,
        )
        self.setores_checks_frame.pack(fill="x", expand=True)

        acoes_setor = tk.Frame(setor_frame)
        acoes_setor.pack(fill="x", pady=(6, 0))

        tk.Button(
            acoes_setor,
            text="Selecionar todos",
            width=18,
            command=self._selecionar_todos_setores,
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            acoes_setor,
            text="Limpar selecao",
            width=16,
            command=self._limpar_selecao_setores,
        ).pack(side="left")

        self.setores_resumo_var = tk.StringVar(value="Nenhum setor selecionado.")
        self.setores_resumo_label = tk.Label(
            setor_frame,
            textvariable=self.setores_resumo_var,
            anchor="w",
            fg="#134f5c",
        )
        self.setores_resumo_label.pack(fill="x", pady=(4, 0))

        acoes_cache = tk.Frame(self.root, padx=12, pady=8)
        acoes_cache.pack(fill="x")

        tk.Button(
            acoes_cache,
            text="Salvar cache",
            width=14,
            command=self.salvar_cache,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            acoes_cache,
            text="Recarregar cache",
            width=16,
            command=self.carregar_cache,
        ).pack(side="left")

        rodape = tk.Frame(self.root, padx=12, pady=0)
        rodape.pack(fill="x", pady=(0, 12))

        self.status = tk.Label(rodape, text="", anchor="w", fg="#0b5394")
        self.status.pack(fill="x")
        self.animacao_var = tk.StringVar(value="")
        self.animacao_label = tk.Label(rodape, textvariable=self.animacao_var, anchor="w", fg="#38761d")
        self.animacao_label.pack(fill="x", pady=(4, 0))
        self.progresso = ttk.Progressbar(rodape, mode="indeterminate")
        self.progresso.pack(fill="x", pady=(4, 0))

    def _carregar_cache_inicial(self) -> None:
        self.carregar_cache(exibir_msg=False)

    def _atualizar_listas(self) -> None:
        self.lista.delete(0, tk.END)
        for endereco in self.enderecos_listas:
            self.lista.insert(tk.END, endereco)
        self._atualizar_status_total()

    def _atualizar_ops(self) -> None:
        self.ops_var.set(self.endereco_ops)
        self._atualizar_setores_da_op()
        self._atualizar_status_total()

    def _atualizar_status_total(self) -> None:
        total_listas = len(self.enderecos_listas)
        total_ops = 1 if self.endereco_ops else 0
        self._set_status(f"Cache atual -> listas: {total_listas} | ops: {total_ops}")

    def _set_status(self, texto: str) -> None:
        self.status.config(text=texto)

    def _aplicar_setores_em_checks(self, setores: list[str]) -> None:
        self.setores_ops = setores
        self.setor_vars = {}

        for widget in self.setores_checks_frame.winfo_children():
            widget.destroy()

        if not self.endereco_ops or not setores:
            if self.endereco_ops:
                self.setores_resumo_var.set(
                    f"Nenhum setor disponivel em {SETORES_FILE.name}. Verifique o arquivo."
                )
            else:
                self.setores_resumo_var.set("Carregue um arquivo de OP para listar setores.")
            return

        for setor in setores:
            var = tk.BooleanVar(value=True)
            self.setor_vars[setor] = var
            chk = tk.Checkbutton(
                self.setores_checks_frame,
                text=setor,
                variable=var,
                anchor="w",
                command=self._atualizar_resumo_setores,
            )
            chk.pack(fill="x", anchor="w", padx=2, pady=1)

        self._atualizar_resumo_setores()

    def _selecionar_todos_setores(self) -> None:
        if not self.setor_vars:
            return
        for var in self.setor_vars.values():
            var.set(True)
        self._atualizar_resumo_setores()

    def _limpar_selecao_setores(self) -> None:
        if not self.setor_vars:
            return
        for var in self.setor_vars.values():
            var.set(False)
        self._atualizar_resumo_setores()

    def _obter_setores_selecionados(self) -> list[str]:
        return [setor for setor, var in self.setor_vars.items() if var.get()]

    def _atualizar_resumo_setores(self) -> None:
        selecionados = self._obter_setores_selecionados()
        if not selecionados:
            self.setores_resumo_var.set("Nenhum setor selecionado.")
            return
        if len(selecionados) == 1:
            self.setores_resumo_var.set(f"1 setor selecionado: {selecionados[0]}")
            return
        self.setores_resumo_var.set(f"{len(selecionados)} setores selecionados.")

    def _atualizar_setores_da_op(self) -> None:
        if not self.endereco_ops:
            self._aplicar_setores_em_checks([])
            return

        try:
            # Recarrega setores.csv a cada atualizacao para refletir alteracoes sem reiniciar app.
            self.setores_referencia = _carregar_setores_referencia()
            if self.setores_referencia:
                # Exibe os setores do CSV para marcacao; o match com OP ocorre na explosao.
                self._aplicar_setores_em_checks(self.setores_referencia)
                self._set_status(
                    f"Setores carregados de {SETORES_FILE}: {len(self.setores_referencia)}"
                )
            else:
                # Fallback: se CSV estiver vazio/invalido, tenta levantar setores direto da OP.
                df_op = _preparar_op(_ler_tabular(Path(self.endereco_ops)))
                setores_op = _extrair_setores(df_op)
                self._aplicar_setores_em_checks(setores_op)
                if setores_op:
                    self._set_status(
                        f"{SETORES_FILE.name} vazio/invalido. Setores carregados da OP: {len(setores_op)}"
                    )
                else:
                    self._set_status(
                        f"Nao foi possivel obter setores de {SETORES_FILE.name} nem da OP."
                    )
        except Exception:
            # Mantem a tela funcional mesmo quando o arquivo ainda nao esta legivel.
            self._aplicar_setores_em_checks([])
            self._set_status("Nao foi possivel ler setores da OP. Verifique o arquivo de OP.")

    def _iniciar_animacao(self, setores: list[str]) -> None:
        self.animacao_ativa = True
        self.animacao_indice = 0
        if setores:
            if len(setores) <= 2:
                self.animacao_contexto = f"Setores: {', '.join(setores)}"
            else:
                self.animacao_contexto = f"Setores: {len(setores)} selecionados"
        else:
            self.animacao_contexto = "Setores: nenhum"
        self.progresso.start(12)
        self.botao_explosao.config(text=self.animacao_botao_frames[0])
        self._rodar_animacao()

    def _rodar_animacao(self) -> None:
        if not self.animacao_ativa:
            return
        quadro = self.animacao_frames[self.animacao_indice % len(self.animacao_frames)]
        quadro_botao = self.animacao_botao_frames[self.animacao_indice % len(self.animacao_botao_frames)]
        self.animacao_indice += 1
        self.animacao_var.set(f"Explodindo relatorio {quadro}  {self.animacao_contexto} | Aguarde...")
        self.botao_explosao.config(text=quadro_botao)
        self.animacao_after_id = self.root.after(120, self._rodar_animacao)

    def _parar_animacao(self) -> None:
        self.animacao_ativa = False
        if self.animacao_after_id is not None:
            try:
                self.root.after_cancel(self.animacao_after_id)
            except Exception:
                pass
        self.animacao_after_id = None
        self.progresso.stop()
        self.animacao_var.set("")
        self.botao_explosao.config(text="Fazer explosao")

    def _oferecer_salvar_arquivo(self, arquivo_saida: Path) -> None:
        try:
            escolha = messagebox.askyesnocancel(
                "Salvar arquivo",
                "Deseja salvar o CSV em outra pasta?\n\n"
                "Sim: escolher pasta.\n"
                "Nao: salvar na pasta Downloads deste computador.",
            )
            if escolha is None:
                self._set_status("Explosao concluida. Salvamento opcional cancelado.")
                return

            if escolha:
                destino = filedialog.asksaveasfilename(
                    title="Salvar explosao como",
                    initialfile=arquivo_saida.name,
                    defaultextension=".csv",
                    filetypes=[("CSV", "*.csv"), ("Todos os arquivos", "*.*")],
                )
                if not destino:
                    self._set_status("Explosao concluida. Nenhum destino adicional selecionado.")
                    return

                caminho_destino = Path(destino)
                shutil.copy2(arquivo_saida, caminho_destino)
                self._set_status(f"Explosao concluida e salva em: {caminho_destino}")
                messagebox.showinfo("Arquivo salvo", f"Arquivo salvo em:\n{caminho_destino}")
                return

            pasta_downloads = _obter_pasta_downloads()
            pasta_downloads.mkdir(parents=True, exist_ok=True)
            destino_download = _caminho_unico(pasta_downloads / arquivo_saida.name)
            shutil.copy2(arquivo_saida, destino_download)
            self._set_status(f"Explosao concluida. Arquivo salvo em Downloads: {destino_download}")
            messagebox.showinfo("Arquivo salvo", f"Arquivo salvo em Downloads:\n{destino_download}")
        except Exception as exc:
            self._set_status(f"Explosao concluida, mas houve falha ao salvar copia: {exc}")
            messagebox.showerror("Falha ao salvar", f"Nao foi possivel salvar copia do arquivo:\n{exc}")

    def carregar_listas(self) -> None:
        selecionados = self.root.tk.splitlist(
            filedialog.askopenfilenames(
                title="Selecione um ou mais arquivos de lista tecnica",
            )
        )
        if not selecionados:
            return

        novos = 0
        for caminho in selecionados:
            if caminho not in self.enderecos_listas:
                self.enderecos_listas.append(caminho)
                novos += 1

        self._atualizar_listas()
        self._set_status(f"{novos} arquivo(s) adicionados em LISTAS.")

    def carregar_ops(self) -> None:
        selecionado = filedialog.askopenfilename(
            title="Selecione 1 arquivo de OP",
        )
        if not selecionado:
            return

        self.endereco_ops = selecionado
        self._atualizar_ops()
        total_setores = len(self.setores_ops)
        self._set_status(f"Arquivo de OPS atualizado. Setores encontrados: {total_setores}")

    def remover_listas_selecionadas(self) -> None:
        selecionados = list(self.lista.curselection())
        if not selecionados:
            self._set_status("Nenhum item selecionado para remover.")
            return

        for indice in reversed(selecionados):
            del self.enderecos_listas[indice]

        self._atualizar_listas()
        self._set_status("Itens selecionados removidos de LISTAS.")

    def limpar_listas(self) -> None:
        if not self.enderecos_listas:
            self._set_status("A lista de LISTAS ja esta vazia.")
            return

        confirmar = messagebox.askyesno(
            "Limpar listas",
            "Deseja limpar todos os enderecos da categoria LISTAS?",
        )
        if not confirmar:
            return

        self.enderecos_listas = []
        self._atualizar_listas()
        self._set_status("Categoria LISTAS limpa.")

    def limpar_ops(self) -> None:
        if not self.endereco_ops:
            self._set_status("A categoria OPS ja esta vazia.")
            return

        self.endereco_ops = ""
        self._atualizar_ops()
        self._set_status("Categoria OPS limpa.")

    def salvar_cache(self) -> None:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        payload = {
            "listas": self.enderecos_listas,
            "ops": self.endereco_ops,
        }
        with CACHE_FILE.open("w", encoding="utf-8") as arquivo:
            json.dump(payload, arquivo, ensure_ascii=False, indent=2)

        self._set_status(f"Cache salvo em: {CACHE_FILE}")
        messagebox.showinfo("Cache salvo", f"Enderecos salvos em:\n{CACHE_FILE}")

    def _carregar_cache_arquivo(self, caminho_cache: Path) -> dict | list:
        with caminho_cache.open("r", encoding="utf-8") as arquivo:
            return json.load(arquivo)

    def carregar_cache(self, exibir_msg: bool = True) -> None:
        caminho_cache = CACHE_FILE
        if not caminho_cache.exists():
            for legado in LEGACY_CACHE_FILES:
                if legado.exists():
                    caminho_cache = legado
                    break

        if not caminho_cache.exists():
            self.enderecos_listas = []
            self.endereco_ops = ""
            self._atualizar_listas()
            self._atualizar_ops()
            if exibir_msg:
                self._set_status("Arquivo de cache ainda nao existe.")
            return

        try:
            dados = self._carregar_cache_arquivo(caminho_cache)

            if isinstance(dados, list):
                self.enderecos_listas = [str(item) for item in dados]
                self.endereco_ops = ""
            elif isinstance(dados, dict):
                listas = dados.get("listas", [])
                ops = dados.get("ops", "")
                if not isinstance(listas, list):
                    raise ValueError("Campo 'listas' invalido no cache.")
                self.enderecos_listas = [str(item) for item in listas]
                self.endereco_ops = str(ops) if ops else ""
            else:
                raise ValueError("Estrutura invalida de cache.")

            self._atualizar_listas()
            self._atualizar_ops()
            if exibir_msg:
                self._set_status(f"Cache carregado de: {caminho_cache}")
        except Exception as exc:
            self.enderecos_listas = []
            self.endereco_ops = ""
            self._atualizar_listas()
            self._atualizar_ops()
            messagebox.showerror("Erro ao carregar cache", f"Falha ao carregar cache:\n{exc}")

    def _copiar_arquivos_para_temp(self) -> tuple[Path, list[Path], Path]:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        pasta_execucao = TEMP_DIR / timestamp
        pasta_execucao.mkdir(parents=True, exist_ok=True)

        arquivos_listas_temp: list[Path] = []
        for indice, caminho in enumerate(self.enderecos_listas, start=1):
            origem = Path(caminho)
            if not origem.exists():
                raise FileNotFoundError(f"Arquivo de lista nao encontrado: {origem}")
            destino = pasta_execucao / f"lista_{indice:03d}_{origem.name}"
            shutil.copy2(origem, destino)
            arquivos_listas_temp.append(destino)

        origem_op = Path(self.endereco_ops)
        if not origem_op.exists():
            raise FileNotFoundError(f"Arquivo de OP nao encontrado: {origem_op}")
        arquivo_op_temp = pasta_execucao / f"op_{origem_op.name}"
        shutil.copy2(origem_op, arquivo_op_temp)

        return pasta_execucao, arquivos_listas_temp, arquivo_op_temp

    def _processar_explosao(self, setores_filtro: list[str]) -> tuple[Path, Path, pd.Timestamp, int]:
        pasta_temp_execucao, listas_temp, op_temp = self._copiar_arquivos_para_temp()

        dfs_listas = []
        for caminho_lista in listas_temp:
            df_bruto = _ler_tabular(caminho_lista)
            df_lista = _preparar_lista_tecnica(df_bruto)
            dfs_listas.append(df_lista)

        if not dfs_listas:
            raise ValueError("Nenhuma lista tecnica valida foi processada.")

        df_listas_unificado = pd.concat(dfs_listas, ignore_index=True)
        df_listas_unificado = (
            df_listas_unificado.groupby(
                ["cod_produto_pai", "cod_materia_prima", "desc_materia_prima"],
                as_index=False,
            )["qtd_mp"]
            .sum()
        )

        df_op_bruto = _ler_tabular(op_temp)
        df_op = _preparar_op(df_op_bruto)
        if setores_filtro:
            chaves_setores = {_normalizar_setor_chave(s) for s in setores_filtro}
            chaves_op = df_op["setor"].map(_normalizar_setor_chave)
            df_op = df_op.loc[chaves_op.isin(chaves_setores)].copy()
            if df_op.empty:
                setores_op = _extrair_setores(_preparar_op(df_op_bruto))
                sugestao = ", ".join(setores_op[:8]) if setores_op else "nenhum setor identificado na OP"
                raise ValueError(
                    "Nenhuma OP valida encontrada para os setores selecionados. "
                    f"Setores na coluna 7 da OP: {sugestao}"
                )
        else:
            raise ValueError("Selecione ao menos 1 setor para fazer a explosao.")

        df_producao, data_base = _producao_por_periodo(df_op)

        df_resultado = _gerar_resultado_explosao(df_listas_unificado, df_producao)

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        df_resultado.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig", sep=";", decimal=",")

        return OUTPUT_FILE, pasta_temp_execucao, data_base, len(df_resultado)

    def _finalizar_explosao_sucesso(
        self,
        arquivo_saida: Path,
        pasta_temp_execucao: Path,
        data_base: pd.Timestamp,
        total_itens: int,
        setores_filtro: list[str],
    ) -> None:
        self._parar_animacao()
        self.botao_explosao.config(state="normal")
        if len(setores_filtro) <= 2:
            setor_txt = ", ".join(setores_filtro)
        else:
            setor_txt = f"{len(setores_filtro)} setores"
        self._set_status(
            f"Explosao concluida com sucesso. Itens: {total_itens} | Data base: {data_base.date()} | Setores: {setor_txt}"
        )
        messagebox.showinfo(
            "Explosao concluida",
            "Arquivo gerado com sucesso.\n\n"
            f"Saida:\n{arquivo_saida}\n\n"
            f"Temp da execucao:\n{pasta_temp_execucao}",
        )
        self._oferecer_salvar_arquivo(arquivo_saida)

    def _finalizar_explosao_erro(self, exc: Exception, detalhes: str) -> None:
        self._parar_animacao()
        self.botao_explosao.config(state="normal")
        self._set_status(f"Falha na explosao: {exc}")
        messagebox.showerror("Erro na explosao", f"{exc}")
        print(detalhes)

    def fazer_explosao(self) -> None:
        if not self.enderecos_listas:
            messagebox.showwarning("Listas vazias", "Carregue ao menos 1 arquivo na categoria LISTAS.")
            return
        if not self.endereco_ops:
            messagebox.showwarning("OP vazio", "Carregue 1 arquivo na categoria OPS.")
            return

        setores_filtro = self._obter_setores_selecionados()
        if not setores_filtro:
            messagebox.showwarning("Setor vazio", "Selecione ao menos 1 setor para fazer a explosao.")
            return

        self.botao_explosao.config(state="disabled")
        self._set_status("Executando explosao... aguarde.")
        self._iniciar_animacao(setores_filtro)

        def _run() -> None:
            try:
                arquivo_saida, pasta_temp_execucao, data_base, total_itens = self._processar_explosao(
                    setores_filtro=setores_filtro
                )
                self.root.after(
                    0,
                    lambda: self._finalizar_explosao_sucesso(
                        arquivo_saida=arquivo_saida,
                        pasta_temp_execucao=pasta_temp_execucao,
                        data_base=data_base,
                        total_itens=total_itens,
                        setores_filtro=setores_filtro,
                    ),
                )
            except Exception as exc:
                detalhes = traceback.format_exc()
                self.root.after(0, lambda: self._finalizar_explosao_erro(exc, detalhes))

        threading.Thread(target=_run, daemon=True).start()

    def executar(self) -> None:
        self.root.mainloop()


def main() -> None:
    app = CacheEnderecosApp()
    app.executar()


if __name__ == "__main__":
    main()

