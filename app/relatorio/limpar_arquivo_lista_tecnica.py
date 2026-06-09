from __future__ import annotations

import csv
import re
import shutil
from pathlib import Path

import pandas as pd


PROJECT_DIR = Path(__file__).resolve().parents[2]
CSV_DIR = PROJECT_DIR / "csv"
TEMP_DIR = PROJECT_DIR / "temp" / "lista_tecnica"
RE_CODIGO_7 = re.compile(r"^\d{6,7}(?:\.0+)?$")
RE_DECIMAL = re.compile(r"^-?\d+(?:\.\d+)?$")


def corrigir_mojibake(texto: object) -> object:
    if not isinstance(texto, str):
        return texto

    valor = texto.strip()
    if not valor:
        return valor

    def score(candidato: str) -> tuple[int, int, int]:
        return (
            candidato.count("Ã") + candidato.count("Â") + candidato.count("�"),
            candidato.count("\uFFFD"),
            len(candidato),
        )

    candidatos = [valor]
    for encoding in ("latin1", "cp1252"):
        try:
            reparado = valor.encode(encoding).decode("utf-8")
        except Exception:
            continue
        if reparado not in candidatos:
            candidatos.append(reparado)

    return min(candidatos, key=score)


def _texto_preservado(valor: object) -> str:
    if valor is None:
        return ""
    if pd.isna(valor):
        return ""

    texto = corrigir_mojibake(str(valor)).strip()
    if not texto:
        return ""

    if texto.endswith(".0"):
        inteiro = texto[:-2]
        if inteiro.isdigit():
            return inteiro

    return texto


def _decimal_com_virgula(valor: object) -> str:
    texto = _texto_preservado(valor)
    if not texto:
        return ""
    if RE_DECIMAL.fullmatch(texto) and "." in texto:
        return texto.replace(".", ",")
    return texto


def normalizar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [corrigir_mojibake(str(col)).strip() for col in df.columns]
    df = df.fillna("")
    for col in df.columns:
        df[col] = df[col].map(_texto_preservado)
    return df


def ler_arquivos_caminhos(paths_csv: Path) -> list[Path]:
    if not paths_csv.exists():
        raise FileNotFoundError(f"Arquivo de caminhos não encontrado: {paths_csv}")

    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            df = pd.read_csv(paths_csv, dtype=str, encoding=encoding)
            if df.empty:
                continue
            coluna = df.columns[0]
            caminhos = (
                df[coluna]
                .dropna()
                .astype(str)
                .map(corrigir_mojibake)
                .map(lambda v: v.strip().strip('"'))
            )
            resultado = [Path(c) for c in caminhos if c and c.lower() != "caminhos"]
            if resultado:
                return resultado
        except Exception as exc:
            last_error = exc
            continue

    raise RuntimeError(f"Não foi possível ler os caminhos da lista técnica: {last_error}")


def _ler_csv(path: Path) -> pd.DataFrame:
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            return pd.read_csv(
                path,
                header=None,
                dtype=str,
                keep_default_na=False,
                sep=None,
                engine="python",
                encoding=encoding,
            )
        except Exception as exc:
            last_error = exc
            continue
    raise RuntimeError(f"Falha ao ler CSV {path}: {last_error}")


def _ler_excel(path: Path) -> pd.DataFrame:
    ext = path.suffix.lower()
    engine = "openpyxl" if ext in {".xlsx", ".xlsm"} else None
    return pd.read_excel(path, header=None, dtype=str, keep_default_na=False, engine=engine)


def _normalizar_bloco_linha(valores: list[str]) -> list[list[str]]:
    blocos: list[list[str]] = []
    total = len(valores)
    if total < 5:
        return blocos

    for indice in range(0, total - 4):
        codigo = _texto_preservado(valores[indice])
        if not RE_CODIGO_7.fullmatch(codigo):
            continue

        bloco = [_texto_preservado(valores[indice + deslocamento]) for deslocamento in range(5)]
        if any(campo for campo in bloco[1:]):
            blocos.append(bloco)

    return blocos


def ler_arquivo_tabelado(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {path}")

    ext = path.suffix.lower()
    if ext in {".csv", ".txt"}:
        df = _ler_csv(path)
    elif ext in {".xlsx", ".xlsm", ".xls"}:
        df = _ler_excel(path)
    else:
        raise ValueError(f"Formato não suportado: {path.suffix}")

    df = normalizar_dataframe(df)

    registros: list[list[str]] = []
    for _, linha in df.iterrows():
        valores = [_texto_preservado(valor) for valor in linha.tolist()]
        registros.extend(_normalizar_bloco_linha(valores))

    if not registros:
        raise ValueError(f"Nenhum bloco válido de 5 colunas encontrado em {path.name}")

    return pd.DataFrame(registros, columns=["Column1", "Column2", "Column3", "Column4", "Column5"])


def copiar_para_temp(origem: Path, destino_dir: Path, indice: int) -> Path:
    destino_dir.mkdir(parents=True, exist_ok=True)
    destino = destino_dir / f"{indice:03d}_{origem.stem}{origem.suffix}"
    shutil.copy2(origem, destino)
    return destino


def salvar_csv_normalizado(df: pd.DataFrame, destino: Path) -> None:
    destino.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(destino, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_ALL)


def preparar_lista_tecnica(paths_csv: Path, output_csv: Path) -> tuple[pd.DataFrame, list[str]]:
    caminhos = ler_arquivos_caminhos(paths_csv)

    frames: list[pd.DataFrame] = []
    erros: list[str] = []

    origem_dir = TEMP_DIR / "originais"
    normalizados_dir = TEMP_DIR / "normalizados"
    origem_dir.mkdir(parents=True, exist_ok=True)
    normalizados_dir.mkdir(parents=True, exist_ok=True)

    for indice, caminho in enumerate(caminhos, start=1):
        caminho_str = corrigir_mojibake(str(caminho))
        origem = Path(caminho_str)
        if not origem.is_absolute() and not origem.exists():
            origem = (PROJECT_DIR / origem).resolve()

        if not origem.exists():
            erros.append(f"[{indice:03d}] não encontrado: {origem}")
            continue

        copia_temp = copiar_para_temp(origem, origem_dir, indice)

        try:
            df = ler_arquivo_tabelado(copia_temp)
            if "Column5" in df.columns:
                df["Column5"] = df["Column5"].map(_decimal_com_virgula)
            saida_temp = normalizados_dir / f"{indice:03d}_{origem.stem}.csv"
            salvar_csv_normalizado(df, saida_temp)
            frames.append(df)
        except Exception as exc:
            erros.append(f"[{indice:03d}] {origem}: {exc}")

    if not frames:
        erro_texto = "\n".join(erros) if erros else "nenhum arquivo pôde ser processado"
        raise RuntimeError(f"Não foi possível montar a lista técnica.\n{erro_texto}")

    resultado = pd.concat(frames, ignore_index=True)
    resultado = resultado.dropna(how="all").drop_duplicates().reset_index(drop=True)
    resultado = resultado.fillna("")
    for col in resultado.columns:
        if col == "Column5":
            resultado[col] = resultado[col].map(_decimal_com_virgula)
        else:
            resultado[col] = resultado[col].map(_texto_preservado)

    salvar_csv_normalizado(resultado, output_csv)

    return resultado, erros


def main() -> int:
    paths_csv = CSV_DIR / "caminhos_lista_tecnica.csv"
    output_lt = CSV_DIR / "lt_tabela_geral.csv"
    output_compat = CSV_DIR / "TabelaGeral.csv"

    df, erros = preparar_lista_tecnica(paths_csv, output_lt)
    salvar_csv_normalizado(df, output_compat)

    print(f"Lista técnica consolidada em {output_lt}")
    print(f"Cópia de compatibilidade gerada em {output_compat}")
    print(f"Linhas consolidadas: {len(df)}")

    if erros:
        print("Alguns arquivos foram ignorados:")
        for erro in erros:
            print(f"- {erro}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
