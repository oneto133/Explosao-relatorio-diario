from __future__ import annotations

import datetime as dt
import math
import csv
from calendar import monthrange
from pathlib import Path
from typing import Iterable

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
CSV_DIR = BASE_DIR / "csv"

LISTA_TECNICA_PATH = CSV_DIR / "TabelaGeral.csv"
OP_PATH = Path(r"Z:\PCP\05 - Motores\Dados\LeadTimeProd.xlsx")
OUTPUT_PATH = CSV_DIR / "explosao mes a mes .csv"

MESES_PT = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]


def _normalize_code(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    if isinstance(value, (int, float)):
        return str(int(round(float(value))))
    text = str(value).strip()
    if not text:
        return ""
    numeric_like = text.replace(".", "").replace(",", "").isdigit()
    if numeric_like:
        text_normalized = text.replace(",", ".")
        try:
            num = float(text_normalized)
            if num.is_integer():
                return str(int(num))
        except Exception:
            pass
    digits = "".join(ch for ch in text if ch.isdigit())
    return digits if digits else text


def _parse_number(value: object) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return 0.0
        return float(value)
    text = str(value).strip()
    if not text:
        return 0.0
    text = text.replace(" ", "")
    if text.count(",") == 1 and text.count(".") >= 1:
        text = text.replace(".", "").replace(",", ".")
    elif text.count(",") == 1 and text.count(".") == 0:
        text = text.replace(",", ".")
    try:
        return float(text)
    except Exception:
        return 0.0


def _parse_date(value: object) -> pd.Timestamp | pd.NaT:
    if value is None:
        return pd.NaT
    if isinstance(value, float) and math.isnan(value):
        return pd.NaT
    if isinstance(value, dt.datetime):
        return pd.Timestamp(year=value.year, month=value.month, day=value.day)
    if isinstance(value, dt.date):
        return pd.Timestamp(year=value.year, month=value.month, day=value.day)
    if isinstance(value, (int, float)):
        return pd.to_datetime(value, errors="coerce", unit="D", origin="1899-12-30")
    text = str(value).strip()
    if not text:
        return pd.NaT
    return pd.to_datetime(text, errors="coerce", dayfirst=True)


def _month_label(ts: pd.Timestamp) -> str:
    return ts.strftime("%m/%Y")


def _format_number_for_csv(value: object) -> object:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    try:
        num = float(value)
    except Exception:
        return str(value)
    return int(round(num))


def _read_lista_tecnica(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Lista técnica não encontrada: {path}")

    df_raw = pd.read_csv(path, dtype=str, sep=None, engine="python")
    if df_raw.shape[1] < 5:
        raise ValueError("TabelaGeral.csv inválida: esperado no mínimo 5 colunas.")

    if len(df_raw.columns) >= 5 and all(str(c).lower().startswith("column") for c in df_raw.columns[:5]):
        df = pd.DataFrame(
            {
                "cod_produto_pai": df_raw.iloc[:, 0],
                "codigo": df_raw.iloc[:, 2],
                "desc_materia_prima": df_raw.iloc[:, 3],
                "qtd_mp": df_raw.iloc[:, 4],
            }
        )
    else:
        # fallback por posição para manter compatível com o layout atual conhecido
        df = pd.DataFrame(
            {
                "cod_produto_pai": df_raw.iloc[:, 0],
                "codigo": df_raw.iloc[:, 2],
                "desc_materia_prima": df_raw.iloc[:, 3],
                "qtd_mp": df_raw.iloc[:, 4],
            }
        )

    df["cod_produto_pai"] = df["cod_produto_pai"].map(_normalize_code)
    df["codigo"] = df["codigo"].map(_normalize_code)
    df["desc_materia_prima"] = df["desc_materia_prima"].fillna("").astype(str).str.strip()
    df["qtd_mp"] = df["qtd_mp"].map(_parse_number)

    valid = (
        df["cod_produto_pai"].str.fullmatch(r"\d+")
        & df["codigo"].str.fullmatch(r"\d+")
        & (df["qtd_mp"] != 0)
    )
    df = df.loc[valid].copy()

    if df.empty:
        raise ValueError("Lista técnica sem linhas válidas após limpeza.")
    return df


def _read_op_via_pandas(path: Path) -> pd.DataFrame:
    # Alguns arquivos de OP quebram no openpyxl; por isso essa função pode falhar.
    df_raw = pd.read_excel(path, sheet_name=0, header=None)
    if df_raw.shape[1] < 9:
        raise ValueError("LeadTimeProd.xlsx inválido: esperado mínimo de 9 colunas.")
    return df_raw


def _read_op_via_com(path: Path) -> pd.DataFrame:
    try:
        import win32com.client as win32
    except Exception as exc:
        raise RuntimeError("win32com indisponível para fallback de leitura do OP.") from exc

    try:
        excel = win32.DispatchEx("Excel.Application")
    except Exception:
        excel = win32.Dispatch("Excel.Application")
    wb = None
    try:
        wb = excel.Workbooks.Open(str(path))
        ws = wb.Worksheets(1)
        values = ws.UsedRange.Value
    finally:
        if wb is not None:
            wb.Close(False)
        excel.Quit()

    rows = [list(row) for row in values if row and any(cell is not None for cell in row)]
    if not rows:
        raise ValueError("LeadTimeProd.xlsx vazio.")
    return pd.DataFrame(rows, dtype=object)


def _read_op(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Arquivo de OP não encontrado: {path}")

    try:
        return _read_op_via_pandas(path)
    except Exception:
        return _read_op_via_com(path)


def _prepare_op(df_raw: pd.DataFrame) -> pd.DataFrame:
    if df_raw.shape[1] < 9:
        raise ValueError("Arquivo de OP inválido: esperado mínimo de 9 colunas.")

    col_data = [_parse_date(v) for v in df_raw.iloc[:, 1].tolist()]
    col_status = [str(v).strip() if v is not None else "" for v in df_raw.iloc[:, 3].tolist()]
    col_cod = [_normalize_code(v) for v in df_raw.iloc[:, 4].tolist()]
    col_qtd = [_parse_number(v) for v in df_raw.iloc[:, 8].tolist()]

    df = pd.DataFrame(
        {
            "data_referencia": col_data,
            "status": col_status,
            "cod_produto_pai": col_cod,
            "qtd_produzida": col_qtd,
        }
    )

    valid = (
        df["cod_produto_pai"].str.fullmatch(r"\d+")
        & df["data_referencia"].notna()
        & (df["qtd_produzida"] > 0)
    )
    df = df.loc[valid].copy()

    if df.empty:
        raise ValueError("OP sem linhas válidas após limpeza.")
    return df


def _last_12_month_starts(data_base: pd.Timestamp) -> list[pd.Timestamp]:
    fim_mes_base = pd.Timestamp(year=data_base.year, month=data_base.month, day=1)
    return [fim_mes_base - pd.DateOffset(months=offset) for offset in range(11, -1, -1)]


def _build_monthly_production(df_op: pd.DataFrame, month_starts: Iterable[pd.Timestamp]) -> pd.DataFrame:
    prod = pd.DataFrame({"cod_produto_pai": sorted(df_op["cod_produto_pai"].unique())})
    for month_start in month_starts:
        month_end = month_start + pd.offsets.MonthEnd(1)
        label = _month_label(month_start)
        col_prod = f"prod_{label}"

        grouped = (
            df_op.loc[(df_op["data_referencia"] >= month_start) & (df_op["data_referencia"] <= month_end)]
            .groupby("cod_produto_pai", as_index=False)["qtd_produzida"]
            .sum()
            .rename(columns={"qtd_produzida": col_prod})
        )
        prod = prod.merge(grouped, on="cod_produto_pai", how="left")
        prod[col_prod] = prod[col_prod].fillna(0.0)
    return prod


def _explode_monthly(
    df_lt: pd.DataFrame,
    df_prod: pd.DataFrame,
    month_starts: Iterable[pd.Timestamp],
) -> pd.DataFrame:
    base = df_lt.merge(df_prod, on="cod_produto_pai", how="left")
    for month_start in month_starts:
        label = _month_label(month_start)
        col_prod = f"prod_{label}"
        base[col_prod] = base[col_prod].fillna(0.0)
        base[label] = base["qtd_mp"] * base[col_prod]

    consumo_cols = [_month_label(m) for m in month_starts]
    resultado = (
        base.groupby(["codigo", "desc_materia_prima"], as_index=False)[consumo_cols]
        .sum()
        .rename(
            columns={
                "codigo": "codigo",
                "desc_materia_prima": "descricao",
            }
        )
    )

    for month_start in month_starts:
        label = _month_label(month_start)
        days = monthrange(month_start.year, month_start.month)[1]
        resultado[f"m_{label}"] = resultado[label] / float(days)

    resultado["consumo_total_12_meses"] = resultado[consumo_cols].sum(axis=1)
    resultado["media_diaria_12_meses"] = resultado["consumo_total_12_meses"] / 365.0

    numeric_cols = [c for c in resultado.columns if c in consumo_cols or c.startswith("m_") or c.startswith("media_diaria_")]
    resultado[numeric_cols] = resultado[numeric_cols].round(4)
    resultado = resultado.loc[resultado["consumo_total_12_meses"] > 0].copy()
    resultado = resultado.sort_values("consumo_total_12_meses", ascending=False).reset_index(drop=True)
    return resultado


def main() -> None:
    df_lt = _read_lista_tecnica(LISTA_TECNICA_PATH)
    df_op_raw = _read_op(OP_PATH)
    df_op = _prepare_op(df_op_raw)

    data_base = df_op["data_referencia"].max()
    month_starts = _last_12_month_starts(data_base)

    df_prod = _build_monthly_production(df_op, month_starts)
    df_final = _explode_monthly(df_lt, df_prod, month_starts)

    colunas_numericas = [col for col in df_final.columns if col not in {"codigo", "descricao"}]
    df_saida = df_final.copy()
    for col in colunas_numericas:
        df_saida[col] = df_saida[col].map(_format_number_for_csv)

    CSV_DIR.mkdir(parents=True, exist_ok=True)
    # Escrita manual para preservar os inteiros como numéricos no CSV.
    with OUTPUT_PATH.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_NONNUMERIC)
        writer.writerow(df_saida.columns.tolist())
        for row in df_saida.itertuples(index=False, name=None):
            writer.writerow(row)

    periodo = f"{month_starts[0].strftime('%m/%Y')} até {month_starts[-1].strftime('%m/%Y')}"
    print(f"OK: explosão mês a mês gerada em {OUTPUT_PATH}")
    print(f"Período considerado: {periodo}")
    print(f"Itens (MP) gerados: {len(df_final)}")


if __name__ == "__main__":
    main()
