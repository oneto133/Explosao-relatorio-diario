import os
import sys
import pandas as pd



def _coerce_integer_like_columns(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        series = df[col]

        # Tenta converter também colunas object (ex: "123.0")
        numeric = pd.to_numeric(series, errors="coerce")
        if numeric.notna().sum() == 0:
            continue

        # Converte para inteiro quando todos os valores numéricos são inteiros
        if (numeric.dropna() % 1 == 0).all():
            df[col] = numeric.round().astype("Int64")

    return df


def _format_integer_if_numeric(value: object) -> object:
    if pd.isna(value):
        return ""
    try:
        num = float(value)
    except Exception:
        return value
    return str(int(round(num)))


def export_sheet_to_csv(excel_path: str, sheet_name: str, csv_path: str) -> None:
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Arquivo não encontrado: {excel_path}")

    ext = os.path.splitext(excel_path)[1].lower()
    read_kwargs = {"sheet_name": sheet_name}

    # xlsb requer o engine pyxlsb
    if ext == ".xlsb":
        read_kwargs["engine"] = "pyxlsb"

    df = pd.read_excel(excel_path, **read_kwargs)
    df = _coerce_integer_like_columns(df)

    if sheet_name == "Pendências - Geral":
        df = df.map(_format_integer_if_numeric)

    # Força formatação da Column5 para não sair com ".0" quando é inteiro
    if "Column5" in df.columns:
        col5_num = pd.to_numeric(df["Column5"], errors="coerce")
        if col5_num.notna().sum() > 0:
            def _fmt_col5(value: object) -> object:
                if pd.isna(value):
                    return ""
                try:
                    num = float(value)
                except Exception:
                    return value
                if num % 1 == 0:
                    return str(int(round(num)))
                return str(num).replace(".", ",") if isinstance(value, str) else num

            df["Column5"] = df["Column5"].map(_fmt_col5)
    os.makedirs(os.path.dirname(csv_path), exist_ok=True)
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")


def limpar_relatorio_csv() -> None:
    df = pd.read_csv(r"csv/Pendencias - Geral.csv", encoding="utf-8-sig", skiprows=1)

    motores = df[(df["Departamento"] == "MOTORES") & (df["Unnamed: 11"] != 0)]
    cancelas = df[(df["Departamento"] == "CANCELAS") & (df["Unnamed: 11"] != 0)]

    novo = pd.concat([motores, cancelas])
    novo.to_csv(r"csv/relatorio_diario.csv", index=False, encoding="latin1")


def main() -> int:
    relatorio_path = r"temp/relatorio_semanal_temp.xlsx"

    export_sheet_to_csv(relatorio_path, "Pendências - Geral", os.path.join("csv", "Pendencias - Geral.csv"))

    limpar_relatorio_csv()

    print("CSV gerados com sucesso em ./csv")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Erro: {exc}")
        raise
