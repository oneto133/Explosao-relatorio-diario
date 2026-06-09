from __future__ import annotations

from pathlib import Path

if __package__ in {None, ""}:
    import sys

    sys.path.append(str(Path(__file__).resolve().parent))
    from limpar_arquivo_lista_tecnica import (
        CSV_DIR,
        preparar_lista_tecnica,
        salvar_csv_normalizado,
    )
else:
    from .limpar_arquivo_lista_tecnica import (
        CSV_DIR,
        preparar_lista_tecnica,
        salvar_csv_normalizado,
    )


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
