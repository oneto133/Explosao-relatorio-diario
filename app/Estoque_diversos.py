import os
import re
import shutil
import zipfile
from pathlib import Path

import pandas as pd

ARQUIVOS = {
    "Aluminio": r"Z:\PCP\05 - Motores\Estoque\Diversos\ALUMINIO.xlsx",
    "Bobinagem": r"Z:\PCP\05 - Motores\Estoque\Diversos\BOBINAGEM.xlsx",
    "Plastico": r"Z:\PCP\05 - Motores\Estoque\Diversos\PLASTICO.xlsx",
    "Prensas": r"Z:\PCP\05 - Motores\Estoque\Diversos\PRENSAS.xlsx",
    "Zamac": r"Z:\PCP\05 - Motores\Estoque\Diversos\ZAMAC.xlsx",
}


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


def _ler_excel(caminho_xlsx: Path) -> pd.DataFrame:
    _normalizar_xlsx(caminho_xlsx)
    try:
        return pd.read_excel(caminho_xlsx, header=None)
    except Exception:
        if _regravar_xlsx_com_excel(caminho_xlsx):
            return pd.read_excel(caminho_xlsx, header=None)
        raise


def _remover_casas_decimais(df: pd.DataFrame) -> pd.DataFrame:
    for coluna in df.columns:
        if coluna == "chave":
            continue

        numeros = pd.to_numeric(df[coluna], errors="coerce")
        mascara_numeros = numeros.notna()
        if not mascara_numeros.any():
            continue

        serie = df[coluna].astype("object")
        serie.loc[mascara_numeros] = numeros.loc[mascara_numeros].astype(float).astype("int64")
        df[coluna] = serie

    return df


def gerar_csv_diversos(
    arquivos: dict[str, str] | None = None,
    pasta_temp: str = r"temp\diversos",
    caminho_saida_csv: str = r"csv\diversos.csv",
) -> str:
    if arquivos is None:
        arquivos = ARQUIVOS

    temp_dir = Path(pasta_temp)
    csv_path = Path(caminho_saida_csv)
    temp_dir.mkdir(parents=True, exist_ok=True)
    csv_path.parent.mkdir(parents=True, exist_ok=True)

    tabelas: list[pd.DataFrame] = []

    for chave, caminho_origem in arquivos.items():
        if not os.path.exists(caminho_origem):
            raise FileNotFoundError(f"Arquivo não encontrado para '{chave}': {caminho_origem}")

        sufixo = Path(caminho_origem).suffix or ".xlsx"
        caminho_temp = temp_dir / f"{chave}{sufixo}"

        shutil.copy2(caminho_origem, caminho_temp)

        df = _ler_excel(caminho_temp)
        df.columns = [f"coluna_{indice + 1}" for indice in range(df.shape[1])]

        df.insert(0, "chave", chave)

        tabelas.append(df)

    if tabelas:
        diversos = pd.concat(tabelas, ignore_index=True, sort=False)
    else:
        diversos = pd.DataFrame(columns=["chave"])

    diversos = _remover_casas_decimais(diversos)

    diversos.to_csv(csv_path, index=False, encoding="utf-8-sig")

    return str(csv_path)


if __name__ == "__main__":
    try:
        saida = gerar_csv_diversos()
        print(f"Arquivo consolidado gerado com sucesso: {saida}")
    except Exception as exc:
        print(f"Erro ao gerar csv de diversos: {exc}")
        raise

