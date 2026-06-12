import os
from datetime import datetime
from pathlib import Path
import pandas as pd
from copiar_relatorio import main as copiar
from extrair_csvs import main as extracao
from atualizar_lista_tecnica import main as atualizarLT
from reposicao_e_diversos import main as relatorioDiversos
from time import sleep

def main() -> None:

    RELATORIO = Path(__file__).resolve().parent
    APP = RELATORIO.parent
    PAI = APP.parent
    CSV = PAI / "csv"

    while True:
        caminho = r"Z:\PCP\2.2- Relatório Semanal - NOVO.xlsb"

        timestamp = os.path.getmtime(caminho)

        modificação = datetime.fromtimestamp(timestamp)

        log_caminho = CSV / "log_atualizacao.csv"
        
        df = pd.read_csv(log_caminho, nrows=1)

        data_hora = df.columns[0]
        if str(modificação) != str(data_hora):
            executar()
            registrar_log(PAI, modificação=modificação)

        else:
             sleep(60)


def registrar_log(caminho, modificação) -> None:
        with open (caminho / "csv/log_atualizacao.csv", "w") as file:
            file.write(f"{modificação}")


def executar() -> str:

    copiar().relatorio_vendas()
    extracao()
    atualizarLT()
    relatorioDiversos().filtro()

    print("Dados atualizados")

            

if __name__ == "__main__":
    main()