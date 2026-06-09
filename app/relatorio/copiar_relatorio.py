import os
import shutil

class main:
    def __init__(self):
        pass

    def relatorio_vendas(self):
                
        origem = r"Z:\PCP\2.2- Relatório Semanal - NOVO.xlsb"

        destino = r"temp\relatorio_semanal_temp.xlsx"
        os.makedirs(os.path.dirname(destino), exist_ok=True)


        shutil.copy2(origem, destino)

if __name__ ==" __main__":
    relatorio = main()
    relatorio.relatorio_vendas()

    