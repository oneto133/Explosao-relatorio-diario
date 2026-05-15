import pyautogui as pi
from coletaEstoque import main as es
from time import sleep
from datetime import date, timedelta

class main:
    def __init__(self):
        self.abrir_piloto()
        sleep(0.2)
        self.abrir_relatorio_status_op()
        sleep(0.5)
        locais = {
                  "10": {"1": r"z:\pcp\05 - mOTORES\aNALISES\prod-12\producao-12meses-motores.XLSX",
                         "2": r"z:\pcp\05 - mOTORES\aNALISES\prod-6\producao-6meses-motores.XLSX",
                         "3": r"z:\pcp\05 - mOTORES\aNALISES\prod-3\producao-3meses-motores.XLSX"},
                    "22": {"1": r"z:\pcp\05 - mOTORES\aNALISES\prod-12\producao-12meses-tornos.XLSX",
                         "2": r"z:\pcp\05 - mOTORES\aNALISES\prod-6\producao-6meses-tornos.XLSX",
                         "3": r"z:\pcp\05 - mOTORES\aNALISES\prod-3\producao-3meses-tornos.XLSX"}}
        
        for local, chave in locais.items():
            indice = 1
            for mes, caminho in chave.items():
                self.relatorio_op(local=local, inicio=self.data()[indice], fim=self.data()[0])
                self.imprimir_relatorio(caminho=caminho)
                indice += 1
        print("Fim")

    def abrir_piloto(self):
        pi.click(x=453, y=1064)

    def relatorio_op(self, local, inicio, fim):
        pi.click(x=833, y=435) #Data Inicial
        pi.press("backspace", presses=4, interval=0.1)
        sleep(0.5)
        pi.write(inicio, interval=0.1)
        sleep(0.5)
        pi.click(x=1082, y=441) #Data Final
        sleep(0.5)
        pi.press("backspace", presses=4, interval=0.1)
        pi.write(fim, interval=0.1)
        sleep(0.5)
        pi.click(x=864, y=559)
        sleep(0.5)
        pi.press("backspace", presses=4, interval=0.2)
        sleep(0.5)
        pi.write(local, interval=0.1)
        sleep(0.5)
        pi.click(x=845, y=668)
        sleep(13)
        print(pi.position())

    def data(self):
        hoje = str(date.today().strftime("%d/%m/%Y")).replace("/","")
        ano = str((date.today() - timedelta(365)).strftime("%d/%m/%Y")).replace("/","") # ano anterior
        seis = str((date.today() - timedelta(182)).strftime("%d/%m/%Y")).replace("/","") # Seis meses
        tres = str((date.today() - timedelta(90)).strftime("%d/%m/%Y")).replace("/","") # tres meses
    
        return hoje, ano, seis, tres

    def imprimir_relatorio(self, caminho):
        pi.click(x=27, y=43)
        sleep(0.2)
        pi.click(x=743, y=554)
        sleep(0.5)
        pi.click(x=1081, y=577)
        pi.press('x')
        """Tecla 'X' e depois pressiona duas vezes para baixo
        para que possa selecionar o tipo de impressão que é xlsx dados do arquivo. 
        Se mudar algo na estrutura do programa também tem que mudar aqui."""
        sleep(0.2)
        pi.press('down', presses=2)
        pi.press('enter') #seleciona o tipo de arquivo xlsx dados do arquivo
        sleep(0.05)
        pi.click(x=1175, y=609)
        sleep(1)
        pi.write(caminho, interval=0.07)
        pi.click(x=1240, y=600)
        sleep(0.05)
        pi.click(x=1080, y=660)
        sleep(0.05)
        pi.click(x=918, y=558) #Imprimindo o arquivo
        sleep(6)
        pi.click(x=1907, y=4) #Fechando o relatório
        sleep(1)
        print(pi.position())

    def abrir_relatorio_status_op(self):
        pi.click(x=19, y=28) #Industria
        sleep(0.5)
        pi.click(x=74, y=68) #Relatórios
        sleep(0.5)
        pi.click(x=220, y=73) #Status das OP´S
        sleep(0.5)

    indice = 1

if __name__ == "__main__":
    main()
