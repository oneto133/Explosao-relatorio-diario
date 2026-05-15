import pyautogui as pi
from time import sleep
import tkinter as tk
from tkinter import messagebox
from datetime import date, datetime, timedelta

#Tempo de coleta de dados 03:15 segundos
#Tempo de coleta de dados 03:19 segundos
#Tempo de coleta de dados 23/02/2026 - 03:12 segundos
    
class main:
    modelo8_localestx = 1053
    modelo8_localesty = 749
    
    def __init__(self):
        pass

    def executar(self):
        root = tk.Tk()
        root.withdraw()
        sleep(1)
        self.abrir_piloto()
        sleep(0.1)
        self.consulta_estoque()
        self.controle()

    def controle(self):
        locais = {"10": r"z:\pcp\05 - mOTORES\eSTOQUE\motores.XLSX", 
                  "4": r"z:\pcp\05 - mOTORES\eSTOQUE\almoxarifado.XLSX",
                  "15": r"z:\pcp\05 - mOTORES\eSTOQUE\dIVERSOS\bobinagem.XLSX",
                  "18": r"z:\pcp\05 - mOTORES\eSTOQUE\dIVERSOS\aluminio.XLSX",
                  "19": r"z:\pcp\05 - mOTORES\eSTOQUE\dIVERSOS\zamac.XLSX",
                  "20": r"z:\pcp\05 - mOTORES\eSTOQUE\dIVERSOS\plastico.XLSX",
                  "21": r"z:\pcp\05 - mOTORES\eSTOQUE\dIVERSOS\prensas.XLSX",
                  "1": r"z:\pcp\05 - mOTORES\eSTOQUE\dIVERSOS\expedicao.XLSX"}
        indice = 1
        for chave, caminho in locais.items():
            self.parametros_de_busca(str(chave), indice)
            self.imprimir_relatorio(str(caminho), str(chave) )
            indice +=1
        
        print("fim")
        
    def abrir_piloto(self):
        pi.click(x=453, y=1064)
        sleep(0.1)
    

    def consulta_estoque(self):
        pi.click(x=92, y=28) #Compras
        sleep(0.1)
        pi.click(x=154, y=98) #Relatórios
        sleep(0.1)
        pi.click(x=392, y=125) #Produtos
        sleep(0.1)
        pi.click(x=663, y=454) #Estoue de produtos
        sleep(0.1)

    def parametros_de_busca(self, local, indice):
        sleep(4.00)
        pi.click(x=self.modelo8_localestx, y=self.modelo8_localesty)
        pi.click(x=833, y=490)
        pi.press('backspace', 2)
        pi.write(local, interval=0.4)
        pi.click(x=803, y=273)
        if indice == 1:
            pi.click(x=750, y=293) #SELECIONA o filtro para só obter dados de produto com estoque
        
        #if indice == 8:
            #pi.click(x=716, y=335) #Seleção para se obter estoque reservado na expe\eSdição
        pi.click(x=852, y=832)


    def imprimir_relatorio(self, caminho, chave):
        if chave == "10":
            sleep(1) #Tempo de carregamento do relatório
            while True:
                if self.verificar():
                    print("Não carregou")
                    sleep(10)
                    
                    continue
                
                break
        elif chave == "4":
            sleep(7) #Tempo de carregamento do relatório do almoxarifado
        elif chave == "1":
            sleep(14) #Tempo de carregamento de estoque da expedição
        elif chave in ("19", "21"):
            sleep(4) #Tempo de carregamento de estoque de ZAMAC e PRENSAS
        else:
            sleep(5.5)

        
        pi.click(x=27, y=43)
        sleep(0.2)
        pi.click(x=743, y=554)
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
        if chave == "1":
            sleep(7)
        else: 
            sleep(4)

        pi.click(x=1882, y=3) #Fechando o relatório
        sleep(0.05)
        pi.click(x=743, y=272)
        sleep(0.05)
        print(pi.position())

    def data(self):
        hoje = str(date.today().strftime("%d/%m/%Y")).replace("/","")
        ano = str((date.today() - timedelta(365)).strftime("%d/%m/%Y")).replace("/","") # ano anterior
        seis = str((date.today() - timedelta(182)).strftime("%d/%m/%Y")).replace("/","") # Seis meses
        tres = str((date.today() - timedelta(90)).strftime("%d/%m/%Y")).replace("/","") # tres meses
    
        return hoje, ano, seis, tres
    
    def verificar(self):
        img = r"img\estoqueProdutos.png"


        try:
            for pos in pi.locateAllOnScreen(img, confidence=0.9):
                centro = pi.center(pos)
                pi.click(centro)
                return True
        
        except Exception as e:

            return False
    
if __name__ == "__main__":
    programa = main()
    programa.executar()