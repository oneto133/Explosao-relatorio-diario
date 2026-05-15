from app.coletaEstoque import main
import tkinter as tk
from tkinter import messagebox, Button
from app.controle import main as controle
import threading

class App():
    def __init__(self):
        pass

    def telaInicial(self):
        self.root = tk.Tk()
        self.root.title("Explosão de relatório diário")
        alturaJanela = 300
        larguraJanela = 300

        altura = self.root.winfo_screenwidth()
        largura = self.root.winfo_screenheight()

        posy = int((largura / 2) - (larguraJanela / 2))
        posx =  int((altura / 2) - (larguraJanela / 2))

        self.root.geometry(f"{larguraJanela}x{alturaJanela}+{posx}+{posy}")


        button_coletaEstoque = tk.Button(self.root, text="Coleta Estoque", width=20, height=5, bg="gray", font=("Arial", 10, "bold"), command=self.coletaEstoque)
        button_coletaEstoque.pack()

        self.button_relatorio = tk.Button(self.root, text="Extrair relatório", bg="gray", width=20, height=5, font=("Arial", 10, "bold"))
        self.button_relatorio.pack()
        self.button_relatorio.config(command=lambda: self.tarefa(self.chamarControle, self.button_relatorio))

        self.button_explosao = tk.Button(self.root, text="Explodir relatório diário", bg="gray", width=20, height=5, font=("Arial", 10, "bold"))
        self.button_explosao.pack()
        self.button_explosao.config(command=lambda: self.explosao())

        self.root.mainloop()

    def explosao(self):
        self.coletaEstoque()
        self.tarefa(self.chamarControle, self.button_explosao)
        

        messagebox("Relatório explodido com sucesso!")


    def tarefa(self, func, button, *args):
        button.config(state="disabled")

        def run():
            func(*args)
            self.root.after(0, lambda: button.config(state="normal"))
        threading.Thread(target=run).start()

        

    def chamarControle(self):
        relatorio = controle()
        extrairRelatorio = relatorio.atualizar_dados()


    def coletaEstoque(self):    
        digitado = messagebox.askokcancel("Coleta de dados de estoque", "Consulta de estoque\n" \
        "Certifique-se de que tudo está em perfeito estado\n\n\n" \
        "✔ O programa PILOTO está logado e com a tela limpa.\n\n" \
        "✔ Certifique-se de que há uma boa conexão com a internet\n\n" \
        "❌ Não use o mouse ou o teclado, isso pode causar mal-funcionamento na coleta de dados\n\n" \
        "Os dados serão salvos nas pastas de destinos, não é necessário abrir.\n\n" \
        "No piloto, vá em industria ->" \
                "Relatórios -> Relatórios de status das OP´s\nE extraia todas as ordens geradas desde o dia 17 de outubro" \
                " até a data atual do local de estoque 10 (Motores)\nClique em imprimir para arquivo " \
                "escolha o formato xlsx dados do arquivo, salve na pasta dados e salve no arquivo OP, após, apenas abra o arquivo de explosão" \
                " copie os dados de venda do relatório de venda, no excel, clique em atualizar tudo, aguarde alguns segundos" \
                " e a explosão estará feita."
        "\n\n\nApós a checagem, clique em ok e o programa fará a sua parte de coleta de dados\n")


        if digitado == True:
            main()
            messagebox.showinfo("Finalizado", "Coleta de estoque finalizada com sucesso\n\n\n" \
                "Os dados foram salvos nas pastas de destinos, não é necessário abrir.\n\n" \
                "No piloto, vá em industria ->" \
                "Relatórios -> Relatórios de status das OP´s\nE extraia todas as ordens geradas desde o dia 17 de outubro" \
                " até a data atual do local de estoque 10 (Motores)\nClique em imprimir para arquivo " \
                "escolha o formato xlsx dados do arquivo, salve na pasta dados e salve no arquivo OP, após, apenas abra o arquivo de explosão" \
                " copie os dados de venda do relatório de venda, no excel, clique em atualizar tudo, aguarde alguns segundos" \
                " e a explosão estará feita.")
        else:
            "Cancelado"



if __name__ == "__main__":
    programa = App()
    programa.telaInicial()