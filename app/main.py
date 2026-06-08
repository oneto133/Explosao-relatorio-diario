from .estoque import (
    ColetarEstoque,
    agrupar_diversos
)

from .relatorio import (
    CopiarRelatorio as copiar,
    extraircsv as extracao,
    relatorioDiversos
)

import asyncio



class main:
    """
    Arquivo responsável por orquetrar scraping, limpeza e carregamentos de dados do relatório diário
    """
    def __init__(self):
        pass

    async def main_async(self):
        try:
            await asyncio.gather(
                self.coletar_estoque(),
                self.atualizar_dados()
            )
        except asyncio.CancelledError:
            print("Execução cancelada.")
            raise

    async def coletar_estoque(self):
        await asyncio.to_thread(ColetarEstoque().executar)
        print("Coleta de estoque concluída!")

        await asyncio.to_thread(agrupar_diversos)
        print("Estoque diversos agrupados")


    async def atualizar_dados(self):
        await asyncio.to_thread(copiar().relatorio_vendas)
        print("Relatório copiado")

        await asyncio.to_thread(extracao)
        print("Extração realizada")
        print("Relatório Consolidado")
        



if __name__ == "__main__":
    programa = main()
    try:
        asyncio.run(programa.main_async())
    except KeyboardInterrupt:
        print("Execução interrompida pelo usuário.")
