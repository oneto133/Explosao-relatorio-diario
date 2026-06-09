from estoque import (
    ColetarEstoque,
    agrupar_diversos
)

from relatorio import (
    CopiarRelatorio as copiar,
    extraircsv as extracao,
    relatorioDiversos, atualizarLT
)

import asyncio
import traceback


class main:
    """
    Arquivo responsável por orquestrar scraping, limpeza e carregamentos de dados do relatório diário
    """
    def __init__(self):
        pass

    async def main_async(self):
        try:
            await asyncio.gather(
                self.coletar_estoque(),
                self.atualizar_dados(),
                return_exceptions=False
            )
        except Exception:
            print("Falha na execução:")
            traceback.print_exc()
            raise

    async def coletar_estoque(self):
        print("Iniciando coleta de estoque...")
        try:
            await asyncio.to_thread(ColetarEstoque().executar)
            print("Coleta de estoque concluída!")

            await asyncio.to_thread(agrupar_diversos)
            print("Estoque diversos agrupados")

        except BaseException:
            print("Falha durante a coleta de estoque:")
            traceback.print_exc()
            raise

    async def atualizar_dados(self):
        print("Iniciando atualização dos relatórios...")
        try:
            await asyncio.to_thread(copiar().relatorio_vendas)
            print("Relatório copiado")

            await asyncio.to_thread(extracao)
            print("Extração realizada")

            await asyncio.to_thread(atualizarLT)
            print("Lista técnica atualizada")
            print("Relatório consolidado")

            await asyncio.to_thread(relatorioDiversos().filtro)
            print("Itens de reposição carregados")
            
        except BaseException:
            print("Falha durante a atualização dos relatórios:")
            traceback.print_exc()
            raise


if __name__ == "__main__":
    programa = main()
    try:
        asyncio.run(programa.main_async())
    except KeyboardInterrupt:
        print("Execução interrompida pelo usuário.")
    except Exception:
        traceback.print_exc()
        raise
