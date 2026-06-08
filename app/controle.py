try:
    from app.copiar_relatorio import main as copiar
    from app.extrair_csvs import main as extracao
    from app.coletaEstoque import main as estoque
    from app.Estoque_diversos import gerar_csv_diversos as diversos

except:
    from copiar_relatorio import main as copiar
    from extrair_csvs import main as extracao
    from coletaEstoque import main as estoque
    from Estoque_diversos import gerar_csv_diversos as diversos
finally:
    import asyncio


class main:
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
        await asyncio.to_thread(estoque().executar)
        print("Coleta de estoque concluída!")

        await asyncio.to_thread(diversos)
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
