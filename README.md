# Aplicação de apoio ao PCP

Este projeto foi organizado para automatizar tarefas do dia a dia ligadas a estoque e relatórios internos.

A parte mais importante da aplicação está dentro da pasta [`app`](app/), que concentra as rotinas que coletam dados, tratam arquivos e organizam os relatórios gerados.

## Visão geral

A aplicação foi pensada para fazer três coisas principais:

1. Coletar informações de estoque diretamente do sistema por meio de automação/scraping.
2. Gerar e organizar arquivos de relatórios em formatos mais fáceis de usar, como CSV.
3. Separar os dados por finalidade, como estoque, reposição e outros relatórios de apoio.

Em outras palavras, o objetivo é reduzir o trabalho manual na coleta e no tratamento dessas informações.

## Como a pasta `app` está organizada

### `app/main.py`

É o ponto de entrada da aplicação.

Esse arquivo funciona como um orquestrador: ele chama as rotinas de estoque e as rotinas de relatório na sequência correta.

### `app/estoque/`

Essa pasta concentra o que está ligado à extração de estoque.

- [`coletaEstoque.py`](app/estoque/coletaEstoque.py): faz a extração de estoque via scraping no sistema.
- [`Estoque_diversos.py`](app/estoque/Estoque_diversos.py): trata os dados de estoque diversos e gera o arquivo final separado.
- [`__init__.py`](app/estoque/__init__.py): facilita as importações usadas pelo restante da aplicação.

Resumo simples:
- aqui a aplicação entra no sistema
- coleta os dados
- e salva o resultado para uso posterior

### `app/relatorio/`

Essa pasta reúne as rotinas relacionadas aos relatórios.

- [`copiar_relatorio.py`](app/relatorio/copiar_relatorio.py): copia os arquivos base que serão usados no processamento.
- [`extrair_csvs.py`](app/relatorio/extrair_csvs.py): lê os arquivos de origem e gera os CSVs necessários.
- [`reposicao_e_diversos.py`](app/relatorio/reposicao_e_diversos.py): filtra os dados para montar a base de reposição e diversos.
- [`Explosao_mes_a_mes.py`](app/relatorio/Explosao_mes_a_mes.py): módulo auxiliar de análise e organização dos dados de explosão mensal para análise ao solicitar itens comprados.
- [`extrair_lead_time.py`](app/relatorio/extrair_lead_time.py): módulo auxiliar para extração de informações de lead time.
- [`__init__.py`](app/relatorio/__init__.py): organiza os imports da pasta.

Resumo simples:
- aqui a aplicação pega arquivos de origem
- transforma os dados
- e gera relatórios prontos para uso

### `app/__init__.py`

Serve para tornar a pasta `app` um pacote Python e facilitar os imports internos.

## Fluxo da aplicação

De forma resumida, o fluxo é este:

1. `app/main.py` inicia a execução.
2. As rotinas da pasta `app/estoque/` coletam os dados de estoque.
3. As rotinas da pasta `app/relatorio/` copiam, extraem e tratam os relatórios.
4. Os arquivos finais são gerados para consulta ou uso em outros processos.

## Intuito do projeto

O intuito da aplicação é automatizar etapas repetitivas que normalmente exigiriam muito trabalho manual.

Ela ajuda a:

- capturar dados do sistema com menos intervenção
- padronizar a geração de arquivos
- reduzir erros de digitação e de tratamento manual
- deixar os dados prontos para análise e uso operacional

## Leitura rápida do propósito de cada pasta

- `app/estoque`: extrair informações de estoque do sistema
- `app/relatorio`: copiar, converter e organizar relatórios
- `app/main.py`: coordenar tudo em uma única execução

## Observação

Esta documentação foi escrita para explicar o propósito da pasta `app` de forma simples e direta, sem entrar em detalhes técnicos de implementação.

## Considerações finais

Com essa aplicação, estima-se ter diminuido no mínimo 1 hora por dia coletando dados de estoque, limpando relatórios e carregando/revisando dados para outros usuários.
