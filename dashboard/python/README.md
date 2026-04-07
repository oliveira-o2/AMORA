# Consolidador Local

Este diretorio permite executar o `amora_consolidador.py` de forma local e consumir o resultado direto no dashboard.

## Pre-requisitos

1. Python 3 instalado no Windows.
2. Dependencias instaladas:

```bash
pip install -r requirements.txt
```

## Como iniciar

1. Execute `start_consolidador_server.bat`.
2. O servidor local sera exposto em `http://127.0.0.1:8765`.
3. No dashboard, abra a area `Base` e use o bloco `Consolidador local`.

## Fluxo na aplicacao

1. Selecione um ou mais arquivos de NFe.
2. Opcionalmente selecione `Vendas por Item`.
3. Opcionalmente selecione `Lista de Vendas`.
4. Clique em `Processar e carregar`.
5. O resultado sera processado no Python local e carregado automaticamente no dashboard.

## Saidas

- Excel consolidado em `dashboard/output/consolidador`
- JSON compativel com o dashboard no mesmo diretorio
