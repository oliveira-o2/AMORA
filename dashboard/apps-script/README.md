# Apps Script - AMORA Dashboard

Camada de persistência para snapshots e padrão global usando Google Sheets.

## O que este back-end expõe

- `GET ?action=ping`
- `GET ?action=list_versions`
- `GET ?action=get_version&payload=...`
- `GET ?action=get_global_config`
- `POST save_snapshot`
- `POST save_global_config`

## Estrutura esperada no Google Sheets

O script cria e mantém as abas abaixo:

- `versions`
- `sales_raw`
- `stock_raw`
- `snapshot_config`
- `cfop_overrides`
- `global_config`

## Implantação

1. Crie um projeto Apps Script standalone.
2. Copie `Code.gs` e `appsscript.json`.
3. Defina a Script Property `SPREADSHEET_ID` com o ID da planilha operacional.
4. Publique como Web App.
5. Permita acesso conforme a política interna do cliente.
6. Copie a URL `/exec` e salve no campo `URL do Web App` do dashboard.

## Observações

- Sem `SPREADSHEET_ID`, o script tenta usar a planilha ativa apenas em contexto bound.
- O front-end possui fallback local para testes, mas o fluxo oficial de equipe deve usar o Web App publicado.
