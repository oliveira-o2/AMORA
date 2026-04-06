# Dashboard AMORA

Dashboard operacional para carga de bases, configuração de regras financeiras, versionamento de snapshots e análise gerencial.

## Estrutura

- `app/`: front-end estático.
- `apps-script/`: camada de persistência e leitura no Google Sheets.
- `docs/`: arquitetura, contrato de dados e handoff UX.

## Fluxo operacional

1. Carregar a base na página `Base`.
2. Revisar regras e CFOPs em `Configurações`.
3. Salvar snapshot oficial.
4. Consultar histórico em `Versões`.
5. Ler KPIs e drill-down em `Análises`.

## Como rodar localmente

1. Abra um servidor estático dentro de `dashboard/app/`.
2. Exemplo com PowerShell:

```powershell
cd dashboard/app
python -m http.server 8080
```

3. Acesse `http://localhost:8080`.

## Integração Google

O front-end funciona sem back-end para carga e análise local. Para snapshots compartilhados e padrão global, publique o Web App do Apps Script e informe a URL no próprio dashboard.

Guia de implantação: [apps-script/README.md](./apps-script/README.md)
