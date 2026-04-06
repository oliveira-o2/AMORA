# AMORA

Repositório operacional da frente AMORA.

## Estrutura atual

- `AMORA_BASE_CONSOLIDADA.ipynb`: notebook legado preservado na raiz.
- `dashboard/`: dashboard CFO em HTML/CSS/JS com versionamento de bases via Google Sheets + Apps Script.

## Dashboard

O dashboard fica isolado em `dashboard/` para permitir evolução da interface, do motor de regras e da integração com Google sem interferir no notebook existente.

- App: [dashboard/app/index.html](./dashboard/app/index.html)
- Back-end Apps Script: [dashboard/apps-script/Code.gs](./dashboard/apps-script/Code.gs)
- Documentação: [dashboard/README.md](./dashboard/README.md)
