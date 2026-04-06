# Arquitetura

## Visão geral

O produto foi dividido em duas camadas:

- Front-end estático em `dashboard/app/`
- Persistência Google em `dashboard/apps-script/`

O front mantém a lógica analítica existente e adiciona governança operacional:

- carga manual de base
- configuração obrigatória
- padrão global reaproveitável
- snapshot congelado por versão
- restauração de versão histórica

## Páginas do produto

- `Base`: upload, conexão com Apps Script, resumo da base e ação de salvar snapshot.
- `Configurações`: composição de receita, tributos, margem, estoque e reclassificação de CFOP.
- `Versões`: índice de snapshots, abertura de histórico e duplicação para nova revisão.
- `Análises`: KPIs, gráficos, rankings e drill-down.

## Estado da aplicação

O `app.js` passa a controlar:

- `config`: configuração ativa da sessão
- `globalConfig`: padrão global carregado do Google ou localStorage
- `currentVersionMeta`: snapshot atualmente restaurado
- `workingVersionParentId`: origem de uma revisão em andamento
- `versions`: índice do histórico disponível
- `apiBaseUrl`: URL do Web App do Apps Script

## Regras de persistência

- Nova carga: usa `globalConfig` como ponto de partida.
- Snapshot salvo: congela `config` + `cfopOverrides` + base normalizada.
- Versão antiga aberta: restaura a regra congelada do snapshot.
- Ajuste posterior em versão aberta: transforma a sessão em revisão derivada, não sobrescreve a versão original.

## Estratégia de fallback

Para permitir validação local sem back-end publicado:

- padrão global pode ser salvo em `localStorage`
- snapshots podem ser gravados localmente para testes

Esse fallback nao substitui o fluxo oficial via Google Sheets.
