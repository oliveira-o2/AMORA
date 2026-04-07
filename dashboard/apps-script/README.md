# Apps Script - AMORA Dashboard

Camada oficial de persistencia para snapshots, configuracao global e historico versionado usando Google Sheets.

## O que esta API expõe

### GET

- `?action=ping`
- `?action=workspace_status`
- `?action=list_versions`
- `?action=get_version&payload=...`
- `?action=get_global_config`

### POST

- `setup_workspace`
- `save_snapshot`
- `save_global_config`

## Estrutura esperada no Google Sheets

O script cria e mantem automaticamente as abas abaixo:

- `versions`
- `sales_raw`
- `stock_raw`
- `snapshot_config`
- `cfop_overrides`
- `global_config`

## Caminho mais rapido para publicar

Se o repositorio estiver em drive compartilhado de rede, execute primeiro:

```powershell
powershell -ExecutionPolicy Bypass -File .\bootstrap_local_workspace.ps1
```

Isso copia o projeto Apps Script para uma pasta local no Windows, onde o `npm` e o `clasp` costumam funcionar melhor.

### 1. Instalar dependencias

No diretorio `dashboard/apps-script`:

```bash
npm install
```

### 2. Autenticar no Google

```bash
npx clasp login
```

### 3. Criar o projeto Apps Script

```bash
npx clasp create --type standalone --title "AMORA Dashboard API"
```

Isso vai gerar o arquivo `.clasp.json` local com o `scriptId`.

### 4. Enviar o codigo

```bash
npx clasp push
```

### 5. Definir a planilha operacional

No Apps Script:

1. Abra `Project Settings`
2. Vá em `Script Properties`
3. Crie a chave `SPREADSHEET_ID`
4. Informe o ID da planilha onde o historico sera salvo

### 6. Publicar como Web App

No Apps Script:

1. `Deploy`
2. `New deployment`
3. Tipo: `Web app`
4. Execute como: sua conta
5. Acesso: conforme a politica interna do cliente
6. Copie a URL final `/exec`

### 7. Testar no dashboard

1. Cole a URL em `URL do Web App`
2. Clique em `Salvar URL`
3. Clique em `Testar API`
4. Clique em `Inicializar estrutura`

## Respostas esperadas

### Ping

```json
{
  "ok": true,
  "data": {
    "status": "ready",
    "spreadsheetId": "..."
  }
}
```

### Inicializacao

```json
{
  "ok": true,
  "data": {
    "status": "ready",
    "message": "Estrutura operacional criada e validada."
  }
}
```

## Observacoes

- Sem `SPREADSHEET_ID`, o script tenta usar a planilha ativa apenas em contexto bound.
- O front-end possui fallback local para testes, mas o fluxo oficial de equipe deve usar o Web App publicado.
- O ambiente atual nao tem `clasp` instalado. Por isso este diretorio ja ficou preparado com `package.json` e `.claspignore` para facilitar a publicacao.
- Em teste local, `npm install` falhou ao escrever no drive compartilhado. Se isso acontecer com voce, use `bootstrap_local_workspace.ps1` e publique a partir de uma pasta local.
