# Contrato de Dados

## Entrada do front-end

### Excel

- Aba contendo `Consolidado` no nome, ou primeira aba como fallback.
- Colunas reconhecidas por normalização:
  - `data`, `data_nf`, `emissao`, `data_emissao`
  - `nota`
  - `cliente`
  - `uf`
  - `nome`, `item`, `produto`
  - `quant`, `quantidade`
  - `valor`
  - `icms`, `pis`, `cofins`, `ipi`
  - `custo_total_y`, `custo_total`, `custo`
  - `cfop`
  - `tipo_cfop`, `tipo`, `natureza`

### JSON

```json
{
  "meta": {},
  "sales": [],
  "stock": []
}
```

## Endpoints do Apps Script

### `POST save_snapshot`

Payload principal:

```json
{
  "clientName": "AMORA DISTRIBUIDORA LTDA",
  "sourceFileName": "base_abril.xlsx",
  "sourceFormat": "Excel",
  "sheetName": "Consolidado",
  "summary": {
    "totalRows": 0,
    "ignoredRows": 0,
    "unclassifiedRows": 0,
    "coverageFrom": "2026-01-01",
    "coverageTo": "2026-03-31",
    "salesCount": 0,
    "stockCount": 0
  },
  "config": {},
  "cfopOverrides": [],
  "sales": [],
  "stock": [],
  "parentVersionId": ""
}
```

Resposta:

```json
{
  "ok": true,
  "version": {
    "versionId": "VER-20260406-150000-ABC123",
    "createdAt": "2026-04-06T18:00:00.000Z"
  }
}
```

### `GET list_versions`

Resposta:

```json
{
  "ok": true,
  "versions": []
}
```

### `GET get_version`

Resposta compatível com a hidratação do front:

```json
{
  "ok": true,
  "data": {
    "meta": {},
    "config": {},
    "cfopOverrides": [],
    "sales": [],
    "stock": []
  }
}
```

### `POST save_global_config`

```json
{
  "action": "save_global_config",
  "config": {},
  "cfopOverrides": []
}
```

### `GET get_global_config`

```json
{
  "ok": true,
  "data": {
    "config": {},
    "cfopOverrides": []
  }
}
```
