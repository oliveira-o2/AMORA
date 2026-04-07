const SHEET_NAMES = {
  versions: 'versions',
  salesRaw: 'sales_raw',
  stockRaw: 'stock_raw',
  snapshotConfig: 'snapshot_config',
  cfopOverrides: 'cfop_overrides',
  globalConfig: 'global_config',
};

const SHEET_HEADERS = {
  versions: [
    'version_id',
    'created_at',
    'client_name',
    'source_file_name',
    'source_format',
    'sheet_name',
    'coverage_from',
    'coverage_to',
    'total_rows',
    'ignored_rows',
    'unclassified_rows',
    'sales_count',
    'stock_count',
    'parent_version_id',
    'hash',
    'saved_by',
  ],
  sales_raw: [
    'version_id',
    'line_no',
    'date',
    'note',
    'client',
    'uf',
    'item',
    'quantity',
    'revenue',
    'icms',
    'pis',
    'cofins',
    'ipi',
    'cost',
    'margin_value',
    'margin_pct',
    'cfop',
    'type',
  ],
  stock_raw: [
    'version_id',
    'line_no',
    'date',
    'note',
    'client',
    'uf',
    'item',
    'quantity',
    'value',
    'cfop',
    'type',
  ],
  snapshot_config: [
    'version_id',
    'saved_at',
    'config_json',
  ],
  cfop_overrides: [
    'version_id',
    'code',
    'analysis_type',
    'official_type',
    'saved_at',
  ],
  global_config: [
    'updated_at',
    'config_json',
    'cfop_overrides_json',
    'updated_by',
  ],
};

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'ping';
    const payload = parsePayloadParam_(e);
    let result;

    switch (action) {
      case 'ping':
        result = { ok: true, data: getApiInfo_() };
        break;
      case 'workspace_status':
        result = { ok: true, data: getWorkspaceStatus_() };
        break;
      case 'list_versions':
        result = { ok: true, versions: listVersions_() };
        break;
      case 'get_version':
        result = { ok: true, data: getVersion_(payload) };
        break;
      case 'get_global_config':
        result = { ok: true, data: getGlobalConfig_() };
        break;
      default:
        throw new Error('Acao GET nao suportada: ' + action);
    }

    return jsonResponse_(result);
  } catch (error) {
    return jsonResponse_({ ok: false, error: error.message });
  }
}

function doPost(e) {
  try {
    const body = parsePostBody_(e);
    const action = body.action || 'ping';
    let result;

    switch (action) {
      case 'setup_workspace':
        result = { ok: true, data: setupWorkspace_() };
        break;
      case 'save_snapshot':
        result = { ok: true, version: saveSnapshot_(body) };
        break;
      case 'save_global_config':
        result = { ok: true, data: saveGlobalConfig_(body) };
        break;
      default:
        throw new Error('Acao POST nao suportada: ' + action);
    }

    return jsonResponse_(result);
  } catch (error) {
    return jsonResponse_({ ok: false, error: error.message });
  }
}

function parsePayloadParam_(e) {
  if (!e || !e.parameter || !e.parameter.payload) return {};
  try {
    return JSON.parse(e.parameter.payload);
  } catch (error) {
    throw new Error('Payload GET invalido.');
  }
}

function parsePostBody_(e) {
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
  try {
    return JSON.parse(raw);
  } catch (error) {
    throw new Error('Body JSON invalido.');
  }
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (spreadsheetId) return SpreadsheetApp.openById(spreadsheetId);

  try {
    return SpreadsheetApp.getActive();
  } catch (error) {
    throw new Error('Configure a Script Property SPREADSHEET_ID antes de publicar o Web App.');
  }
}

function getAvailableActions_() {
  return {
    get: ['ping', 'workspace_status', 'list_versions', 'get_version', 'get_global_config'],
    post: ['setup_workspace', 'save_snapshot', 'save_global_config'],
  };
}

function ensureWorkspaceSheets_() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, SHEET_NAMES.versions, SHEET_HEADERS.versions);
  ensureSheet_(ss, SHEET_NAMES.salesRaw, SHEET_HEADERS.sales_raw);
  ensureSheet_(ss, SHEET_NAMES.stockRaw, SHEET_HEADERS.stock_raw);
  ensureSheet_(ss, SHEET_NAMES.snapshotConfig, SHEET_HEADERS.snapshot_config);
  ensureSheet_(ss, SHEET_NAMES.cfopOverrides, SHEET_HEADERS.cfop_overrides);
  ensureSheet_(ss, SHEET_NAMES.globalConfig, SHEET_HEADERS.global_config);
  return ss;
}

function ensureSheet_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function countDataRows_(sheet) {
  return Math.max(sheet.getLastRow() - 1, 0);
}

function getApiInfo_() {
  const ss = ensureWorkspaceSheets_();
  return {
    status: 'ready',
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    spreadsheetUrl: ss.getUrl(),
    availableActions: getAvailableActions_(),
    sheets: buildSheetsStatus_(ss),
  };
}

function buildSheetsStatus_(ss) {
  return Object.keys(SHEET_NAMES).map(function(key) {
    const sheetName = SHEET_NAMES[key];
    const sheet = ss.getSheetByName(sheetName);
    const expectedHeaders = SHEET_HEADERS[sheetName] || [];
    const currentHeaders = sheet
      ? sheet.getRange(1, 1, 1, expectedHeaders.length).getDisplayValues()[0]
      : [];
    const headerIsReady = expectedHeaders.join('|') === currentHeaders.join('|');
    return {
      key: key,
      name: sheetName,
      exists: !!sheet,
      headerIsReady: headerIsReady,
      rowCount: sheet ? countDataRows_(sheet) : 0,
    };
  });
}

function getWorkspaceStatus_() {
  const ss = ensureWorkspaceSheets_();
  const sheets = buildSheetsStatus_(ss);
  return {
    status: sheets.every(function(sheet) { return sheet.exists && sheet.headerIsReady; }) ? 'ready' : 'incomplete',
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    spreadsheetUrl: ss.getUrl(),
    sheetCount: sheets.length,
    sheets: sheets,
  };
}

function setupWorkspace_() {
  const ss = ensureWorkspaceSheets_();
  return {
    status: 'ready',
    message: 'Estrutura operacional criada e validada.',
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    spreadsheetUrl: ss.getUrl(),
    sheets: buildSheetsStatus_(ss),
  };
}

function rowsToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) return [];
  const headers = values[0];
  return values.slice(1).map(function(row) {
    const result = {};
    headers.forEach(function(header, index) {
      result[header] = row[index];
    });
    return result;
  });
}

function appendRows_(sheet, rows) {
  if (!rows.length) return;
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

function toNumber_(value) {
  if (value === '' || value == null) return null;
  const numeric = Number(value);
  return isNaN(numeric) ? null : numeric;
}

function toInt_(value) {
  return Number(value || 0);
}

function parseJson_(value, fallback) {
  if (!value) return fallback;
  try {
    return JSON.parse(value);
  } catch (error) {
    return fallback;
  }
}

function buildVersionId_() {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Sao_Paulo', 'yyyyMMdd-HHmmss');
  return 'VER-' + stamp + '-' + Utilities.getUuid().slice(0, 6).toUpperCase();
}

function computeHash_(value) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value);
  return digest.map(function(byte) {
    const normalized = byte < 0 ? byte + 256 : byte;
    const piece = normalized.toString(16);
    return piece.length === 1 ? '0' + piece : piece;
  }).join('');
}

function saveSnapshot_(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = ensureWorkspaceSheets_();
    const versionId = buildVersionId_();
    const createdAt = new Date().toISOString();
    const savedBy = Session.getActiveUser().getEmail() || '';
    const sales = Array.isArray(input.sales) ? input.sales : [];
    const stock = Array.isArray(input.stock) ? input.stock : [];
    const summary = input.summary || {};
    const hash = computeHash_(JSON.stringify({
      clientName: input.clientName || '',
      sourceFileName: input.sourceFileName || '',
      summary: summary,
      config: input.config || {},
      cfopOverrides: input.cfopOverrides || [],
      salesCount: sales.length,
      stockCount: stock.length,
    }));

    const versionsSheet = ss.getSheetByName(SHEET_NAMES.versions);
    const salesSheet = ss.getSheetByName(SHEET_NAMES.salesRaw);
    const stockSheet = ss.getSheetByName(SHEET_NAMES.stockRaw);
    const configSheet = ss.getSheetByName(SHEET_NAMES.snapshotConfig);
    const cfopSheet = ss.getSheetByName(SHEET_NAMES.cfopOverrides);

    appendRows_(versionsSheet, [[
      versionId,
      createdAt,
      input.clientName || '',
      input.sourceFileName || '',
      input.sourceFormat || '',
      input.sheetName || '',
      summary.coverageFrom || '',
      summary.coverageTo || '',
      toInt_(summary.totalRows),
      toInt_(summary.ignoredRows),
      toInt_(summary.unclassifiedRows),
      sales.length,
      stock.length,
      input.parentVersionId || '',
      hash,
      savedBy,
    ]]);

    appendRows_(salesSheet, sales.map(function(row, index) {
      return [
        versionId,
        index + 1,
        row.date || '',
        row.note || '',
        row.client || '',
        row.uf || '',
        row.item || '',
        toNumber_(row.quantity),
        toNumber_(row.revenue),
        toNumber_(row.icms),
        toNumber_(row.pis),
        toNumber_(row.cofins),
        toNumber_(row.ipi),
        toNumber_(row.cost),
        toNumber_(row.margin_value),
        toNumber_(row.margin_pct),
        row.cfop || '',
        row.type || '',
      ];
    }));

    appendRows_(stockSheet, stock.map(function(row, index) {
      return [
        versionId,
        index + 1,
        row.date || '',
        row.note || '',
        row.client || '',
        row.uf || '',
        row.item || '',
        toNumber_(row.quantity),
        toNumber_(row.value),
        row.cfop || '',
        row.type || '',
      ];
    }));

    appendRows_(configSheet, [[
      versionId,
      createdAt,
      JSON.stringify(input.config || {}),
    ]]);

    appendRows_(cfopSheet, (Array.isArray(input.cfopOverrides) ? input.cfopOverrides : []).map(function(row) {
      return [
        versionId,
        row.code || '',
        row.analysisType || '',
        row.officialType || '',
        createdAt,
      ];
    }));

    return {
      versionId: versionId,
      createdAt: createdAt,
      clientName: input.clientName || '',
      sourceFileName: input.sourceFileName || '',
      sourceFormat: input.sourceFormat || '',
      coverageFrom: summary.coverageFrom || '',
      coverageTo: summary.coverageTo || '',
      salesCount: sales.length,
      stockCount: stock.length,
      parentVersionId: input.parentVersionId || '',
      hash: hash,
    };
  } finally {
    lock.releaseLock();
  }
}

function saveGlobalConfig_(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = ensureWorkspaceSheets_();
    const sheet = ss.getSheetByName(SHEET_NAMES.globalConfig);
    const updatedAt = new Date().toISOString();
    appendRows_(sheet, [[
      updatedAt,
      JSON.stringify(input.config || {}),
      JSON.stringify(input.cfopOverrides || []),
      Session.getActiveUser().getEmail() || '',
    ]]);
    return {
      updatedAt: updatedAt,
    };
  } finally {
    lock.releaseLock();
  }
}

function getGlobalConfig_() {
  const ss = ensureWorkspaceSheets_();
  const rows = rowsToObjects_(ss.getSheetByName(SHEET_NAMES.globalConfig));
  if (!rows.length) {
    return {
      config: {},
      cfopOverrides: [],
    };
  }

  const latest = rows.sort(function(a, b) {
    return String(b.updated_at).localeCompare(String(a.updated_at));
  })[0];

  return {
    config: parseJson_(latest.config_json, {}),
    cfopOverrides: parseJson_(latest.cfop_overrides_json, []),
  };
}

function listVersions_() {
  const ss = ensureWorkspaceSheets_();
  const rows = rowsToObjects_(ss.getSheetByName(SHEET_NAMES.versions));
  return rows
    .map(function(row) {
      return {
        versionId: row.version_id,
        createdAt: row.created_at,
        clientName: row.client_name,
        sourceFileName: row.source_file_name,
        sourceFormat: row.source_format,
        coverageFrom: row.coverage_from,
        coverageTo: row.coverage_to,
        salesCount: toInt_(row.sales_count),
        stockCount: toInt_(row.stock_count),
        parentVersionId: row.parent_version_id,
        hash: row.hash,
      };
    })
    .sort(function(a, b) {
      return String(b.createdAt).localeCompare(String(a.createdAt));
    });
}

function getVersion_(payload) {
  const versionId = payload && payload.versionId ? String(payload.versionId) : '';
  if (!versionId) throw new Error('Informe versionId para carregar o snapshot.');

  const ss = ensureWorkspaceSheets_();
  const versions = rowsToObjects_(ss.getSheetByName(SHEET_NAMES.versions));
  const meta = versions.find(function(row) { return row.version_id === versionId; });
  if (!meta) throw new Error('Versao nao encontrada: ' + versionId);

  const sales = rowsToObjects_(ss.getSheetByName(SHEET_NAMES.salesRaw))
    .filter(function(row) { return row.version_id === versionId; })
    .sort(function(a, b) { return toInt_(a.line_no) - toInt_(b.line_no); })
    .map(function(row) {
      return {
        date: row.date,
        note: row.note,
        client: row.client,
        uf: row.uf,
        item: row.item,
        quantity: toNumber_(row.quantity),
        revenue: toNumber_(row.revenue),
        icms: toNumber_(row.icms),
        pis: toNumber_(row.pis),
        cofins: toNumber_(row.cofins),
        ipi: toNumber_(row.ipi),
        cost: toNumber_(row.cost),
        margin_value: toNumber_(row.margin_value),
        margin_pct: toNumber_(row.margin_pct),
        cfop: row.cfop,
        type: row.type,
      };
    });

  const stock = rowsToObjects_(ss.getSheetByName(SHEET_NAMES.stockRaw))
    .filter(function(row) { return row.version_id === versionId; })
    .sort(function(a, b) { return toInt_(a.line_no) - toInt_(b.line_no); })
    .map(function(row) {
      return {
        date: row.date,
        note: row.note,
        client: row.client,
        uf: row.uf,
        item: row.item,
        quantity: toNumber_(row.quantity),
        value: toNumber_(row.value),
        cfop: row.cfop,
        type: row.type,
      };
    });

  const configRow = rowsToObjects_(ss.getSheetByName(SHEET_NAMES.snapshotConfig))
    .filter(function(row) { return row.version_id === versionId; })
    .sort(function(a, b) { return String(b.saved_at).localeCompare(String(a.saved_at)); })[0];

  const cfopOverrides = rowsToObjects_(ss.getSheetByName(SHEET_NAMES.cfopOverrides))
    .filter(function(row) { return row.version_id === versionId; })
    .map(function(row) {
      return {
        code: row.code,
        analysisType: row.analysis_type,
        officialType: row.official_type,
      };
    });

  return {
    meta: {
      versionId: meta.version_id,
      createdAt: meta.created_at,
      clientName: meta.client_name,
      sourceFileName: meta.source_file_name,
      sourceFormat: meta.source_format,
      sheetName: meta.sheet_name,
      coverageFrom: meta.coverage_from,
      coverageTo: meta.coverage_to,
      salesCount: toInt_(meta.sales_count),
      stockCount: toInt_(meta.stock_count),
      parentVersionId: meta.parent_version_id,
      hash: meta.hash,
    },
    config: parseJson_(configRow ? configRow.config_json : '{}', {}),
    cfopOverrides: cfopOverrides,
    sales: sales,
    stock: stock,
  };
}
