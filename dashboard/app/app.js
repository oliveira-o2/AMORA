/* ═══════════════════════════════════════════════════════════════════
   CFO DASHBOARD — app.js
   Versão: 2.0
   Melhorias:
   - Corrigido fallback perigoso em classifyType (CFOP /^[56]/ → venda)
   - Corrigido taxesTotal:0 hardcoded no rebuildRowsFromCfopRegistry
   - Eliminada dupla chamada de calculateMarginPctForRow no render
   - calculateMetrics cacheado: só recalcula quando dados mudam
   - Processamento de Excel em setTimeout para não travar a UI
   - KPIs adicionados: ticket médio, concentração top-3, mix de devolução
   - Análise CFO com thresholds calibrados e severity badges
   - Exportação CSV por tabela
   - Modal com KPI summary interno
   - Barra de carregamento visual
═══════════════════════════════════════════════════════════════════ */

const DEFAULT_CLIENT_NAME = "AMORA DISTRIBUIDORA LTDA";
const DEFAULT_API_BASE_URL = "";
const LOCAL_STORAGE_KEYS = {
  clientName: "amoraDashboardClientName",
  apiBaseUrl: "amoraDashboardApiBaseUrl",
  globalConfig: "amoraDashboardGlobalConfig",
  localVersions: "amoraDashboardLocalVersions",
};

function createDefaultConfig() {
  return {
    revenue: { sales: true, returns: true },
    taxes:   { icms: true, pis: true, cofins: true, ipi: true },
    stock:   { remessa: true, baixa: true },
    margin:  { deductTaxes: false, requireCost: true },
  };
}

function createEmptyGlobalConfig() {
  return {
    config: createDefaultConfig(),
    cfopOverrides: [],
  };
}

const PAGE_META = {
  base: {
    eyebrow: "Operação financeira",
    title: "Base e governança do dashboard",
  },
  config: {
    eyebrow: "Configuração analítica",
    title: "Regras financeiras e tratamento de CFOP",
  },
  versions: {
    eyebrow: "Histórico versionado",
    title: "Snapshots congelados e reutilização de versões",
  },
  analysis: {
    eyebrow: "Leitura executiva",
    title: "Indicadores, comparativos e drill-down gerencial",
  },
};

const state = {
  source:          null,
  baseRows:        [],
  salesRows:       [],
  stockRows:       [],
  filteredSales:   [],
  filteredStock:   [],
  compareSales:    [],
  compareStock:    [],
  metrics:         null,
  compareMetrics:  null,
  charts:          new Map(),
  modalRows:       [],
  modalMode:       "sales",
  statusTimer:     null,
  compareDirty:    false,
  clientName:      DEFAULT_CLIENT_NAME,
  activePage:      "base",
  cfopRegistry:    new Map(),
  config:          createDefaultConfig(),
  globalConfig:    createEmptyGlobalConfig(),
  apiBaseUrl:      DEFAULT_API_BASE_URL,
  versions:        [],
  currentVersionMeta: null,
  workingVersionParentId: null,
  versionsLoadedAt: null,
};

/* ─── FORMATTERS ─────────────────────────────────────────────────── */
const currencyFormatter = new Intl.NumberFormat("pt-BR", {
  style:"currency", currency:"BRL", maximumFractionDigits:0,
});
const numberFormatter = new Intl.NumberFormat("pt-BR", { maximumFractionDigits:0 });
const decimalFormatter = new Intl.NumberFormat("pt-BR", {
  minimumFractionDigits:2, maximumFractionDigits:2,
});
const shortDateFormatter = new Intl.DateTimeFormat("pt-BR", {
  day:"2-digit", month:"2-digit", year:"2-digit",
});
const fullDateFormatter = new Intl.DateTimeFormat("pt-BR", {
  day:"2-digit", month:"2-digit", year:"numeric",
});
const monthFormatter = new Intl.DateTimeFormat("pt-BR", {
  month:"short", year:"2-digit",
});

const byId = (id) => document.getElementById(id);

function cloneJson(value) {
  return value == null ? value : JSON.parse(JSON.stringify(value));
}

/* ─── CFOP REGISTRY ──────────────────────────────────────────────── */
const CFOP_TYPE_OPTIONS = [
  { value:"venda",         label:"Venda" },
  { value:"devolucao",     label:"Devolução" },
  { value:"remessa",       label:"Remessa / transferência" },
  { value:"bonificacao",   label:"Bonificação" },
  { value:"doacao",        label:"Doação" },
  { value:"brinde",        label:"Brinde" },
  { value:"baixa_estoque", label:"Baixa de estoque" },
  { value:"ignorar",       label:"Ignorar na análise" },
];

const CFOP_REGISTRY_SEED = {
  "5102":{ description:"Venda de mercadoria adquirida ou recebida de terceiros.", officialType:"venda", source:"Ajuste SINIEF 03/24" },
  "5202":{ description:"Devolução de compra para comercialização.", officialType:"devolucao", source:"Ajuste SINIEF 03/24" },
  "5405":{ description:"Venda com substituição tributária — contribuinte substituído.", officialType:"venda", source:"Ajuste SINIEF 03/24" },
  "5411":{ description:"Devolução de compra com substituição tributária.", officialType:"devolucao", source:"Ajuste SINIEF 03/24" },
  "5905":{ description:"Remessa para depósito fechado, armazém geral ou filial.", officialType:"remessa", source:"Ajuste SINIEF 03/24" },
  "5910":{ description:"Remessa em bonificação, doação ou brinde.", officialType:"bonificacao", source:"Ajuste SINIEF 03/24" },
  "5911":{ description:"Remessa de amostra grátis.", officialType:"remessa", source:"Ajuste SINIEF 03/24" },
  "5927":{ description:"Baixa de estoque — perda, roubo ou deterioração.", officialType:"baixa_estoque", source:"Ajuste SINIEF 03/24" },
  "5949":{ description:"Outra saída não especificada.", officialType:"ignorar", source:"Ajuste SINIEF 03/24" },
  "6102":{ description:"Venda interestadual de mercadoria.", officialType:"venda", source:"Ajuste SINIEF 03/24" },
  "6152":{ description:"Transferência interestadual de mercadoria.", officialType:"remessa", source:"Ajuste SINIEF 03/24" },
  "6202":{ description:"Devolução interestadual de compra.", officialType:"devolucao", source:"Ajuste SINIEF 03/24" },
  "6910":{ description:"Remessa interestadual em bonificação, doação ou brinde.", officialType:"bonificacao", source:"Ajuste SINIEF 03/24" },
  "6911":{ description:"Remessa interestadual de amostra grátis.", officialType:"remessa", source:"Ajuste SINIEF 03/24" },
};

/* ─── UTILS ──────────────────────────────────────────────────────── */
function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&","&amp;").replaceAll("<","&lt;")
    .replaceAll(">","&gt;").replaceAll('"',"&quot;")
    .replaceAll("'","&#39;");
}

function formatCurrency(value)  { return currencyFormatter.format(Number(value || 0)); }
function formatNumber(value)    { return numberFormatter.format(Number(value || 0)); }
function formatDecimal(value)   { return decimalFormatter.format(Number(value || 0)); }

function formatPercent(value) {
  if (value == null || Number.isNaN(value)) return "—";
  return `${formatDecimal(value)}%`;
}

function parseIsoDate(isoDate) {
  if (!isoDate) return null;
  const [year, month, day] = String(isoDate).split("-").map(Number);
  if (!year || !month || !day) return null;
  return new Date(year, month - 1, day);
}

function formatDateBr(isoDate, { shortYear = false } = {}) {
  if (!isoDate) return "—";
  const date = parseIsoDate(isoDate);
  if (!date) return String(isoDate);
  return shortYear ? shortDateFormatter.format(date) : fullDateFormatter.format(date);
}

function formatRangeBr(from, to, { shortYear = true } = {}) {
  if (!from && !to) return "Sem período";
  if (!from) return `Até ${formatDateBr(to, { shortYear })}`;
  if (!to)   return `A partir de ${formatDateBr(from, { shortYear })}`;
  return `${formatDateBr(from, { shortYear })} – ${formatDateBr(to, { shortYear })}`;
}

function formatMonth(isoMonth) {
  if (!isoMonth) return "";
  const [year, month] = isoMonth.split("-");
  const date = new Date(Number(year), Number(month) - 1, 1);
  return monthFormatter.format(date).replace(".", "");
}

function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g,"")
    .toLowerCase().replace(/[^a-z0-9]+/g,"_")
    .replace(/^_+|_+$/g,"");
}

function parseNumber(value, { allowNull = false } = {}) {
  if (value == null || value === "") return allowNull ? null : 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : (allowNull ? null : 0);
  const trimmed = String(value).trim();
  if (!trimmed) return allowNull ? null : 0;
  const normalized = trimmed.replace(/\s/g,"")
    .replace(/\.(?=\d{3}(?:\D|$))/g,"").replace(",",".");
  const parsed = Number(normalized);
  if (!Number.isFinite(parsed)) return allowNull ? null : 0;
  return parsed;
}

function toIsoDate(year, month, day) {
  return [String(year).padStart(4,"0"), String(month).padStart(2,"0"), String(day).padStart(2,"0")].join("-");
}

function addDaysToIso(isoDate, days) {
  const date = parseIsoDate(isoDate);
  if (!date) return "";
  date.setDate(date.getDate() + days);
  return toIsoDate(date.getFullYear(), date.getMonth() + 1, date.getDate());
}

function diffDaysInclusive(from, to) {
  const start = parseIsoDate(from), end = parseIsoDate(to);
  if (!start || !end) return 0;
  return Math.round((end - start) / 86400000) + 1;
}

function parseDate(value) {
  if (value == null || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime()))
    return toIsoDate(value.getFullYear(), value.getMonth() + 1, value.getDate());
  if (typeof value === "number" && Number.isFinite(value) && value > 20000 && value < 70000) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return toIsoDate(parsed.y, parsed.m, parsed.d);
  }
  const text = String(value).trim();
  if (!text) return null;
  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) return toIsoDate(isoMatch[1], isoMatch[2], isoMatch[3]);
  const brMatch = text.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (brMatch) return toIsoDate(brMatch[3], brMatch[2], brMatch[1]);
  const fallback = new Date(text);
  if (!Number.isNaN(fallback.getTime()))
    return toIsoDate(fallback.getFullYear(), fallback.getMonth() + 1, fallback.getDate());
  return null;
}

function monthFromDate(value) { return value ? value.slice(0, 7) : ""; }

function mergeConfig(config) {
  const base = createDefaultConfig();
  return {
    revenue: { ...base.revenue, ...(config?.revenue || {}) },
    taxes: { ...base.taxes, ...(config?.taxes || {}) },
    stock: { ...base.stock, ...(config?.stock || {}) },
    margin: { ...base.margin, ...(config?.margin || {}) },
  };
}

function normalizeCfopOverrides(rawOverrides) {
  if (!rawOverrides) return [];
  if (Array.isArray(rawOverrides)) {
    return rawOverrides
      .map((item) => ({
        code: String(item.code ?? item.cfop ?? "").replace(/\D/g, ""),
        analysisType: String(item.analysisType ?? item.type ?? item.customType ?? "").trim(),
      }))
      .filter((item) => item.code && item.analysisType);
  }
  if (typeof rawOverrides === "object") {
    return Object.entries(rawOverrides)
      .map(([code, analysisType]) => ({
        code: String(code).replace(/\D/g, ""),
        analysisType: String(analysisType ?? "").trim(),
      }))
      .filter((item) => item.code && item.analysisType);
  }
  return [];
}

function getCfopOverrides() {
  return [...state.cfopRegistry.values()]
    .filter((meta) => meta.customType)
    .sort((a, b) => a.code.localeCompare(b.code))
    .map((meta) => ({
      code: meta.code,
      analysisType: meta.analysisType,
      officialType: meta.officialType,
    }));
}

function setConfigInputs(config) {
  const merged = mergeConfig(config);
  byId("configRevenueSales").checked = !!merged.revenue.sales;
  byId("configRevenueReturns").checked = !!merged.revenue.returns;
  byId("configTaxIcms").checked = !!merged.taxes.icms;
  byId("configTaxPis").checked = !!merged.taxes.pis;
  byId("configTaxCofins").checked = !!merged.taxes.cofins;
  byId("configTaxIpi").checked = !!merged.taxes.ipi;
  byId("configStockRemessa").checked = !!merged.stock.remessa;
  byId("configStockBaixa").checked = !!merged.stock.baixa;
  byId("configMarginDeductTaxes").checked = !!merged.margin.deductTaxes;
  byId("configMarginRequireCost").checked = !!merged.margin.requireCost;
}

function applyConfigState(config, { rebuild = false, rerender = false } = {}) {
  setConfigInputs(config);
  state.config = mergeConfig(config);
  syncConfigFromInputs();
  if (rebuild && state.baseRows.length) rebuildRowsFromCfopRegistry();
  if (rerender && state.source) applyFilters();
}

function applyCfopOverrides(rawOverrides, { resetMissing = true, rebuild = true, rerender = true } = {}) {
  const overrides = normalizeCfopOverrides(rawOverrides);
  const map = new Map(overrides.map((item) => [item.code, item.analysisType]));

  state.cfopRegistry.forEach((meta, code) => {
    if (map.has(code)) {
      meta.customType = map.get(code) === meta.officialType ? null : map.get(code);
      meta.analysisType = meta.customType || meta.officialType;
      return;
    }
    if (resetMissing) {
      meta.customType = null;
      meta.analysisType = meta.officialType;
    }
  });

  if (rebuild && state.baseRows.length) rebuildRowsFromCfopRegistry();
  renderCfopRegistryMeta();
  renderCfopConfigTable();
  if (rerender && state.source) applyFilters();
}

function formatTimestamp(value) {
  if (!value) return "—";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return new Intl.DateTimeFormat("pt-BR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
}

/* ─── CFOP CLASSIFICATION ────────────────────────────────────────── */
/* FIX: fallback antigo (/^[56]/ → "venda") era perigoso.
   Agora CFOPs desconhecidos viram "ignorar" em vez de inflar receita. */
function classifyType(typeValue, cfopValue) {
  const type = normalizeText(typeValue);
  const cfop = String(cfopValue ?? "").trim();

  // Prefer explicit type label from the row
  if (type.includes("devol"))                       return "devolucao";
  if (type.includes("bonific"))                     return "bonificacao";
  if (type.includes("doac"))                        return "doacao";
  if (type.includes("brinde"))                      return "brinde";
  if (type.includes("remessa") || type.includes("transfer")) return "remessa";
  if (type.includes("baixa") && type.includes("estoque"))    return "baixa_estoque";
  if (type.includes("venda"))                       return "venda";

  // Known CFOP fallback — explicit safe list only
  if (/^[56]202$/.test(cfop))   return "devolucao";
  if (/^[56]411$/.test(cfop))   return "devolucao";
  if (/^[56]910$/.test(cfop))   return "bonificacao";
  if (/^[56]911$/.test(cfop))   return "remessa";
  if (/^[56]927$/.test(cfop))   return "baixa_estoque";
  if (/^[56]905$/.test(cfop))   return "remessa";
  if (/^[56]152$/.test(cfop))   return "remessa";
  if (/^[56]949$/.test(cfop))   return "ignorar";
  // Known venda CFOPs only
  if (/^[56]102$/.test(cfop))   return "venda";
  if (/^[56]405$/.test(cfop))   return "venda";

  // SAFE default: ignorar (never silently inflate revenue)
  return "ignorar";
}

function friendlyType(type) {
  const map = {
    venda:"Venda", devolucao:"Devolução", remessa:"Remessa",
    bonificacao:"Bonificação", doacao:"Doação", brinde:"Brinde",
    baixa_estoque:"Baixa de estoque", ignorar:"Ignorar",
  };
  return map[type] || "Outro";
}

function isRemessaLikeType(type) {
  return ["remessa","bonificacao","doacao","brinde"].includes(type);
}
function isStockMovementType(type) {
  return isRemessaLikeType(type) || type === "baixa_estoque";
}
function inferTypeFromCfop(cfop, fallbackType = "ignorar") {
  const code = String(cfop ?? "").replace(/\D/g,"");
  if (!code) return fallbackType;
  if (CFOP_REGISTRY_SEED[code]?.officialType) return CFOP_REGISTRY_SEED[code].officialType;
  return classifyType(fallbackType, code);
}
function getCfopMeta(cfop) {
  return state.cfopRegistry.get(String(cfop ?? "").replace(/\D/g,"")) || null;
}
function getCfopAnalysisType(cfop, fallbackType = "ignorar") {
  const meta = getCfopMeta(cfop);
  if (meta?.analysisType) return meta.analysisType;
  return inferTypeFromCfop(cfop, fallbackType);
}

/* ─── MARGIN HELPERS ─────────────────────────────────────────────── */
function calculateRowTaxes(row) {
  let total = 0;
  if (state.config.taxes.icms)   total += row.icms   || 0;
  if (state.config.taxes.pis)    total += row.pis    || 0;
  if (state.config.taxes.cofins) total += row.cofins || 0;
  if (state.config.taxes.ipi)    total += row.ipi    || 0;
  return total;
}

/* Returns { value, pct } or { value:null, pct:null } — computed once per call */
function calculateMarginForRow(row) {
  if (state.config.margin.requireCost && row.cost == null)
    return { value:null, pct:null };
  if (row.cost == null)
    return { value:null, pct:null };
  let value = row.revenue - row.cost;
  if (state.config.margin.deductTaxes) value -= calculateRowTaxes(row);
  const pct = row.revenue ? (value / row.revenue) * 100 : null;
  return { value, pct };
}

/* ─── ROW HELPERS ────────────────────────────────────────────────── */
function createHeaderMap(rawRow) {
  const map = new Map();
  Object.keys(rawRow || {}).forEach((key) => map.set(normalizeText(key), rawRow[key]));
  return map;
}

function pickValue(headerMap, keys) {
  for (const key of keys) {
    if (headerMap.has(key)) {
      const value = headerMap.get(key);
      if (value !== null && value !== "") return value;
    }
  }
  return null;
}

function ensureSignedValue(value, type) {
  const parsed = parseNumber(value, { allowNull:true });
  if (parsed == null) return null;
  if (type === "devolucao") return parsed > 0 ? -parsed : parsed;
  if (type === "venda")     return parsed < 0 ? Math.abs(parsed) : parsed;
  return parsed < 0 ? Math.abs(parsed) : parsed;
}

/* Build row from Excel raw row */
function buildRow(rawRow) {
  const headerMap = createHeaderMap(rawRow);
  const date   = parseDate(pickValue(headerMap, ["data","data_nf","emissao","data_emissao"]));
  const note   = pickValue(headerMap, ["nota"]);
  const client = pickValue(headerMap, ["cliente"]);
  const uf     = pickValue(headerMap, ["uf"]);
  const item   = pickValue(headerMap, ["nome","item","produto"]);
  const quantity = parseNumber(pickValue(headerMap, ["quant","quantidade"]));
  const revenue  = parseNumber(pickValue(headerMap, ["valor"]));
  const icms     = parseNumber(pickValue(headerMap, ["icms"]));
  const pis      = parseNumber(pickValue(headerMap, ["pis"]));
  const cofins   = parseNumber(pickValue(headerMap, ["cofins"]));
  const ipi      = parseNumber(pickValue(headerMap, ["ipi"]));
  const cost     = parseNumber(pickValue(headerMap, ["custo_total_y","custo_total","custo"]), { allowNull:true });
  const cfop     = pickValue(headerMap, ["cfop"]);
  const typeRaw  = pickValue(headerMap, ["tipo_cfop","tipo","natureza"]) ?? "";
  const type     = classifyType(typeRaw, cfop);

  // Missing required fields → discard (real data quality issue, not CFOP issue)
  if (!date || !note || !client || !item || !revenue) return null;

  if (type === "venda" || type === "devolucao") {
    const sign  = type === "devolucao" ? -1 : 1;
    const sRev  = Math.abs(revenue)  * sign;
    const sQty  = Math.abs(quantity) * sign;
    const sIcms = Math.abs(icms)     * sign;
    const sPis  = Math.abs(pis)      * sign;
    const sCof  = Math.abs(cofins)   * sign;
    const sIpi  = Math.abs(ipi)      * sign;
    const sCost = cost == null ? null : Math.abs(cost) * sign;
    const mVal  = sCost == null ? null : sRev - sCost;
    const mPct  = mVal == null || sRev === 0 ? null : (mVal / sRev) * 100;
    return {
      dataset:"sales",
      row:{
        date, month:monthFromDate(date),
        note:String(note), client:String(client), uf:String(uf ?? ""),
        item:String(item), quantity:sQty,
        revenue:sRev, icms:sIcms, pis:sPis, cofins:sCof, ipi:sIpi,
        taxesTotal:sIcms + sPis + sCof + sIpi,
        cost:sCost, marginValue:mVal, marginPct:mPct,
        cfop:String(cfop ?? ""), type, typeLabel:friendlyType(type),
      },
    };
  }

  if (isStockMovementType(type)) {
    return {
      dataset:"stock",
      row:{
        date, month:monthFromDate(date),
        note:String(note), client:String(client), uf:String(uf ?? ""),
        item:String(item), quantity:Math.abs(quantity),
        value:Math.abs(revenue),
        cfop:String(cfop ?? ""), type, typeLabel:friendlyType(type),
      },
    };
  }

  // FIX: type === "ignorar" → preserve row in base for CFOP registry visibility.
  // The row will NOT enter sales or stock but the CFOP will appear in the configurator
  // so the analyst can reclassify it. We return a special "unclassified" marker.
  return {
    dataset:"unclassified",
    row:{
      date, month:monthFromDate(date),
      note:String(note), client:String(client), uf:String(uf ?? ""),
      item:String(item), quantity:Math.abs(quantity),
      revenue:Math.abs(revenue),
      cfop:String(cfop ?? ""), type:"ignorar", typeLabel:"Não classificado",
    },
  };
}

function normalizeJsonSalesRow(rawRow) {
  const type = classifyType(rawRow.type ?? rawRow.type_label ?? rawRow.tipo_cfop, rawRow.cfop);
  if (type !== "venda" && type !== "devolucao") return null;
  const date   = parseDate(rawRow.date);
  const note   = rawRow.note   ?? rawRow.nota;
  const client = rawRow.client ?? rawRow.cliente;
  const item   = rawRow.item   ?? rawRow.nome;
  if (!date || !note || !client || !item) return null;
  const revenue = ensureSignedValue(rawRow.revenue ?? rawRow.valor, type);
  const quantity = ensureSignedValue(rawRow.quantity ?? rawRow.quant, type) ?? 0;
  const icms     = ensureSignedValue(rawRow.icms, type) ?? 0;
  const pis      = ensureSignedValue(rawRow.pis, type) ?? 0;
  const cofins   = ensureSignedValue(rawRow.cofins, type) ?? 0;
  const ipi      = ensureSignedValue(rawRow.ipi, type) ?? 0;
  const cost     = ensureSignedValue(rawRow.cost ?? rawRow.custo, type);
  const marginValue = rawRow.margin_value != null
    ? ensureSignedValue(rawRow.margin_value, type)
    : cost == null || revenue == null ? null : revenue - cost;
  const marginPct = rawRow.margin_pct != null
    ? parseNumber(rawRow.margin_pct, { allowNull:true })
    : marginValue == null || !revenue ? null : (marginValue / revenue) * 100;
  if (revenue == null) return null;
  return {
    date, month:monthFromDate(date),
    note:String(note), client:String(client), uf:String(rawRow.uf ?? ""),
    item:String(item), quantity, revenue, icms, pis, cofins, ipi,
    taxesTotal:icms + pis + cofins + ipi,
    cost, marginValue, marginPct,
    cfop:String(rawRow.cfop ?? ""), type, typeLabel:friendlyType(type),
  };
}

function normalizeJsonStockRow(rawRow) {
  const type = classifyType(rawRow.type ?? rawRow.type_label ?? rawRow.tipo_cfop, rawRow.cfop);
  if (!isStockMovementType(type)) return null;
  const date   = parseDate(rawRow.date);
  const note   = rawRow.note   ?? rawRow.nota;
  const client = rawRow.client ?? rawRow.cliente;
  const item   = rawRow.item   ?? rawRow.nome;
  const value  = parseNumber(rawRow.value ?? rawRow.valor, { allowNull:true });
  if (!date || !note || !client || !item || value == null) return null;
  return {
    date, month:monthFromDate(date),
    note:String(note), client:String(client), uf:String(rawRow.uf ?? ""),
    item:String(item), quantity:Math.abs(parseNumber(rawRow.quantity ?? rawRow.quant)),
    value:Math.abs(value), cfop:String(rawRow.cfop ?? ""), type, typeLabel:friendlyType(type),
  };
}

/* ─── CFOP REGISTRY ──────────────────────────────────────────────── */
function buildCfopRegistry(baseRows) {
  const registry = new Map();
  baseRows.forEach((row) => {
    const code = String(row.cfop ?? "").replace(/\D/g,"");
    if (!code) return;
    const seeded = CFOP_REGISTRY_SEED[code] || {};
    if (!registry.has(code)) {
      registry.set(code, {
        code,
        description:  seeded.description || "CFOP sem descrição oficial semeada.",
        officialType: seeded.officialType || inferTypeFromCfop(code, row.originalType),
        customType:   null,
        analysisType: seeded.officialType || inferTypeFromCfop(code, row.originalType),
        count:0, totalValue:0,
        source: seeded.source || "Heurística operacional",
      });
    }
    const current = registry.get(code);
    current.count      += 1;
    current.totalValue += Math.abs(row.baseValue || row.revenue || row.value || 0);
  });

  // Preserve custom overrides
  state.cfopRegistry.forEach((existing, code) => {
    if (!registry.has(code) || !existing.customType) return;
    const target = registry.get(code);
    target.customType   = existing.customType;
    target.analysisType = existing.customType;
  });

  state.cfopRegistry = registry;
}

function toBaseRow(row, dataset) {
  if (dataset === "sales") {
    return { ...row, dataset, baseValue:Math.abs(row.revenue || 0), originalType:row.type, analysisType:row.type };
  }
  return {
    date:row.date, month:row.month, note:row.note, client:row.client, uf:row.uf,
    item:row.item, quantity:row.quantity, revenue:row.value,
    icms:0, pis:0, cofins:0, ipi:0, taxesTotal:0,
    cost:null, marginValue:null, marginPct:null,
    cfop:row.cfop, type:row.type, typeLabel:row.typeLabel,
    dataset, value:row.value, baseValue:row.value, originalType:row.type, analysisType:row.type,
  };
}

function rebuildRowsFromCfopRegistry() {
  const salesRows = [], stockRows = [];
  state.baseRows.forEach((row) => {
    const analysisType  = getCfopAnalysisType(row.cfop, row.originalType);
    row.analysisType    = analysisType;
    row.type            = analysisType;
    row.typeLabel       = friendlyType(analysisType);

    if (analysisType === "venda" || analysisType === "devolucao") {
      const sign       = analysisType === "devolucao" ? -1 : 1;
      const revenue    = Math.abs(row.baseValue || row.revenue || 0) * sign;
      const quantity   = Math.abs(row.quantity || 0) * sign;
      const taxScale   = row.revenue ? Math.abs(revenue / row.revenue) : 1;
      const cost       = row.cost == null ? null : Math.abs(row.cost) * sign;
      const sIcms      = Math.abs(row.icms   || 0) * sign * taxScale;
      const sPis       = Math.abs(row.pis    || 0) * sign * taxScale;
      const sCofins    = Math.abs(row.cofins || 0) * sign * taxScale;
      const sIpi       = Math.abs(row.ipi    || 0) * sign * taxScale;
      salesRows.push({
        date:row.date, month:row.month, note:row.note, client:row.client, uf:row.uf,
        item:row.item, quantity, revenue,
        icms:sIcms, pis:sPis, cofins:sCofins, ipi:sIpi,
        taxesTotal:sIcms + sPis + sCofins + sIpi,   // FIX
        cost, marginValue:null, marginPct:null,
        cfop:String(row.cfop || ""), type:analysisType, typeLabel:friendlyType(analysisType),
      });
      return;
    }

    if (isStockMovementType(analysisType)) {
      stockRows.push({
        date:row.date, month:row.month, note:row.note, client:row.client, uf:row.uf,
        item:row.item, quantity:Math.abs(row.quantity || 0),
        value:Math.abs(row.baseValue || row.value || row.revenue || 0),
        cfop:String(row.cfop || ""), type:analysisType, typeLabel:friendlyType(analysisType),
      });
    }
  });

  state.salesRows = salesRows.sort((a,b) => a.date.localeCompare(b.date) || a.note.localeCompare(b.note));
  state.stockRows = stockRows.sort((a,b) => a.date.localeCompare(b.date));
  if (state.source) {
    state.source.salesRows = state.salesRows;
    state.source.stockRows = state.stockRows;
  }
  invalidateMetricsCache(); // rebuild always invalidates
}

/* ─── FINANCIAL HELPERS ──────────────────────────────────────────── */
function sum(rows, field) { return rows.reduce((t,r) => t + Number(r[field] || 0), 0); }
function uniqueCount(rows, field) { return new Set(rows.map(r => r[field]).filter(Boolean)).size; }

function sortByMetric(rows, metric, direction = "desc") {
  return [...rows].sort((a,b) => {
    const av = a[metric] == null ? -Infinity : Number(a[metric]);
    const bv = b[metric] == null ? -Infinity : Number(b[metric]);
    return direction === "asc" ? av - bv : bv - av;
  });
}

function isRevenueRowIncluded(row) {
  if (row.type === "venda")     return state.config.revenue.sales;
  if (row.type === "devolucao") return state.config.revenue.returns;
  return false;
}
function isStockRowSelected(row) {
  if (isRemessaLikeType(row.type)) return state.config.stock.remessa;
  if (row.type === "baixa_estoque") return state.config.stock.baixa;
  return false;
}

function getSelectedTaxFields() {
  return ["icms","pis","cofins","ipi"].filter(f => state.config.taxes[f]);
}
function getSelectedTaxLabel() {
  const labels = getSelectedTaxFields().map(f => f.toUpperCase());
  return labels.length ? labels.join(" + ") : "Nenhum tributo";
}
function getRevenueCompositionLabel() {
  const labels = [];
  if (state.config.revenue.sales)   labels.push("vendas");
  if (state.config.revenue.returns) labels.push("devoluções");
  return labels.length ? labels.join(" + ") : "nenhum componente";
}

function percentageChange(current, previous) {
  if (previous == null || Number(previous) === 0 || !Number.isFinite(Number(previous))) return null;
  return ((current - previous) / Math.abs(previous)) * 100;
}

function buildDeltaChip(current, previous, { inverse = false, label = "vs período B" } = {}) {
  const delta = percentageChange(current, previous);
  if (delta == null) return { text:`Sem base ${label}`, tone:"neutral" };
  const positiveIsGood = inverse ? delta <= 0 : delta >= 0;
  return {
    text:`${delta >= 0 ? "+" : ""}${formatDecimal(delta)}% ${label}`,
    tone:positiveIsGood ? "good" : "bad",
  };
}

/* ─── METRICS ENGINE ─────────────────────────────────────────────── */
function calculateMetrics(salesRows, stockRows) {
  const revenueRows      = salesRows.filter(isRevenueRowIncluded);
  const positiveSales    = salesRows.filter(r => r.type === "venda");
  const returnRows       = salesRows.filter(r => r.type === "devolucao");
  const selectedStock    = stockRows.filter(isStockRowSelected);
  const remessaRows      = stockRows.filter(r => isRemessaLikeType(r.type));
  const baixaRows        = stockRows.filter(r => r.type === "baixa_estoque");

  // Monthly aggregation
  const monthlyMap = new Map();
  revenueRows.forEach((row) => {
    const { value:mVal } = calculateMarginForRow(row);
    if (!monthlyMap.has(row.month)) {
      monthlyMap.set(row.month, {
        month:row.month, revenue:0, taxes:0, returns:0,
        mappedRevenue:0, marginValue:0,
        icms:0, pis:0, cofins:0, ipi:0, notes:new Set(),
      });
    }
    const m = monthlyMap.get(row.month);
    m.revenue += row.revenue;
    m.taxes   += calculateRowTaxes(row);
    m.icms    += row.icms   || 0;
    m.pis     += row.pis    || 0;
    m.cofins  += row.cofins || 0;
    m.ipi     += row.ipi    || 0;
    m.notes.add(row.note);
    if (row.type === "devolucao") m.returns += Math.abs(row.revenue);
    if (mVal != null) { m.marginValue += mVal; m.mappedRevenue += row.revenue; }
  });

  const monthly = [...monthlyMap.values()]
    .sort((a,b) => a.month.localeCompare(b.month))
    .map(m => ({
      month:m.month, revenue:m.revenue, taxes:m.taxes, returns:m.returns,
      noteCount:m.notes.size, marginValue:m.marginValue,
      icms:m.icms, pis:m.pis, cofins:m.cofins, ipi:m.ipi,
      marginPct: m.mappedRevenue ? (m.marginValue / m.mappedRevenue) * 100 : null,
      taxPct:    m.revenue       ? (m.taxes / m.revenue) * 100 : null,
    }));

  // Totals
  const totalRevenue         = sum(revenueRows, "revenue");
  const grossPositiveRevenue = sum(positiveSales, "revenue");
  const totalReturns         = Math.abs(sum(returnRows, "revenue"));
  const totalTaxes           = revenueRows.reduce((t,r) => t + calculateRowTaxes(r), 0);

  // FIX: compute margin once per row, store result, avoid double-call
  const rowMargins = revenueRows.map(r => ({ r, m: calculateMarginForRow(r) }));
  const mappedPairs   = rowMargins.filter(({ m }) => m.value != null);
  const mappedRevenue = mappedPairs.reduce((t, { r }) => t + r.revenue, 0);
  const totalMargin   = mappedPairs.reduce((t, { m }) => t + m.value, 0);

  const grossMarginPct = mappedRevenue ? (totalMargin / mappedRevenue) * 100 : null;
  const taxPct         = totalRevenue  ? (totalTaxes / totalRevenue) * 100 : null;
  const returnRate     = grossPositiveRevenue ? (totalReturns / grossPositiveRevenue) * 100 : null;

  // Total notes for ticket calculation
  const totalNotesCount = uniqueCount(positiveSales, "note");
  const avgTicket = totalNotesCount ? grossPositiveRevenue / totalNotesCount : null;

  // Client aggregation
  const clientMap = new Map();
  revenueRows.forEach((row) => {
    const { value:mVal } = calculateMarginForRow(row);
    if (!clientMap.has(row.client)) {
      clientMap.set(row.client, {
        client:row.client, uf:row.uf || "", revenue:0, quantity:0,
        taxes:0, returns:0, marginValue:0, mappedRevenue:0, notes:new Set(),
      });
    }
    const c = clientMap.get(row.client);
    c.revenue  += row.revenue;
    c.quantity += row.quantity || 0;
    c.taxes    += calculateRowTaxes(row);
    if (row.type === "devolucao") c.returns += Math.abs(row.revenue);
    if (mVal != null) { c.marginValue += mVal; c.mappedRevenue += row.revenue; }
    c.notes.add(row.note);
  });

  const clients = [...clientMap.values()]
    .map(c => ({
      client:c.client, uf:c.uf, revenue:c.revenue, quantity:c.quantity,
      taxes:c.taxes, returns:c.returns, notes:c.notes.size,
      marginValue:c.marginValue,
      marginPct:c.mappedRevenue ? (c.marginValue / c.mappedRevenue) * 100 : null,
    }))
    .sort((a,b) => b.revenue - a.revenue);

  // Client concentration (Herfindahl-proxy: top-3 share)
  const top3Revenue = clients.slice(0,3).reduce((t,c) => t + c.revenue, 0);
  const top3Share   = totalRevenue ? (top3Revenue / totalRevenue) * 100 : null;

  // Product aggregation
  const productMap = new Map();
  revenueRows.forEach((row) => {
    const { value:mVal } = calculateMarginForRow(row);
    if (!productMap.has(row.item)) {
      productMap.set(row.item, {
        item:row.item, revenue:0, quantity:0, taxes:0, returns:0,
        marginValue:0, mappedRevenue:0, notes:new Set(),
      });
    }
    const p = productMap.get(row.item);
    p.revenue  += row.revenue;
    p.quantity += row.quantity || 0;
    p.taxes    += calculateRowTaxes(row);
    if (row.type === "devolucao") p.returns += Math.abs(row.revenue);
    if (mVal != null) { p.marginValue += mVal; p.mappedRevenue += row.revenue; }
    p.notes.add(row.note);
  });

  const products = [...productMap.values()]
    .map(p => ({
      item:p.item, revenue:p.revenue, quantity:p.quantity,
      taxes:p.taxes, returns:p.returns, notes:p.notes.size,
      marginValue:p.marginValue,
      marginPct:p.mappedRevenue ? (p.marginValue / p.mappedRevenue) * 100 : null,
    }))
    .sort((a,b) => b.revenue - a.revenue);

  const positiveSkuMargin = products.filter(p => p.marginValue > 0).sort((a,b) => b.marginValue - a.marginValue);
  const negativeSkuMargin = products.filter(p => p.marginValue < 0).sort((a,b) => a.marginValue - b.marginValue);
  const badClients = clients.filter(c => c.marginValue < 0 || (c.marginPct != null && c.marginPct < 0))
    .sort((a,b) => a.marginValue - b.marginValue);

  // Tax outliers (z-score)
  const taxProducts = products
    .filter(p => p.revenue > 5000)
    .map(p => ({ item:p.item, revenue:p.revenue, taxPct:p.revenue ? (p.taxes/p.revenue)*100 : null }))
    .filter(p => p.taxPct != null && Number.isFinite(p.taxPct));
  const avgTaxPct = taxProducts.length ? taxProducts.reduce((t,p) => t+p.taxPct,0)/taxProducts.length : null;
  const taxStd = taxProducts.length && avgTaxPct != null
    ? Math.sqrt(taxProducts.reduce((t,p) => t + Math.pow(p.taxPct-avgTaxPct,2), 0) / taxProducts.length)
    : null;
  const taxOutliers = taxStd
    ? taxProducts.map(p => ({ ...p, zScore:(p.taxPct-avgTaxPct)/taxStd }))
        .filter(p => Math.abs(p.zScore) >= 1.5)
        .sort((a,b) => Math.abs(b.zScore)-Math.abs(a.zScore))
    : [];

  // Stock products helpers
  function buildStockProducts(rows) {
    const map = new Map();
    rows.forEach(r => {
      if (!map.has(r.item)) map.set(r.item,{item:r.item,value:0,quantity:0,notes:new Set()});
      const s = map.get(r.item);
      s.value    += r.value;
      s.quantity += r.quantity || 0;
      s.notes.add(r.note);
    });
    return [...map.values()].map(r=>({...r,notes:r.notes.size})).sort((a,b)=>b.value-a.value);
  }

  function buildStockCfops(rows) {
    const map = new Map();
    rows.forEach(r => {
      const code = String(r.cfop || "");
      const meta = getCfopMeta(code);
      if (!map.has(code)) map.set(code,{cfop:code, description:meta?.description||"CFOP sem descrição", value:0, quantity:0, notes:new Set()});
      const s = map.get(code);
      s.value    += r.value;
      s.quantity += r.quantity || 0;
      s.notes.add(r.note);
    });
    return [...map.values()].map(r=>({...r,notes:r.notes.size})).sort((a,b)=>b.value-a.value);
  }

  // Monthly stock maps
  function toMonthlyMap(rows, field = "value") {
    const m = new Map();
    rows.forEach(r => {
      if (!m.has(r.month)) m.set(r.month,{month:r.month,value:0});
      m.get(r.month).value += r[field];
    });
    return [...m.values()].sort((a,b)=>a.month.localeCompare(b.month));
  }

  const taxBreakdown = {
    icms:   revenueRows.reduce((t,r) => t+(state.config.taxes.icms   ? r.icms   ||0 : 0), 0),
    pis:    revenueRows.reduce((t,r) => t+(state.config.taxes.pis    ? r.pis    ||0 : 0), 0),
    cofins: revenueRows.reduce((t,r) => t+(state.config.taxes.cofins ? r.cofins ||0 : 0), 0),
    ipi:    revenueRows.reduce((t,r) => t+(state.config.taxes.ipi    ? r.ipi    ||0 : 0), 0),
  };

  const lastMonth     = monthly[monthly.length - 1] || null;
  const previousMonth = monthly[monthly.length - 2] || null;
  const monthOverMonthRevenue = lastMonth && previousMonth && previousMonth.revenue
    ? ((lastMonth.revenue - previousMonth.revenue) / previousMonth.revenue) * 100 : null;

  // ── CPV (Custo dos Produtos Vendidos) ─────────────────────────────
  // CPV = soma dos custos das linhas de venda com custo mapeado
  const cpvRows   = positiveSales.filter(r => r.cost != null && r.cost > 0);
  const totalCpv  = cpvRows.reduce((t,r) => t + Math.abs(r.cost), 0);
  // Base de receita sobre a qual o CPV foi calculado (linhas com custo)
  const cpvRevBase = cpvRows.reduce((t,r) => t + r.revenue, 0);
  const cpvPct     = cpvRevBase ? (totalCpv / cpvRevBase) * 100 : null;

  // ── RECEITA BRUTA (apenas vendas positivas, sem devoluções) ───────
  const receitaBruta  = grossPositiveRevenue;
  // Deduções = devoluções + impostos totais (sobre receita bruta)
  const totalDeducoes = totalReturns + totalTaxes;
  const deducoesPct   = receitaBruta ? (totalDeducoes / receitaBruta) * 100 : null;

  // ── RECEITA LÍQUIDA JÁ CALCULADA COMO totalRevenue (vendas – devoluções)
  // Mas para consistência com o DRE: receita líquida após impostos também
  const receitaLiquidaAposImpostos = totalRevenue - totalTaxes;

  // ── TAXA DE CONVERSÃO DE RECEITA ──────────────────────────────────
  // Quanto da receita bruta chega como líquida (após deduções)
  const taxaConversao = receitaBruta ? (totalRevenue / receitaBruta) * 100 : null;

  // ── IMPACTO DAS DEVOLUÇÕES SOBRE RECEITA LÍQUIDA ──────────────────
  const devolucaoSobreLiquida = totalRevenue ? (totalReturns / (totalRevenue + totalReturns)) * 100 : null;

  // ── CARGA TRIBUTÁRIA SOBRE RECEITA LÍQUIDA ────────────────────────
  const cargaSobreLiquida = totalRevenue ? (totalTaxes / totalRevenue) * 100 : null;

  // ── MARKUP ────────────────────────────────────────────────────────
  // Markup = Receita Líquida / CPV (mapeado)
  const markup = totalCpv ? (cpvRevBase / totalCpv) : null;

  // ── MARGEM BRUTA SOBRE RECEITA LÍQUIDA ───────────────────────────
  // Já temos grossMarginPct (sobre mappedRevenue). Calculamos também sobre liquida total
  const margemBrutaLiquida = mappedRevenue ? (totalMargin / mappedRevenue) * 100 : null;

  // ── ÍNDICE DE PERDA TOTAL ─────────────────────────────────────────
  // (Devoluções + Impostos) / Receita Bruta — quanto "morre" antes de virar margem
  const indicePerdaTotal = receitaBruta ? (totalDeducoes / receitaBruta) * 100 : null;

  // ── EFICIÊNCIA COMERCIAL ──────────────────────────────────────────
  // Margem Bruta / Receita Bruta — KPI para diretoria
  const eficienciaComercial = (receitaBruta && grossMarginPct != null)
    ? (totalMargin / receitaBruta) * 100 : null;

  // ── CPV MENSAL (para gráfico de tendência) ─────────────────────────
  const monthlyWithCpv = monthly.map(m => {
    // For monthly CPV, we approximate from the ratio if available
    const mRows = positiveSales.filter(r => r.month === m.month && r.cost != null && r.cost > 0);
    const mCpv  = mRows.reduce((t,r) => t + Math.abs(r.cost), 0);
    const mCpvRev = mRows.reduce((t,r) => t + r.revenue, 0);
    return {
      ...m,
      cpv:    mCpv,
      cpvPct: mCpvRev ? (mCpv / mCpvRev) * 100 : null,
      // gross revenue for the month (sales only, no returns)
      grossRevenue: positiveSales.filter(r => r.month === m.month).reduce((t,r) => t + r.revenue, 0),
      returnAmt:    returnRows.filter(r => r.month === m.month).reduce((t,r) => t + Math.abs(r.revenue), 0),
    };
  });

  // ── ALERTAS AUTOMÁTICOS ───────────────────────────────────────────
  const alerts = [];

  if (returnRate != null && returnRate > 3)
    alerts.push({ level:"critical", msg:`Devolução em ${formatPercent(returnRate)} — acima do limite de 3%` });

  if (cpvPct != null) {
    // Detecta CPV subindo M/M
    const cpvTrend = monthlyWithCpv.filter(m => m.cpvPct != null);
    if (cpvTrend.length >= 2) {
      const last2 = cpvTrend.slice(-2);
      if (last2[1].cpvPct > last2[0].cpvPct + 2)
        alerts.push({ level:"warning", msg:`CPV subindo: ${formatPercent(last2[0].cpvPct)} → ${formatPercent(last2[1].cpvPct)} no último mês` });
    }
  }

  if (taxOutliers.length > 0)
    alerts.push({ level:"warning", msg:`${taxOutliers.length} SKU${taxOutliers.length>1?"s":""} com imposto atípico — revisar CFOP/NCM` });

  if (grossMarginPct != null && monthOverMonthRevenue != null && monthOverMonthRevenue > 5 && lastMonth?.marginPct != null) {
    const prevMargPct = previousMonth?.marginPct ?? null;
    if (prevMargPct != null && lastMonth.marginPct < prevMargPct - 2)
      alerts.push({ level:"critical", msg:`Margem caindo (${formatPercent(lastMonth.marginPct)}) com receita subindo (+${formatDecimal(monthOverMonthRevenue)}%) — crescimento ruim` });
  }

  if (top3Share != null && top3Share > 50)
    alerts.push({ level:"warning", msg:`Concentração elevada: top-3 clientes = ${formatPercent(top3Share)} da receita` });

  return {
    // ── receita
    totalRevenue, grossPositiveRevenue, receitaBruta,
    totalReturns, returnRate, devolucaoSobreLiquida,
    totalTaxes, taxBreakdown, taxPct, cargaSobreLiquida,
    // ── deduções
    totalDeducoes, deducoesPct, taxaConversao,
    receitaLiquidaAposImpostos,
    // ── custo
    totalCpv, cpvPct, cpvRevBase,
    // ── margem
    totalMargin, grossMarginPct, margemBrutaLiquida,
    mappedRevenue, markup,
    // ── operacional
    indicePerdaTotal, eficienciaComercial,
    // ── outros
    avgTicket, top3Share, top3Revenue,
    totalClients: uniqueCount(positiveSales, "client"),
    totalSkus:    uniqueCount(positiveSales, "item"),
    totalNotes:   uniqueCount(revenueRows,   "note"),
    totalSelectedStock: sum(selectedStock, "value"),
    totalRemessa:       sum(remessaRows, "value"),
    totalBaixa:         sum(baixaRows, "value"),
    monthly: monthlyWithCpv,
    clients, products, positiveSkuMargin, negativeSkuMargin,
    badClients, taxOutliers, avgTaxPct, alerts,
    remessaMonthly:  toMonthlyMap(remessaRows),
    baixaMonthly:    toMonthlyMap(baixaRows),
    remessaProducts: buildStockProducts(remessaRows),
    baixaProducts:   buildStockProducts(baixaRows),
    baixaCfops:      buildStockCfops(baixaRows),
    lastMonth, previousMonth, monthOverMonthRevenue,
  };
}

/* Cached metrics — avoids full recalc when only chart controls change.
   FIX: cache separado por conjunto de linhas (A vs B).
   A chave inclui primeiro+último date para distinguir recortes com mesmo tamanho. */
const metricsCache = new Map(); // key → metrics

function rowsCacheKey(salesRows, stockRows) {
  const firstSale = salesRows[0]?.date  ?? "";
  const lastSale  = salesRows[salesRows.length - 1]?.date ?? "";
  const firstStk  = stockRows[0]?.date  ?? "";
  const lastStk   = stockRows[stockRows.length - 1]?.date ?? "";
  return `${salesRows.length}:${firstSale}:${lastSale}:${stockRows.length}:${firstStk}:${lastStk}:${JSON.stringify(state.config)}`;
}

function getMetrics(salesRows, stockRows) {
  const key = rowsCacheKey(salesRows, stockRows);
  if (metricsCache.has(key)) return metricsCache.get(key);
  const metrics = calculateMetrics(salesRows, stockRows);
  metricsCache.set(key, metrics);
  return metrics;
}

function invalidateMetricsCache() {
  metricsCache.clear();
}

/* ─── LOADER BAR ─────────────────────────────────────────────────── */
function showLoader() {
  const bar = byId("loaderBar");
  if (!bar) return;
  bar.className = "loader-bar active running";
}
function hideLoader() {
  const bar = byId("loaderBar");
  if (!bar) return;
  bar.className = "loader-bar active";
  setTimeout(() => { bar.className = "loader-bar"; }, 350);
}

/* ─── STATUS ─────────────────────────────────────────────────────── */
function setStatus(message, tone = "info") {
  const bar = byId("statusBar");
  if (!bar) return;
  bar.textContent = message;
  bar.className = `status-inline ${tone}`;
  if (state.statusTimer) clearTimeout(state.statusTimer);
  state.statusTimer = window.setTimeout(() => {
    bar.textContent = "";
    bar.className   = "status-inline";
  }, 6000);
}

/* ─── CHARTS ─────────────────────────────────────────────────────── */
function destroyChart(chartId) {
  if (!state.charts.has(chartId)) return;
  state.charts.get(chartId).destroy();
  state.charts.delete(chartId);
}

function initChartTheme() {
  if (!window.Chart) return;
  Chart.defaults.color              = "#5d6b70";
  Chart.defaults.borderColor        = "rgba(20,32,37,0.08)";
  Chart.defaults.font.family        = '"Manrope",system-ui,sans-serif';
  Chart.defaults.font.size          = 12;
  Chart.defaults.plugins.tooltip.cornerRadius = 6;
}

function renderMonthlyChart(metrics) {
  destroyChart("monthlyChart");
  const hasCpv = metrics.monthly.some(m => m.cpv > 0);
  const datasets = [
    {
      label:"Receita Líquida",
      data:metrics.monthly.map(m => m.revenue),
      backgroundColor:"rgba(91,156,246,0.18)",
      borderColor:"#5b9cf6", borderWidth:1,
      borderRadius:6, grouped:false,
      categoryPercentage:0.72, barPercentage:0.98, maxBarThickness:56,
    },
    {
      label:"Margem Bruta",
      data:metrics.monthly.map(m => m.marginValue),
      backgroundColor(ctx) {
        return ctx.raw >= 0 ? "rgba(46,204,113,0.75)" : "rgba(231,76,60,0.80)";
      },
      borderRadius:6, grouped:false,
      categoryPercentage:0.46, barPercentage:0.96, maxBarThickness:34,
    },
  ];
  if (hasCpv) {
    datasets.push({
      label:"CPV",
      data:metrics.monthly.map(m => m.cpv || 0),
      type:"line",
      borderColor:"#f39c12", borderWidth:2, borderDash:[4,3],
      backgroundColor:"transparent",
      pointBackgroundColor:"#f39c12", pointRadius:4,
      tension:0.3, yAxisID:"y",
    });
  }
  const chart = new Chart(byId("monthlyChart"), {
    type:"bar",
    data:{ labels:metrics.monthly.map(m => formatMonth(m.month)), datasets },
    options:{
      responsive:true, maintainAspectRatio:false,
      onClick(_e, els) {
        if (!els.length) return;
        const month = metrics.monthly[els[0].index].month;
        const rows  = state.filteredSales.filter(r => r.month === month && isRevenueRowIncluded(r));
        openModal(`Lançamentos — ${formatMonth(month)}`, `${formatNumber(rows.length)} linhas`, rows, "sales");
      },
      plugins:{
        legend:{ position:"bottom" },
        tooltip:{ callbacks:{ label:ctx => `${ctx.dataset.label}: ${formatCurrency(ctx.parsed.y)}` } },
      },
      scales:{
        x:{ grid:{display:false} },
        y:{ ticks:{ callback:v => formatCurrency(v) } },
      },
    },
  });
  state.charts.set("monthlyChart", chart);
}

function renderTaxChart(metrics) {
  destroyChart("taxChart");
  const chart = new Chart(byId("taxChart"), {
    type:"bar",
    data:{
      labels:metrics.monthly.map(m => formatMonth(m.month)),
      datasets:[
        {
          label:"Receita", data:metrics.monthly.map(m => m.revenue),
          backgroundColor:"rgba(91,156,246,0.15)", borderColor:"rgba(91,156,246,0.5)",
          borderWidth:1, borderRadius:6, grouped:false,
          categoryPercentage:0.80, barPercentage:0.98, maxBarThickness:58,
        },
        {
          label:"ICMS", data:metrics.monthly.map(m => state.config.taxes.icms ? m.icms : 0),
          backgroundColor:"rgba(243,156,18,0.75)", borderRadius:6, grouped:false,
          categoryPercentage:0.60, barPercentage:0.96, maxBarThickness:44,
          hidden:!state.config.taxes.icms,
        },
        {
          label:"PIS",  data:metrics.monthly.map(m => state.config.taxes.pis ? m.pis : 0),
          backgroundColor:"rgba(46,204,113,0.75)", borderRadius:6, grouped:false,
          categoryPercentage:0.46, barPercentage:0.96, maxBarThickness:32,
          hidden:!state.config.taxes.pis,
        },
        {
          label:"COFINS", data:metrics.monthly.map(m => state.config.taxes.cofins ? m.cofins : 0),
          backgroundColor:"rgba(91,156,246,0.80)", borderRadius:6, grouped:false,
          categoryPercentage:0.32, barPercentage:0.96, maxBarThickness:22,
          hidden:!state.config.taxes.cofins,
        },
        {
          label:"IPI", data:metrics.monthly.map(m => state.config.taxes.ipi ? m.ipi : 0),
          backgroundColor:"rgba(231,76,60,0.80)", borderRadius:6, grouped:false,
          categoryPercentage:0.20, barPercentage:0.96, maxBarThickness:14,
          hidden:!state.config.taxes.ipi,
        },
      ],
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      onClick(_e, els) {
        if (!els.length) return;
        const mMonth = metrics.monthly[els[0].index];
        const rows   = state.filteredSales.filter(r => r.month === mMonth.month && isRevenueRowIncluded(r));
        openModal(`Tributos — ${formatMonth(mMonth.month)}`,
          `Receita ${formatCurrency(mMonth.revenue)} • Tributos ${formatCurrency(mMonth.taxes)}`,
          rows, "sales");
      },
      plugins:{
        legend:{ position:"bottom" },
        tooltip:{ callbacks:{ label:ctx => `${ctx.dataset.label}: ${formatCurrency(ctx.parsed.y)}` } },
      },
      scales:{
        x:{ grid:{display:false} },
        y:{ ticks:{ callback:v => formatCurrency(v) } },
      },
    },
  });
  state.charts.set("taxChart", chart);
  renderTaxBreakdown(metrics);
}

function renderTaxBreakdown(metrics) {
  const totalTaxes = metrics.totalTaxes || 0;
  const items = [
    { label:"ICMS", field:"icms", color:"#f39c12" },
    { label:"PIS",  field:"pis",  color:"#2ecc71" },
    { label:"COFINS", field:"cofins", color:"#5b9cf6" },
    { label:"IPI",  field:"ipi",  color:"#e74c3c" },
  ];
  byId("taxBreakdown").innerHTML = items.map(({ label, field, color }) => {
    const value  = metrics.taxBreakdown[field] || 0;
    const ofRev  = metrics.totalRevenue ? (value / metrics.totalRevenue) * 100 : null;
    const ofTax  = totalTaxes ? (value / totalTaxes) * 100 : null;
    const barPct = (ofTax || 0).toFixed(1);
    return `
      <article class="tax-item">
        <strong>${escapeHtml(label)}</strong>
        <div class="tax-item-value">${escapeHtml(formatCurrency(value))}</div>
        <div class="tax-item-meta">
          <span>${escapeHtml(formatPercent(ofRev))} da receita</span>
          <span>${escapeHtml(formatPercent(ofTax))} dos tributos</span>
        </div>
        <div class="tax-item-bar">
          <div class="tax-item-bar-fill" style="width:${barPct}%;background:${color}"></div>
        </div>
      </article>`;
  }).join("");
}

function renderStockTypeChart(chartId, monthlyRows, type) {
  destroyChart(chartId);
  const label = type === "remessa" ? "Remessas e similares" : "Baixas de estoque";
  const color = type === "remessa" ? "#5b9cf6" : "#f39c12";
  const chart = new Chart(byId(chartId), {
    type:"bar",
    data:{
      labels:monthlyRows.map(r => formatMonth(r.month)),
      datasets:[{ label, data:monthlyRows.map(r => r.value), backgroundColor:color, borderRadius:6 }],
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      onClick(_e, els) {
        if (!els.length) return;
        const month = monthlyRows[els[0].index].month;
        const rows  = state.filteredStock.filter(r =>
          type === "remessa" ? isRemessaLikeType(r.type) && r.month === month : r.type === type && r.month === month
        );
        openModal(`${label} — ${formatMonth(month)}`, `${formatCurrency(sum(rows,"value"))} no período`, rows, "stock");
      },
      plugins:{
        legend:{ display:false },
        tooltip:{ callbacks:{ label:ctx => `${label}: ${formatCurrency(ctx.parsed.y)}` } },
      },
      scales:{
        x:{ grid:{display:false} },
        y:{ ticks:{ callback:v => formatCurrency(v) } },
      },
    },
  });
  state.charts.set(chartId, chart);
}

/* ─── GENERIC TABLE RENDERER ─────────────────────────────────────── */
function renderTable(tableId, headers, rows, clickHandler = null) {
  const table = byId(tableId);
  if (!table) return;
  const headerHtml = headers.map(h =>
    `<th class="${h.numeric ? "num" : ""}">${escapeHtml(h.label)}</th>`
  ).join("");

  if (!rows.length) {
    table.innerHTML = `
      <thead><tr>${headerHtml}</tr></thead>
      <tbody><tr><td colspan="${headers.length}" class="empty-cell">Nenhum dado para este recorte.</td></tr></tbody>`;
    return;
  }

  table.innerHTML = `
    <thead><tr>${headerHtml}</tr></thead>
    <tbody>
      ${rows.map((row, i) =>
        `<tr class="${clickHandler ? "clickable" : ""}" data-index="${i}">
          ${row.cells.map(c => `<td class="${c.numeric ? "num" : ""}">${c.html}</td>`).join("")}
        </tr>`
      ).join("")}
    </tbody>`;

  if (clickHandler) {
    table.querySelectorAll("tbody tr.clickable").forEach(tr => {
      tr.addEventListener("click", () => clickHandler(Number(tr.dataset.index)));
    });
  }
}

function getTopLimit(id, fallback = 15) {
  const v = Number(byId(id)?.value);
  return Number.isFinite(v) && v > 0 ? v : fallback;
}

/* ─── TABLE RENDERERS ────────────────────────────────────────────── */
function renderClientsTable(metrics) {
  const rows = sortByMetric(metrics.clients, byId("topClientsSort").value || "revenue")
    .slice(0, getTopLimit("topClientsLimit", 15));
  renderTable("clientsTable",
    [{ label:"#" },{ label:"Cliente" },{ label:"UF" },
     { label:"Receita",numeric:true },{ label:"NFs",numeric:true },
     { label:"Devolução",numeric:true },{ label:"Margem %",numeric:true },{ label:"Margem R$",numeric:true }],
    rows.map((row, i) => ({
      cells:[
        { html:escapeHtml(i+1) },
        { html:escapeHtml(row.client) },
        { html:`<span class="pill info">${escapeHtml(row.uf||"—")}</span>` },
        { html:escapeHtml(formatCurrency(row.revenue)), numeric:true },
        { html:escapeHtml(formatNumber(row.notes)), numeric:true },
        { html:row.returns > 0 ? `<span class="pill warn">${escapeHtml(formatCurrency(row.returns))}</span>` : `<span style="color:var(--muted-2)">—</span>`, numeric:true },
        { html:row.marginPct == null ? "—" : `<span class="pill ${row.marginPct>=0?"good":"bad"}">${escapeHtml(formatPercent(row.marginPct))}</span>`, numeric:true },
        { html:escapeHtml(formatCurrency(row.marginValue||0)), numeric:true },
      ],
    })),
    (i) => {
      const client = rows[i];
      const detailRows = state.filteredSales.filter(r => r.client === client.client);
      openModal(`Cliente: ${client.client}`, `${formatCurrency(client.revenue)} no período`, detailRows, "sales");
    }
  );
}

function renderProductsTable(metrics) {
  const rows = sortByMetric(metrics.products, byId("topProductsSort").value || "revenue")
    .slice(0, getTopLimit("topProductsLimit", 15));
  renderTable("productsTable",
    [{ label:"#" },{ label:"SKU / Produto" },
     { label:"Receita",numeric:true },{ label:"Qtd",numeric:true },
     { label:"Part.%",numeric:true },{ label:"Margem %",numeric:true },{ label:"Margem R$",numeric:true }],
    rows.map((row, i) => {
      const share = metrics.totalRevenue ? (row.revenue / metrics.totalRevenue) * 100 : null;
      return { cells:[
        { html:escapeHtml(i+1) },
        { html:escapeHtml(row.item) },
        { html:escapeHtml(formatCurrency(row.revenue)), numeric:true },
        { html:escapeHtml(formatNumber(row.quantity)), numeric:true },
        { html:escapeHtml(formatPercent(share)), numeric:true },
        { html:row.marginPct == null ? "—" : `<span class="pill ${row.marginPct>=0?"good":"bad"}">${escapeHtml(formatPercent(row.marginPct))}</span>`, numeric:true },
        { html:escapeHtml(formatCurrency(row.marginValue||0)), numeric:true },
      ]};
    }),
    (i) => {
      const product = rows[i];
      const detailRows = state.filteredSales.filter(r => r.item === product.item);
      openModal(`Produto: ${product.item}`, `${formatCurrency(product.revenue)} no período`, detailRows, "sales");
    }
  );
}

function renderPositiveSkuTable(metrics) {
  const rows = metrics.positiveSkuMargin.slice(0, getTopLimit("topPositiveSkuLimit", 15));
  renderTable("positiveSkuTable",
    [{ label:"#" },{ label:"SKU / Produto" },
     { label:"Receita",numeric:true },{ label:"Margem %",numeric:true },{ label:"Margem R$",numeric:true }],
    rows.map((row, i) => ({ cells:[
      { html:escapeHtml(i+1) },
      { html:escapeHtml(row.item) },
      { html:escapeHtml(formatCurrency(row.revenue)), numeric:true },
      { html:`<span class="pill good">${escapeHtml(formatPercent(row.marginPct))}</span>`, numeric:true },
      { html:escapeHtml(formatCurrency(row.marginValue)), numeric:true },
    ]})),
    (i) => {
      const p = rows[i];
      openModal(`Margem positiva: ${p.item}`, `${formatCurrency(p.marginValue)} de margem`,
        state.filteredSales.filter(r => r.item === p.item), "sales");
    }
  );
}

function renderNegativeMarginTable(metrics) {
  const rows = metrics.negativeSkuMargin.slice(0, getTopLimit("topNegativeSkuLimit", 15));
  renderTable("negativeMarginTable",
    [{ label:"#" },{ label:"SKU / Produto" },
     { label:"Receita",numeric:true },{ label:"Margem %",numeric:true },{ label:"Prejuízo R$",numeric:true }],
    rows.map((row, i) => ({ cells:[
      { html:escapeHtml(i+1) },
      { html:escapeHtml(row.item) },
      { html:escapeHtml(formatCurrency(row.revenue)), numeric:true },
      { html:`<span class="pill bad">${escapeHtml(formatPercent(row.marginPct))}</span>`, numeric:true },
      { html:`<span style="color:var(--red)">${escapeHtml(formatCurrency(Math.abs(row.marginValue)))}</span>`, numeric:true },
    ]})),
    (i) => {
      const p = rows[i];
      openModal(`Margem negativa: ${p.item}`, `${formatCurrency(p.marginValue)} de margem`,
        state.filteredSales.filter(r => r.item === p.item), "sales");
    }
  );
}

function renderBadClientsTable(metrics) {
  const rows = metrics.badClients.slice(0, getTopLimit("topBadClientsLimit", 15));
  renderTable("badClientsTable",
    [{ label:"#" },{ label:"Cliente" },
     { label:"Receita",numeric:true },{ label:"Devoluções",numeric:true },
     { label:"Margem %",numeric:true },{ label:"Prejuízo R$",numeric:true }],
    rows.map((row, i) => ({ cells:[
      { html:escapeHtml(i+1) },
      { html:escapeHtml(row.client) },
      { html:escapeHtml(formatCurrency(row.revenue)), numeric:true },
      { html:row.returns > 0 ? `<span class="pill warn">${escapeHtml(formatCurrency(row.returns))}</span>` : "—", numeric:true },
      { html:row.marginPct == null ? "—" : `<span class="pill bad">${escapeHtml(formatPercent(row.marginPct))}</span>`, numeric:true },
      { html:`<span style="color:var(--red)">${escapeHtml(formatCurrency(Math.abs(row.marginValue)))}</span>`, numeric:true },
    ]})),
    (i) => {
      const client = rows[i];
      openModal(`Cliente crítico: ${client.client}`, `${formatCurrency(client.marginValue)} de margem`,
        state.filteredSales.filter(r => r.client === client.client), "sales");
    }
  );
}

function renderTaxOutlierTable(metrics) {
  const rows = metrics.taxOutliers.slice(0, getTopLimit("topTaxLimit", 15));
  renderTable("taxOutlierTable",
    [{ label:"SKU / Produto" },{ label:"Receita",numeric:true },
     { label:"Imposto %",numeric:true },{ label:"Desvio σ",numeric:true }],
    rows.map((row) => ({ cells:[
      { html:escapeHtml(row.item) },
      { html:escapeHtml(formatCurrency(row.revenue)), numeric:true },
      { html:`<span class="pill ${row.zScore>=0?"bad":"info"}">${escapeHtml(formatPercent(row.taxPct))}</span>`, numeric:true },
      { html:escapeHtml(`${row.zScore>=0?"+":""}${formatDecimal(row.zScore)}σ`), numeric:true },
    ]})),
    (i) => {
      const p = rows[i];
      openModal(`Tributação atípica: ${p.item}`, `Imposto efetivo ${formatPercent(p.taxPct)}`,
        state.filteredSales.filter(r => r.item === p.item && r.type === "venda"), "sales");
    }
  );
}

function renderStockTypeTable(tableId, rows, type) {
  const sliced = rows.slice(0, 15);
  if (type === "baixa_estoque" && rows.length && rows[0]?.cfop !== undefined) {
    renderTable(tableId,
      [{ label:"#" },{ label:"CFOP" },{ label:"Descrição" },
       { label:"Valor",numeric:true },{ label:"Qtd",numeric:true },{ label:"NFs",numeric:true }],
      sliced.map((row,i) => ({ cells:[
        { html:escapeHtml(i+1) },
        { html:`<code>${escapeHtml(row.cfop)}</code>` },
        { html:escapeHtml(row.description) },
        { html:escapeHtml(formatCurrency(row.value)), numeric:true },
        { html:escapeHtml(formatNumber(row.quantity)), numeric:true },
        { html:escapeHtml(formatNumber(row.notes)), numeric:true },
      ]})),
      (i) => {
        const r = sliced[i];
        openModal(`Baixa CFOP ${r.cfop}`, r.description,
          state.filteredStock.filter(row => row.type === type && row.cfop === r.cfop), "stock");
      }
    );
    return;
  }
  renderTable(tableId,
    [{ label:"#" },{ label:"SKU / Produto" },
     { label:"Valor",numeric:true },{ label:"Qtd",numeric:true },{ label:"NFs",numeric:true }],
    sliced.map((row,i) => ({ cells:[
      { html:escapeHtml(i+1) },
      { html:escapeHtml(row.item) },
      { html:escapeHtml(formatCurrency(row.value)), numeric:true },
      { html:escapeHtml(formatNumber(row.quantity)), numeric:true },
      { html:escapeHtml(formatNumber(row.notes)), numeric:true },
    ]})),
    (i) => {
      const p = sliced[i];
      const detailRows = state.filteredStock.filter(r =>
        type === "remessa" ? isRemessaLikeType(r.type) && r.item === p.item : r.type === type && r.item === p.item
      );
      openModal(type === "remessa" ? `Remessa: ${p.item}` : `Baixa: ${p.item}`,
        `${formatCurrency(p.value)} em movimentação`, detailRows, "stock");
    }
  );
}

/* ─── SUMMARY CARDS (CFO KPI BLOCKS) ────────────────────────────── */
function kpiColor(value, thresholds) {
  // thresholds: [{ limit, cls }] sorted from critical → ok
  // value OP limit triggers cls
  for (const { limit, cls, op = "lt" } of thresholds) {
    if (op === "lt"  && value <  limit) return cls;
    if (op === "lte" && value <= limit) return cls;
    if (op === "gt"  && value >  limit) return cls;
    if (op === "gte" && value >= limit) return cls;
  }
  return "kpi-green";
}

function renderSummaryCards(metrics, compareMetrics) {
  const cmp = compareMetrics;

  // ── helpers ──
  const chip = (val, cVal, { inverse=false, suffix="" }={}) => {
    if (cVal != null && val != null) return buildDeltaChip(val, cVal, { inverse });
    return { text: suffix || "—", tone:"neutral" };
  };
  const momChip = metrics.monthOverMonthRevenue;

  // ─────────────────────────────────────────────────────────────────
  // BLOCO 1 — DRE WATERFALL: Receita Bruta → Deduções → Receita Líquida
  // ─────────────────────────────────────────────────────────────────
  const dreCards = [
    {
      label:"Receita Bruta",
      kpiClass:"kpi-blue",
      value: formatCurrency(metrics.receitaBruta),
      note: `${formatNumber(metrics.totalNotes)} NFs de venda no período`,
      chip: cmp ? buildDeltaChip(metrics.receitaBruta, cmp.receitaBruta)
               : momChip == null ? { text:"Sem comparativo mensal", tone:"neutral" }
               : { text:`${momChip>=0?"+":""}${formatDecimal(momChip)}% vs mês ant.`, tone:momChip>=0?"good":"bad" },
    },
    {
      label:"(–) Devoluções",
      kpiClass: metrics.returnRate != null && metrics.returnRate > 3 ? "kpi-red" : "kpi-amber",
      value: formatCurrency(metrics.totalReturns),
      note: `${formatPercent(metrics.returnRate)} da receita bruta · ${formatPercent(metrics.devolucaoSobreLiquida)} da líquida`,
      chip: cmp ? buildDeltaChip(metrics.totalReturns, cmp.totalReturns, { inverse:true })
               : { text: metrics.returnRate != null && metrics.returnRate > 3 ? "⚠ Acima de 3%" : "Dentro do limite", tone: metrics.returnRate != null && metrics.returnRate > 3 ? "bad" : "good" },
    },
    {
      label:"(–) Impostos",
      kpiClass:"kpi-red",
      value: formatCurrency(metrics.totalTaxes),
      note: `${formatPercent(metrics.taxPct)} sobre receita líquida · ${formatPercent(metrics.cargaSobreLiquida)} sobre líquida`,
      chip: cmp ? buildDeltaChip(metrics.totalTaxes, cmp.totalTaxes, { inverse:true })
               : { text: getSelectedTaxLabel(), tone:"neutral" },
    },
    {
      label:"Receita Líquida",
      kpiClass:"kpi-blue",
      value: formatCurrency(metrics.totalRevenue),
      note: `Taxa de conversão: ${formatPercent(metrics.taxaConversao)} da bruta`,
      chip: cmp ? buildDeltaChip(metrics.totalRevenue, cmp.totalRevenue)
               : { text:`Deduções: ${formatPercent(metrics.deducoesPct)} da bruta`, tone: metrics.deducoesPct != null && metrics.deducoesPct > 25 ? "bad" : "neutral" },
    },
    {
      label:"(–) CPV",
      kpiClass: metrics.cpvPct != null && metrics.cpvPct > 80 ? "kpi-red" : metrics.cpvPct != null && metrics.cpvPct > 60 ? "kpi-amber" : "",
      value: metrics.totalCpv ? formatCurrency(metrics.totalCpv) : "—",
      note: metrics.cpvPct != null ? `${formatPercent(metrics.cpvPct)} sobre receita com custo mapeado` : "Custo não mapeado para todas as linhas",
      chip: cmp
        ? (cmp.cpvPct != null ? buildDeltaChip(metrics.cpvPct, cmp.cpvPct, { inverse:true }) : { text:"—", tone:"neutral" })
        : { text: metrics.markup != null ? `Markup: ${formatDecimal(metrics.markup)}×` : "—", tone:"neutral" },
    },
    {
      label:"Margem Bruta",
      kpiClass: metrics.grossMarginPct == null ? "" : metrics.grossMarginPct < 0 ? "kpi-red" : metrics.grossMarginPct < 15 ? "kpi-amber" : "kpi-green",
      value: metrics.grossMarginPct == null ? "—" : formatPercent(metrics.grossMarginPct),
      note: metrics.mappedRevenue
        ? `${formatCurrency(metrics.totalMargin)} · base ${formatCurrency(metrics.mappedRevenue)} ${state.config.margin.deductTaxes?"pós-tributos":"pré-tributos"}`
        : "Sem custo mapeado",
      chip: cmp
        ? buildDeltaChip(metrics.grossMarginPct||0, cmp.grossMarginPct||0)
        : { text: metrics.totalMargin ? formatCurrency(metrics.totalMargin) : "Sem margem", tone: metrics.totalMargin != null ? (metrics.totalMargin>=0?"good":"bad") : "neutral" },
    },
  ];

  // ─────────────────────────────────────────────────────────────────
  // BLOCO 2 — KPIs OPERACIONAIS DERIVADOS
  // ─────────────────────────────────────────────────────────────────
  const opsCards = [
    {
      label:"Taxa de Conversão",
      kpiClass: metrics.taxaConversao != null && metrics.taxaConversao < 70 ? "kpi-red" : metrics.taxaConversao != null && metrics.taxaConversao < 80 ? "kpi-amber" : "kpi-green",
      value: metrics.taxaConversao != null ? formatPercent(metrics.taxaConversao) : "—",
      note: "Receita Líquida / Receita Bruta — quanto chega após deduções",
      chip: cmp ? buildDeltaChip(metrics.taxaConversao||0, cmp.taxaConversao||0)
               : { text:"Mínimo saudável: 75%", tone:"neutral" },
    },
    {
      label:"Índice de Perda Total",
      kpiClass: metrics.indicePerdaTotal != null && metrics.indicePerdaTotal > 30 ? "kpi-red" : metrics.indicePerdaTotal != null && metrics.indicePerdaTotal > 20 ? "kpi-amber" : "kpi-green",
      value: metrics.indicePerdaTotal != null ? formatPercent(metrics.indicePerdaTotal) : "—",
      note: "(Devoluções + Impostos) / Receita Bruta — quanto morre antes de virar margem",
      chip: cmp ? buildDeltaChip(metrics.indicePerdaTotal||0, cmp.indicePerdaTotal||0, { inverse:true })
               : { text: metrics.indicePerdaTotal != null && metrics.indicePerdaTotal > 20 ? "⚠ Acima de 20%" : "Dentro do padrão", tone: metrics.indicePerdaTotal != null && metrics.indicePerdaTotal > 20 ? "bad" : "good" },
    },
    {
      label:"Eficiência Comercial",
      kpiClass: metrics.eficienciaComercial == null ? "" : metrics.eficienciaComercial < 0 ? "kpi-red" : metrics.eficienciaComercial < 10 ? "kpi-amber" : "kpi-green",
      value: metrics.eficienciaComercial != null ? formatPercent(metrics.eficienciaComercial) : "—",
      note: "Margem Bruta / Receita Bruta — KPI de diretoria",
      chip: cmp ? buildDeltaChip(metrics.eficienciaComercial||0, cmp.eficienciaComercial||0)
               : { text:"Meta recomendada: > 10%", tone:"neutral" },
    },
    {
      label:"Markup Médio",
      kpiClass: metrics.markup != null && metrics.markup < 1.2 ? "kpi-red" : metrics.markup != null && metrics.markup < 1.5 ? "kpi-amber" : "",
      value: metrics.markup != null ? `${formatDecimal(metrics.markup)}×` : "—",
      note: `Receita / CPV — ${metrics.cpvPct != null ? `CPV = ${formatPercent(metrics.cpvPct)} da receita` : "CPV parcialmente mapeado"}`,
      chip: cmp && cmp.markup != null ? buildDeltaChip(metrics.markup||0, cmp.markup||0)
               : { text: metrics.markup != null && metrics.markup < 1.3 ? "⚠ Markup baixo" : "Markup calculado", tone: metrics.markup != null && metrics.markup < 1.3 ? "bad" : "neutral" },
    },
    {
      label:"Ticket Médio / NF",
      kpiClass:"",
      value: metrics.avgTicket != null ? formatCurrency(metrics.avgTicket) : "—",
      note: `Receita bruta ÷ ${formatNumber(metrics.totalNotes)} NFs de venda`,
      chip: cmp && cmp.avgTicket != null ? buildDeltaChip(metrics.avgTicket||0, cmp.avgTicket||0)
               : { text:"por NF emitida", tone:"neutral" },
    },
    {
      label:"Concentração Top-3",
      kpiClass: metrics.top3Share != null && metrics.top3Share > 50 ? "kpi-red" : metrics.top3Share != null && metrics.top3Share > 30 ? "kpi-amber" : "kpi-green",
      value: metrics.top3Share != null ? formatPercent(metrics.top3Share) : "—",
      note: "Participação dos 3 maiores clientes na receita total",
      chip: cmp ? buildDeltaChip(metrics.top3Share||0, cmp.top3Share||0, { inverse:true })
               : { text: metrics.top3Share != null && metrics.top3Share > 50 ? "⚠ Risco de concentração" : "Diversificada", tone: metrics.top3Share != null && metrics.top3Share > 50 ? "bad" : "good" },
    },
    {
      label:"Clientes Ativos",
      kpiClass:"kpi-blue",
      value: formatNumber(metrics.totalClients),
      note: `${formatNumber(metrics.totalSkus)} SKUs ativos no período`,
      chip: cmp ? buildDeltaChip(metrics.totalClients, cmp.totalClients)
               : { text:`${formatNumber(metrics.totalSkus)} SKUs`, tone:"neutral" },
    },
    {
      label:"Remessas",
      kpiClass:"",
      value: formatCurrency(metrics.totalRemessa),
      note: `${formatNumber(state.filteredStock.filter(r=>isRemessaLikeType(r.type)).length)} linhas · remessa, bonificação e similares`,
      chip: cmp ? buildDeltaChip(metrics.totalRemessa, cmp.totalRemessa, { inverse:true })
               : { text: state.config.stock.remessa ? "Incluída" : "Fora da composição", tone:"neutral" },
    },
    {
      label:"Baixas de Estoque",
      kpiClass:"",
      value: formatCurrency(metrics.totalBaixa),
      note: `${formatNumber(state.filteredStock.filter(r=>r.type==="baixa_estoque").length)} linhas de baixa`,
      chip: cmp ? buildDeltaChip(metrics.totalBaixa, cmp.totalBaixa, { inverse:true })
               : { text: state.config.stock.baixa ? "Incluída" : "Fora da composição", tone:"neutral" },
    },
  ];

  // ─────────────────────────────────────────────────────────────────
  // RENDER
  // ─────────────────────────────────────────────────────────────────
  const renderBlock = (id, cards) => {
    const el = byId(id);
    if (!el) return;
    el.innerHTML = cards.map(card => `
      <article class="panel summary-card ${escapeHtml(card.kpiClass||"")}">
        <span class="summary-label">${escapeHtml(card.label)}</span>
        <div class="summary-value">${escapeHtml(card.value)}</div>
        <div class="summary-note">${escapeHtml(card.note)}</div>
        <span class="summary-chip ${card.chip.tone}">${escapeHtml(card.chip.text)}</span>
      </article>`).join("");
  };

  renderBlock("summaryGridDre", dreCards);
  renderBlock("summaryGridOps", opsCards);
  renderAlerts(metrics.alerts || []);
}

/* ─── ALERTS PANEL ───────────────────────────────────────────────── */
function renderAlerts(alerts) {
  const el = byId("alertsPanel");
  if (!el) return;
  if (!alerts.length) {
    el.innerHTML = `<div class="alert-item alert-ok"><span class="alert-icon">✓</span><span>Sem alertas ativos no período selecionado.</span></div>`;
    return;
  }
  el.innerHTML = alerts.map(a => `
    <div class="alert-item alert-${a.level}">
      <span class="alert-icon">${a.level === "critical" ? "⚠" : "○"}</span>
      <span>${escapeHtml(a.msg)}</span>
    </div>`).join("");
}

/* ─── ANALYSIS PANEL ─────────────────────────────────────────────── */
function severityBadge(level) {
  const map = {
    critical:{ label:"Crítico",  cls:"bad" },
    warning: { label:"Atenção",  cls:"warn" },
    ok:      { label:"Saudável", cls:"good" },
    info:    { label:"Info",     cls:"info" },
  };
  const { label, cls } = map[level] || map.info;
  return `<span class="pill ${cls} analysis-item-badge">${escapeHtml(label)}</span>`;
}

function renderAnalysisPanel(metrics, compareMetrics) {
  const analysis = [];

  // 1. Receita
  if (compareMetrics) {
    const delta = percentageChange(metrics.totalRevenue, compareMetrics.totalRevenue);
    analysis.push({
      title:"Comparativo de receita", severity:delta==null?"info":delta>=0?"ok":"warning",
      fact: delta==null
        ? "Período B sem base suficiente para comparação."
        : `Receita líquida do período A ficou ${delta>=0?"acima":"abaixo"} do período B em ${formatPercent(Math.abs(delta))}.`,
      impact: delta==null ? "Leitura qualitativa apenas."
        : delta>=0 ? "Tração comercial superior ao comparativo."
        : "Queda de faturamento pressiona margem e diluição de custos.",
      recommendation: delta==null
        ? "Preencha um período B comparável para fechar a análise."
        : delta>=0 ? "Validar se o crescimento veio de preço, volume ou mix de produto."
        : "Abrir receita por cliente, SKU e UF para localizar a queda.",
    });
  } else {
    const mom = metrics.monthOverMonthRevenue;
    analysis.push({
      title:"Leitura de receita", severity:mom==null?"info":mom>=0?"ok":"warning",
      fact: mom==null
        ? "Período insuficiente para comparação mensal."
        : `Receita líquida variou ${mom>=0?"+":""}${formatDecimal(mom)}% no último mês. Taxa de conversão: ${formatPercent(metrics.taxaConversao)}.`,
      impact: mom==null ? "Amplie o período para leitura de tendência."
        : mom>=0 ? "Curva de vendas em aceleração." : "Curva de vendas em desaceleração — pressão sobre margem e diluição de custos fixos.",
      recommendation: mom==null ? "Selecione ao menos 2 meses completos."
        : "Confirmar se variação veio de preço, volume ou perda de clientes ativos.",
    });
  }

  // 2. Devoluções — com thresholds calibrados
  const ret = metrics.returnRate;
  const devImpLiq = metrics.devolucaoSobreLiquida;
  analysis.push({
    title:"Devoluções", severity:ret==null?"info":ret>5?"critical":ret>3?"critical":ret>1?"warning":"ok",
    fact:`Devoluções somam ${formatCurrency(metrics.totalReturns)} — ${formatPercent(ret)} da receita bruta e ${formatPercent(devImpLiq)} da receita líquida.`,
    impact:ret!=null && ret>3
      ? `⚠ Devolução acima de 3%: impacto direto sobre receita líquida e compressão de margem.`
      : ret!=null && ret>1
        ? "Devolução entre 1–3%: monitorar de perto por cliente e SKU."
        : "Devolução abaixo de 1%: nível saudável no período.",
    recommendation:ret!=null && ret>3
      ? `Cruzar com ranking de clientes e SKUs. Investigar causa-raiz: qualidade, logística ou precificação.`
      : ret!=null && ret>1
        ? "Acompanhar evolução mensal e identificar produtos com maior índice de retorno."
        : "Manter monitoramento preventivo para evitar deterioração silenciosa.",
  });

  // 3. Impostos — com carga sobre líquida
  const taxOutCnt = metrics.taxOutliers.length;
  analysis.push({
    title:"Tributação", severity:taxOutCnt>3?"critical":taxOutCnt>0?"warning":"ok",
    fact:metrics.taxPct==null ? "Sem base para calcular carga."
      : `Carga tributária: ${formatPercent(metrics.taxPct)} sobre receita líquida · ${formatPercent(metrics.cargaSobreLiquida)} sobre receita após deduções · ${getSelectedTaxLabel()}.`,
    impact:taxOutCnt
      ? `${formatNumber(taxOutCnt)} SKU${taxOutCnt>1?"s":""} com tributação atípica — risco de distorção de rentabilidade e erro de classificação fiscal.`
      : "Carga tributária sem outliers relevantes neste recorte.",
    recommendation:taxOutCnt
      ? `Revisar CFOP, NCM e enquadramento para: ${escapeHtml(metrics.taxOutliers.slice(0,2).map(t=>t.item).join(", "))}.`
      : "Usar este painel para validar mudanças fiscais e oportunidades de economia tributária.",
  });

  // 4. Margem e CPV
  const marg = metrics.grossMarginPct;
  const cpv  = metrics.cpvPct;
  analysis.push({
    title:"Margem Bruta e CPV", severity:marg==null?"info":marg<0?"critical":marg<15?"warning":"ok",
    fact:marg==null
      ? "Custo insuficientemente mapeado — margem não confiável para decisão."
      : `Margem Bruta: ${formatPercent(marg)} · CPV: ${cpv!=null?formatPercent(cpv):"não mapeado"} · Markup: ${metrics.markup!=null?`${formatDecimal(metrics.markup)}×`:"—"} · Eficiência Comercial: ${metrics.eficienciaComercial!=null?formatPercent(metrics.eficienciaComercial):"—"}.`,
    impact:marg==null ? "Dashboard forte em receita, fraco em rentabilidade. Preencher custo é ação prioritária."
      : marg<0 ? "Resultado destruindo valor econômico — cada venda gera prejuízo."
      : marg<15 ? "Margem apertada: qualquer desvio de preço, devolução ou imposto pode eliminar o resultado."
      : "Margem positiva e saudável — defender mix e tributação.",
    recommendation:marg==null ? "Priorizar preenchimento de custo por SKU para ativar cálculo de rentabilidade."
      : `Atacar primeiro os ${formatNumber(metrics.negativeSkuMargin.length)} SKUs com margem negativa e ${formatNumber(metrics.badClients.length)} clientes com resultado negativo.`,
  });

  // 5. Concentração e eficiência comercial
  const conc = metrics.top3Share;
  analysis.push({
    title:"Concentração de carteira", severity:conc==null?"info":conc>50?"critical":conc>30?"warning":"ok",
    fact:metrics.clients[0]
      ? `Top-3 clientes respondem por ${formatPercent(conc)} da receita. Maior: ${metrics.clients[0].client} com ${formatCurrency(metrics.clients[0].revenue)}.`
      : "Sem cliente dominante no recorte.",
    impact:conc!=null && conc>50
      ? "Dependência elevada — ruptura comercial com top cliente gera impacto imediato no resultado."
      : conc!=null && conc>30
        ? "Concentração relevante — monitorar saúde financeira e margem dos top clientes."
        : "Base de clientes relativamente diversificada.",
    recommendation:conc!=null && conc>30
      ? "Analisar margem e devolução dos top-3 antes de ampliar volume. Diversificar carteira."
      : "Usar remessas e baixas para identificar movimentação sem conversão em receita.",
  });

  // 6. Índice de perda total
  if (metrics.indicePerdaTotal != null) {
    analysis.push({
      title:"Índice de Perda Total", severity:metrics.indicePerdaTotal>30?"critical":metrics.indicePerdaTotal>20?"warning":"ok",
      fact:`${formatPercent(metrics.indicePerdaTotal)} da receita bruta se perde em devoluções e impostos antes de virar margem.`,
      impact:metrics.indicePerdaTotal>30
        ? "Perda elevada: menos de 70% da receita bruta resta para cobrir custo e gerar margem."
        : metrics.indicePerdaTotal>20
          ? "Perda moderada: monitorar evolução para evitar compressão de margem bruta."
          : "Nível de perda aceitável no período.",
      recommendation:`Taxa de conversão de receita atual: ${formatPercent(metrics.taxaConversao)}. Meta recomendada: acima de 75%.`,
    });
  }

  byId("analysisList").innerHTML = analysis.map(item => `
    <article class="analysis-item">
      <div class="analysis-item-head">
        ${severityBadge(item.severity)}
        <span class="analysis-item-title">${escapeHtml(item.title)}</span>
      </div>
      <div class="analysis-item-body">
        <span><strong>Fato:</strong> ${escapeHtml(item.fact)}</span>
        <span><strong>Impacto:</strong> ${escapeHtml(item.impact)}</span>
        <span><strong>Recomendação:</strong> ${escapeHtml(item.recommendation)}</span>
      </div>
    </article>`).join("");
}

/* ─── FILTER INFO ────────────────────────────────────────────────── */
function renderFilterInfo(metrics, compareMetrics) {
  const parts = [
    `Período A: ${formatRangeBr(byId("dateFrom").value, byId("dateTo").value)}`,
    `${formatNumber(state.filteredSales.length)} linhas`,
    `${formatNumber(metrics.totalNotes)} NFs`,
  ];
  if (compareMetrics && byId("compareToggle").checked) {
    parts.push(`Período B: ${formatRangeBr(byId("compareFrom").value, byId("compareTo").value)}`);
    parts.push(`${formatNumber(state.compareSales.length)} linhas comp.`);
  }
  byId("filterInfo").textContent = parts.join(" · ");
}

/* ─── SOURCE SUMMARY ─────────────────────────────────────────────── */
function summarizeSource(source) {
  const items = [
    `<strong>Formato:</strong> ${escapeHtml(source.format)}`,
    `<strong>Arquivo:</strong> ${escapeHtml(source.fileName)}`,
    `<strong>Aba:</strong> ${escapeHtml(source.sheetName)}`,
    `<strong>Cobertura:</strong> ${escapeHtml(formatRangeBr(source.coverageFrom, source.coverageTo, { shortYear:false }))}`,
    `<strong>Linhas totais:</strong> ${formatNumber(source.totalRows)}`,
    `<strong>Vendas/devoluções válidas:</strong> ${formatNumber(source.salesRows.length)}`,
    `<strong>Remessas/baixas válidas:</strong> ${formatNumber(source.stockRows.length)}`,
  ];
  if ((source.unclassifiedRows || 0) > 0) {
    items.push(`<strong>CFOPs não classificados:</strong> <span style="color:var(--amber)">${formatNumber(source.unclassifiedRows)} linhas — verifique o configurador de CFOP</span>`);
  }
  if (source.ignoredRows > 0) {
    items.push(`<strong>Linhas ignoradas (dados incompletos):</strong> ${formatNumber(source.ignoredRows)}`);
  }
  byId("sourceSummary").classList.remove("empty");
  byId("sourceSummary").innerHTML = items.join("<br>");
}

function buildAssumptions() {
  const items = [
    `Receita líquida = ${getRevenueCompositionLabel()}.`,
    `Impostos = ${getSelectedTaxLabel()}.`,
    state.config.margin.deductTaxes
      ? "Margem = receita − custo − tributos selecionados (pós-tributos)."
      : "Margem = receita − custo (pré-tributos).",
    "Remessas, bonificações, doações, brindes e baixas ficam fora da receita.",
    "Tributação atípica: desvio padrão por SKU em vendas com receita > R$ 5.000.",
    "Ticket médio = receita bruta positiva ÷ número de NFs de venda.",
    "Concentração = participação acumulada dos 3 maiores clientes.",
  ];
  byId("assumptionsList").innerHTML = items.map(t => `<li>${escapeHtml(t)}</li>`).join("");
  byId("marginExplain").textContent = state.config.margin.deductTaxes
    ? "Margem pós-tributos: receita − custo − tributos selecionados."
    : "Margem pré-tributos: receita − custo.";
}

/* ─── CFOP REGISTRY UI ───────────────────────────────────────────── */
function renderCfopRegistryMeta() {
  const codes     = [...state.cfopRegistry.values()];
  const overrides = codes.filter(r => r.customType).length;
  byId("cfopRegistryMeta").textContent = !codes.length
    ? "A base oficial será aplicada aos CFOPs detectados na base carregada."
    : `${formatNumber(codes.length)} CFOPs detectados · ${formatNumber(overrides)} reclassificações · referência: Ajuste SINIEF 03/24`;
}

function renderCfopConfigTable() {
  const query = normalizeText(byId("cfopSearch")?.value);
  const rows  = [...state.cfopRegistry.values()]
    .sort((a,b) => a.code.localeCompare(b.code))
    .filter(row => {
      if (!query) return true;
      return [row.code, row.description, row.analysisType, row.officialType]
        .map(normalizeText).some(v => v.includes(query));
    });

  renderTable("cfopConfigTable",
    [{ label:"CFOP" },{ label:"Descrição oficial" },{ label:"Padrão" },{ label:"Análise" },
     { label:"Linhas",numeric:true },{ label:"Valor",numeric:true }],
    rows.map(row => ({ cells:[
      { html:`<code>${escapeHtml(row.code)}</code>` },
      { html:escapeHtml(row.description) },
      { html:`<span class="pill info">${escapeHtml(friendlyType(row.officialType))}</span>` },
      { html:`<select class="cfop-type-select" data-cfop="${escapeHtml(row.code)}">
          ${CFOP_TYPE_OPTIONS.map(o =>
            `<option value="${escapeHtml(o.value)}" ${o.value===row.analysisType?"selected":""}>${escapeHtml(o.label)}</option>`
          ).join("")}
        </select>` },
      { html:escapeHtml(formatNumber(row.count)), numeric:true },
      { html:escapeHtml(formatCurrency(row.totalValue)), numeric:true },
    ]}))
  );

  byId("cfopConfigTable").querySelectorAll(".cfop-type-select").forEach(sel => {
    sel.addEventListener("change", (e) => {
      markWorkingCopyFromLoadedVersion();
      const code = e.target.dataset.cfop;
      const meta = state.cfopRegistry.get(code);
      if (!meta) return;
      meta.customType   = e.target.value === meta.officialType ? null : e.target.value;
      meta.analysisType = meta.customType || meta.officialType;
      rebuildRowsFromCfopRegistry();
      renderCfopRegistryMeta();
      renderCfopConfigTable();
      if (state.source) applyFilters();
    });
  });
}

/* ─── MODAL ──────────────────────────────────────────────────────── */
function renderModalKpis(rows, mode) {
  if (mode === "stock") {
    const totalVal = rows.reduce((t,r) => t + (r.value||0), 0);
    byId("modalKpis").innerHTML = `
      <div class="modal-kpi"><div class="modal-kpi-label">Valor total</div><div class="modal-kpi-value">${escapeHtml(formatCurrency(totalVal))}</div></div>
      <div class="modal-kpi"><div class="modal-kpi-label">Linhas</div><div class="modal-kpi-value">${escapeHtml(formatNumber(rows.length))}</div></div>`;
    return;
  }
  const totalRev  = rows.reduce((t,r) => t + (r.revenue||0), 0);
  const totalTax  = rows.reduce((t,r) => t + calculateRowTaxes(r), 0);
  const taxPct    = totalRev ? (totalTax/totalRev)*100 : null;
  // FIX: compute once per row
  const modalMargins = rows.map(r => calculateMarginForRow(r));
  const withCost  = rows.filter((_, i) => modalMargins[i].value != null);
  const withCostM = modalMargins.filter(m => m.value != null);
  const totalMarg = withCostM.reduce((t, m) => t + m.value, 0);
  const margRev   = withCost.reduce((t,r) => t + r.revenue, 0);
  const margPct   = margRev ? (totalMarg/margRev)*100 : null;
  const notes     = new Set(rows.map(r => r.note)).size;

  byId("modalKpis").innerHTML = [
    { lbl:"Receita",     val:formatCurrency(totalRev) },
    { lbl:"Impostos",    val:`${formatCurrency(totalTax)} (${formatPercent(taxPct)})` },
    { lbl:"Margem",      val:margPct != null ? `${formatPercent(margPct)} · ${formatCurrency(totalMarg)}` : "—" },
    { lbl:"NFs",         val:formatNumber(notes) },
    { lbl:"Linhas",      val:formatNumber(rows.length) },
  ].map(k => `<div class="modal-kpi"><div class="modal-kpi-label">${escapeHtml(k.lbl)}</div><div class="modal-kpi-value">${escapeHtml(k.val)}</div></div>`).join("");
}

function renderModalTable(rows) {
  byId("modalCount").textContent = `${formatNumber(rows.length)} linhas exibidas`;

  if (state.modalMode === "stock") {
    renderTable("modalTable",
      [{ label:"Data" },{ label:"NF" },{ label:"Cliente" },{ label:"UF" },
       { label:"Produto" },{ label:"Tipo" },{ label:"Qtd",numeric:true },
       { label:"Valor",numeric:true },{ label:"CFOP" }],
      rows.map(row => ({ cells:[
        { html:escapeHtml(formatDateBr(row.date)) },
        { html:escapeHtml(row.note) },
        { html:escapeHtml(row.client) },
        { html:escapeHtml(row.uf||"—") },
        { html:escapeHtml(row.item) },
        { html:`<span class="pill info">${escapeHtml(row.typeLabel)}</span>` },
        { html:escapeHtml(formatNumber(row.quantity)), numeric:true },
        { html:escapeHtml(formatCurrency(row.value)), numeric:true },
        { html:escapeHtml(row.cfop||"—") },
      ]}))
    );
    return;
  }

  renderTable("modalTable",
    [{ label:"Data" },{ label:"NF" },{ label:"Cliente" },{ label:"UF" },
     { label:"Produto" },{ label:"Tipo" },
     { label:"Receita",numeric:true },{ label:"Impostos",numeric:true },
     { label:"Custo",numeric:true },{ label:"Margem %",numeric:true },{ label:"Margem R$",numeric:true }],
    rows.map(row => {
      const { value:mVal, pct:mPct } = calculateMarginForRow(row);
      return { cells:[
        { html:escapeHtml(formatDateBr(row.date)) },
        { html:escapeHtml(row.note) },
        { html:escapeHtml(row.client) },
        { html:escapeHtml(row.uf||"—") },
        { html:escapeHtml(row.item) },
        { html:`<span class="pill ${row.type==="devolucao"?"bad":"good"}">${escapeHtml(row.typeLabel)}</span>` },
        { html:escapeHtml(formatCurrency(row.revenue)), numeric:true },
        { html:escapeHtml(formatCurrency(calculateRowTaxes(row))), numeric:true },
        { html:row.cost == null ? "—" : escapeHtml(formatCurrency(row.cost)), numeric:true },
        { html:mPct == null ? "—" : `<span class="pill ${mPct>=0?"good":"bad"}">${escapeHtml(formatPercent(mPct))}</span>`, numeric:true },
        { html:mVal == null ? "—" : escapeHtml(formatCurrency(mVal)), numeric:true },
      ]};
    })
  );
}

function filterModalRows() {
  const query = normalizeText(byId("modalSearch").value);
  const rows  = !query ? state.modalRows
    : state.modalRows.filter(row =>
        [row.date, row.note, row.client, row.uf, row.item, row.typeLabel, row.cfop]
          .map(normalizeText).some(v => v.includes(query))
      );
  renderModalTable(rows);
}

function openModal(title, subtitle, rows, mode) {
  state.modalRows = rows;
  state.modalMode = mode;
  byId("modalTitle").textContent    = title;
  byId("modalSubtitle").textContent = subtitle;
  byId("modalSearch").value         = "";
  renderModalKpis(rows, mode);
  renderModalTable(rows);
  byId("modalBackdrop").classList.remove("hidden");
}

function closeModal() {
  byId("modalBackdrop").classList.add("hidden");
  state.modalRows = [];
}

/* ─── CSV EXPORT ─────────────────────────────────────────────────── */
function downloadCsv(rows, filename) {
  if (!rows.length) return;
  const cols = Object.keys(rows[0]);
  const lines = [
    cols.join(";"),
    ...rows.map(r => cols.map(c => JSON.stringify(r[c] ?? "")).join(";")),
  ];
  const blob = new Blob(["\uFEFF" + lines.join("\r\n")], { type:"text/csv;charset=utf-8;" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href     = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

/* ─── MAIN RENDER ORCHESTRATOR ───────────────────────────────────── */
function renderDashboard() {
  if (!state.source) return;
  state.metrics        = getMetrics(state.filteredSales, state.filteredStock);
  state.compareMetrics = byId("compareToggle").checked
    ? getMetrics(state.compareSales, state.compareStock)
    : null;

  summarizeSource(state.source);
  renderFilterInfo(state.metrics, state.compareMetrics);
  renderSummaryCards(state.metrics, state.compareMetrics);
  renderMonthlyChart(state.metrics);
  renderTaxChart(state.metrics);
  renderAnalysisPanel(state.metrics, state.compareMetrics);
  renderClientsTable(state.metrics);
  renderProductsTable(state.metrics);
  renderPositiveSkuTable(state.metrics);
  renderNegativeMarginTable(state.metrics);
  renderBadClientsTable(state.metrics);
  renderTaxOutlierTable(state.metrics);
  renderStockTypeChart("remessaChart", state.metrics.remessaMonthly, "remessa");
  renderStockTypeChart("baixaChart",   state.metrics.baixaMonthly,   "baixa_estoque");
  renderStockTypeTable("remessaTable", state.metrics.remessaProducts, "remessa");
  renderStockTypeTable("baixaTable",   state.metrics.baixaCfops,      "baixa_estoque");

  byId("emptyState").classList.add("hidden");
  byId("dashboardContent").classList.remove("hidden");
}

/* ─── FILTERS ────────────────────────────────────────────────────── */
function filterRowsByDate(rows, from, to) {
  return rows.filter(r => (!from || r.date >= from) && (!to || r.date <= to));
}

function applyFilters() {
  if (!state.source) return;
  const from = byId("dateFrom").value, to = byId("dateTo").value;
  state.filteredSales = filterRowsByDate(state.salesRows, from, to);
  state.filteredStock = filterRowsByDate(state.stockRows, from, to);
  if (byId("compareToggle").checked) {
    state.compareSales = filterRowsByDate(state.salesRows, byId("compareFrom").value, byId("compareTo").value);
    state.compareStock = filterRowsByDate(state.stockRows, byId("compareFrom").value, byId("compareTo").value);
  } else {
    state.compareSales = [];
    state.compareStock = [];
  }
  invalidateMetricsCache();
  renderDashboard();
}

function resetFilters() {
  if (!state.source) return;
  byId("dateFrom").value = state.source.coverageFrom;
  byId("dateTo").value   = state.source.coverageTo;
  if (byId("compareToggle").checked) setDefaultCompareRange();
  applyFilters();
}

/* ─── COMPARE RANGE ──────────────────────────────────────────────── */
function getCoverageDates(rowsA, rowsB) {
  const dates = [...rowsA,...rowsB].map(r => r.date).filter(Boolean).sort();
  return { from:dates[0]||"", to:dates[dates.length-1]||"" };
}

function setDefaultCompareRange() {
  const dateFrom = byId("dateFrom").value, dateTo = byId("dateTo").value;
  if (!dateFrom || !dateTo) return;
  const duration = diffDaysInclusive(dateFrom, dateTo);
  if (!duration) return;
  const compareTo   = addDaysToIso(dateFrom, -1);
  const compareFrom = addDaysToIso(compareTo, -(duration-1));
  byId("compareFrom").value = compareFrom;
  byId("compareTo").value   = compareTo;
  state.compareDirty = false;
}

function updateCompareInputsAvailability() {
  const enabled = !!state.source && byId("compareToggle").checked;
  byId("compareFrom").disabled = !enabled;
  byId("compareTo").disabled   = !enabled;
}

/* ─── SOURCE LOAD ────────────────────────────────────────────────── */
function finalizeSourceLoad({
  fileName,
  format,
  sheetName,
  totalRows,
  ignoredRows = 0,
  unclassifiedRows = 0,
  salesRows,
  stockRows,
  configOverride = null,
  cfopOverrides = null,
  versionMeta = null,
  workingVersionParentId = "",
}) {
  if (!salesRows.length && !stockRows.length)
    throw new Error("Nenhuma linha válida encontrada para o dashboard.");

  const coverage = getCoverageDates(salesRows, stockRows);
  state.source = {
    fileName, format, sheetName, totalRows, ignoredRows, unclassifiedRows,
    coverageFrom:coverage.from, coverageTo:coverage.to, salesRows, stockRows,
  };

  // Base rows for CFOP registry: include unclassified stubs so their CFOPs appear in the configurator
  const unclassifiedStubs = salesRows._unclassified || [];
  state.baseRows = [
    ...salesRows.map(r => toBaseRow(r,"sales")),
    ...stockRows.map(r => toBaseRow(r,"stock")),
    ...unclassifiedStubs.map(r => ({ ...r, dataset:"unclassified", baseValue:r.revenue||0, originalType:"ignorar", analysisType:"ignorar" })),
  ];
  buildCfopRegistry(state.baseRows);

  const effectiveConfig = configOverride || state.globalConfig?.config || createDefaultConfig();
  applyConfigState(effectiveConfig, { rebuild: false, rerender: false });
  rebuildRowsFromCfopRegistry();

  const effectiveOverrides = cfopOverrides ?? state.globalConfig?.cfopOverrides ?? [];
  applyCfopOverrides(effectiveOverrides, { resetMissing: true, rebuild: true, rerender: false });

  state.currentVersionMeta = versionMeta ? normalizeVersionRecord(versionMeta) : null;
  state.workingVersionParentId = workingVersionParentId || "";
  updateWorkingVersionSummary();

  byId("dateFrom").disabled      = false;
  byId("dateTo").disabled        = false;
  byId("compareToggle").disabled = false;
  byId("resetFilters").disabled  = false;
  byId("dateFrom").value         = coverage.from;
  byId("dateTo").value           = coverage.to;

  summarizeSource(state.source);
  updateCompareInputsAvailability();
  renderCfopRegistryMeta();
  renderCfopConfigTable();
  if (byId("compareToggle").checked) setDefaultCompareRange();
  applyFilters();
}

/* ─── FILE HANDLERS ──────────────────────────────────────────────── */
function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try { resolve(XLSX.read(e.target.result, { type:"array", cellDates:true })); }
      catch(err) { reject(err); }
    };
    reader.onerror = () => reject(reader.error || new Error("Falha ao ler o arquivo."));
    reader.readAsArrayBuffer(file);
  });
}

async function handleJsonFile(file) {
  const text    = await file.text();
  const payload = JSON.parse(text);
  if (!payload || typeof payload !== "object") throw new Error("JSON inválido.");
  const salesSource = Array.isArray(payload.sales) ? payload.sales : [];
  const stockSource = Array.isArray(payload.stock) ? payload.stock : [];
  const salesRows   = salesSource.map(normalizeJsonSalesRow).filter(Boolean);
  const stockRows   = stockSource.map(normalizeJsonStockRow).filter(Boolean);
  finalizeSourceLoad({
    fileName:file.name, format:"JSON",
    sheetName:payload.meta?.version ? `Payload ${payload.meta.version}` : "Payload JSON",
    totalRows:salesSource.length + stockSource.length,
    ignoredRows:salesSource.length + stockSource.length - salesRows.length - stockRows.length,
    salesRows, stockRows,
  });
}

async function handleWorkbookFile(file) {
  const workbook  = await readWorkbook(file);
  const sheetName = workbook.SheetNames.find(n => normalizeText(n).includes("consolidado")) || workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) throw new Error('Aba "Consolidado" não encontrada.');

  // Yield one frame to the UI so the loader renders before the heavy loop
  await new Promise(resolve => setTimeout(resolve, 16));

  const rawRows   = XLSX.utils.sheet_to_json(worksheet, { defval:null, raw:true });
  const salesRows = [], stockRows = [];
  let ignoredRows = 0, unclassifiedRows = 0;

  rawRows.forEach((rawRow) => {
    const built = buildRow(rawRow);
    if (!built) { ignoredRows++; return; }
    if (built.dataset === "sales")        salesRows.push(built.row);
    else if (built.dataset === "stock")   stockRows.push(built.row);
    else if (built.dataset === "unclassified") {
      // Preserve the CFOP in the registry so analyst can reclassify,
      // but do NOT add to salesRows or stockRows
      unclassifiedRows++;
      // We still need to register the CFOP — done via buildCfopRegistry on baseRows.
      // Push a minimal base-compatible stub for registry purposes only.
      salesRows._unclassified = salesRows._unclassified || [];
      salesRows._unclassified.push(built.row);
    }
  });

  finalizeSourceLoad({
    fileName:file.name, format:"Excel", sheetName,
    totalRows:rawRows.length, ignoredRows,
    unclassifiedRows,
    salesRows, stockRows,
  });
}

async function handleInputFile(file) {
  if (!file) return;
  showLoader();
  try {
    setStatus(`Processando ${file.name}…`, "info");
    resetWorkingVersionContext();
    if (file.name.toLowerCase().endsWith(".json")) await handleJsonFile(file);
    else await handleWorkbookFile(file);
    setStatus(`Base carregada: ${file.name}`, "success");
    setActivePage("config");
  } catch (err) {
    console.error(err);
    setStatus(`Erro: ${err.message}`, "error");
  } finally {
    hideLoader();
  }
}

/* ─── CLIENT NAME ────────────────────────────────────────────────── */
function applyClientName(name) {
  const sanitized = String(name||"").trim() || DEFAULT_CLIENT_NAME;
  state.clientName = sanitized;
  byId("clientNameInput").value     = sanitized;
  byId("clientNameBadge").textContent = sanitized;
  try { window.localStorage.setItem(LOCAL_STORAGE_KEYS.clientName, sanitized); } catch (_) {}
}

function hydrateClientName() {
  try {
    const saved = window.localStorage.getItem(LOCAL_STORAGE_KEYS.clientName);
    applyClientName(saved || DEFAULT_CLIENT_NAME);
  } catch (_) { applyClientName(DEFAULT_CLIENT_NAME); }
}

function apiIsConfigured() {
  return Boolean(String(state.apiBaseUrl || "").trim());
}

function applyApiBaseUrl(url, { persist = true } = {}) {
  const normalized = String(url || "").trim();
  state.apiBaseUrl = normalized || DEFAULT_API_BASE_URL;
  byId("apiBaseUrlInput").value = state.apiBaseUrl;
  if (persist) {
    try { window.localStorage.setItem(LOCAL_STORAGE_KEYS.apiBaseUrl, state.apiBaseUrl); } catch (_) {}
  }
  updateApiConfigStatus();
}

function hydrateApiBaseUrl() {
  try {
    applyApiBaseUrl(window.localStorage.getItem(LOCAL_STORAGE_KEYS.apiBaseUrl) || DEFAULT_API_BASE_URL, { persist: false });
  } catch (_) {
    applyApiBaseUrl(DEFAULT_API_BASE_URL, { persist: false });
  }
}

function updateApiConfigStatus() {
  const node = byId("apiConfigStatus");
  if (!node) return;
  node.textContent = apiIsConfigured()
    ? "Apps Script configurado. Os snapshots e padrões globais podem ser lidos e gravados no Google Sheets."
    : "Sem URL configurada. O dashboard continua funcionando localmente, mas snapshots no Google ficam desabilitados.";
}

function readLocalGlobalConfig() {
  try {
    const raw = window.localStorage.getItem(LOCAL_STORAGE_KEYS.globalConfig);
    if (!raw) return createEmptyGlobalConfig();
    const parsed = JSON.parse(raw);
    return {
      config: mergeConfig(parsed?.config),
      cfopOverrides: normalizeCfopOverrides(parsed?.cfopOverrides),
    };
  } catch (_) {
    return createEmptyGlobalConfig();
  }
}

function writeLocalGlobalConfig(payload) {
  try {
    window.localStorage.setItem(LOCAL_STORAGE_KEYS.globalConfig, JSON.stringify(payload));
  } catch (_) {}
}

function readLocalVersions() {
  try {
    const raw = window.localStorage.getItem(LOCAL_STORAGE_KEYS.localVersions);
    return raw ? JSON.parse(raw) : [];
  } catch (_) {
    return [];
  }
}

function writeLocalVersions(versions) {
  try {
    window.localStorage.setItem(LOCAL_STORAGE_KEYS.localVersions, JSON.stringify(versions));
  } catch (_) {}
}

async function callApi(action, { method = "GET", payload = null } = {}) {
  if (!apiIsConfigured()) throw new Error("URL do Apps Script não configurada.");
  const base = state.apiBaseUrl;
  const url = method === "GET"
    ? `${base}${base.includes("?") ? "&" : "?"}action=${encodeURIComponent(action)}${payload ? `&payload=${encodeURIComponent(JSON.stringify(payload))}` : ""}`
    : base;

  const response = await fetch(url, {
    method,
    headers: method === "POST" ? { "Content-Type": "text/plain;charset=utf-8" } : undefined,
    body: method === "POST" ? JSON.stringify({ action, ...payload }) : undefined,
  });

  const text = await response.text();
  let data = null;
  try { data = text ? JSON.parse(text) : null; } catch (_) { data = { ok: false, error: text || "Resposta inválida do Apps Script." }; }

  if (!response.ok || data?.ok === false) {
    throw new Error(data?.error || `Falha na integração (${response.status}).`);
  }
  return data;
}

function buildGlobalConfigPayload() {
  return {
    config: cloneJson(state.config),
    cfopOverrides: getCfopOverrides(),
  };
}

function serializeSalesRows(rows = []) {
  return rows.map((row) => ({
    date: row.date,
    note: row.note,
    client: row.client,
    uf: row.uf,
    item: row.item,
    quantity: row.quantity,
    revenue: row.revenue,
    icms: row.icms,
    pis: row.pis,
    cofins: row.cofins,
    ipi: row.ipi,
    cost: row.cost,
    margin_value: row.marginValue,
    margin_pct: row.marginPct,
    cfop: row.cfop,
    type: row.type,
  }));
}

function serializeStockRows(rows = []) {
  return rows.map((row) => ({
    date: row.date,
    note: row.note,
    client: row.client,
    uf: row.uf,
    item: row.item,
    quantity: row.quantity,
    value: row.value,
    cfop: row.cfop,
    type: row.type,
  }));
}

function buildSnapshotPayload() {
  return {
    clientName: state.clientName,
    sourceFileName: state.source?.fileName || "upload_local",
    sourceFormat: state.source?.format || "manual",
    sheetName: state.source?.sheetName || "Consolidado",
    summary: {
      totalRows: state.source?.totalRows || 0,
      ignoredRows: state.source?.ignoredRows || 0,
      unclassifiedRows: state.source?.unclassifiedRows || 0,
      coverageFrom: state.source?.coverageFrom || "",
      coverageTo: state.source?.coverageTo || "",
      salesCount: state.salesRows.length,
      stockCount: state.stockRows.length,
    },
    config: cloneJson(state.config),
    cfopOverrides: getCfopOverrides(),
    sales: serializeSalesRows(state.salesRows),
    stock: serializeStockRows(state.stockRows),
    parentVersionId: state.currentVersionMeta?.versionId || state.workingVersionParentId || "",
    meta: {
      sourceMode: state.currentVersionMeta ? "snapshot" : "upload",
    },
  };
}

function normalizeVersionRecord(record) {
  if (!record) return null;
  return {
    versionId: record.versionId ?? record.version_id ?? "",
    createdAt: record.createdAt ?? record.created_at ?? "",
    clientName: record.clientName ?? record.client_name ?? state.clientName,
    sourceFileName: record.sourceFileName ?? record.source_file_name ?? "",
    sourceFormat: record.sourceFormat ?? record.source_format ?? "",
    coverageFrom: record.coverageFrom ?? record.coverage_from ?? "",
    coverageTo: record.coverageTo ?? record.coverage_to ?? "",
    salesCount: Number(record.salesCount ?? record.sales_count ?? 0),
    stockCount: Number(record.stockCount ?? record.stock_count ?? 0),
    parentVersionId: record.parentVersionId ?? record.parent_version_id ?? "",
    hash: record.hash ?? "",
    payload: record.payload ?? null,
  };
}

function updateWorkingVersionSummary() {
  const node = byId("workingVersionSummary");
  const detail = byId("versionDetailCard");
  if (!node || !detail) return;

  if (state.currentVersionMeta?.versionId) {
    const coverage = formatRangeBr(state.currentVersionMeta.coverageFrom, state.currentVersionMeta.coverageTo, { shortYear: false });
    node.className = "callout callout-neutral";
    node.innerHTML = `<strong>Versão carregada:</strong> ${escapeHtml(state.currentVersionMeta.versionId)}<br><span>${escapeHtml(coverage)} · ${escapeHtml(formatTimestamp(state.currentVersionMeta.createdAt))}</span>`;
    detail.className = "callout callout-neutral";
    detail.innerHTML = `<strong>${escapeHtml(state.currentVersionMeta.versionId)}</strong><br><span>${escapeHtml(state.currentVersionMeta.sourceFileName || "Snapshot")}</span><br><span>${escapeHtml(coverage)}</span>`;
    return;
  }

  if (state.workingVersionParentId) {
    node.className = "callout callout-warn";
    node.innerHTML = `<strong>Rascunho derivado:</strong> ${escapeHtml(state.workingVersionParentId)}<br><span>Ao salvar, uma nova versão será criada a partir deste snapshot.</span>`;
    detail.className = "callout callout-warn";
    detail.innerHTML = `<strong>Rascunho derivado</strong><br><span>Origem: ${escapeHtml(state.workingVersionParentId)}</span>`;
    return;
  }

  node.className = "callout callout-neutral";
  node.textContent = state.source
    ? "Base carregada e pronta para configuração. Salve um snapshot para congelar esta leitura."
    : "Nenhum snapshot salvo nesta sessão.";
  detail.className = "callout callout-neutral";
  detail.textContent = state.source
    ? "Sessão baseada em upload manual. Após revisar as regras, salve um snapshot para criar histórico."
    : "Nenhuma versão carregada nesta sessão.";
}

function markWorkingCopyFromLoadedVersion() {
  if (!state.currentVersionMeta?.versionId) return;
  state.workingVersionParentId = state.currentVersionMeta.versionId;
  state.currentVersionMeta = null;
  updateWorkingVersionSummary();
}

function renderVersionsTable() {
  const versions = [...state.versions]
    .map(normalizeVersionRecord)
    .filter(Boolean)
    .sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));

  state.versions = versions;

  renderTable(
    "versionsTable",
    [
      { label: "Versão" },
      { label: "Criada em" },
      { label: "Cobertura" },
      { label: "Base" },
      { label: "Linhas", numeric: true },
      { label: "Ações" },
    ],
    versions.map((version) => ({
      cells: [
        { html: `<strong>${escapeHtml(version.versionId || "sem-id")}</strong>` },
        { html: escapeHtml(formatTimestamp(version.createdAt)) },
        { html: escapeHtml(formatRangeBr(version.coverageFrom, version.coverageTo, { shortYear: false })) },
        { html: escapeHtml(version.sourceFileName || version.sourceFormat || "Snapshot") },
        { html: escapeHtml(formatNumber((version.salesCount || 0) + (version.stockCount || 0))), numeric: true },
        { html: `<div class="table-actions"><button class="table-action" type="button" data-version-action="open" data-version-id="${escapeHtml(version.versionId)}">Abrir</button><button class="table-action secondary" type="button" data-version-action="duplicate" data-version-id="${escapeHtml(version.versionId)}">Duplicar</button></div>` },
      ],
    }))
  );

  byId("versionsTable").querySelectorAll("[data-version-action]").forEach((button) => {
    button.addEventListener("click", async (event) => {
      event.stopPropagation();
      const versionId = event.currentTarget.dataset.versionId;
      const action = event.currentTarget.dataset.versionAction;
      if (action === "open") {
        await openVersion(versionId);
        return;
      }
      await duplicateVersion(versionId);
    });
  });

  const status = byId("versionsStatus");
  if (!status) return;
  status.textContent = versions.length
    ? `${formatNumber(versions.length)} versão(ões) disponíveis${apiIsConfigured() ? " no Google" : " em armazenamento local"}.`
    : apiIsConfigured()
      ? "Nenhum snapshot encontrado no índice do Google Sheets."
      : "Nenhum snapshot local encontrado. Configure o Apps Script para usar histórico compartilhado.";
}

function resetWorkingVersionContext() {
  state.currentVersionMeta = null;
  state.workingVersionParentId = null;
  updateWorkingVersionSummary();
}

async function loadGlobalConfig({ silent = false, applyToSession = false } = {}) {
  try {
    let payload = null;
    if (apiIsConfigured()) {
      const response = await callApi("get_global_config");
      payload = response?.data || response;
      if (payload?.config || payload?.cfopOverrides) writeLocalGlobalConfig(payload);
    } else {
      payload = readLocalGlobalConfig();
    }

    state.globalConfig = {
      config: mergeConfig(payload?.config),
      cfopOverrides: normalizeCfopOverrides(payload?.cfopOverrides),
    };

    if (applyToSession) {
      applyConfigState(state.globalConfig.config, { rebuild: false, rerender: false });
      applyCfopOverrides(state.globalConfig.cfopOverrides, { resetMissing: true, rebuild: true, rerender: !!state.source });
    }

    if (!silent) setStatus(apiIsConfigured() ? "Padrão global carregado do Google." : "Padrão global carregado do armazenamento local.", "success");
    updateWorkingVersionSummary();
    return state.globalConfig;
  } catch (error) {
    if (!silent) setStatus(`Erro ao carregar padrão global: ${error.message}`, "error");
    return createEmptyGlobalConfig();
  }
}

async function saveGlobalConfig() {
  const payload = buildGlobalConfigPayload();
  state.globalConfig = {
    config: mergeConfig(payload.config),
    cfopOverrides: normalizeCfopOverrides(payload.cfopOverrides),
  };
  writeLocalGlobalConfig(state.globalConfig);

  try {
    if (apiIsConfigured()) {
      await callApi("save_global_config", { method: "POST", payload });
      setStatus("Padrão global salvo no Google Sheets.", "success");
    } else {
      setStatus("Padrão global salvo localmente. Configure o Apps Script para compartilhar com o time.", "info");
    }
  } catch (error) {
    setStatus(`Erro ao salvar padrão global: ${error.message}`, "error");
  }
}

function buildLocalVersionRecord(snapshotPayload) {
  const createdAt = new Date().toISOString();
  const versionId = `local-${createdAt.replace(/\D/g, "").slice(0, 14)}`;
  return normalizeVersionRecord({
    versionId,
    createdAt,
    clientName: snapshotPayload.clientName,
    sourceFileName: snapshotPayload.sourceFileName,
    sourceFormat: snapshotPayload.sourceFormat,
    coverageFrom: snapshotPayload.summary.coverageFrom,
    coverageTo: snapshotPayload.summary.coverageTo,
    salesCount: snapshotPayload.summary.salesCount,
    stockCount: snapshotPayload.summary.stockCount,
    parentVersionId: snapshotPayload.parentVersionId,
    payload: {
      meta: {
        versionId,
        createdAt,
        sourceFileName: snapshotPayload.sourceFileName,
        sourceFormat: snapshotPayload.sourceFormat,
        sheetName: snapshotPayload.sheetName,
        coverageFrom: snapshotPayload.summary.coverageFrom,
        coverageTo: snapshotPayload.summary.coverageTo,
        parentVersionId: snapshotPayload.parentVersionId,
      },
      config: snapshotPayload.config,
      cfopOverrides: snapshotPayload.cfopOverrides,
      sales: snapshotPayload.sales,
      stock: snapshotPayload.stock,
    },
  });
}

async function saveSnapshot() {
  if (!state.source) {
    setStatus("Carregue uma base antes de salvar um snapshot.", "error");
    setActivePage("base");
    return;
  }

  const snapshotPayload = buildSnapshotPayload();

  try {
    let version = null;
    if (apiIsConfigured()) {
      const response = await callApi("save_snapshot", { method: "POST", payload: snapshotPayload });
      version = normalizeVersionRecord(response?.version || response?.data?.version || response?.meta || snapshotPayload.meta);
      setStatus(`Snapshot salvo: ${version?.versionId || "nova versão"}.`, "success");
    } else {
      const localVersions = readLocalVersions();
      const localVersion = buildLocalVersionRecord(snapshotPayload);
      localVersions.unshift(localVersion);
      writeLocalVersions(localVersions);
      version = localVersion;
      setStatus("Snapshot salvo localmente. Configure o Apps Script para compartilhar o histórico.", "info");
    }

    if (version) {
      state.currentVersionMeta = version;
      state.workingVersionParentId = "";
      await refreshVersions({ silent: true });
      updateWorkingVersionSummary();
      setActivePage("versions");
    }
  } catch (error) {
    setStatus(`Erro ao salvar snapshot: ${error.message}`, "error");
  }
}

async function refreshVersions({ silent = false } = {}) {
  try {
    let versions = readLocalVersions();
    if (apiIsConfigured()) {
      const response = await callApi("list_versions");
      versions = response?.versions || response?.data?.versions || [];
    }

    state.versions = (versions || []).map(normalizeVersionRecord).filter(Boolean);
    state.versionsLoadedAt = new Date().toISOString();
    renderVersionsTable();
    if (!silent) setStatus(apiIsConfigured() ? "Histórico atualizado do Google Sheets." : "Histórico local atualizado.", "success");
  } catch (error) {
    state.versions = [];
    renderVersionsTable();
    if (!silent) setStatus(`Erro ao carregar versões: ${error.message}`, "error");
  }
}

async function getVersionPayload(versionId) {
  if (apiIsConfigured()) {
    const response = await callApi("get_version", { payload: { versionId } });
    return response?.data || response;
  }
  const version = readLocalVersions().map(normalizeVersionRecord).find((item) => item.versionId === versionId);
  if (!version?.payload) throw new Error("Versão local não encontrada.");
  return version.payload;
}

async function openVersion(versionId) {
  try {
    const payload = await getVersionPayload(versionId);
    hydrateSnapshotPayload(payload, { duplicate: false });
    setActivePage("analysis");
    setStatus(`Versão ${versionId} carregada.`, "success");
  } catch (error) {
    setStatus(`Erro ao abrir versão: ${error.message}`, "error");
  }
}

async function duplicateVersion(versionId) {
  try {
    const payload = await getVersionPayload(versionId);
    hydrateSnapshotPayload(payload, { duplicate: true });
    setActivePage("config");
    setStatus(`Versão ${versionId} carregada como nova revisão.`, "success");
  } catch (error) {
    setStatus(`Erro ao duplicar versão: ${error.message}`, "error");
  }
}

function hydrateSnapshotPayload(payload, { duplicate = false } = {}) {
  const salesSource = Array.isArray(payload?.sales) ? payload.sales : [];
  const stockSource = Array.isArray(payload?.stock) ? payload.stock : [];
  const salesRows = salesSource.map(normalizeJsonSalesRow).filter(Boolean);
  const stockRows = stockSource.map(normalizeJsonStockRow).filter(Boolean);
  const meta = normalizeVersionRecord(payload?.meta || {});

  finalizeSourceLoad({
    fileName: meta?.sourceFileName || "snapshot.json",
    format: meta?.sourceFormat || "Snapshot",
    sheetName: payload?.meta?.sheetName || "Snapshot",
    totalRows: salesSource.length + stockSource.length,
    ignoredRows: 0,
    unclassifiedRows: 0,
    salesRows,
    stockRows,
    configOverride: payload?.config || createDefaultConfig(),
    cfopOverrides: payload?.cfopOverrides || [],
    versionMeta: duplicate ? null : meta,
    workingVersionParentId: duplicate ? meta?.versionId || "" : "",
  });
}

/* ─── PAGE NAV ───────────────────────────────────────────────────── */
function setActivePage(page) {
  state.activePage = page;
  byId("pageBase").classList.toggle("hidden", page !== "base");
  byId("pageConfig").classList.toggle("hidden", page !== "config");
  byId("pageVersions").classList.toggle("hidden", page !== "versions");
  byId("pageAnalysis").classList.toggle("hidden", page !== "analysis");
  document.querySelectorAll(".rail-nav-btn").forEach(btn =>
    btn.classList.toggle("active", btn.dataset.page === page)
  );
  const meta = PAGE_META[page] || PAGE_META.base;
  byId("pageEyebrow").textContent = meta.eyebrow;
  byId("pageTitle").textContent = meta.title;
}

/* ─── SYNC CONFIG ────────────────────────────────────────────────── */
function syncConfigFromInputs() {
  state.config.revenue.sales          = byId("configRevenueSales").checked;
  state.config.revenue.returns        = byId("configRevenueReturns").checked;
  state.config.taxes.icms             = byId("configTaxIcms").checked;
  state.config.taxes.pis              = byId("configTaxPis").checked;
  state.config.taxes.cofins           = byId("configTaxCofins").checked;
  state.config.taxes.ipi              = byId("configTaxIpi").checked;
  state.config.stock.remessa          = byId("configStockRemessa").checked;
  state.config.stock.baixa            = byId("configStockBaixa").checked;
  state.config.margin.deductTaxes     = byId("configMarginDeductTaxes").checked;
  state.config.margin.requireCost     = byId("configMarginRequireCost").checked;
  invalidateMetricsCache();
  buildAssumptions();
}

/* ─── BINDINGS ───────────────────────────────────────────────────── */
function bindUploads() {
  [{ trigger:"jsonTrigger", input:"jsonInput" },{ trigger:"excelTrigger", input:"excelInput" }]
    .forEach(({ trigger, input }) => {
      byId(trigger).addEventListener("click", () => byId(input).click());
      byId(input).addEventListener("change", (e) => {
        handleInputFile(e.target.files?.[0]);
        e.target.value = "";
      });
    });
}

function bindWorkspaceActions() {
  byId("saveApiBaseUrlBtn").addEventListener("click", async () => {
    applyApiBaseUrl(byId("apiBaseUrlInput").value);
    if (apiIsConfigured()) {
      await loadGlobalConfig({ silent: true });
      await refreshVersions({ silent: true });
      setStatus("Integração Apps Script configurada.", "success");
    } else {
      setStatus("URL do Apps Script removida. O dashboard segue em modo local.", "info");
    }
  });

  byId("apiBaseUrlInput").addEventListener("keydown", async (event) => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    byId("saveApiBaseUrlBtn").click();
  });

  byId("loadGlobalConfigBtn").addEventListener("click", async () => {
    if (state.source) markWorkingCopyFromLoadedVersion();
    await loadGlobalConfig({ applyToSession: !!state.source });
  });

  byId("saveGlobalConfigBtn").addEventListener("click", async () => {
    await saveGlobalConfig();
  });

  byId("resetGlobalConfigBtn").addEventListener("click", async () => {
    markWorkingCopyFromLoadedVersion();
    await loadGlobalConfig({ applyToSession: true });
  });

  byId("saveSnapshotBtn").addEventListener("click", async () => {
    await saveSnapshot();
  });

  byId("refreshVersionsBtn").addEventListener("click", async () => {
    await refreshVersions();
  });

  byId("goToConfigBtn").addEventListener("click", () => setActivePage("config"));
  byId("goToAnalysisBtn").addEventListener("click", () => {
    if (!state.source) {
      setStatus("Carregue uma base antes de abrir as análises.", "error");
      return;
    }
    setActivePage("analysis");
  });
  byId("openBaseFromVersionsBtn").addEventListener("click", () => setActivePage("base"));
}

function bindFilters() {
  ["dateFrom","dateTo"].forEach(id => {
    byId(id).addEventListener("change", () => {
      if (byId("compareToggle").checked && !state.compareDirty) setDefaultCompareRange();
      applyFilters();
    });
  });
  byId("compareToggle").addEventListener("change", () => {
    updateCompareInputsAvailability();
    if (byId("compareToggle").checked && !state.compareDirty) setDefaultCompareRange();
    applyFilters();
  });
  ["compareFrom","compareTo"].forEach(id => {
    byId(id).addEventListener("change", () => { state.compareDirty = true; applyFilters(); });
  });
  byId("resetFilters").addEventListener("click", resetFilters);
}

function bindRankingControls() {
  // FIX: ranking selects only re-render the affected table, not recalculate metrics
  const tableMap = {
    topClientsLimit:      () => renderClientsTable(state.metrics),
    topClientsSort:       () => renderClientsTable(state.metrics),
    topProductsLimit:     () => renderProductsTable(state.metrics),
    topProductsSort:      () => renderProductsTable(state.metrics),
    topPositiveSkuLimit:  () => renderPositiveSkuTable(state.metrics),
    topNegativeSkuLimit:  () => renderNegativeMarginTable(state.metrics),
    topBadClientsLimit:   () => renderBadClientsTable(state.metrics),
    topTaxLimit:          () => renderTaxOutlierTable(state.metrics),
  };
  Object.entries(tableMap).forEach(([id, fn]) => {
    byId(id).addEventListener("change", () => { if (state.metrics) fn(); });
  });
}

function bindModal() {
  byId("closeModal").addEventListener("click", closeModal);
  byId("modalBackdrop").addEventListener("click", (e) => { if (e.target === e.currentTarget) closeModal(); });
  byId("modalSearch").addEventListener("input", filterModalRows);
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeModal(); });
}

function bindConfig() {
  ["configRevenueSales","configRevenueReturns","configTaxIcms","configTaxPis",
   "configTaxCofins","configTaxIpi","configStockRemessa","configStockBaixa",
   "configMarginDeductTaxes","configMarginRequireCost"].forEach(id => {
    byId(id).addEventListener("change", () => {
      markWorkingCopyFromLoadedVersion();
      syncConfigFromInputs();
      if (state.baseRows.length) rebuildRowsFromCfopRegistry();
      if (state.source) applyFilters();
    });
  });
}

function bindCfopConfig() {
  byId("cfopSearch").addEventListener("input", renderCfopConfigTable);
  byId("resetCfopConfig").addEventListener("click", () => {
    markWorkingCopyFromLoadedVersion();
    state.cfopRegistry.forEach(meta => { meta.customType = null; meta.analysisType = meta.officialType; });
    rebuildRowsFromCfopRegistry();
    renderCfopRegistryMeta();
    renderCfopConfigTable();
    if (state.source) applyFilters();
  });
}

function bindExport() {
  byId("exportCsvBtn")?.addEventListener("click", () => {
    if (!state.filteredSales.length) return;
    downloadCsv(
      state.filteredSales.filter(isRevenueRowIncluded).map(r => {
        const mg = calculateMarginForRow(r);  // compute once
        return {
          data:           r.date,
          nota:           r.note,
          cliente:        r.client,
          uf:             r.uf,
          produto:        r.item,
          tipo:           r.typeLabel,
          cfop:           r.cfop,
          quantidade:     r.quantity,
          receita:        r.revenue,
          icms:           r.icms,
          pis:            r.pis,
          cofins:         r.cofins,
          ipi:            r.ipi,
          impostos_total: calculateRowTaxes(r),
          custo:          r.cost ?? "",
          margem_r:       mg.value ?? "",
          margem_pct:     mg.pct   ?? "",
        };
      }),
      `dashboard_${state.source?.fileName?.replace(/\.[^.]+$/,"")}_${byId("dateFrom").value}_${byId("dateTo").value}.csv`
    );
  });
}

/* ─── INIT ───────────────────────────────────────────────────────── */
async function init() {
  hydrateClientName();
  hydrateApiBaseUrl();
  syncConfigFromInputs();
  initChartTheme();
  setActivePage("base");
  bindUploads();
  bindWorkspaceActions();
  bindFilters();
  byId("clientNameInput").addEventListener("input", (e) => applyClientName(e.target.value));
  bindConfig();
  bindPageNavigation();
  bindCfopConfig();
  bindRankingControls();
  bindModal();
  bindExport();
  buildAssumptions();
  renderCfopRegistryMeta();
  renderCfopConfigTable();
  renderVersionsTable();
  updateWorkingVersionSummary();
  await loadGlobalConfig({ silent: true });
  await refreshVersions({ silent: true });
}

function bindPageNavigation() {
  document.querySelectorAll(".rail-nav-btn").forEach(btn => {
    btn.addEventListener("click", () => setActivePage(btn.dataset.page));
  });
}

window.addEventListener("load", init);
