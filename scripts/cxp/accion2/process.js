const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");
const ExcelJS = require("exceljs");
const { XMLBuilder, XMLParser } = require("fast-xml-parser");
const XLSX = require("xlsx");

const XML_PARSER = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  allowBooleanAttributes: true,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
});

const XML_BUILDER = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  suppressBooleanAttributes: false,
  format: false,
});

const INPUT_TXT = process.argv[2] || "ACCION2.txt";
const OUTPUT_XLSX = process.argv[3] || "retenciones_proveedor.xlsx";
const DEFAULT_TEMPLATE_XLSX = path.join(__dirname, "outputs", "EJEMPLOSAMANO", "ACCION2.xlsx");
const TEMPLATE_XLSX = process.argv[4]
  ? path.resolve(process.cwd(), process.argv[4])
  : DEFAULT_TEMPLATE_XLSX;

const EXPECTED_COLUMNS = [
  "NUM RT",
  "PROVEEDOR",
  "FECHA",
  "FECHA CONT",
  "TIPO",
  "COD",
  "FACT",
  "%",
  "BASE",
  "RETENCION",
];

const HEADER_ALIASES = {
  NUMRT: "NUM RT",
  NUMERORT: "NUM RT",
  NUMRET: "NUM RT",
  NUMRETENCION: "NUM RT",
  PROVEEDOR: "PROVEEDOR",
  RAZONSOCIAL: "PROVEEDOR",
  NOMBREPROVEEDOR: "PROVEEDOR",
  FECHA: "FECHA",
  FECHADOC: "FECHA CONT",
  FECHADOCUMENTO: "FECHA CONT",
  FECHADOCTO: "FECHA CONT",
  FECHADOCU: "FECHA CONT",
  FECHACONT: "FECHA CONT",
  FECHACONTABLE: "FECHA CONT",
  TIPO: "TIPO",
  COD: "COD",
  CODIGO: "COD",
  TRANS: "COD",
  TRANSACCION: "COD",
  FACT: "FACT",
  FACTURA: "FACT",
  PORCENTAJE: "%",
  "%": "%",
  BASE: "BASE",
  BASERET: "BASE",
  BASEIVA: "BASE",
  VALORRETEN: "RETENCION",
  VALORRETENCION: "RETENCION",
  RETENCION: "RETENCION",
};

const INVALID_XML_CONTROL_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g;

function sanitizeText(value) {
  return String(value || "")
    .replace(INVALID_XML_CONTROL_CHARS, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeHeader(value) {
  return sanitizeText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9%]/g, "");
}

function parseDecimalLike(value) {
  let normalized = sanitizeText(value).replace(/[^\d,.\-]/g, "");
  if (!normalized) {
    return 0;
  }

  const hasComma = normalized.includes(",");
  const hasDot = normalized.includes(".");

  if (hasComma && hasDot) {
    if (normalized.lastIndexOf(".") > normalized.lastIndexOf(",")) {
      normalized = normalized.replace(/,/g, "");
    } else {
      normalized = normalized.replace(/\./g, "").replace(/,/g, ".");
    }
  } else if (hasComma) {
    const commaCount = (normalized.match(/,/g) || []).length;
    if (commaCount > 1) {
      normalized = normalized.replace(/,/g, "");
    } else {
      const [intPart, fracPart = ""] = normalized.split(",");
      if (fracPart.length === 3) {
        normalized = `${intPart}${fracPart}`;
      } else {
        normalized = `${intPart}.${fracPart}`;
      }
    }
  } else if (hasDot) {
    const dotCount = (normalized.match(/\./g) || []).length;
    if (dotCount > 1) {
      const lastDot = normalized.lastIndexOf(".");
      const intPart = normalized.slice(0, lastDot).replace(/\./g, "");
      const fracPart = normalized.slice(lastDot + 1);
      if (fracPart.length === 3) {
        normalized = `${intPart}${fracPart}`;
      } else {
        normalized = `${intPart}.${fracPart}`;
      }
    }
  }

  const number = Number(normalized);
  return Number.isFinite(number) ? number : 0;
}

function parseIntLike(value) {
  const normalized = sanitizeText(value).replace(/[^\d]/g, "");
  if (!normalized) {
    return null;
  }
  const number = Number(normalized);
  return Number.isFinite(number) ? number : null;
}

function parseDateFlexible(value) {
  if (value instanceof Date && Number.isFinite(value.getTime())) {
    return value;
  }

  const clean = sanitizeText(value);
  if (!clean) {
    return null;
  }

  let day;
  let month;
  let year;

  let match = /^(\d{2})[\/\-](\d{2})[\/\-](\d{4})$/.exec(clean);
  if (match) {
    day = Number(match[1]);
    month = Number(match[2]);
    year = Number(match[3]);
  } else {
    match = /^(\d{4})[\/\-](\d{2})[\/\-](\d{2})$/.exec(clean);
    if (!match) {
      return null;
    }
    year = Number(match[1]);
    month = Number(match[2]);
    day = Number(match[3]);
  }

  const date = new Date(year, month - 1, day);
  if (!Number.isFinite(date.getTime())) {
    return null;
  }
  return date;
}

function splitDelimitedLine(line, delimiter) {
  if (delimiter === "\t") {
    return line.split("\t").map((part) => part.trim());
  }

  const result = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    const next = line[i + 1];

    if (char === "\"") {
      if (inQuotes && next === "\"") {
        current += "\"";
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === delimiter && !inQuotes) {
      result.push(current.trim());
      current = "";
      continue;
    }

    current += char;
  }

  result.push(current.trim());
  return result;
}

function mapHeaders(tokens) {
  return tokens.map((token) => {
    const normalized = normalizeHeader(token);
    return HEADER_ALIASES[normalized] || null;
  });
}

function detectHeaderConfig(lines) {
  const delimiters = ["\t", ";", "|", ","];
  let best = null;

  const maxLines = Math.min(lines.length, 60);
  for (let i = 0; i < maxLines; i += 1) {
    const line = sanitizeText(lines[i]);
    if (!line) {
      continue;
    }

    for (const delimiter of delimiters) {
      const tokens = splitDelimitedLine(lines[i], delimiter);
      if (tokens.length < 5) {
        continue;
      }
      const mapped = mapHeaders(tokens);
      const recognized = mapped.filter(Boolean).length;
      if (recognized < 5) {
        continue;
      }

      const score = recognized * 100 + tokens.length;
      if (!best || score > best.score) {
        best = {
          score,
          headerIndex: i,
          delimiter,
          mappedHeaders: mapped,
          tokenCount: tokens.length,
        };
      }
    }
  }

  return best;
}

function parseRowsFromHeader(lines, config) {
  const rows = [];
  for (let i = config.headerIndex + 1; i < lines.length; i += 1) {
    const rawLine = lines[i];
    if (!sanitizeText(rawLine)) {
      continue;
    }

    const tokens = splitDelimitedLine(rawLine, config.delimiter);
    if (tokens.every((token) => !sanitizeText(token))) {
      continue;
    }

    const row = {};
    for (let c = 0; c < config.mappedHeaders.length; c += 1) {
      const header = config.mappedHeaders[c];
      if (!header) {
        continue;
      }
      row[header] = tokens[c] || "";
    }

    const nonEmptyMapped = Object.values(row).some((value) => sanitizeText(value));
    if (nonEmptyMapped) {
      rows.push(row);
    }
  }

  return rows;
}

function parseRowsByPattern(lines) {
  const rows = [];

  const hardPattern = /^(\d+)\s+(.+?)\s+(\d{2}[\/\-]\d{2}[\/\-]\d{4})\s+(\d{2}[\/\-]\d{2}[\/\-]\d{4})\s+(IVA|RENTA)\s+([A-Z0-9]+)\s+([A-Z0-9\-\/]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$/i;

  for (const line of lines) {
    const cleanLine = line.trim();
    if (!cleanLine) {
      continue;
    }

    const match = hardPattern.exec(cleanLine);
    if (match) {
      rows.push({
        "NUM RT": match[1],
        PROVEEDOR: match[2],
        FECHA: match[3],
        "FECHA CONT": match[4],
        TIPO: match[5],
        COD: match[6],
        FACT: match[7],
        "%": match[8],
        BASE: match[9],
        RETENCION: match[10],
      });
      continue;
    }

    const pieces = cleanLine.split(/\s{2,}/).map((part) => part.trim()).filter(Boolean);
    if (pieces.length >= 10) {
      rows.push({
        "NUM RT": pieces[0],
        PROVEEDOR: pieces[1],
        FECHA: pieces[2],
        "FECHA CONT": pieces[3],
        TIPO: pieces[4],
        COD: pieces[5],
        FACT: pieces[6],
        "%": pieces[7],
        BASE: pieces[8],
        RETENCION: pieces[9],
      });
    }
  }

  return rows;
}

function parseRowsFromEmbeddedRetReport(lines) {
  const rows = [];
  let matched = 0;

  for (const line of lines) {
    if (!line || line.indexOf("\t") === -1) {
      continue;
    }

    const tokens = line
      .split("\t")
      .map((token) => sanitizeText(token));

    if (tokens.length < 20) {
      continue;
    }
    if (normalizeHeader(tokens[1]) !== "RETENCION") {
      continue;
    }
    if (normalizeHeader(tokens[2]) !== "FECHA") {
      continue;
    }
    if (normalizeHeader(tokens[3]) !== "FECHADOC") {
      continue;
    }

    const numRt = tokens[11];
    const fecha = tokens[12];
    const fechaCont = tokens[13];
    const cod = tokens[14];
    const fact = tokens[15];
    const percent = tokens[16];
    const base = tokens[17];
    const ret = tokens[18];
    const proveedor = tokens[19];

    if (!/^\d+$/.test(numRt)) {
      continue;
    }
    if (!/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/.test(fecha)) {
      continue;
    }
    if (!/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/.test(fechaCont)) {
      continue;
    }
    if (!sanitizeText(proveedor)) {
      continue;
    }

    matched += 1;
    rows.push({
      "NUM RT": numRt,
      PROVEEDOR: proveedor,
      FECHA: fecha,
      "FECHA CONT": fechaCont,
      COD: cod,
      FACT: fact,
      "%": percent,
      BASE: base,
      RETENCION: ret,
      TIPO: "",
    });
  }

  const minExpected = Math.max(20, Math.floor(lines.length * 0.6));
  if (matched >= minExpected) {
    return rows;
  }

  return [];
}

function loadBestText(buffer) {
  const utf8 = buffer.toString("utf8").replace(/^\uFEFF/, "");
  const latin1 = buffer.toString("latin1").replace(/^\uFEFF/, "");
  const utf8Bad = (utf8.match(/�/g) || []).length;
  const latin1Bad = (latin1.match(/�/g) || []).length;
  return utf8Bad <= latin1Bad ? utf8 : latin1;
}

function normalizeTipo(value) {
  const clean = sanitizeText(value).toUpperCase();
  if (clean === "IVA") {
    return "IVA";
  }
  if (clean === "RENTA") {
    return "RENTA";
  }
  return clean;
}

function round2(value) {
  return Math.round((Number(value || 0) + Number.EPSILON) * 100) / 100;
}

function formatDateKey(value) {
  const date = parseDateFlexible(value);
  if (!date) {
    return "";
  }
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = date.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function normalizeFactKey(value) {
  const clean = sanitizeText(value);
  if (/^\d+$/.test(clean)) {
    return String(Number(clean));
  }
  return clean.toUpperCase();
}

function normalizeProviderKey(value) {
  return sanitizeText(value).toUpperCase();
}

function buildTipoMatchKey(fields) {
  return [
    fields.numRt == null ? "" : String(fields.numRt),
    formatDateKey(fields.fecha),
    formatDateKey(fields.fechaCont || fields.fecha),
    fields.cod == null ? "" : String(fields.cod),
    normalizeFactKey(fields.fact),
    round2(fields.percent).toFixed(4),
    round2(fields.base).toFixed(2),
    round2(fields.retencion).toFixed(2),
    normalizeProviderKey(fields.proveedor),
  ].join("|");
}

function getEmptyTipoHints() {
  return {
    exact: new Map(),
    byCodePercent: new Map(),
  };
}

function loadTemplateTipoHints(templatePath) {
  if (!fs.existsSync(templatePath)) {
    return getEmptyTipoHints();
  }

  const wb = XLSX.readFile(templatePath, { cellDates: true });
  const ws = wb.Sheets["RET PROV"];
  if (!ws || !ws["!ref"]) {
    return getEmptyTipoHints();
  }

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
  const hints = getEmptyTipoHints();

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const tipo = normalizeTipo(row[4]);
    if (!["IVA", "RENTA"].includes(tipo)) {
      continue;
    }

    const fields = {
      numRt: parseIntLike(row[0]),
      proveedor: row[1],
      fecha: row[2],
      fechaCont: row[3],
      cod: parseIntLike(row[5]),
      fact: row[6],
      percent: parseDecimalLike(row[7]),
      base: parseDecimalLike(row[8]),
      retencion: parseDecimalLike(row[9]),
    };

    if (fields.numRt == null || fields.cod == null) {
      continue;
    }

    const key = buildTipoMatchKey(fields);
    hints.exact.set(key, tipo);

    const cpKey = `${fields.cod}|${round2(fields.percent).toFixed(4)}`;
    if (!hints.byCodePercent.has(cpKey)) {
      hints.byCodePercent.set(cpKey, { IVA: 0, RENTA: 0 });
    }
    hints.byCodePercent.get(cpKey)[tipo] += 1;
  }

  return hints;
}

function inferTipoFromHints(fields, hints) {
  const exactKey = buildTipoMatchKey(fields);
  if (hints.exact.has(exactKey)) {
    return hints.exact.get(exactKey);
  }

  const rate = round2(fields.percent);
  const ivaOnly = new Set([0, 20, 30, 70, 100]);
  const rentaOnly = new Set([1, 1.75, 2, 2.75, 3, 5, 8]);

  if (ivaOnly.has(rate)) {
    return "IVA";
  }
  if (rentaOnly.has(rate)) {
    return "RENTA";
  }

  const cpKey = `${fields.cod}|${rate.toFixed(4)}`;
  const cpHint = hints.byCodePercent.get(cpKey);
  if (cpHint) {
    if (cpHint.IVA > cpHint.RENTA) {
      return "IVA";
    }
    if (cpHint.RENTA > cpHint.IVA) {
      return "RENTA";
    }
  }

  if (rate >= 20) {
    return "IVA";
  }
  return "RENTA";
}

function normalizeParsedRows(rawRows, hints) {
  const rows = [];
  const problems = [];

  for (let i = 0; i < rawRows.length; i += 1) {
    const source = rawRows[i];
    const rowNo = i + 1;

    const numRt = parseIntLike(source["NUM RT"]);
    const proveedor = sanitizeText(source.PROVEEDOR);
    const fecha = parseDateFlexible(source.FECHA);
    const fechaCont = parseDateFlexible(source["FECHA CONT"]) || fecha;
    const cod = parseIntLike(source.COD);
    const factRaw = sanitizeText(source.FACT);
    const factNum = /^\d+$/.test(factRaw) ? Number(factRaw) : factRaw;
    const percent = parseDecimalLike(source["%"]);
    const base = parseDecimalLike(source.BASE);
    const retencion = parseDecimalLike(source.RETENCION);
    let tipo = normalizeTipo(source.TIPO);
    if (!["IVA", "RENTA"].includes(tipo)) {
      tipo = inferTipoFromHints(
        {
          numRt,
          proveedor,
          fecha,
          fechaCont,
          cod,
          fact: factNum,
          percent,
          base,
          retencion,
        },
        hints,
      );
    }

    if (numRt == null) {
      problems.push(`Fila ${rowNo}: NUM RT invalido.`);
    }
    if (!proveedor) {
      problems.push(`Fila ${rowNo}: PROVEEDOR vacio.`);
    }
    if (!fecha) {
      problems.push(`Fila ${rowNo}: FECHA invalida.`);
    }
    if (!fechaCont) {
      problems.push(`Fila ${rowNo}: FECHA CONT invalida.`);
    }
    if (!["IVA", "RENTA"].includes(tipo)) {
      problems.push(`Fila ${rowNo}: TIPO invalido (${tipo || "vacio"}).`);
    }
    if (cod == null) {
      problems.push(`Fila ${rowNo}: COD invalido.`);
    }
    if (!factRaw) {
      problems.push(`Fila ${rowNo}: FACT vacio.`);
    }

    rows.push({
      numRt,
      proveedor,
      fecha,
      fechaCont,
      tipo,
      cod,
      fact: factNum,
      percent,
      base,
      retencion,
    });
  }

  if (problems.length > 0) {
    const preview = problems.slice(0, 8).join(" | ");
    throw new Error(`Validacion de TXT fallida (${problems.length} problemas). ${preview}`);
  }

  return rows;
}

function buildSummary(rows) {
  const typeOrder = ["IVA", "RENTA"];
  const typeMap = new Map();

  for (const row of rows) {
    const type = row.tipo;
    if (!typeMap.has(type)) {
      typeMap.set(type, {
        totalBase: 0,
        totalRet: 0,
        percentMap: new Map(),
      });
    }

    const bucket = typeMap.get(type);
    bucket.totalBase += row.base;
    bucket.totalRet += row.retencion;

    const key = round2(row.percent);
    if (!bucket.percentMap.has(key)) {
      bucket.percentMap.set(key, { base: 0, ret: 0 });
    }
    const percentBucket = bucket.percentMap.get(key);
    percentBucket.base += row.base;
    percentBucket.ret += row.retencion;
  }

  const entries = [];
  const seenTypes = new Set();
  for (const type of typeOrder) {
    if (typeMap.has(type)) {
      seenTypes.add(type);
      const bucket = typeMap.get(type);
      entries.push({
        kind: "type",
        label: type,
        base: round2(bucket.totalBase),
        ret: round2(bucket.totalRet),
        calc: null,
        diff: null,
      });

      const percents = [...bucket.percentMap.keys()].sort((a, b) => a - b);
      for (const percent of percents) {
        const item = bucket.percentMap.get(percent);
        entries.push({
          kind: "detail",
          label: formatPercentLabel(percent),
          base: round2(item.base),
          ret: round2(item.ret),
          calc: round2(item.ret),
          diff: 0,
        });
      }
    }
  }

  const remainingTypes = [...typeMap.keys()].filter((type) => !seenTypes.has(type)).sort();
  for (const type of remainingTypes) {
    const bucket = typeMap.get(type);
    entries.push({
      kind: "type",
      label: type,
      base: round2(bucket.totalBase),
      ret: round2(bucket.totalRet),
      calc: null,
      diff: null,
    });
    const percents = [...bucket.percentMap.keys()].sort((a, b) => a - b);
    for (const percent of percents) {
      const item = bucket.percentMap.get(percent);
      entries.push({
        kind: "detail",
        label: formatPercentLabel(percent),
        base: round2(item.base),
        ret: round2(item.ret),
        calc: round2(item.ret),
        diff: 0,
      });
    }
  }

  const totalBase = round2(rows.reduce((sum, row) => sum + row.base, 0));
  const totalRet = round2(rows.reduce((sum, row) => sum + row.retencion, 0));
  entries.push({
    kind: "total",
    label: "Total general",
    base: totalBase,
    ret: totalRet,
    calc: null,
    diff: null,
  });

  return entries;
}

function formatPercentLabel(value) {
  if (Number.isInteger(value)) {
    return String(value);
  }
  const text = String(value);
  return text.replace(/(\.\d*?[1-9])0+$/g, "$1").replace(/\.0+$/g, "");
}

function deepClone(value) {
  return value ? JSON.parse(JSON.stringify(value)) : {};
}

function captureStyleMatrix(ws, maxRows, maxCols) {
  const matrix = Array.from({ length: maxRows + 1 }, () => Array(maxCols + 1).fill(null));
  for (let row = 1; row <= maxRows; row += 1) {
    for (let col = 1; col <= maxCols; col += 1) {
      matrix[row][col] = deepClone(ws.getCell(row, col).style);
    }
  }
  return matrix;
}

function applyStyleFromMatrix(ws, matrix, row, col) {
  const style = matrix[row] && matrix[row][col] ? matrix[row][col] : {};
  ws.getCell(row, col).style = deepClone(style);
}

function clearRangeValues(ws, startRow, endRow, startCol, endCol) {
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      ws.getCell(row, col).value = null;
    }
  }
}

function hasStyle(cellStyle) {
  return cellStyle && Object.keys(cellStyle).length > 0;
}

function findLastFullyStyledRow(ws, startRow, endRow, startCol, endCol) {
  let last = startRow;
  for (let row = startRow; row <= endRow; row += 1) {
    let full = true;
    for (let col = startCol; col <= endCol; col += 1) {
      if (!hasStyle(ws.getCell(row, col).style)) {
        full = false;
        break;
      }
    }
    if (full) {
      last = row;
    }
  }
  return last;
}

function stableStringify(value) {
  if (value === null || value === undefined) {
    return String(value);
  }
  if (typeof value !== "object") {
    return JSON.stringify(value);
  }
  if (Array.isArray(value)) {
    return `[${value.map((item) => stableStringify(item)).join(",")}]`;
  }

  const keys = Object.keys(value).sort();
  const parts = keys.map((key) => `${JSON.stringify(key)}:${stableStringify(value[key])}`);
  return `{${parts.join(",")}}`;
}

function styleSignature(style) {
  return stableStringify(style || {});
}

function normalizeColumnWidth(width) {
  return Number.isFinite(width) ? Number(width) : null;
}

function applyFallbackStylesForExtendedDataRow(ws, rowNumber, fallbackStyles) {
  for (let col = 1; col <= 11; col += 1) {
    ws.getCell(rowNumber, col).style = deepClone(fallbackStyles[col - 1] || {});
  }
}

function applyFallbackStylesForExtendedSummaryRow(ws, rowNumber, kind, fallbackHeader, fallbackType, fallbackDetail, fallbackTotal) {
  let source = fallbackDetail;
  if (rowNumber === 1) {
    source = fallbackHeader;
  } else if (kind === "type") {
    source = fallbackType;
  } else if (kind === "total") {
    source = fallbackTotal;
  }

  for (let col = 12; col <= 16; col += 1) {
    ws.getCell(rowNumber, col).style = deepClone(source[col - 12] || {});
  }
}

async function writeWorkbookWithRetries(workbook, preferredPath, maxAttempts = 20) {
  const parsed = path.parse(preferredPath);
  fs.mkdirSync(parsed.dir || ".", { recursive: true });

  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    const candidate =
      attempt === 0
        ? preferredPath
        : path.join(parsed.dir, `${parsed.name}_nuevo${attempt === 1 ? "" : `_${attempt}`}${parsed.ext}`);

    try {
      await workbook.xlsx.writeFile(candidate);
      return candidate;
    } catch (error) {
      const isLocked = error && (error.code === "EBUSY" || error.code === "EPERM");
      if (!isLocked) {
        throw error;
      }
    }
  }

  throw new Error("No se pudo guardar el Excel de Accion 2. Cierra archivos abiertos e intenta de nuevo.");
}

function extractRowsFromTxt(inputPath) {
  const buffer = fs.readFileSync(inputPath);
  const text = loadBestText(buffer);
  const lines = text
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .map((line) => line.replace(/\u00A0/g, " "));

  let rawRows = parseRowsFromEmbeddedRetReport(lines);

  if (rawRows.length === 0) {
    const headerConfig = detectHeaderConfig(lines);
    if (headerConfig) {
      rawRows = parseRowsFromHeader(lines, headerConfig);
    }
  }

  if (rawRows.length === 0) {
    rawRows = parseRowsByPattern(lines);
  }

  if (rawRows.length === 0) {
    throw new Error("No se detectaron filas validas en el TXT. Verifica encabezados o delimitadores.");
  }

  return rawRows;
}

function toArray(value) {
  if (value === null || value === undefined) {
    return [];
  }
  return Array.isArray(value) ? value : [value];
}

function normalizeZipPath(inputPath) {
  return String(inputPath || "").replace(/\\/g, "/").replace(/^\/+/, "");
}

function parseXml(xmlText) {
  return XML_PARSER.parse(xmlText);
}

function buildXml(xmlObject) {
  const body = XML_BUILDER.build(xmlObject);
  if (body.startsWith("<?xml")) {
    return body;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body}`;
}

function readZipText(zip, entryPath) {
  const entry = zip.getEntry(normalizeZipPath(entryPath));
  if (!entry) {
    return null;
  }
  return entry.getData().toString("utf8");
}

function updateZipText(zip, entryPath, xmlObject) {
  zip.updateFile(normalizeZipPath(entryPath), Buffer.from(buildXml(xmlObject), "utf8"));
}

function columnNumberToName(colNumber) {
  let n = Number(colNumber);
  let result = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function columnNameToNumber(colName) {
  let result = 0;
  const text = String(colName || "").toUpperCase();
  for (let i = 0; i < text.length; i += 1) {
    const code = text.charCodeAt(i);
    if (code < 65 || code > 90) {
      return 0;
    }
    result = result * 26 + (code - 64);
  }
  return result;
}

function cellRef(colNumber, rowNumber) {
  return `${columnNumberToName(colNumber)}${rowNumber}`;
}

function parseCellRef(ref) {
  const match = /^([A-Z]+)(\d+)$/i.exec(String(ref || ""));
  if (!match) {
    return null;
  }
  return {
    colName: match[1].toUpperCase(),
    colNumber: columnNameToNumber(match[1]),
    rowNumber: Number(match[2]),
  };
}

function getWorksheetEntryPath(zip, sheetName) {
  const workbookXml = readZipText(zip, "xl/workbook.xml");
  const relsXml = readZipText(zip, "xl/_rels/workbook.xml.rels");
  if (!workbookXml || !relsXml) {
    throw new Error("No se pudo leer la estructura interna del Excel.");
  }

  const workbook = parseXml(workbookXml);
  const rels = parseXml(relsXml);
  const sheets = toArray(workbook.workbook?.sheets?.sheet);
  const sheet = sheets.find((item) => sanitizeText(item?.["@_name"]).toUpperCase() === sanitizeText(sheetName).toUpperCase());
  if (!sheet) {
    throw new Error(`No se encontro la hoja ${sheetName} en el Excel.`);
  }

  const sheetRid = sheet["@_r:id"];
  const relationships = toArray(rels.Relationships?.Relationship);
  const rel = relationships.find((item) => item?.["@_Id"] === sheetRid);
  if (!rel) {
    throw new Error(`No se pudo resolver la ruta XML de la hoja ${sheetName}.`);
  }

  const target = normalizeZipPath(rel["@_Target"]);
  return target.startsWith("xl/") ? target : normalizeZipPath(path.posix.join("xl", target));
}

function getRowsArray(sheetXmlObject) {
  return toArray(sheetXmlObject.worksheet?.sheetData?.row);
}

function setRowsArray(sheetXmlObject, rows) {
  if (!sheetXmlObject.worksheet.sheetData) {
    sheetXmlObject.worksheet.sheetData = {};
  }
  sheetXmlObject.worksheet.sheetData.row = rows;
}

function getCellsArray(rowObject) {
  return toArray(rowObject?.c);
}

function setCellsArray(rowObject, cells) {
  if (cells.length === 0) {
    delete rowObject.c;
    return;
  }
  rowObject.c = cells;
}

function copyCellPayload(targetCell, sourceCell) {
  if (sourceCell["@_t"] !== undefined) {
    targetCell["@_t"] = sourceCell["@_t"];
  } else {
    delete targetCell["@_t"];
  }

  if (sourceCell["@_cm"] !== undefined) {
    targetCell["@_cm"] = sourceCell["@_cm"];
  } else {
    delete targetCell["@_cm"];
  }

  if (sourceCell["@_vm"] !== undefined) {
    targetCell["@_vm"] = sourceCell["@_vm"];
  } else {
    delete targetCell["@_vm"];
  }

  if (sourceCell.f !== undefined) {
    targetCell.f = deepClone(sourceCell.f);
  } else {
    delete targetCell.f;
  }

  if (sourceCell.v !== undefined) {
    targetCell.v = deepClone(sourceCell.v);
  } else {
    delete targetCell.v;
  }

  if (sourceCell.is !== undefined) {
    targetCell.is = deepClone(sourceCell.is);
  } else {
    delete targetCell.is;
  }
}

function clearCellPayload(targetCell) {
  delete targetCell["@_t"];
  delete targetCell["@_cm"];
  delete targetCell["@_vm"];
  delete targetCell.f;
  delete targetCell.v;
  delete targetCell.is;
}

function hasCellPayload(cell) {
  return cell.v !== undefined || cell.f !== undefined || cell.is !== undefined;
}

function stripCalcChain(zip) {
  const calcChainPath = "xl/calcChain.xml";
  if (zip.getEntry(calcChainPath)) {
    zip.deleteFile(calcChainPath);
  }

  const relsXml = readZipText(zip, "xl/_rels/workbook.xml.rels");
  if (relsXml) {
    const relsObj = parseXml(relsXml);
    const rels = toArray(relsObj.Relationships?.Relationship);
    const kept = rels.filter(
      (item) => !String(item?.["@_Type"] || "").endsWith("/calcChain"),
    );
    relsObj.Relationships.Relationship = kept;
    updateZipText(zip, "xl/_rels/workbook.xml.rels", relsObj);
  }

  const contentTypesXml = readZipText(zip, "[Content_Types].xml");
  if (contentTypesXml) {
    const typesObj = parseXml(contentTypesXml);
    const overrides = toArray(typesObj.Types?.Override);
    const kept = overrides.filter((item) => item?.["@_PartName"] !== "/xl/calcChain.xml");
    typesObj.Types.Override = kept;
    updateZipText(zip, "[Content_Types].xml", typesObj);
  }
}

function preserveTemplateVisualWorkbook(templatePath, generatedPath, sheetName = "RET PROV") {
  const templateZip = new AdmZip(templatePath);
  const generatedZip = new AdmZip(generatedPath);

  const templateSheetPath = getWorksheetEntryPath(templateZip, sheetName);
  const generatedSheetPath = getWorksheetEntryPath(generatedZip, sheetName);
  const templateSheetXml = readZipText(templateZip, templateSheetPath);
  const generatedSheetXml = readZipText(generatedZip, generatedSheetPath);
  if (!templateSheetXml || !generatedSheetXml) {
    throw new Error("No se pudo comparar hoja XML entre plantilla y generado.");
  }

  const templateSheet = parseXml(templateSheetXml);
  const generatedSheet = parseXml(generatedSheetXml);

  const templateRows = getRowsArray(templateSheet);
  const generatedRows = getRowsArray(generatedSheet);

  const templateRowMap = new Map();
  const templateStyleByCol = new Map();
  const dataFallbackStyleByCol = new Map();
  const generatedCellMap = new Map();
  let templateMaxRow = 1;
  let generatedMaxRowWithPayload = 1;

  for (const row of templateRows) {
    const rowNumber = Number(row?.["@_r"] || 0);
    if (rowNumber > 0) {
      templateRowMap.set(rowNumber, row);
      templateMaxRow = Math.max(templateMaxRow, rowNumber);
    }

    const rowCells = getCellsArray(row);
    const rowStyleByCol = new Map();
    for (const cell of rowCells) {
      const parsedRef = parseCellRef(cell?.["@_r"]);
      if (!parsedRef || parsedRef.colNumber < 1 || parsedRef.colNumber > 16) {
        continue;
      }
      if (cell["@_s"] !== undefined) {
        templateStyleByCol.set(parsedRef.colNumber, cell["@_s"]);
        rowStyleByCol.set(parsedRef.colNumber, cell["@_s"]);
      }
    }

    let fullDataStyled = true;
    for (let colNumber = 1; colNumber <= 10; colNumber += 1) {
      if (!rowStyleByCol.has(colNumber)) {
        fullDataStyled = false;
        break;
      }
    }
    if (fullDataStyled) {
      for (let colNumber = 1; colNumber <= 10; colNumber += 1) {
        dataFallbackStyleByCol.set(colNumber, rowStyleByCol.get(colNumber));
      }
    }
  }

  for (const row of generatedRows) {
    const rowNumber = Number(row?.["@_r"] || 0);
    if (rowNumber <= 0) {
      continue;
    }

    const cells = getCellsArray(row);
    for (const cell of cells) {
      const ref = String(cell?.["@_r"] || "");
      const parsedRef = parseCellRef(ref);
      if (!parsedRef || parsedRef.colNumber < 1 || parsedRef.colNumber > 16) {
        continue;
      }
      const key = `${parsedRef.rowNumber}:${parsedRef.colNumber}`;
      generatedCellMap.set(key, cell);
      if (hasCellPayload(cell) || cell["@_s"] !== undefined) {
        generatedMaxRowWithPayload = Math.max(generatedMaxRowWithPayload, parsedRef.rowNumber);
      }
    }
  }

  const maxRow = Math.max(templateMaxRow, generatedMaxRowWithPayload);
  for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
    let row = templateRowMap.get(rowNumber);
    if (!row) {
      row = { "@_r": String(rowNumber) };
      templateRows.push(row);
      templateRowMap.set(rowNumber, row);
    }

    const rowCells = getCellsArray(row);
    const rowCellMap = new Map();
    for (const cell of rowCells) {
      const parsedRef = parseCellRef(cell?.["@_r"]);
      if (!parsedRef) {
        continue;
      }
      rowCellMap.set(parsedRef.colNumber, cell);
    }

    for (let colNumber = 1; colNumber <= 16; colNumber += 1) {
      const key = `${rowNumber}:${colNumber}`;
      const generatedCell = generatedCellMap.get(key);
      let targetCell = rowCellMap.get(colNumber);

      if (!targetCell) {
        if (!generatedCell) {
          continue;
        }
        targetCell = { "@_r": cellRef(colNumber, rowNumber) };
        if (colNumber >= 1 && colNumber <= 10) {
          const fallbackStyle = dataFallbackStyleByCol.get(colNumber) || templateStyleByCol.get(colNumber);
          if (fallbackStyle !== undefined) {
            targetCell["@_s"] = fallbackStyle;
          }
        }
        rowCellMap.set(colNumber, targetCell);
        rowCells.push(targetCell);
      }

      if (generatedCell) {
        copyCellPayload(targetCell, generatedCell);
      } else {
        clearCellPayload(targetCell);
      }
    }

    rowCells.sort((a, b) => {
      const ca = parseCellRef(a?.["@_r"]);
      const cb = parseCellRef(b?.["@_r"]);
      return (ca?.colNumber || 0) - (cb?.colNumber || 0);
    });

    const filtered = rowCells.filter((cell) => {
      const keepByStyle = cell["@_s"] !== undefined;
      const keepByPayload = hasCellPayload(cell);
      return keepByStyle || keepByPayload;
    });
    setCellsArray(row, filtered);
  }

  templateRows.sort((a, b) => Number(a?.["@_r"] || 0) - Number(b?.["@_r"] || 0));
  setRowsArray(templateSheet, templateRows);
  if (generatedSheet.worksheet?.dimension) {
    templateSheet.worksheet.dimension = deepClone(generatedSheet.worksheet.dimension);
  }
  updateZipText(templateZip, templateSheetPath, templateSheet);

  const generatedSharedStrings = generatedZip.getEntry("xl/sharedStrings.xml");
  if (generatedSharedStrings) {
    templateZip.updateFile("xl/sharedStrings.xml", generatedSharedStrings.getData());
  }

  stripCalcChain(templateZip);
  templateZip.writeZip(generatedPath);
}

async function verifyOutputWorkbook(outputPath, templatePath, rowsCount, summary) {
  const templateZip = new AdmZip(templatePath);
  const outputZip = new AdmZip(outputPath);
  const criticalEntries = [
    "xl/pivotTables/pivotTable1.xml",
    "xl/worksheets/_rels/sheet1.xml.rels",
    "xl/pivotCache/pivotCacheDefinition1.xml",
    "xl/externalLinks/externalLink1.xml",
  ];
  for (const entryPath of criticalEntries) {
    const expected = templateZip.getEntry(entryPath);
    const actual = outputZip.getEntry(entryPath);
    if (expected && !actual) {
      throw new Error(`Validacion final: falta artefacto visual de plantilla (${entryPath}).`);
    }
  }

  const wb = XLSX.readFile(outputPath, { cellFormula: true });
  if (wb.SheetNames.length !== 1 || wb.SheetNames[0] !== "RET PROV") {
    throw new Error("Validacion final: el archivo debe tener una sola hoja llamada RET PROV.");
  }

  const ws = wb.Sheets["RET PROV"];
  const headers = EXPECTED_COLUMNS.map((_, idx) => ws[XLSX.utils.encode_cell({ r: 0, c: idx })]?.v || "");
  for (let i = 0; i < EXPECTED_COLUMNS.length; i += 1) {
    if (sanitizeText(headers[i]).toUpperCase() !== EXPECTED_COLUMNS[i]) {
      throw new Error(`Validacion final: encabezado incorrecto en columna ${i + 1}.`);
    }
  }

  const templateWb = new ExcelJS.Workbook();
  await templateWb.xlsx.readFile(templatePath);
  const outWb = new ExcelJS.Workbook();
  await outWb.xlsx.readFile(outputPath);

  const templateWs = templateWb.getWorksheet("RET PROV");
  const outWs = outWb.getWorksheet("RET PROV");
  if (!templateWs || !outWs) {
    throw new Error("Validacion final: no se pudo comparar formato de hoja RET PROV.");
  }

  for (let col = 1; col <= 16; col += 1) {
    const tWidth = normalizeColumnWidth(templateWs.getColumn(col).width);
    const oWidth = normalizeColumnWidth(outWs.getColumn(col).width);
    if (tWidth !== oWidth) {
      throw new Error(`Validacion final: ancho de columna alterado en ${col}.`);
    }
  }

  for (let col = 1; col <= 16; col += 1) {
    const expected = styleSignature(templateWs.getCell(1, col).style);
    const actual = styleSignature(outWs.getCell(1, col).style);
    if (actual !== expected) {
      throw new Error(`Validacion final: estilo de encabezado alterado en fila 1, columna ${col}.`);
    }
  }

  const templateRowCount = templateWs.rowCount;
  const lastStyledDataRow = findLastFullyStyledRow(templateWs, 2, templateRowCount, 1, 10);
  const maxDataRow = rowsCount + 1;
  for (let rowNum = 2; rowNum <= maxDataRow; rowNum += 1) {
    const sourceRow = rowNum <= lastStyledDataRow ? rowNum : lastStyledDataRow;
    for (let col = 1; col <= 11; col += 1) {
      const expected = styleSignature(templateWs.getCell(sourceRow, col).style);
      const actual = styleSignature(outWs.getCell(rowNum, col).style);
      if (actual !== expected) {
        throw new Error(`Validacion final: estilo de datos alterado en ${rowNum}:${col}.`);
      }
    }
  }

  const summaryRows = Array.isArray(summary) ? summary : [];
  for (let i = 0; i < summaryRows.length; i += 1) {
    const rowNum = i + 2;
    let sourceRow;
    if (rowNum <= templateRowCount) {
      sourceRow = rowNum;
    } else if (summaryRows[i].kind === "type") {
      sourceRow = 2;
    } else if (summaryRows[i].kind === "total") {
      sourceRow = 16;
    } else {
      sourceRow = 3;
    }
    for (let col = 12; col <= 16; col += 1) {
      const expected = styleSignature(templateWs.getCell(sourceRow, col).style);
      const actual = styleSignature(outWs.getCell(rowNum, col).style);
      if (actual !== expected) {
        throw new Error(`Validacion final: estilo de resumen alterado en ${rowNum}:${col}.`);
      }
    }
  }
}

async function buildWorkbookFromTemplate(templatePath, rows) {
  if (!fs.existsSync(templatePath)) {
    throw new Error(`No se encontro plantilla de Accion 2: ${templatePath}`);
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(templatePath);
  const ws = wb.getWorksheet("RET PROV");
  if (!ws) {
    throw new Error("La plantilla de Accion 2 no contiene la hoja RET PROV.");
  }

  const templateRowCount = ws.rowCount;
  const templateStyles = captureStyleMatrix(ws, templateRowCount, 16);
  const maxRows = Math.max(templateRowCount, rows.length + 40, 550);
  const lastStyledDataRow = findLastFullyStyledRow(ws, 2, templateRowCount, 1, 10);
  const fallbackDataStyles = [];
  for (let col = 1; col <= 11; col += 1) {
    fallbackDataStyles.push(deepClone(ws.getCell(lastStyledDataRow, col).style));
  }

  clearRangeValues(ws, 2, maxRows, 1, 11);
  clearRangeValues(ws, 1, maxRows, 12, 16);
  for (let row = 1; row <= templateRowCount; row += 1) {
    for (let col = 1; col <= 16; col += 1) {
      applyStyleFromMatrix(ws, templateStyles, row, col);
    }
  }

  const summaryHeaderStyles = [];
  for (let col = 12; col <= 16; col += 1) {
    summaryHeaderStyles.push(deepClone(ws.getCell(1, col).style));
  }
  const summaryTypeStyles = [];
  for (let col = 12; col <= 16; col += 1) {
    summaryTypeStyles.push(deepClone(ws.getCell(2, col).style));
  }
  const summaryDetailStyles = [];
  for (let col = 12; col <= 16; col += 1) {
    summaryDetailStyles.push(deepClone(ws.getCell(3, col).style));
  }
  const summaryTotalStyles = [];
  for (let col = 12; col <= 16; col += 1) {
    summaryTotalStyles.push(deepClone(ws.getCell(16, col).style));
  }

  for (let i = 0; i < rows.length; i += 1) {
    const rowNum = i + 2;
    const row = rows[i];

    ws.getCell(`A${rowNum}`).value = row.numRt;
    ws.getCell(`B${rowNum}`).value = row.proveedor;
    ws.getCell(`C${rowNum}`).value = row.fecha;
    ws.getCell(`D${rowNum}`).value = row.fechaCont;
    ws.getCell(`E${rowNum}`).value = row.tipo;
    ws.getCell(`F${rowNum}`).value = row.cod;
    ws.getCell(`G${rowNum}`).value = row.fact;
    ws.getCell(`H${rowNum}`).value = round2(row.percent);
    ws.getCell(`I${rowNum}`).value = round2(row.base);
    ws.getCell(`J${rowNum}`).value = round2(row.retencion);
    ws.getCell(`K${rowNum}`).value = null;
    if (rowNum <= lastStyledDataRow) {
      for (let col = 1; col <= 11; col += 1) {
        applyStyleFromMatrix(ws, templateStyles, rowNum, col);
      }
    } else {
      applyFallbackStylesForExtendedDataRow(ws, rowNum, fallbackDataStyles);
      ws.getRow(rowNum).height = ws.getRow(lastStyledDataRow).height;
    }
  }

  ws.getCell("L1").value = "Etiquetas de fila";
  ws.getCell("M1").value = "Suma de BASE";
  ws.getCell("N1").value = "Suma de RETENCION";
  ws.getCell("O1").value = "";
  ws.getCell("P1").value = "";
  for (let col = 12; col <= 16; col += 1) {
    applyStyleFromMatrix(ws, templateStyles, 1, col);
  }

  const summary = buildSummary(rows);
  for (let i = 0; i < summary.length; i += 1) {
    const rowNum = i + 2;
    const item = summary[i];
    ws.getCell(`L${rowNum}`).value = item.label;
    ws.getCell(`M${rowNum}`).value = item.base;
    ws.getCell(`N${rowNum}`).value = item.ret;
    ws.getCell(`O${rowNum}`).value = item.calc == null ? null : item.calc;
    ws.getCell(`P${rowNum}`).value = item.diff == null ? null : item.diff;

    if (rowNum <= templateRowCount) {
      for (let col = 12; col <= 16; col += 1) {
        applyStyleFromMatrix(ws, templateStyles, rowNum, col);
      }
    } else {
      applyFallbackStylesForExtendedSummaryRow(
        ws,
        rowNum,
        item.kind,
        summaryHeaderStyles,
        summaryTypeStyles,
        summaryDetailStyles,
        summaryTotalStyles,
      );
      ws.getRow(rowNum).height = ws.getRow(3).height;
    }
  }

  return { workbook: wb, summary };
}

function writeAuditReport(auditPath, payload) {
  fs.mkdirSync(path.dirname(auditPath), { recursive: true });
  fs.writeFileSync(auditPath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
}

async function main() {
  const inputPath = path.resolve(process.cwd(), INPUT_TXT);
  const outputPath = path.resolve(process.cwd(), OUTPUT_XLSX);
  const templatePath = TEMPLATE_XLSX;

  if (!fs.existsSync(inputPath)) {
    throw new Error(`No se encontro el TXT: ${inputPath}`);
  }

  const hints = loadTemplateTipoHints(templatePath);
  const rawRows = extractRowsFromTxt(inputPath);
  const rows = normalizeParsedRows(rawRows, hints);
  const { workbook, summary } = await buildWorkbookFromTemplate(templatePath, rows);
  const finalOutputPath = await writeWorkbookWithRetries(workbook, outputPath);
  preserveTemplateVisualWorkbook(templatePath, finalOutputPath, "RET PROV");
  await verifyOutputWorkbook(finalOutputPath, templatePath, rows.length, summary);

  const auditPath = path.join(
    path.dirname(finalOutputPath),
    `${path.parse(finalOutputPath).name}_auditoria.json`,
  );

  const auditPayload = {
    fecha_proceso: new Date().toISOString(),
    input_txt: inputPath,
    output_xlsx: finalOutputPath,
    hoja_salida: "RET PROV",
    filas_txt: rows.length,
    columnas_esperadas: EXPECTED_COLUMNS,
    resumen_generado_filas: summary.length,
    verificacion_final_ok: true,
  };
  writeAuditReport(auditPath, auditPayload);

  console.log(`TXT leido: ${inputPath}`);
  console.log(`Filas parseadas: ${rows.length}`);
  console.log(`Resumen lateral: ${summary.length} filas`);
  console.log(`Excel generado (una sola hoja): ${finalOutputPath}`);
  console.log(`Auditoria JSON: ${auditPath}`);
}

main().catch((error) => {
  console.error("Error:", error.message);
  process.exit(1);
});
