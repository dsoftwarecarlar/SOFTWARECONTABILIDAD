const fs = require("fs");
const http = require("http");
const os = require("os");
const path = require("path");
const { spawn } = require("child_process");

const AdmZip = require("adm-zip");
const ExcelJS = require("exceljs");
const XLSX = require("xlsx");

const {
  extractRowsFromTxt,
  loadTemplateTipoHints,
  normalizeParsedRows,
} = require("../cxp/accion2/parser");
const accion2Constants = require("../cxp/accion2/constants");
const {
  buildWorkbookFromTemplate: buildAction2Workbook,
} = require("../cxp/accion2/workbook");
const { parseInputSources: parseAction3Inputs } = require("../cxp/accion3/parser");
const {
  buildWorkbookFromTemplate: buildAction3Workbook,
} = require("../cxp/accion3/workbook");
const {
  parseInputSources: parseAction4Inputs,
  buildAction4OutputPlan,
} = require("../cxp/accion4/parser");
const {
  buildWorkbookFromTemplate: buildAction4Workbook,
} = require("../cxp/accion4/workbook");

const ROOT = path.resolve(__dirname, "..", "..");
const APP_ROOT = path.join(ROOT, "laravel_app");
const ROUTER = path.join(APP_ROOT, "server.php");
const PHP_EXE = firstExistingPath([
  path.join(ROOT, ".tools", "php82", "php.exe"),
  process.env.PHP_EXE || "",
  "C:\\xampp\\php\\php.exe",
]);
const HOST = "127.0.0.1";
const PORT = 18997;
const BASE_URL = `http://${HOST}:${PORT}`;

function assertCondition(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

function firstExistingPath(candidates) {
  for (const candidate of candidates) {
    if (!candidate) {
      continue;
    }
    if (candidate.includes("\\") || candidate.includes("/")) {
      if (fs.existsSync(candidate)) {
        return candidate;
      }
      continue;
    }
    return candidate;
  }

  return candidates[0];
}

function actionFixturePath(fileName) {
  return firstExistingPath([
    path.join(ROOT, "resources", "cxp", "acciones", "fixtures", fileName),
    path.join(ROOT, "outputs", "EJEMPLOSAMANO1", fileName),
  ]);
}

function actionTemplatePath(fileName) {
  return firstExistingPath([
    path.join(ROOT, "resources", "cxp", "acciones", "templates", fileName),
    path.join(ROOT, "outputs", "EJEMPLOSAMANO1", fileName),
  ]);
}

function actionContractPath(fileName) {
  return firstExistingPath([
    path.join(ROOT, "resources", "cxp", "acciones", "contracts", fileName),
    path.join(ROOT, "outputs", fileName),
  ]);
}

function actionMonthlyPath(matchers, label) {
  const monthlyDir = path.join(ROOT, "resources", "cxp", "acciones", "PLANTILLAYARCHIVOS");
  assertCondition(fs.existsSync(monthlyDir), `${label}: no existe PLANTILLAYARCHIVOS.`);
  const entries = fs.readdirSync(monthlyDir);
  const found = entries.find((name) => matchers.some((matcher) => matcher.test(name)));
  assertCondition(!!found, `${label}: no se encontro archivo mensual esperado.`);
  return path.join(monthlyDir, found);
}

function sanitizeText(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function round2(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return value;
  }
  return Math.round((numeric + Number.EPSILON) * 100) / 100;
}

function normalizeTextComparison(value) {
  const normalized = sanitizeText(value)
    .replace(/�/g, "N")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
  return normalized === "" ? null : normalized;
}

function startPhpServer() {
  const server = spawn(PHP_EXE, ["-d", "max_execution_time=180", "-S", `${HOST}:${PORT}`, ROUTER], {
    cwd: APP_ROOT,
    stdio: ["ignore", "pipe", "pipe"],
  });

  server.stdout.on("data", (chunk) => process.stdout.write(`[php] ${chunk}`));
  server.stderr.on("data", (chunk) => process.stderr.write(`[php] ${chunk}`));
  return server;
}

function waitServerReady(retries = 60) {
  return new Promise((resolve, reject) => {
    let attempts = 0;
    const tick = () => {
      attempts += 1;
      const req = http.request(
        { hostname: HOST, port: PORT, path: "/", method: "GET" },
        (res) => {
          res.resume();
          resolve();
        },
      );
      req.on("error", (error) => {
        if (attempts >= retries) {
          reject(new Error(`No se pudo iniciar servidor Laravel: ${error.message}`));
          return;
        }
        setTimeout(tick, 200);
      });
      req.end();
    };
    tick();
  });
}

function httpRequest(method, requestPath, body = null, headers = {}) {
  return new Promise((resolve, reject) => {
    const req = http.request(
      { hostname: HOST, port: PORT, path: requestPath, method, headers },
      (res) => {
        const chunks = [];
        res.on("data", (chunk) => chunks.push(chunk));
        res.on("end", () => {
          const buffer = Buffer.concat(chunks);
          resolve({
            status: res.statusCode || 0,
            headers: res.headers,
            buffer,
            text: buffer.toString("utf8"),
          });
        });
      },
    );
    req.on("error", reject);
    if (body) {
      req.write(body);
    }
    req.end();
  });
}

async function httpRequestWithRedirects(method, requestPath, body = null, headers = {}, maxRedirects = 5) {
  let currentPath = requestPath;
  let currentMethod = method;
  let currentBody = body;

  for (let attempt = 0; attempt <= maxRedirects; attempt += 1) {
    const response = await httpRequest(currentMethod, currentPath, currentBody, headers);
    const location = response.headers.location;
    if (![301, 302, 303, 307, 308].includes(response.status) || !location) {
      return response;
    }

    const parsed = location.startsWith("http") ? new URL(location) : new URL(location, BASE_URL);
    currentPath = parsed.pathname + parsed.search;
    currentMethod = "GET";
    currentBody = null;
  }

  throw new Error(`Demasiados redirects al solicitar ${requestPath}.`);
}

function buildMultipartBody(token, files = []) {
  const boundary = `----CodexBoundary${Date.now()}${Math.random().toString(16).slice(2)}`;
  const chunks = [];
  const push = (value) => chunks.push(Buffer.isBuffer(value) ? value : Buffer.from(value, "utf8"));

  push(`--${boundary}\r\n`);
  push('Content-Disposition: form-data; name="_token"\r\n\r\n');
  push(token);
  push("\r\n");

  for (const file of files) {
    push(`--${boundary}\r\n`);
    push(`Content-Disposition: form-data; name="${file.field}"; filename="${path.basename(file.path)}"\r\n`);
    push(`Content-Type: ${file.contentType}\r\n\r\n`);
    push(fs.readFileSync(file.path));
    push("\r\n");
  }

  push(`--${boundary}--\r\n`);
  return { boundary, body: Buffer.concat(chunks) };
}

function extractCsrfToken(html, label) {
  const match = html.match(/name="_token"\s+value="([^"]+)"/i);
  assertCondition(match, `No se encontro CSRF token en ${label}.`);
  return match[1];
}

function extractCookie(headers, label) {
  const setCookie = headers["set-cookie"];
  assertCondition(Array.isArray(setCookie) && setCookie.length > 0, `Laravel no devolvio cookie en ${label}.`);
  return setCookie.map((cookie) => cookie.split(";", 1)[0]).join("; ");
}

function extractDownloadInfo(html, label) {
  const linkMatch = html.match(/href="([^"]*\/downloads\/[^"]+)"/i);
  assertCondition(linkMatch, `${label}: no se encontro enlace de descarga.`);
  const link = linkMatch[1];
  const absoluteUrl = link.startsWith("http") ? link : `${BASE_URL}${link.startsWith("/") ? "" : "/"}${link}`;
  const parsed = new URL(absoluteUrl);
  return {
    link: parsed.pathname + parsed.search,
  };
}

async function submitModuleFiles(moduleSlug, files) {
  const page = await httpRequest("GET", `/cxp/modules/${moduleSlug}`);
  assertCondition(page.status === 200, `${moduleSlug}: GET devolvio HTTP ${page.status}.`);
  const cookie = extractCookie(page.headers, moduleSlug);
  const token = extractCsrfToken(page.text, moduleSlug);
  const payload = buildMultipartBody(token, files);

  const response = await httpRequest("POST", `/cxp/modules/${moduleSlug}`, payload.body, {
    "Content-Type": `multipart/form-data; boundary=${payload.boundary}`,
    "Content-Length": String(payload.body.length),
    Cookie: cookie,
  });

  assertCondition(response.status === 200, `${moduleSlug}: POST devolvio HTTP ${response.status}.`);
  const downloadInfo = extractDownloadInfo(response.text, moduleSlug);
  const download = await httpRequestWithRedirects("GET", downloadInfo.link, null, { Cookie: cookie });
  assertCondition(download.status === 200, `${moduleSlug}: descarga devolvio HTTP ${download.status}.`);
  return download.buffer;
}

async function submitBundleModule() {
  return submitModuleFiles("consolidado-acciones", []);
}

function readZipEntryText(zip, entryName) {
  const entry = zip.getEntry(entryName);
  assertCondition(!!entry, `No existe ${entryName} dentro del XLSX.`);
  return entry.getData().toString("utf8");
}

function assertNoExternalLinks(buffer, label) {
  const zip = new AdmZip(buffer);
  const names = zip.getEntries().map((entry) => entry.entryName);
  const externalEntries = names.filter((name) => /^xl\/externalLinks\//i.test(name));
  assertCondition(externalEntries.length === 0, `${label}: el XLSX contiene externalLinks heredados.`);

  const workbookXml = readZipEntryText(zip, "xl/workbook.xml");
  assertCondition(!/<externalReferences\b/i.test(workbookXml), `${label}: xl/workbook.xml conserva externalReferences.`);

  const workbookRels = readZipEntryText(zip, "xl/_rels/workbook.xml.rels");
  assertCondition(!/relationships\/externalLink/i.test(workbookRels), `${label}: workbook.xml.rels conserva externalLink.`);
}

function assertWorkbookViewStructure(buffer, label) {
  const zip = new AdmZip(buffer);
  const workbookXml = readZipEntryText(zip, "xl/workbook.xml");
  const worksheetEntries = zip
    .getEntries()
    .filter((entry) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(entry.entryName));
  const requiresBookViews = worksheetEntries.some((entry) =>
    /<sheetView\b[^>]*\bworkbookViewId="(\d+)"/.test(entry.getData().toString("utf8")),
  );

  if (!requiresBookViews) {
    return;
  }

  assertCondition(
    /<(?:\w+:)?bookViews\b/.test(workbookXml) && /<(?:\w+:)?workbookView\b/.test(workbookXml),
    `${label}: faltan bookViews en xl/workbook.xml.`,
  );
}

function assertStylesXmlReadable(buffer, label) {
  const zip = new AdmZip(buffer);
  const stylesXml = readZipEntryText(zip, "xl/styles.xml");
  assertCondition(/<styleSheet\b/i.test(stylesXml), `${label}: falta styleSheet en xl/styles.xml.`);
  assertCondition(/<fonts\b/i.test(stylesXml), `${label}: falta bloque fonts en xl/styles.xml.`);
}

async function assertExcelJsReadable(buffer, label) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  assertCondition(workbook.worksheets.length > 0, `${label}: ExcelJS no detecto hojas.`);
}

async function loadWorkbook(source) {
  const workbook = new ExcelJS.Workbook();
  if (Buffer.isBuffer(source)) {
    await workbook.xlsx.load(source);
  } else {
    await workbook.xlsx.readFile(source);
  }
  return workbook;
}

function normalizeComparable(value) {
  if (Array.isArray(value)) {
    return value.map((item) => normalizeComparable(item));
  }
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === "string") {
    return normalizeTextComparison(value);
  }
  if (typeof value === "number") {
    return round2(value);
  }
  if (value && typeof value === "object") {
    return Object.keys(value)
      .sort()
      .reduce((result, key) => {
        result[key] = normalizeComparable(value[key]);
        return result;
      }, {});
  }
  return value ?? null;
}

function getFormula(value) {
  if (!value || typeof value !== "object") {
    return null;
  }
  if (Object.prototype.hasOwnProperty.call(value, "formula")) {
    return value.formula || null;
  }
  if (Object.prototype.hasOwnProperty.call(value, "sharedFormula")) {
    return `SHARED:${value.sharedFormula}`;
  }
  return null;
}

function isDateFormat(numFmt) {
  return typeof numFmt === "string" && /[dmy]/i.test(numFmt);
}

function excelSerialToIsoDate(serial) {
  const numeric = Number(serial);
  if (!Number.isFinite(numeric)) {
    return null;
  }
  const utcDays = Math.floor(numeric - 25569);
  const utcValue = utcDays * 86400;
  const dateInfo = new Date(utcValue * 1000);
  if (!Number.isFinite(dateInfo.getTime())) {
    return null;
  }
  return dateInfo.toISOString().slice(0, 10);
}

function normalizeCellValue(cell, formulaNormalizer = (formula) => formula) {
  const formula = getFormula(cell.value);
  if (formula) {
    return { formula: formulaNormalizer(formula) };
  }
  if (cell.value instanceof Date) {
    return cell.value.toISOString().slice(0, 10);
  }
  if ((typeof cell.value === "number" || typeof cell.value === "string") && isDateFormat(cell.numFmt)) {
    return excelSerialToIsoDate(cell.value) || cell.value;
  }
  return normalizeComparable(cell.value);
}

function normalizeSheetDisplayValue(cell, columnIndex) {
  const value = cell?.v ?? null;
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  if (columnIndex === 4 && (typeof value === "string" || typeof value === "number")) {
    const normalizedDate = excelSerialToIsoDate(value);
    if (normalizedDate) {
      return normalizedDate;
    }
  }
  if (typeof value === "number") {
    return round2(value);
  }
  if (typeof value === "string") {
    return normalizeTextComparison(value);
  }
  return value;
}

function action1ValueSignature(source) {
  const workbook = Buffer.isBuffer(source)
    ? XLSX.read(source, { type: "buffer", cellFormula: true, cellDates: true })
    : XLSX.readFile(source, { cellFormula: true, cellDates: true });
  const sheet = workbook.Sheets["LIBRO COMPRAS"] || workbook.Sheets[workbook.SheetNames[0]];
  assertCondition(!!sheet, "Accion 1: no existe hoja LIBRO COMPRAS.");

  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:L1");
  let lastMeaningful = 0;
  for (let row = 0; row <= range.e.r; row += 1) {
    let meaningful = false;
    for (let col = 0; col < 12; col += 1) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
      if (cell && normalizeSheetDisplayValue(cell, col + 1) != null && sanitizeText(normalizeSheetDisplayValue(cell, col + 1)) !== "") {
        meaningful = true;
        break;
      }
    }
    if (meaningful) {
      lastMeaningful = row;
    }
  }

  const rows = [];
  for (let row = 0; row <= lastMeaningful; row += 1) {
    const values = [];
    for (let col = 0; col < 12; col += 1) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
      values.push(JSON.stringify(normalizeSheetDisplayValue(cell, col + 1)));
    }
    rows.push(values.join("|"));
  }
  return rows.join("\n");
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function quoteSheetReference(sheetName) {
  return /[\s']/.test(sheetName) ? `'${sheetName.replace(/'/g, "''")}'!` : `${sheetName}!`;
}

function rewriteFormulaSheetReference(formula, sourceSheetName, targetSheetName) {
  let updated = formula;
  const replacement = quoteSheetReference(targetSheetName);
  const quotedExternalPattern = new RegExp(`'(?:\\[[^\\]]+\\])?${escapeRegExp(sourceSheetName)}'!`, "g");
  updated = updated.replace(quotedExternalPattern, replacement);
  if (sourceSheetName !== targetSheetName) {
    const quotedInternalPattern = new RegExp(`'${escapeRegExp(sourceSheetName)}'!`, "g");
    updated = updated.replace(quotedInternalPattern, replacement);
  }
  return updated;
}

function normalizeAction2BundleFormula(formula) {
  if (!formula) {
    return formula;
  }
  return rewriteFormulaSheetReference(formula, "MAYOR RET", "ACCION 3 MAYOR RET");
}

function normalizeBundleFormula(formula) {
  if (!formula) {
    return formula;
  }

  const sheetMap = new Map([
    ["LIBRO COMPRAS", "ACCION 1 LIBRO COMPRAS"],
    ["RET PROV", "ACCION 2 RET PROV"],
    ["MAYOR RET", "ACCION 3 MAYOR RET"],
    ["MAYOR IVA", "ACCION 4 MAYOR IVA"],
  ]);

  let updated = formula;
  for (const [sourceSheet, targetSheet] of sheetMap.entries()) {
    updated = rewriteFormulaSheetReference(updated, sourceSheet, targetSheet);
  }
  return updated;
}

function lastMeaningfulRow(worksheet, startCol, endCol) {
  let last = 1;
  for (let rowIndex = 1; rowIndex <= worksheet.rowCount; rowIndex += 1) {
    let meaningful = false;
    for (let col = startCol; col <= endCol; col += 1) {
      const value = worksheet.getRow(rowIndex).getCell(col).value;
      if (getFormula(value)) {
        meaningful = true;
        break;
      }
      if (value != null && sanitizeText(value) !== "") {
        meaningful = true;
        break;
      }
    }
    if (meaningful) {
      last = rowIndex;
    }
  }
  return last;
}

function worksheetValueSignature(worksheet, startCol, endCol, formulaNormalizer = (formula) => formula) {
  const rows = [];
  const lastRow = lastMeaningfulRow(worksheet, startCol, endCol);
  for (let rowIndex = 1; rowIndex <= lastRow; rowIndex += 1) {
    const values = [];
    for (let col = startCol; col <= endCol; col += 1) {
      const cell = worksheet.getRow(rowIndex).getCell(col);
      values.push(JSON.stringify(normalizeCellValue(cell, formulaNormalizer)));
    }
    rows.push(values.join("|"));
  }
  return rows.join("\n");
}

function assertWorksheetValuesMatch(expectedWs, outputWs, startCol, endCol, label, formulaNormalizer = (formula) => formula) {
  const expectedSignature = worksheetValueSignature(expectedWs, startCol, endCol, formulaNormalizer);
  const outputSignature = worksheetValueSignature(outputWs, startCol, endCol, formulaNormalizer);
  assertCondition(expectedSignature === outputSignature, `${label}: los valores/celdas no coinciden.`);
}

function detailRowSignatures(worksheet, startCol, endCol) {
  const rows = [];
  const lastRow = lastMeaningfulRow(worksheet, startCol, endCol);
  for (let rowIndex = 2; rowIndex <= lastRow; rowIndex += 1) {
    const anchor = worksheet.getRow(rowIndex).getCell(startCol).value;
    if (anchor == null || sanitizeText(anchor) === "") {
      continue;
    }
    const values = [];
    for (let col = startCol; col <= endCol; col += 1) {
      const cell = worksheet.getRow(rowIndex).getCell(col);
      values.push(JSON.stringify(normalizeCellValue(cell)));
    }
    rows.push(values.join("|"));
  }
  return rows.sort();
}

function assertDetailRowSetMatch(expectedWs, outputWs, startCol, endCol, label) {
  const expectedRows = detailRowSignatures(expectedWs, startCol, endCol);
  const outputRows = detailRowSignatures(outputWs, startCol, endCol);
  assertCondition(expectedRows.length === outputRows.length, `${label}: cantidad de movimientos distinta.`);
  assertCondition(expectedRows.join("\n") === outputRows.join("\n"), `${label}: los movimientos no coinciden.`);
}

function parseAuditReportDate(value) {
  const text = normalizeTextComparison(value);
  const match = /^(\d{2})-([A-Z]{3})-(\d{2})$/.exec(text || "");
  if (!match) {
    return null;
  }

  const monthMap = {
    JAN: 1,
    ENE: 1,
    FEB: 2,
    MAR: 3,
    APR: 4,
    ABR: 4,
    MAY: 5,
    JUN: 6,
    JUL: 7,
    AUG: 8,
    AGO: 8,
    SEP: 9,
    SET: 9,
    OCT: 10,
    NOV: 11,
    DEC: 12,
    DIC: 12,
  };

  const month = monthMap[match[2]];
  if (!month) {
    return null;
  }
  return `20${match[3]}-${String(month).padStart(2, "0")}-${match[1]}`;
}

function normalizeAction3MonthlyCell(cell, columnIndex) {
  const formula = getFormula(cell.value);
  if (formula) {
    return { formula };
  }
  if (cell.value instanceof Date) {
    return cell.value.toISOString().slice(0, 10);
  }
  if (columnIndex === 4 && typeof cell.value === "string") {
    return parseAuditReportDate(cell.value) || normalizeTextComparison(cell.value);
  }
  if (columnIndex === 7) {
    const text = sanitizeText(cell.value);
    return /^\d+(?:\.0+)?$/.test(text) ? String(parseInt(text, 10)) : normalizeTextComparison(text);
  }
  return normalizeComparable(cell.value);
}

function action3MonthlyDetailRowSignatures(worksheet) {
  const rows = [];
  for (let rowIndex = 1; rowIndex <= worksheet.rowCount; rowIndex += 1) {
    const code = normalizeTextComparison(worksheet.getRow(rowIndex).getCell(1).value);
    if (!/^\d{2}\.\d{2}\.\d{2}\.\d{2}\.\d{4}$/.test(code || "")) {
      continue;
    }

    const values = [];
    for (let col = 1; col <= 11; col += 1) {
      values.push(JSON.stringify(normalizeAction3MonthlyCell(worksheet.getRow(rowIndex).getCell(col), col)));
    }
    rows.push(values.join("|"));
  }
  return rows.sort();
}

function assertAction3MonthlyDetailMatch(expectedWs, outputWs, label) {
  const expectedRows = action3MonthlyDetailRowSignatures(expectedWs);
  const outputRows = action3MonthlyDetailRowSignatures(outputWs);
  assertCondition(expectedRows.length === outputRows.length, `${label}: cantidad de movimientos distinta.`);
  assertCondition(expectedRows.join("\n") === outputRows.join("\n"), `${label}: los movimientos no coinciden.`);
}

function normalizeAction4MonthlyCell(cell, columnIndex) {
  const formula = getFormula(cell.value);
  if (formula) {
    return { formula };
  }
  if (cell.value instanceof Date) {
    return cell.value.toISOString().slice(0, 10);
  }
  if (columnIndex === 4 && typeof cell.value === "string") {
    return parseAuditReportDate(cell.value) || normalizeTextComparison(cell.value);
  }
  if (columnIndex === 7) {
    const text = sanitizeText(cell.value);
    return /^\d+(?:\.0+)?$/.test(text) ? String(parseInt(text, 10)) : normalizeTextComparison(text);
  }
  return normalizeComparable(cell.value);
}

function worksheetValueSignatureByCellNormalizer(worksheet, startCol, endCol, cellNormalizer) {
  const rows = [];
  const lastRow = lastMeaningfulRow(worksheet, startCol, endCol);
  for (let rowIndex = 1; rowIndex <= lastRow; rowIndex += 1) {
    const values = [];
    for (let col = startCol; col <= endCol; col += 1) {
      values.push(JSON.stringify(cellNormalizer(worksheet.getRow(rowIndex).getCell(col), col)));
    }
    rows.push(values.join("|"));
  }
  return rows.join("\n");
}

function assertAction4MonthlyWorksheetMatch(expectedWs, outputWs, label) {
  const expectedSignature = worksheetValueSignatureByCellNormalizer(expectedWs, 1, 16, normalizeAction4MonthlyCell);
  const outputSignature = worksheetValueSignatureByCellNormalizer(outputWs, 1, 16, normalizeAction4MonthlyCell);
  assertCondition(expectedSignature === outputSignature, `${label}: los valores/celdas no coinciden.`);
}

function assertFormulaMapMatches(expectedWs, outputWs, label) {
  const mismatches = [];
  expectedWs.eachRow((row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      const expectedFormula = getFormula(cell.value);
      if (!expectedFormula) {
        return;
      }
      const outputFormula = getFormula(outputWs.getCell(cell.address).value);
      if (expectedFormula !== outputFormula) {
        mismatches.push(`${cell.address}: ${expectedFormula} <> ${outputFormula || "null"}`);
      }
    });
  });
  assertCondition(mismatches.length === 0, `${label}: formulas distintas. ${mismatches.slice(0, 5).join(" | ")}`);
}

function normalizeAction2SummaryLabel(label) {
  return normalizeTextComparison(label)?.toUpperCase() || null;
}

function formatAction2PercentLabel(value) {
  if (Number.isInteger(value)) {
    return String(value);
  }
  const text = String(value);
  return text.replace(/(\.\d*?[1-9])0+$/g, "$1").replace(/\.0+$/g, "");
}

function toCents(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return 0;
  }
  return Math.round((numeric + Number.EPSILON) * 100);
}

function fromCents(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return 0;
  }
  return numeric / 100;
}

function buildAction2SummaryEntries(rows) {
  const typeOrder = ["IVA", "RENTA"];
  const typeMap = new Map();
  let totalBaseCents = 0;
  let totalRetCents = 0;

  for (const row of rows) {
    if (Number(row?.numRt) === 999999999) {
      continue;
    }

    const type = normalizeTextComparison(row.tipo);
    if (!typeMap.has(type)) {
      typeMap.set(type, {
        totalBaseCents: 0,
        totalRetCents: 0,
        percentMap: new Map(),
      });
    }

    const bucket = typeMap.get(type);
    const baseCents = toCents(row.base);
    const retCents = toCents(row.retencion);
    bucket.totalBaseCents += baseCents;
    bucket.totalRetCents += retCents;
    totalBaseCents += baseCents;
    totalRetCents += retCents;

    const percentKey = round2(row.percent);
    if (!bucket.percentMap.has(percentKey)) {
      bucket.percentMap.set(percentKey, { baseCents: 0, retCents: 0 });
    }
    const percentBucket = bucket.percentMap.get(percentKey);
    percentBucket.baseCents += baseCents;
    percentBucket.retCents += retCents;
  }

  const entries = [];
  const seenTypes = new Set();
  for (const type of typeOrder) {
    if (!typeMap.has(type)) {
      continue;
    }
    seenTypes.add(type);
    const bucket = typeMap.get(type);
    entries.push({ label: type, base: fromCents(bucket.totalBaseCents), ret: fromCents(bucket.totalRetCents) });

    const percents = [...bucket.percentMap.keys()].sort((a, b) => a - b);
    for (const percent of percents) {
      const item = bucket.percentMap.get(percent);
      entries.push({
        label: formatAction2PercentLabel(percent),
        base: fromCents(item.baseCents),
        ret: fromCents(item.retCents),
      });
    }
  }

  const remainingTypes = [...typeMap.keys()].filter((type) => !seenTypes.has(type)).sort();
  for (const type of remainingTypes) {
    const bucket = typeMap.get(type);
    entries.push({ label: type, base: fromCents(bucket.totalBaseCents), ret: fromCents(bucket.totalRetCents) });

    const percents = [...bucket.percentMap.keys()].sort((a, b) => a - b);
    for (const percent of percents) {
      const item = bucket.percentMap.get(percent);
      entries.push({
        label: formatAction2PercentLabel(percent),
        base: fromCents(item.baseCents),
        ret: fromCents(item.retCents),
      });
    }
  }

  entries.push({ label: "Total general", base: fromCents(totalBaseCents), ret: fromCents(totalRetCents) });
  return entries;
}

function assertAction2SummaryMatches(outputWs, rows, label) {
  const summaryLookup = new Map(
    buildAction2SummaryEntries(rows).map((item) => [normalizeAction2SummaryLabel(item.label), item]),
  );

  let started = false;
  for (let rowIndex = 2; rowIndex <= 200; rowIndex += 1) {
    const rawLabel = outputWs.getCell(rowIndex, 12).value;
    const normalizedLabel = normalizeAction2SummaryLabel(rawLabel);
    if (!normalizedLabel) {
      if (started) {
        break;
      }
      continue;
    }
    started = true;
    const expected = summaryLookup.get(normalizedLabel) || { base: 0, ret: 0 };
    const actualBase = normalizeCellValue(outputWs.getCell(rowIndex, 13));
    const actualRet = normalizeCellValue(outputWs.getCell(rowIndex, 14));
    assertCondition(
      JSON.stringify(actualBase) === JSON.stringify(round2(expected.base)),
      `${label}: base lateral incorrecta en fila ${rowIndex}.`,
    );
    assertCondition(
      JSON.stringify(actualRet) === JSON.stringify(round2(expected.ret)),
      `${label}: retencion lateral incorrecta en fila ${rowIndex}.`,
    );
  }
}

async function verifyWorkbookBuffer(buffer, label, verifyFn, ...args) {
  const outputPath = path.join(os.tmpdir(), `${label}_${Date.now()}_${Math.random().toString(16).slice(2)}.xlsx`);
  fs.writeFileSync(outputPath, buffer);
  try {
    await verifyFn(outputPath, ...args);
  } finally {
    if (fs.existsSync(outputPath)) {
      fs.unlinkSync(outputPath);
    }
  }
}

async function buildAction2Expected(inputPath, mode = "legacy") {
  const templatePath = mode === "monthly"
    ? actionMonthlyPath([/PLANTILLAHECHAAMANO.*\.xlsx$/i], "accion2-mensual-plantilla")
    : actionTemplatePath("ACCION2.xlsx");
  const rows = normalizeParsedRows(
    extractRowsFromTxt(inputPath),
    loadTemplateTipoHints(templatePath),
  );
  const { workbook, summary } = await buildAction2Workbook(templatePath, rows);
  return { rows, templatePath, workbook, summary };
}

async function buildAction3Expected(inputPaths, mode = "legacy") {
  if (mode === "monthly") {
    const templatePath = actionMonthlyPath([/PLANTILLAHECHAAMANO.*\.xlsx$/i], "accion3-mensual-plantilla");
    const workbook = await loadWorkbook(templatePath);
    return { rows: [], templatePath, workbook };
  }

  const parsed = await parseAction3Inputs(inputPaths);
  const templatePath = actionTemplatePath("MAYOR RET_ACCION3.xlsx");
  const dropAgcmFromRender = true;
  const { workbook } = await buildAction3Workbook(templatePath, parsed.rows, { dropAgcmFromRender });
  return { rows: parsed.rows, templatePath, workbook };
}

async function buildAction4Expected(inputPath, mode = "legacy") {
  if (mode === "monthly") {
    const templatePath = actionMonthlyPath([/PLANTILLAHECHAAMANO.*\.xlsx$/i], "accion4-mensual-plantilla");
    const workbook = await loadWorkbook(templatePath);
    return { parsed: null, plan: null, templatePath, workbook };
  }

  const parsed = await parseAction4Inputs([inputPath]);
  const plan = buildAction4OutputPlan(parsed.rows);
  const templatePath = actionTemplatePath("MAYORIVAACCION4.xlsx");
  const { workbook } = await buildAction4Workbook(templatePath, plan.rowPlan, plan.movementRows);
  return { parsed, plan, templatePath, workbook };
}

async function assertOutputHealth(buffer, label) {
  assertNoExternalLinks(buffer, label);
  assertWorkbookViewStructure(buffer, label);
  assertStylesXmlReadable(buffer, label);
  await assertExcelJsReadable(buffer, label);
}

async function runAction1Scenario(label, inputPath, referencePath, mode = "legacy") {
  const outputBuffer = await submitModuleFiles("accion1", [
    { field: "source_files", path: inputPath, contentType: "application/pdf" },
  ]);
  await assertOutputHealth(outputBuffer, label);

  if (mode === "monthly") {
    const outputWorkbook = await loadWorkbook(outputBuffer);
    const expectedWorkbook = await loadWorkbook(referencePath);
    const outputWs = outputWorkbook.getWorksheet("LIBRO COMPRAS");
    const expectedWs = expectedWorkbook.getWorksheet("LIBRO COMPRAS");
    assertCondition(!!outputWs, `${label}: no existe hoja LIBRO COMPRAS en salida.`);
    assertCondition(!!expectedWs, `${label}: no existe hoja LIBRO COMPRAS en referencia.`);
    assertWorksheetValuesMatch(expectedWs, outputWs, 1, 12, label);
  } else {
    const outputSignature = action1ValueSignature(outputBuffer);
    const expectedSignature = action1ValueSignature(referencePath);
    assertCondition(outputSignature === expectedSignature, `${label}: los valores efectivos del libro no coinciden.`);
  }

  console.log(`OK: ${label}`);
  return outputBuffer;
}

async function runAction2Scenario(label, inputPath, mode = "legacy") {
  const expected = await buildAction2Expected(inputPath, mode);
  const outputBuffer = await submitModuleFiles("accion2", [
    { field: "source_files", path: inputPath, contentType: "text/plain" },
  ]);
  await assertOutputHealth(outputBuffer, label);

  const outputWorkbook = await loadWorkbook(outputBuffer);
  const outputWs = outputWorkbook.getWorksheet("RET PROV");
  const expectedWs = expected.workbook.getWorksheet("RET PROV");
  assertCondition(!!outputWs, `${label}: no existe hoja RET PROV.`);
  assertCondition(!!expectedWs, `${label}: no existe hoja RET PROV esperada.`);
  assertWorksheetValuesMatch(expectedWs, outputWs, 1, 11, `${label}:detalle`);
  assertAction2SummaryMatches(outputWs, expected.rows, `${label}:resumen`);
  assertFormulaMapMatches(expectedWs, outputWs, `${label}:formulas`);

  console.log(`OK: ${label}`);
  return outputBuffer;
}

async function runAction3Scenario(label, inputPaths, contentType, mode = "legacy") {
  const expected = await buildAction3Expected(inputPaths, mode);
  const outputBuffer = await submitModuleFiles(
    "accion3",
    inputPaths.map((inputPath) => ({
      field: "source_files[]",
      path: inputPath,
      contentType,
    })),
  );
  await assertOutputHealth(outputBuffer, label);

  const outputWorkbook = await loadWorkbook(outputBuffer);
  const outputWs = outputWorkbook.getWorksheet("MAYOR RET");
  const expectedWs = expected.workbook.getWorksheet("MAYOR RET");
  assertCondition(!!outputWs, `${label}: no existe hoja MAYOR RET.`);
  assertCondition(!!expectedWs, `${label}: no existe hoja MAYOR RET esperada.`);
  if (mode === "monthly") {
    assertAction3MonthlyDetailMatch(expectedWs, outputWs, `${label}:detalle`);
  } else {
    assertDetailRowSetMatch(expectedWs, outputWs, 1, 11, `${label}:detalle`);
  }
  assertWorksheetValuesMatch(expectedWs, outputWs, 13, 16, `${label}:resumen`);

  console.log(`OK: ${label}`);
  return outputBuffer;
}

async function runAction4Scenario(label, inputPath, mode = "legacy") {
  const expected = await buildAction4Expected(inputPath, mode);
  const outputBuffer = await submitModuleFiles("accion4", [
    { field: "source_files", path: inputPath, contentType: "text/plain" },
  ]);
  await assertOutputHealth(outputBuffer, label);

  const outputWorkbook = await loadWorkbook(outputBuffer);
  const outputWs = outputWorkbook.getWorksheet("MAYOR IVA");
  const expectedWs = expected.workbook.getWorksheet("MAYOR IVA");
  assertCondition(!!outputWs, `${label}: no existe hoja MAYOR IVA.`);
  assertCondition(!!expectedWs, `${label}: no existe hoja MAYOR IVA esperada.`);
  if (mode === "monthly") {
    assertAction4MonthlyWorksheetMatch(expectedWs, outputWs, label);
  } else {
    assertWorksheetValuesMatch(expectedWs, outputWs, 1, 16, label);
  }

  console.log(`OK: ${label}`);
  return outputBuffer;
}

async function runBundleScenario(label, standaloneOutputs) {
  const outputsDir = path.join(ROOT, "storage", "outputs");
  const probeStamp = `${Date.now()}`;
  const probeFiles = [
    { key: "accion1", path: path.join(outputsDir, `bundle_probe_${label}_${probeStamp}_resultado.xlsx`) },
    { key: "accion2", path: path.join(outputsDir, `bundle_probe_${label}_${probeStamp}_20991231_235959_accion2.xlsx`) },
    { key: "accion3", path: path.join(outputsDir, `bundle_probe_${label}_${probeStamp}_20991231_235959_accion3.xlsx`) },
    { key: "accion4", path: path.join(outputsDir, `bundle_probe_${label}_${probeStamp}_accion4.xlsx`) },
  ];

  for (const item of probeFiles) {
    fs.writeFileSync(item.path, standaloneOutputs[item.key]);
  }

  let bundleBuffer;
  try {
    bundleBuffer = await submitBundleModule();
    await assertOutputHealth(bundleBuffer, label);
  } finally {
    for (const item of probeFiles) {
      if (fs.existsSync(item.path)) {
        fs.unlinkSync(item.path);
      }
    }
  }

  const bundleWorkbook = await loadWorkbook(bundleBuffer);
  assertCondition(bundleWorkbook.worksheets.length === 4, `${label}: el consolidado debe tener 4 hojas.`);

  const sheetExpectations = [
    {
      bundleSheet: "ACCION 1 LIBRO COMPRAS",
      standaloneSheet: "LIBRO COMPRAS",
      outputBuffer: standaloneOutputs.accion1,
      startCol: 1,
      endCol: 12,
      formulaNormalizer: normalizeBundleFormula,
    },
    {
      bundleSheet: "ACCION 2 RET PROV",
      standaloneSheet: "RET PROV",
      outputBuffer: standaloneOutputs.accion2,
      startCol: 1,
      endCol: 16,
      formulaNormalizer: normalizeBundleFormula,
    },
    {
      bundleSheet: "ACCION 3 MAYOR RET",
      standaloneSheet: "MAYOR RET",
      outputBuffer: standaloneOutputs.accion3,
      startCol: 1,
      endCol: 16,
      formulaNormalizer: normalizeBundleFormula,
    },
    {
      bundleSheet: "ACCION 4 MAYOR IVA",
      standaloneSheet: "MAYOR IVA",
      outputBuffer: standaloneOutputs.accion4,
      startCol: 1,
      endCol: 16,
      formulaNormalizer: normalizeBundleFormula,
    },
  ];

  for (const item of sheetExpectations) {
    const standaloneWorkbook = await loadWorkbook(item.outputBuffer);
    const expectedWs = standaloneWorkbook.getWorksheet(item.standaloneSheet);
    const outputWs = bundleWorkbook.getWorksheet(item.bundleSheet);
    assertCondition(!!expectedWs, `${label}: no existe hoja standalone ${item.standaloneSheet}.`);
    assertCondition(!!outputWs, `${label}: no existe hoja bundle ${item.bundleSheet}.`);
    assertWorksheetValuesMatch(
      expectedWs,
      outputWs,
      item.startCol,
      item.endCol,
      `${label}:${item.bundleSheet}`,
      item.formulaNormalizer,
    );
  }

  const standaloneAction2Workbook = await loadWorkbook(standaloneOutputs.accion2);
  const standaloneAction2Ws = standaloneAction2Workbook.getWorksheet("RET PROV");
  const action2Ws = bundleWorkbook.getWorksheet("ACCION 2 RET PROV");
  const rewrittenCells = ["O3", "O4", "O5", "O6", "O7", "O10", "O11", "O12", "O13", "O14", "O15"];
  for (const cellAddress of rewrittenCells) {
    const expectedFormula = getFormula(standaloneAction2Ws.getCell(cellAddress).value);
    const formula = getFormula(action2Ws.getCell(cellAddress).value);
    if (!expectedFormula) {
      assertCondition(formula == null, `${label}: ${cellAddress} deberia permanecer sin formula.`);
      continue;
    }
    assertCondition(typeof formula === "string" && !formula.includes("["), `${label}: ${cellAddress} conserva referencia externa.`);
    assertCondition(
      formula.includes("'ACCION 3 MAYOR RET'!"),
      `${label}: ${cellAddress} no apunta a ACCION 3 MAYOR RET.`,
    );
  }

  console.log(`OK: ${label}`);
}

async function main() {
  const promptPath = path.join(ROOT, "docs", "PROMPT_AUDITORIA_TOTAL_ARCHIVOS_VENTANA1_2026-03-25.md");
  assertCondition(fs.existsSync(promptPath), "Falta el prompt de auditoria total de Ventana 1.");

  const scenarios = {
    legacy: {
      action1Input: actionFixturePath("CXPREP_docproveedor.pdf"),
      action1Reference: actionContractPath("CXPREP_docproveedor_20260306_092810_resultado.xlsx"),
      action2Input: actionFixturePath("CXPREP_RET_GENERALACCION2.txt"),
      action3TxtInputs: [actionFixturePath("CON_MAYORGEN2ACCION3.txt")],
      action3PdfInputs: [actionFixturePath("CON_MAYORGEN2ACCION3.pdf")],
      action4Input: actionFixturePath("CON_MAYORGEN2IVAACCION4.TXT"),
    },
    monthly: {
      action1Input: actionMonthlyPath([/^CXPREP_docproveedor.*\.pdf$/i], "accion1-mensual-pdf"),
      action1Reference: actionMonthlyPath([/PLANTILLAHECHAAMANO.*\.xlsx$/i], "accion1-mensual-plantilla"),
      action2Input: actionMonthlyPath([/^CXPREP_RET_GENERAL.*\.txt$/i], "accion2-mensual-txt"),
      action3TxtInputs: [
        actionMonthlyPath([/^CON_MAYORGEN2RETACCION3\.TXT$/i], "accion3-mensual-txt-1"),
        actionMonthlyPath([/^CON_MAYORGEN2RET2ACCION3\.TXT$/i], "accion3-mensual-txt-2"),
      ],
      action4Input: actionMonthlyPath([/^CON_MAYORGEN2IVAENERO\.TXT$/i], "accion4-mensual-txt"),
    },
  };

  console.log(`PROMPT: ${promptPath}`);
  console.log("FILES LEGACY:");
  Object.values(scenarios.legacy).flat().forEach((file) => console.log(` - ${file}`));
  console.log("FILES MONTHLY:");
  Object.values(scenarios.monthly).flat().forEach((file) => console.log(` - ${file}`));

  const server = startPhpServer();
  try {
    await waitServerReady();

    const legacyOutputs = {};
    legacyOutputs.accion1 = await runAction1Scenario("accion1-legacy", scenarios.legacy.action1Input, scenarios.legacy.action1Reference, "legacy");
    legacyOutputs.accion2 = await runAction2Scenario("accion2-legacy", scenarios.legacy.action2Input, "legacy");
    legacyOutputs.accion3 = await runAction3Scenario("accion3-legacy-txt", scenarios.legacy.action3TxtInputs, "text/plain", "legacy");
    legacyOutputs.accion4 = await runAction4Scenario("accion4-legacy", scenarios.legacy.action4Input, "legacy");
    await runBundleScenario("bundle-legacy", legacyOutputs);
    await runAction3Scenario("accion3-legacy-pdf", scenarios.legacy.action3PdfInputs, "application/pdf", "legacy");

    const monthlyOutputs = {};
    monthlyOutputs.accion1 = await runAction1Scenario("accion1-monthly", scenarios.monthly.action1Input, scenarios.monthly.action1Reference, "monthly");
    monthlyOutputs.accion2 = await runAction2Scenario("accion2-monthly", scenarios.monthly.action2Input, "monthly");
    monthlyOutputs.accion3 = await runAction3Scenario("accion3-monthly-txt-multi", scenarios.monthly.action3TxtInputs, "text/plain", "monthly");
    monthlyOutputs.accion4 = await runAction4Scenario("accion4-monthly", scenarios.monthly.action4Input, "monthly");
    await runBundleScenario("bundle-monthly", monthlyOutputs);

    console.log("OK: auditoria total de archivos legacy y mensuales de Ventana 1 validada.");
  } finally {
    server.kill();
  }
}

main().catch((error) => {
  console.error(`FAIL: ${error.message}`);
  process.exit(1);
});
