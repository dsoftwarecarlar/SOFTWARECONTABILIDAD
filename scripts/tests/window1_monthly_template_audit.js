const fs = require("fs");
const http = require("http");
const os = require("os");
const path = require("path");
const { spawn } = require("child_process");

const ExcelJS = require("exceljs");

const {
  extractRowsFromTxt,
  loadTemplateTipoHints,
  normalizeParsedRows,
} = require("../cxp/accion2/parser");
const accion2Constants = require("../cxp/accion2/constants");
const { parseInputSources: parseAction3InputSources } = require("../cxp/accion3/parser");
const {
  parseInputSources: parseAction4InputSources,
  buildAction4OutputPlan,
} = require("../cxp/accion4/parser");

const ROOT = path.resolve(__dirname, "..", "..");
const APP_ROOT = path.join(ROOT, "laravel_app");
const ROUTER = path.join(APP_ROOT, "server.php");
const PHP_EXE = firstExistingPath([
  path.join(ROOT, ".tools", "php82", "php.exe"),
  process.env.PHP_EXE || "",
  "C:\\xampp\\php\\php.exe",
]);
const HOST = "127.0.0.1";
const PORT = 18983;
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

function actionMonthlyPath(matchers, label) {
  const monthlyDir = path.join(ROOT, "resources", "cxp", "acciones", "PLANTILLAYARCHIVOS");
  assertCondition(fs.existsSync(monthlyDir), `${label}: no existe PLANTILLAYARCHIVOS.`);
  const entries = fs.readdirSync(monthlyDir);
  const found = entries.find((name) => matchers.some((matcher) => matcher.test(name)));
  assertCondition(!!found, `${label}: no se encontro archivo mensual esperado en PLANTILLAYARCHIVOS.`);
  return path.join(monthlyDir, found);
}

function round2(value) {
  return Math.round((Number(value || 0) + Number.EPSILON) * 100) / 100;
}

function sanitizeText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function formatSlashDate(value) {
  const date = value instanceof Date ? value : new Date(value);
  return [
    String(date.getDate()).padStart(2, "0"),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getFullYear()),
  ].join("/");
}

function formatReportDate(value) {
  const date = value instanceof Date ? value : new Date(value);
  const monthNames = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"];
  return `${String(date.getDate()).padStart(2, "0")}-${monthNames[date.getMonth()]}-${String(date.getFullYear()).slice(-2)}`;
}

function formatReportAmount(value) {
  return round2(value).toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function mutateDocu(value, index) {
  const text = sanitizeText(value);
  if (/^\d+$/.test(text)) {
    return String(Number(text) + 8000 + index);
  }
  return `${text}X${index + 1}`;
}

function mutateAction2Row(row, index) {
  return {
    ...row,
    numRt: Number(row.numRt) + 800000 + index,
    proveedor: `${sanitizeText(row.proveedor)} EXTRA ${index + 1}`.slice(0, 120),
    fact: typeof row.fact === "number" ? Number(row.fact) + 9000 + index : `${sanitizeText(row.fact)}-${index + 1}`,
  };
}

function mutateMovementRow(row, index) {
  return {
    ...row,
    ASIENTO: Number(row.ASIENTO || 0) + 7000 + index,
    DOCU: mutateDocu(row.DOCU, index),
    DETALLE: `${sanitizeText(row.DETALLE)} EXTRA ${index + 1}`.slice(0, 120),
  };
}

function buildAction2InputFile(rows, label) {
  const header = "NUM RT\tPROVEEDOR\tFECHA\tFECHA CONT\tTIPO\tCOD\tFACT\t%\tBASE\tRETENCION";
  const lines = rows.map((row) => [
    row.numRt,
    sanitizeText(row.proveedor),
    formatSlashDate(row.fecha),
    formatSlashDate(row.fechaCont),
    row.tipo,
    row.cod,
    row.fact,
    String(row.percent),
    String(round2(row.base)),
    String(round2(row.retencion)),
  ].join("\t"));

  const filePath = path.join(os.tmpdir(), `accion2_${label}_${Date.now()}.txt`);
  fs.writeFileSync(filePath, [header, ...lines].join("\r\n"), "utf8");
  return filePath;
}

function deriveOpeningBalance(rows) {
  assertCondition(rows.length > 0, "No hay filas para derivar saldo inicial.");
  const first = rows[0];
  return round2(Number(first.SALDO || 0) - Number(first.DEBE || 0) + Number(first.HABER || 0));
}

function buildSingleAccountMayorTxtFile(rows, label) {
  const openingBalance = deriveOpeningBalance(rows);
  const totalDebe = round2(rows.reduce((sum, row) => sum + Number(row.DEBE || 0), 0));
  const totalHaber = round2(rows.reduce((sum, row) => sum + Number(row.HABER || 0), 0));
  const closingBalance = round2(openingBalance + totalDebe - totalHaber);
  let runningBalance = openingBalance;

  const lines = rows.map((row) => {
    runningBalance = round2(runningBalance + Number(row.DEBE || 0) - Number(row.HABER || 0));
    const cols = new Array(30).fill("");
    cols[6] = sanitizeText(row.COD);
    cols[7] = sanitizeText(row.CUENTA);
    cols[8] = formatReportAmount(openingBalance);
    cols[9] = formatReportAmount(totalDebe);
    cols[10] = formatReportAmount(totalHaber);
    cols[11] = formatReportAmount(closingBalance);
    cols[21] = sanitizeText(row.EXT) || "N";
    cols[22] = formatReportDate(row.FECHA);
    cols[23] = sanitizeText(row.ORIGEN);
    cols[24] = String(Number(row.ASIENTO || 0));
    cols[25] = sanitizeText(row.DOCU);
    cols[26] = sanitizeText(row.DETALLE).replace(/\t/g, " ");
    cols[27] = formatReportAmount(row.DEBE);
    cols[28] = formatReportAmount(row.HABER);
    cols[29] = formatReportAmount(runningBalance);
    return cols.join("\t");
  });

  const filePath = path.join(os.tmpdir(), `${label}_${Date.now()}.txt`);
  fs.writeFileSync(filePath, lines.join("\r\n"), "utf8");
  return filePath;
}

function buildAction3TxtFile(rows, label) {
  const groups = new Map();

  for (const row of rows) {
    const key = `${sanitizeText(row.COD)}|${sanitizeText(row.CUENTA)}`;
    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key).push(row);
  }

  const lines = [];
  for (const groupRows of groups.values()) {
    const openingBalance = deriveOpeningBalance(groupRows);
    const totalDebe = round2(groupRows.reduce((sum, row) => sum + Number(row.DEBE || 0), 0));
    const totalHaber = round2(groupRows.reduce((sum, row) => sum + Number(row.HABER || 0), 0));
    const closingBalance = round2(openingBalance + totalDebe - totalHaber);
    let runningBalance = openingBalance;

    for (const row of groupRows) {
      runningBalance = round2(runningBalance + Number(row.DEBE || 0) - Number(row.HABER || 0));
      const cols = new Array(30).fill("");
      cols[6] = sanitizeText(row.COD);
      cols[7] = sanitizeText(row.CUENTA);
      cols[8] = formatReportAmount(openingBalance);
      cols[9] = formatReportAmount(totalDebe);
      cols[10] = formatReportAmount(totalHaber);
      cols[11] = formatReportAmount(closingBalance);
      cols[21] = sanitizeText(row.EXT) || "N";
      cols[22] = formatReportDate(row.FECHA);
      cols[23] = sanitizeText(row.ORIGEN);
      cols[24] = String(Number(row.ASIENTO || 0));
      cols[25] = sanitizeText(row.DOCU);
      cols[26] = sanitizeText(row.DETALLE).replace(/\t/g, " ");
      cols[27] = formatReportAmount(row.DEBE);
      cols[28] = formatReportAmount(row.HABER);
      cols[29] = formatReportAmount(runningBalance);
      lines.push(cols.join("\t"));
    }
  }

  const filePath = path.join(os.tmpdir(), `${label}_${Date.now()}.txt`);
  fs.writeFileSync(filePath, lines.join("\r\n"), "utf8");
  return filePath;
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

function buildMultipartBody(token, files) {
  const boundary = `----CodexBoundary${Date.now()}${Math.random().toString(16).slice(2)}`;
  const chunks = [];

  const push = (value) => {
    chunks.push(Buffer.isBuffer(value) ? value : Buffer.from(value, "utf8"));
  };

  push(`--${boundary}\r\n`);
  push('Content-Disposition: form-data; name="_token"\r\n\r\n');
  push(token);
  push("\r\n");

  for (const file of files) {
    const content = fs.readFileSync(file.path);
    push(`--${boundary}\r\n`);
    push(`Content-Disposition: form-data; name="${file.field}"; filename="${path.basename(file.path)}"\r\n`);
    push(`Content-Type: ${file.contentType}\r\n\r\n`);
    push(content);
    push("\r\n");
  }

  push(`--${boundary}--\r\n`);

  return {
    boundary,
    body: Buffer.concat(chunks),
  };
}

function extractCsrfToken(html, label) {
  const match = html.match(/name="_token"\s+value="([^"]+)"/i);
  assertCondition(match, `No se encontro CSRF token en ${label}.`);
  return match[1];
}

function extractCookie(headers, label) {
  const setCookie = headers["set-cookie"];
  assertCondition(Array.isArray(setCookie) && setCookie.length > 0, `Laravel no devolvio cookie de sesion en ${label}.`);
  return setCookie.map((cookie) => cookie.split(";", 1)[0]).join("; ");
}

function extractDownloadInfo(html, label) {
  const linkMatch = html.match(/href="([^"]*\/downloads\/[^"]+)"/i);
  assertCondition(linkMatch, `${label}: no se encontro enlace de descarga Laravel.`);

  const link = linkMatch[1];
  const absoluteUrl = link.startsWith("http") ? link : `${BASE_URL}${link.startsWith("/") ? "" : "/"}${link}`;
  const parsed = new URL(absoluteUrl);
  const codeMatch = html.match(/<code>([^<]+\.xlsx)<\/code>/i);
  const fileName = codeMatch ? codeMatch[1].trim() : path.basename(parsed.pathname);

  assertCondition(fileName.toLowerCase().endsWith(".xlsx"), `${label}: no se pudo derivar nombre XLSX de la respuesta.`);

  return {
    fileName,
    link: parsed.pathname + parsed.search,
  };
}

async function submitModuleFile(moduleSlug, field, filePath, contentType) {
  const page = await httpRequest("GET", `/cxp/modules/${moduleSlug}`);
  assertCondition(page.status === 200, `${moduleSlug}: GET devolvio HTTP ${page.status}.`);
  const cookie = extractCookie(page.headers, moduleSlug);
  const token = extractCsrfToken(page.text, moduleSlug);
  const payload = buildMultipartBody(token, [
    { field, path: filePath, contentType },
  ]);

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

async function submitBundleModule(moduleSlug) {
  const page = await httpRequest("GET", `/cxp/modules/${moduleSlug}`);
  assertCondition(page.status === 200, `${moduleSlug}: GET devolvio HTTP ${page.status}.`);
  const cookie = extractCookie(page.headers, moduleSlug);
  const token = extractCsrfToken(page.text, moduleSlug);
  const payload = buildMultipartBody(token, []);

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

async function loadWorkbook(source) {
  const workbook = new ExcelJS.Workbook();
  if (Buffer.isBuffer(source)) {
    await workbook.xlsx.load(source);
  } else {
    await workbook.xlsx.readFile(source);
  }
  return workbook;
}

function getFormula(cell) {
  const value = cell.value;
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

function normalizeFormulaText(formula) {
  if (typeof formula !== "string" || formula === "" || formula.startsWith("SHARED:")) {
    return null;
  }
  return formula.startsWith("=") ? formula : `=${formula}`;
}

function extractFormulaSheetReferences(formula) {
  const normalized = normalizeFormulaText(formula);
  if (!normalized) {
    return [];
  }

  const references = [];
  const quotedPattern = /'((?:[^']|'')+)'!/g;
  let match;
  while ((match = quotedPattern.exec(normalized)) !== null) {
    references.push(match[1].replace(/''/g, "'"));
  }

  const unquotedFormula = normalized.replace(quotedPattern, " ");
  const unquotedPattern = /(^|[^A-Z0-9_])([A-Za-z_][A-Za-z0-9_ .-]*)!/g;
  while ((match = unquotedPattern.exec(unquotedFormula)) !== null) {
    const candidate = sanitizeText(match[2]);
    if (candidate) {
      references.push(candidate);
    }
  }

  return references;
}

function formulaHasOrphanReference(formula, sheetNames) {
  const normalized = normalizeFormulaText(formula);
  if (!normalized) {
    return false;
  }

  const references = extractFormulaSheetReferences(normalized);
  return references.some((reference) => reference.includes("[") || !sheetNames.has(reference));
}

function normalizeFormulaAgainstWorkbook(formula, sheetNames) {
  if (!formula) {
    return null;
  }
  return formulaHasOrphanReference(formula, sheetNames) ? null : formula;
}

function assertNoOrphanWorkbookFormulas(workbook, label) {
  const sheetNames = new Set(workbook.worksheets.map((worksheet) => worksheet.name));
  const orphans = [];

  workbook.worksheets.forEach((worksheet) => {
    worksheet.eachRow((row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        const formula = getFormula(cell);
        if (!formula || !formulaHasOrphanReference(formula, sheetNames)) {
          return;
        }
        orphans.push(`${worksheet.name}!${cell.address}:${formula}`);
      });
    });
  });

  assertCondition(orphans.length === 0, `${label}: formulas huerfanas. ${orphans.slice(0, 5).join(" | ")}`);
}

function normalizeComparable(value) {
  if (Array.isArray(value)) {
    return value.map((item) => normalizeComparable(item));
  }
  if (value && typeof value === "object") {
    const normalized = Object.keys(value)
      .sort()
      .reduce((result, key) => {
        result[key] = normalizeComparable(value[key]);
        return result;
      }, {});
    if (normalized.type === "pattern" && (normalized.pattern == null || normalized.pattern === "")) {
      normalized.pattern = "none";
    }
    return normalized;
  }
  return value ?? null;
}

function styleSignature(cell) {
  return JSON.stringify({
    numFmt: normalizeComparable(cell.numFmt || null),
    font: normalizeComparable(cell.font || null),
    alignment: normalizeComparable(cell.alignment || null),
    border: normalizeComparable(cell.border || null),
    fill: normalizeComparable(cell.fill || null),
    protection: normalizeComparable(cell.protection || null),
  });
}

function assertStyleEquals(templateCell, outputCell, label) {
  assertCondition(
    styleSignature(templateCell) === styleSignature(outputCell),
    `${label}: estilo distinto en ${outputCell.address}.`,
  );
}

function assertRowStyles(templateWs, outputWs, templateRowIndex, outputRowIndex, startCol, endCol, label) {
  for (let col = startCol; col <= endCol; col += 1) {
    assertStyleEquals(templateWs.getRow(templateRowIndex).getCell(col), outputWs.getRow(outputRowIndex).getCell(col), label);
  }
}

function assertHeaderStyles(templateWs, outputWs, startCol, endCol, label) {
  assertRowStyles(templateWs, outputWs, 1, 1, startCol, endCol, label);
}

function assertFormulaMapMatches(templateWs, outputWs, label, expectedFormulaNormalizer = (formula) => formula) {
  const mismatches = [];
  templateWs.eachRow((row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      const formula = getFormula(cell);
      const expectedFormula = expectedFormulaNormalizer(formula);
      if (!expectedFormula) {
        const outputFormula = getFormula(outputWs.getCell(cell.address));
        if (outputFormula != null) {
          mismatches.push(`${cell.address}: null <> ${outputFormula}`);
        }
        return;
      }
      const outputFormula = getFormula(outputWs.getCell(cell.address));
      if (expectedFormula !== outputFormula) {
        mismatches.push(`${cell.address}: ${expectedFormula} <> ${outputFormula || "null"}`);
      }
    });
  });

  assertCondition(mismatches.length === 0, `${label}: formulas distintas. ${mismatches.slice(0, 5).join(" | ")}`);
}

function countNonEmptyRows(worksheet, columnIndex, startRow) {
  let count = 0;
  for (let rowIndex = startRow; rowIndex <= worksheet.rowCount; rowIndex += 1) {
    const value = worksheet.getRow(rowIndex).getCell(columnIndex).value;
    if (value == null || sanitizeText(value) === "") {
      continue;
    }
    count += 1;
  }
  return count;
}

function findLastNonEmptyRow(worksheet, columnIndex, startRow) {
  for (let rowIndex = worksheet.rowCount; rowIndex >= startRow; rowIndex -= 1) {
    const value = worksheet.getRow(rowIndex).getCell(columnIndex).value;
    if (value != null && sanitizeText(value) !== "") {
      return rowIndex;
    }
  }
  return startRow;
}

function hasMeaningfulStyle(worksheet, rowIndex, startCol, endCol) {
  for (let col = startCol; col <= endCol; col += 1) {
    const signature = styleSignature(worksheet.getRow(rowIndex).getCell(col));
    if (signature === styleSignature(new ExcelJS.Workbook().addWorksheet("tmp").getCell("A1"))) {
      return false;
    }
  }
  return true;
}

function findLastStyledRow(worksheet, startRow, startCol, endCol) {
  let last = startRow;
  for (let rowIndex = startRow; rowIndex <= worksheet.rowCount; rowIndex += 1) {
    if (hasMeaningfulStyle(worksheet, rowIndex, startCol, endCol)) {
      last = rowIndex;
    }
  }
  return last;
}

function worksheetValueSignature(worksheet, startCol, endCol) {
  const rows = [];
  for (let rowIndex = 1; rowIndex <= worksheet.rowCount; rowIndex += 1) {
    const values = [];
    for (let col = startCol; col <= endCol; col += 1) {
      const cellValue = worksheet.getRow(rowIndex).getCell(col).value;
      const normalizedValue = (cellValue && typeof cellValue === "object" && Object.prototype.hasOwnProperty.call(cellValue, "formula"))
        ? { formula: cellValue.formula || null }
        : normalizeComparable(cellValue);
      values.push(JSON.stringify(normalizedValue));
    }
    rows.push(values.join("|"));
  }
  return rows.join("\n");
}

function assertWorksheetValuesMatch(expectedWs, outputWs, startCol, endCol, label) {
  assertCondition(
    expectedWs.rowCount === outputWs.rowCount,
    `${label}: cantidad de filas distinta (${outputWs.rowCount}/${expectedWs.rowCount}).`,
  );

  const expectedSignature = worksheetValueSignature(expectedWs, startCol, endCol);
  const outputSignature = worksheetValueSignature(outputWs, startCol, endCol);
  assertCondition(expectedSignature === outputSignature, `${label}: los valores no coinciden con la plantilla mensual.`);
}

function findTemplateDataRowForOutput(outputRowIndex, templateLastStyledRow) {
  return Math.min(outputRowIndex, templateLastStyledRow);
}

function buildAction2Variants() {
  const sourcePath = actionFixturePath("CXPREP_RET_GENERALACCION2.txt");
  const hints = loadTemplateTipoHints(accion2Constants.DEFAULT_TEMPLATE_XLSX);
  const baseRows = normalizeParsedRows(extractRowsFromTxt(sourcePath), hints);
  const shortRows = baseRows.slice(0, Math.min(12, baseRows.length));
  const extraRows = baseRows.slice(0, Math.min(10, baseRows.length)).map((row, index) => mutateAction2Row(row, index));
  const longRows = baseRows.concat(extraRows);
  return [
    { label: "accion2-corto", rows: shortRows },
    { label: "accion2-largo", rows: longRows },
  ];
}

async function buildAction3Variants() {
  const sourcePath = actionFixturePath("CON_MAYORGEN2ACCION3.txt");
  const parsed = await parseAction3InputSources([sourcePath]);
  const baseRows = parsed.rows;
  const shortRows = baseRows.slice(0, Math.min(18, baseRows.length));
  const extraRows = baseRows.slice(0, Math.min(8, baseRows.length)).map((row, index) => mutateMovementRow(row, index));
  const longRows = baseRows.concat(extraRows);
  return [
    { label: "accion3-corto", rows: shortRows },
    { label: "accion3-largo", rows: longRows },
  ];
}

async function buildAction4Variants() {
  const sourcePath = actionFixturePath("CON_MAYORGEN2IVAACCION4.TXT");
  const parsed = await parseAction4InputSources([sourcePath]);
  const baseRows = parsed.rows;
  const firstInvenIndex = baseRows.findIndex((row) => sanitizeText(row.ORIGEN).toUpperCase() === "INVEN");
  const shortTarget = firstInvenIndex >= 0 ? Math.min(baseRows.length, firstInvenIndex + 12) : Math.min(80, baseRows.length);
  const shortRows = baseRows.slice(0, Math.max(20, shortTarget));
  const extraRows = baseRows.slice(0, Math.min(12, baseRows.length)).map((row, index) => mutateMovementRow(row, index));
  const longRows = baseRows.concat(extraRows);
  return [
    { label: "accion4-corto", rows: shortRows },
    { label: "accion4-largo", rows: longRows },
  ];
}

async function runAction2Audit(templateBundle) {
  for (const variant of buildAction2Variants()) {
    const filePath = buildAction2InputFile(variant.rows, variant.label);
    try {
      const outputBuffer = await submitModuleFile("accion2", "source_files", filePath, "text/plain");
      const outputWorkbook = await loadWorkbook(outputBuffer);
      const outputWs = outputWorkbook.getWorksheet("RET PROV");
      const outputSheetNames = new Set(outputWorkbook.worksheets.map((worksheet) => worksheet.name));
      assertCondition(!!outputWs, `${variant.label}: no existe hoja RET PROV.`);
      assertCondition(
        countNonEmptyRows(outputWs, 1, 2) === variant.rows.length,
        `${variant.label}: cantidad de filas de salida distinta a la esperada.`,
      );
      assertNoOrphanWorkbookFormulas(outputWorkbook, variant.label);
      assertHeaderStyles(templateBundle.ws, outputWs, 1, 10, variant.label);
      assertFormulaMapMatches(
        templateBundle.ws,
        outputWs,
        variant.label,
        (formula) => normalizeFormulaAgainstWorkbook(formula, outputSheetNames),
      );
      assertRowStyles(templateBundle.ws, outputWs, 2, 2, 1, 10, `${variant.label}-primera-fila`);
      const lastOutputRow = variant.rows.length + 1;
      const templateRow = findTemplateDataRowForOutput(lastOutputRow, templateBundle.lastStyledDataRow);
      assertRowStyles(templateBundle.ws, outputWs, templateRow, lastOutputRow, 1, 10, `${variant.label}-ultima-fila`);
    } finally {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    }
  }
}

async function runAction1MonthlyAudit() {
  const monthlyPdfPath = actionMonthlyPath([/^CXPREP_docproveedor.*\.pdf$/i], "accion1-mensual-pdf");
  const monthlyTemplatePath = actionMonthlyPath([/PLANTILLAHECHAAMANO.*\.xlsx$/i], "accion1-mensual-plantilla");

  const outputBuffer = await submitModuleFile("accion1", "source_files", monthlyPdfPath, "application/pdf");
  const monthlyTemplateBuffer = fs.readFileSync(monthlyTemplatePath);
  const outputWorkbook = await loadWorkbook(outputBuffer);
  const expectedWorkbook = await loadWorkbook(monthlyTemplatePath);
  const outputWs = outputWorkbook.getWorksheet("LIBRO COMPRAS");
  const expectedWs = expectedWorkbook.getWorksheet("LIBRO COMPRAS");

  assertCondition(!!outputWs, "accion1-mensual: no existe hoja LIBRO COMPRAS en la salida.");
  assertCondition(!!expectedWs, "accion1-mensual: no existe hoja LIBRO COMPRAS en la plantilla manual.");
  assertCondition(!outputBuffer.equals(monthlyTemplateBuffer), "accion1-mensual: la descarga no debe ser un clon binario de la plantilla manual.");
  assertWorksheetValuesMatch(expectedWs, outputWs, 1, 12, "accion1-mensual");
}

async function runAction3Audit(templateBundle) {
  for (const variant of await buildAction3Variants()) {
    const filePath = buildAction3TxtFile(variant.rows, variant.label);
    try {
      const outputBuffer = await submitModuleFile("accion3", "source_files[]", filePath, "text/plain");
      const outputWorkbook = await loadWorkbook(outputBuffer);
      const outputWs = outputWorkbook.getWorksheet("MAYOR RET");
      assertCondition(!!outputWs, `${variant.label}: no existe hoja MAYOR RET.`);
      const expectedRenderedRows = variant.rows.filter((row) => sanitizeText(row.ORIGEN).toUpperCase() !== "AGCM").length;
      assertCondition(
        countNonEmptyRows(outputWs, 1, 2) === expectedRenderedRows,
        `${variant.label}: cantidad de movimientos distinta a la esperada.`,
      );
      assertHeaderStyles(templateBundle.ws, outputWs, 1, 11, variant.label);
      assertFormulaMapMatches(templateBundle.ws, outputWs, variant.label);
      assertRowStyles(templateBundle.ws, outputWs, 2, 2, 1, 11, `${variant.label}-primera-fila`);
      const lastOutputRow = expectedRenderedRows + 1;
      const templateRow = Math.min(lastOutputRow, templateBundle.ws.rowCount);
      assertRowStyles(templateBundle.ws, outputWs, templateRow, lastOutputRow, 1, 11, `${variant.label}-ultima-fila`);
    } finally {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    }
  }
}

async function runAction4Audit(templateBundle) {
  for (const variant of await buildAction4Variants()) {
    const filePath = buildSingleAccountMayorTxtFile(variant.rows, variant.label);
    try {
      const outputBuffer = await submitModuleFile("accion4", "source_files", filePath, "text/plain");
      const outputWorkbook = await loadWorkbook(outputBuffer);
      const outputWs = outputWorkbook.getWorksheet("MAYOR IVA");
      assertCondition(!!outputWs, `${variant.label}: no existe hoja MAYOR IVA.`);

      const expectedPlan = buildAction4OutputPlan(variant.rows);
      const movementCount = countNonEmptyRows(outputWs, 1, 2);
      assertCondition(
        movementCount === expectedPlan.movementRows.length,
        `${variant.label}: cantidad de movimientos distinta a la esperada.`,
      );

      assertHeaderStyles(templateBundle.ws, outputWs, 1, 11, variant.label);
      assertRowStyles(templateBundle.ws, outputWs, 2, 2, 1, 11, `${variant.label}-primera-fila`);

      const lastMovementRow = findLastNonEmptyRow(outputWs, 1, 2);
      const templateRow = Math.min(lastMovementRow, templateBundle.ws.rowCount);
      assertRowStyles(templateBundle.ws, outputWs, templateRow, lastMovementRow, 1, 11, `${variant.label}-ultima-fila`);

      for (let index = 0; index < expectedPlan.rowPlan.length; index += 1) {
        const item = expectedPlan.rowPlan[index];
        if (item.type !== "subtotal") {
          continue;
        }
        const subtotalRow = index + 2;
        const balanceRow = subtotalRow + 1;
        assertCondition(
          getFormula(outputWs.getCell(`I${subtotalRow}`)) === `SUM(I${item.fromRow}:I${item.toRow})`,
          `${variant.label}: formula subtotal DEBE incorrecta en fila ${subtotalRow}.`,
        );
        assertCondition(
          getFormula(outputWs.getCell(`J${subtotalRow}`)) === `SUM(J${item.fromRow}:J${item.toRow})`,
          `${variant.label}: formula subtotal HABER incorrecta en fila ${subtotalRow}.`,
        );
        const expectedBalanceFormula = expectedPlan.rowPlan[index + 1]?.mode === "haber_minus_debe"
          ? `+J${balanceRow - 1}-I${balanceRow - 1}`
          : `+I${balanceRow - 1}-J${balanceRow - 1}`;
        assertCondition(
          getFormula(outputWs.getCell(`K${balanceRow}`)) === expectedBalanceFormula,
          `${variant.label}: formula subtotal SALDO incorrecta en fila ${balanceRow}.`,
        );
      }
    } finally {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    }
  }
}

async function runBundleAudit() {
  const outputBuffer = await submitBundleModule("consolidado-acciones");
  const outputWorkbook = await loadWorkbook(outputBuffer);
  const outputWs = outputWorkbook.getWorksheet("ACCION 2 RET PROV");
  assertCondition(!!outputWs, "bundle: no existe hoja ACCION 2 RET PROV.");
  assertCondition(outputWorkbook.worksheets.length === 4, "bundle: el consolidado debe tener 4 hojas.");

  const externalFormulas = [];
  outputWs.eachRow((row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      const formula = getFormula(cell);
      if (formula && formula.includes("[")) {
        externalFormulas.push(`${cell.address}:${formula}`);
      }
    });
  });
  assertCondition(
    externalFormulas.length === 0,
    `bundle: quedaron formulas externas en ACCION 2. ${externalFormulas.slice(0, 5).join(" | ")}`,
  );

  const expectedCells = ["O3", "O4", "O5", "O10", "O11", "O12", "O13", "O14", "O15"];
  for (const cellAddress of expectedCells) {
    const formula = getFormula(outputWs.getCell(cellAddress));
    if (formula == null) {
      continue;
    }
    assertCondition(
      typeof formula === "string" && formula.includes("'ACCION 3 MAYOR RET'!"),
      `bundle: ${cellAddress} no apunta a la hoja interna ACCION 3 MAYOR RET.`,
    );
  }
}

async function main() {
  const action2Template = await loadWorkbook(actionTemplatePath("ACCION2.xlsx"));
  const action3Template = await loadWorkbook(actionTemplatePath("MAYOR RET_ACCION3.xlsx"));
  const action4Template = await loadWorkbook(actionTemplatePath("MAYORIVAACCION4.xlsx"));

  const templateBundles = {
    action2: {
      ws: action2Template.getWorksheet("RET PROV"),
      lastStyledDataRow: findLastStyledRow(action2Template.getWorksheet("RET PROV"), 2, 1, 10),
    },
    action3: {
      ws: action3Template.getWorksheet("MAYOR RET"),
    },
    action4: {
      ws: action4Template.getWorksheet("MAYOR IVA"),
    },
  };

  const server = startPhpServer();
  try {
    await waitServerReady();
    await runAction1MonthlyAudit();
    await runAction2Audit(templateBundles.action2);
    await runAction3Audit(templateBundles.action3);
    await runAction4Audit(templateBundles.action4);
    await runBundleAudit();
    console.log("OK: auditoria mensual de plantillas y resiliencia de ventana 1 validada.");
  } finally {
    server.kill();
  }
}

main().catch((error) => {
  console.error(`FAIL: ${error.message}`);
  process.exit(1);
});
