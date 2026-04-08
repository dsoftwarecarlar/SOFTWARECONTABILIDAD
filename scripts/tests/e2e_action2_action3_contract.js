const fs = require("fs");
const http = require("http");
const os = require("os");
const path = require("path");
const crypto = require("crypto");
const { spawn, spawnSync } = require("child_process");

const AdmZip = require("adm-zip");
const ExcelJS = require("exceljs");
const XLSX = require("xlsx");

const {
  extractRowsFromTxt,
  loadTemplateTipoHints,
  normalizeParsedRows,
} = require("../cxp/accion2/parser");
const accion2Constants = require("../cxp/accion2/constants");
const { parseInputSources } = require("../cxp/accion3/parser");

const ROOT = path.resolve(__dirname, "..", "..");
const APP_ROOT = path.join(ROOT, "laravel_app");
const ROUTER = path.join(APP_ROOT, "server.php");
const PHP_EXE = firstExistingPath([
  path.join(ROOT, ".tools", "php82", "php.exe"),
  process.env.PHP_EXE || "",
  "C:\\xampp\\php\\php.exe",
]);
const HOST = "127.0.0.1";
const PORT = 18987;
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

function normalizeText(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
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
  const codeMatch = html.match(/<code>([^<]+\.xlsx)<\/code>/i);
  assertCondition(codeMatch, `${label}: no se encontro nombre de archivo generado en la respuesta.`);
  const linkMatch = html.match(/href="([^"]*\/downloads\/[^"]+)"/i);
  assertCondition(linkMatch, `${label}: no se encontro enlace de descarga Laravel.`);

  const link = linkMatch[1];
  const absoluteUrl = link.startsWith("http") ? link : `${BASE_URL}${link.startsWith("/") ? "" : "/"}${link}`;
  const parsed = new URL(absoluteUrl);

  return {
    fileName: codeMatch[1].trim(),
    link: parsed.pathname + parsed.search,
  };
}

function listColumnValues(sheet, columnIndex, startRow = 2) {
  const values = [];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  for (let row = startRow; row <= range.e.r + 1; row += 1) {
    const cellRef = XLSX.utils.encode_cell({ r: row - 1, c: columnIndex - 1 });
    const cell = sheet[cellRef];
    if (!cell || cell.v == null || String(cell.v).trim() === "") {
      continue;
    }
    values.push(cell.v);
  }
  return values;
}

function normalizeWorkbookDate(value) {
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed && parsed.y && parsed.m && parsed.d) {
      return `${String(parsed.y).padStart(4, "0")}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
  }
  return String(value || "").trim();
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
  assertCondition(!/relationships\/externalLink/i.test(workbookRels), `${label}: workbook.xml.rels conserva relaciones externalLink.`);

  const contentTypes = readZipEntryText(zip, "[Content_Types].xml");
  assertCondition(
    !/spreadsheetml\.externalLink\+xml/i.test(contentTypes),
    `${label}: [Content_Types].xml conserva overrides de externalLink.`,
  );
}

async function assertExcelJsReadable(buffer, label) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  assertCondition(workbook.worksheets.length > 0, `${label}: ExcelJS no detecto hojas dentro del XLSX.`);
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
    const candidate = normalizeText(match[2]);
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

async function assertNoOrphanWorkbookFormulas(buffer, label) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const sheetNames = new Set(workbook.worksheets.map((worksheet) => worksheet.name));
  const orphans = [];

  workbook.worksheets.forEach((worksheet) => {
    worksheet.eachRow((row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        const formula = getFormula(cell.value);
        if (!formula || !formulaHasOrphanReference(formula, sheetNames)) {
          return;
        }
        orphans.push(`${worksheet.name}!${cell.address}:${formula}`);
      });
    });
  });

  assertCondition(orphans.length === 0, `${label}: formulas huerfanas. ${orphans.slice(0, 5).join(" | ")}`);
}

function buildAction3WorkbookSignature(source) {
  const workbook = Buffer.isBuffer(source)
    ? XLSX.read(source, { type: "buffer", cellDates: true, cellFormula: false })
    : XLSX.readFile(source, { cellDates: true, cellFormula: false });
  const sheet = workbook.Sheets["MAYOR RET"];
  assertCondition(!!sheet, "Accion 3 PDF: no existe hoja MAYOR RET.");

  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:K1");
  const rows = [];
  for (let rowIndex = 1; rowIndex <= range.e.r; rowIndex += 1) {
    const values = [];
    let nonEmpty = false;
    for (let columnIndex = 0; columnIndex < 11; columnIndex += 1) {
      const ref = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      const cell = sheet[ref];
      let value = "";
      if (cell && cell.v != null) {
        if (columnIndex === 3) {
          value = normalizeWorkbookDate(cell.v);
        } else if (typeof cell.v === "number") {
          value = String(Math.round((cell.v + Number.EPSILON) * 100) / 100);
        } else {
          value = String(cell.v).trim();
        }
      }
      if (value !== "") {
        nonEmpty = true;
      }
      values.push(value);
    }
    if (nonEmpty) {
      rows.push(values.join("|"));
    }
  }

  return {
    rows: rows.length,
    hash: crypto.createHash("sha256").update(rows.join("\n")).digest("hex"),
  };
}

function buildLegacyAction3PdfWorkbookSignature(pdfInputPath) {
  const outputPath = path.join(os.tmpdir(), `accion3_pdf_legacy_${Date.now()}.xlsx`);
  const scriptPath = path.join(ROOT, "run_bot_accion3.js");
  const templatePath = actionTemplatePath("MAYOR RET_ACCION3.xlsx");

  const completed = spawnSync("node", [scriptPath, pdfInputPath, outputPath, templatePath], {
    cwd: ROOT,
    encoding: "utf8",
    windowsHide: true,
    timeout: 240000,
  });

  if (completed.error) {
    throw completed.error;
  }

  if (completed.status !== 0) {
    throw new Error(
      `Accion 3 PDF legacy fallo: ${[completed.stdout, completed.stderr].filter(Boolean).join("\n").trim() || completed.status}`,
    );
  }

  try {
    assertCondition(fs.existsSync(outputPath), "Accion 3 PDF legacy no genero workbook.");
    return buildAction3WorkbookSignature(outputPath);
  } finally {
    if (fs.existsSync(outputPath)) {
      fs.unlinkSync(outputPath);
    }
    const auditPath = outputPath.replace(/\.xlsx$/i, "_auditoria.json");
    if (fs.existsSync(auditPath)) {
      fs.unlinkSync(auditPath);
    }
  }
}

function createAction2CanaryFile() {
  const sourcePath = actionFixturePath("CXPREP_RET_GENERALACCION2.txt");
  const sourceText = fs.readFileSync(sourcePath, "utf8");

  const rows = normalizeParsedRows(extractRowsFromTxt(sourcePath), loadTemplateTipoHints(accion2Constants.DEFAULT_TEMPLATE_XLSX));
  assertCondition(rows.length > 0, "No se pudieron extraer filas base para Accion 2.");
  const originalProvider = String(rows[0].proveedor || "").trim();
  assertCondition(originalProvider !== "", "No se detecto proveedor base en Accion 2.");

  const canary = `PROVEEDOR_CANARIO_${Date.now()}`;
  const position = sourceText.indexOf(originalProvider);
  assertCondition(position >= 0, "No se encontro proveedor base para inyectar canario en Accion 2.");

  const mutated = sourceText.slice(0, position) + canary + sourceText.slice(position + originalProvider.length);
  const filePath = path.join(os.tmpdir(), `accion2_canary_${Date.now()}.txt`);
  fs.writeFileSync(filePath, mutated, "utf8");

  const parsedCanaryRows = normalizeParsedRows(
    extractRowsFromTxt(filePath),
    loadTemplateTipoHints(accion2Constants.DEFAULT_TEMPLATE_XLSX),
  );
  const providerValues = parsedCanaryRows.map((item) => normalizeText(item.proveedor));
  assertCondition(
    providerValues.includes(normalizeText(canary)),
    "No se pudo confirmar proveedor canario en parsing de Accion 2.",
  );

  return { filePath, canary };
}

async function createAction3CanaryFile() {
  const sourcePath = actionFixturePath("CON_MAYORGEN2ACCION3.txt");
  const sourceText = fs.readFileSync(sourcePath, "utf8");
  const parsed = await parseInputSources([sourcePath]);
  const baseRow = parsed.rows.find((item) => String(item.DOCU || "").trim() !== "");
  assertCondition(baseRow, "No se encontro DOCU base en Accion 3.");

  const originalDocu = String(baseRow.DOCU).trim();
  const canary = `DOCCANARIO${String(Date.now()).slice(-8)}`;
  const position = sourceText.indexOf(originalDocu);
  assertCondition(position >= 0, "No se encontro DOCU base para inyectar canario en Accion 3.");

  const mutated = sourceText.slice(0, position) + canary + sourceText.slice(position + originalDocu.length);
  const filePath = path.join(os.tmpdir(), `accion3_canary_${Date.now()}.txt`);
  fs.writeFileSync(filePath, mutated, "utf8");

  const parsedCanary = await parseInputSources([filePath]);
  const docValues = parsedCanary.rows.map((item) => normalizeText(item.DOCU));
  assertCondition(docValues.includes(normalizeText(canary)), "No se pudo confirmar DOCU canario en parsing de Accion 3.");

  return { filePath, canary };
}

async function runAction2Contract() {
  const canaryInput = createAction2CanaryFile();
  try {
    const page = await httpRequest("GET", "/cxp/modules/accion2");
    assertCondition(page.status === 200, `Accion 2 Laravel devolvio HTTP ${page.status}.`);
    const cookie = extractCookie(page.headers, "accion2");
    const token = extractCsrfToken(page.text, "accion2");
    const payload = buildMultipartBody(token, [
      { field: "source_files", path: canaryInput.filePath, contentType: "text/plain" },
    ]);

    const response = await httpRequest("POST", "/cxp/modules/accion2", payload.body, {
      "Content-Type": `multipart/form-data; boundary=${payload.boundary}`,
      "Content-Length": String(payload.body.length),
      Cookie: cookie,
    });

    assertCondition(response.status === 200, `Accion 2 Laravel devolvio estado HTTP ${response.status}.`);
    const parsed = extractDownloadInfo(response.text, "Accion 2");
    const download = await httpRequestWithRedirects("GET", parsed.link, null, { Cookie: cookie });
    assertCondition(download.status === 200, `No se pudo descargar salida de Accion 2 (${download.status}).`);
    assertNoExternalLinks(download.buffer, "Accion 2");
    await assertExcelJsReadable(download.buffer, "Accion 2");
    await assertNoOrphanWorkbookFormulas(download.buffer, "Accion 2");

    const workbook = XLSX.read(download.buffer, { type: "buffer" });
    const sheet = workbook.Sheets.RET_PROV || workbook.Sheets["RET PROV"];
    assertCondition(!!sheet, "Accion 2: no existe hoja RET PROV.");

    const providers = listColumnValues(sheet, 2, 2).map((item) => normalizeText(item));
    assertCondition(
      providers.includes(normalizeText(canaryInput.canary)),
      "Accion 2: el XLSX descargado no contiene el proveedor canario del TXT subido.",
    );

    const historyPage = await httpRequest("GET", "/cxp/modules/accion2", null, { Cookie: cookie });
    assertCondition(historyPage.status === 200, `Accion 2: historial Laravel devolvio ${historyPage.status}.`);
    assertCondition(
      historyPage.text.includes(parsed.fileName),
      "Accion 2: el historial Laravel no incluye el archivo recien generado.",
    );
  } finally {
    if (fs.existsSync(canaryInput.filePath)) {
      fs.unlinkSync(canaryInput.filePath);
    }
  }
}

async function runAction3Contract() {
  const canaryInput = await createAction3CanaryFile();
  try {
    const page = await httpRequest("GET", "/cxp/modules/accion3");
    assertCondition(page.status === 200, `Accion 3 Laravel devolvio HTTP ${page.status}.`);
    const cookie = extractCookie(page.headers, "accion3");
    const token = extractCsrfToken(page.text, "accion3");
    const payload = buildMultipartBody(token, [
      { field: "source_files[]", path: canaryInput.filePath, contentType: "text/plain" },
    ]);

    const response = await httpRequest("POST", "/cxp/modules/accion3", payload.body, {
      "Content-Type": `multipart/form-data; boundary=${payload.boundary}`,
      "Content-Length": String(payload.body.length),
      Cookie: cookie,
    });

    assertCondition(response.status === 200, `Accion 3 Laravel devolvio estado HTTP ${response.status}.`);
    const parsed = extractDownloadInfo(response.text, "Accion 3");
    const download = await httpRequestWithRedirects("GET", parsed.link, null, { Cookie: cookie });
    assertCondition(download.status === 200, `No se pudo descargar salida de Accion 3 (${download.status}).`);
    assertNoExternalLinks(download.buffer, "Accion 3");
    await assertExcelJsReadable(download.buffer, "Accion 3");

    const workbook = XLSX.read(download.buffer, { type: "buffer" });
    const sheet = workbook.Sheets["MAYOR RET"];
    assertCondition(!!sheet, "Accion 3: no existe hoja MAYOR RET.");

    const docValues = listColumnValues(sheet, 7, 2).map((item) => normalizeText(item));
    assertCondition(
      docValues.includes(normalizeText(canaryInput.canary)),
      "Accion 3: el XLSX descargado no contiene el DOCU canario del TXT subido.",
    );

    const historyPage = await httpRequest("GET", "/cxp/modules/accion3", null, { Cookie: cookie });
    assertCondition(historyPage.status === 200, `Accion 3: historial Laravel devolvio ${historyPage.status}.`);
    assertCondition(
      historyPage.text.includes(parsed.fileName),
      "Accion 3: el historial Laravel no incluye el archivo recien generado.",
    );
  } finally {
    if (fs.existsSync(canaryInput.filePath)) {
      fs.unlinkSync(canaryInput.filePath);
    }
  }
}

async function runAction3PdfContract() {
  const inputPath = actionFixturePath("CON_MAYORGEN2ACCION3.pdf");
  const page = await httpRequest("GET", "/cxp/modules/accion3");
  assertCondition(page.status === 200, `Accion 3 PDF Laravel devolvio HTTP ${page.status}.`);
  const cookie = extractCookie(page.headers, "accion3-pdf");
  const token = extractCsrfToken(page.text, "accion3-pdf");
  const payload = buildMultipartBody(token, [
    { field: "source_files[]", path: inputPath, contentType: "application/pdf" },
  ]);

  const response = await httpRequest("POST", "/cxp/modules/accion3", payload.body, {
    "Content-Type": `multipart/form-data; boundary=${payload.boundary}`,
    "Content-Length": String(payload.body.length),
    Cookie: cookie,
  });

  assertCondition(response.status === 200, `Accion 3 PDF Laravel devolvio estado HTTP ${response.status}.`);
  const parsed = extractDownloadInfo(response.text, "Accion 3 PDF");
  const download = await httpRequestWithRedirects("GET", parsed.link, null, { Cookie: cookie });
  assertCondition(download.status === 200, `No se pudo descargar salida PDF de Accion 3 (${download.status}).`);
  assertNoExternalLinks(download.buffer, "Accion 3 PDF");
  await assertExcelJsReadable(download.buffer, "Accion 3 PDF");

  const generatedSignature = buildAction3WorkbookSignature(download.buffer);
  const legacySignature = buildLegacyAction3PdfWorkbookSignature(inputPath);
  assertCondition(
    generatedSignature.rows === legacySignature.rows && generatedSignature.hash === legacySignature.hash,
    `Accion 3 PDF: la salida Laravel+Python no coincide con el workbook legacy (rows=${generatedSignature.rows}/${legacySignature.rows}).`,
  );

  const historyPage = await httpRequest("GET", "/cxp/modules/accion3", null, { Cookie: cookie });
  assertCondition(historyPage.status === 200, `Accion 3 PDF: historial Laravel devolvio ${historyPage.status}.`);
  assertCondition(
    historyPage.text.includes(parsed.fileName),
    "Accion 3 PDF: el historial Laravel no incluye el archivo recien generado.",
  );
}

async function main() {
  const server = startPhpServer();
  try {
    await waitServerReady();
    await runAction2Contract();
    await runAction3Contract();
    await runAction3PdfContract();
    console.log("OK: contrato E2E Accion 2/3 validado en Laravel.");
  } finally {
    server.kill();
  }
}

main().catch((error) => {
  console.error(`FAIL: ${error.message}`);
  process.exit(1);
});
