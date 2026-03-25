const fs = require("fs");
const http = require("http");
const os = require("os");
const path = require("path");
const crypto = require("crypto");
const { spawn } = require("child_process");

const AdmZip = require("adm-zip");
const ExcelJS = require("exceljs");
const XLSX = require("xlsx");

const { parseInputSources } = require("../cxp/accion4/parser");

const ROOT = path.resolve(__dirname, "..", "..");
const APP_ROOT = path.join(ROOT, "laravel_app");
const ROUTER = path.join(APP_ROOT, "server.php");
const PHP_EXE = firstExistingPath([
  path.join(ROOT, ".tools", "php82", "php.exe"),
  process.env.PHP_EXE || "",
  "C:\\xampp\\php\\php.exe",
]);
const HOST = "127.0.0.1";
const PORT = 18991;
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

function action1FixturePath(fileName) {
  return firstExistingPath([
    path.join(ROOT, "resources", "cxp", "acciones", "fixtures", fileName),
    path.join(ROOT, fileName),
  ]);
}

function action1ContractPath(fileName) {
  return firstExistingPath([
    path.join(ROOT, "resources", "cxp", "acciones", "contracts", fileName),
    path.join(ROOT, "outputs", fileName),
  ]);
}

function normalizeText(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function normalizeDocuText(value) {
  const normalized = normalizeText(value);
  if (/^\d+$/.test(normalized)) {
    return String(Number(normalized));
  }
  return normalized;
}

function round2(value) {
  return Math.round((Number(value || 0) + Number.EPSILON) * 100) / 100;
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

function readZipEntryText(zip, entryName) {
  const entry = zip.getEntry(entryName);
  assertCondition(!!entry, `Consolidado: no existe ${entryName} dentro del XLSX.`);
  return entry.getData().toString("utf8");
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
    `${label}: faltan bookViews en xl/workbook.xml aunque hay sheetViews con workbookViewId.`,
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
  assertCondition(workbook.worksheets.length > 0, `${label}: ExcelJS no detecto hojas dentro del XLSX.`);
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

function buildAction1NormalizedSignature(source) {
  const workbook = Buffer.isBuffer(source)
    ? XLSX.read(source, { type: "buffer", cellFormula: false })
    : XLSX.readFile(source, { cellFormula: false });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  assertCondition(!!sheet, "Accion 1: el archivo no contiene hoja principal.");

  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:M1");
  const rows = [];
  const stableColumns = [0, 1, 3, 4, 5, 6, 7, 8, 9, 10];
  for (let row = 1; row <= range.e.r; row += 1) {
    const values = [];
    let nonEmpty = false;
    for (const column of stableColumns) {
      const ref = XLSX.utils.encode_cell({ r: row, c: column });
      const cell = sheet[ref];
      let value = "";
      if (cell && cell.v != null) {
        value = typeof cell.v === "number" ? String(round2(cell.v)) : String(cell.v).trim();
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

async function createAction4CanaryFile() {
  const sourcePath = actionFixturePath("CON_MAYORGEN2IVAACCION4.TXT");
  const sourceText = fs.readFileSync(sourcePath, "latin1");
  const parsed = await parseInputSources([sourcePath]);
  const firstRow = parsed.rows.find((item) => String(item.DOCU || "").trim() !== "");
  const originalDocu = String(firstRow ? firstRow.DOCU : "").trim();
  assertCondition(originalDocu !== "", "No se encontro DOCU base en Accion 4.");

  let canary = `IVACAN${String(Date.now()).slice(-8)}`;
  if (/^\d+$/.test(originalDocu)) {
    canary = String(Date.now()).slice(-originalDocu.length).padStart(originalDocu.length, "8");
    if (canary === originalDocu) {
      canary = `${"9".repeat(Math.max(originalDocu.length - 1, 0))}8`;
    }
  }
  const position = sourceText.indexOf(originalDocu);
  assertCondition(position >= 0, "No se encontro DOCU base para inyectar canario en Accion 4.");

  const mutated = sourceText.slice(0, position) + canary + sourceText.slice(position + originalDocu.length);
  const filePath = path.join(os.tmpdir(), `accion4_canary_${Date.now()}.txt`);
  fs.writeFileSync(filePath, mutated, "latin1");

  return { filePath, canary };
}

function columnValues(sheet, columnIndex, startRow = 2) {
  const values = [];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  for (let row = startRow; row <= range.e.r + 1; row += 1) {
    const ref = XLSX.utils.encode_cell({ r: row - 1, c: columnIndex - 1 });
    const cell = sheet[ref];
    if (!cell || cell.v == null || String(cell.v).trim() === "") {
      continue;
    }
    values.push(cell.v);
  }
  return values;
}

async function runAction1Contract() {
  const inputPath = action1FixturePath("CXPREP_docproveedor.pdf");
  const page = await httpRequest("GET", "/cxp/modules/accion1");
  assertCondition(page.status === 200, `Accion 1 Laravel devolvio HTTP ${page.status}.`);
  const cookie = extractCookie(page.headers, "accion1");
  const token = extractCsrfToken(page.text, "accion1");
  const payload = buildMultipartBody(token, [
    { field: "source_files", path: inputPath, contentType: "application/pdf" },
  ]);

  const response = await httpRequest("POST", "/cxp/modules/accion1", payload.body, {
    "Content-Type": `multipart/form-data; boundary=${payload.boundary}`,
    "Content-Length": String(payload.body.length),
    Cookie: cookie,
  });

  assertCondition(response.status === 200, `Accion 1 Laravel devolvio estado HTTP ${response.status}.`);
  const parsed = extractDownloadInfo(response.text, "Accion 1");
  const download = await httpRequestWithRedirects("GET", parsed.link, null, { Cookie: cookie });
  assertCondition(download.status === 200, `No se pudo descargar salida de Accion 1 (${download.status}).`);

  const outputWorkbook = XLSX.read(download.buffer, { type: "buffer", cellFormula: false });
  assertCondition(
    outputWorkbook.SheetNames.includes("LIBRO COMPRAS"),
    "Accion 1: no existe hoja LIBRO COMPRAS en el Excel descargado.",
  );
  assertNoExternalLinks(download.buffer, "Accion 1");
  assertWorkbookViewStructure(download.buffer, "Accion 1");
  assertStylesXmlReadable(download.buffer, "Accion 1");
  await assertExcelJsReadable(download.buffer, "Accion 1");

  const generatedSignature = buildAction1NormalizedSignature(download.buffer);
  const referenceSignature = buildAction1NormalizedSignature(
    action1ContractPath("CXPREP_docproveedor_20260306_092810_resultado.xlsx"),
  );
  assertCondition(
    generatedSignature.rows === referenceSignature.rows && generatedSignature.hash === referenceSignature.hash,
    `Accion 1: la salida no coincide con el contrato base (rows=${generatedSignature.rows}/${referenceSignature.rows}).`,
  );

  const historyPage = await httpRequest("GET", "/cxp/modules/accion1", null, { Cookie: cookie });
  assertCondition(historyPage.status === 200, `Accion 1: historial Laravel devolvio ${historyPage.status}.`);
  assertCondition(
    historyPage.text.includes(parsed.fileName),
    "Accion 1: el historial Laravel no incluye el archivo recien generado.",
  );
}

async function runAction4Contract() {
  const canaryInput = await createAction4CanaryFile();
  try {
    const page = await httpRequest("GET", "/cxp/modules/accion4");
    assertCondition(page.status === 200, `Accion 4 Laravel devolvio HTTP ${page.status}.`);
    const cookie = extractCookie(page.headers, "accion4");
    const token = extractCsrfToken(page.text, "accion4");
    const payload = buildMultipartBody(token, [
      { field: "source_files", path: canaryInput.filePath, contentType: "text/plain" },
    ]);

    const response = await httpRequest("POST", "/cxp/modules/accion4", payload.body, {
      "Content-Type": `multipart/form-data; boundary=${payload.boundary}`,
      "Content-Length": String(payload.body.length),
      Cookie: cookie,
    });

    assertCondition(response.status === 200, `Accion 4 Laravel devolvio estado HTTP ${response.status}.`);
    const parsed = extractDownloadInfo(response.text, "Accion 4");
    const download = await httpRequestWithRedirects("GET", parsed.link, null, { Cookie: cookie });
    assertCondition(download.status === 200, `No se pudo descargar salida de Accion 4 (${download.status}).`);

    const workbook = XLSX.read(download.buffer, { type: "buffer" });
    const sheet = workbook.Sheets["MAYOR IVA"];
    assertCondition(!!sheet, "Accion 4: no existe hoja MAYOR IVA.");
    assertNoExternalLinks(download.buffer, "Accion 4");
    await assertExcelJsReadable(download.buffer, "Accion 4");

    const docValues = columnValues(sheet, 7, 2).map((item) => normalizeText(item));
    assertCondition(
      docValues.map((item) => normalizeDocuText(item)).includes(normalizeDocuText(canaryInput.canary)),
      "Accion 4: el XLSX descargado no contiene el DOCU canario del TXT subido.",
    );

    const historyPage = await httpRequest("GET", "/cxp/modules/accion4", null, { Cookie: cookie });
    assertCondition(historyPage.status === 200, `Accion 4: historial Laravel devolvio ${historyPage.status}.`);
    assertCondition(
      historyPage.text.includes(parsed.fileName),
      "Accion 4: el historial Laravel no incluye el archivo recien generado.",
    );
  } finally {
    if (fs.existsSync(canaryInput.filePath)) {
      fs.unlinkSync(canaryInput.filePath);
    }
  }
}

async function runBundleCheck() {
  const page = await httpRequest("GET", "/cxp/modules/consolidado-acciones");
  assertCondition(page.status === 200, `Consolidado Laravel devolvio HTTP ${page.status}.`);
  const cookie = extractCookie(page.headers, "consolidado");
  const token = extractCsrfToken(page.text, "consolidado");
  const payload = buildMultipartBody(token, []);
  const response = await httpRequest("POST", "/cxp/modules/consolidado-acciones", payload.body, {
    "Content-Type": `multipart/form-data; boundary=${payload.boundary}`,
    "Content-Length": String(payload.body.length),
    Cookie: cookie,
  });

  assertCondition(response.status === 200, `Consolidado Laravel devolvio estado HTTP ${response.status}.`);
  if (/Falta generar salida reciente para esta accion/i.test(response.text)) {
    console.log("SKIP: consolidado Laravel necesita salidas recientes de las 4 acciones.");
    return;
  }

  const parsed = extractDownloadInfo(response.text, "Consolidado");
  const download = await httpRequestWithRedirects("GET", parsed.link, null, { Cookie: cookie });
  assertCondition(download.status === 200, `No se pudo descargar consolidado (${download.status}).`);

  const workbook = XLSX.read(download.buffer, { type: "buffer", cellFormula: false });
  assertCondition(
    workbook.SheetNames.length === 4,
    `Consolidado: se esperaban 4 hojas y llegaron ${workbook.SheetNames.length}.`,
  );
  assertNoExternalLinks(download.buffer, "Consolidado");
  assertWorkbookViewStructure(download.buffer, "Consolidado");
  assertStylesXmlReadable(download.buffer, "Consolidado");
  await assertExcelJsReadable(download.buffer, "Consolidado");
}

async function main() {
  const server = startPhpServer();
  try {
    await waitServerReady();
    await runAction1Contract();
    await runAction4Contract();
    await runBundleCheck();
    console.log("OK: contrato E2E Accion 1/4 validado en Laravel.");
  } finally {
    server.kill();
  }
}

main().catch((error) => {
  console.error(`FAIL: ${error.message}`);
  process.exit(1);
});
