const fs = require("fs");
const http = require("http");
const os = require("os");
const path = require("path");
const { spawn, spawnSync } = require("child_process");

const XLSX = require("xlsx");

const {
  extractRowsFromTxt,
  loadTemplateTipoHints,
  normalizeParsedRows,
} = require("../cxp/accion2/parser");
const accion2Constants = require("../cxp/accion2/constants");
const { parseInputSources } = require("../cxp/accion3/parser");

const ROOT = path.resolve(__dirname, "..", "..");
const PHP_EXE = process.env.PHP_EXE || "C:\\xampp\\php\\php.exe";
const HOST = "127.0.0.1";
const PORT = 18987;
const BASE_URL = `http://${HOST}:${PORT}`;

function assertCondition(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

function normalizeText(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function startPhpServer() {
  const server = spawn(PHP_EXE, ["-S", `${HOST}:${PORT}`, "-t", ROOT], {
    cwd: ROOT,
    stdio: ["ignore", "pipe", "pipe"],
  });

  server.stdout.on("data", (chunk) => process.stdout.write(`[php] ${chunk}`));
  server.stderr.on("data", (chunk) => process.stderr.write(`[php] ${chunk}`));
  return server;
}

function waitServerReady(retries = 50) {
  return new Promise((resolve, reject) => {
    let attempts = 0;
    const tick = () => {
      attempts += 1;
      const req = http.request(
        {
          hostname: HOST,
          port: PORT,
          path: "/",
          method: "GET",
        },
        (res) => {
          res.resume();
          resolve();
        },
      );
      req.on("error", (error) => {
        if (attempts >= retries) {
          reject(new Error(`No se pudo iniciar servidor PHP: ${error.message}`));
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
      {
        hostname: HOST,
        port: PORT,
        path: requestPath,
        method,
        headers,
      },
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

function buildMultipartBody(fields, files) {
  const boundary = `----CodexBoundary${Date.now()}${Math.random().toString(16).slice(2)}`;
  const chunks = [];

  const push = (value) => {
    chunks.push(Buffer.isBuffer(value) ? value : Buffer.from(value, "utf8"));
  };

  for (const field of fields) {
    push(`--${boundary}\r\n`);
    push(`Content-Disposition: form-data; name="${field.name}"\r\n\r\n`);
    push(`${field.value}\r\n`);
  }

  for (const file of files) {
    const content = fs.readFileSync(file.path);
    push(`--${boundary}\r\n`);
    push(
      `Content-Disposition: form-data; name="${file.field}"; filename="${path.basename(file.path)}"\r\n`,
    );
    push("Content-Type: text/plain\r\n\r\n");
    push(content);
    push("\r\n");
  }

  push(`--${boundary}--\r\n`);

  return {
    boundary,
    body: Buffer.concat(chunks),
  };
}

function extractDownloadNameAndSuccess(html) {
  const linkMatch = html.match(/href="([^"]*download\.php\?file=[^"]+)"/i);
  assertCondition(linkMatch, "No se encontro enlace de descarga en la respuesta HTML.");

  const link = linkMatch[1];
  const absoluteUrl = link.startsWith("http") ? link : `${BASE_URL}${link.startsWith("/") ? "" : "/"}${link}`;
  const parsed = new URL(absoluteUrl);
  const fileName = decodeURIComponent(parsed.searchParams.get("file") || "");
  assertCondition(fileName !== "", "No se pudo extraer nombre de archivo desde enlace de descarga.");

  const successMatch = html.match(/Archivo generado correctamente:\s*([^|<\r\n]+)/i);
  assertCondition(successMatch, "No se encontro mensaje de exito con excel_name.");
  const successName = successMatch[1].trim();

  return {
    link: parsed.pathname + parsed.search,
    fileName,
    successName,
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

function fetchHistoryFromPhp(actionKey) {
  const phpCode = `require 'C:/xampp/htdocs/SOFTWARECONTABILIDAD/includes/app.php'; echo json_encode(app_list_output_files_for_action('${actionKey}', 20), JSON_UNESCAPED_SLASHES);`;
  const result = spawnSync(PHP_EXE, ["-r", phpCode], { cwd: ROOT, encoding: "utf8" });
  if (result.status !== 0) {
    throw new Error(`No se pudo leer historial de ${actionKey}: ${result.stderr || result.stdout}`);
  }
  const output = (result.stdout || "").trim();
  return output ? JSON.parse(output) : [];
}

function createAction2CanaryFile() {
  const sourcePath = path.join(ROOT, "outputs", "EJEMPLOSAMANO1", "CXPREP_RET_GENERALACCION2.txt");
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
  const sourcePath = path.join(ROOT, "outputs", "EJEMPLOSAMANO1", "CON_MAYORGEN2ACCION3.txt");
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
  const payload = buildMultipartBody([], [{ field: "txt_file", path: canaryInput.filePath }]);
  const response = await httpRequest("POST", "/modules/cxp_txt/index.php", payload.body, {
    "Content-Type": `multipart/form-data; boundary=${payload.boundary}`,
    "Content-Length": String(payload.body.length),
  });

  assertCondition(response.status === 200, `Accion 2 devolvio estado HTTP ${response.status}.`);
  const parsed = extractDownloadNameAndSuccess(response.text);
  assertCondition(
    parsed.successName === parsed.fileName,
    `Accion 2 inconsistente: success_message='${parsed.successName}' y descarga='${parsed.fileName}'.`,
  );

  const download = await httpRequest("GET", parsed.link);
  assertCondition(download.status === 200, `No se pudo descargar salida de Accion 2 (${download.status}).`);

  const workbook = XLSX.read(download.buffer, { type: "buffer" });
  const sheet = workbook.Sheets.RET_PROV || workbook.Sheets["RET PROV"];
  assertCondition(!!sheet, "Accion 2: no existe hoja RET PROV.");

  const providers = listColumnValues(sheet, 2, 2).map((item) => normalizeText(item));
  assertCondition(
    providers.includes(normalizeText(canaryInput.canary)),
    "Accion 2: el XLSX descargado no contiene el proveedor canario del TXT subido.",
  );

  const history = fetchHistoryFromPhp("accion2");
  const historyNames = history.map((item) => String(item.name || ""));
  assertCondition(
    historyNames.includes(parsed.fileName),
    "Accion 2: el historial no incluye el archivo recien generado.",
  );
  assertCondition(
    !historyNames.includes("ACCION2.xlsx"),
    "Accion 2: el historial esta incluyendo plantilla base.",
  );
}

async function runAction3Contract() {
  const canaryInput = await createAction3CanaryFile();
  const payload = buildMultipartBody([], [{ field: "txt_file[]", path: canaryInput.filePath }]);
  const response = await httpRequest("POST", "/modules/cxp_accion3/index.php", payload.body, {
    "Content-Type": `multipart/form-data; boundary=${payload.boundary}`,
    "Content-Length": String(payload.body.length),
  });

  assertCondition(response.status === 200, `Accion 3 devolvio estado HTTP ${response.status}.`);
  const parsed = extractDownloadNameAndSuccess(response.text);
  assertCondition(
    parsed.successName === parsed.fileName,
    `Accion 3 inconsistente: success_message='${parsed.successName}' y descarga='${parsed.fileName}'.`,
  );

  const download = await httpRequest("GET", parsed.link);
  assertCondition(download.status === 200, `No se pudo descargar salida de Accion 3 (${download.status}).`);

  const workbook = XLSX.read(download.buffer, { type: "buffer" });
  const sheet = workbook.Sheets["MAYOR RET"];
  assertCondition(!!sheet, "Accion 3: no existe hoja MAYOR RET.");

  const docValues = listColumnValues(sheet, 7, 2).map((item) => normalizeText(item));
  assertCondition(
    docValues.includes(normalizeText(canaryInput.canary)),
    "Accion 3: el XLSX descargado no contiene el DOCU canario del TXT subido.",
  );

  const history = fetchHistoryFromPhp("accion3");
  const historyNames = history.map((item) => String(item.name || ""));
  assertCondition(
    historyNames.includes(parsed.fileName),
    "Accion 3: el historial no incluye el archivo recien generado.",
  );
  assertCondition(
    !historyNames.includes("MAYOR RET_ACCION3.xlsx"),
    "Accion 3: el historial esta incluyendo plantilla base.",
  );
}

async function main() {
  const server = startPhpServer();
  try {
    await waitServerReady();
    await runAction2Contract();
    await runAction3Contract();
    console.log("OK: contrato E2E Accion 2/3 validado.");
  } finally {
    server.kill();
  }
}

main().catch((error) => {
  console.error(`FAIL: ${error.message}`);
  process.exit(1);
});
