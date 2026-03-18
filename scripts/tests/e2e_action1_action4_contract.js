const fs = require("fs");
const http = require("http");
const os = require("os");
const path = require("path");
const crypto = require("crypto");
const { spawn, spawnSync } = require("child_process");

const XLSX = require("xlsx");

const { parseInputSources } = require("../cxp/accion4/parser");

const ROOT = path.resolve(__dirname, "..", "..");
const PHP_EXE = process.env.PHP_EXE || "C:\\xampp\\php\\php.exe";
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
    if (candidate && fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return candidates[0];
}

function actionFixturePath(fileName) {
  return firstExistingPath([
    path.join(ROOT, "resources", "cxp", "acciones", "fixtures", fileName),
    path.join(ROOT, "outputs", "EJEMPLOSAMANO1", fileName),
  ]);
}

function normalizeText(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function round2(value) {
  return Math.round((Number(value || 0) + Number.EPSILON) * 100) / 100;
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

function waitServerReady(retries = 60) {
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

function buildMultipartBody(files) {
  const boundary = `----CodexBoundary${Date.now()}${Math.random().toString(16).slice(2)}`;
  const chunks = [];

  const push = (value) => {
    chunks.push(Buffer.isBuffer(value) ? value : Buffer.from(value, "utf8"));
  };

  for (const file of files) {
    const content = fs.readFileSync(file.path);
    push(`--${boundary}\r\n`);
    push(
      `Content-Disposition: form-data; name="${file.field}"; filename="${path.basename(file.path)}"\r\n`,
    );
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

function extractDownloadNameAndSuccess(html) {
  const linkMatch = html.match(/href="([^"]*download\.php\?file=[^"]+)"/i);
  assertCondition(linkMatch, "No se encontro enlace de descarga en la respuesta HTML.");

  const link = linkMatch[1];
  const absoluteUrl = link.startsWith("http")
    ? link
    : `${BASE_URL}${link.startsWith("/") ? "" : "/"}${link}`;
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

function fetchHistoryFromPhp(actionKey) {
  const phpCode = `require 'C:/xampp/htdocs/SOFTWARECONTABILIDAD/includes/app.php'; echo json_encode(app_list_output_files_for_action('${actionKey}', 20), JSON_UNESCAPED_SLASHES);`;
  const result = spawnSync(PHP_EXE, ["-r", phpCode], { cwd: ROOT, encoding: "utf8" });
  if (result.status !== 0) {
    throw new Error(`No se pudo leer historial de ${actionKey}: ${result.stderr || result.stdout}`);
  }
  const output = (result.stdout || "").trim();
  return output ? JSON.parse(output) : [];
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
  const inputPath = path.join(ROOT, "CXPREP_docproveedor.pdf");
  const payload = buildMultipartBody([
    { field: "pdf_file", path: inputPath, contentType: "application/pdf" },
  ]);

  const response = await httpRequest("POST", "/modules/cxp_pdf/index.php", payload.body, {
    "Content-Type": `multipart/form-data; boundary=${payload.boundary}`,
    "Content-Length": String(payload.body.length),
  });

  assertCondition(response.status === 200, `Accion 1 devolvio estado HTTP ${response.status}.`);
  const parsed = extractDownloadNameAndSuccess(response.text);
  assertCondition(
    parsed.successName === parsed.fileName,
    `Accion 1 inconsistente: success_message='${parsed.successName}' y descarga='${parsed.fileName}'.`,
  );

  const download = await httpRequest("GET", parsed.link);
  assertCondition(download.status === 200, `No se pudo descargar salida de Accion 1 (${download.status}).`);

  const outputWorkbook = XLSX.read(download.buffer, { type: "buffer", cellFormula: false });
  assertCondition(
    outputWorkbook.SheetNames.includes("LIBRO COMPRAS"),
    "Accion 1: no existe hoja LIBRO COMPRAS en el Excel descargado.",
  );

  const generatedSignature = buildAction1NormalizedSignature(download.buffer);
  const referenceSignature = buildAction1NormalizedSignature(
    path.join(ROOT, "outputs", "CXPREP_docproveedor_20260306_092810_resultado.xlsx"),
  );
  assertCondition(
    generatedSignature.rows === referenceSignature.rows && generatedSignature.hash === referenceSignature.hash,
    `Accion 1: la salida no coincide con el contrato base (rows=${generatedSignature.rows}/${referenceSignature.rows}).`,
  );

  const history = fetchHistoryFromPhp("accion1");
  const historyNames = history.map((item) => String(item.name || ""));
  assertCondition(
    historyNames.includes(parsed.fileName),
    "Accion 1: el historial no incluye el archivo recien generado.",
  );
}

async function runAction4Contract() {
  const canaryInput = await createAction4CanaryFile();
  const payload = buildMultipartBody([
    { field: "txt_file", path: canaryInput.filePath, contentType: "text/plain" },
  ]);

  const response = await httpRequest("POST", "/modules/cxp_accion4/index.php", payload.body, {
    "Content-Type": `multipart/form-data; boundary=${payload.boundary}`,
    "Content-Length": String(payload.body.length),
  });

  assertCondition(response.status === 200, `Accion 4 devolvio estado HTTP ${response.status}.`);
  const parsed = extractDownloadNameAndSuccess(response.text);
  assertCondition(
    parsed.successName === parsed.fileName,
    `Accion 4 inconsistente: success_message='${parsed.successName}' y descarga='${parsed.fileName}'.`,
  );

  const download = await httpRequest("GET", parsed.link);
  assertCondition(download.status === 200, `No se pudo descargar salida de Accion 4 (${download.status}).`);

  const workbook = XLSX.read(download.buffer, { type: "buffer" });
  const sheet = workbook.Sheets["MAYOR IVA"];
  assertCondition(!!sheet, "Accion 4: no existe hoja MAYOR IVA.");

  const docValues = columnValues(sheet, 7, 2).map((item) => normalizeText(item));
  assertCondition(
    docValues.includes(normalizeText(canaryInput.canary)),
    "Accion 4: el XLSX descargado no contiene el DOCU canario del TXT subido.",
  );

  const history = fetchHistoryFromPhp("accion4");
  const historyNames = history.map((item) => String(item.name || ""));
  assertCondition(
    historyNames.includes(parsed.fileName),
    "Accion 4: el historial no incluye el archivo recien generado.",
  );
}

async function runBundleCheck() {
  const response = await httpRequest("GET", "/export_all_actions.php");
  if (response.status === 409) {
    console.log("SKIP: export_all_actions necesita salidas recientes de las 4 acciones.");
    return;
  }

  assertCondition(response.status === 200, `Consolidado devolvio estado HTTP ${response.status}.`);
  const workbook = XLSX.read(response.buffer, { type: "buffer", cellFormula: false });
  assertCondition(
    workbook.SheetNames.length === 4,
    `Consolidado: se esperaban 4 hojas y llegaron ${workbook.SheetNames.length}.`,
  );
}

async function main() {
  const server = startPhpServer();
  try {
    await waitServerReady();
    await runAction1Contract();
    await runAction4Contract();
    await runBundleCheck();
    console.log("OK: contrato E2E Accion 1/4 validado.");
  } finally {
    server.kill();
  }
}

main().catch((error) => {
  console.error(`FAIL: ${error.message}`);
  process.exit(1);
});
