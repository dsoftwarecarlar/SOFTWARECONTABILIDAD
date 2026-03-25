const fs = require("fs");
const http = require("http");
const path = require("path");
const crypto = require("crypto");
const { spawn } = require("child_process");

const XLSX = require("xlsx");

const ROOT = path.resolve(__dirname, "..", "..");
const APP_ROOT = path.join(ROOT, "laravel_app");
const ROUTER = path.join(APP_ROOT, "server.php");
const PHP_EXE = firstExistingPath([
  path.join(ROOT, ".tools", "php82", "php.exe"),
  process.env.PHP_EXE || "",
  "C:\\xampp\\php\\php.exe",
]);
const HOST = "127.0.0.1";
const PORT = 18993;
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

function repuestosContractPath(fileName) {
  return firstExistingPath([
    path.join(ROOT, "resources", "cxp", "repuestos_tytserv", "contracts", fileName),
  ]);
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

    const resolved = location.startsWith("http")
      ? new URL(location).pathname + new URL(location).search
      : (location.startsWith("/") ? location : `/${location}`);
    currentPath = resolved;
    currentMethod = "GET";
    currentBody = null;
  }

  throw new Error(`Demasiados redirects al solicitar ${requestPath}.`);
}

function buildMultipartBody(files, token) {
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
    push(
      `Content-Disposition: form-data; name="${file.field}"; filename="${path.basename(file.path)}"\r\n`,
    );
    push("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n");
    push(content);
    push("\r\n");
  }

  push(`--${boundary}--\r\n`);

  return {
    boundary,
    body: Buffer.concat(chunks),
  };
}

function extractCsrfToken(html) {
  const match = html.match(/name="_token"\s+value="([^"]+)"/i);
  assertCondition(match, "No se encontro CSRF token en el formulario Laravel de repuestos.");
  return match[1];
}

function extractCookie(headers) {
  const setCookie = headers["set-cookie"];
  assertCondition(Array.isArray(setCookie) && setCookie.length > 0, "Laravel no devolvio cookie de sesion.");
  return setCookie.map((cookie) => cookie.split(";", 1)[0]).join("; ");
}

function extractDownloadLink(html) {
  const match = html.match(/href="([^"]*downloads[^"]+)"/i);
  assertCondition(match, "No se encontro enlace de descarga en la respuesta Laravel.");
  const raw = match[1];
  if (raw.startsWith("http")) {
    return new URL(raw).pathname;
  }
  return raw.startsWith("/") ? raw : `/${raw}`;
}

function cellText(sheet, row, column) {
  const cell = sheet[XLSX.utils.encode_cell({ r: row - 1, c: column - 1 })];
  if (!cell) {
    return "";
  }

  const value = cell.w != null && String(cell.w).trim() !== "" ? cell.w : cell.v;
  return value == null ? "" : String(value).trim();
}

function rowContainsTotalGeneral(sheet, row, lastColumn = 41) {
  for (let column = 1; column <= lastColumn; column += 1) {
    if (/^TOTAL\s+GENERAL/i.test(cellText(sheet, row, column))) {
      return true;
    }
  }

  return false;
}

function payloadSignature(sheet) {
  const payloadColumns = [
    1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 13, 16,
    18, 20, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31,
    32, 33, 34, 35, 36, 37, 38, 39, 40,
  ];
  const rows = [];
  for (let row = 11; row <= 2000; row += 1) {
    if (rowContainsTotalGeneral(sheet, row)) {
      break;
    }

    const document = cellText(sheet, row, 5);
    if (document === "") {
      continue;
    }

    rows.push(payloadColumns.map((column) => cellText(sheet, row, column)).join("|"));
  }

  return {
    rows: rows.length,
    hash: crypto.createHash("sha256").update(rows.join("\n")).digest("hex"),
  };
}

function normalizedVisibleSheetSignature(sheet, ignoredHeaderCells = []) {
  const ref = sheet["!ref"] || "A1";
  const range = XLSX.utils.decode_range(ref);
  const rows = [];

  for (let row = range.s.r + 1; row <= range.e.r + 1; row += 1) {
    const values = [];
    let nonEmpty = false;
    for (let column = range.s.c + 1; column <= range.e.c + 1; column += 1) {
      let value = cellText(sheet, row, column);
      if (row === 2 && ignoredHeaderCells.includes(column)) {
        value = "";
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
    nonEmptyRows: rows.length,
    hash: crypto.createHash("sha256").update(rows.join("\n")).digest("hex"),
  };
}

function loadRepuestosContract() {
  const contractPath = repuestosContractPath("repuestos_tytserv_fixture_contract.json");
  assertCondition(fs.existsSync(contractPath), `No existe contrato de repuestos: ${contractPath}`);
  return JSON.parse(fs.readFileSync(contractPath, "utf8"));
}

function buildContractSources() {
  const sourceDir = firstExistingPath([
    path.join(ROOT, "resources", "cxp", "repuestos_tytserv", "fixtures"),
    path.join(ROOT, "outputs", "EJEMPLOAMANOTAREA3"),
  ]);
  const ncSourceDir = path.join(sourceDir, "archivosnc_rep");
  const files = [
    { field: "excel_tyt", path: path.join(sourceDir, "RepLibroVentasGeneral.xlsx"), sheet: "REP TYT" },
    { field: "excel_nc_tyt", path: path.join(ncSourceDir, "RepLibroDevolucionesGeneral.xlsx") },
    { field: "excel_peug", path: path.join(sourceDir, "RepLibroVentasGeneral (1).xlsx"), sheet: "REP PEUGT" },
    { field: "excel_nc_peug", path: path.join(ncSourceDir, "RepLibroDevolucionesGeneral (1).xlsx") },
    { field: "excel_chgn", path: path.join(sourceDir, "RepLibroVentasGeneral (2).xlsx"), sheet: "REP CHGN" },
    { field: "excel_nc_chgn", path: path.join(ncSourceDir, "RepLibroDevolucionesGeneral (2).xlsx") },
    { field: "excel_szk", path: path.join(sourceDir, "RepLibroVentasGeneral (3).xlsx"), sheet: "REP SZK" },
    { field: "excel_nc_szk", path: path.join(ncSourceDir, "RepLibroDevolucionesGeneral (3).xlsx") },
  ];

  for (const file of files) {
    assertCondition(fs.existsSync(file.path), `No existe archivo de contrato: ${file.path}`);
  }

  return files;
}

async function runContract() {
  const files = buildContractSources();
  const contract = loadRepuestosContract();
  const page = await httpRequest("GET", "/cxp/modules/repuestos-tytserv");
  assertCondition(page.status === 200, `Formulario Laravel devolvio estado ${page.status}.`);

  const cookie = extractCookie(page.headers);
  const token = extractCsrfToken(page.text);
  const payload = buildMultipartBody(files, token);
  const response = await httpRequest("POST", "/cxp/modules/repuestos-tytserv", payload.body, {
    "Content-Type": `multipart/form-data; boundary=${payload.boundary}`,
    "Content-Length": String(payload.body.length),
    Cookie: cookie,
  });

  assertCondition(response.status === 200, `Repuestos Laravel devolvio estado HTTP ${response.status}.`);
  const downloadLink = extractDownloadLink(response.text);
  const download = await httpRequestWithRedirects("GET", downloadLink, null, { Cookie: cookie });
  assertCondition(download.status === 200, `No se pudo descargar salida de repuestos (${download.status}).`);

  const outputWorkbook = XLSX.read(download.buffer, { type: "buffer", cellFormula: false, cellText: true, sheetStubs: true });
  assertCondition(
    JSON.stringify(outputWorkbook.SheetNames) === JSON.stringify(contract.sheet_order),
    `El orden de hojas no coincide con el contrato actual (${outputWorkbook.SheetNames.join(" | ")}).`,
  );

  for (const sourceFile of files) {
    if (!sourceFile.sheet) {
      continue;
    }

    const sourceWorkbook = XLSX.readFile(sourceFile.path, { cellFormula: false, cellText: true, sheetStubs: true });
    const sourceSheet = sourceWorkbook.Sheets[sourceWorkbook.SheetNames[0]];
    const outputSheet = outputWorkbook.Sheets[sourceFile.sheet];

    assertCondition(!!outputSheet, `No existe hoja ${sourceFile.sheet} en el Excel descargado.`);

    const sourceSignature = payloadSignature(sourceSheet);
    const outputSignature = payloadSignature(outputSheet);
    assertCondition(
      sourceSignature.rows === outputSignature.rows,
      `El payload de ${sourceFile.sheet} no coincide con el archivo subido (srcRows=${sourceSignature.rows}, outRows=${outputSignature.rows}).`,
    );
    assertCondition(
      sourceSignature.hash === outputSignature.hash,
      `El detalle visible de ${sourceFile.sheet} no coincide con el archivo subido.`,
    );
  }

  for (const [sheetName, expectedSignature] of Object.entries(contract.sheets || {})) {
    const outputSheet = outputWorkbook.Sheets[sheetName];
    assertCondition(!!outputSheet, `No existe hoja ${sheetName} en el Excel descargado.`);

    const ignoredHeaderCells = Array.isArray(contract.visible_signature_rules?.ignored_header_cells?.[sheetName])
      ? contract.visible_signature_rules.ignored_header_cells[sheetName]
      : [];
    const actualSignature = normalizedVisibleSheetSignature(outputSheet, ignoredHeaderCells);
    assertCondition(
      actualSignature.nonEmptyRows === expectedSignature.visible_non_empty_rows,
      `${sheetName}: filas visibles inesperadas (${actualSignature.nonEmptyRows} != ${expectedSignature.visible_non_empty_rows}).`,
    );
    assertCondition(
      actualSignature.hash === expectedSignature.visible_hash,
      `${sheetName}: hash visible no coincide con el contrato actual.`,
    );
  }
}

async function main() {
  const server = startPhpServer();
  try {
    await waitServerReady();
    await runContract();
    console.log("OK: contrato E2E repuestos validado en Laravel + Python.");
  } finally {
    server.kill();
  }
}

main().catch((error) => {
  console.error(`FAIL: ${error.message}`);
  process.exit(1);
});
