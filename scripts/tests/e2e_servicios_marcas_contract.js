const fs = require("fs");
const http = require("http");
const os = require("os");
const path = require("path");
const crypto = require("crypto");
const { spawn, spawnSync } = require("child_process");

const XLSX = require("xlsx");

const ROOT = path.resolve(__dirname, "..", "..");
const PHP_EXE = process.env.PHP_EXE || "C:\\xampp\\php\\php.exe";
const HOST = "127.0.0.1";
const PORT = 18994;
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

function serviciosFixturePath(fileKey) {
  const candidatesByKey = {
    px: [
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "detalle-vtas-xliquidar (2).xlsx"),
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "detalle-vtas-xliquidar.xlsx"),
      path.join(ROOT, "outputs", "EJEMPLOAMANOTAREA2", "detalle-vtas-xliquidar.xlsx"),
    ],
    repventas: [
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "RepFacturacionServContabilidad (3).xls"),
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "RepFacturacionServContabilidad.xls"),
      path.join(ROOT, "outputs", "EJEMPLOAMANOTAREA2", "RepFacturacionServContabilidad.xls"),
    ],
    factura_tyt: [
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "archivosasubirtoy", "SERREP_FACTURAS_NAFTOY.TXT"),
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "SERREP_FACTURAS_NAF_REPFACT.txt"),
      path.join(ROOT, "outputs", "EJEMPLOAMANOTAREA2", "SERREP_FACTURAS_NAF_REPFACT.txt"),
    ],
    nota_tyt: [
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "archivosasubirtoy", "SERREP_NOTACRED_NAFTOY.TXT"),
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "SERREP_NOTACRED_NAF.txt"),
      path.join(ROOT, "outputs", "EJEMPLOAMANOTAREA2", "SERREP_NOTACRED_NAF.txt"),
    ],
    mayor_tyt: [
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "archivosasubirtoy", "CON_MAYORGEN2TOY.TXT"),
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "CON_MAYORGEN2TOY.TXT"),
      path.join(ROOT, "outputs", "EJEMPLOAMANOTAREA2", "CON_MAYORGEN2TOY.TXT"),
    ],
  };

  const candidates = candidatesByKey[fileKey] || [
    path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", fileKey),
    path.join(ROOT, "outputs", "EJEMPLOAMANOTAREA2", fileKey),
  ];

  return firstExistingPath(candidates);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
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
        setTimeout(tick, 250);
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

function buildMultipartBody(files, fields = {}) {
  const boundary = `----CodexBoundary${Date.now()}${Math.random().toString(16).slice(2)}`;
  const chunks = [];

  const push = (value) => {
    chunks.push(Buffer.isBuffer(value) ? value : Buffer.from(value, "utf8"));
  };

  for (const [field, value] of Object.entries(fields)) {
    push(`--${boundary}\r\n`);
    push(`Content-Disposition: form-data; name="${field}"\r\n\r\n`);
    push(String(value));
    push("\r\n");
  }

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

function createTempFileName(prefix, extension) {
  return path.join(
    os.tmpdir(),
    `${prefix}_${Date.now()}_${crypto.randomBytes(4).toString("hex")}.${extension}`,
  );
}

function setSheetTextCell(sheet, rowIndex, columnIndex, value) {
  const ref = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
  sheet[ref] = { t: "s", v: value, w: value };
}

function createFacturaCanaryFile() {
  const sourcePath = serviciosFixturePath("factura_tyt");
  const lines = fs.readFileSync(sourcePath, "utf8").split(/\r?\n/);

  let targetLine = -1;
  for (let index = 0; index < lines.length; index += 1) {
    const columns = lines[index].split("\t");
    if (
      normalizeText(columns[0]) === "AGENCIA :" &&
      normalizeText(columns[1]) === "MATRIZ" &&
      normalizeText(columns[22]) !== ""
    ) {
      targetLine = index;
      break;
    }
  }

  assertCondition(targetLine >= 0, "Servicios: no se encontro una fila MATRIZ para inyectar canario en REP FACTURACION.");

  const columns = lines[targetLine].split("\t");
  const canary = `TXTTYT_${Date.now().toString().slice(-8)}`;
  columns[24] = canary;
  lines[targetLine] = columns.join("\t");

  const targetPath = createTempFileName("servicios_factura_canary", "txt");
  fs.writeFileSync(targetPath, lines.join("\r\n"), "utf8");

  return { filePath: targetPath, canary };
}

function createPxCanaryFile() {
  const sourcePath = serviciosFixturePath("px");
  const workbook = XLSX.readFile(sourcePath, { cellFormula: false, cellNF: true, cellText: true });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

  let inToyotaSection = false;
  let targetRow = -1;
  for (let index = 0; index < rows.length; index += 1) {
    const row = Array.isArray(rows[index]) ? rows[index] : [];
    const marker = normalizeText(row[1]);
    const brand = normalizeText(row[3]);
    if (marker === "MARCA:") {
      inToyotaSection = brand === "TOYOTA";
      continue;
    }

    if (!inToyotaSection) {
      continue;
    }

    const agencia = normalizeText(row[1]);
    const factura = normalizeText(row[4]);
    const item = normalizeText(row[15]);
    if (/^\d+$/.test(agencia) && factura !== "" && item !== "" && item !== "ITEM") {
      targetRow = index;
      break;
    }
  }

  assertCondition(targetRow >= 0, "Servicios: no se encontro una fila TOYOTA para inyectar canario en PX.");

  const canary = `PXTYT_${Date.now().toString().slice(-8)}`;
  setSheetTextCell(sheet, targetRow, 15, canary);

  const targetPath = createTempFileName("servicios_px_canary", "xlsx");
  XLSX.writeFile(workbook, targetPath);

  return { filePath: targetPath, canary };
}

function activeServiciosJobs() {
  const jobsDir = path.join(ROOT, "storage", "jobs");
  if (!fs.existsSync(jobsDir)) {
    return [];
  }

  const active = [];
  for (const name of fs.readdirSync(jobsDir).filter((entry) => /^servicios_marcas_.*\.json$/i.test(entry))) {
    const filePath = path.join(jobsDir, name);
    try {
      const raw = fs.readFileSync(filePath, "utf8").trim();
      if (raw === "") {
        continue;
      }

      const payload = JSON.parse(raw);
      const status = String(payload.status || "");
      if (["queued", "running", "cancel_requested"].includes(status)) {
        active.push({
          name,
          jobId: String(payload.job_id || ""),
          status,
        });
      }
    } catch (_error) {
      // Ignore partially written JSON snapshots while the worker updates its state.
    }
  }

  return active;
}

async function ensureNoActiveServiciosJobs(timeoutMs = 120000) {
  const startedAt = Date.now();
  let active = activeServiciosJobs();
  while (active.length > 0) {
    if (Date.now() - startedAt > timeoutMs) {
      const labels = active.map((job) => `${job.jobId}:${job.status}`).join(", ");
      throw new Error(`Servicios por Marca tiene jobs activos bloqueando la prueba: ${labels}`);
    }
    await sleep(3000);
    active = activeServiciosJobs();
  }
}

async function pollJobUntilFinished(jobId, timeoutMs = 900000, intervalMs = 5000) {
  const startedAt = Date.now();
  let lastDigest = "";
  let lastLogAt = 0;

  while (Date.now() - startedAt < timeoutMs) {
    const response = await httpRequest(
      "GET",
      `/modules/cxp_servicios_marcas/index.php?status=${encodeURIComponent(jobId)}`,
    );
    assertCondition(response.status === 200, `Servicios: status poll devolvio HTTP ${response.status}.`);
    const job = JSON.parse(response.text);
    const digest = `${job.status}|${job.message || ""}`;
    if (digest !== lastDigest || Date.now() - lastLogAt >= 60000) {
      console.log(`INFO|servicios_job|${job.status}|${job.message || ""}`);
      lastDigest = digest;
      lastLogAt = Date.now();
    }

    if (["complete", "error", "cancelled"].includes(job.status)) {
      return job;
    }

    await sleep(intervalMs);
  }

  throw new Error(`Servicios: timeout esperando el job ${jobId}.`);
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

function sheetContainsText(sheet, expected) {
  const needle = normalizeText(expected);
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  return rows.some((row) =>
    (Array.isArray(row) ? row : []).some((cell) => normalizeText(cell).includes(needle)),
  );
}

function countRowsByAgency(sheet, agencyLabel) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  const agency = normalizeText(agencyLabel);
  let count = 0;
  for (const row of rows) {
    const cells = Array.isArray(row) ? row : [];
    if (normalizeText(cells[0]) === agency && normalizeText(cells[2]) !== "") {
      count += 1;
    }
  }
  return count;
}

async function runServiciosContract() {
  await ensureNoActiveServiciosJobs();

  const facturaCanary = createFacturaCanaryFile();
  const pxCanary = createPxCanaryFile();
  const cleanupPaths = [facturaCanary.filePath, pxCanary.filePath];

  try {
    const files = [
      {
        field: "factura_tyt_file",
        path: facturaCanary.filePath,
        contentType: "text/plain",
      },
      {
        field: "nota_tyt_file",
        path: serviciosFixturePath("nota_tyt"),
        contentType: "text/plain",
      },
      {
        field: "px_file",
        path: pxCanary.filePath,
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
      {
        field: "repventas_file",
        path: serviciosFixturePath("repventas"),
        contentType: "application/vnd.ms-excel",
      },
      {
        field: "mayor_tyt_file",
        path: serviciosFixturePath("mayor_tyt"),
        contentType: "text/plain",
      },
    ];

    const payload = buildMultipartBody(files, {
      action: "process",
      brand_key: "tyt",
    });

    const processResponse = await httpRequest("POST", "/modules/cxp_servicios_marcas/index.php", payload.body, {
      "Content-Type": `multipart/form-data; boundary=${payload.boundary}`,
      "Content-Length": String(payload.body.length),
    });

    assertCondition(
      processResponse.status === 302,
      `Servicios: el POST inicial debe redirigir al job y devolvio ${processResponse.status}.`,
    );

    const location = String(processResponse.headers.location || "");
    const match = location.match(/[?&]job=([^&]+)/);
    assertCondition(match, "Servicios: no se encontro el job en la cabecera Location.");
    const jobId = decodeURIComponent(match[1]);

    const job = await pollJobUntilFinished(jobId);
    assertCondition(job.status === "complete", `Servicios: el job termino en estado ${job.status}.`);
    assertCondition(Array.isArray(job.downloads) && job.downloads.length === 1, "Servicios: se esperaba una sola salida MATRIZ para brand_key=tyt.");
    assertCondition(
      Array.isArray(job.summary) && job.summary.length === 1 && String(job.summary[0].key || "") === "tyt",
      "Servicios: la respuesta del job no devolvio resumen exclusivo de tyt.",
    );

    const download = job.downloads[0];
    assertCondition(
      String(download.label || "") === "MATRIZ",
      `Servicios: la salida debe etiquetarse como MATRIZ y llego '${download.label || ""}'.`,
    );
    assertCondition(
      /^servicios_tyt_.*\.xls$/i.test(String(download.name || "")),
      `Servicios: nombre de salida inesperado '${download.name || ""}'.`,
    );

    const consoleText = String(job.console || "");
    assertCondition(
      consoleText.includes(`OUTPUT|${download.name}|MATRIZ`),
      "Servicios: la consola del job no registro la salida esperada.",
    );
    assertCondition(
      consoleText.includes("INFO|processing|tyt|rows="),
      "Servicios: la consola del job no registro el procesamiento de tyt.",
    );

    const pageResponse = await httpRequest(
      "GET",
      `/modules/cxp_servicios_marcas/index.php?job=${encodeURIComponent(jobId)}`,
    );
    assertCondition(pageResponse.status === 200, `Servicios: la pagina final del job devolvio ${pageResponse.status}.`);
    assertCondition(
      pageResponse.text.includes(String(download.name || "")) && pageResponse.text.includes(String(download.download_url || "")),
      "Servicios: la pagina final no publico el enlace de descarga del job.",
    );

    const downloadUrl = String(download.download_url || "");
    const absoluteDownloadUrl = downloadUrl.startsWith("http")
      ? downloadUrl
      : `${BASE_URL}${downloadUrl.startsWith("/") ? "" : "/"}${downloadUrl}`;
    const parsedDownload = new URL(absoluteDownloadUrl);
    const outputResponse = await httpRequest("GET", parsedDownload.pathname + parsedDownload.search);
    assertCondition(outputResponse.status === 200, `Servicios: no se pudo descargar la salida (${outputResponse.status}).`);

    const workbook = XLSX.read(outputResponse.buffer, { type: "buffer", cellFormula: false, cellNF: true, cellText: true });
    const repFactSheet = workbook.Sheets["REP FACTURACION"] || workbook.Sheets["REP FACTURACIÓN"];
    const notaSheet = workbook.Sheets["NOTA DE CREDITO"] || workbook.Sheets["NOTA DE CRÉDITO"];
    const pxSheet = workbook.Sheets.PX;
    const repVtasSheet = workbook.Sheets["REP VTAS"];
    const precontVentasSheet = workbook.Sheets.PrecontabilizacionVentas;

    assertCondition(!!repFactSheet, "Servicios: la salida no contiene la hoja REP FACTURACION.");
    assertCondition(!!notaSheet, "Servicios: la salida no contiene la hoja NOTA DE CREDITO.");
    assertCondition(!!pxSheet, "Servicios: la salida no contiene la hoja PX.");
    assertCondition(!!repVtasSheet, "Servicios: la salida no contiene la hoja REP VTAS.");
    assertCondition(!!precontVentasSheet, "Servicios: la salida no contiene la hoja PrecontabilizacionVentas.");

    assertCondition(
      sheetContainsText(repFactSheet, facturaCanary.canary),
      "Servicios: la hoja REP FACTURACION no contiene el cliente canario del TXT subido.",
    );
    assertCondition(
      sheetContainsText(pxSheet, pxCanary.canary),
      "Servicios: la hoja PX no contiene el item canario del upload PX.",
    );
    assertCondition(
      !sheetContainsText(precontVentasSheet, "#REF!"),
      "Servicios: PrecontabilizacionVentas contiene #REF! y rompe la conciliacion.",
    );

    const repFactRows = countRowsByAgency(repFactSheet, "MATRIZ");
    const noteRows = countRowsByAgency(notaSheet, "MATRIZ");
    const repVtasRows = countRowsByAgency(repVtasSheet, "MATRIZ");

    assertCondition(repFactRows > 0, "Servicios: la hoja REP FACTURACION no contiene filas MATRIZ.");
    assertCondition(noteRows > 0, "Servicios: la hoja NOTA DE CREDITO no contiene filas MATRIZ.");
    assertCondition(repVtasRows > 0, "Servicios: la hoja REP VTAS no contiene filas MATRIZ.");
    assertCondition(
      Number(job.summary[0].rows || 0) === repVtasRows,
      `Servicios: resumen (${job.summary[0].rows || 0}) no coincide con REP VTAS (${repVtasRows}).`,
    );

    const history = fetchHistoryFromPhp("servicios");
    const historyNames = history.map((item) => String(item.name || ""));
    assertCondition(
      historyNames.includes(String(download.name || "")),
      "Servicios: el historial no incluye la salida recien generada.",
    );

    const persistedOutputPath = path.join(ROOT, "storage", "outputs", String(download.name || ""));
    if (fs.existsSync(persistedOutputPath)) {
      cleanupPaths.push(persistedOutputPath);
    }
  } finally {
    for (const filePath of cleanupPaths) {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    }
  }
}

async function main() {
  const server = startPhpServer();
  try {
    await waitServerReady();
    await runServiciosContract();
    console.log("OK: contrato E2E Servicios por Marca validado.");
  } finally {
    server.kill();
  }
}

main().catch((error) => {
  console.error(`FAIL: ${error.message}`);
  process.exit(1);
});
