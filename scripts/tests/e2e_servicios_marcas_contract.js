const fs = require("fs");
const http = require("http");
const os = require("os");
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
const PORT = 18994;
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

function serviciosTemplatePath(brandKey) {
  const candidatesByKey = {
    tyt: [
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "templates", "11. Concili. Servicios TYT 2026.xls"),
    ],
  };

  return firstExistingPath(candidatesByKey[brandKey] || []);
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

function round2(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

function assertClose(actual, expected, label, tolerance = 0.02) {
  const actualRounded = round2(actual);
  const expectedRounded = round2(expected);
  assertCondition(
    Math.abs(actualRounded - expectedRounded) <= tolerance,
    `${label}: esperado ${expectedRounded}, actual ${actualRounded}.`,
  );
}

function parseDecimalLike(raw) {
  let normalized = String(raw || "").replace(/[^\d,.\-]/g, "");
  if (normalized === "") {
    return 0;
  }

  const hasComma = normalized.includes(",");
  const hasDot = normalized.includes(".");
  if (hasComma && hasDot) {
    if (normalized.lastIndexOf(".") > normalized.lastIndexOf(",")) {
      normalized = normalized.replace(/,/g, "");
    } else {
      normalized = normalized.replace(/\./g, "").replace(",", ".");
    }
  } else if (hasComma) {
    const commaCount = (normalized.match(/,/g) || []).length;
    if (commaCount > 1) {
      normalized = normalized.replace(/,/g, "");
    } else {
      const [intPart, fracPart = ""] = normalized.split(",");
      normalized = fracPart.length === 3 ? `${intPart}${fracPart}` : `${intPart}.${fracPart}`;
    }
  } else if (hasDot) {
    const dotCount = (normalized.match(/\./g) || []).length;
    if (dotCount > 1) {
      const lastDot = normalized.lastIndexOf(".");
      const intPart = normalized.slice(0, lastDot).replace(/\./g, "");
      const fracPart = normalized.slice(lastDot + 1);
      normalized = fracPart.length === 3 ? `${intPart}${fracPart}` : `${intPart}.${fracPart}`;
    }
  }

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function readMayorRows(filePath) {
  return fs
    .readFileSync(filePath, "utf8")
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => line.split("\t"))
    .filter((columns) => columns.length >= 30 && /^\d{2}\.\d{2}\.\d{2}\.\d{2}\.\d{4}$/.test(String(columns[6] || "").trim()))
    .map((columns) => ({
      account: String(columns[6] || "").trim(),
      name: String(columns[7] || "").trim(),
      origin: String(columns[23] || "").trim().toUpperCase(),
      seat: String(columns[24] || "").trim(),
      detail: String(columns[26] || "").trim().toUpperCase(),
      debit: parseDecimalLike(columns[27]),
      credit: parseDecimalLike(columns[28]),
    }));
}

function compactAccountCode(account) {
  const digits = String(account || "").replace(/\D/g, "");
  if (digits.length > 0 && digits.length < 12) {
    return digits.padStart(12, "0");
  }
  return digits;
}

function isMayorPxAdjustmentRow(row) {
  const account = compactAccountCode(row.account);
  if (!/^040101\d{2}(0003|0012)$/.test(account)) {
    return false;
  }
  if (String(row.detail || "").includes("REGISTRO DE PX AJUSTE DE EGRESO")) {
    return true;
  }
  return String(row.origin || "") === "AGCM" && String(row.seat || "") === "435";
}

function filterMayorRowsForWorkbook(rows) {
  return rows.filter((row) => !isMayorPxAdjustmentRow(row));
}

function classifyMayorControlBucket(row) {
  const account = String(row.account || "").trim();
  const name = String(row.name || "").trim().toUpperCase();
  if (!account && !name) {
    return "";
  }
  if (/^01\.01\.05\.\d{2}\.\d{4}$/.test(account) || name.includes("GARANT")) {
    return "guarantee";
  }
  if (!/^04\.01\.01\.\d{2}\.\d{4}$/.test(account)) {
    return "";
  }
  const suffix = account.split(".").pop() || "";
  if (suffix === "0014" || name.includes("DEVOL")) {
    return "return";
  }
  if (["0010", "0011", "0012"].includes(suffix) || name.includes("DESC")) {
    return "discount";
  }
  return "sales";
}

function getMayorControlMetrics(rows) {
  const metrics = {
    InvoiceSales: 0,
    InvoiceDiscounts: 0,
    NoteSales: 0,
    NoteDiscounts: 0,
    NetSales: 0,
  };

  for (const row of rows) {
    const bucket = classifyMayorControlBucket(row);
    const debit = Number(row.debit || 0);
    const credit = Number(row.credit || 0);
    if (bucket === "sales") {
      metrics.InvoiceSales += credit;
      metrics.NetSales += credit - debit;
      continue;
    }
    if (bucket === "discount") {
      metrics.InvoiceDiscounts += debit;
      metrics.NoteDiscounts += credit;
      metrics.NetSales += credit - debit;
      continue;
    }
    if (bucket === "return") {
      metrics.NoteSales += debit;
      metrics.NetSales += credit - debit;
    }
  }

  return Object.fromEntries(Object.entries(metrics).map(([key, value]) => [key, round2(value)]));
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

    const parsed = location.startsWith("http")
      ? new URL(location)
      : new URL(location, BASE_URL);
    currentPath = parsed.pathname + parsed.search;
    currentMethod = "GET";
    currentBody = null;
  }

  throw new Error(`Demasiados redirects al solicitar ${requestPath}.`);
}

function buildMultipartBody(files, token, fields = {}) {
  const boundary = `----CodexBoundary${Date.now()}${Math.random().toString(16).slice(2)}`;
  const chunks = [];

  const push = (value) => {
    chunks.push(Buffer.isBuffer(value) ? value : Buffer.from(value, "utf8"));
  };

  push(`--${boundary}\r\n`);
  push('Content-Disposition: form-data; name="_token"\r\n\r\n');
  push(token);
  push("\r\n");

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

function extractCsrfToken(html) {
  const match = html.match(/name="_token"\s+value="([^"]+)"/i);
  assertCondition(match, "No se encontro CSRF token en el formulario Laravel de servicios.");
  return match[1];
}

function extractCookie(headers) {
  const setCookie = headers["set-cookie"];
  assertCondition(Array.isArray(setCookie) && setCookie.length > 0, "Laravel no devolvio cookie de sesion.");
  return setCookie.map((cookie) => cookie.split(";", 1)[0]).join("; ");
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

async function pollJobUntilFinished(jobId, cookie, timeoutMs = 900000, intervalMs = 5000) {
  const startedAt = Date.now();
  let lastDigest = "";
  let lastLogAt = 0;

  while (Date.now() - startedAt < timeoutMs) {
    const response = await httpRequest(
      "GET",
      `/cxp/modules/servicios-marcas/jobs/${encodeURIComponent(jobId)}`,
      null,
      { Cookie: cookie },
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
    const hasPayload = cells.some((cell, index) => index > 0 && normalizeText(cell) !== "");
    if (normalizeText(cells[0]) === agency && hasPayload) {
      count += 1;
    }
  }
  return count;
}

function cellText(sheet, row, column) {
  const ref = XLSX.utils.encode_cell({ r: row - 1, c: column - 1 });
  const cell = sheet?.[ref];
  if (!cell) {
    return "";
  }

  const value = cell.w != null && String(cell.w).trim() !== "" ? cell.w : cell.v;
  return value == null ? "" : String(value).trim();
}

function cellNumber(sheet, row, column) {
  return round2(parseDecimalLike(cellText(sheet, row, column)));
}

function findSheet(workbook, candidates) {
  for (const candidate of candidates) {
    if (workbook.Sheets[candidate]) {
      return workbook.Sheets[candidate];
    }
  }

  return null;
}

function assertSeedColumnsMatch(outputSheet, templateSheet, rowNumbers, columns, label) {
  for (const rowNumber of rowNumbers) {
    for (const column of columns) {
      const expected = cellText(templateSheet, rowNumber, column);
      const actual = cellText(outputSheet, rowNumber, column);
      assertCondition(
        normalizeText(actual) === normalizeText(expected),
        `${label}: la fila ${rowNumber} columna ${column} no conserva el seed de plantilla.`,
      );
    }
  }
}

function findPrecontCostos2Rows(sheet) {
  const rows = [];
  for (let row = 2; row <= 60; row += 1) {
    const account = cellText(sheet, row, 4);
    if (account === "") {
      continue;
    }

    rows.push({
      row,
      account,
      ag: cellText(sheet, row, 2),
      line: cellText(sheet, row, 3),
      number: cellText(sheet, row, 5),
      description: cellText(sheet, row, 6),
      costCenter: cellText(sheet, row, 7),
      asiento: cellText(sheet, row, 10),
    });
  }

  return rows;
}

function assertPrecontCostos2Seeds(outputSheet, templateSheet) {
  const templateRows = findPrecontCostos2Rows(templateSheet);
  const outputRows = findPrecontCostos2Rows(outputSheet);
  const expectedAccounts = ["050201010001", "050201010002", "050201010003", "050201010005"];

  for (const account of expectedAccounts) {
    const expected = templateRows.find((row) => normalizeText(row.account) === account);
    const actual = outputRows.find((row) => normalizeText(row.account) === account);
    assertCondition(!!expected, `PrecontabilizacionCostos (2): la plantilla no contiene la cuenta esperada ${account}.`);
    assertCondition(!!actual, `PrecontabilizacionCostos (2): la salida no contiene la cuenta esperada ${account}.`);
    assertCondition(
      normalizeText(actual.ag) === normalizeText(expected.ag) &&
      normalizeText(actual.line) === normalizeText(expected.line) &&
      normalizeText(actual.number) === normalizeText(expected.number) &&
      normalizeText(actual.description) === normalizeText(expected.description) &&
      normalizeText(actual.costCenter) === normalizeText(expected.costCenter) &&
      normalizeText(actual.asiento) === normalizeText(expected.asiento),
      `PrecontabilizacionCostos (2): la estructura seed de la cuenta ${account} no coincide con la plantilla.`,
    );
  }
}

async function runServiciosContract() {
  await ensureNoActiveServiciosJobs();

  const facturaCanary = createFacturaCanaryFile();
  const pxCanary = createPxCanaryFile();
  const cleanupPaths = [facturaCanary.filePath, pxCanary.filePath];

  try {
    const formResponse = await httpRequest("GET", "/cxp/modules/servicios-marcas");
    assertCondition(formResponse.status === 200, `Servicios: formulario Laravel devolvio HTTP ${formResponse.status}.`);
    const cookie = extractCookie(formResponse.headers);
    const token = extractCsrfToken(formResponse.text);

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

    const payload = buildMultipartBody(files, token, {
      action: "process",
      brand_key: "tyt",
    });

    const processResponse = await httpRequest("POST", "/cxp/modules/servicios-marcas", payload.body, {
      "Content-Type": `multipart/form-data; boundary=${payload.boundary}`,
      "Content-Length": String(payload.body.length),
      Cookie: cookie,
    });

    assertCondition(
      processResponse.status === 302,
      `Servicios: el POST inicial debe redirigir al job y devolvio ${processResponse.status}.`,
    );

    const location = String(processResponse.headers.location || "");
    const match = location.match(/[?&]job=([^&]+)/);
    assertCondition(match, "Servicios: no se encontro el job en la cabecera Location.");
    const jobId = decodeURIComponent(match[1]);

    const job = await pollJobUntilFinished(jobId, cookie);
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
      `/cxp/modules/servicios-marcas?job=${encodeURIComponent(jobId)}`,
      null,
      { Cookie: cookie },
    );
    assertCondition(pageResponse.status === 200, `Servicios: la pagina final del job devolvio ${pageResponse.status}.`);
    assertCondition(
      pageResponse.text.includes(String(download.name || "")) && pageResponse.text.includes(String(download.download_url || "")),
      "Servicios: la pagina final no publico el enlace de descarga del job.",
    );

    const historyPage = await httpRequest("GET", "/cxp/modules/servicios-marcas", null, { Cookie: cookie });
    assertCondition(historyPage.status === 200, `Servicios: la pagina base devolvio ${historyPage.status}.`);
    assertCondition(
      historyPage.text.includes(String(download.name || "")),
      "Servicios: el historial Laravel no incluye la salida recien generada.",
    );

    const downloadUrl = String(download.download_url || "");
    const absoluteDownloadUrl = downloadUrl.startsWith("http")
      ? downloadUrl
      : `${BASE_URL}${downloadUrl.startsWith("/") ? "" : "/"}${downloadUrl}`;
    const parsedDownload = new URL(absoluteDownloadUrl);
    const outputResponse = await httpRequestWithRedirects(
      "GET",
      parsedDownload.pathname + parsedDownload.search,
      null,
      { Cookie: cookie },
    );
    assertCondition(outputResponse.status === 200, `Servicios: no se pudo descargar la salida (${outputResponse.status}).`);

    const workbook = XLSX.read(outputResponse.buffer, { type: "buffer", cellFormula: false, cellNF: true, cellText: true });
    const repFactSheet =
      workbook.Sheets["REP FACTURACION"] ||
      workbook.Sheets["REP FACTURACIÓN"] ||
      workbook.Sheets["REP FACTURACIÃ“N"];
    const notaSheet = workbook.Sheets["NOTA DE CREDITO"] || workbook.Sheets["NOTA DE CRÃ‰DITO"];
    const pxSheet = workbook.Sheets.PX;
    const repVtasSheet = workbook.Sheets["REP VTAS"];
    const precontVentasSheet = workbook.Sheets.PrecontabilizacionVentas;
    const precontCostos2Sheet = workbook.Sheets["PrecontabilizacionCostos (2)"];
    const costoSheet = workbook.Sheets.COSTO;
    const estadisticasSheet = workbook.Sheets.ESTADISTICAS;

    assertCondition(!!repFactSheet, "Servicios: la salida no contiene la hoja REP FACTURACION.");
    assertCondition(!!notaSheet, "Servicios: la salida no contiene la hoja NOTA DE CREDITO.");
    assertCondition(!!pxSheet, "Servicios: la salida no contiene la hoja PX.");
    assertCondition(!!repVtasSheet, "Servicios: la salida no contiene la hoja REP VTAS.");
    assertCondition(!!precontVentasSheet, "Servicios: la salida no contiene la hoja PrecontabilizacionVentas.");
    assertCondition(!!precontCostos2Sheet, "Servicios: la salida no contiene la hoja PrecontabilizacionCostos (2).");
    assertCondition(!!costoSheet, "Servicios: la salida no contiene la hoja COSTO.");
    assertCondition(!!estadisticasSheet, "Servicios: la salida no contiene la hoja ESTADISTICAS.");

    const templateWorkbook = XLSX.readFile(serviciosTemplatePath("tyt"), { cellFormula: false, cellNF: true, cellText: true });
    const templatePrecontCostos2Sheet = findSheet(templateWorkbook, ["PrecontabilizacionCostos (2)"]);
    const templateCostoSheet = findSheet(templateWorkbook, ["COSTO"]);
    const templateEstadisticasSheet = findSheet(templateWorkbook, ["ESTADISTICAS"]);

    assertCondition(!!templatePrecontCostos2Sheet, "Servicios: no se pudo abrir la plantilla base de PrecontabilizacionCostos (2).");
    assertCondition(!!templateCostoSheet, "Servicios: no se pudo abrir la plantilla base de COSTO.");
    assertCondition(!!templateEstadisticasSheet, "Servicios: no se pudo abrir la plantilla base de ESTADISTICAS.");

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

    assertSeedColumnsMatch(costoSheet, templateCostoSheet, [6, 7, 8, 9], [1, 2, 4, 5, 7], "COSTO");
    assertSeedColumnsMatch(estadisticasSheet, templateEstadisticasSheet, [6, 131, 222, 290, 365], [1, 2, 5, 6, 8], "ESTADISTICAS");
    assertPrecontCostos2Seeds(precontCostos2Sheet, templatePrecontCostos2Sheet);

    const mayorMetrics = getMayorControlMetrics(filterMayorRowsForWorkbook(readMayorRows(serviciosFixturePath("mayor_tyt"))));
    assertClose(cellNumber(repFactSheet, 9, 4), mayorMetrics.InvoiceSales, "Servicios REP FACTURACION D9");
    assertClose(cellNumber(repFactSheet, 9, 5), mayorMetrics.InvoiceDiscounts, "Servicios REP FACTURACION E9");
    assertClose(cellNumber(notaSheet, 4, 4), mayorMetrics.NoteSales, "Servicios NOTA DE CREDITO D4");
    assertClose(cellNumber(notaSheet, 4, 5), mayorMetrics.NoteDiscounts, "Servicios NOTA DE CREDITO E4");
    assertClose(cellNumber(notaSheet, 4, 6), mayorMetrics.NoteSales, "Servicios NOTA DE CREDITO F4");
    assertClose(cellNumber(notaSheet, 4, 7), mayorMetrics.NoteDiscounts, "Servicios NOTA DE CREDITO G4");
    assertClose(cellNumber(repVtasSheet, 6, 4), mayorMetrics.NetSales, "Servicios REP VTAS D6");
    assertClose(cellNumber(repVtasSheet, 6, 5), cellNumber(costoSheet, 4, 10), "Servicios REP VTAS E6");

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
    console.log("OK: contrato E2E Servicios por Marca validado en Laravel + Python.");
  } finally {
    server.kill();
  }
}

main().catch((error) => {
  console.error(`FAIL: ${error.message}`);
  process.exit(1);
});
