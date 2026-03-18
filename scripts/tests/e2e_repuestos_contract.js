const fs = require("fs");
const http = require("http");
const path = require("path");
const crypto = require("crypto");
const { spawn } = require("child_process");

const XLSX = require("xlsx");

const ROOT = path.resolve(__dirname, "..", "..");
const PHP_EXE = process.env.PHP_EXE || "C:\\xampp\\php\\php.exe";
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
    if (candidate && fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return candidates[0];
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

  const successMatch = html.match(/Archivo generado correctamente:\s*([^<\r\n]+)/i);
  assertCondition(successMatch, "No se encontro mensaje de exito con excel_name.");
  const successName = successMatch[1].trim();

  return {
    link: parsed.pathname + parsed.search,
    fileName,
    successName,
  };
}

function cellText(sheet, address) {
  const cell = sheet[address];
  if (!cell || cell.v == null) {
    return "";
  }

  return String(cell.v).trim();
}

function cellTextAt(sheet, row, column) {
  return cellText(sheet, XLSX.utils.encode_cell({ r: row - 1, c: column - 1 }));
}

function rowContainsTotalGeneral(sheet, row, lastColumn = 41) {
  for (let column = 1; column <= lastColumn; column += 1) {
    if (/^TOTAL\s+GENERAL/i.test(cellTextAt(sheet, row, column))) {
      return true;
    }
  }

  return false;
}

function sheetSignature(sheet) {
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

    const document = cellTextAt(sheet, row, 5);
    if (document === "") {
      continue;
    }

    rows.push(payloadColumns.map((column) => cellTextAt(sheet, row, column)).join("|"));
  }

  return {
    rows: rows.length,
    hash: crypto.createHash("sha256").update(rows.join("\n")).digest("hex"),
  };
}

function buildContractSources() {
  const sourceDir = firstExistingPath([
    path.join(ROOT, "resources", "cxp", "repuestos_tytserv", "fixtures"),
    path.join(ROOT, "outputs", "EJEMPLOAMANOTAREA3"),
  ]);
  const files = [
    { field: "excel_tyt", path: path.join(sourceDir, "RepLibroVentasGeneral.xlsx"), sheet: "REP TYT" },
    { field: "excel_peug", path: path.join(sourceDir, "RepLibroVentasGeneral (1).xlsx"), sheet: "REP PEUGT" },
    { field: "excel_chgn", path: path.join(sourceDir, "RepLibroVentasGeneral (2).xlsx"), sheet: "REP CHGN" },
    { field: "excel_szk", path: path.join(sourceDir, "RepLibroVentasGeneral (3).xlsx"), sheet: "REP SZK" },
  ];

  for (const file of files) {
    assertCondition(fs.existsSync(file.path), `No existe archivo de contrato: ${file.path}`);
  }

  return files;
}

async function runContract() {
  const files = buildContractSources();
  const payload = buildMultipartBody(files);
  const response = await httpRequest("POST", "/modules/cxp_repuestos_tytserv/index.php", payload.body, {
    "Content-Type": `multipart/form-data; boundary=${payload.boundary}`,
    "Content-Length": String(payload.body.length),
  });

  assertCondition(response.status === 200, `Repuestos devolvio estado HTTP ${response.status}.`);
  const parsed = extractDownloadNameAndSuccess(response.text);
  assertCondition(
    parsed.successName === parsed.fileName,
    `Repuestos inconsistente: success_message='${parsed.successName}' y descarga='${parsed.fileName}'.`,
  );

  const download = await httpRequest("GET", parsed.link);
  assertCondition(download.status === 200, `No se pudo descargar salida de repuestos (${download.status}).`);

  const outputWorkbook = XLSX.read(download.buffer, { type: "buffer", cellFormula: false });
  for (const sourceFile of files) {
    const sourceWorkbook = XLSX.readFile(sourceFile.path, { cellFormula: false });
    const sourceSheet = sourceWorkbook.Sheets[sourceWorkbook.SheetNames[0]];
    const outputSheet = outputWorkbook.Sheets[sourceFile.sheet];

    assertCondition(!!outputSheet, `No existe hoja ${sourceFile.sheet} en el Excel descargado.`);

    const sourceSignature = sheetSignature(sourceSheet);
    const outputSignature = sheetSignature(outputSheet);
    assertCondition(
      sourceSignature.rows === outputSignature.rows,
      `El payload de ${sourceFile.sheet} no coincide con el archivo subido (srcRows=${sourceSignature.rows}, outRows=${outputSignature.rows}).`,
    );
  }
}

async function main() {
  const server = startPhpServer();
  try {
    await waitServerReady();
    await runContract();
    console.log("OK: contrato E2E repuestos validado.");
  } finally {
    server.kill();
  }
}

main().catch((error) => {
  console.error(`FAIL: ${error.message}`);
  process.exit(1);
});
