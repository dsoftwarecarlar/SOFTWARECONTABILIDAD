const http = require("http");
const { spawn } = require("child_process");

const ROOT = "C:/xampp/htdocs/SOFTWARECONTABILIDAD";
const PHP_EXE = process.env.PHP_EXE || "C:/xampp/php/php.exe";
const HOST = "127.0.0.1";
const PORT = 18996;

const PAGES = [
  "/",
  "/areas/cxp/index.php",
  "/areas/cxp/libro-compras-aclt.php",
  "/areas/cxp/conciliacion-servicios-marcas.php",
  "/areas/cxp/facturacion-repuestos-tytserv.php",
  "/modules/cxp_txt/index.php",
  "/modules/cxp_accion3/index.php",
  "/modules/cxp_repuestos_tytserv/index.php",
  "/modules/cxp_servicios_marcas/index.php",
];

function assertCondition(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

function request(pathname) {
  return new Promise((resolve, reject) => {
    const req = http.request(
      {
        hostname: HOST,
        port: PORT,
        path: pathname,
        method: "GET",
      },
      (res) => {
        const chunks = [];
        res.on("data", (chunk) => chunks.push(chunk));
        res.on("end", () => {
          const buffer = Buffer.concat(chunks);
          resolve({
            status: res.statusCode || 0,
            text: buffer.toString("utf8"),
          });
        });
      },
    );

    req.on("error", reject);
    req.end();
  });
}

async function waitServerReady(maxAttempts = 60) {
  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      const response = await request("/");
      if (response.status === 200) {
        return;
      }
    } catch (_error) {
      // retry
    }
    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  throw new Error("No se pudo iniciar el servidor PHP para smoke de recursos.");
}

function extractAssetUrls(html) {
  const urls = new Set();
  const regex = /(?:href|src)="([^"]+)"/gi;
  let match = regex.exec(html);

  while (match) {
    const raw = String(match[1] || "");
    if (
      raw === ""
      || raw.startsWith("#")
      || raw.startsWith("http://")
      || raw.startsWith("https://")
      || raw.startsWith("mailto:")
      || raw.includes("download.php?file=")
    ) {
      match = regex.exec(html);
      continue;
    }

    const normalized = raw.startsWith("/") ? raw : `/${raw}`;
    if (
      normalized.endsWith(".xlsx")
      || normalized.endsWith(".xls")
      || normalized.endsWith(".txt")
      || normalized.endsWith(".pdf")
    ) {
      match = regex.exec(html);
      continue;
    }

    if (
      normalized.startsWith("/assets/")
      || normalized.startsWith("/images/")
      || normalized.startsWith("/modules/")
    ) {
      urls.add(normalized);
    }

    match = regex.exec(html);
  }

  return [...urls];
}

async function runSmoke() {
  for (const page of PAGES) {
    const pageResponse = await request(page);
    assertCondition(pageResponse.status === 200, `Vista no disponible (${page}) -> ${pageResponse.status}`);

    const assets = extractAssetUrls(pageResponse.text);
    for (const asset of assets) {
      const assetResponse = await request(asset);
      assertCondition(assetResponse.status === 200, `Recurso no disponible (${asset}) -> ${assetResponse.status}`);
    }
  }
}

async function main() {
  const server = spawn(PHP_EXE, ["-S", `${HOST}:${PORT}`, "-t", ROOT], {
    cwd: ROOT,
    stdio: ["ignore", "pipe", "pipe"],
  });

  server.stdout.on("data", (chunk) => process.stdout.write(`[php] ${chunk}`));
  server.stderr.on("data", (chunk) => process.stderr.write(`[php] ${chunk}`));

  try {
    await waitServerReady();
    await runSmoke();
    console.log("OK: smoke de vistas y recursos.");
  } finally {
    server.kill();
  }
}

main().catch((error) => {
  console.error(`FAIL: ${error.message}`);
  process.exit(1);
});
