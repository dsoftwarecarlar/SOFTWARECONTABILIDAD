const http = require("http");
const fs = require("fs");
const path = require("path");
const { spawn } = require("child_process");

const ROOT = "C:/xampp/htdocs/SOFTWARECONTABILIDAD";
const APP_ROOT = path.join(ROOT, "laravel_app");
const ROUTER = path.join(APP_ROOT, "server.php");
const PHP_EXE = firstExistingPath([
  path.join(ROOT, ".tools", "php82", "php.exe"),
  process.env.PHP_EXE || "",
  "C:/xampp/php/php.exe",
]);
const HOST = "127.0.0.1";
const PORT = 18996;
const ORIGIN = `http://${HOST}:${PORT}`;
const BANNED_VISIBLE_TERMS = [
  { label: "laravel", pattern: /\blaravel\b/i },
  { label: "python", pattern: /\bpython\b/i },
  { label: "php", pattern: /\bphp\b/i },
  { label: "node", pattern: /\bnode\b/i },
  { label: "powershell", pattern: /\bpowershell\b/i },
  { label: "legacy", pattern: /\blegacy\b/i },
  { label: "worker", pattern: /\bworker\b/i },
  { label: "pipeline", pattern: /\bpipeline\b/i },
  { label: "runtime", pattern: /\bruntime\b/i },
  { label: "preflight", pattern: /\bpreflight\b/i },
  { label: "dispatch", pattern: /\bdispatch\b/i },
  { label: "job", pattern: /\bjob\b/i },
  { label: "queued", pattern: /\bqueued\b/i },
  { label: "running", pattern: /\brunning\b/i },
  { label: "excel com", pattern: /\bexcel com\b/i },
];

const PAGES = [
  "/",
  "/cxp",
  "/contabilidad-general",
  "/cxp/windows/libro-compras-aclt",
  "/cxp/windows/conciliacion-servicios-marcas",
  "/cxp/windows/facturacion-repuestos-tytserv",
  "/cxp/modules/accion1",
  "/cxp/modules/accion2",
  "/cxp/modules/accion3",
  "/cxp/modules/accion4",
  "/cxp/modules/consolidado-acciones",
  "/cxp/modules/servicios-marcas",
  "/cxp/modules/repuestos-tytserv",
];
const REQUIRED_VISIBLE_HINTS = {
  "/cxp/modules/accion1": ["cxprep_docproveedor"],
  "/cxp/modules/accion2": ["cxprep_ret_generalaccion2"],
  "/cxp/modules/accion3": ["con_mayorgen2accion3"],
  "/cxp/modules/accion4": ["con_mayorgen2ivaaccion4"],
  "/cxp/modules/servicios-marcas": ["detalle-vtas-xliquidar", "repfacturacionservcontabilidad", "serrep_facturas_nafchan"],
  "/cxp/modules/repuestos-tytserv": ["replibroventasgeneral", "replibrodevolucionesgeneral"],
};

function firstExistingPath(candidates) {
  for (const candidate of candidates) {
    if (!candidate) {
      continue;
    }
    if ((candidate.includes("\\") || candidate.includes("/")) && fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return candidates[0];
}

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
          resolve({
            status: res.statusCode || 0,
            text: Buffer.concat(chunks).toString("utf8"),
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

  throw new Error("No se pudo iniciar el servidor Laravel para smoke de recursos.");
}

async function runSmoke() {
  const seenUrls = new Set();

  for (const page of PAGES) {
    const response = await request(page);
    assertCondition(response.status === 200, `Vista no disponible (${page}) -> ${response.status}`);
    assertCondition(
      /<html[\s>]/i.test(response.text) || /<!doctype html>/i.test(response.text),
      `La respuesta de ${page} no parece HTML valido.`,
    );

    const visibleText = extractVisibleText(response.text);
    assertNoTechnicalText(page, visibleText);
    assertVisibleHints(page, visibleText);

    for (const internalUrl of collectInternalUrls(page, response.text)) {
      if (seenUrls.has(internalUrl)) {
        continue;
      }

      seenUrls.add(internalUrl);
      const linkedResponse = await request(internalUrl);
      assertCondition(
        linkedResponse.status === 200,
        `Recurso o enlace interno no disponible (${internalUrl}) -> ${linkedResponse.status}`,
      );
    }
  }
}

function extractVisibleText(html) {
  return html
    .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, " ")
    .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, " ")
    .replace(/<!--[\s\S]*?-->/g, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, "\"")
    .replace(/&#39;/gi, "'")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function assertNoTechnicalText(page, visibleText) {
  for (const term of BANNED_VISIBLE_TERMS) {
    assertCondition(
      !term.pattern.test(visibleText),
      `Texto tecnico visible para usuario en ${page}: "${term.label}"`,
    );
  }
}

function assertVisibleHints(page, visibleText) {
  for (const hint of REQUIRED_VISIBLE_HINTS[page] || []) {
    assertCondition(
      visibleText.includes(hint),
      `No se mostro la referencia esperada de archivo en ${page}: ${hint}`,
    );
  }
}

function collectInternalUrls(page, html) {
  const urls = new Set();
  const attributePattern = /\b(?:href|src)=["']([^"']+)["']/gi;

  for (const match of html.matchAll(attributePattern)) {
    const rawValue = (match[1] || "").trim();
    if (!rawValue) {
      continue;
    }

    const lowered = rawValue.toLowerCase();
    if (
      lowered.startsWith("#") ||
      lowered.startsWith("javascript:") ||
      lowered.startsWith("mailto:") ||
      lowered.startsWith("tel:") ||
      lowered.startsWith("data:")
    ) {
      continue;
    }

    const resolved = new URL(rawValue, `${ORIGIN}${page}`);
    const isLegacyPhp = /\.php(?:[?#].*)?$/i.test(resolved.pathname + resolved.search);
    const isLegacyAreaPath = /^\/areas\//i.test(resolved.pathname);
    const isLegacyModulePath = /^\/modules\//i.test(resolved.pathname);

    assertCondition(
      !isLegacyPhp && !isLegacyAreaPath && !isLegacyModulePath,
      `Enlace legacy visible en ${page}: ${rawValue}`,
    );

    if (resolved.origin !== ORIGIN) {
      continue;
    }

    urls.add(`${resolved.pathname}${resolved.search}`);
  }

  return urls;
}

async function main() {
  const server = spawn(PHP_EXE, ["-d", "max_execution_time=180", "-S", `${HOST}:${PORT}`, ROUTER], {
    cwd: APP_ROOT,
    stdio: ["ignore", "pipe", "pipe"],
  });

  server.stdout.on("data", (chunk) => process.stdout.write(`[php] ${chunk}`));
  server.stderr.on("data", (chunk) => process.stderr.write(`[php] ${chunk}`));

  try {
    await waitServerReady();
    await runSmoke();
    console.log("OK: smoke de vistas y recursos Laravel.");
  } finally {
    server.kill();
  }
}

main().catch((error) => {
  console.error(`FAIL: ${error.message}`);
  process.exit(1);
});
