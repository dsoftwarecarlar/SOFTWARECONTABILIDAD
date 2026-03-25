const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..", "..");
const JOBS_DIR = path.join(ROOT, "storage", "jobs");

const TOTAL_BRAND_MS_MAX = 120000;
const PRECONT_VENTAS_MS_MAX = 10000;
const PRECONT_COSTOS2_MS_MAX = 2000;

function assertCondition(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

function latestServiciosJob() {
  const files = fs.readdirSync(JOBS_DIR)
    .filter((name) => /^servicios_marcas_.*\.json$/i.test(name))
    .map((name) => {
      const filePath = path.join(JOBS_DIR, name);
      return {
        name,
        filePath,
        stat: fs.statSync(filePath),
      };
    })
    .sort((left, right) => right.stat.mtimeMs - left.stat.mtimeMs);

  assertCondition(files.length > 0, "No existe ningun job de Servicios por Marca para medir rendimiento.");
  const raw = fs.readFileSync(files[0].filePath, "utf8");
  return {
    name: files[0].name,
    payload: JSON.parse(raw),
  };
}

function metric(consoleText, label) {
  const match = consoleText.match(new RegExp(`INFO\\|${label}\\|[^|]+\\|(\\d+)`));
  return match ? Number(match[1]) : null;
}

function main() {
  const { name, payload } = latestServiciosJob();
  const consoleText = String(payload.console || "");

  const totalBrandMs = metric(consoleText, "total_brand_ms");
  const fillPrecontVentasMs = metric(consoleText, "fill_precont_ventas_ms");
  const fillPrecontCostos2Ms = metric(consoleText, "fill_precont_costos2_ms");

  assertCondition(payload.status === "complete", `El ultimo job de servicios no termino en complete (${payload.status || "desconocido"}).`);
  assertCondition(totalBrandMs !== null, "No se encontro total_brand_ms en el ultimo job de servicios.");
  assertCondition(fillPrecontVentasMs !== null, "No se encontro fill_precont_ventas_ms en el ultimo job de servicios.");
  assertCondition(fillPrecontCostos2Ms !== null, "No se encontro fill_precont_costos2_ms en el ultimo job de servicios.");

  assertCondition(
    totalBrandMs <= TOTAL_BRAND_MS_MAX,
    `Servicios por Marca excedio el guardrail total (${totalBrandMs} ms > ${TOTAL_BRAND_MS_MAX} ms).`,
  );
  assertCondition(
    fillPrecontVentasMs <= PRECONT_VENTAS_MS_MAX,
    `Servicios por Marca excedio el guardrail de PrecontabilizacionVentas (${fillPrecontVentasMs} ms > ${PRECONT_VENTAS_MS_MAX} ms).`,
  );
  assertCondition(
    fillPrecontCostos2Ms <= PRECONT_COSTOS2_MS_MAX,
    `Servicios por Marca excedio el guardrail de PrecontabilizacionCostos (2) (${fillPrecontCostos2Ms} ms > ${PRECONT_COSTOS2_MS_MAX} ms).`,
  );

  console.log(
    `OK: rendimiento Servicios por Marca dentro de guardrail (${name}) total=${totalBrandMs}ms precontVentas=${fillPrecontVentasMs}ms precontCostos2=${fillPrecontCostos2Ms}ms.`,
  );
}

main();
