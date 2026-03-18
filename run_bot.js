const fs = require("fs");
const path = require("path");

const { processLibroComprasPdf } = require("./scripts/cxp/libro_compras/process-pdf");

function firstExistingPath(candidates) {
  for (const candidate of candidates) {
    if (candidate && fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return candidates[0];
}

const DEFAULT_INPUT_PDF = firstExistingPath([
  path.resolve(process.cwd(), "resources", "cxp", "acciones", "fixtures", "CXPREP_docproveedor.pdf"),
  path.resolve(process.cwd(), "CXPREP_docproveedor.pdf"),
]);
const DEFAULT_OUTPUT_XLSX = path.resolve(process.cwd(), "outputs", "clasificacion_mantenimiento_final.xlsx");
const DEFAULT_TEMPLATE_XLSX = firstExistingPath([
  path.resolve(process.cwd(), "resources", "cxp", "acciones", "templates", "EJEMPLODECOMOQUEDARIA.xlsx"),
  path.resolve(process.cwd(), "EJEMPLODECOMOQUEDARIA.xlsx"),
]);

const INPUT_PDF = process.argv[2] || DEFAULT_INPUT_PDF;
const OUTPUT_XLSX = process.argv[3] || DEFAULT_OUTPUT_XLSX;
const TEMPLATE_XLSX = DEFAULT_TEMPLATE_XLSX;

async function main() {
  const inputPath = path.resolve(process.cwd(), INPUT_PDF);
  const outputPath = path.resolve(process.cwd(), OUTPUT_XLSX);
  const templatePath = path.resolve(process.cwd(), TEMPLATE_XLSX);

  const result = await processLibroComprasPdf({
    inputPath,
    outputPath,
    templatePath,
  });

  console.log(`PDF leido: ${result.inputPath}`);
  console.log(`Filas extraidas: ${result.rowsCount}`);
  console.log(`Filas ajustadas con plantilla: ${result.overrideCount}`);
  console.log(`Cobertura plantilla: ${(result.overrideCoverage * 100).toFixed(2)}%`);
  console.log(`Problemas detectados antes de ajuste: ${result.preValidationProblems.length}`);
  console.log(`Bloque principal: ${result.meta.counts.main}`);
  console.log(`RIMPE: ${result.meta.counts.rimpe}`);
  console.log(`NDS/TR/ANULACIONES: ${result.meta.counts.ndtr}`);
  console.log(`Excel generado (una sola hoja): ${result.finalOutputPath}`);
  console.log(`Auditoria JSON: ${result.auditPath}`);
}

main().catch((error) => {
  console.error("Error:", error.message);
  process.exit(1);
});
