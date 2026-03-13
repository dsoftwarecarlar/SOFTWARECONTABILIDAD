const path = require("path");

const { processLibroComprasPdf } = require("./scripts/cxp/libro_compras/process-pdf");

const INPUT_PDF = process.argv[2] || "CXPREP_docproveedor.pdf";
const OUTPUT_XLSX = process.argv[3] || "clasificacion_mantenimiento.xlsx";
const TEMPLATE_XLSX = "EJEMPLODECOMOQUEDARIA.xlsx";

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
