const fs = require("fs");
const path = require("path");

const {
  writeWorkbookWithRetries,
  writeAuditReport,
} = require("../shared/core-utils");
const {
  DEFAULT_INPUT_SOURCE,
  DEFAULT_OUTPUT_XLSX,
  DEFAULT_TEMPLATE_XLSX,
  SHEET_NAME,
  EXPECTED_COLUMNS,
} = require("./constants");
const {
  extractRowsFromTxt,
  loadTemplateTipoHints,
  normalizeParsedRows,
} = require("./parser");
const {
  buildWorkbookFromTemplate,
  preserveTemplateVisualWorkbook,
  verifyOutputWorkbook,
} = require("./workbook");

async function main() {
  const inputPath = path.resolve(process.cwd(), process.argv[2] || DEFAULT_INPUT_SOURCE);
  const outputPath = path.resolve(process.cwd(), process.argv[3] || DEFAULT_OUTPUT_XLSX);
  const templatePath = process.argv[4]
    ? path.resolve(process.cwd(), process.argv[4])
    : DEFAULT_TEMPLATE_XLSX;

  if (!fs.existsSync(inputPath)) {
    throw new Error(`No se encontro el TXT: ${inputPath}`);
  }

  const hints = loadTemplateTipoHints(templatePath);
  const rawRows = extractRowsFromTxt(inputPath);
  const rows = normalizeParsedRows(rawRows, hints);
  const { workbook, summary } = await buildWorkbookFromTemplate(templatePath, rows);
  const finalOutputPath = await writeWorkbookWithRetries(workbook, outputPath);
  preserveTemplateVisualWorkbook(templatePath, finalOutputPath, SHEET_NAME);
  await verifyOutputWorkbook(finalOutputPath, templatePath, rows.length, summary);

  const auditPath = path.join(
    path.dirname(finalOutputPath),
    `${path.parse(finalOutputPath).name}_auditoria.json`,
  );

  const auditPayload = {
    fecha_proceso: new Date().toISOString(),
    input_txt: inputPath,
    output_xlsx: finalOutputPath,
    hoja_salida: SHEET_NAME,
    filas_txt: rows.length,
    columnas_esperadas: EXPECTED_COLUMNS,
    resumen_generado_filas: summary.length,
    verificacion_final_ok: true,
  };
  writeAuditReport(auditPath, auditPayload);

  console.log(`TXT leido: ${inputPath}`);
  console.log(`Filas parseadas: ${rows.length}`);
  console.log(`Resumen lateral: ${summary.length} filas`);
  console.log(`Excel generado (una sola hoja): ${finalOutputPath}`);
  console.log(`Auditoria JSON: ${auditPath}`);
}

module.exports = { main };

if (require.main === module) {
  main().catch((error) => {
    console.error("Error:", error.message);
    process.exit(1);
  });
}
