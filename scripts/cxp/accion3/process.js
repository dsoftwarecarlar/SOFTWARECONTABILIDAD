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
} = require("./constants");
const {
  parseInputSources,
  validateRows,
} = require("./parser");
const {
  buildWorkbookFromTemplate,
  preserveTemplateVisualWorkbook,
  verifyOutputWorkbook,
} = require("./workbook");

function parseCliArguments(argv = process.argv.slice(2)) {
  const hasFlags = argv.includes("--output") || argv.includes("--template");
  if (!hasFlags) {
    return {
      inputSources: [argv[0] || DEFAULT_INPUT_SOURCE],
      outputXlsx: argv[1] || DEFAULT_OUTPUT_XLSX,
      templateXlsx: argv[2] ? path.resolve(process.cwd(), argv[2]) : DEFAULT_TEMPLATE_XLSX,
    };
  }

  const inputSources = [];
  let outputXlsx = DEFAULT_OUTPUT_XLSX;
  let templateXlsx = DEFAULT_TEMPLATE_XLSX;

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--output") {
      if (!argv[index + 1]) {
        throw new Error("Falta el valor de --output.");
      }
      outputXlsx = argv[index + 1];
      index += 1;
      continue;
    }
    if (arg === "--template") {
      if (!argv[index + 1]) {
        throw new Error("Falta el valor de --template.");
      }
      templateXlsx = path.resolve(process.cwd(), argv[index + 1]);
      index += 1;
      continue;
    }
    inputSources.push(arg);
  }

  if (inputSources.length === 0) {
    inputSources.push(DEFAULT_INPUT_SOURCE);
  }

  return { inputSources, outputXlsx, templateXlsx };
}

async function main() {
  const processStart = Date.now();
  const cli = parseCliArguments();
  const outputPath = path.resolve(process.cwd(), cli.outputXlsx);
  const templatePath = cli.templateXlsx;
  const timingsMs = {};

  let stageStart = Date.now();
  const parsed = await parseInputSources(cli.inputSources);
  const rows = parsed.rows;
  timingsMs.parse = Date.now() - stageStart;

  stageStart = Date.now();
  validateRows(rows);
  timingsMs.validate = Date.now() - stageStart;

  stageStart = Date.now();
  const { workbook, summary } = await buildWorkbookFromTemplate(templatePath, rows);
  timingsMs.build = Date.now() - stageStart;

  stageStart = Date.now();
  const finalOutputPath = await writeWorkbookWithRetries(workbook, outputPath);
  timingsMs.write = Date.now() - stageStart;

  stageStart = Date.now();
  const mergeMetrics = preserveTemplateVisualWorkbook(templatePath, finalOutputPath, SHEET_NAME);
  timingsMs.merge = Date.now() - stageStart;

  stageStart = Date.now();
  await verifyOutputWorkbook(finalOutputPath, templatePath, rows.length);
  timingsMs.verify = Date.now() - stageStart;
  timingsMs.total = Date.now() - processStart;

  const memory = process.memoryUsage();

  const auditPath = path.join(
    path.dirname(finalOutputPath),
    `${path.parse(finalOutputPath).name}_auditoria.json`,
  );

  writeAuditReport(auditPath, {
    fecha_proceso: new Date().toISOString(),
    input_source: parsed.inputPaths.length === 1 ? parsed.inputPaths[0] : parsed.inputPaths,
    input_sources: parsed.inputPaths,
    input_pdf: parsed.inputExt === ".pdf" ? parsed.inputPaths[0] : null,
    input_tipo: parsed.diagnostics.source_type || parsed.inputExt.replace(".", ""),
    archivos_origen: parsed.sourceFiles,
    total_archivos_origen: parsed.inputPaths.length,
    output_xlsx: finalOutputPath,
    hoja_salida: SHEET_NAME,
    movimientos_extraidos: rows.length,
    filas_fecha_detectadas: parsed.diagnostics.date_rows_detected,
    filas_fecha_omitidas: parsed.diagnostics.skipped_date_rows,
    cuentas_validadas_fuente: parsed.diagnostics.account_totals_checked,
    cuentas_descuadradas_fuente: parsed.diagnostics.account_total_mismatches,
    cuentas_validadas_pdf: parsed.diagnostics.account_totals_checked,
    cuentas_descuadradas_pdf: parsed.diagnostics.account_total_mismatches,
    resumen_filtrado: summary.length,
    rows_before_merge: mergeMetrics.rows_before_merge,
    rows_after_merge: mergeMetrics.rows_after_merge,
    payload_cells_before: mergeMetrics.payload_cells_before,
    payload_cells_after: mergeMetrics.payload_cells_after,
    payload_hash_before: mergeMetrics.payload_hash_before,
    payload_hash_after: mergeMetrics.payload_hash_after,
    merge_integrity_ok: mergeMetrics.merge_integrity_ok === true,
    timings_ms: timingsMs,
    memory_bytes: {
      rss: memory.rss,
      heap_total: memory.heapTotal,
      heap_used: memory.heapUsed,
    },
    verificacion_final_ok: true,
  });

  if (parsed.inputPaths.length === 1) {
    console.log(`Archivo leido: ${parsed.inputPaths[0]}`);
  } else {
    console.log(`Archivos leidos: ${parsed.inputPaths.length}`);
    for (const inputPath of parsed.inputPaths) {
      console.log(`- ${inputPath}`);
    }
  }
  console.log(`Movimientos extraidos: ${rows.length}`);
  console.log(`Resumen lateral: ${summary.length} filas`);
  console.log(`Rendimiento (ms): parse=${timingsMs.parse}, validate=${timingsMs.validate}, build=${timingsMs.build}, write=${timingsMs.write}, merge=${timingsMs.merge}, verify=${timingsMs.verify}, total=${timingsMs.total}`);
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
