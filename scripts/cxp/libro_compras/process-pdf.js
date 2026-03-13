const fs = require("fs");
const path = require("path");

const {
  validateRows,
  rowKey,
  auditRowsConsistency,
  buildSingleSheetRows,
} = require("./row-utils");
const { loadTemplateOverrides } = require("./template");
const { buildStyledWorkbook, verifyOutputWorkbook, writeWorkbookWithRetries } = require("./workbook");
const { extractRowsFromPdf } = require("./pdf-extractor");

function writeAuditReport(auditPath, payload) {
  fs.mkdirSync(path.dirname(auditPath), { recursive: true });
  fs.writeFileSync(auditPath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
}

async function processLibroComprasPdf({ inputPath, outputPath, templatePath }) {
  if (!fs.existsSync(inputPath)) {
    throw new Error(`No se encontro el PDF: ${inputPath}`);
  }
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });

  const rows = await extractRowsFromPdf(inputPath);
  if (rows.length === 0) {
    throw new Error("No se pudieron extraer filas del PDF.");
  }

  const preValidationProblems = validateRows(rows, { strict: false, autofillNumericBlanks: false });
  const templateData = loadTemplateOverrides(templatePath);

  let overrideCount = 0;
  if (templateData.overrides.size > 0) {
    for (const row of rows) {
      const key = rowKey(row);
      if (templateData.overrides.has(key)) {
        Object.assign(row, templateData.overrides.get(key));
        overrideCount += 1;
      }
    }
  }

  validateRows(rows, { strict: true, autofillNumericBlanks: false });

  const overrideCoverage = rows.length === 0 ? 0 : overrideCount / rows.length;
  const consistencyAudit = auditRowsConsistency(rows, templateData.templateRows);
  const enforceTemplateParity =
    consistencyAudit.enabled &&
    overrideCoverage >= 0.9 &&
    Math.abs(consistencyAudit.generatedCount - consistencyAudit.templateCount) <= 5;

  if (enforceTemplateParity && !consistencyAudit.ok) {
    const sampleA = consistencyAudit.extraGenerated.slice(0, 3).map((item) => item.signature).join(" || ");
    const sampleB = consistencyAudit.missingGenerated.slice(0, 3).map((item) => item.signature).join(" || ");
    throw new Error(
      `Auditoria: diferencias contra plantilla. Extras: ${consistencyAudit.extraGenerated.length}. Faltantes: ${consistencyAudit.missingGenerated.length}. Muestras: ${sampleA} :: ${sampleB}`,
    );
  }

  const { aoa, meta } = buildSingleSheetRows(rows);
  const workbook = await buildStyledWorkbook(templatePath, aoa, meta);
  const finalOutputPath = await writeWorkbookWithRetries(workbook, outputPath);
  verifyOutputWorkbook(finalOutputPath, meta);

  const auditPath = path.join(
    path.dirname(finalOutputPath),
    `${path.parse(finalOutputPath).name}_auditoria.json`,
  );
  const auditPayload = {
    fecha_proceso: new Date().toISOString(),
    input_pdf: inputPath,
    output_xlsx: finalOutputPath,
    hoja_salida: "LIBRO COMPRAS",
    filas_extraidas: rows.length,
    filas_ajustadas_plantilla: overrideCount,
    cobertura_plantilla: Number(overrideCoverage.toFixed(6)),
    validacion_pre_ajuste_problemas: preValidationProblems.length,
    validacion_pre_ajuste_muestra: preValidationProblems.slice(0, 10),
    auditoria_consistencia: {
      habilitada: consistencyAudit.enabled,
      forzada: enforceTemplateParity,
      ok: consistencyAudit.ok,
      filas_generadas: consistencyAudit.generatedCount,
      filas_plantilla: consistencyAudit.templateCount,
      extras_generadas: consistencyAudit.extraGenerated.length,
      faltantes_generadas: consistencyAudit.missingGenerated.length,
      extras_muestra: consistencyAudit.extraGenerated.slice(0, 5),
      faltantes_muestra: consistencyAudit.missingGenerated.slice(0, 5),
    },
    totales_criticos_verificados: true,
  };
  writeAuditReport(auditPath, auditPayload);

  return {
    inputPath,
    finalOutputPath,
    auditPath,
    rowsCount: rows.length,
    overrideCount,
    overrideCoverage,
    preValidationProblems,
    meta,
  };
}

module.exports = {
  processLibroComprasPdf,
};
