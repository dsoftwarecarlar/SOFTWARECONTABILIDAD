const fs = require("fs");
const crypto = require("crypto");

const AdmZip = require("adm-zip");
const ExcelJS = require("exceljs");
const XLSX = require("xlsx");

const {
  sanitizeText,
  deepClone,
  round2,
  toExcelDateSerial,
} = require("../shared/core-utils");
const {
  clearRangeValues,
  parseXml,
  readZipText,
  updateZipText,
  getWorksheetEntryPath,
  getRowsArray,
  setRowsArray,
  getCellsArray,
  setCellsArray,
  parseCellRef,
  cellRef,
  copyCellPayload,
  cloneTemplateRow,
  clearCellPayload,
  hasCellPayload,
  stripCalcChain,
  styleSignature,
} = require("../shared/excel-template-utils");
const {
  SHEET_NAME,
  EXPECTED_HEADERS,
} = require("./constants");

function ensureTemplateCapacity(ws, requiredRowCount) {
  const currentRowCount = Math.max(ws.rowCount, 1);
  if (requiredRowCount <= currentRowCount) {
    return;
  }
  ws.duplicateRow(Math.max(currentRowCount, 2), requiredRowCount - currentRowCount, true);
}

function readTemplateSummaryLabels(ws) {
  const labels = [];
  for (let row = 2; row <= 120; row += 1) {
    const label = sanitizeText(ws.getCell(row, 13).value);
    if (!label) {
      if (labels.length > 0) {
        break;
      }
      continue;
    }
    labels.push({ row, label });
    if (label.toUpperCase() === "TOTAL GENERAL") {
      break;
    }
  }
  return labels;
}

function toCents(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return 0;
  }
  return Math.round((numeric + Number.EPSILON) * 100);
}

function fromCents(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return 0;
  }
  return numeric / 100;
}

function createStyleSignatureCache() {
  const cache = new WeakMap();
  return (style) => {
    if (!style || typeof style !== "object") {
      return styleSignature(style);
    }

    const cached = cache.get(style);
    if (cached) {
      return cached;
    }

    const computed = styleSignature(style);
    cache.set(style, computed);
    return computed;
  };
}

function buildSummary(rows, templateLabels) {
  const bucket = new Map();
  let totalDebeCents = 0;
  let totalHaberCents = 0;

  for (const row of rows) {
    const label = sanitizeText(row.CUENTA);
    if (!bucket.has(label)) {
      bucket.set(label, { debeCents: 0, haberCents: 0 });
    }
    const item = bucket.get(label);
    const debeCents = toCents(row.DEBE);
    const haberCents = toCents(row.HABER);
    item.debeCents += debeCents;
    item.haberCents += haberCents;
    totalDebeCents += debeCents;
    totalHaberCents += haberCents;
  }

  const result = [];
  for (const labelItem of templateLabels) {
    if (labelItem.label.toUpperCase() === "TOTAL GENERAL") {
      result.push({
        row: labelItem.row,
        label: labelItem.label,
        debe: fromCents(totalDebeCents),
        haber: fromCents(totalHaberCents),
      });
      continue;
    }

    const values = bucket.get(labelItem.label) || { debeCents: 0, haberCents: 0 };
    result.push({
      row: labelItem.row,
      label: labelItem.label,
      debe: fromCents(values.debeCents),
      haber: fromCents(values.haberCents),
    });
  }

  return result;
}

function cellPayloadSignature(cell) {
  return JSON.stringify({
    t: cell?.["@_t"] ?? null,
    v: cell?.v ?? null,
    f: cell?.f ?? null,
    is: cell?.is ?? null,
  });
}

function capturePayloadSnapshot(rows, maxColumn) {
  const rowsWithPayload = new Set();
  const payloadEntries = [];

  for (const row of rows) {
    const rowNumber = Number(row?.["@_r"] || 0);
    if (rowNumber < 2) {
      continue;
    }

    const cells = getCellsArray(row);
    for (const cell of cells) {
      const ref = parseCellRef(cell?.["@_r"]);
      if (!ref || ref.col < 1 || ref.col > maxColumn) {
        continue;
      }
      if (!hasCellPayload(cell)) {
        continue;
      }

      rowsWithPayload.add(ref.row);
      payloadEntries.push(`${ref.row}:${ref.col}:${cellPayloadSignature(cell)}`);
    }
  }

  payloadEntries.sort();
  const payloadHash = crypto
    .createHash("sha256")
    .update(payloadEntries.join("\n"), "utf8")
    .digest("hex");

  return {
    rows: rowsWithPayload.size,
    payloadCells: payloadEntries.length,
    payloadHash,
  };
}

async function buildWorkbookFromTemplate(templatePath, rows) {
  if (!fs.existsSync(templatePath)) {
    throw new Error(`No se encontro plantilla de Accion 3: ${templatePath}`);
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(templatePath);
  const ws = wb.getWorksheet(SHEET_NAME);
  if (!ws) {
    throw new Error(`La plantilla de Accion 3 no contiene la hoja ${SHEET_NAME}.`);
  }

  ensureTemplateCapacity(ws, Math.max(rows.length + 1, ws.rowCount));
  const templateRowCount = ws.rowCount;

  clearRangeValues(ws, 2, templateRowCount, 1, 11);
  clearRangeValues(ws, 2, templateRowCount, 14, 15);

  for (let i = 0; i < rows.length; i += 1) {
    const rowNum = i + 2;
    const row = rows[i];

    ws.getCell(rowNum, 1).value = row.COD;
    ws.getCell(rowNum, 2).value = row.CUENTA;
    ws.getCell(rowNum, 3).value = row.EXT;
    ws.getCell(rowNum, 4).value = toExcelDateSerial(row.FECHA);
    ws.getCell(rowNum, 5).value = row.ORIGEN;
    ws.getCell(rowNum, 6).value = row.ASIENTO;
    ws.getCell(rowNum, 7).value = row.DOCU;
    ws.getCell(rowNum, 8).value = row.DETALLE;
    ws.getCell(rowNum, 9).value = round2(row.DEBE);
    ws.getCell(rowNum, 10).value = round2(row.HABER);
    ws.getCell(rowNum, 11).value = round2(row.SALDO);
  }

  const summaryLabels = readTemplateSummaryLabels(ws);
  const summary = buildSummary(rows, summaryLabels);
  for (const item of summary) {
    ws.getCell(item.row, 14).value = round2(item.debe);
    ws.getCell(item.row, 15).value = round2(item.haber);
  }

  return { workbook: wb, summary };
}

function preserveTemplateVisualWorkbook(templatePath, generatedPath, sheetName = SHEET_NAME) {
  const templateZip = new AdmZip(templatePath);
  const generatedZip = new AdmZip(generatedPath);

  const templateSheetPath = getWorksheetEntryPath(templateZip, sheetName);
  const generatedSheetPath = getWorksheetEntryPath(generatedZip, sheetName);
  const templateSheetXml = readZipText(templateZip, templateSheetPath);
  const generatedSheetXml = readZipText(generatedZip, generatedSheetPath);
  if (!templateSheetXml || !generatedSheetXml) {
    throw new Error("No se pudo comparar hoja XML entre plantilla y generado.");
  }

  const templateSheet = parseXml(templateSheetXml);
  const generatedSheet = parseXml(generatedSheetXml);
  const templateRows = getRowsArray(templateSheet);
  const generatedRows = getRowsArray(generatedSheet);
  const payloadBeforeMerge = capturePayloadSnapshot(generatedRows, 11);

  const templateRowMap = new Map();
  const generatedCellMap = new Map();
  let templateMaxRow = 1;
  let generatedMaxRow = 1;
  let templatePrototypeRow = null;

  for (const row of templateRows) {
    const rowNumber = Number(row?.["@_r"] || 0);
    if (rowNumber > 0) {
      templateRowMap.set(rowNumber, row);
      templateMaxRow = Math.max(templateMaxRow, rowNumber);
    }
  }

  if (templateRowMap.has(templateMaxRow)) {
    templatePrototypeRow = cloneTemplateRow(templateRowMap.get(templateMaxRow), templateMaxRow);
  }

  for (const row of generatedRows) {
    const rowNumber = Number(row?.["@_r"] || 0);
    if (rowNumber <= 0) {
      continue;
    }
    const cells = getCellsArray(row);
    for (const cell of cells) {
      const ref = parseCellRef(cell?.["@_r"]);
      if (!ref || ref.col < 1 || ref.col > 16) {
        continue;
      }
      generatedCellMap.set(`${ref.row}:${ref.col}`, cell);
      if (hasCellPayload(cell)) {
        generatedMaxRow = Math.max(generatedMaxRow, ref.row);
      }
    }
  }

  const maxRow = Math.max(templateMaxRow, generatedMaxRow);
  for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
    let row = templateRowMap.get(rowNumber);
    if (!row) {
      row = templatePrototypeRow ? cloneTemplateRow(templatePrototypeRow, rowNumber) : { "@_r": String(rowNumber) };
      templateRows.push(row);
      templateRowMap.set(rowNumber, row);
    }

    const rowCells = getCellsArray(row);
    const rowCellMap = new Map();
    for (const cell of rowCells) {
      const ref = parseCellRef(cell?.["@_r"]);
      if (ref) {
        rowCellMap.set(ref.col, cell);
      }
    }

    for (let col = 1; col <= 16; col += 1) {
      const generatedCell = generatedCellMap.get(`${rowNumber}:${col}`);
      let targetCell = rowCellMap.get(col);

      if (!targetCell) {
        if (!generatedCell) {
          continue;
        }
        targetCell = { "@_r": cellRef(col, rowNumber) };
        rowCellMap.set(col, targetCell);
        rowCells.push(targetCell);
      }

      if (generatedCell) {
        copyCellPayload(targetCell, generatedCell);
      } else {
        clearCellPayload(targetCell);
      }
    }

    rowCells.sort((a, b) => {
      const ca = parseCellRef(a?.["@_r"]);
      const cb = parseCellRef(b?.["@_r"]);
      return (ca?.col || 0) - (cb?.col || 0);
    });

    setCellsArray(row, rowCells.filter((cell) => cell["@_s"] !== undefined || hasCellPayload(cell)));
  }

  templateRows.sort((a, b) => Number(a?.["@_r"] || 0) - Number(b?.["@_r"] || 0));
  const payloadAfterMerge = capturePayloadSnapshot(templateRows, 11);

  if (
    payloadBeforeMerge.rows !== payloadAfterMerge.rows
    || payloadBeforeMerge.payloadCells !== payloadAfterMerge.payloadCells
    || payloadBeforeMerge.payloadHash !== payloadAfterMerge.payloadHash
  ) {
    throw new Error(
      `Validacion post-merge Accion 3 fallida: rows ${payloadBeforeMerge.rows}=>${payloadAfterMerge.rows}, `
      + `payload ${payloadBeforeMerge.payloadCells}=>${payloadAfterMerge.payloadCells}.`,
    );
  }

  setRowsArray(templateSheet, templateRows);
  if (generatedSheet.worksheet?.dimension) {
    templateSheet.worksheet.dimension = deepClone(generatedSheet.worksheet.dimension);
  }
  updateZipText(templateZip, templateSheetPath, templateSheet);

  const generatedSharedStrings = generatedZip.getEntry("xl/sharedStrings.xml");
  if (generatedSharedStrings) {
    templateZip.updateFile("xl/sharedStrings.xml", generatedSharedStrings.getData());
  }

  stripCalcChain(templateZip);
  templateZip.writeZip(generatedPath);

  return {
    rows_before_merge: payloadBeforeMerge.rows,
    rows_after_merge: payloadAfterMerge.rows,
    payload_cells_before: payloadBeforeMerge.payloadCells,
    payload_cells_after: payloadAfterMerge.payloadCells,
    payload_hash_before: payloadBeforeMerge.payloadHash,
    payload_hash_after: payloadAfterMerge.payloadHash,
    merge_integrity_ok: true,
  };
}

async function verifyOutputWorkbook(outputPath, templatePath, rowsCount) {
  const templateZip = new AdmZip(templatePath);
  const outputZip = new AdmZip(outputPath);
  const criticalEntries = [
    "xl/pivotTables/pivotTable1.xml",
    "xl/worksheets/_rels/sheet1.xml.rels",
    "xl/pivotCache/pivotCacheDefinition1.xml",
  ];
  for (const entryPath of criticalEntries) {
    const expected = templateZip.getEntry(entryPath);
    const actual = outputZip.getEntry(entryPath);
    if (expected && !actual) {
      throw new Error(`Validacion final: falta artefacto visual de plantilla (${entryPath}).`);
    }
  }

  const wb = XLSX.readFile(outputPath, { cellFormula: true });
  if (wb.SheetNames.length !== 1 || wb.SheetNames[0] !== SHEET_NAME) {
    throw new Error(`Validacion final: el archivo debe tener una sola hoja llamada ${SHEET_NAME}.`);
  }

  const ws = wb.Sheets[SHEET_NAME];
  const headers = EXPECTED_HEADERS.map((_, idx) => ws[XLSX.utils.encode_cell({ r: 0, c: idx })]?.v || "");
  for (let i = 0; i < EXPECTED_HEADERS.length; i += 1) {
    if (sanitizeText(headers[i]).toUpperCase() !== EXPECTED_HEADERS[i]) {
      throw new Error(`Validacion final: encabezado incorrecto en columna ${i + 1}.`);
    }
  }

  for (let row = 2; row <= rowsCount + 1; row += 1) {
    const dateCell = ws[`D${row}`];
    if (!dateCell || typeof dateCell.v !== "number") {
      throw new Error(`Validacion final: FECHA invalida en fila ${row}.`);
    }
    if (Math.abs(dateCell.v - Math.trunc(dateCell.v)) > 1e-9) {
      throw new Error(`Validacion final: FECHA con fraccion horaria en fila ${row}.`);
    }
  }

  const templateWb = new ExcelJS.Workbook();
  await templateWb.xlsx.readFile(templatePath);
  const outWb = new ExcelJS.Workbook();
  await outWb.xlsx.readFile(outputPath);

  const templateWs = templateWb.getWorksheet(SHEET_NAME);
  const outWs = outWb.getWorksheet(SHEET_NAME);
  if (!templateWs || !outWs) {
    throw new Error(`Validacion final: no se pudo comparar formato de hoja ${SHEET_NAME}.`);
  }
  const styleSig = createStyleSignatureCache();

  for (let col = 1; col <= 16; col += 1) {
    const expected = styleSig(templateWs.getCell(1, col).style);
    const actual = styleSig(outWs.getCell(1, col).style);
    if (expected !== actual) {
      throw new Error(`Validacion final: estilo de encabezado alterado en fila 1, columna ${col}.`);
    }
  }

  const lastDataRow = rowsCount + 1;
  if (lastDataRow >= 2) {
    const expectedDataStyleRow = Math.min(lastDataRow, templateWs.rowCount);
    for (let col = 1; col <= 11; col += 1) {
      const expected = styleSig(templateWs.getCell(expectedDataStyleRow, col).style);
      const actual = styleSig(outWs.getCell(lastDataRow, col).style);
      if (expected !== actual) {
        throw new Error(`Validacion final: estilo de datos alterado en ${lastDataRow}:${col}.`);
      }
    }
  }

  const summaryLabels = readTemplateSummaryLabels(templateWs);
  for (const item of summaryLabels) {
    for (let col = 13; col <= 16; col += 1) {
      const expected = styleSig(templateWs.getCell(item.row, col).style);
      const actual = styleSig(outWs.getCell(item.row, col).style);
      if (expected !== actual) {
        throw new Error(`Validacion final: estilo lateral alterado en ${item.row}:${col}.`);
      }
    }
  }
}

module.exports = {
  buildWorkbookFromTemplate,
  preserveTemplateVisualWorkbook,
  verifyOutputWorkbook,
};
