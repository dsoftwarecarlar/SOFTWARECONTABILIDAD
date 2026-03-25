const fs = require("fs");

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
  captureStyleMatrix,
  applyStyleFromMatrix,
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

function buildSummary(rows, templateLabels) {
  const bucket = new Map();
  let totalDebe = 0;
  let totalHaber = 0;

  for (const row of rows) {
    const label = sanitizeText(row.CUENTA);
    if (!bucket.has(label)) {
      bucket.set(label, { debe: 0, haber: 0 });
    }
    const item = bucket.get(label);
    item.debe += row.DEBE;
    item.haber += row.HABER;
    totalDebe += row.DEBE;
    totalHaber += row.HABER;
  }

  const result = [];
  for (const labelItem of templateLabels) {
    if (labelItem.label.toUpperCase() === "TOTAL GENERAL") {
      result.push({
        row: labelItem.row,
        label: labelItem.label,
        debe: round2(totalDebe),
        haber: round2(totalHaber),
      });
      continue;
    }

    const values = bucket.get(labelItem.label) || { debe: 0, haber: 0 };
    result.push({
      row: labelItem.row,
      label: labelItem.label,
      debe: round2(values.debe),
      haber: round2(values.haber),
    });
  }

  return result;
}

async function buildWorkbookFromTemplate(templatePath, rowPlan, movementRows) {
  if (!fs.existsSync(templatePath)) {
    throw new Error(`No se encontro plantilla de Accion 4: ${templatePath}`);
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(templatePath);
  const ws = wb.getWorksheet(SHEET_NAME);
  if (!ws) {
    throw new Error(`La plantilla de Accion 4 no contiene la hoja ${SHEET_NAME}.`);
  }

  ensureTemplateCapacity(ws, Math.max(rowPlan.length + 1, ws.rowCount));
  const templateRowCount = ws.rowCount;

  const templateStyles = captureStyleMatrix(ws, templateRowCount, 16);
  clearRangeValues(ws, 2, templateRowCount, 1, 11);
  clearRangeValues(ws, 2, templateRowCount, 14, 15);

  for (let row = 1; row <= templateRowCount; row += 1) {
    for (let col = 1; col <= 16; col += 1) {
      applyStyleFromMatrix(ws, templateStyles, row, col);
    }
  }

  for (let i = 0; i < rowPlan.length; i += 1) {
    const rowNum = i + 2;
    const planItem = rowPlan[i];

    if (planItem.type === "movement") {
      const row = planItem.row;
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
    } else if (planItem.type === "subtotal") {
      ws.getCell(rowNum, 9).value = { formula: `SUM(I${planItem.fromRow}:I${planItem.toRow})` };
      ws.getCell(rowNum, 10).value = { formula: `SUM(J${planItem.fromRow}:J${planItem.toRow})` };
    } else if (planItem.type === "subtotal_balance") {
      ws.getCell(rowNum, 11).value = {
        formula: planItem.mode === "haber_minus_debe"
          ? `+J${rowNum - 1}-I${rowNum - 1}`
          : `+I${rowNum - 1}-J${rowNum - 1}`,
      };
    }

    for (let col = 1; col <= 11; col += 1) {
      applyStyleFromMatrix(ws, templateStyles, rowNum, col);
    }
  }

  const summaryLabels = readTemplateSummaryLabels(ws);
  const summary = buildSummary(movementRows, summaryLabels);
  for (const item of summary) {
    ws.getCell(item.row, 14).value = round2(item.debe);
    ws.getCell(item.row, 15).value = round2(item.haber);
    for (let col = 13; col <= 16; col += 1) {
      applyStyleFromMatrix(ws, templateStyles, item.row, col);
    }
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
}

async function verifyOutputWorkbook(outputPath, templatePath, rowPlan) {
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

  for (let index = 0; index < rowPlan.length; index += 1) {
    if (rowPlan[index].type !== "movement") {
      continue;
    }
    const row = index + 2;
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

  for (let col = 1; col <= 16; col += 1) {
    const expected = styleSignature(templateWs.getCell(1, col).style);
    const actual = styleSignature(outWs.getCell(1, col).style);
    if (expected !== actual) {
      throw new Error(`Validacion final: estilo de encabezado alterado en fila 1, columna ${col}.`);
    }
  }

  let lastMovementRow = 1;
  for (let index = rowPlan.length - 1; index >= 0; index -= 1) {
    if (rowPlan[index].type === "movement") {
      lastMovementRow = index + 2;
      break;
    }
  }

  if (lastMovementRow >= 2) {
    const expectedDataStyleRow = Math.min(lastMovementRow, templateWs.rowCount);
    for (let col = 1; col <= 11; col += 1) {
      const expected = styleSignature(templateWs.getCell(expectedDataStyleRow, col).style);
      const actual = styleSignature(outWs.getCell(lastMovementRow, col).style);
      if (expected !== actual) {
        throw new Error(`Validacion final: estilo de datos alterado en ${lastMovementRow}:${col}.`);
      }
    }
  }

  const summaryLabels = readTemplateSummaryLabels(templateWs);
  for (const item of summaryLabels) {
    for (let col = 13; col <= 16; col += 1) {
      const expected = styleSignature(templateWs.getCell(item.row, col).style);
      const actual = styleSignature(outWs.getCell(item.row, col).style);
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
