const fs = require("fs");
const crypto = require("crypto");

const AdmZip = require("adm-zip");
const ExcelJS = require("exceljs");
const XLSX = require("xlsx");

const {
  sanitizeText,
  deepClone,
  round2,
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
  columnNumberToName,
  cellRef,
  copyCellPayload,
  clearCellPayload,
  hasCellPayload,
  stripCalcChain,
  styleSignature,
} = require("../shared/excel-template-utils");
const {
  EXPECTED_COLUMNS,
  SHEET_NAME,
} = require("./constants");

function formatPercentLabel(value) {
  if (Number.isInteger(value)) {
    return String(value);
  }
  const text = String(value);
  return text.replace(/(\.\d*?[1-9])0+$/g, "$1").replace(/\.0+$/g, "");
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

function shouldIncludeRowInSummary(row) {
  // Placeholder retentions like 999999999 stay in the detail grid,
  // but the manual example excludes them from the lateral summary totals.
  return Number(row?.numRt) !== 999999999;
}

function buildSummary(rows) {
  const typeOrder = ["IVA", "RENTA"];
  const typeMap = new Map();
  let totalBaseCents = 0;
  let totalRetCents = 0;

  for (const row of rows) {
    if (!shouldIncludeRowInSummary(row)) {
      continue;
    }

    const type = row.tipo;
    if (!typeMap.has(type)) {
      typeMap.set(type, {
        totalBaseCents: 0,
        totalRetCents: 0,
        percentMap: new Map(),
      });
    }

    const bucket = typeMap.get(type);
    const baseCents = toCents(row.base);
    const retCents = toCents(row.retencion);
    bucket.totalBaseCents += baseCents;
    bucket.totalRetCents += retCents;
    totalBaseCents += baseCents;
    totalRetCents += retCents;

    const key = round2(row.percent);
    if (!bucket.percentMap.has(key)) {
      bucket.percentMap.set(key, { baseCents: 0, retCents: 0 });
    }
    const percentBucket = bucket.percentMap.get(key);
    percentBucket.baseCents += baseCents;
    percentBucket.retCents += retCents;
  }

  const entries = [];
  const seenTypes = new Set();
  for (const type of typeOrder) {
    if (typeMap.has(type)) {
      seenTypes.add(type);
      const bucket = typeMap.get(type);
      entries.push({
        kind: "type",
        label: type,
        base: fromCents(bucket.totalBaseCents),
        ret: fromCents(bucket.totalRetCents),
        calc: null,
        diff: null,
      });

      const percents = [...bucket.percentMap.keys()].sort((a, b) => a - b);
      for (const percent of percents) {
        const item = bucket.percentMap.get(percent);
        entries.push({
          kind: "detail",
          label: formatPercentLabel(percent),
          base: fromCents(item.baseCents),
          ret: fromCents(item.retCents),
          calc: fromCents(item.retCents),
          diff: 0,
        });
      }
    }
  }

  const remainingTypes = [...typeMap.keys()].filter((type) => !seenTypes.has(type)).sort();
  for (const type of remainingTypes) {
    const bucket = typeMap.get(type);
    entries.push({
      kind: "type",
      label: type,
      base: fromCents(bucket.totalBaseCents),
      ret: fromCents(bucket.totalRetCents),
      calc: null,
      diff: null,
    });
    const percents = [...bucket.percentMap.keys()].sort((a, b) => a - b);
    for (const percent of percents) {
      const item = bucket.percentMap.get(percent);
      entries.push({
        kind: "detail",
        label: formatPercentLabel(percent),
        base: fromCents(item.baseCents),
        ret: fromCents(item.retCents),
        calc: fromCents(item.retCents),
        diff: 0,
      });
    }
  }

  entries.push({
    kind: "total",
    label: "Total general",
    base: fromCents(totalBaseCents),
    ret: fromCents(totalRetCents),
    calc: null,
    diff: null,
  });

  return entries;
}

function hasStyle(cellStyle) {
  return cellStyle && Object.keys(cellStyle).length > 0;
}

function findLastFullyStyledRow(ws, startRow, endRow, startCol, endCol) {
  let last = startRow;
  for (let row = startRow; row <= endRow; row += 1) {
    let full = true;
    for (let col = startCol; col <= endCol; col += 1) {
      if (!hasStyle(ws.getCell(row, col).style)) {
        full = false;
        break;
      }
    }
    if (full) {
      last = row;
    }
  }
  return last;
}

function normalizeColumnWidth(width) {
  return Number.isFinite(width) ? Number(width) : null;
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

function applyFallbackStylesForExtendedDataRow(ws, rowNumber, fallbackStyles) {
  for (let col = 1; col <= 11; col += 1) {
    ws.getCell(rowNumber, col).style = deepClone(fallbackStyles[col - 1] || {});
  }
}

function applyFallbackStylesForExtendedSummaryRow(ws, rowNumber, kind, fallbackHeader, fallbackType, fallbackDetail, fallbackTotal) {
  let source = fallbackDetail;
  if (rowNumber === 1) {
    source = fallbackHeader;
  } else if (kind === "type") {
    source = fallbackType;
  } else if (kind === "total") {
    source = fallbackTotal;
  }

  for (let col = 12; col <= 16; col += 1) {
    ws.getCell(rowNumber, col).style = deepClone(source[col - 12] || {});
  }
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

    const rowCells = getCellsArray(row);
    for (const cell of rowCells) {
      const parsedRef = parseCellRef(cell?.["@_r"]);
      if (!parsedRef || parsedRef.colNumber < 1 || parsedRef.colNumber > maxColumn) {
        continue;
      }
      if (!hasCellPayload(cell)) {
        continue;
      }

      rowsWithPayload.add(parsedRef.rowNumber);
      payloadEntries.push(`${parsedRef.rowNumber}:${parsedRef.colNumber}:${cellPayloadSignature(cell)}`);
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

function normalizeSummaryLabel(label) {
  return sanitizeText(label).replace(/\s+/g, " ").trim().toUpperCase();
}

function readTemplateSummaryLayout(ws) {
  const layout = [];
  for (let row = 2; row <= 200; row += 1) {
    const label = sanitizeText(ws.getCell(row, 12).value);
    if (!label) {
      if (layout.length > 0) {
        break;
      }
      continue;
    }
    const normalized = normalizeSummaryLabel(label);
    let kind = "detail";
    if (normalized === "TOTAL GENERAL") {
      kind = "total";
    } else if (!/^[0-9.]+$/.test(normalized)) {
      kind = "type";
    }
    layout.push({ row, label, kind });
  }
  return layout;
}

function preserveTemplateVisualWorkbook(templatePath, generatedPath, sheetName = SHEET_NAME, dataRowCount = null) {
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
  const payloadBeforeMerge = capturePayloadSnapshot(generatedRows, 16);

  const templateRowMap = new Map();
  const templateStyleByCol = new Map();
  const dataFallbackStyleByCol = new Map();
  const generatedCellMap = new Map();
  let templateMaxRow = 1;
  let generatedMaxRowWithPayload = 1;

  for (const row of templateRows) {
    const rowNumber = Number(row?.["@_r"] || 0);
    if (rowNumber > 0) {
      templateRowMap.set(rowNumber, row);
      templateMaxRow = Math.max(templateMaxRow, rowNumber);
    }

    const rowCells = getCellsArray(row);
    const rowStyleByCol = new Map();
    for (const cell of rowCells) {
      const parsedRef = parseCellRef(cell?.["@_r"]);
      if (!parsedRef || parsedRef.colNumber < 1 || parsedRef.colNumber > 16) {
        continue;
      }
      if (cell["@_s"] !== undefined) {
        templateStyleByCol.set(parsedRef.colNumber, cell["@_s"]);
        rowStyleByCol.set(parsedRef.colNumber, cell["@_s"]);
      }
    }

    let fullDataStyled = true;
    for (let colNumber = 1; colNumber <= 10; colNumber += 1) {
      if (!rowStyleByCol.has(colNumber)) {
        fullDataStyled = false;
        break;
      }
    }
    if (fullDataStyled) {
      for (let colNumber = 1; colNumber <= 10; colNumber += 1) {
        dataFallbackStyleByCol.set(colNumber, rowStyleByCol.get(colNumber));
      }
    }
  }

  for (const row of generatedRows) {
    const rowNumber = Number(row?.["@_r"] || 0);
    if (rowNumber <= 0) {
      continue;
    }

    const cells = getCellsArray(row);
    for (const cell of cells) {
      const ref = String(cell?.["@_r"] || "");
      const parsedRef = parseCellRef(ref);
      if (!parsedRef || parsedRef.colNumber < 1 || parsedRef.colNumber > 16) {
        continue;
      }
      const key = `${parsedRef.rowNumber}:${parsedRef.colNumber}`;
      generatedCellMap.set(key, cell);
      if (hasCellPayload(cell) || cell["@_s"] !== undefined) {
        generatedMaxRowWithPayload = Math.max(generatedMaxRowWithPayload, parsedRef.rowNumber);
      }
    }
  }

  const maxRow = Math.max(templateMaxRow, generatedMaxRowWithPayload);
  for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
    let row = templateRowMap.get(rowNumber);
    if (!row) {
      row = { "@_r": String(rowNumber) };
      templateRows.push(row);
      templateRowMap.set(rowNumber, row);
    }

    const rowCells = getCellsArray(row);
    const rowCellMap = new Map();
    for (const cell of rowCells) {
      const parsedRef = parseCellRef(cell?.["@_r"]);
      if (!parsedRef) {
        continue;
      }
      rowCellMap.set(parsedRef.colNumber, cell);
    }

    for (let colNumber = 1; colNumber <= 16; colNumber += 1) {
      const key = `${rowNumber}:${colNumber}`;
      const generatedCell = generatedCellMap.get(key);
      let targetCell = rowCellMap.get(colNumber);

      if (!targetCell) {
        if (!generatedCell) {
          continue;
        }
        targetCell = { "@_r": cellRef(colNumber, rowNumber) };
        if (colNumber >= 1 && colNumber <= 10) {
          const fallbackStyle = dataFallbackStyleByCol.get(colNumber) || templateStyleByCol.get(colNumber);
          if (fallbackStyle !== undefined) {
            targetCell["@_s"] = fallbackStyle;
          }
        }
        rowCellMap.set(colNumber, targetCell);
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
      return (ca?.colNumber || 0) - (cb?.colNumber || 0);
    });

    const filtered = rowCells.filter((cell) => {
      const keepByStyle = cell["@_s"] !== undefined;
      const keepByPayload = hasCellPayload(cell);
      return keepByStyle || keepByPayload;
    });
    setCellsArray(row, filtered);
  }

  templateRows.sort((a, b) => Number(a?.["@_r"] || 0) - Number(b?.["@_r"] || 0));
  const payloadAfterMerge = capturePayloadSnapshot(templateRows, 16);

  if (
    payloadBeforeMerge.rows !== payloadAfterMerge.rows
    || payloadBeforeMerge.payloadCells !== payloadAfterMerge.payloadCells
    || payloadBeforeMerge.payloadHash !== payloadAfterMerge.payloadHash
  ) {
    throw new Error(
      `Validacion post-merge Accion 2 fallida: rows ${payloadBeforeMerge.rows}=>${payloadAfterMerge.rows}, `
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

  if (Number.isFinite(dataRowCount)) {
    const endRow = Math.max(2, Number(dataRowCount) + 1);
    const lastCol = columnNumberToName(EXPECTED_COLUMNS.length);
    const ref = `A1:${lastCol}${endRow}`;

    const cachePath = "xl/pivotCache/pivotCacheDefinition1.xml";
    const cacheXmlText = readZipText(templateZip, cachePath);
    if (cacheXmlText) {
      const cacheObj = parseXml(cacheXmlText);
      if (cacheObj.pivotCacheDefinition?.cacheSource?.worksheetSource) {
        cacheObj.pivotCacheDefinition.cacheSource.worksheetSource["@_ref"] = ref;
        cacheObj.pivotCacheDefinition.cacheSource.worksheetSource["@_sheet"] = sheetName;
      }
      if (cacheObj.pivotCacheDefinition) {
        cacheObj.pivotCacheDefinition["@_recordCount"] = String(Math.max(0, Number(dataRowCount)));
        cacheObj.pivotCacheDefinition["@_refreshOnLoad"] = "1";
      }
      updateZipText(templateZip, cachePath, cacheObj);
    }

    const pivotPath = "xl/pivotTables/pivotTable1.xml";
    const pivotXmlText = readZipText(templateZip, pivotPath);
    if (pivotXmlText) {
      const pivotObj = parseXml(pivotXmlText);
      if (pivotObj.pivotTableDefinition) {
        pivotObj.pivotTableDefinition["@_refreshOnLoad"] = "1";
      }
      updateZipText(templateZip, pivotPath, pivotObj);
    }

    if (templateSheet?.worksheet?.autoFilter?.["@_ref"]) {
      templateSheet.worksheet.autoFilter["@_ref"] = ref;
    }
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

async function verifyOutputWorkbook(outputPath, templatePath, rowsCount, summary) {
  const templateZip = new AdmZip(templatePath);
  const outputZip = new AdmZip(outputPath);
  const criticalEntries = [
    "xl/pivotTables/pivotTable1.xml",
    "xl/worksheets/_rels/sheet1.xml.rels",
    "xl/pivotCache/pivotCacheDefinition1.xml",
    "xl/externalLinks/externalLink1.xml",
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
  const headers = EXPECTED_COLUMNS.map((_, idx) => ws[XLSX.utils.encode_cell({ r: 0, c: idx })]?.v || "");
  for (let i = 0; i < EXPECTED_COLUMNS.length; i += 1) {
    if (sanitizeText(headers[i]).toUpperCase() !== EXPECTED_COLUMNS[i]) {
      throw new Error(`Validacion final: encabezado incorrecto en columna ${i + 1}.`);
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
    const tWidth = normalizeColumnWidth(templateWs.getColumn(col).width);
    const oWidth = normalizeColumnWidth(outWs.getColumn(col).width);
    if (tWidth !== oWidth) {
      throw new Error(`Validacion final: ancho de columna alterado en ${col}.`);
    }
  }

  for (let col = 1; col <= 16; col += 1) {
    const expected = styleSig(templateWs.getCell(1, col).style);
    const actual = styleSig(outWs.getCell(1, col).style);
    if (actual !== expected) {
      throw new Error(`Validacion final: estilo de encabezado alterado en fila 1, columna ${col}.`);
    }
  }

  const templateRowCount = templateWs.rowCount;
  const lastStyledDataRow = findLastFullyStyledRow(templateWs, 2, templateRowCount, 1, 10);
  const maxDataRow = rowsCount + 1;
  for (let rowNum = 2; rowNum <= maxDataRow; rowNum += 1) {
    for (let col = 1; col <= 11; col += 1) {
      const templateStyle = rowNum <= templateRowCount ? templateWs.getCell(rowNum, col).style : null;
      const expectedTemplate = styleSig(templateStyle || {});
      const expectedFallback = styleSig(templateWs.getCell(lastStyledDataRow, col).style || {});
      const expectedEmpty = styleSig({});
      const actual = styleSig(outWs.getCell(rowNum, col).style);
      if (actual !== expectedTemplate && actual !== expectedFallback && actual !== expectedEmpty) {
        throw new Error(`Validacion final: estilo de datos alterado en ${rowNum}:${col}.`);
      }
    }
  }

  const summaryRows = Array.isArray(summary) ? summary : [];
  for (let i = 0; i < summaryRows.length; i += 1) {
    const rowNum = i + 2;
    let sourceRow;
    if (rowNum <= templateRowCount) {
      sourceRow = rowNum;
    } else if (summaryRows[i].kind === "type") {
      sourceRow = 2;
    } else if (summaryRows[i].kind === "total") {
      sourceRow = 16;
    } else {
      sourceRow = 3;
    }
    for (let col = 12; col <= 16; col += 1) {
      const expected = styleSig(templateWs.getCell(sourceRow, col).style);
      const actual = styleSig(outWs.getCell(rowNum, col).style);
      if (actual !== expected) {
        throw new Error(`Validacion final: estilo de resumen alterado en ${rowNum}:${col}.`);
      }
    }
  }
}

async function buildWorkbookFromTemplate(templatePath, rows) {
  if (!fs.existsSync(templatePath)) {
    throw new Error(`No se encontro plantilla de Accion 2: ${templatePath}`);
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(templatePath);
  const ws = wb.getWorksheet(SHEET_NAME);
  if (!ws) {
    throw new Error(`La plantilla de Accion 2 no contiene la hoja ${SHEET_NAME}.`);
  }

  const templateRowCount = ws.rowCount;
  const maxRows = Math.max(templateRowCount, rows.length + 40, 550);
  const lastStyledDataRow = findLastFullyStyledRow(ws, 2, templateRowCount, 1, 10);
  const fallbackDataStyles = [];
  for (let col = 1; col <= 11; col += 1) {
    fallbackDataStyles.push(deepClone(ws.getCell(lastStyledDataRow, col).style));
  }

  clearRangeValues(ws, 2, maxRows, 1, 11);

  for (let i = 0; i < rows.length; i += 1) {
    const rowNum = i + 2;
    const row = rows[i];

    ws.getCell(`A${rowNum}`).value = row.numRt;
    ws.getCell(`B${rowNum}`).value = row.proveedor;
    ws.getCell(`C${rowNum}`).value = row.fecha;
    ws.getCell(`D${rowNum}`).value = row.fechaCont;
    ws.getCell(`E${rowNum}`).value = row.tipo;
    ws.getCell(`F${rowNum}`).value = row.cod;
    ws.getCell(`G${rowNum}`).value = row.fact;
    ws.getCell(`H${rowNum}`).value = round2(row.percent);
    ws.getCell(`I${rowNum}`).value = round2(row.base);
    ws.getCell(`J${rowNum}`).value = round2(row.retencion);
    ws.getCell(`K${rowNum}`).value = null;
    if (rowNum > lastStyledDataRow) {
      applyFallbackStylesForExtendedDataRow(ws, rowNum, fallbackDataStyles);
      ws.getRow(rowNum).height = ws.getRow(lastStyledDataRow).height;
    }
  }

  const templateSummaryLayout = readTemplateSummaryLayout(ws);
  const computedSummary = buildSummary(rows);
  const summaryLookup = new Map(
    computedSummary.map((item) => [normalizeSummaryLabel(item.label), item]),
  );
  const summary = [];
  for (const slot of templateSummaryLayout) {
    const match = summaryLookup.get(normalizeSummaryLabel(slot.label));
    const base = match ? match.base : 0;
    const ret = match ? match.ret : 0;

    ws.getCell(slot.row, 13).value = base;
    ws.getCell(slot.row, 14).value = ret;

    summary.push({
      kind: slot.kind,
      label: slot.label,
      base,
      ret,
      calc: match?.calc ?? null,
      diff: match?.diff ?? null,
    });
  }

  return { workbook: wb, summary };
}

module.exports = {
  buildWorkbookFromTemplate,
  preserveTemplateVisualWorkbook,
  verifyOutputWorkbook,
};
