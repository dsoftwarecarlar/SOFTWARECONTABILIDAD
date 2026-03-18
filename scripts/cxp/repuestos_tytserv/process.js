const path = require("path");
const fs = require("fs");
const crypto = require("crypto");
const ExcelJS = require("exceljs");
const XLSX = require("xlsx");
const AdmZip = require("adm-zip");

const { sanitizeText } = require("../shared/core-utils");

const REP_MAX_COLUMN = 41;
const NC_MAX_COLUMN = 31;
const REP_DETAIL_START_ROW = 11;
const NC_DETAIL_START_ROW = 8;
const MAYOR_IVA_START_ROW = 299;
const MAYOR_IVA_END_ROW = 366;
const MAYOR_IVA_LAST_COLUMN = 10;

const REP_PAYLOAD_COLUMNS = Array.from({ length: REP_MAX_COLUMN }, (_value, index) => index + 1);

const REP_TEXT_COLUMNS = new Set([
  1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16,
  37, 38, 39, 40, 41,
]);
const NC_TOTAL_COLUMNS = [22, 23, 24, 25, 26, 27, 28, 30];

const SHEET_CONFIGS = [
  { key: "tyt", label: "MATRIZ", argKeys: ["inputtyt", "input-tyt"], targetSheet: "REP TYT" },
  { key: "peug", label: "PEUGEOT", argKeys: ["inputpeug", "input-peug"], targetSheet: "REP PEUGT" },
  { key: "chgn", label: "CHANGAN", argKeys: ["inputchgn", "input-chgn"], targetSheet: "REP CHGN" },
  { key: "szk", label: "SUZUKI", argKeys: ["inputszk", "input-szk"], targetSheet: "REP SZK" },
];

const MY_LAYOUTS = {
  tyt: {
    mySheetName: "MY REP TYT",
    detailColumn: 7,
    seatColumn: 6,
    dateColumn: 4,
    debitColumn: 8,
    creditColumn: 9,
    saldoColumn: 10,
    sections: [
      { name: "sales", startRow: 2, endRow: 37, amountColumn: 9, oppositeColumn: 8 },
      { name: "discount", startRow: 42, endRow: 52, amountColumn: 8, oppositeColumn: 9 },
    ],
  },
  peug: {
    mySheetName: "MY REP PEUG",
    detailColumn: 7,
    seatColumn: 6,
    dateColumn: 4,
    debitColumn: 8,
    creditColumn: 9,
    saldoColumn: 10,
    sections: [
      { name: "sales", startRow: 2, endRow: 18, amountColumn: 9, oppositeColumn: 8 },
      { name: "discount", startRow: 21, endRow: 24, amountColumn: 8, oppositeColumn: 9 },
    ],
  },
  chgn: {
    mySheetName: "MY REP CHGN",
    detailColumn: 8,
    seatColumn: 6,
    dateColumn: 4,
    debitColumn: 9,
    creditColumn: 10,
    saldoColumn: 11,
    sections: [
      { name: "sales", startRow: 2, endRow: 7, amountColumn: 10, oppositeColumn: 9 },
      { name: "discount", startRow: 12, endRow: 14, amountColumn: 9, oppositeColumn: 10 },
    ],
  },
  szk: {
    mySheetName: "MY REP SZK",
    detailColumn: 7,
    seatColumn: 6,
    dateColumn: 4,
    debitColumn: 8,
    creditColumn: 9,
    saldoColumn: 10,
    sections: [
      { name: "sales", startRow: 2, endRow: 30, amountColumn: 9, oppositeColumn: 8 },
      { name: "discount", startRow: 34, endRow: 49, amountColumn: 8, oppositeColumn: 9 },
    ],
  },
};

const MONTHS = {
  ENE: 0,
  FEB: 1,
  MAR: 2,
  ABR: 3,
  MAY: 4,
  JUN: 5,
  JUL: 6,
  AGO: 7,
  SEP: 8,
  OCT: 9,
  NOV: 10,
  DIC: 11,
  JAN: 0,
  APR: 3,
  AUG: 7,
  DEC: 11,
};

function parseArgs(argv) {
  const args = {};
  for (let index = 2; index < argv.length; index += 1) {
    const current = String(argv[index] || "");
    if (!current.startsWith("-")) {
      continue;
    }

    const key = current.replace(/^-+/, "").toLowerCase();
    const next = argv[index + 1];
    if (typeof next === "string" && !next.startsWith("-")) {
      args[key] = next;
      index += 1;
      continue;
    }

    args[key] = "";
  }

  if (args.template) {
    args.templatepath = args.template;
  }
  if (args.output) {
    args.outputpath = args.output;
  }

  return args;
}

function getArg(args, keys) {
  for (const key of keys) {
    if (typeof args[key] === "string" && args[key] !== "") {
      return args[key];
    }
  }

  return "";
}

function cloneDeep(value) {
  return value == null ? value : JSON.parse(JSON.stringify(value));
}

function normalizeText(value) {
  if (value == null) {
    return "";
  }

  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return sanitizeText(String(value));
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return sanitizeText(value.toISOString());
  }

  if (Array.isArray(value.richText)) {
    return sanitizeText(value.richText.map((part) => part.text || "").join(""));
  }

  if (typeof value.text === "string") {
    return sanitizeText(value.text);
  }

  if (value.result != null) {
    return normalizeText(value.result);
  }

  if (value.formula) {
    return sanitizeText(String(value.formula));
  }

  return sanitizeText(String(value));
}

function normalizeDocNumber(value) {
  return normalizeText(value).replace(/\s+/g, "");
}

function sourceAddress(row, column) {
  return XLSX.utils.encode_cell({ r: row - 1, c: column - 1 });
}

function getSourceCell(sheet, row, column) {
  return sheet[sourceAddress(row, column)] || null;
}

function sourceRenderableText(cell) {
  if (!cell) {
    return "";
  }

  if (cell.w != null && String(cell.w).trim() !== "") {
    return sanitizeText(String(cell.w));
  }

  if (cell.v == null) {
    return "";
  }

  return sanitizeText(String(cell.v));
}

function sourceLiteralText(cell) {
  if (!cell) {
    return "";
  }

  if (cell.w != null && String(cell.w) !== "") {
    return String(cell.w);
  }

  if (cell.v == null) {
    return "";
  }

  return String(cell.v);
}

function sourceCellText(sheet, row, column) {
  return sourceRenderableText(getSourceCell(sheet, row, column));
}

function parseNumericText(text) {
  const normalized = normalizeText(text).replace(/,/g, "");
  if (normalized === "") {
    return 0;
  }

  const numeric = Number(normalized);
  return Number.isFinite(numeric) ? numeric : 0;
}

function sourceCellNumber(sheet, row, column) {
  const cell = getSourceCell(sheet, row, column);
  if (!cell || cell.v == null) {
    return 0;
  }

  if (typeof cell.v === "number") {
    return cell.v;
  }

  return parseNumericText(sourceRenderableText(cell));
}

function sourceRowHasNeedle(sheet, row, needle, lastColumn = REP_MAX_COLUMN) {
  const expected = normalizeText(needle).toUpperCase();
  if (expected === "") {
    return false;
  }

  for (let column = 1; column <= lastColumn; column += 1) {
    if (sourceCellText(sheet, row, column).toUpperCase().includes(expected)) {
      return true;
    }
  }

  return false;
}

function getSourceLastRow(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:AO1");
  return Math.max(1, range.e.r + 1);
}

function findSourceTotalRow(sheet) {
  const maxRow = Math.max(200, getSourceLastRow(sheet) + 10);
  for (let row = 1; row <= maxRow; row += 1) {
    if (sourceRowHasNeedle(sheet, row, "TOTAL GENERAL")) {
      return row;
    }
  }

  return getSourceLastRow(sheet);
}

function findLastPopulatedSourceDetailRow(sheet, totalRow) {
  for (let row = totalRow - 1; row >= REP_DETAIL_START_ROW; row -= 1) {
    for (let column = 1; column <= REP_MAX_COLUMN; column += 1) {
      if (sourceCellText(sheet, row, column) !== "") {
        return row;
      }
    }
  }

  return REP_DETAIL_START_ROW - 1;
}

function createSha256(text) {
  return crypto.createHash("sha256").update(normalizeText(text)).digest("hex");
}

function buildSourcePayloadSignature(sheet) {
  const rows = [];
  const totalRow = findSourceTotalRow(sheet);
  for (let row = REP_DETAIL_START_ROW; row < totalRow; row += 1) {
    const document = sourceCellText(sheet, row, 5);
    // Se ignoran las filas vacías o anuladas para un conteo preciso.
    if (document === "" || document.toUpperCase().startsWith("ANULAD")) {
      continue;
    }

    rows.push(REP_PAYLOAD_COLUMNS.map((column) => sourceCellText(sheet, row, column)).join("|"));
  }

  return {
    rows: rows.length,
    hash: createSha256(rows.join("\n")),
  };
}

function assertRepSourceWorksheet(sheet, label) {
  const checks = [
    { row: 9, column: 5, needle: "# DOC" },
    { row: 9, column: 9, needle: "RUC" },
    { row: 9, column: 10, needle: "CLIENTE" },
    { row: 9, column: 11, needle: "CLIENTE" },
    { row: 9, column: 16, needle: "ITEM" },
    { row: 9, column: 18, needle: "SUBTOT" },
  ];

  for (const check of checks) {
    const actual = sourceCellText(sheet, check.row, check.column).toUpperCase();
    const expected = normalizeText(check.needle).toUpperCase();
    if (!actual.includes(expected)) {
      throw new Error(
        `La hoja fuente ${label} no coincide con la estructura esperada en fila ${check.row} columna ${check.column}. Esperado contiene '${check.needle}' y llego '${actual}'.`,
      );
    }
  }
}

function readSourceSheet(sourcePath, label) {
  const workbook = XLSX.readFile(sourcePath, {
    cellDates: true,
    cellFormula: true,
    cellNF: true,
    cellText: true,
    sheetStubs: true,
  });

  const preferredName = workbook.SheetNames.find(
    (sheetName) => normalizeText(sheetName).toUpperCase() === "REPLIBROVENTASGENERAL",
  );
  const sheetName = preferredName || "";
  if (!sheetName) {
    throw new Error(
      `El archivo fuente ${label} debe contener la hoja 'RepLibroVentasGeneral'. No se aceptan plantillas ni salidas ya generadas.`,
    );
  }

  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    throw new Error(`No se pudo abrir la hoja de ${label} dentro del archivo ${sourcePath}.`);
  }

  assertRepSourceWorksheet(sheet, label);

  return {
    path: sourcePath,
    label,
    sheetName,
    sheet,
    lastRow: getSourceLastRow(sheet),
    totalRow: findSourceTotalRow(sheet),
    payload: buildSourcePayloadSignature(sheet),
    dateRange: getSourceDateRange(sheet),
  };
}

function isMergedSlave(cell) {
  return !!(cell.isMerged && cell.master && cell.master.address !== cell.address);
}

function getWorksheetCellText(worksheet, row, column) {
  const cell = worksheet.getRow(row).getCell(column);
  if (isMergedSlave(cell)) {
    return "";
  }

  if (cell.value == null) {
    return "";
  }

  try {
    if (typeof cell.text === "string" && cell.text.trim() !== "") {
      return sanitizeText(cell.text);
    }
  } catch (_error) {
  }

  return normalizeText(cell.value);
}

function getWorksheetCellNumber(worksheet, row, column) {
  const cell = worksheet.getRow(row).getCell(column);
  if (isMergedSlave(cell)) {
    return 0;
  }
  const { value } = cell;
  if (value == null) {
    return 0;
  }

  if (typeof value === "number") {
    return value;
  }

  if (typeof value === "object" && value.result != null) {
    if (typeof value.result === "number") {
      return value.result;
    }

    return parseNumericText(value.result);
  }

  return parseNumericText(cell.text || value);
}

function worksheetRowHasNeedle(worksheet, row, needle, lastColumn = REP_MAX_COLUMN) {
  const expected = normalizeText(needle).toUpperCase();
  if (expected === "") {
    return false;
  }

  for (let column = 1; column <= lastColumn; column += 1) {
    if (getWorksheetCellText(worksheet, row, column).toUpperCase().includes(expected)) {
      return true;
    }
  }

  return false;
}

function findWorksheetRowContaining(worksheet, needle, startRow = 1, endRow = 400, lastColumn = REP_MAX_COLUMN) {
  for (let row = startRow; row <= endRow; row += 1) {
    if (worksheetRowHasNeedle(worksheet, row, needle, lastColumn)) {
      return row;
    }
  }

  return null;
}

function buildWorksheetPayloadSignature(worksheet, totalRow) {
  const rows = [];
  for (let row = REP_DETAIL_START_ROW; row < totalRow; row += 1) {
    const document = getWorksheetCellText(worksheet, row, 5);
    if (document === "") {
      continue;
    }

    rows.push(REP_PAYLOAD_COLUMNS.map((column) => getWorksheetCellText(worksheet, row, column)).join("|"));
  }

  return {
    rows: rows.length,
    hash: createSha256(rows.join("\n")),
  };
}

function clearWorksheetRange(worksheet, startRow, endRow, startColumn = 1, endColumn = REP_MAX_COLUMN) {
  if (startRow > endRow) {
    return;
  }

  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    for (let column = startColumn; column <= endColumn; column += 1) {
      const cell = row.getCell(column);
      if (isMergedSlave(cell)) {
        continue;
      }
      cell.value = null;
    }
  }
}

function clearWorksheetColumns(worksheet, startRow, endRow, columns) {
  if (startRow > endRow) {
    return;
  }

  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    for (const column of columns) {
      const cell = row.getCell(column);
      if (isMergedSlave(cell)) {
        continue;
      }
      cell.value = null;
    }
  }
}

function copyPayloadRowFromSource(sourceSheet, targetSheet, sourceRowNumber, targetRowNumber) {
  const targetRow = targetSheet.getRow(targetRowNumber);
  for (const column of REP_PAYLOAD_COLUMNS) {
    const targetCell = targetRow.getCell(column);
    if (isMergedSlave(targetCell)) {
      continue;
    }
    targetCell.value = convertSourceCellValue(getSourceCell(sourceSheet, sourceRowNumber, column), column);
  }
}

function rowHasRepPayloadContent(worksheet, rowNumber) {
  for (const column of REP_PAYLOAD_COLUMNS) {
    if (getWorksheetCellText(worksheet, rowNumber, column) !== "") {
      return true;
    }
  }

  return false;
}

function findRepFooterStartRow(worksheet, totalRow) {
  if (!Number.isInteger(totalRow) || totalRow <= REP_DETAIL_START_ROW) {
    return totalRow;
  }

  let footerStartRow = totalRow;
  for (let row = totalRow - 1; row >= REP_DETAIL_START_ROW; row -= 1) {
    if (rowHasRepPayloadContent(worksheet, row)) {
      break;
    }
    footerStartRow = row;
  }

  return footerStartRow;
}

function findRepStyleRow(worksheet, footerStartRow) {
  for (let row = footerStartRow - 1; row >= REP_DETAIL_START_ROW; row -= 1) {
    if (rowHasRepPayloadContent(worksheet, row)) {
      return row;
    }
  }

  return REP_DETAIL_START_ROW;
}

function getRowMergeDefinitions(worksheet, rowNumber) {
  return Object.values(worksheet._merges || {})
    .filter((range) => range && range.top === rowNumber && range.bottom === rowNumber)
    .map((range) => ({
      left: range.left,
      right: range.right,
    }));
}

function clearMergesInRowRange(worksheet, startRow, endRow) {
  const mergeEntries = Object.entries(worksheet._merges || {});
  for (const [_key, range] of mergeEntries) {
    if (!range) {
      continue;
    }
    if (range.top < startRow || range.bottom > endRow) {
      continue;
    }
    worksheet.unMergeCells(range.top, range.left, range.bottom, range.right);
  }
}

function applyRowMergeDefinitions(worksheet, rowNumber, mergeDefinitions) {
  clearMergesInRowRange(worksheet, rowNumber, rowNumber);
  for (const merge of mergeDefinitions) {
    worksheet.mergeCells(rowNumber, merge.left, rowNumber, merge.right);
  }
}

function ensureRepDetailCapacity(worksheet, totalRow, mayorRow, requiredLastDetailRow) {
  if (!Number.isInteger(totalRow) || requiredLastDetailRow < REP_DETAIL_START_ROW) {
    return {
      totalRow,
      mayorRow,
      formatMayorRow: mayorRow,
      templateTotalRow: totalRow,
    };
  }

  const footerStartRow = findRepFooterStartRow(worksheet, totalRow);
  const spacerCount = Math.max(0, totalRow - footerStartRow);
  const desiredTotalRow = Math.max(REP_DETAIL_START_ROW + spacerCount + 1, requiredLastDetailRow + spacerCount + 1);
  const extraRows = Math.max(0, desiredTotalRow - totalRow);
  const styleRow = findRepStyleRow(worksheet, footerStartRow);
  const detailMerges = getRowMergeDefinitions(worksheet, styleRow);
  const spacerMerges = getRowMergeDefinitions(worksheet, footerStartRow);
  const totalMerges = getRowMergeDefinitions(worksheet, totalRow);
  if (extraRows > 0) {
    worksheet.insertRows(footerStartRow, Array.from({ length: extraRows }, () => []), "n");
  }

  const templateSpacerRow = footerStartRow + extraRows;
  const templateTotalRow = totalRow + extraRows;
  const formatMayorRow = Number.isInteger(mayorRow) ? mayorRow + extraRows : mayorRow;
  const desiredSpacerRow = desiredTotalRow - spacerCount;

  clearMergesInRowRange(
    worksheet,
    Math.min(desiredSpacerRow, templateSpacerRow),
    Math.max(templateTotalRow, Number.isInteger(formatMayorRow) ? formatMayorRow : templateTotalRow),
  );

  for (let row = footerStartRow; row < footerStartRow + extraRows; row += 1) {
    copyRowStyle(worksheet, styleRow, row, REP_MAX_COLUMN);
    applyRowMergeDefinitions(worksheet, row, detailMerges);
  }

  copyRowStyle(worksheet, templateSpacerRow, desiredSpacerRow, REP_MAX_COLUMN);
  applyRowMergeDefinitions(worksheet, desiredSpacerRow, spacerMerges);

  copyRowStyle(worksheet, templateTotalRow, desiredTotalRow, REP_MAX_COLUMN);
  applyRowMergeDefinitions(worksheet, desiredTotalRow, totalMerges);

  return {
    totalRow: desiredTotalRow,
    mayorRow: desiredTotalRow + 1,
    formatMayorRow,
    templateTotalRow,
  };
}

function copyCellStyle(sourceCell, targetCell) {
  if (isMergedSlave(targetCell)) {
    return;
  }

  targetCell.style = cloneDeep(sourceCell.style || {});
}

function copyRowStyle(worksheet, sourceRowNumber, targetRowNumber, lastColumn = REP_MAX_COLUMN) {
  const sourceRow = worksheet.getRow(sourceRowNumber);
  const targetRow = worksheet.getRow(targetRowNumber);
  if (sourceRow.height) {
    targetRow.height = sourceRow.height;
  }
  targetRow.hidden = sourceRow.hidden === true;
  targetRow.outlineLevel = sourceRow.outlineLevel || 0;

  for (let column = 1; column <= lastColumn; column += 1) {
    copyCellStyle(sourceRow.getCell(column), targetRow.getCell(column));
  }
}

function cloneCellValuePreservingDates(value) {
  if (value == null) {
    return value;
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getTime());
  }

  if (Array.isArray(value)) {
    return value.map((item) => cloneCellValuePreservingDates(item));
  }

  if (typeof value === "object") {
    const cloned = {};
    for (const [key, nestedValue] of Object.entries(value)) {
      cloned[key] = cloneCellValuePreservingDates(nestedValue);
    }
    return cloned;
  }

  return value;
}

function copyRowValues(worksheet, sourceRowNumber, targetRowNumber, lastColumn) {
  const sourceRow = worksheet.getRow(sourceRowNumber);
  const targetRow = worksheet.getRow(targetRowNumber);
  clearWorksheetRange(worksheet, targetRowNumber, targetRowNumber, 1, lastColumn);

  for (let column = 1; column <= lastColumn; column += 1) {
    const sourceCell = sourceRow.getCell(column);
    const targetCell = targetRow.getCell(column);
    if (isMergedSlave(targetCell)) {
      continue;
    }
    targetCell.value = cloneCellValuePreservingDates(sourceCell.value);
  }
}

function convertSourceCellValue(cell, column) {
  if (!cell) {
    return null;
  }

  if (REP_TEXT_COLUMNS.has(column)) {
    const text = sourceLiteralText(cell);
    return text === "" ? null : text;
  }

  if (cell.f) {
    return {
      formula: cell.f,
      result: cell.v == null ? undefined : cell.v,
    };
  }

  if (cell.v == null) {
    return null;
  }

  return cell.v;
}

function copyRepSourceToTarget(sourceSheet, targetSheet) {
  const sourceLastRow = getSourceLastRow(sourceSheet);
  for (let rowNumber = 1; rowNumber <= sourceLastRow; rowNumber += 1) {
    const row = targetSheet.getRow(rowNumber);
    for (let column = 1; column <= REP_MAX_COLUMN; column += 1) {
      const targetCell = row.getCell(column);
      if (isMergedSlave(targetCell)) {
        continue;
      }

      targetCell.value = convertSourceCellValue(getSourceCell(sourceSheet, rowNumber, column), column);
    }
  }
}

function getRepMayorFormulaMap(key, totalRow) {
  switch (key) {
    case "tyt":
      return { 16: "'MY REP TYT'!I38", 18: "'MY REP TYT'!H53" };
    case "peug":
      return {
        16: "'MY REP PEUG'!I19",
        18: "'MY REP PEUG'!H25",
        25: `X${totalRow}-'NC REP PEUG'!AD9:AE9`,
        26: `X${totalRow}`,
      };
    case "chgn":
      return { 16: "'MY REP CHGN'!J8", 19: "'MY REP CHGN'!I15" };
    case "szk":
      return { 16: "'MY REP SZK'!I31", 18: "'MY REP SZK'!H50" };
    default:
      throw new Error(`Clave no soportada: ${key}`);
  }
}

function applyMayorRow(worksheet, key, totalRow, formatRow) {
  const mayorRow = totalRow + 1;
  if (Number.isInteger(formatRow) && formatRow > 0 && formatRow !== mayorRow) {
    copyRowStyle(worksheet, formatRow, mayorRow, REP_MAX_COLUMN);
  }

  clearWorksheetRange(worksheet, mayorRow, mayorRow, 1, REP_MAX_COLUMN);
  worksheet.getRow(mayorRow).getCell(14).value = "MAYOR";

  const formulas = getRepMayorFormulaMap(key, totalRow);
  for (const [column, formula] of Object.entries(formulas)) {
    worksheet.getRow(mayorRow).getCell(Number(column)).value = { formula, result: 0 };
  }

  return mayorRow;
}

function applyRepSheet(sourceData, workbook, config) {
  const targetSheet = workbook.getWorksheet(config.targetSheet);
  if (!targetSheet) {
    throw new Error(`No existe la hoja requerida en plantilla: ${config.targetSheet}`);
  }

  const oldUsedLastRow = Math.max(targetSheet.rowCount || 0, 1);
  let oldTotalRow = findWorksheetRowContaining(
    targetSheet,
    "TOTAL GENERAL",
    1,
    Math.max(400, oldUsedLastRow + 20),
    REP_MAX_COLUMN,
  );
  let oldMayorRow = findWorksheetRowContaining(
    targetSheet,
    "MAYOR",
    1,
    Math.max(400, oldUsedLastRow + 20),
    REP_MAX_COLUMN,
  );
  const capacity = ensureRepDetailCapacity(
    targetSheet,
    oldTotalRow,
    oldMayorRow,
    findLastPopulatedSourceDetailRow(sourceData.sheet, sourceData.totalRow),
  );
  oldTotalRow = capacity.totalRow;
  oldMayorRow = capacity.mayorRow;
  copyRepSourceToTarget(sourceData.sheet, targetSheet);

  const resolvedTotalRow = findWorksheetRowContaining(
    targetSheet,
    "TOTAL GENERAL",
    1,
    Math.max(400, sourceData.lastRow + 20, oldUsedLastRow + 20),
    REP_MAX_COLUMN,
  ) || oldTotalRow;

  const mayorRow = applyMayorRow(
    targetSheet,
    config.key,
    resolvedTotalRow,
    capacity.formatMayorRow || oldMayorRow || oldTotalRow,
  );
  clearWorksheetColumns(
    targetSheet,
    mayorRow + 1,
    Math.max(oldUsedLastRow, sourceData.lastRow + 10),
    REP_PAYLOAD_COLUMNS,
  );

  const targetPayload = buildWorksheetPayloadSignature(targetSheet, resolvedTotalRow);
  if (targetPayload.rows !== sourceData.payload.rows) {
    throw new Error(
      `La hoja ${config.targetSheet} no conserva los datos del archivo subido. Filas fuente=${sourceData.payload.rows}, filas salida=${targetPayload.rows}.`,
    );
  }

  return { rowCount: targetPayload.rows };
}
function roundAmount(value, decimals = 2) {
  const factor = 10 ** decimals;
  return Math.round((Number(value) + Number.EPSILON) * factor) / factor;
}

function compareTextAscending(left, right) {
  const leftText = normalizeText(left).toUpperCase();
  const rightText = normalizeText(right).toUpperCase();
  if (leftText < rightText) {
    return -1;
  }
  if (leftText > rightText) {
    return 1;
  }
  return 0;
}

function compareDateAscending(left, right) {
  const leftTime = left instanceof Date && !Number.isNaN(left.getTime()) ? left.getTime() : Number.POSITIVE_INFINITY;
  const rightTime = right instanceof Date && !Number.isNaN(right.getTime()) ? right.getTime() : Number.POSITIVE_INFINITY;
  if (leftTime < rightTime) {
    return -1;
  }
  if (leftTime > rightTime) {
    return 1;
  }
  return 0;
}

function compareSeatAscending(left, right) {
  const leftText = normalizeText(left);
  const rightText = normalizeText(right);
  const leftNumeric = /^\d+$/.test(leftText) ? Number(leftText) : Number.NaN;
  const rightNumeric = /^\d+$/.test(rightText) ? Number(rightText) : Number.NaN;

  if (Number.isFinite(leftNumeric) && Number.isFinite(rightNumeric) && leftNumeric !== rightNumeric) {
    return leftNumeric - rightNumeric;
  }

  return compareTextAscending(leftText, rightText);
}

function comparePostingEntries(left, right) {
  return compareDateAscending(left.dateValue, right.dateValue)
    || compareSeatAscending(left.seat, right.seat)
    || compareTextAscending(left.detail, right.detail)
    || compareTextAscending(left.account, right.account)
    || compareTextAscending(left.side, right.side);
}

function buildPostingEntries(groupMap, side) {
  return Object.entries(groupMap || {})
    .map(([groupKey, group]) => {
      const [account, seat, detail] = groupKey.split("|");
      return {
        account,
        seat: seat || "",
        detail: detail || "",
        dateValue: group?.dateValue || null,
        dateText: group?.dateText || "",
        amount: roundAmount(group?.amount || 0, 2),
        side,
      };
    })
    .filter((entry) => entry.account !== "" && entry.detail !== "" && Math.abs(entry.amount) >= 0.0000001)
    .sort(comparePostingEntries);
}

function getSectionLastColumn(layout) {
  return Math.max(
    layout.dateColumn,
    layout.seatColumn,
    layout.detailColumn,
    layout.debitColumn,
    layout.creditColumn,
    layout.saldoColumn,
  );
}

function getTemplateAccountBlocks(templateWorksheet, section) {
  const blocks = [];
  let current = null;

  for (let row = section.startRow; row <= section.endRow; row += 1) {
    const account = getWorksheetCellText(templateWorksheet, row, 1);
    if (account === "") {
      if (current) {
        blocks.push(current);
        current = null;
      }
      continue;
    }

    if (!current || current.account !== account) {
      if (current) {
        blocks.push(current);
      }
      current = { account, startRow: row, endRow: row };
      continue;
    }

    current.endRow = row;
  }

  if (current) {
    blocks.push(current);
  }

  return blocks;
}

function getSectionOpeningBalances(templateWorksheet, layout, section) {
  const balances = {};
  for (const block of getTemplateAccountBlocks(templateWorksheet, section)) {
    const templateDebit = getWorksheetCellNumber(templateWorksheet, block.startRow, layout.debitColumn);
    const templateCredit = getWorksheetCellNumber(templateWorksheet, block.startRow, layout.creditColumn);
    const templateSaldo = getWorksheetCellNumber(templateWorksheet, block.startRow, layout.saldoColumn);
    balances[block.account] = roundAmount(templateSaldo - templateDebit + templateCredit, 2);
  }
  return balances;
}

function getSectionStaticColumns(layout) {
  const dynamicColumns = new Set([
    layout.dateColumn,
    layout.seatColumn,
    layout.detailColumn,
    layout.debitColumn,
    layout.creditColumn,
    layout.saldoColumn,
  ]);

  const columns = [];
  for (let column = 1; column <= getSectionLastColumn(layout); column += 1) {
    if (!dynamicColumns.has(column)) {
      columns.push(column);
    }
  }
  return columns;
}

function getSectionStaticTemplates(templateWorksheet, layout, section) {
  const templates = {};
  const staticColumns = getSectionStaticColumns(layout);
  for (const block of getTemplateAccountBlocks(templateWorksheet, section)) {
    templates[block.account] = {};
    for (const column of staticColumns) {
      templates[block.account][column] = cloneDeep(templateWorksheet.getRow(block.startRow).getCell(column).value);
    }
  }
  return templates;
}

function setFormulaResult(cell, result, formulaOverride = null) {
  const currentValue = cell.value;
  const formula = formulaOverride || (currentValue && typeof currentValue === "object" ? currentValue.formula : null);
  if (!formula) {
    cell.value = result;
    return;
  }

  const nextValue = {
    formula,
    result,
  };

  if (currentValue && typeof currentValue === "object" && currentValue.shareType) {
    nextValue.shareType = currentValue.shareType;
  }

  if (currentValue && typeof currentValue === "object" && currentValue.ref) {
    nextValue.ref = currentValue.ref;
  }

  cell.value = nextValue;
}

function sumWorksheetColumn(worksheet, startRow, endRow, column) {
  let total = 0;
  for (let row = startRow; row <= endRow; row += 1) {
    total += getWorksheetCellNumber(worksheet, row, column);
  }
  return roundAmount(total, 2);
}

function addGroupedAmount(groups, groupKey, amount, dateValue, dateText, seat, detail) {
  if (Math.abs(amount) < 0.0000001) {
    return;
  }

  if (!groups[groupKey]) {
    groups[groupKey] = {
      amount: 0,
      dateValue,
      dateText,
      seat,
      detail,
    };
  }

  const entry = groups[groupKey];
  entry.amount = roundAmount(entry.amount + amount, 6);
  if (dateValue != null && dateValue !== "") {
    if (entry.dateValue == null || entry.dateValue === "") {
      entry.dateValue = dateValue;
    } else if (entry.dateValue instanceof Date && dateValue instanceof Date) {
      if (dateValue < entry.dateValue) {
        entry.dateValue = dateValue;
        entry.dateText = dateText;
      }
    }
  }
  if (normalizeText(entry.dateText) === "" && normalizeText(dateText) !== "") {
    entry.dateText = dateText;
  }
}

function getRepDetailName(key, agency) {
  switch (key) {
    case "tyt":
      return "MOD. REPUESTOS REP01";
    case "peug":
      return "MOD. REPUESTOS REP06";
    case "chgn":
      return "MOD. REPUESTOS REP05";
    case "szk":
      return normalizeText(agency) === "08" ? "MOD. REPUESTOS REP08" : "MOD. REPUESTOS REP07";
    default:
      throw new Error(`Clave no soportada para detalle REP: ${key}`);
  }
}

function getMayorIvaDetailsForKey(key) {
  switch (key) {
    case "tyt":
      return ["MOD. REPUESTOS REP01"];
    case "peug":
      return ["MOD. REPUESTOS REP06"];
    case "chgn":
      return ["MOD. REPUESTOS REP05"];
    case "szk":
      return ["MOD. REPUESTOS REP07", "MOD. REPUESTOS REP08"];
    default:
      return [];
  }
}

function getPostingAccount(key, category, form) {
  const normalizedForm = normalizeText(form).toUpperCase();
  const mappings = {
    tyt: {
      sales: { CONTADO: "04.01.01.01.0001", CREDITO: "04.01.01.01.0003" },
      discount: { CONTADO: "04.01.01.01.0005", CREDITO: "04.01.01.01.0007" },
    },
    peug: {
      sales: { CONTADO: "04.01.01.03.0001", CREDITO: "04.01.01.03.0003" },
      discount: { CONTADO: "04.01.01.03.0005", CREDITO: "04.01.01.03.0007" },
    },
    chgn: {
      sales: { CONTADO: "04.01.01.02.0001", CREDITO: "04.01.01.02.0003" },
      discount: { CONTADO: "04.01.01.02.0005" },
    },
    szk: {
      sales: { CONTADO: "04.01.01.04.0001", CREDITO: "04.01.01.04.0003" },
      discount: { CONTADO: "04.01.01.04.0005", CREDITO: "04.01.01.04.0007" },
    },
  };

  return mappings[key]?.[category]?.[normalizedForm] || "";
}

function getDevolAccount(key, form) {
  const normalizedForm = normalizeText(form).toUpperCase();
  const mappings = {
    tyt: { CONTADO: "04.01.01.01.0009", CREDITO: "04.01.01.01.0011" },
    peug: { CONTADO: "04.01.01.03.0009", CREDITO: "04.01.01.03.0011" },
    chgn: { CONTADO: "04.01.01.02.0009", CREDITO: "04.01.01.02.0011" },
    szk: { CONTADO: "04.01.01.04.0009", CREDITO: "04.01.01.04.0011" },
  };

  return mappings[key]?.[normalizedForm] || "";
}

function inferFormFromDescription(description) {
  const text = normalizeText(description).toUpperCase();
  if (text.includes("CREDITO")) {
    return "CREDITO";
  }
  if (text.includes("CONTADO")) {
    return "CONTADO";
  }
  return "";
}

function getNcSheetName(key) {
  switch (key) {
    case "tyt":
      return "NC REP TYT";
    case "peug":
      return "NC REP PEUG";
    case "chgn":
      return "NC REP CHGN";
    case "szk":
      return "NC REP SZK";
    default:
      return "";
  }
}

function buildRepPostingGroups(sourceSheet, key) {
  const sales = {};
  const discount = {};
  const vat = {};
  const totalRow = findSourceTotalRow(sourceSheet);

  for (let row = REP_DETAIL_START_ROW; row < totalRow; row += 1) {
    const seat = sourceCellText(sourceSheet, row, 38);
    if (seat === "" || seat === "ASIENTO") {
      continue;
    }

    const form = sourceCellText(sourceSheet, row, 40).toUpperCase();
    if (form !== "CONTADO" && form !== "CREDITO") {
      continue;
    }

    const agency = sourceCellText(sourceSheet, row, 39);
    const detail = getRepDetailName(key, agency);
    const rawDateValue = getSourceCell(sourceSheet, row, 3)?.v ?? null;
    const dateText = sourceCellText(sourceSheet, row, 3);
    const dateValue = convertDateValue(rawDateValue, dateText);
    const salesAmount = sourceCellNumber(sourceSheet, row, 18);
    const discountAmount = sourceCellNumber(sourceSheet, row, 20);
    const vatAmount = sourceCellNumber(sourceSheet, row, 26);

    const salesAccount = getPostingAccount(key, "sales", form);
    if (salesAccount !== "") {
      addGroupedAmount(sales, `${salesAccount}|${seat}|${detail}`, salesAmount, dateValue, dateText, seat, detail);
    }

    const discountAccount = getPostingAccount(key, "discount", form);
    if (discountAccount !== "") {
      addGroupedAmount(
        discount,
        `${discountAccount}|${seat}|${detail}`,
        discountAmount,
        dateValue,
        dateText,
        seat,
        detail,
      );
    }

    addGroupedAmount(vat, `${seat}|${detail}`, vatAmount, dateValue, dateText, seat, detail);
  }

  return { sales, discount, vat };
}

function buildRepLookup(sourceSheet, key) {
  const lookup = new Map();
  const totalRow = findSourceTotalRow(sourceSheet);

  for (let row = REP_DETAIL_START_ROW; row < totalRow; row += 1) {
    const document = normalizeDocNumber(sourceCellText(sourceSheet, row, 5));
    if (document === "" || document.toUpperCase().startsWith("ANULAD")) {
      continue;
    }

    if (lookup.has(document)) {
      continue;
    }

    const seat = sourceCellText(sourceSheet, row, 38);
    const form = normalizeText(sourceCellText(sourceSheet, row, 40)).toUpperCase();
    const agency = sourceCellText(sourceSheet, row, 39);
    const detail = getRepDetailName(key, agency);
    const rawDateValue = getSourceCell(sourceSheet, row, 3)?.v ?? null;
    const dateText = sourceCellText(sourceSheet, row, 3);
    const dateValue = convertDateValue(rawDateValue, dateText);

    lookup.set(document, {
      seat,
      form,
      agency,
      detail,
      dateValue,
      dateText,
    });
  }

  return lookup;
}

function buildNcTotals(ncWorksheet, key) {
  const discountTotals = {};
  const devolTotals = {};
  const vatTotals = {};

  if (!ncWorksheet) {
    return { discountTotals, devolTotals, vatTotals };
  }

  const totalRow = findWorksheetRowContaining(ncWorksheet, "TOTAL GENERAL", 1, 200, NC_MAX_COLUMN)
    || NC_DETAIL_START_ROW;

  for (let row = NC_DETAIL_START_ROW; row < totalRow; row += 1) {
    const annulText = normalizeText(getWorksheetCellText(ncWorksheet, row, 1))
      || normalizeText(getWorksheetCellText(ncWorksheet, row, 2));
    if (annulText.toUpperCase().startsWith("SI") || annulText.toUpperCase().startsWith("ANULAD")) {
      continue;
    }

    const document = normalizeDocNumber(getWorksheetCellText(ncWorksheet, row, 10));
    if (document === "") {
      continue;
    }

    const description = getWorksheetCellText(ncWorksheet, row, 11);
    const form = inferFormFromDescription(description);

    const dateCell = ncWorksheet.getRow(row).getCell(4).value;
    const dateText = getWorksheetCellText(ncWorksheet, row, 4);
    const dateValue = convertDateValue(dateCell, dateText);

    const subTotal = getWorksheetCellNumber(ncWorksheet, row, 23);
    const discountAmount = getWorksheetCellNumber(ncWorksheet, row, 24);
    const vatAmount = getWorksheetCellNumber(ncWorksheet, row, 30);

    const discountAccount = getPostingAccount(key, "discount", form);
    if (discountAccount !== "") {
      addGroupedAmount(
        discountTotals,
        discountAccount,
        discountAmount,
        dateValue,
        dateText,
        "",
        "",
      );
    }

    const devolAccount = getDevolAccount(key, form);
    if (devolAccount !== "") {
      addGroupedAmount(
        devolTotals,
        devolAccount,
        subTotal,
        dateValue,
        dateText,
        "",
        "",
      );
    }

    addGroupedAmount(vatTotals, "vat", vatAmount, dateValue, dateText, "", "");
  }

  return { discountTotals, devolTotals, vatTotals };
}

function getTemplateRowDate(worksheet, rowNumber, dateColumn) {
  const dateCell = worksheet.getRow(rowNumber).getCell(dateColumn).value;
  const dateText = getWorksheetCellText(worksheet, rowNumber, dateColumn);
  return { dateValue: convertDateValue(dateCell, dateText), dateText };
}

function sameCalendarDate(left, right) {
  if (!(left instanceof Date) || Number.isNaN(left.getTime()) || !(right instanceof Date) || Number.isNaN(right.getTime())) {
    return false;
  }

  return left.getFullYear() === right.getFullYear()
    && left.getMonth() === right.getMonth()
    && left.getDate() === right.getDate();
}

function findTemplateNcContext(templateWorkbook, key, form, subTotal, vatAmount, ncDateValue, fallbackDateText = "") {
  const layout = MY_LAYOUTS[key];
  if (!layout) {
    return null;
  }

  const templateWorksheet = templateWorkbook.getWorksheet(layout.mySheetName);
  if (!templateWorksheet) {
    return null;
  }

  const devolAccount = getDevolAccount(key, form);
  const devolSection = getDevolSectionRange(templateWorksheet, layout);
  const subTotalAmount = roundAmount(Number(subTotal || 0), 2);
  const vatTarget = roundAmount(Number(vatAmount || 0), 2);

  const candidates = [];
  if (devolAccount && devolSection) {
    for (let row = devolSection.startRow; row <= devolSection.endRow; row += 1) {
      if (getWorksheetCellText(templateWorksheet, row, 1) !== devolAccount) {
        continue;
      }

      const rowAmount = roundAmount(getWorksheetCellNumber(templateWorksheet, row, layout.debitColumn), 2);
      if (Math.abs(rowAmount - subTotalAmount) >= 0.0000001) {
        continue;
      }

      const rowDate = getTemplateRowDate(templateWorksheet, row, layout.dateColumn);
      const seat = getWorksheetCellText(templateWorksheet, row, layout.seatColumn);
      const detail = getWorksheetCellText(templateWorksheet, row, layout.detailColumn);
      if (seat === "" || detail === "") {
        continue;
      }

      const sameDate = sameCalendarDate(rowDate.dateValue, ncDateValue);
      candidates.push({
        seat,
        detail,
        dateValue: rowDate.dateValue || ncDateValue || null,
        dateText: rowDate.dateText || fallbackDateText,
        score: sameDate ? 0 : 1,
      });
    }
  }

  if (candidates.length > 0) {
    candidates.sort((left, right) => left.score - right.score);
    return candidates[0];
  }

  const mayorWorksheet = templateWorkbook.getWorksheet("MAYOR IVA");
  if (!mayorWorksheet || Math.abs(vatTarget) < 0.0000001) {
    return null;
  }

  const mayorCandidates = [];
  for (let row = MAYOR_IVA_START_ROW; row <= MAYOR_IVA_END_ROW; row += 1) {
    const type = getWorksheetCellText(mayorWorksheet, row, 5).toUpperCase();
    if (type !== "REPTO") {
      continue;
    }

    const debit = roundAmount(getWorksheetCellNumber(mayorWorksheet, row, 8), 2);
    if (Math.abs(debit - vatTarget) >= 0.0000001) {
      continue;
    }

    const seat = getWorksheetCellText(mayorWorksheet, row, 6);
    const detail = getWorksheetCellText(mayorWorksheet, row, 7);
    if (seat === "" || detail === "") {
      continue;
    }

    const rowDate = getTemplateRowDate(mayorWorksheet, row, 4);
    const sameDate = sameCalendarDate(rowDate.dateValue, ncDateValue);
    mayorCandidates.push({
      seat,
      detail,
      dateValue: rowDate.dateValue || ncDateValue || null,
      dateText: rowDate.dateText || fallbackDateText,
      score: sameDate ? 0 : 1,
    });
  }

  if (mayorCandidates.length === 0) {
    return null;
  }

  mayorCandidates.sort((left, right) => left.score - right.score);
  return mayorCandidates[0];
}

function buildGroupsFromAccountTotals(templateWorksheet, layout, groupMap, totalsByAccount) {
  const result = {};
  const groupsByAccount = {};

  for (const [groupKey, group] of Object.entries(groupMap)) {
    const [account, seat, detail] = groupKey.split("|");
    if (!groupsByAccount[account]) {
      groupsByAccount[account] = [];
    }
    groupsByAccount[account].push({ groupKey, group, seat, detail });
  }

  for (const [account, groups] of Object.entries(groupsByAccount)) {
    const totalEntry = totalsByAccount[account] || null;
    const targetTotal = totalEntry ? Number(totalEntry.amount) : 0;
    const baseTotal = groups.reduce((sum, info) => sum + Number(info.group.baseTotal || 0), 0);
    let remaining = roundAmount(targetTotal, 6);

    groups.forEach((info, index) => {
      const isLast = index === groups.length - 1;
      const scaledAmount = isLast
        ? roundAmount(remaining, 2)
        : baseTotal !== 0
          ? roundAmount((targetTotal * info.group.baseTotal) / baseTotal, 2)
          : 0;

      if (!isLast) {
        remaining = roundAmount(remaining - scaledAmount, 6);
      }

      const rowNumber = info.group.rows[0]?.rowNumber;
      const templateDate = rowNumber
        ? getTemplateRowDate(templateWorksheet, rowNumber, layout.dateColumn)
        : { dateValue: null, dateText: "" };
      const dateValue = templateDate.dateValue || totalEntry?.dateValue || null;
      const dateText = templateDate.dateText || totalEntry?.dateText || "";

      result[info.groupKey] = {
        amount: scaledAmount,
        dateValue,
        dateText,
        seat: info.seat || "",
        detail: info.detail || "",
      };
    });
  }

  return result;
}

function buildGroupsFromTotal(templateWorksheet, groupMap, totalEntry, dateColumn = 4) {
  const result = {};
  const groups = Object.entries(groupMap);
  const targetTotal = totalEntry ? Number(totalEntry.amount) : 0;
  const baseTotal = groups.reduce((sum, [, group]) => sum + Number(group.baseTotal || 0), 0);
  let remaining = roundAmount(targetTotal, 6);

  groups.forEach(([groupKey, group], index) => {
    const isLast = index === groups.length - 1;
    const scaledAmount = isLast
      ? roundAmount(remaining, 2)
      : baseTotal !== 0
        ? roundAmount((targetTotal * group.baseTotal) / baseTotal, 2)
        : 0;

    if (!isLast) {
      remaining = roundAmount(remaining - scaledAmount, 6);
    }

    const [seat, detail] = groupKey.split("|");
    const rowNumber = group.rows[0]?.rowNumber;
    const templateDate = rowNumber
      ? getTemplateRowDate(templateWorksheet, rowNumber, dateColumn)
      : { dateValue: null, dateText: "" };
    const dateValue = templateDate.dateValue || totalEntry?.dateValue || null;
    const dateText = templateDate.dateText || totalEntry?.dateText || "";

    result[groupKey] = {
      amount: scaledAmount,
      dateValue,
      dateText,
      seat: seat || "",
      detail: detail || "",
    };
  });

  return result;
}

function filterMayorIvaGroupsByDetail(groupMap, allowedDetails) {
  if (!allowedDetails || allowedDetails.length === 0) {
    return {};
  }
  const allowed = new Set(allowedDetails);
  const filtered = {};
  for (const [groupKey, group] of Object.entries(groupMap)) {
    const [, detail] = groupKey.split("|");
    if (allowed.has(detail)) {
      filtered[groupKey] = group;
    }
  }
  return filtered;
}

function buildNcPostingGroups(templateWorkbook, key, repLookup = new Map()) {
  const discountCredit = {};
  const devol = {};
  const vat = {};

  const ncSheetName = getNcSheetName(key);
  const ncWorksheet = ncSheetName ? templateWorkbook.getWorksheet(ncSheetName) : null;
  if (!ncWorksheet) {
    return { discountCredit, devol, vat };
  }

  const totalRow = findWorksheetRowContaining(ncWorksheet, "TOTAL GENERAL", 1, 200, NC_MAX_COLUMN)
    || NC_DETAIL_START_ROW;

  for (let row = NC_DETAIL_START_ROW; row < totalRow; row += 1) {
    const annulText = normalizeText(getWorksheetCellText(ncWorksheet, row, 1))
      || normalizeText(getWorksheetCellText(ncWorksheet, row, 2));
    if (annulText.toUpperCase().startsWith("SI") || annulText.toUpperCase().startsWith("ANULAD")) {
      continue;
    }

    const document = normalizeDocNumber(getWorksheetCellText(ncWorksheet, row, 10));
    if (document === "") {
      continue;
    }

    const repInfo = repLookup.get(document) || null;
    const description = getWorksheetCellText(ncWorksheet, row, 11);
    const form = inferFormFromDescription(description) || repInfo?.form || "";
    const dateCell = ncWorksheet.getRow(row).getCell(4).value;
    const subTotal = getWorksheetCellNumber(ncWorksheet, row, 23);
    const discountAmount = getWorksheetCellNumber(ncWorksheet, row, 24);
    const vatAmount = getWorksheetCellNumber(ncWorksheet, row, 30);
    const ncDateText = getWorksheetCellText(ncWorksheet, row, 4);
    const ncDateValue = convertDateValue(dateCell, ncDateText);
    const templateInfo = findTemplateNcContext(templateWorkbook, key, form, subTotal, vatAmount, ncDateValue, ncDateText);
    const dateText = templateInfo?.dateText || repInfo?.dateText || ncDateText;
    const dateValue = templateInfo?.dateValue || repInfo?.dateValue || ncDateValue;
    const seat = templateInfo?.seat || repInfo?.seat || "";
    const detail = templateInfo?.detail || repInfo?.detail || "";

    const discountAccount = getPostingAccount(key, "discount", form);
    if (discountAccount !== "" && seat !== "" && detail !== "") {
      addGroupedAmount(
        discountCredit,
        `${discountAccount}|${seat}|${detail}`,
        discountAmount,
        dateValue,
        dateText,
        seat,
        detail,
      );
    }

    const devolAccount = getDevolAccount(key, form);
    if (devolAccount !== "" && seat !== "" && detail !== "") {
      addGroupedAmount(
        devol,
        `${devolAccount}|${seat}|${detail}`,
        subTotal,
        dateValue,
        dateText,
        seat,
        detail,
      );
    }

    if (seat !== "" && detail !== "") {
      addGroupedAmount(vat, `${seat}|${detail}`, vatAmount, dateValue, dateText, seat, detail);
    }
  }

  return { discountCredit, devol, vat };
}

function getTemplateSectionGroups(worksheet, layout, section) {
  const debitGroups = {};
  const creditGroups = {};

  for (let row = section.startRow; row <= section.endRow; row += 1) {
    const amount = getWorksheetCellNumber(worksheet, row, section.amountColumn);
    const opposite = getWorksheetCellNumber(worksheet, row, section.oppositeColumn);
    const account = getWorksheetCellText(worksheet, row, 1);
    const seat = getWorksheetCellText(worksheet, row, layout.seatColumn);
    const detail = getWorksheetCellText(worksheet, row, layout.detailColumn);

    if (account === "" || seat === "" || detail === "") {
      continue;
    }

    if (Math.abs(amount) >= 0.0000001 && Math.abs(opposite) < 0.0000001) {
      const groupKey = `${account}|${seat}|${detail}`;
      if (!debitGroups[groupKey]) {
        debitGroups[groupKey] = { baseTotal: 0, rows: [] };
      }
      debitGroups[groupKey].baseTotal = roundAmount(debitGroups[groupKey].baseTotal + amount, 6);
      debitGroups[groupKey].rows.push({ rowNumber: row, baseAmount: amount });
      continue;
    }

    if (Math.abs(opposite) >= 0.0000001 && Math.abs(amount) < 0.0000001) {
      const groupKey = `${account}|${seat}|${detail}`;
      if (!creditGroups[groupKey]) {
        creditGroups[groupKey] = { baseTotal: 0, rows: [] };
      }
      creditGroups[groupKey].baseTotal = roundAmount(creditGroups[groupKey].baseTotal + opposite, 6);
      creditGroups[groupKey].rows.push({ rowNumber: row, baseAmount: opposite });
    }
  }

  return { debit: debitGroups, credit: creditGroups };
}
function parseDateText(text) {
  const normalized = normalizeText(text).toUpperCase();
  if (normalized === "") {
    return null;
  }

  let match = normalized.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]) - 1;
    let year = Number(match[3]);
    if (year < 100) {
      year += year >= 70 ? 1900 : 2000;
    }

    const date = new Date(year, month, day);
    if (!Number.isNaN(date.getTime())) {
      return date;
    }
  }

  match = normalized.match(/^(\d{1,2})[\/-]([A-Z]{3})[\/-](\d{2,4})$/);
  if (match) {
    const day = Number(match[1]);
    const month = MONTHS[match[2]];
    let year = Number(match[3]);
    if (year < 100) {
      year += year >= 70 ? 1900 : 2000;
    }

    if (month != null) {
      const date = new Date(year, month, day);
      if (!Number.isNaN(date.getTime())) {
        return date;
      }
    }
  }

  const native = new Date(normalized);
  return Number.isNaN(native.getTime()) ? null : native;
}

function convertDateValue(value, fallbackText = "") {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === "number" && Number.isFinite(value) && XLSX.SSF?.parse_date_code) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed && parsed.y && parsed.m && parsed.d) {
      return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H || 0, parsed.M || 0, parsed.S || 0);
    }
  }

  const parsedText = parseDateText(normalizeText(value) || normalizeText(fallbackText));
  if (parsedText) {
    return parsedText;
  }

  const text = normalizeText(value) || normalizeText(fallbackText);
  return text === "" ? null : text;
}

function setDateCellValue(worksheet, row, column, value, fallbackText = "") {
  worksheet.getRow(row).getCell(column).value = convertDateValue(value, fallbackText);
}

function formatDateDDMMYYYY(dateValue) {
  const day = String(dateValue.getDate()).padStart(2, "0");
  const month = String(dateValue.getMonth() + 1).padStart(2, "0");
  const year = String(dateValue.getFullYear());
  return `${day}/${month}/${year}`;
}

function getSourceDateRange(sourceSheet) {
  const totalRow = findSourceTotalRow(sourceSheet);
  let minDate = null;
  let maxDate = null;

  for (let row = REP_DETAIL_START_ROW; row < totalRow; row += 1) {
    const rawValue = getSourceCell(sourceSheet, row, 3)?.v ?? null;
    const rawText = sourceCellText(sourceSheet, row, 3);
    const resolved = convertDateValue(rawValue, rawText);
    if (!(resolved instanceof Date) || Number.isNaN(resolved.getTime())) {
      continue;
    }

    if (!minDate || resolved < minDate) {
      minDate = resolved;
    }
    if (!maxDate || resolved > maxDate) {
      maxDate = resolved;
    }
  }

  return { minDate, maxDate };
}

function updateSheetDateRange(worksheet, minDate, maxDate) {
  if (!worksheet || !(minDate instanceof Date) || !(maxDate instanceof Date)) {
    return;
  }

  const minLabel = `FECHA INICIAL: ${formatDateDDMMYYYY(minDate)}`;
  const maxLabel = `FECHA FINAL: ${formatDateDDMMYYYY(maxDate)}`;

  for (let row = 1; row <= 12; row += 1) {
    const currentRow = worksheet.getRow(row);
    for (let column = 1; column <= 14; column += 1) {
      const cell = currentRow.getCell(column);
      if (isMergedSlave(cell) || cell.value == null) {
        continue;
      }

      const text = normalizeText(cell.value).toUpperCase();
      if (text.includes("FECHA INICIAL")) {
        cell.value = minLabel;
      } else if (text.includes("FECHA FINAL")) {
        cell.value = maxLabel;
      }
    }
  }
}

function getWorksheetDateRange(
  worksheet,
  startRow = REP_DETAIL_START_ROW,
  dateColumn = 3,
  totalNeedle = "TOTAL GENERAL",
  lastColumn = REP_MAX_COLUMN,
) {
  if (!worksheet) {
    return { minDate: null, maxDate: null };
  }

  const totalRow = findWorksheetRowContaining(
    worksheet,
    totalNeedle,
    1,
    Math.max(400, (worksheet.rowCount || 0) + 20),
    lastColumn,
  ) || Math.max(worksheet.rowCount || 0, startRow);

  let minDate = null;
  let maxDate = null;
  for (let row = startRow; row < totalRow; row += 1) {
    const dateCell = worksheet.getRow(row).getCell(dateColumn).value;
    const dateText = getWorksheetCellText(worksheet, row, dateColumn);
    const resolved = convertDateValue(dateCell, dateText);
    if (!(resolved instanceof Date) || Number.isNaN(resolved.getTime())) {
      continue;
    }

    if (!minDate || resolved < minDate) {
      minDate = resolved;
    }
    if (!maxDate || resolved > maxDate) {
      maxDate = resolved;
    }
  }

  return { minDate, maxDate };
}

function buildCombinedDateRange(sourceDataList) {
  let minDate = null;
  let maxDate = null;

  for (const sourceData of sourceDataList || []) {
    const range = sourceData?.dateRange || null;
    if (!(range?.minDate instanceof Date) || Number.isNaN(range.minDate.getTime())) {
      continue;
    }
    if (!(range?.maxDate instanceof Date) || Number.isNaN(range.maxDate.getTime())) {
      continue;
    }

    if (!minDate || range.minDate < minDate) {
      minDate = range.minDate;
    }
    if (!maxDate || range.maxDate > maxDate) {
      maxDate = range.maxDate;
    }
  }

  return { minDate, maxDate };
}

function getDateMonthKey(dateValue) {
  if (!(dateValue instanceof Date) || Number.isNaN(dateValue.getTime())) {
    return "";
  }

  return `${dateValue.getFullYear()}-${String(dateValue.getMonth() + 1).padStart(2, "0")}`;
}

function getMonthStartDate(dateValue) {
  if (!(dateValue instanceof Date) || Number.isNaN(dateValue.getTime())) {
    return null;
  }

  return new Date(dateValue.getFullYear(), dateValue.getMonth(), 1);
}

function shouldUseTemplateCarryover(templateWorkbook, sourceDataList) {
  const sourceRange = buildCombinedDateRange(sourceDataList);
  const sourceKey = getDateMonthKey(sourceRange.minDate);
  if (!sourceKey) {
    return true;
  }

  for (const config of SHEET_CONFIGS) {
    const worksheet = templateWorkbook.getWorksheet(config.targetSheet);
    const templateRange = getWorksheetDateRange(worksheet, REP_DETAIL_START_ROW, 3, "TOTAL GENERAL", REP_MAX_COLUMN);
    const templateKey = getDateMonthKey(templateRange.minDate);
    if (!templateKey) {
      continue;
    }
    return templateKey === sourceKey;
  }

  return true;
}

function applySectionScaling(outputWorksheet, templateWorksheet, layout, section, sourceGroups) {
  const lastColumn = getSectionLastColumn(layout);
  for (let row = section.startRow; row <= section.endRow; row += 1) {
    clearWorksheetRange(outputWorksheet, row, row, 1, lastColumn);
  }

  const hasSplitGroups = sourceGroups && (Object.prototype.hasOwnProperty.call(sourceGroups, "debit")
    || Object.prototype.hasOwnProperty.call(sourceGroups, "credit"));
  const debitGroups = hasSplitGroups ? (sourceGroups.debit || {}) : (sourceGroups || {});
  const creditGroups = hasSplitGroups ? (sourceGroups.credit || {}) : {};
  const staticTemplates = getSectionStaticTemplates(templateWorksheet, layout, section);

  const accountOrder = new Map(
    getTemplateAccountBlocks(templateWorksheet, section).map((block, index) => [block.account, index]),
  );
  const entries = [
    ...buildPostingEntries(debitGroups, "debit"),
    ...buildPostingEntries(creditGroups, "credit"),
  ].sort((left, right) => {
    const leftOrder = accountOrder.has(left.account) ? accountOrder.get(left.account) : Number.POSITIVE_INFINITY;
    const rightOrder = accountOrder.has(right.account) ? accountOrder.get(right.account) : Number.POSITIVE_INFINITY;
    return (leftOrder - rightOrder) || comparePostingEntries(left, right);
  });

  const capacity = section.endRow - section.startRow + 1;
  if (entries.length > capacity) {
    throw new Error(`La hoja ${layout.mySheetName} no tiene filas suficientes en la sección ${section.startRow}-${section.endRow}. Capacidad=${capacity}, requeridas=${entries.length}.`);
  }

  const activeRows = new Set();
  entries.forEach((entry, index) => {
    const rowNumber = section.startRow + index;
    const row = outputWorksheet.getRow(rowNumber);
    activeRows.add(rowNumber);
    const staticValues = staticTemplates[entry.account] || {};
    for (const [column, value] of Object.entries(staticValues)) {
      row.getCell(Number(column)).value = cloneDeep(value);
    }
    row.getCell(1).value = entry.account;
    setDateCellValue(outputWorksheet, rowNumber, layout.dateColumn, entry.dateValue, entry.dateText);
    row.getCell(layout.seatColumn).value = entry.seat;
    row.getCell(layout.detailColumn).value = entry.detail;
    row.getCell(layout.debitColumn).value = entry.side === "debit" ? entry.amount : 0;
    row.getCell(layout.creditColumn).value = entry.side === "credit" ? entry.amount : 0;
  });

  return activeRows;
}

function recalculateSectionSaldo(
  outputWorksheet,
  templateWorksheet,
  layout,
  section,
  activeRows = new Set(),
  options = {},
) {
  const useTemplateOpeningBalance = options.useTemplateOpeningBalance !== false;
  const openingBalances = useTemplateOpeningBalance
    ? getSectionOpeningBalances(templateWorksheet, layout, section)
    : {};
  let runningBalance = 0;
  let currentAccount = "";

  for (let row = section.startRow; row <= section.endRow; row += 1) {
    if (!activeRows.has(row)) {
      outputWorksheet.getRow(row).getCell(layout.saldoColumn).value = null;
      continue;
    }

    const account = getWorksheetCellText(outputWorksheet, row, 1);
    if (account === "") {
      outputWorksheet.getRow(row).getCell(layout.saldoColumn).value = null;
      continue;
    }

    if (account !== currentAccount) {
      currentAccount = account;
      runningBalance = useTemplateOpeningBalance ? (openingBalances[account] || 0) : 0;
    }

    const debit = getWorksheetCellNumber(outputWorksheet, row, layout.debitColumn);
    const credit = getWorksheetCellNumber(outputWorksheet, row, layout.creditColumn);
    runningBalance = roundAmount(runningBalance + debit - credit, 2);
    outputWorksheet.getRow(row).getCell(layout.saldoColumn).value = runningBalance;
  }
}

function countSectionEntries(sourceGroups) {
  const hasSplitGroups = sourceGroups && (Object.prototype.hasOwnProperty.call(sourceGroups, "debit")
    || Object.prototype.hasOwnProperty.call(sourceGroups, "credit"));
  const debitGroups = hasSplitGroups ? (sourceGroups.debit || {}) : (sourceGroups || {});
  const creditGroups = hasSplitGroups ? (sourceGroups.credit || {}) : {};
  return buildPostingEntries(debitGroups, "debit").length + buildPostingEntries(creditGroups, "credit").length;
}

function prepareSectionTotalRow(outputWorksheet, layout, originalSection, adjustedSection) {
  const originalTotalRow = originalSection.endRow + 1;
  const adjustedTotalRow = adjustedSection.endRow + 1;
  if (originalTotalRow === adjustedTotalRow) {
    return;
  }

  const lastColumn = getSectionLastColumn(layout);
  copyRowStyle(outputWorksheet, originalTotalRow, adjustedTotalRow, lastColumn);
  copyRowValues(outputWorksheet, originalTotalRow, adjustedTotalRow, lastColumn);
  clearWorksheetRange(outputWorksheet, originalTotalRow, originalTotalRow, 1, lastColumn);
}

function updateSectionTotalResults(outputWorksheet, section) {
  const totalRow = section.endRow + 1;
  const amountLetter = XLSX.utils.encode_col(section.amountColumn - 1);
  const oppositeLetter = XLSX.utils.encode_col(section.oppositeColumn - 1);
  setFormulaResult(
    outputWorksheet.getRow(totalRow).getCell(section.amountColumn),
    sumWorksheetColumn(outputWorksheet, section.startRow, section.endRow, section.amountColumn),
    `SUM(${amountLetter}${section.startRow}:${amountLetter}${section.endRow})`,
  );
  setFormulaResult(
    outputWorksheet.getRow(totalRow).getCell(section.oppositeColumn),
    sumWorksheetColumn(outputWorksheet, section.startRow, section.endRow, section.oppositeColumn),
    `SUM(${oppositeLetter}${section.startRow}:${oppositeLetter}${section.endRow})`,
  );
}

function getDevolSectionRange(templateWorksheet, layout) {
  const lastSectionEnd = Math.max(...layout.sections.map((section) => section.endRow));
  const lastRow = Math.max(templateWorksheet.rowCount || 0, 1);
  let startRow = null;
  let endRow = null;

  for (let row = lastSectionEnd + 1; row <= lastRow; row += 1) {
    const account = getWorksheetCellText(templateWorksheet, row, 1);
    const name = getWorksheetCellText(templateWorksheet, row, 2).toUpperCase();
    if (account === "" || !name.includes("DEVOL")) {
      continue;
    }

    if (startRow == null) {
      startRow = row;
    }
    endRow = row;
  }

  if (startRow == null || endRow == null) {
    return null;
  }

  return {
    startRow,
    endRow,
    amountColumn: layout.debitColumn,
    oppositeColumn: layout.creditColumn,
  };
}

function updateMySheetFromRep(workbook, templateWorkbook, key, repGroups, ncGroups, options = {}) {
  const layout = MY_LAYOUTS[key];
  if (!layout) {
    throw new Error(`No existe layout MY para ${key}.`);
  }

  const outputWorksheet = workbook.getWorksheet(layout.mySheetName);
  const templateWorksheet = templateWorkbook.getWorksheet(layout.mySheetName);
  if (!outputWorksheet || !templateWorksheet) {
    throw new Error(`No existe la hoja requerida en plantilla: ${layout.mySheetName}`);
  }

  const resolvedNcGroups = ncGroups || {};
  const discountCreditGroups = resolvedNcGroups.discountCredit || {};
  const devolGroups = resolvedNcGroups.devol || {};
  const totalRows = {};

  for (let index = 0; index < layout.sections.length; index += 1) {
    const baseSection = layout.sections[index];
    const sourceGroups = baseSection.name === "sales"
      ? { debit: {}, credit: repGroups.sales }
      : { debit: repGroups.discount, credit: discountCreditGroups };
    const requiredRows = countSectionEntries(sourceGroups);
    const nextStart = index < layout.sections.length - 1 ? layout.sections[index + 1].startRow : null;
    const availableEnd = nextStart ? nextStart - 2 : baseSection.endRow;
    const adjustedSection = {
      ...baseSection,
      endRow: Math.max(baseSection.endRow, baseSection.startRow + requiredRows - 1),
    };

    if (adjustedSection.endRow > availableEnd) {
      throw new Error(
        `La hoja ${layout.mySheetName} no tiene espacio suficiente en la sección ${baseSection.startRow}-${baseSection.endRow}. `
        + `Capacidad máxima=${availableEnd - baseSection.startRow + 1}, requeridas=${requiredRows}.`,
      );
    }

    prepareSectionTotalRow(outputWorksheet, layout, baseSection, adjustedSection);
    const activeRows = applySectionScaling(outputWorksheet, templateWorksheet, layout, adjustedSection, sourceGroups);
    recalculateSectionSaldo(outputWorksheet, templateWorksheet, layout, adjustedSection, activeRows, options);
    updateSectionTotalResults(outputWorksheet, adjustedSection);
    totalRows[baseSection.name] = adjustedSection.endRow + 1;
  }

  const baseDevolSection = getDevolSectionRange(templateWorksheet, layout);
  if (baseDevolSection) {
    const requiredRows = countSectionEntries({ debit: devolGroups, credit: {} });
    const devolSection = {
      ...baseDevolSection,
      endRow: Math.max(baseDevolSection.endRow, baseDevolSection.startRow + requiredRows - 1),
    };
    prepareSectionTotalRow(outputWorksheet, layout, baseDevolSection, devolSection);
    const activeRows = applySectionScaling(
      outputWorksheet,
      templateWorksheet,
      layout,
      devolSection,
      { debit: devolGroups, credit: {} },
    );
    recalculateSectionSaldo(outputWorksheet, templateWorksheet, layout, devolSection, activeRows, options);
    updateSectionTotalResults(outputWorksheet, devolSection);
    totalRows.devol = devolSection.endRow + 1;
  }

  return totalRows;
}

function clearNcSheetNoSource(workbook, sheetName) {
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) {
    return;
  }

  const totalRow = findWorksheetRowContaining(worksheet, "TOTAL GENERAL", 1, 40, NC_MAX_COLUMN);
  if (!totalRow) {
    return;
  }

  const mayorRow = findWorksheetRowContaining(worksheet, "MAYOR", 1, 40, NC_MAX_COLUMN);
  if (totalRow > 8) {
    clearWorksheetRange(worksheet, 8, totalRow - 1, 1, NC_MAX_COLUMN);
  }

  for (const column of NC_TOTAL_COLUMNS) {
    worksheet.getRow(totalRow).getCell(column).value = 0;
  }

  if (mayorRow) {
    worksheet.getRow(mayorRow).getCell(22).value = "MAYOR";
  }
}

function getMayorIvaTemplateGroups(worksheet) {
  const debitGroups = {};
  const creditGroups = {};
  for (let row = MAYOR_IVA_START_ROW; row <= MAYOR_IVA_END_ROW; row += 1) {
    const type = getWorksheetCellText(worksheet, row, 5).toUpperCase();
    const seat = getWorksheetCellText(worksheet, row, 6);
    const detail = getWorksheetCellText(worksheet, row, 7);
    const debit = getWorksheetCellNumber(worksheet, row, 8);
    const credit = getWorksheetCellNumber(worksheet, row, 9);

    if (type !== "REPTO" || seat === "" || detail === "") {
      continue;
    }

    const groupKey = `${seat}|${detail}`;
    if (Math.abs(credit) >= 0.0000001 && Math.abs(debit) < 0.0000001) {
      if (!creditGroups[groupKey]) {
        creditGroups[groupKey] = {
          baseTotal: 0,
          rows: [],
        };
      }

      creditGroups[groupKey].baseTotal = roundAmount(creditGroups[groupKey].baseTotal + credit, 6);
      creditGroups[groupKey].rows.push({ rowNumber: row, baseAmount: credit });
      continue;
    }

    if (Math.abs(debit) >= 0.0000001 && Math.abs(credit) < 0.0000001) {
      if (!debitGroups[groupKey]) {
        debitGroups[groupKey] = {
          baseTotal: 0,
          rows: [],
        };
      }

      debitGroups[groupKey].baseTotal = roundAmount(debitGroups[groupKey].baseTotal + debit, 6);
      debitGroups[groupKey].rows.push({ rowNumber: row, baseAmount: debit });
    }
  }

  return { debit: debitGroups, credit: creditGroups };
}

function getMayorIvaOpeningBalance(templateWorksheet) {
  const saldo = getWorksheetCellNumber(templateWorksheet, MAYOR_IVA_START_ROW, 10);
  const debit = getWorksheetCellNumber(templateWorksheet, MAYOR_IVA_START_ROW, 8);
  const credit = getWorksheetCellNumber(templateWorksheet, MAYOR_IVA_START_ROW, 9);
  return roundAmount(saldo - debit + credit, 2);
}

function getMayorIvaFamilyOrder(detail) {
  switch (normalizeText(detail).toUpperCase()) {
    case "MOD. REPUESTOS REP01":
      return 0;
    case "MOD. REPUESTOS REP07":
    case "MOD. REPUESTOS REP08":
      return 1;
    case "MOD. REPUESTOS REP06":
      return 2;
    case "MOD. REPUESTOS REP05":
      return 3;
    default:
      return 99;
  }
}

function getMayorIvaDetailOrder(detail) {
  switch (normalizeText(detail).toUpperCase()) {
    case "MOD. REPUESTOS REP07":
      return 0;
    case "MOD. REPUESTOS REP08":
      return 1;
    default:
      return 0;
  }
}

function compareMayorIvaEntries(left, right) {
  return getMayorIvaFamilyOrder(left.detail) - getMayorIvaFamilyOrder(right.detail)
    || compareDateAscending(left.dateValue, right.dateValue)
    || compareSeatAscending(left.seat, right.seat)
    || getMayorIvaDetailOrder(left.detail) - getMayorIvaDetailOrder(right.detail)
    || compareTextAscending(left.side, right.side);
}

function buildMayorIvaEntries(repGroupsByKey, ncGroupsByKey) {
  const entries = [];

  for (const repGroups of Object.values(repGroupsByKey || {})) {
    for (const group of Object.values(repGroups?.vat || {})) {
      const amount = roundAmount(group?.amount || 0, 2);
      if (Math.abs(amount) < 0.0000001) {
        continue;
      }
      entries.push({
        side: "credit",
        amount,
        dateValue: group.dateValue || null,
        dateText: group.dateText || "",
        seat: group.seat || "",
        detail: group.detail || "",
      });
    }
  }

  for (const ncGroups of Object.values(ncGroupsByKey || {})) {
    for (const group of Object.values(ncGroups?.vat || {})) {
      const amount = roundAmount(group?.amount || 0, 2);
      if (Math.abs(amount) < 0.0000001) {
        continue;
      }
      entries.push({
        side: "debit",
        amount,
        dateValue: group.dateValue || null,
        dateText: group.dateText || "",
        seat: group.seat || "",
        detail: group.detail || "",
      });
    }
  }

  return entries
    .filter((entry) => entry.seat !== "" && entry.detail !== "")
    .sort(compareMayorIvaEntries);
}

function updateMayorIvaSummaryResults(worksheet, options = {}) {
  const summaryStartRow = Number.isInteger(options.summaryStartRow) ? options.summaryStartRow : 2;
  const windowStartRow = Number.isInteger(options.windowStartRow) ? options.windowStartRow : 283;
  const debitAll = sumWorksheetColumn(worksheet, summaryStartRow, MAYOR_IVA_END_ROW, 8);
  const creditAll = sumWorksheetColumn(worksheet, summaryStartRow, MAYOR_IVA_END_ROW, 9);
  const debitWindow = sumWorksheetColumn(worksheet, windowStartRow, MAYOR_IVA_END_ROW, 8);
  const creditWindow = sumWorksheetColumn(worksheet, windowStartRow, MAYOR_IVA_END_ROW, 9);

  setFormulaResult(worksheet.getRow(367).getCell(8), debitAll);
  setFormulaResult(worksheet.getRow(367).getCell(9), creditAll);
  setFormulaResult(worksheet.getRow(368).getCell(9), roundAmount(creditAll - debitAll, 2));
  setFormulaResult(worksheet.getRow(369).getCell(8), debitWindow);
  setFormulaResult(worksheet.getRow(369).getCell(9), creditWindow);
  setFormulaResult(worksheet.getRow(370).getCell(9), roundAmount(creditWindow - debitWindow, 2));
}

function updateMayorIvaFromRep(workbook, templateWorkbook, repGroupsByKey, ncGroupsByKey, options = {}) {
  const outputWorksheet = workbook.getWorksheet("MAYOR IVA");
  const templateWorksheet = templateWorkbook.getWorksheet("MAYOR IVA");
  if (!outputWorksheet || !templateWorksheet) {
    throw new Error("No existe la hoja requerida en plantilla: MAYOR IVA");
  }

  const useTemplateOpeningBalance = options.useTemplateOpeningBalance !== false;
  const clearCarryoverWindow = options.clearCarryoverWindow === true;
  const summaryStartRow = Number.isInteger(options.summaryStartRow) ? options.summaryStartRow : 2;
  const windowStartRow = Number.isInteger(options.windowStartRow) ? options.windowStartRow : 283;

  if (clearCarryoverWindow) {
    clearWorksheetRange(outputWorksheet, 280, MAYOR_IVA_START_ROW - 1, 1, MAYOR_IVA_LAST_COLUMN);
  }

  for (let row = MAYOR_IVA_START_ROW; row <= MAYOR_IVA_END_ROW; row += 1) {
    outputWorksheet.getRow(row).getCell(4).value = null;
    outputWorksheet.getRow(row).getCell(6).value = null;
    outputWorksheet.getRow(row).getCell(7).value = null;
    outputWorksheet.getRow(row).getCell(8).value = 0;
    outputWorksheet.getRow(row).getCell(9).value = 0;
    outputWorksheet.getRow(row).getCell(10).value = null;
  }

  const entries = buildMayorIvaEntries(repGroupsByKey, ncGroupsByKey);
  const capacity = MAYOR_IVA_END_ROW - MAYOR_IVA_START_ROW + 1;
  if (entries.length > capacity) {
    throw new Error(`La hoja MAYOR IVA no tiene filas suficientes. Capacidad=${capacity}, requeridas=${entries.length}.`);
  }

  const activeRows = new Set();
  entries.forEach((entry, index) => {
    const rowNumber = MAYOR_IVA_START_ROW + index;
    const row = outputWorksheet.getRow(rowNumber);
    activeRows.add(rowNumber);
    setDateCellValue(outputWorksheet, rowNumber, 4, entry.dateValue, entry.dateText);
    row.getCell(6).value = entry.seat;
    row.getCell(7).value = entry.detail;
    row.getCell(8).value = entry.side === "debit" ? entry.amount : 0;
    row.getCell(9).value = entry.side === "credit" ? entry.amount : 0;
  });

  for (let row = MAYOR_IVA_START_ROW; row <= MAYOR_IVA_END_ROW; row += 1) {
    if (!activeRows.has(row)) {
      clearWorksheetRange(outputWorksheet, row, row, 1, MAYOR_IVA_LAST_COLUMN);
    }
  }

  let runningBalance = useTemplateOpeningBalance ? getMayorIvaOpeningBalance(templateWorksheet) : 0;
  for (let row = MAYOR_IVA_START_ROW; row <= MAYOR_IVA_END_ROW; row += 1) {
    const templateType = getWorksheetCellText(templateWorksheet, row, 5);
    const templateSeat = getWorksheetCellText(templateWorksheet, row, 6);
    const templateDetail = getWorksheetCellText(templateWorksheet, row, 7);
    const templateDebit = getWorksheetCellNumber(templateWorksheet, row, 8);
    const templateCredit = getWorksheetCellNumber(templateWorksheet, row, 9);
    const templateSaldo = getWorksheetCellNumber(templateWorksheet, row, 10);

    if (
      templateType === "" &&
      templateSeat === "" &&
      templateDetail === "" &&
      Math.abs(templateDebit) < 0.0000001 &&
      Math.abs(templateCredit) < 0.0000001 &&
      Math.abs(templateSaldo) < 0.0000001
    ) {
      outputWorksheet.getRow(row).getCell(10).value = null;
      continue;
    }

    if (!activeRows.has(row)) {
      outputWorksheet.getRow(row).getCell(10).value = null;
      continue;
    }

    const debit = getWorksheetCellNumber(outputWorksheet, row, 8);
    const credit = getWorksheetCellNumber(outputWorksheet, row, 9);
    runningBalance = roundAmount(runningBalance + debit - credit, 2);
    outputWorksheet.getRow(row).getCell(10).value = runningBalance;
  }

  updateMayorIvaSummaryResults(outputWorksheet, { summaryStartRow, windowStartRow });
}

function updateRepMayorRows(workbook, myTotalRowsByKey) {
  const specs = {
    tyt: { sheet: "REP TYT", mySheet: "MY REP TYT", salesColumn: 9, discountColumn: 8, discountTargetColumn: 18 },
    peug: { sheet: "REP PEUGT", mySheet: "MY REP PEUG", salesColumn: 9, discountColumn: 8, discountTargetColumn: 18 },
    chgn: { sheet: "REP CHGN", mySheet: "MY REP CHGN", salesColumn: 10, discountColumn: 9, discountTargetColumn: 19 },
    szk: { sheet: "REP SZK", mySheet: "MY REP SZK", salesColumn: 9, discountColumn: 8, discountTargetColumn: 18 },
  };

  for (const config of SHEET_CONFIGS) {
    const spec = specs[config.key];
    const worksheet = workbook.getWorksheet(spec.sheet);
    const totals = myTotalRowsByKey?.[config.key];
    if (!worksheet) {
      continue;
    }

    const totalRow = findWorksheetRowContaining(worksheet, "TOTAL GENERAL", 1, 200, REP_MAX_COLUMN);
    const mayorRow = findWorksheetRowContaining(worksheet, "MAYOR", 1, 200, REP_MAX_COLUMN);
    if (!totalRow || !mayorRow) {
      continue;
    }

    const myWorksheet = workbook.getWorksheet(spec.mySheet);
    if (myWorksheet && totals?.sales) {
      setFormulaResult(
        worksheet.getRow(mayorRow).getCell(16),
        roundAmount(getWorksheetCellNumber(myWorksheet, totals.sales, spec.salesColumn), 2),
        `+'${spec.mySheet}'!${XLSX.utils.encode_col(spec.salesColumn - 1)}${totals.sales}`,
      );
    }

    if (myWorksheet && totals?.discount) {
      setFormulaResult(
        worksheet.getRow(mayorRow).getCell(spec.discountTargetColumn),
        roundAmount(getWorksheetCellNumber(myWorksheet, totals.discount, spec.discountColumn), 2),
        `+'${spec.mySheet}'!${XLSX.utils.encode_col(spec.discountColumn - 1)}${totals.discount}`,
      );
    }

    if (config.key === "peug") {
      const ncSheet = workbook.getWorksheet("NC REP PEUG");
      const ncVat = ncSheet ? getWorksheetCellNumber(ncSheet, 9, 30) : 0;
      const totalBase = getWorksheetCellNumber(worksheet, totalRow, 24);
      worksheet.getRow(mayorRow).getCell(25).value = roundAmount(totalBase - ncVat, 2);
      worksheet.getRow(mayorRow).getCell(26).value = roundAmount(totalBase, 2);
    }
  }
}

function extractXmlBlock(xml, tagName) {
  const pattern = new RegExp(`<${tagName}(?:\\s[^>]*)?>[\\s\\S]*?<\\/${tagName}>`);
  const match = xml.match(pattern);
  return match ? match[0] : "";
}

function replaceOrInsertXmlBlock(xml, tagName, block, insertBeforeTag) {
  const pattern = new RegExp(`<${tagName}(?:\\s[^>]*)?>[\\s\\S]*?<\\/${tagName}>`);
  if (pattern.test(xml)) {
    if (block) {
      return xml.replace(pattern, block);
    }
    return xml.replace(pattern, "");
  }

  if (!block) {
    return xml;
  }

  const insertIndex = xml.indexOf(insertBeforeTag);
  if (insertIndex >= 0) {
    return `${xml.slice(0, insertIndex)}${block}${xml.slice(insertIndex)}`;
  }

  return xml;
}

function repairWorkbookMetadataFromTemplate(outputPath, templatePath) {
  if (!outputPath || !templatePath || !fs.existsSync(outputPath) || !fs.existsSync(templatePath)) {
    return;
  }

  const outputZip = new AdmZip(outputPath);
  const templateZip = new AdmZip(templatePath);
  const workbookEntry = outputZip.getEntry("xl/workbook.xml");
  const templateWorkbookEntry = templateZip.getEntry("xl/workbook.xml");
  if (!workbookEntry || !templateWorkbookEntry) {
    return;
  }

  const outputXml = workbookEntry.getData().toString("utf8");
  const templateXml = templateWorkbookEntry.getData().toString("utf8");
  const definedNamesBlock = extractXmlBlock(templateXml, "definedNames");
  const extLstBlock = extractXmlBlock(templateXml, "extLst");

  let repairedXml = replaceOrInsertXmlBlock(outputXml, "definedNames", definedNamesBlock, "<calcPr");
  repairedXml = replaceOrInsertXmlBlock(repairedXml, "extLst", extLstBlock, "</workbook>");

  if (repairedXml !== outputXml) {
    outputZip.updateFile("xl/workbook.xml", Buffer.from(repairedXml, "utf8"));
    outputZip.writeZip(outputPath);
  }
}

function normalizeWorksheetCellReferences(outputPath) {
  if (!outputPath || !fs.existsSync(outputPath)) {
    return;
  }

  const outputZip = new AdmZip(outputPath);
  let changed = false;

  for (const entry of outputZip.getEntries()) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/i.test(entry.entryName)) {
      continue;
    }

    const xml = entry.getData().toString("utf8");
    const normalized = xml.replace(
      /<row\b([^>]*)r="(\d+)"([^>]*)>([\s\S]*?)<\/row>/g,
      (match, before, rowNumber, after, body) => {
        const fixedBody = body.replace(
          /(<c\b[^>]*\br=")([A-Z]+)(\d+)(")/g,
          (cellMatch, prefix, columnLetters, cellRow, suffix) => {
            if (cellRow === rowNumber) {
              return cellMatch;
            }
            return `${prefix}${columnLetters}${rowNumber}${suffix}`;
          },
        );

        if (fixedBody === body) {
          return match;
        }

        return `<row${before}r="${rowNumber}"${after}>${fixedBody}</row>`;
      },
    );

    if (normalized !== xml) {
      outputZip.updateFile(entry.entryName, Buffer.from(normalized, "utf8"));
      changed = true;
    }
  }

  if (changed) {
    outputZip.writeZip(outputPath);
  }
}

async function writeWorkbookToRequestedPath(workbook, outputPath, templatePath = "", maxAttempts = 8) {
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      await workbook.xlsx.writeFile(outputPath);
      repairWorkbookMetadataFromTemplate(outputPath, templatePath);
      normalizeWorksheetCellReferences(outputPath);
      return outputPath;
    } catch (error) {
      const locked = error && (error.code === "EBUSY" || error.code === "EPERM");
      if (!locked || attempt === maxAttempts) {
        throw error;
      }
    }
  }

  throw new Error("No se pudo guardar el Excel final en la ruta solicitada.");
}

async function main() {
  const args = parseArgs(process.argv);
  const templatePath = path.resolve(process.cwd(), args.templatepath || args["template-path"] || "");
  const outputPath = path.resolve(process.cwd(), args.outputpath || args["output-path"] || "");

  if (!templatePath || !fs.existsSync(templatePath)) {
    throw new Error(`No existe la plantilla base: ${templatePath}`);
  }

  if (!outputPath) {
    throw new Error("Falta -OutputPath/--output-path.");
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);
  workbook.calcProperties.fullCalcOnLoad = true;
  workbook.calcProperties.forceFullCalc = true;

  const templateWorkbook = new ExcelJS.Workbook();
  await templateWorkbook.xlsx.readFile(templatePath);

  const sourceDataByKey = {};
  const sourceDataList = [];
  for (const config of SHEET_CONFIGS) {
    const sourcePath = path.resolve(process.cwd(), getArg(args, config.argKeys));
    if (!sourcePath || !fs.existsSync(sourcePath)) {
      throw new Error(`No existe archivo de entrada para ${config.label}: ${sourcePath}`);
    }

    const sourceData = readSourceSheet(sourcePath, config.label);
    sourceDataByKey[config.key] = sourceData;
    sourceDataList.push(sourceData);
  }

  const globalSourceRange = buildCombinedDateRange(sourceDataList);
  const useTemplateCarryover = shouldUseTemplateCarryover(templateWorkbook, sourceDataList);

  const repGroupsByKey = {};
  const ncGroupsByKey = {};
  const myTotalRowsByKey = {};

  for (const config of SHEET_CONFIGS) {
    const sourceData = sourceDataByKey[config.key];
    const repLookup = buildRepLookup(sourceData.sheet, config.key);
    const repResult = applyRepSheet(sourceData, workbook, config);
    repGroupsByKey[config.key] = buildRepPostingGroups(sourceData.sheet, config.key);
    ncGroupsByKey[config.key] = useTemplateCarryover
      ? buildNcPostingGroups(templateWorkbook, config.key, repLookup)
      : { discountCredit: {}, devol: {}, vat: {} };
    myTotalRowsByKey[config.key] = updateMySheetFromRep(
      workbook,
      templateWorkbook,
      config.key,
      repGroupsByKey[config.key],
      ncGroupsByKey[config.key],
      { useTemplateOpeningBalance: useTemplateCarryover },
    );

    console.log(`INFO|${config.key}|rows=${repResult.rowCount}`);
    console.log(`INFO|${config.key}|sheet=${config.targetSheet}`);
    console.log(`INFO|${config.key}|label=${config.label}`);
  }

  // La lógica de MAYOR IVA ha sido deshabilitada porque la plantilla no contiene fórmulas y los cálculos eran incorrectos.
  if (!useTemplateCarryover) {
    const monthStartDate = getMonthStartDate(globalSourceRange.minDate) || globalSourceRange.minDate;
    for (const config of SHEET_CONFIGS) {
      const ncSheetName = getNcSheetName(config.key);
      clearNcSheetNoSource(workbook, ncSheetName);
      updateSheetDateRange(workbook.getWorksheet(ncSheetName), monthStartDate, globalSourceRange.maxDate);
    }
  }

  updateMayorIvaFromRep(workbook, templateWorkbook, repGroupsByKey, ncGroupsByKey, {
    useTemplateOpeningBalance: useTemplateCarryover,
    clearCarryoverWindow: !useTemplateCarryover,
    summaryStartRow: useTemplateCarryover ? 2 : MAYOR_IVA_START_ROW,
    windowStartRow: useTemplateCarryover ? 283 : MAYOR_IVA_START_ROW,
  });
  console.log("INFO|mayor_iva|updated=1");
  updateRepMayorRows(workbook, myTotalRowsByKey);

  const finalPath = await writeWorkbookToRequestedPath(workbook, outputPath, templatePath);
  console.log(`OUTPUT|${path.basename(finalPath)}|FACTURACION REPUESTOS TYTSERV`);
}

if (require.main === module) {
  main().catch((error) => {
    if (error && error.stack) {
      console.error(error.stack);
    } else {
      console.error(error && error.message ? error.message : String(error));
    }
    process.exit(1);
  });
}


