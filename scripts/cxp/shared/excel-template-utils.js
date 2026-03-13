const path = require("path");

const AdmZip = require("adm-zip");
const { XMLBuilder, XMLParser } = require("fast-xml-parser");

const { sanitizeText, deepClone } = require("./core-utils");

const XML_PARSER = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  allowBooleanAttributes: true,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
});

const XML_BUILDER = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  suppressBooleanAttributes: false,
  format: false,
});

function captureStyleMatrix(ws, maxRows, maxCols) {
  const matrix = [];
  for (let row = 1; row <= maxRows; row += 1) {
    const rowStyles = [];
    for (let col = 1; col <= maxCols; col += 1) {
      rowStyles.push(JSON.parse(JSON.stringify(ws.getCell(row, col).style || {})));
    }
    matrix.push(rowStyles);
  }
  return matrix;
}

function applyStyleFromMatrix(ws, matrix, row, col) {
  ws.getCell(row, col).style = JSON.parse(JSON.stringify(matrix[row - 1]?.[col - 1] || {}));
}

function clearRangeValues(ws, startRow, endRow, startCol, endCol) {
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      ws.getCell(row, col).value = null;
    }
  }
}

function toArray(value) {
  if (Array.isArray(value)) {
    return value;
  }
  if (value == null) {
    return [];
  }
  return [value];
}

function normalizeZipPath(inputPath) {
  return String(inputPath || "").replace(/\\/g, "/").replace(/^\/+/, "");
}

function parseXml(xmlText) {
  return XML_PARSER.parse(xmlText);
}

function buildXml(xmlObject) {
  const body = XML_BUILDER.build(xmlObject);
  if (body.startsWith("<?xml")) {
    return body;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body}`;
}

function readZipText(zip, entryPath) {
  const entry = zip.getEntry(normalizeZipPath(entryPath));
  if (!entry) {
    return null;
  }
  return entry.getData().toString("utf8");
}

function updateZipText(zip, entryPath, xmlObject) {
  zip.updateFile(normalizeZipPath(entryPath), Buffer.from(buildXml(xmlObject), "utf8"));
}

function getWorksheetEntryPath(zip, sheetName) {
  const workbookText = readZipText(zip, "xl/workbook.xml");
  const workbookRelsText = readZipText(zip, "xl/_rels/workbook.xml.rels");
  if (!workbookText || !workbookRelsText) {
    throw new Error("No se pudo leer la estructura interna del Excel.");
  }

  const workbookXml = parseXml(workbookText);
  const workbookRelsXml = parseXml(workbookRelsText);

  const sheets = toArray(workbookXml.workbook?.sheets?.sheet);
  const rels = toArray(workbookRelsXml.Relationships?.Relationship);
  const normalizedSheetName = sanitizeText(sheetName).toUpperCase();
  const targetSheet = sheets.find(
    (sheet) => sanitizeText(sheet?.["@_name"]).toUpperCase() === normalizedSheetName,
  );
  if (!targetSheet) {
    throw new Error(`No se encontro la hoja ${sheetName} en el Excel.`);
  }

  const relationId = targetSheet["@_r:id"];
  const relation = rels.find((item) => item["@_Id"] === relationId);
  if (!relation) {
    throw new Error(`No se pudo resolver la ruta XML de la hoja ${sheetName}.`);
  }

  const target = normalizeZipPath(relation["@_Target"]);
  return target.startsWith("xl/") ? target : normalizeZipPath(path.posix.join("xl", target));
}

function getRowsArray(sheetXmlObject) {
  return toArray(sheetXmlObject.worksheet?.sheetData?.row);
}

function setRowsArray(sheetXmlObject, rows) {
  if (!sheetXmlObject.worksheet) {
    sheetXmlObject.worksheet = {};
  }
  if (!sheetXmlObject.worksheet.sheetData) {
    sheetXmlObject.worksheet.sheetData = {};
  }
  sheetXmlObject.worksheet.sheetData.row = rows;
}

function getCellsArray(rowObject) {
  return toArray(rowObject?.c);
}

function setCellsArray(rowObject, cells) {
  if (!cells.length) {
    delete rowObject.c;
    return;
  }
  rowObject.c = cells;
}

function parseCellRef(ref) {
  const match = /^([A-Z]+)(\d+)$/i.exec(String(ref || ""));
  if (!match) {
    return null;
  }
  const colName = match[1].toUpperCase();
  const colNumber = columnNameToNumber(colName);
  const rowNumber = Number(match[2]);
  return {
    col: colNumber,
    row: rowNumber,
    colName,
    colNumber,
    rowNumber,
  };
}

function columnNumberToName(colNumber) {
  let result = "";
  let current = Number(colNumber);
  while (current > 0) {
    const remainder = (current - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    current = Math.floor((current - 1) / 26);
  }
  return result;
}

function columnNameToNumber(colName) {
  let result = 0;
  const normalized = String(colName || "").toUpperCase();
  for (let i = 0; i < normalized.length; i += 1) {
    result = result * 26 + (normalized.charCodeAt(i) - 64);
  }
  return result;
}

function cellRef(colNumber, rowNumber) {
  return `${columnNumberToName(colNumber)}${rowNumber}`;
}

function copyCellPayload(targetCell, sourceCell) {
  for (const key of ["v", "f", "is", "vm", "@_t", "@_cm", "@_vm", "extLst"]) {
    if (sourceCell[key] != null) {
      targetCell[key] = deepClone(sourceCell[key]);
    } else {
      delete targetCell[key];
    }
  }
}

function cloneTemplateRow(rowTemplate, targetRowNumber) {
  const clone = deepClone(rowTemplate);
  clone["@_r"] = String(targetRowNumber);
  const cells = getCellsArray(clone).map((cell) => {
    const nextCell = deepClone(cell);
    const parsedRef = parseCellRef(nextCell["@_r"]);
    if (parsedRef) {
      nextCell["@_r"] = cellRef(parsedRef.colNumber, targetRowNumber);
    }
    return nextCell;
  });
  setCellsArray(clone, cells);
  return clone;
}

function clearCellPayload(targetCell) {
  delete targetCell["@_cm"];
  delete targetCell["@_vm"];
  delete targetCell.v;
  delete targetCell.f;
  delete targetCell.is;
  delete targetCell.vm;
  delete targetCell.extLst;
  delete targetCell["@_t"];
}

function hasCellPayload(cell) {
  return cell.v != null || cell.f != null || cell.is != null;
}

function stripCalcChain(zip) {
  const calcChainPath = "xl/calcChain.xml";
  if (zip.getEntry(calcChainPath)) {
    zip.deleteFile(calcChainPath);
  }

  const relsXml = readZipText(zip, "xl/_rels/workbook.xml.rels");
  if (relsXml) {
    const relsObj = parseXml(relsXml);
    const rels = toArray(relsObj.Relationships?.Relationship);
    const kept = rels.filter((item) => !String(item?.["@_Type"] || "").endsWith("/calcChain"));
    relsObj.Relationships.Relationship = kept;
    updateZipText(zip, "xl/_rels/workbook.xml.rels", relsObj);
  }

  const contentTypesXml = readZipText(zip, "[Content_Types].xml");
  if (contentTypesXml) {
    const typesObj = parseXml(contentTypesXml);
    const overrides = toArray(typesObj.Types?.Override);
    const kept = overrides.filter((item) => item?.["@_PartName"] !== "/xl/calcChain.xml");
    typesObj.Types.Override = kept;
    updateZipText(zip, "[Content_Types].xml", typesObj);
  }
}

function stableStringify(value) {
  if (value === null || value === undefined) {
    return String(value);
  }
  if (typeof value !== "object") {
    return JSON.stringify(value);
  }
  if (Array.isArray(value)) {
    return `[${value.map((item) => stableStringify(item)).join(",")}]`;
  }
  const keys = Object.keys(value).sort();
  return `{${keys.map((key) => `${JSON.stringify(key)}:${stableStringify(value[key])}`).join(",")}}`;
}

function styleSignature(style) {
  return stableStringify(style || {});
}

module.exports = {
  AdmZip,
  captureStyleMatrix,
  applyStyleFromMatrix,
  clearRangeValues,
  toArray,
  normalizeZipPath,
  parseXml,
  buildXml,
  readZipText,
  updateZipText,
  getWorksheetEntryPath,
  getRowsArray,
  setRowsArray,
  getCellsArray,
  setCellsArray,
  parseCellRef,
  columnNumberToName,
  columnNameToNumber,
  cellRef,
  copyCellPayload,
  cloneTemplateRow,
  clearCellPayload,
  hasCellPayload,
  stripCalcChain,
  stableStringify,
  styleSignature,
};
