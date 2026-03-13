const AdmZip = require("adm-zip");
const { XMLBuilder, XMLParser } = require("fast-xml-parser");

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
  return String(inputPath || "").replace(/\\/g, "/");
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
    throw new Error(`No se encontro la entrada ZIP: ${entryPath}`);
  }
  return entry.getData().toString("utf8");
}

function updateZipText(zip, entryPath, xmlObject) {
  zip.updateFile(normalizeZipPath(entryPath), Buffer.from(buildXml(xmlObject), "utf8"));
}

function getWorksheetEntryPath(zip, sheetName) {
  const workbookXml = parseXml(readZipText(zip, "xl/workbook.xml"));
  const workbookRelsXml = parseXml(readZipText(zip, "xl/_rels/workbook.xml.rels"));

  const sheets = toArray(workbookXml.workbook.sheets.sheet);
  const rels = toArray(workbookRelsXml.Relationships.Relationship);
  const targetSheet = sheets.find((sheet) => String(sheet["@_name"] || "") === sheetName);
  if (!targetSheet) {
    throw new Error(`No se encontro la hoja ${sheetName} en workbook.xml.`);
  }

  const relationId = targetSheet["@_r:id"];
  const relation = rels.find((item) => item["@_Id"] === relationId);
  if (!relation) {
    throw new Error(`No se encontro la relacion ${relationId} para ${sheetName}.`);
  }

  return `xl/${String(relation["@_Target"] || "").replace(/^\/+/, "")}`;
}

function getRowsArray(sheetXmlObject) {
  return toArray(sheetXmlObject.worksheet.sheetData.row);
}

function setRowsArray(sheetXmlObject, rows) {
  sheetXmlObject.worksheet.sheetData.row = rows;
}

function getCellsArray(rowObject) {
  return toArray(rowObject.c);
}

function setCellsArray(rowObject, cells) {
  rowObject.c = cells;
}

function parseCellRef(ref) {
  const match = /^([A-Z]+)(\d+)$/.exec(String(ref || ""));
  if (!match) {
    return null;
  }
  return {
    colName: match[1],
    rowNumber: Number(match[2]),
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
  for (const key of ["v", "f", "is", "vm", "@_t", "@_s", "@_cm", "extLst"]) {
    if (sourceCell[key] != null) {
      targetCell[key] = JSON.parse(JSON.stringify(sourceCell[key]));
    } else {
      delete targetCell[key];
    }
  }
}

function cloneTemplateRow(rowTemplate, targetRowNumber) {
  const clone = JSON.parse(JSON.stringify(rowTemplate));
  clone["@_r"] = String(targetRowNumber);
  const cells = getCellsArray(clone).map((cell) => {
    const nextCell = JSON.parse(JSON.stringify(cell));
    const parsedRef = parseCellRef(nextCell["@_r"]);
    if (parsedRef) {
      nextCell["@_r"] = `${parsedRef.colName}${targetRowNumber}`;
    }
    return nextCell;
  });
  setCellsArray(clone, cells);
  return clone;
}

function clearCellPayload(targetCell) {
  delete targetCell.v;
  delete targetCell.f;
  delete targetCell.is;
  delete targetCell.vm;
  delete targetCell.extLst;
  targetCell["@_t"] = "n";
}

function hasCellPayload(cell) {
  return cell.v != null || cell.f != null || cell.is != null;
}

function stripCalcChain(zip) {
  const calcChainPath = "xl/calcChain.xml";
  if (zip.getEntry(calcChainPath)) {
    zip.deleteFile(calcChainPath);
  }
}

function stableStringify(value) {
  if (value === null || typeof value !== "object") {
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
