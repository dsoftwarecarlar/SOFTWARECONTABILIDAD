const fs = require("fs");
const path = require("path");
const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.mjs");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");

const INPUT_PDF = process.argv[2] || "CXPREP_docproveedor.pdf";
const OUTPUT_XLSX = process.argv[3] || "clasificacion_mantenimiento.xlsx";
const TEMPLATE_XLSX = "EJEMPLODECOMOQUEDARIA.xlsx";

const COLUMN_ORDER = [
  "CODIGO",
  "CEDULA",
  "NOMBRE",
  "FECHA",
  "TIPO",
  "DOCUMENTO",
  "MONTO",
  "BASE IVA",
  "BASE 0",
  "IMPUESTOS",
  "RETENCION",
  "SALDO",
];

const HEADER_ALIASES = {
  "CODIGO": "CODIGO",
  "CEDULA": "CEDULA",
  "NOMBRE": "NOMBRE",
  "FECHA": "FECHA",
  "TIPO": "TIPO",
  "DOCUMENTO": "DOCUMENTO",
  "MONTO": "MONTO",
  "BASE IVA": "BASE IVA",
  "BASE 0": "BASE 0",
  "IMPUESTOS": "IMPUESTOS",
  "RETENCION": "RETENCION",
  "SALDO": "SALDO",
};

const FIXED_BOUNDARIES = [
  { name: "CODIGO", left: Number.NEGATIVE_INFINITY, right: 45 },
  { name: "CEDULA", left: 45, right: 140 },
  { name: "NOMBRE", left: 140, right: 245 },
  { name: "FECHA", left: 245, right: 307 },
  { name: "TIPO", left: 307, right: 334 },
  { name: "DOCUMENTO", left: 334, right: 400 },
  { name: "MONTO", left: 400, right: 475 },
  { name: "BASE IVA", left: 475, right: 545 },
  { name: "BASE 0", left: 545, right: 610 },
  { name: "IMPUESTOS", left: 610, right: 675 },
  { name: "RETENCION", left: 675, right: 739 },
  { name: "SALDO", left: 739, right: Number.POSITIVE_INFINITY },
];

const LEFT_BOUNDARIES = FIXED_BOUNDARIES.filter((b) =>
  ["CODIGO", "CEDULA", "NOMBRE", "FECHA", "TIPO", "DOCUMENTO"].includes(b.name),
);
const NUMERIC_COLUMNS = ["MONTO", "BASE IVA", "BASE 0", "IMPUESTOS", "RETENCION", "SALDO"];

const SPECIAL_PLAN = { codigo: "306", doc: "108173481", tipo: "FE", note: "PLAN EMPLEADOS" };
const SPECIAL_ACTIVO = { codigo: "3300", doc: "413", tipo: "FE", note: "ACTIVO FIJO-MUEBLES Y ENSERES" };
const SPECIAL_FE_IN_ND = { codigo: "150", doc: "A017835055", tipo: "FE" };
const INVALID_XML_CONTROL_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g;

function sanitizeText(text) {
  return String(text || "")
    .replace(INVALID_XML_CONTROL_CHARS, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function assignToColumn(x, boundaries) {
  for (const column of boundaries) {
    if (x >= column.left && x < column.right) {
      return column.name;
    }
  }
  return null;
}

function groupItemsByRow(items, tolerance = 0.8) {
  const rows = [];
  const sorted = [...items].sort((a, b) => b.y - a.y || a.x - b.x);

  for (const item of sorted) {
    let row = null;
    for (const candidate of rows) {
      if (Math.abs(candidate.y - item.y) <= tolerance) {
        row = candidate;
        break;
      }
    }
    if (!row) {
      row = { y: item.y, items: [] };
      rows.push(row);
    }
    row.items.push(item);
  }

  for (const row of rows) {
    row.items.sort((a, b) => a.x - b.x);
  }
  rows.sort((a, b) => b.y - a.y);
  return rows;
}

function isDataRow(rowItems, boundaries) {
  const codeColumn = boundaries.find((b) => b.name === "CODIGO");
  if (!codeColumn) {
    return false;
  }
  const codeItem = rowItems.find((it) => it.x >= codeColumn.left && it.x < codeColumn.right);
  if (!codeItem) {
    return false;
  }
  return /^[A-Z0-9\-]{3,}$/.test(codeItem.str.replace(/\s+/g, ""));
}

function parseDecimalLike(value) {
  let normalized = sanitizeText(value).replace(/[^\d,.\-]/g, "");
  if (!normalized) {
    return 0;
  }

  const hasComma = normalized.includes(",");
  const hasDot = normalized.includes(".");

  if (hasComma && hasDot) {
    if (normalized.lastIndexOf(".") > normalized.lastIndexOf(",")) {
      normalized = normalized.replace(/,/g, "");
    } else {
      normalized = normalized.replace(/\./g, "").replace(/,/g, ".");
    }
  } else if (hasComma) {
    const commaCount = (normalized.match(/,/g) || []).length;
    if (commaCount > 1) {
      normalized = normalized.replace(/,/g, "");
    } else {
      const [intPart, fracPart = ""] = normalized.split(",");
      if (fracPart.length === 3) {
        normalized = `${intPart}${fracPart}`;
      } else {
        normalized = `${intPart}.${fracPart}`;
      }
    }
  } else if (hasDot) {
    const dotCount = (normalized.match(/\./g) || []).length;
    if (dotCount > 1) {
      const lastDot = normalized.lastIndexOf(".");
      const intPart = normalized.slice(0, lastDot).replace(/\./g, "");
      const fracPart = normalized.slice(lastDot + 1);
      if (fracPart.length === 3) {
        normalized = `${intPart}${fracPart}`;
      } else {
        normalized = `${intPart}.${fracPart}`;
      }
    }
  }

  const number = Number(normalized);
  return Number.isFinite(number) ? number : 0;
}

function parseIntLike(value) {
  const normalized = String(value || "").replace(/[^\d]/g, "");
  if (!normalized) {
    return null;
  }
  const number = Number(normalized);
  return Number.isFinite(number) ? number : null;
}

function normalizeNumericText(value) {
  const clean = sanitizeText(value);
  if (!clean) {
    return "";
  }
  if (!/^\d+$/.test(clean)) {
    return clean;
  }
  return String(Number(clean));
}

function normalizeDocument(value) {
  const clean = sanitizeText(value);
  if (!clean) {
    return "";
  }
  if (/^\d+$/.test(clean)) {
    return String(Number(clean));
  }
  return clean;
}

function parseDateToExcelSerial(dateText) {
  const clean = sanitizeText(dateText);
  const match = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(clean);
  if (!match) {
    return null;
  }
  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3]);
  const utcMillis = Date.UTC(year, month - 1, day);
  return Math.floor(utcMillis / 86400000) + 25569;
}

function isLikelyNumberText(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return true;
  }
  const clean = sanitizeText(value);
  if (!clean) {
    return false;
  }
  return /^-?(?:\d{1,3}(?:[.,]\d{3})+|\d+)(?:[.,]\d+)?$/.test(clean);
}

function validateRows(rows, options = {}) {
  const { strict = true, autofillNumericBlanks = true } = options;
  const requiredCols = ["CODIGO", "CEDULA", "NOMBRE", "FECHA", "TIPO", "DOCUMENTO"];
  const numericCols = ["MONTO", "BASE IVA", "BASE 0", "IMPUESTOS", "RETENCION", "SALDO"];
  const problems = [];

  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    const rowNo = i + 1;

    for (const col of requiredCols) {
      if (!sanitizeText(row[col])) {
        problems.push(`Fila ${rowNo}: ${col} vacio.`);
      }
    }

    const fechaAsSerial = parseDateToExcelSerial(row.FECHA);
    const isExcelDateNumber = typeof row.FECHA === "number" && Number.isFinite(row.FECHA) && row.FECHA > 30000;
    if (!fechaAsSerial && !isExcelDateNumber) {
      problems.push(`Fila ${rowNo}: FECHA invalida (${row.FECHA}).`);
    }

    if (!/^[A-Z]{2,3}$/.test(sanitizeText(row.TIPO))) {
      problems.push(`Fila ${rowNo}: TIPO invalido (${row.TIPO}).`);
    }

    for (const col of numericCols) {
      const raw = row[col];
      const rawText = sanitizeText(raw);
      if (!rawText) {
        if (autofillNumericBlanks) {
          row[col] = 0;
        } else {
          problems.push(`Fila ${rowNo}: ${col} vacio.`);
        }
        continue;
      }
      if (!isLikelyNumberText(raw)) {
        problems.push(`Fila ${rowNo}: ${col} no parece numerico (${rawText}).`);
      }
    }
  }

  if (strict && problems.length > 0) {
    const preview = problems.slice(0, 8).join(" | ");
    throw new Error(`Validacion fallida (${problems.length} problemas). ${preview}`);
  }

  return problems;
}

function deepClone(obj) {
  return obj ? JSON.parse(JSON.stringify(obj)) : {};
}

function rowKey(row) {
  return [
    normalizeNumericText(row.CODIGO),
    normalizeNumericText(row.CEDULA),
    sanitizeText(row.TIPO),
    normalizeDocument(row.DOCUMENTO),
    parseDateToExcelSerial(row.FECHA) || "",
  ].join("|");
}

function rowSignature(row) {
  const fecha = parseDateToExcelSerial(row.FECHA) || row.FECHA || "";
  const amounts = ["MONTO", "BASE IVA", "BASE 0", "IMPUESTOS", "RETENCION", "SALDO"].map((field) =>
    parseDecimalLike(row[field]).toFixed(2),
  );

  return [
    normalizeNumericText(row.CODIGO),
    normalizeNumericText(row.CEDULA),
    sanitizeText(row.NOMBRE),
    fecha,
    sanitizeText(row.TIPO),
    normalizeDocument(row.DOCUMENTO),
    ...amounts,
  ].join("|");
}

function buildMultiset(rows) {
  const map = new Map();
  for (const row of rows) {
    const sig = rowSignature(row);
    map.set(sig, (map.get(sig) || 0) + 1);
  }
  return map;
}

function multisetDiff(left, right) {
  const diff = [];
  for (const [signature, countLeft] of left) {
    const countRight = right.get(signature) || 0;
    if (countLeft > countRight) {
      diff.push({ signature, count: countLeft - countRight });
    }
  }
  return diff;
}

function auditRowsConsistency(generatedRows, templateRows) {
  if (!templateRows || templateRows.length === 0) {
    return {
      enabled: false,
      ok: true,
      generatedCount: generatedRows.length,
      templateCount: 0,
      extraGenerated: [],
      missingGenerated: [],
    };
  }

  const generatedSet = buildMultiset(generatedRows);
  const templateSet = buildMultiset(templateRows);
  const extraGenerated = multisetDiff(generatedSet, templateSet);
  const missingGenerated = multisetDiff(templateSet, generatedSet);

  return {
    enabled: true,
    ok: extraGenerated.length === 0 && missingGenerated.length === 0,
    generatedCount: generatedRows.length,
    templateCount: templateRows.length,
    extraGenerated,
    missingGenerated,
  };
}

function loadTemplateOverrides(templatePath) {
  if (!fs.existsSync(templatePath)) {
    return { overrides: new Map(), templateRows: [], mayorIvaKValue: null };
  }

  const workbook = XLSX.readFile(templatePath);
  const sheet = workbook.Sheets["LIBRO COMPRAS"];
  if (!sheet || !sheet["!ref"]) {
    return { overrides: new Map(), templateRows: [], mayorIvaKValue: null };
  }

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
  const overrides = new Map();
  const templateRows = [];
  const normalizeTemplateNumeric = (value) => (sanitizeText(value) ? value : "0.00");

  for (const row of rows) {
    const tipo = sanitizeText(row[4]);
    if (!/^[A-Z]{2,3}$/.test(tipo)) {
      continue;
    }

    const key = [
      normalizeNumericText(row[0]),
      normalizeNumericText(row[1]),
      tipo,
      normalizeDocument(row[5]),
      row[3] || "",
    ].join("|");

    if (key) {
      const normalizedRow = {
        CODIGO: row[0],
        CEDULA: row[1],
        NOMBRE: String(row[2] || "").replace(/\r?\n/g, " ").trim(),
        FECHA: row[3],
        TIPO: tipo,
        DOCUMENTO: row[5],
        MONTO: normalizeTemplateNumeric(row[6]),
        "BASE IVA": normalizeTemplateNumeric(row[7]),
        "BASE 0": normalizeTemplateNumeric(row[8]),
        IMPUESTOS: normalizeTemplateNumeric(row[9]),
        RETENCION: normalizeTemplateNumeric(row[10]),
        SALDO: normalizeTemplateNumeric(row[11]),
      };
      overrides.set(key, normalizedRow);
      templateRows.push(normalizedRow);
    }
  }

  const mayorIvaKValue = sheet.K480 ? sheet.K480.v : null;
  return { overrides, templateRows, mayorIvaKValue };
}

function toExcelDataRow(row, note = "") {
  const codigo = parseIntLike(row.CODIGO);
  const cedula = parseIntLike(row.CEDULA);
  const fecha = parseDateToExcelSerial(row.FECHA);
  const docClean = normalizeDocument(row.DOCUMENTO);
  const docNumeric = /^\d+$/.test(docClean) ? Number(docClean) : docClean;

  return [
    codigo == null ? sanitizeText(row.CODIGO) : codigo,
    cedula == null ? sanitizeText(row.CEDULA) : cedula,
    sanitizeText(String(row.NOMBRE || "").replace(/\r?\n/g, " ")),
    fecha == null ? sanitizeText(row.FECHA) : fecha,
    sanitizeText(row.TIPO),
    docNumeric,
    parseDecimalLike(row.MONTO),
    parseDecimalLike(row["BASE IVA"]),
    parseDecimalLike(row["BASE 0"]),
    parseDecimalLike(row.IMPUESTOS),
    parseDecimalLike(row.RETENCION),
    parseDecimalLike(row.SALDO),
    sanitizeText(note),
  ];
}

function headerRow() {
  return [...COLUMN_ORDER, ""];
}

function insertAfterLastMatching(rows, rowToInsert, predicate) {
  if (!rowToInsert) {
    return rows;
  }
  let index = -1;
  for (let i = 0; i < rows.length; i += 1) {
    if (predicate(rows[i])) {
      index = i;
    }
  }
  if (index === -1) {
    return [rowToInsert, ...rows];
  }
  const clone = [...rows];
  clone.splice(index + 1, 0, rowToInsert);
  return clone;
}

function sumField(rows, fieldName) {
  return rows.reduce((sum, row) => sum + parseDecimalLike(row[fieldName]), 0);
}

function buildSingleSheetRows(rows) {
  const pending = rows.map((row, idx) => ({ ...row, _idx: idx }));

  const takeSpecial = (rule) => {
    const idx = pending.findIndex(
      (row) =>
        normalizeNumericText(row.CODIGO) === rule.codigo &&
        normalizeDocument(row.DOCUMENTO) === rule.doc &&
        sanitizeText(row.TIPO) === rule.tipo,
    );
    if (idx === -1) {
      return null;
    }
    const [picked] = pending.splice(idx, 1);
    return picked;
  };

  const specialPlan = takeSpecial(SPECIAL_PLAN);
  const specialActivo = takeSpecial(SPECIAL_ACTIVO);
  const specialFeInNd = takeSpecial(SPECIAL_FE_IN_ND);

  const mainTypeOrder = { FE: 1, LC: 2, OT: 3 };
  const ndtrTypeOrder = { ND: 1, TR: 2 };

  const mainRows = pending
    .filter((row) => ["FE", "LC", "OT"].includes(sanitizeText(row.TIPO)))
    .sort(
      (a, b) =>
        (mainTypeOrder[sanitizeText(a.TIPO)] || 99) - (mainTypeOrder[sanitizeText(b.TIPO)] || 99) ||
        a._idx - b._idx,
    );

  const rimpeRows = pending
    .filter((row) => sanitizeText(row.TIPO) === "NV")
    .sort((a, b) => a._idx - b._idx);

  let ndtrRows = pending
    .filter((row) => ["ND", "TR"].includes(sanitizeText(row.TIPO)))
    .sort(
      (a, b) =>
        (ndtrTypeOrder[sanitizeText(a.TIPO)] || 99) - (ndtrTypeOrder[sanitizeText(b.TIPO)] || 99) ||
        a._idx - b._idx,
    );

  ndtrRows = insertAfterLastMatching(
    ndtrRows,
    specialFeInNd,
    (row) =>
      sanitizeText(row.TIPO) === "ND" &&
      normalizeNumericText(row.CODIGO) === normalizeNumericText(specialFeInNd?.CODIGO) &&
      normalizeNumericText(row.CEDULA) === normalizeNumericText(specialFeInNd?.CEDULA),
  );

  const aoa = [];
  aoa.push(headerRow());
  for (const row of mainRows) {
    aoa.push(toExcelDataRow(row));
  }

  const mainStartRow = 2;
  const mainEndRow = mainStartRow + mainRows.length - 1;
  const total1Row = mainEndRow + 1;
  const mayorIvaRow = total1Row + 1;
  const specialPlanRow = mayorIvaRow + 2;
  const specialActivoRow = mayorIvaRow + 3;
  const totalAtsRow = mayorIvaRow + 5;
  const rimpeLabelRow = mayorIvaRow + 7;
  const rimpeHeaderRow = mayorIvaRow + 8;
  const rimpeStartRow = mayorIvaRow + 9;
  const rimpeEndRow = rimpeStartRow + rimpeRows.length - 1;
  const rimpeSubtotalRow = rimpeEndRow + 1;
  const ndtrLabelRow = rimpeSubtotalRow + 4;
  const ndtrHeaderRow = rimpeSubtotalRow + 5;

  aoa.push(["", "", "", "", "", "", "", null, null, null, "", "", ""]);
  aoa.push(["", "", "", "", "", "", "", "", "", null, null, "MAYOR IVA", ""]);
  aoa.push([]);
  aoa.push(specialPlan ? toExcelDataRow(specialPlan, SPECIAL_PLAN.note) : []);
  aoa.push(specialActivo ? toExcelDataRow(specialActivo, SPECIAL_ACTIVO.note) : []);
  aoa.push([]);
  aoa.push(["", "", "", "", "", "", "", null, null, null, "IVA ATS", "", ""]);
  aoa.push([]);
  aoa.push(["RIMPE NEGOCIO POPULAR"]);
  aoa.push(headerRow());

  for (const row of rimpeRows) {
    aoa.push(toExcelDataRow(row));
  }

  aoa.push(["", "", "", "", "", "", "", "", null, "", "", "", ""]);
  aoa.push([]);
  aoa.push([]);
  aoa.push([]);
  aoa.push(["NDS, TR, ANULACIONES"]);
  aoa.push(headerRow());

  for (const row of ndtrRows) {
    aoa.push(toExcelDataRow(row));
  }

  const sums = {
    mainH: sumField(mainRows, "BASE IVA"),
    mainI: sumField(mainRows, "BASE 0"),
    mainJ: sumField(mainRows, "IMPUESTOS"),
    planH: specialPlan ? parseDecimalLike(specialPlan["BASE IVA"]) : 0,
    planJ: specialPlan ? parseDecimalLike(specialPlan.IMPUESTOS) : 0,
    activoH: specialActivo ? parseDecimalLike(specialActivo["BASE IVA"]) : 0,
    activoJ: specialActivo ? parseDecimalLike(specialActivo.IMPUESTOS) : 0,
    rimpeI: sumField(rimpeRows, "BASE 0"),
  };

  return {
    aoa,
    meta: {
      mainStartRow,
      mainEndRow,
      total1Row,
      mayorIvaRow,
      specialPlanRow,
      specialActivoRow,
      totalAtsRow,
      rimpeLabelRow,
      rimpeHeaderRow,
      rimpeStartRow,
      rimpeEndRow,
      rimpeSubtotalRow,
      ndtrLabelRow,
      ndtrHeaderRow,
      sums,
      counts: {
        total: rows.length,
        main: mainRows.length,
        rimpe: rimpeRows.length,
        ndtr: ndtrRows.length,
      },
    },
  };
}

function autosizeColumns(aoa) {
  const maxCols = aoa.reduce((max, row) => Math.max(max, row.length), 0);
  const cols = [];
  for (let c = 0; c < maxCols; c += 1) {
    let maxLen = 8;
    for (const row of aoa) {
      const cell = row[c] == null ? "" : String(row[c]);
      if (cell.length > maxLen) {
        maxLen = cell.length;
      }
    }
    cols.push({ wch: Math.min(Math.max(maxLen + 2, 10), 60) });
  }
  return cols;
}

function roundCurrency(value) {
  return Math.round((Number(value || 0) + Number.EPSILON) * 100) / 100;
}

function normalizeOutputCellValue(value) {
  if (value == null || value === "") {
    return null;
  }
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }
  const clean = sanitizeText(value);
  return clean || null;
}

function writeAoaToWorksheet(ws, aoa) {
  for (let rowIndex = 0; rowIndex < aoa.length; rowIndex += 1) {
    const row = ws.getRow(rowIndex + 1);
    const sourceRow = aoa[rowIndex] || [];
    for (let colIndex = 0; colIndex < 13; colIndex += 1) {
      const value = normalizeOutputCellValue(sourceRow[colIndex]);
      if (value !== null) {
        row.getCell(colIndex + 1).value = value;
      }
    }
  }
}

function applyComputedTotals(ws, meta) {
  const { sums } = meta;
  const setNumeric = (address, value) => {
    ws.getCell(address).value = roundCurrency(value);
  };

  setNumeric(`H${meta.total1Row}`, sums.mainH);
  setNumeric(`I${meta.total1Row}`, sums.mainI);
  setNumeric(`J${meta.total1Row}`, sums.mainJ);
  setNumeric(`J${meta.mayorIvaRow}`, sums.mainJ + sums.activoJ);
  setNumeric(`K${meta.mayorIvaRow}`, sums.mainJ + sums.activoJ);
  setNumeric(`H${meta.totalAtsRow}`, sums.mainH + sums.planH + sums.activoH);
  setNumeric(`I${meta.totalAtsRow}`, sums.mainI + sums.rimpeI);
  setNumeric(`J${meta.totalAtsRow}`, sums.mainJ + sums.planJ + sums.activoJ);
  setNumeric(`I${meta.rimpeSubtotalRow}`, sums.rimpeI);
}

function getFallbackVisualStyles() {
  const header = Array.from({ length: 13 }, () => ({
    font: { bold: true, size: 11, name: "Aptos Narrow" },
    fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9E1F2" } },
    alignment: { horizontal: "center" },
    border: {
      top: { style: "dotted" },
      left: { style: "dotted" },
      bottom: { style: "dotted" },
      right: { style: "dotted" },
    },
  }));

  const data = Array.from({ length: 13 }, (_, idx) => ({
    font: { size: 11, name: "Aptos Narrow" },
    border: {
      top: { style: "dotted" },
      left: { style: "dotted" },
      bottom: { style: "dotted" },
      right: { style: "dotted" },
    },
    ...(idx === 3 ? { numFmt: "mm-dd-yy" } : {}),
  }));

  const sectionLabel = { font: { bold: true, size: 14, name: "Aptos Narrow" } };
  const totalStyle = {
    font: { size: 11, name: "Aptos Narrow" },
    numFmt: '_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ ',
  };

  return {
    header,
    data,
    sectionLabel,
    totalH: totalStyle,
    totalI: totalStyle,
    totalJ: totalStyle,
    mayorJ: totalStyle,
    mayorK: totalStyle,
    atsH: totalStyle,
    atsI: totalStyle,
    atsJ: totalStyle,
    rimpeI: totalStyle,
    columnWidths: Array.from({ length: 13 }, () => undefined),
  };
}

async function getTemplateVisualStyles(templatePath) {
  if (!fs.existsSync(templatePath)) {
    return getFallbackVisualStyles();
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(templatePath);
  const ws = wb.getWorksheet("LIBRO COMPRAS");
  if (!ws) {
    return getFallbackVisualStyles();
  }

  const header = [];
  const data = [];
  const columnWidths = [];
  for (let c = 1; c <= 13; c += 1) {
    header.push(deepClone(ws.getRow(1).getCell(c).style));
    data.push(deepClone(ws.getRow(2).getCell(c).style));
    columnWidths.push(ws.getColumn(c).width);
  }

  return {
    header,
    data,
    sectionLabel: deepClone(ws.getCell("A487").style),
    totalH: deepClone(ws.getCell("H479").style),
    totalI: deepClone(ws.getCell("I479").style),
    totalJ: deepClone(ws.getCell("J479").style),
    mayorJ: deepClone(ws.getCell("J480").style),
    mayorK: deepClone(ws.getCell("K480").style),
    atsH: deepClone(ws.getCell("H485").style),
    atsI: deepClone(ws.getCell("I485").style),
    atsJ: deepClone(ws.getCell("J485").style),
    rimpeI: deepClone(ws.getCell("I493").style),
    columnWidths,
    sheetViews: deepClone(ws.views),
    pageSetup: deepClone(ws.pageSetup),
    headerFooter: deepClone(ws.headerFooter),
    sheetProperties: deepClone(ws.properties),
    sheetState: ws.state || "visible",
  };
}

async function buildStyledWorkbook(templatePath, aoa, meta) {
  const styles = await getTemplateVisualStyles(templatePath);
  const wb = new ExcelJS.Workbook();
  wb.calcProperties = { fullCalcOnLoad: true };
  wb.creator = "Codex";
  wb.created = new Date();
  const ws = wb.addWorksheet("LIBRO COMPRAS");

  if (styles.sheetProperties) {
    ws.properties = deepClone(styles.sheetProperties);
  }
  if (styles.sheetViews) {
    ws.views = deepClone(styles.sheetViews);
  }
  if (styles.pageSetup) {
    ws.pageSetup = deepClone(styles.pageSetup);
  }
  if (styles.headerFooter) {
    ws.headerFooter = deepClone(styles.headerFooter);
  }
  ws.state = styles.sheetState || "visible";

  writeAoaToWorksheet(ws, aoa);
  applyComputedTotals(ws, meta);

  styles.columnWidths.forEach((width, idx) => {
    if (width) {
      ws.getColumn(idx + 1).width = width;
    }
  });

  // Prevent "#######" rendering in date/accounting columns when content is wider.
  const minWidths = new Map([
    [4, 12],  // FECHA
    [7, 14],  // MONTO
    [8, 14],  // BASE IVA
    [9, 14],  // BASE 0
    [10, 14], // IMPUESTOS
    [11, 14], // RETENCION
    [12, 18], // SALDO / MAYOR IVA labels
  ]);
  for (const [colIdx, minWidth] of minWidths) {
    const col = ws.getColumn(colIdx);
    const current = Number(col.width || 0);
    if (!Number.isFinite(current) || current < minWidth) {
      col.width = minWidth;
    }
  }

  const applyRowStyle = (rowNum, stylesByCol) => {
    if (rowNum < 1 || rowNum > ws.rowCount) {
      return;
    }
    const row = ws.getRow(rowNum);
    for (let c = 1; c <= 13; c += 1) {
      row.getCell(c).style = deepClone(stylesByCol[c - 1] || {});
    }
  };

  const applyRangeStyle = (startRow, endRow, stylesByCol) => {
    for (let r = startRow; r <= endRow; r += 1) {
      applyRowStyle(r, stylesByCol);
    }
  };

  applyRowStyle(1, styles.header);
  applyRowStyle(meta.rimpeHeaderRow, styles.header);
  applyRowStyle(meta.ndtrHeaderRow, styles.header);

  applyRangeStyle(meta.mainStartRow, meta.mainEndRow, styles.data);
  applyRangeStyle(meta.specialPlanRow, meta.specialActivoRow, styles.data);
  applyRangeStyle(meta.rimpeStartRow, meta.rimpeEndRow, styles.data);
  applyRangeStyle(meta.ndtrHeaderRow + 1, ws.rowCount, styles.data);

  ws.getCell(`A${meta.rimpeLabelRow}`).style = deepClone(styles.sectionLabel);
  ws.getCell(`A${meta.ndtrLabelRow}`).style = deepClone(styles.sectionLabel);

  ws.getCell(`H${meta.total1Row}`).style = deepClone(styles.totalH);
  ws.getCell(`I${meta.total1Row}`).style = deepClone(styles.totalI);
  ws.getCell(`J${meta.total1Row}`).style = deepClone(styles.totalJ);

  ws.getCell(`J${meta.mayorIvaRow}`).style = deepClone(styles.mayorJ);
  ws.getCell(`K${meta.mayorIvaRow}`).style = deepClone(styles.mayorK);

  ws.getCell(`H${meta.totalAtsRow}`).style = deepClone(styles.atsH);
  ws.getCell(`I${meta.totalAtsRow}`).style = deepClone(styles.atsI);
  ws.getCell(`J${meta.totalAtsRow}`).style = deepClone(styles.atsJ);

  ws.getCell(`I${meta.rimpeSubtotalRow}`).style = deepClone(styles.rimpeI);

  return wb;
}

function verifyOutputWorkbook(outputPath, meta) {
  const wb = XLSX.readFile(outputPath, { cellFormula: true });
  if (wb.SheetNames.length !== 1 || wb.SheetNames[0] !== "LIBRO COMPRAS") {
    throw new Error("Validacion final: el archivo debe tener una sola hoja llamada LIBRO COMPRAS.");
  }

  const ws = wb.Sheets["LIBRO COMPRAS"];
  const { sums } = meta;
  const expectedValues = new Map([
    [`H${meta.total1Row}`, roundCurrency(sums.mainH)],
    [`I${meta.total1Row}`, roundCurrency(sums.mainI)],
    [`J${meta.total1Row}`, roundCurrency(sums.mainJ)],
    [`J${meta.mayorIvaRow}`, roundCurrency(sums.mainJ + sums.activoJ)],
    [`K${meta.mayorIvaRow}`, roundCurrency(sums.mainJ + sums.activoJ)],
    [`H${meta.totalAtsRow}`, roundCurrency(sums.mainH + sums.planH + sums.activoH)],
    [`I${meta.totalAtsRow}`, roundCurrency(sums.mainI + sums.rimpeI)],
    [`J${meta.totalAtsRow}`, roundCurrency(sums.mainJ + sums.planJ + sums.activoJ)],
    [`I${meta.rimpeSubtotalRow}`, roundCurrency(sums.rimpeI)],
  ]);

  for (const [address, expectedValue] of expectedValues) {
    const cell = ws[address];
    if (cell && cell.f) {
      throw new Error(`Validacion final: ${address} no debe quedar como formula.`);
    }
    const found = cell && typeof cell.v === "number" ? roundCurrency(cell.v) : NaN;
    if (!Number.isFinite(found) || Math.abs(found - expectedValue) > 0.01) {
      throw new Error(`Validacion final: valor incorrecto en ${address}. Esperado: ${expectedValue}. Encontrado: ${cell ? cell.v : "vacio"}.`);
    }
  }
}

function writeAuditReport(auditPath, payload) {
  fs.mkdirSync(path.dirname(auditPath), { recursive: true });
  fs.writeFileSync(auditPath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
}

async function writeWorkbookWithRetries(workbook, preferredPath, maxAttempts = 20) {
  const parsed = path.parse(preferredPath);
  fs.mkdirSync(parsed.dir || ".", { recursive: true });
  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    const candidate =
      attempt === 0
        ? preferredPath
        : path.join(parsed.dir, `${parsed.name}_nuevo${attempt === 1 ? "" : `_${attempt}`}${parsed.ext}`);
    try {
      await workbook.xlsx.writeFile(candidate);
      return candidate;
    } catch (error) {
      const isLocked = error && (error.code === "EBUSY" || error.code === "EPERM");
      if (!isLocked) {
        throw error;
      }
    }
  }
  throw new Error(`No se pudo guardar el Excel. Cierra los archivos abiertos y vuelve a intentar.`);
}

async function extractRowsFromPdf(pdfPath) {
  const data = new Uint8Array(fs.readFileSync(pdfPath));
  const standardFontDir = path.join(__dirname, "node_modules", "pdfjs-dist", "standard_fonts");
  const standardFontDataUrl = `${standardFontDir.replace(/\\/g, "/")}/`;

  const loadingTask = pdfjsLib.getDocument({
    data,
    standardFontDataUrl,
    verbosity: pdfjsLib.VerbosityLevel.ERRORS,
  });

  const pdf = await loadingTask.promise;
  const allRows = [];

  try {
    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum += 1) {
      const page = await pdf.getPage(pageNum);
      const content = await page.getTextContent();

      const items = content.items
        .map((it) => ({
          str: sanitizeText(it.str),
          x: Number(it.transform[4]),
          y: Number(it.transform[5]),
        }))
        .filter((it) => it.str.length > 0);

      const headerCandidates = items.filter((it) => HEADER_ALIASES[it.str]);
      if (headerCandidates.length < 6) {
        continue;
      }

      const headerMap = new Map();
      for (const candidate of headerCandidates) {
        const canonical = HEADER_ALIASES[candidate.str];
        if (!headerMap.has(canonical)) {
          headerMap.set(canonical, candidate.x);
        }
      }

      const hasEnoughHeaders = COLUMN_ORDER.filter((name) => headerMap.has(name)).length >= 10;
      if (!hasEnoughHeaders) {
        continue;
      }

      const headerY = headerCandidates[0].y;
      const boundaries = FIXED_BOUNDARIES;
      const belowHeader = items.filter((it) => it.y < headerY - 1);
      const groupedRows = groupItemsByRow(belowHeader);

      for (const groupedRow of groupedRows) {
        if (!isDataRow(groupedRow.items, boundaries)) {
          continue;
        }

        const row = {};
        for (const col of COLUMN_ORDER) {
          row[col] = "";
        }

        const leftItems = groupedRow.items.filter((item) => item.x < 400);
        const numericItems = groupedRow.items.filter((item) => item.x >= 400).sort((a, b) => a.x - b.x);

        for (const item of leftItems) {
          const colName = assignToColumn(item.x, LEFT_BOUNDARIES);
          if (!colName) {
            continue;
          }
          row[colName] = row[colName] ? `${row[colName]} ${item.str}` : item.str;
        }

        if (numericItems.length >= NUMERIC_COLUMNS.length) {
          const selected = numericItems.slice(-NUMERIC_COLUMNS.length);
          for (let i = 0; i < NUMERIC_COLUMNS.length; i += 1) {
            row[NUMERIC_COLUMNS[i]] = selected[i].str;
          }
        } else {
          for (const item of numericItems) {
            const colName = assignToColumn(item.x, boundaries);
            if (!colName) {
              continue;
            }
            row[colName] = row[colName] ? `${row[colName]} ${item.str}` : item.str;
          }
        }

        const hasDate = /^\d{2}\/\d{2}\/\d{4}$/.test(row.FECHA);
        const hasTipo = /^[A-Z]{2,3}$/.test(row.TIPO);
        const hasDoc = row.DOCUMENTO.length > 0;
        if (hasDate && hasTipo && hasDoc) {
          allRows.push(row);
        }
      }
    }
  } finally {
    await loadingTask.destroy();
  }

  return allRows;
}

async function main() {
  const inputPath = path.resolve(process.cwd(), INPUT_PDF);
  const outputPath = path.resolve(process.cwd(), OUTPUT_XLSX);
  const templatePath = path.resolve(process.cwd(), TEMPLATE_XLSX);

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
    const sampleA = consistencyAudit.extraGenerated.slice(0, 3).map((x) => x.signature).join(" || ");
    const sampleB = consistencyAudit.missingGenerated.slice(0, 3).map((x) => x.signature).join(" || ");
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

  console.log(`PDF leido: ${inputPath}`);
  console.log(`Filas extraidas: ${rows.length}`);
  console.log(`Filas ajustadas con plantilla: ${overrideCount}`);
  console.log(`Cobertura plantilla: ${(overrideCoverage * 100).toFixed(2)}%`);
  console.log(`Problemas detectados antes de ajuste: ${preValidationProblems.length}`);
  console.log(`Bloque principal: ${meta.counts.main}`);
  console.log(`RIMPE: ${meta.counts.rimpe}`);
  console.log(`NDS/TR/ANULACIONES: ${meta.counts.ndtr}`);
  console.log(`Excel generado (una sola hoja): ${finalOutputPath}`);
  console.log(`Auditoria JSON: ${auditPath}`);
}

main().catch((error) => {
  console.error("Error:", error.message);
  process.exit(1);
});
