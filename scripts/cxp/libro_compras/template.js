const fs = require("fs");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");

const {
  sanitizeText,
  normalizeNumericText,
  normalizeDocument,
  deepClone,
} = require("./row-utils");

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

  const totalStyle = {
    font: { size: 11, name: "Aptos Narrow" },
    numFmt: '_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ ',
  };

  return {
    header,
    data,
    sectionLabel: { font: { bold: true, size: 14, name: "Aptos Narrow" } },
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

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);
  const worksheet = workbook.getWorksheet("LIBRO COMPRAS");
  if (!worksheet) {
    return getFallbackVisualStyles();
  }

  const header = [];
  const data = [];
  const columnWidths = [];
  for (let column = 1; column <= 13; column += 1) {
    header.push(deepClone(worksheet.getRow(1).getCell(column).style));
    data.push(deepClone(worksheet.getRow(2).getCell(column).style));
    columnWidths.push(worksheet.getColumn(column).width);
  }

  return {
    header,
    data,
    sectionLabel: deepClone(worksheet.getCell("A487").style),
    totalH: deepClone(worksheet.getCell("H479").style),
    totalI: deepClone(worksheet.getCell("I479").style),
    totalJ: deepClone(worksheet.getCell("J479").style),
    mayorJ: deepClone(worksheet.getCell("J480").style),
    mayorK: deepClone(worksheet.getCell("K480").style),
    atsH: deepClone(worksheet.getCell("H485").style),
    atsI: deepClone(worksheet.getCell("I485").style),
    atsJ: deepClone(worksheet.getCell("J485").style),
    rimpeI: deepClone(worksheet.getCell("I493").style),
    columnWidths,
    sheetViews: deepClone(worksheet.views),
    pageSetup: deepClone(worksheet.pageSetup),
    headerFooter: deepClone(worksheet.headerFooter),
    sheetProperties: deepClone(worksheet.properties),
    sheetState: worksheet.state || "visible",
  };
}

module.exports = {
  loadTemplateOverrides,
  getTemplateVisualStyles,
};
