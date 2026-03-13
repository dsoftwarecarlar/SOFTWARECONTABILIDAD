const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");

const {
  roundCurrency,
  deepClone,
  normalizeOutputCellValue,
} = require("./row-utils");
const { getTemplateVisualStyles } = require("./template");

function writeAoaToWorksheet(worksheet, aoa) {
  for (let rowIndex = 0; rowIndex < aoa.length; rowIndex += 1) {
    const row = worksheet.getRow(rowIndex + 1);
    const sourceRow = aoa[rowIndex] || [];
    for (let colIndex = 0; colIndex < 13; colIndex += 1) {
      const value = normalizeOutputCellValue(sourceRow[colIndex]);
      if (value !== null) {
        row.getCell(colIndex + 1).value = value;
      }
    }
  }
}

function applyComputedTotals(worksheet, meta) {
  const { sums } = meta;
  const setNumeric = (address, value) => {
    worksheet.getCell(address).value = roundCurrency(value);
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

async function buildStyledWorkbook(templatePath, aoa, meta) {
  const styles = await getTemplateVisualStyles(templatePath);
  const workbook = new ExcelJS.Workbook();
  workbook.calcProperties = { fullCalcOnLoad: true };
  workbook.creator = "Codex";
  workbook.created = new Date();
  const worksheet = workbook.addWorksheet("LIBRO COMPRAS");

  if (styles.sheetProperties) {
    worksheet.properties = deepClone(styles.sheetProperties);
  }
  if (styles.sheetViews) {
    worksheet.views = deepClone(styles.sheetViews);
  }
  if (styles.pageSetup) {
    worksheet.pageSetup = deepClone(styles.pageSetup);
  }
  if (styles.headerFooter) {
    worksheet.headerFooter = deepClone(styles.headerFooter);
  }
  worksheet.state = styles.sheetState || "visible";

  writeAoaToWorksheet(worksheet, aoa);
  applyComputedTotals(worksheet, meta);

  styles.columnWidths.forEach((width, idx) => {
    if (width) {
      worksheet.getColumn(idx + 1).width = width;
    }
  });

  const minWidths = new Map([
    [4, 12],
    [7, 14],
    [8, 14],
    [9, 14],
    [10, 14],
    [11, 14],
    [12, 18],
  ]);
  for (const [colIdx, minWidth] of minWidths) {
    const column = worksheet.getColumn(colIdx);
    const current = Number(column.width || 0);
    if (!Number.isFinite(current) || current < minWidth) {
      column.width = minWidth;
    }
  }

  const applyRowStyle = (rowNum, stylesByCol) => {
    if (rowNum < 1 || rowNum > worksheet.rowCount) {
      return;
    }
    const row = worksheet.getRow(rowNum);
    for (let col = 1; col <= 13; col += 1) {
      row.getCell(col).style = deepClone(stylesByCol[col - 1] || {});
    }
  };

  const applyRangeStyle = (startRow, endRow, stylesByCol) => {
    for (let row = startRow; row <= endRow; row += 1) {
      applyRowStyle(row, stylesByCol);
    }
  };

  applyRowStyle(1, styles.header);
  applyRowStyle(meta.rimpeHeaderRow, styles.header);
  applyRowStyle(meta.ndtrHeaderRow, styles.header);

  applyRangeStyle(meta.mainStartRow, meta.mainEndRow, styles.data);
  applyRangeStyle(meta.specialPlanRow, meta.specialActivoRow, styles.data);
  applyRangeStyle(meta.rimpeStartRow, meta.rimpeEndRow, styles.data);
  applyRangeStyle(meta.ndtrHeaderRow + 1, worksheet.rowCount, styles.data);

  worksheet.getCell(`A${meta.rimpeLabelRow}`).style = deepClone(styles.sectionLabel);
  worksheet.getCell(`A${meta.ndtrLabelRow}`).style = deepClone(styles.sectionLabel);

  worksheet.getCell(`H${meta.total1Row}`).style = deepClone(styles.totalH);
  worksheet.getCell(`I${meta.total1Row}`).style = deepClone(styles.totalI);
  worksheet.getCell(`J${meta.total1Row}`).style = deepClone(styles.totalJ);
  worksheet.getCell(`J${meta.mayorIvaRow}`).style = deepClone(styles.mayorJ);
  worksheet.getCell(`K${meta.mayorIvaRow}`).style = deepClone(styles.mayorK);
  worksheet.getCell(`H${meta.totalAtsRow}`).style = deepClone(styles.atsH);
  worksheet.getCell(`I${meta.totalAtsRow}`).style = deepClone(styles.atsI);
  worksheet.getCell(`J${meta.totalAtsRow}`).style = deepClone(styles.atsJ);
  worksheet.getCell(`I${meta.rimpeSubtotalRow}`).style = deepClone(styles.rimpeI);

  return workbook;
}

function verifyOutputWorkbook(outputPath, meta) {
  const workbook = XLSX.readFile(outputPath, { cellFormula: true });
  if (workbook.SheetNames.length !== 1 || workbook.SheetNames[0] !== "LIBRO COMPRAS") {
    throw new Error("Validacion final: el archivo debe tener una sola hoja llamada LIBRO COMPRAS.");
  }

  const worksheet = workbook.Sheets["LIBRO COMPRAS"];
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
    const cell = worksheet[address];
    if (cell && cell.f) {
      throw new Error(`Validacion final: ${address} no debe quedar como formula.`);
    }
    const found = cell && typeof cell.v === "number" ? roundCurrency(cell.v) : NaN;
    if (!Number.isFinite(found) || Math.abs(found - expectedValue) > 0.01) {
      throw new Error(`Validacion final: valor incorrecto en ${address}. Esperado: ${expectedValue}. Encontrado: ${cell ? cell.v : "vacio"}.`);
    }
  }
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

  throw new Error("No se pudo guardar el Excel. Cierra los archivos abiertos y vuelve a intentar.");
}

module.exports = {
  buildStyledWorkbook,
  verifyOutputWorkbook,
  writeWorkbookWithRetries,
};
