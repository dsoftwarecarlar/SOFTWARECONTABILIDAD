const path = require("path");
const fs = require("fs");
const ExcelJS = require("exceljs");
const XLSX = require("xlsx");

const { sanitizeText } = require("../shared/core-utils");

const SHEET_CONFIGS = [
  { key: "tyt", label: "MATRIZ", argKeys: ["inputtyt", "input-tyt"], targetSheet: "REP TYT" },
  { key: "peug", label: "PEUGEOT", argKeys: ["inputpeug", "input-peug"], targetSheet: "REP PEUGT" },
  { key: "chgn", label: "CHANGAN", argKeys: ["inputchgn", "input-chgn"], targetSheet: "REP CHGN" },
  { key: "szk", label: "SUZUKI", argKeys: ["inputszk", "input-szk"], targetSheet: "REP SZK" },
];

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

function readSourceSheet(sourcePath) {
  const workbook = XLSX.readFile(sourcePath, {
    cellDates: false,
    cellFormula: true,
    cellNF: true,
    cellText: true,
    sheetStubs: true,
  });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error(`El archivo ${sourcePath} no contiene hojas.`);
  }

  return workbook.Sheets[sheetName];
}

function sourceCellText(sheet, row, column) {
  const address = XLSX.utils.encode_cell({ r: row - 1, c: column - 1 });
  const cell = sheet[address];
  if (!cell) {
    return "";
  }

  if (cell.w != null) {
    return sanitizeText(cell.w);
  }

  if (cell.v == null) {
    return "";
  }

  return sanitizeText(cell.v);
}

function countPayloadRows(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:AO1");
  let count = 0;
  for (let row = 11; row <= range.e.r + 1; row += 1) {
    const document = sourceCellText(sheet, row, 5);
    if (!document) {
      continue;
    }

    if (/^TOTAL\s+GENERAL/i.test(document)) {
      break;
    }

    count += 1;
  }

  return count;
}

function convertSourceCell(cell) {
  if (!cell) {
    return null;
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

function clearTargetTail(worksheet, startRow, endRow, lastColumn) {
  if (startRow > endRow) {
    return;
  }

  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    for (let column = 1; column <= lastColumn; column += 1) {
      const cell = row.getCell(column);
      if (cell.isMerged && cell.master && cell.master.address !== cell.address) {
        continue;
      }
      cell.value = null;
    }
    row.commit();
  }
}

function copySourceSheetToTarget(sourceSheet, targetSheet) {
  const range = XLSX.utils.decode_range(sourceSheet["!ref"] || "A1:AO1");
  const lastRow = range.e.r + 1;
  const lastColumn = Math.max(41, range.e.c + 1);
  const clearUntil = Math.max(lastRow, targetSheet.actualRowCount || 0);

  for (let rowNumber = 1; rowNumber <= lastRow; rowNumber += 1) {
    const row = targetSheet.getRow(rowNumber);
    for (let column = 1; column <= lastColumn; column += 1) {
      const targetCell = row.getCell(column);
      if (targetCell.isMerged && targetCell.master && targetCell.master.address !== targetCell.address) {
        continue;
      }
      const address = XLSX.utils.encode_cell({ r: rowNumber - 1, c: column - 1 });
      targetCell.value = convertSourceCell(sourceSheet[address]);
    }
    row.commit();
  }

  clearTargetTail(targetSheet, lastRow + 1, clearUntil, lastColumn);
}

async function writeWorkbookToRequestedPath(workbook, outputPath, maxAttempts = 8) {
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      await workbook.xlsx.writeFile(outputPath);
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

  for (const config of SHEET_CONFIGS) {
    const sourcePath = path.resolve(process.cwd(), getArg(args, config.argKeys));
    if (!sourcePath || !fs.existsSync(sourcePath)) {
      throw new Error(`No existe archivo de entrada para ${config.label}: ${sourcePath}`);
    }

    const sourceSheet = readSourceSheet(sourcePath);
    const targetSheet = workbook.getWorksheet(config.targetSheet);
    if (!targetSheet) {
      throw new Error(`No existe la hoja requerida en plantilla: ${config.targetSheet}`);
    }

    copySourceSheetToTarget(sourceSheet, targetSheet);
    console.log(`INFO|${config.key}|rows=${countPayloadRows(sourceSheet)}`);
    console.log(`INFO|${config.key}|sheet=${config.targetSheet}`);
    console.log(`INFO|${config.key}|label=${config.label}`);
  }

  const finalPath = await writeWorkbookToRequestedPath(workbook, outputPath);
  console.log(`OUTPUT|${path.basename(finalPath)}|FACTURACION REPUESTOS TYTSERV`);
}

if (require.main === module) {
  main().catch((error) => {
    console.error(error && error.message ? error.message : String(error));
    process.exit(1);
  });
}
