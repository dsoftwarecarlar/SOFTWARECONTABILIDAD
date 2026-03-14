const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const OUTPUTS_DIR = path.join(__dirname, "storage", "outputs");
const DEFAULT_OUTPUT_XLSX = "acciones_resumen.xlsx";

const ACTIONS = [
  {
    key: "accion1",
    label: "Accion 1",
    sheetName: "ACCION 1 LIBRO COMPRAS",
    matches: (name) => /_resultado/i.test(name),
  },
  {
    key: "accion2",
    label: "Accion 2",
    sheetName: "ACCION 2 RET PROV",
    matches: (name) => /_\d{8}_\d{6}_accion2(?:_nuevo(?:_\d+)?)?\.(xlsx|xls)$/i.test(name),
  },
  {
    key: "accion3",
    label: "Accion 3",
    sheetName: "ACCION 3 MAYOR RET",
    matches: (name) => /_\d{8}_\d{6}_accion3(?:_nuevo(?:_\d+)?)?\.(xlsx|xls)$/i.test(name),
  },
  {
    key: "accion4",
    label: "Accion 4",
    sheetName: "ACCION 4 MAYOR IVA",
    matches: (name) => /accion4/i.test(name),
  },
];

function cloneValue(value) {
  if (value instanceof Date) {
    return new Date(value.getTime());
  }
  if (Array.isArray(value)) {
    return value.map(cloneValue);
  }
  if (value && typeof value === "object") {
    const result = {};
    for (const [key, item] of Object.entries(value)) {
      result[key] = cloneValue(item);
    }
    return result;
  }
  return value;
}

function parseCliArguments(argv = process.argv.slice(2)) {
  let outputXlsx = DEFAULT_OUTPUT_XLSX;

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--output") {
      outputXlsx = argv[index + 1] || outputXlsx;
      index += 1;
      continue;
    }
  }

  return {
    outputXlsx,
  };
}

function getFileTimestamp(filePath) {
  const stats = fs.statSync(filePath);
  return Number(stats.birthtimeMs || stats.mtimeMs || 0);
}

function listOutputFiles(outputDir) {
  if (!fs.existsSync(outputDir)) {
    return [];
  }

  return fs
    .readdirSync(outputDir)
    .filter((name) => name.toLowerCase().endsWith(".xlsx"))
    .map((name) => ({
      name,
      path: path.join(outputDir, name),
      timestamp: getFileTimestamp(path.join(outputDir, name)),
    }))
    .sort((left, right) => right.timestamp - left.timestamp || right.name.localeCompare(left.name, "es"));
}

function findLatestActionFiles(outputDir) {
  const files = listOutputFiles(outputDir);
  return ACTIONS.map((action) => ({
    ...action,
    latest: files.find((file) => action.matches(file.name)) || null,
  }));
}

function getWorksheetMaxColumn(worksheet) {
  let maxCol = Math.max(worksheet.columnCount || 0, worksheet.actualColumnCount || 0);

  for (const mergeRange of worksheet.model?.merges || []) {
    const match = /:([A-Z]+)\d+$/i.exec(mergeRange);
    if (!match) {
      continue;
    }

    let col = 0;
    const letters = match[1].toUpperCase();
    for (let i = 0; i < letters.length; i += 1) {
      col = col * 26 + (letters.charCodeAt(i) - 64);
    }
    maxCol = Math.max(maxCol, col);
  }

  return maxCol;
}

function copyWorksheetLayout(sourceSheet, targetSheet, maxCol) {
  targetSheet.properties = cloneValue(sourceSheet.properties || {});
  targetSheet.pageSetup = cloneValue(sourceSheet.pageSetup || {});
  targetSheet.headerFooter = cloneValue(sourceSheet.headerFooter || {});
  targetSheet.views = cloneValue(sourceSheet.views || []);
  targetSheet.state = sourceSheet.state;

  if (sourceSheet.autoFilter) {
    targetSheet.autoFilter = cloneValue(sourceSheet.autoFilter);
  }

  for (let col = 1; col <= maxCol; col += 1) {
    const sourceColumn = sourceSheet.getColumn(col);
    const targetColumn = targetSheet.getColumn(col);
    targetColumn.width = sourceColumn.width;
    targetColumn.hidden = sourceColumn.hidden;
    targetColumn.outlineLevel = sourceColumn.outlineLevel;
    targetColumn.style = cloneValue(sourceColumn.style || {});
  }
}

function copyWorksheetRows(sourceSheet, targetSheet, maxRow, maxCol) {
  for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
    const sourceRow = sourceSheet.getRow(rowNumber);
    const targetRow = targetSheet.getRow(rowNumber);
    targetRow.height = sourceRow.height;
    targetRow.hidden = sourceRow.hidden;
    targetRow.outlineLevel = sourceRow.outlineLevel;
    targetRow.style = cloneValue(sourceRow.style || {});

    for (let col = 1; col <= maxCol; col += 1) {
      const sourceCell = sourceRow.getCell(col);
      const targetCell = targetRow.getCell(col);
      targetCell.value = cloneValue(sourceCell.value);
      targetCell.style = cloneValue(sourceCell.style || {});

      if (sourceCell.note) {
        targetCell.note = cloneValue(sourceCell.note);
      }
      if (sourceCell.dataValidation) {
        targetCell.dataValidation = cloneValue(sourceCell.dataValidation);
      }
      if (sourceCell.protection) {
        targetCell.protection = cloneValue(sourceCell.protection);
      }
    }
  }
}

function copyWorksheetMerges(sourceSheet, targetSheet) {
  for (const mergeRange of sourceSheet.model?.merges || []) {
    targetSheet.mergeCells(mergeRange);
  }
}

async function copyWorkbookSheet(targetWorkbook, sourcePath, targetSheetName) {
  const sourceWorkbook = new ExcelJS.Workbook();
  await sourceWorkbook.xlsx.readFile(sourcePath);

  const sourceSheet = sourceWorkbook.worksheets[0];
  if (!sourceSheet) {
    throw new Error(`El archivo ${path.basename(sourcePath)} no contiene hojas para copiar.`);
  }

  const targetSheet = targetWorkbook.addWorksheet(targetSheetName);
  const maxRow = Math.max(sourceSheet.rowCount || 0, 1);
  const maxCol = Math.max(getWorksheetMaxColumn(sourceSheet), 1);

  copyWorksheetLayout(sourceSheet, targetSheet, maxCol);
  copyWorksheetRows(sourceSheet, targetSheet, maxRow, maxCol);
  copyWorksheetMerges(sourceSheet, targetSheet);
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

  throw new Error("No se pudo guardar el Excel consolidado. Cierra archivos abiertos e intenta de nuevo.");
}

async function main() {
  const cli = parseCliArguments();
  const outputPath = path.resolve(process.cwd(), cli.outputXlsx);
  const latestFiles = findLatestActionFiles(OUTPUTS_DIR);
  const missing = latestFiles.filter((item) => item.latest === null);

  if (missing.length > 0) {
    throw new Error(
      `Faltan archivos generados para: ${missing.map((item) => item.label).join(", ")}.`,
    );
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "BOT1ANDREA";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;

  for (const item of latestFiles) {
    await copyWorkbookSheet(workbook, item.latest.path, item.sheetName);
  }

  const finalOutputPath = await writeWorkbookWithRetries(workbook, outputPath);
  console.log(`Excel consolidado generado: ${finalOutputPath}`);
  for (const item of latestFiles) {
    console.log(`${item.label}: ${item.latest.path}`);
  }
}

main().catch((error) => {
  console.error("Error:", error.message);
  process.exit(1);
});
