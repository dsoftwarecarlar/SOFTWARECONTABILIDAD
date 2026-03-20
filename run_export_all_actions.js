const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const { normalizeWorkbookOpenXmlCompatibility } = require("./scripts/cxp/shared/excel-template-utils");

const OUTPUTS_DIR = path.join(__dirname, "storage", "outputs");
const ACTIONS_CONFIG_PATH = path.join(__dirname, "config", "cxp", "action_exports.json");
const DEFAULT_OUTPUT_XLSX = "acciones_resumen.xlsx";

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

function normalizeExtension(value) {
  return String(value || "")
    .trim()
    .replace(/^\./, "")
    .toLowerCase();
}

function loadActionDefinitions() {
  if (!fs.existsSync(ACTIONS_CONFIG_PATH)) {
    throw new Error("No existe config/cxp/action_exports.json en el proyecto.");
  }

  const parsed = JSON.parse(fs.readFileSync(ACTIONS_CONFIG_PATH, "utf8"));
  if (!Array.isArray(parsed) || parsed.length === 0) {
    throw new Error("config/cxp/action_exports.json no contiene acciones exportables validas.");
  }

  return parsed.map((item) => ({
    key: String(item.key || "").trim(),
    label: String(item.label || item.key || "").trim(),
    sheetName: String(item.sheet_name || item.key || "").trim(),
    bundleExtensions: Array.isArray(item.bundle_extensions)
      ? item.bundle_extensions.map(normalizeExtension).filter(Boolean)
      : ["xlsx"],
    fileMatch: item && typeof item.file_match === "object" ? item.file_match : {},
  })).filter((item) => item.key && item.sheetName);
}

function matchesActionFile(action, fileName) {
  const rule = action.fileMatch || {};
  const type = String(rule.type || "").trim().toLowerCase();
  const value = String(rule.value || "");

  if (type === "contains") {
    return value !== "" && fileName.toLowerCase().includes(value.toLowerCase());
  }

  if (type === "regex") {
    if (value === "") {
      return false;
    }
    const flags = String(rule.flags || "").replace(/[^dgimsuvy]/g, "");
    return new RegExp(value, flags).test(fileName);
  }

  return false;
}

function collectBundleExtensions(actions) {
  const extensions = new Set();
  for (const action of actions) {
    for (const extension of action.bundleExtensions) {
      if (extension) {
        extensions.add(extension);
      }
    }
  }

  return extensions.size > 0 ? extensions : new Set(["xlsx"]);
}

function getFileTimestamp(filePath) {
  const stats = fs.statSync(filePath);
  return Number(stats.birthtimeMs || stats.mtimeMs || 0);
}

function listOutputFiles(outputDir, allowedExtensions) {
  if (!fs.existsSync(outputDir)) {
    return [];
  }

  return fs
    .readdirSync(outputDir)
    .filter((name) => allowedExtensions.has(normalizeExtension(path.extname(name))))
    .map((name) => ({
      name,
      path: path.join(outputDir, name),
      timestamp: getFileTimestamp(path.join(outputDir, name)),
    }))
    .sort((left, right) => right.timestamp - left.timestamp || right.name.localeCompare(left.name, "es"));
}

function findLatestActionFiles(outputDir, actions) {
  const files = listOutputFiles(outputDir, collectBundleExtensions(actions));
  return actions.map((action) => ({
    ...action,
    latest: files.find((file) => matchesActionFile(action, file.name)) || null,
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

function ensureWorkbookViews(targetWorkbook, sourceViews = []) {
  if (Array.isArray(targetWorkbook.views) && targetWorkbook.views.length > 0) {
    return;
  }

  const sourceView =
    Array.isArray(sourceViews) && sourceViews.length > 0 && sourceViews[0] && typeof sourceViews[0] === "object"
      ? sourceViews[0]
      : null;
  const hasSheetViews = targetWorkbook.worksheets.some(
    (worksheet) => Array.isArray(worksheet.views) && worksheet.views.length > 0,
  );

  if (!sourceView && !hasSheetViews) {
    return;
  }

  targetWorkbook.views = [
    {
      x: Number.isFinite(sourceView?.x) ? sourceView.x : 0,
      y: Number.isFinite(sourceView?.y) ? sourceView.y : 0,
      width: Number.isFinite(sourceView?.width) ? sourceView.width : 20000,
      height: Number.isFinite(sourceView?.height) ? sourceView.height : 12000,
      firstSheet: Number.isInteger(sourceView?.firstSheet) ? sourceView.firstSheet : 0,
      activeTab: Number.isInteger(sourceView?.activeTab) ? sourceView.activeTab : 0,
      visibility:
        typeof sourceView?.visibility === "string" && sourceView.visibility.trim() !== ""
          ? sourceView.visibility
          : "visible",
    },
  ];
}

async function copyWorkbookSheet(targetWorkbook, sourcePath, targetSheetName) {
  const sourceWorkbook = new ExcelJS.Workbook();
  await sourceWorkbook.xlsx.readFile(sourcePath);

  const sourceSheet = sourceWorkbook.worksheets[0];
  if (!sourceSheet) {
    throw new Error(`El archivo ${path.basename(sourcePath)} no contiene hojas para copiar.`);
  }

  ensureWorkbookViews(targetWorkbook, sourceWorkbook.views);

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
      normalizeWorkbookOpenXmlCompatibility(candidate);
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
  const actions = loadActionDefinitions();
  const latestFiles = findLatestActionFiles(OUTPUTS_DIR, actions);
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

  ensureWorkbookViews(workbook);

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
