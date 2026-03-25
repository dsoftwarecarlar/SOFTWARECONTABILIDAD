const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");

const XLSX = require("xlsx");

const ROOT = path.resolve(__dirname, "..", "..");
const PYTHON = process.env.PYTHON_BINARY || "python";
const REP_LAST_COLUMN = 41;

const REP_CONFIGS = [
  {
    key: "tyt",
    field: "excel_tyt",
    label: "TOYOTA",
    targetSheet: "REP TYT",
    sourceFile: "RepLibroVentasGeneral.xlsx",
    extraRows: 7,
  },
  {
    key: "peug",
    field: "excel_peug",
    label: "PEUGEOT",
    targetSheet: "REP PEUGT",
    sourceFile: "RepLibroVentasGeneral (1).xlsx",
    extraRows: 9,
  },
  {
    key: "chgn",
    field: "excel_chgn",
    label: "CHANGAN",
    targetSheet: "REP CHGN",
    sourceFile: "RepLibroVentasGeneral (2).xlsx",
    extraRows: 11,
  },
  {
    key: "szk",
    field: "excel_szk",
    label: "SUZUKI",
    targetSheet: "REP SZK",
    sourceFile: "RepLibroVentasGeneral (3).xlsx",
    extraRows: 27,
  },
];

function assertCondition(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

function fixturePath(relativePath) {
  return path.join(ROOT, relativePath);
}

function ensureDir(relativePath) {
  const directory = fixturePath(relativePath);
  fs.mkdirSync(directory, { recursive: true });
  return directory;
}

function buildTempDir() {
  const directory = path.join(
    ensureDir(path.join("storage", "verify_runs", "__tmp")),
    `repuestos-rep-capacity-${Date.now()}-${Math.random().toString(16).slice(2)}`,
  );
  fs.mkdirSync(directory, { recursive: true });
  return directory;
}

function buildManifestPath(tempDir) {
  return path.join(tempDir, "manifest.json");
}

function buildOutputPath(tempDir) {
  return path.join(tempDir, "repuestos_rep_stage_capacity.xlsx");
}

function findRowContaining(sheet, needle, lastColumn = REP_LAST_COLUMN) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  const expected = String(needle || "").trim().toUpperCase();
  for (let row = 1; row <= range.e.r + 1; row += 1) {
    for (let column = 1; column <= lastColumn; column += 1) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row - 1, c: column - 1 })];
      const value = cell ? (cell.w != null && String(cell.w).trim() !== "" ? cell.w : cell.v) : null;
      const text = value == null ? "" : String(value).trim();
      if (text.toUpperCase().includes(expected)) {
        return row;
      }
    }
  }
  return null;
}

function sheetText(sheet, row, column) {
  const cell = sheet[XLSX.utils.encode_cell({ r: row - 1, c: column - 1 })];
  if (!cell) {
    return "";
  }
  const value = cell.w != null && String(cell.w).trim() !== "" ? cell.w : cell.v;
  return value == null ? "" : String(value).trim();
}

function createOversizedSource({ inputPath, outputPath, extraRows, sheetName, lastColumn, label }) {
  const script = `
from copy import copy
from pathlib import Path
import warnings
from openpyxl import load_workbook

input_path = Path(${JSON.stringify(inputPath)})
output_path = Path(${JSON.stringify(outputPath)})
sheet_name = ${JSON.stringify(sheetName)}
last_column = ${lastColumn}
extra_rows = ${extraRows}
label = ${JSON.stringify(label)}

with warnings.catch_warnings():
    warnings.filterwarnings("ignore", message="Workbook contains no default style.*", category=UserWarning)
    workbook = load_workbook(input_path, data_only=False, keep_links=True)

worksheet = workbook[sheet_name]

def find_total_row():
    for row in range(1, max(worksheet.max_row, 200) + 1):
        for column in range(1, last_column + 1):
            value = worksheet.cell(row=row, column=column).value
            if value is not None and "TOTAL GENERAL" in str(value).upper():
                return row
    raise RuntimeError(f"TOTAL GENERAL no encontrado en la fuente {label}.")

total_row = find_total_row()
detail_start_row = 11
sample_rows = list(range(detail_start_row, min(total_row, detail_start_row + 10)))
if not sample_rows:
    sample_rows = [detail_start_row]

worksheet.insert_rows(total_row, amount=extra_rows)
for offset in range(extra_rows):
    source_row = sample_rows[offset % len(sample_rows)]
    target_row = total_row + offset
    for column in range(1, last_column + 1):
        source_cell = worksheet.cell(row=source_row, column=column)
        target_cell = worksheet.cell(row=target_row, column=column)
        target_cell.value = source_cell.value
        if source_cell.has_style:
            target_cell._style = copy(source_cell._style)
    worksheet.row_dimensions[target_row].height = worksheet.row_dimensions[source_row].height

output_path.parent.mkdir(parents=True, exist_ok=True)
workbook.save(output_path)
`;

  const result = spawnSync(PYTHON, ["-c", script], {
    cwd: ROOT,
    encoding: "utf8",
  });

  if (result.status !== 0) {
    throw new Error(
      `No se pudo crear la fuente sobredimensionada ${label}: ${result.stderr || result.stdout || result.error?.message}`,
    );
  }
}

function runProcessor(manifestPath) {
  const result = spawnSync(PYTHON, [path.join(ROOT, "python_services", "cli.py"), manifestPath], {
    cwd: ROOT,
    encoding: "utf8",
  });

  if (result.status !== 0) {
    throw new Error(`repuestos_rep_stage fallo: ${result.stderr || result.stdout || result.error?.message}`);
  }

  let payload = null;
  try {
    payload = JSON.parse((result.stdout || "").trim());
  } catch (error) {
    throw new Error(`Respuesta JSON invalida de rep_stage: ${result.stdout}`);
  }

  assertCondition(payload && payload.success === true, "rep_stage no reporto exito.");
  return payload;
}

function assertNoResidualData(sheet, startRow, lastColumn, label) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  for (let row = startRow; row <= range.e.r + 1; row += 1) {
    let hasResidualData = false;
    for (let column = 1; column <= lastColumn; column += 1) {
      if (sheetText(sheet, row, column) !== "") {
        hasResidualData = true;
        break;
      }
    }
    assertCondition(!hasResidualData, `${label} conserva residuo debajo de MAYOR en la fila ${row}.`);
  }
}

function main() {
  const tempDir = buildTempDir();
  const outputPath = buildOutputPath(tempDir);
  const manifestPath = buildManifestPath(tempDir);
  const fixturesRoot = fixturePath("resources/cxp/repuestos_tytserv/fixtures");

  const oversizedSources = {};
  for (const config of REP_CONFIGS) {
    const inputPath = path.join(fixturesRoot, config.sourceFile);
    const outputSourcePath = path.join(tempDir, `${config.key}_oversized.xlsx`);
    createOversizedSource({
      inputPath,
      outputPath: outputSourcePath,
      extraRows: config.extraRows,
      sheetName: "RepLibroVentasGeneral",
      lastColumn: REP_LAST_COLUMN,
      label: config.label,
    });
    oversizedSources[config.field] = { path: outputSourcePath };
  }

  fs.writeFileSync(
    manifestPath,
    JSON.stringify(
      {
        processor: "repuestos_tytserv.rep_stage",
        input_paths: Object.values(oversizedSources).map((entry) => entry.path),
        output_path: outputPath,
        template_path: fixturePath("resources/cxp/repuestos_tytserv/templates/FACTURACION REPUESTOS TYTSERV FEBRERO 2026.xlsx"),
        options: { saved_inputs: oversizedSources },
      },
      null,
      2,
    ),
  );

  const payload = runProcessor(manifestPath);
  const workbook = XLSX.readFile(payload.output_path, { cellFormula: false, cellText: true, sheetStubs: true });

  for (const config of REP_CONFIGS) {
    const sourceWorkbook = XLSX.readFile(oversizedSources[config.field].path, { cellFormula: false, cellText: true, sheetStubs: true });
    const sourceSheet = sourceWorkbook.Sheets.RepLibroVentasGeneral;
    const outputSheet = workbook.Sheets[config.targetSheet];

    assertCondition(!!outputSheet, `No existe ${config.targetSheet} en la salida.`);
    const expectedTotalRow = findRowContaining(sourceSheet, "TOTAL GENERAL", REP_LAST_COLUMN);
    const totalRow = findRowContaining(outputSheet, "TOTAL GENERAL", REP_LAST_COLUMN);
    const mayorRow = findRowContaining(outputSheet, "MAYOR", REP_LAST_COLUMN);

    assertCondition(
      totalRow === expectedTotalRow,
      `${config.targetSheet} no movio TOTAL GENERAL a la fila ${expectedTotalRow}. Valor actual: ${totalRow}`,
    );
    assertCondition(
      mayorRow === expectedTotalRow + 1,
      `${config.targetSheet} no movio MAYOR a la fila ${expectedTotalRow + 1}. Valor actual: ${mayorRow}`,
    );
    assertNoResidualData(outputSheet, mayorRow + 1, REP_LAST_COLUMN, config.targetSheet);
  }

  console.log("OK: REP soporta fuentes mayores que la plantilla para todas las marcas.");
}

main();
