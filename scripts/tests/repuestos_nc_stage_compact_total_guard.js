const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");

const XLSX = require("xlsx");

const ROOT = path.resolve(__dirname, "..", "..");
const PYTHON = process.env.PYTHON_BINARY || "python";

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

function buildManifestPath() {
  const dir = ensureDir("storage/verify_runs");
  return path.join(dir, "repuestos_nc_compact_total_manifest.json");
}

function buildOutputPath() {
  const dir = ensureDir("storage/verify_runs");
  return path.join(dir, `repuestos_nc_compact_total_${Date.now()}.xlsx`);
}

function buildCompactSzkSource() {
  const dir = ensureDir("storage/verify_runs");
  const outputPath = path.join(dir, `repuestos_nc_szk_compact_${Date.now()}.xlsx`);
  const sourcePath = fixturePath("resources/cxp/repuestos_tytserv/fixtures/archivosnc_rep/RepLibroDevolucionesGeneral (3).xlsx");
  const script = `
from copy import copy
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

source = Path(r"""${sourcePath}""")
target = Path(r"""${outputPath}""")
workbook = load_workbook(source, data_only=False, keep_links=True)
worksheet = workbook["RepLibroDevolucionesGeneral"]
row9_merges = [str(rng) for rng in list(worksheet.merged_cells.ranges) if rng.min_row <= 9 <= rng.max_row]
row10_merges = [str(rng) for rng in list(worksheet.merged_cells.ranges) if rng.min_row <= 10 <= rng.max_row]
for range_ref in row9_merges:
    worksheet.unmerge_cells(range_ref)
for column in range(1, 43):
    source_cell = worksheet.cell(row=10, column=column)
    target_cell = worksheet.cell(row=9, column=column)
    if isinstance(source_cell, MergedCell) or isinstance(target_cell, MergedCell):
        continue
    target_cell.value = copy(source_cell.value)
    source_cell.value = None
for range_ref in row10_merges:
    worksheet.unmerge_cells(range_ref)
workbook.save(target)
print(target)
`;
  const result = spawnSync(PYTHON, ["-c", script], {
    cwd: ROOT,
    encoding: "utf8",
  });

  if (result.status !== 0) {
    throw new Error(result.stderr || result.stdout || "No se pudo crear la fuente compacta NC SZK.");
  }

  return outputPath;
}

function buildManifest(outputPath, compactSzkPath) {
  const ncRoot = fixturePath("resources/cxp/repuestos_tytserv/fixtures/archivosnc_rep");
  return {
    processor: "repuestos_tytserv.nc_stage",
    input_paths: [],
    output_path: outputPath,
    template_path: fixturePath("resources/cxp/repuestos_tytserv/templates/FACTURACION REPUESTOS TYTSERV FEBRERO 2026.xlsx"),
    options: {
      saved_inputs: {
        excel_nc_tyt: { path: path.join(ncRoot, "RepLibroDevolucionesGeneral.xlsx") },
        excel_nc_peug: { path: path.join(ncRoot, "RepLibroDevolucionesGeneral (1).xlsx") },
        excel_nc_szk: { path: compactSzkPath },
      },
    },
  };
}

function runProcessor(manifestPath) {
  const result = spawnSync(PYTHON, [path.join(ROOT, "python_services", "cli.py"), manifestPath], {
    cwd: ROOT,
    encoding: "utf8",
  });

  if (result.status !== 0) {
    throw new Error(result.stderr || result.stdout || "nc_stage fallo");
  }

  const payload = JSON.parse((result.stdout || "").trim());
  assertCondition(payload && payload.success === true, "nc_stage no reporto exito.");
  return payload;
}

function sheetText(sheet, row, column) {
  const cell = sheet[XLSX.utils.encode_cell({ r: row - 1, c: column - 1 })];
  if (!cell) {
    return "";
  }

  const value = cell.w != null && String(cell.w).trim() !== "" ? cell.w : cell.v;
  return value == null ? "" : String(value).trim();
}

function findRowContaining(sheet, needle, lastColumn = 42) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  const expected = String(needle || "").trim().toUpperCase();
  for (let row = 1; row <= range.e.r + 1; row += 1) {
    for (let column = 1; column <= lastColumn; column += 1) {
      if (sheetText(sheet, row, column).toUpperCase().includes(expected)) {
        return row;
      }
    }
  }

  return null;
}

function main() {
  const compactSzkPath = buildCompactSzkSource();
  const compactWorkbook = XLSX.readFile(compactSzkPath, { cellFormula: true, cellText: true, sheetStubs: true });
  const compactSheet = compactWorkbook.Sheets.RepLibroDevolucionesGeneral;
  const compactTotalRow = findRowContaining(compactSheet, "TOTAL GENERAL");
  assertCondition(compactTotalRow === 9, `La fuente compacta NC SZK no quedo en fila 9; llego ${compactTotalRow}.`);

  const outputPath = buildOutputPath();
  const manifestPath = buildManifestPath();
  fs.writeFileSync(manifestPath, JSON.stringify(buildManifest(outputPath, compactSzkPath), null, 2));

  const payload = runProcessor(manifestPath);
  const workbook = XLSX.readFile(payload.output_path, { cellFormula: true, cellText: true, sheetStubs: true });
  const sheet = workbook.Sheets["NC REP SZK"];
  assertCondition(!!sheet, "No existe hoja NC REP SZK en la salida.");

  const totalRow = findRowContaining(sheet, "TOTAL GENERAL");
  const mayorRow = findRowContaining(sheet, "MAYOR");
  assertCondition(totalRow === 9, `NC REP SZK: TOTAL GENERAL debe quedar en fila 9 para fuente compacta y llego ${totalRow}.`);
  assertCondition(mayorRow === 10, `NC REP SZK: MAYOR debe quedar en fila 10 para fuente compacta y llego ${mayorRow}.`);
  assertCondition(sheetText(sheet, 9, 13).toUpperCase().includes("TOTAL GENERAL"), "NC REP SZK: se perdio la etiqueta TOTAL GENERAL.");

  console.log("OK: NC REP SZK compacto conserva detalle y fila TOTAL GENERAL.");
}

main();
