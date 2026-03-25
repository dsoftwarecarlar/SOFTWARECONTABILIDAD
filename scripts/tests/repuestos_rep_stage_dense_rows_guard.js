const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");

const ROOT = path.resolve(__dirname, "..", "..");
const PYTHON = process.env.PYTHON_BINARY || "python";
const DETAIL_START_ROW = 11;

const CONFIGS = [
  { field: "excel_tyt", targetSheet: "REP TYT", sourceFile: "RepLibroVentasGeneral.xlsx", label: "TOYOTA", extraRows: 2 },
  { field: "excel_peug", targetSheet: "REP PEUGT", sourceFile: "RepLibroVentasGeneral (1).xlsx", label: "PEUGEOT", extraRows: 2 },
  { field: "excel_chgn", targetSheet: "REP CHGN", sourceFile: "RepLibroVentasGeneral (2).xlsx", label: "CHANGAN", extraRows: 2 },
  { field: "excel_szk", targetSheet: "REP SZK", sourceFile: "RepLibroVentasGeneral (3).xlsx", label: "SUZUKI", extraRows: 2 },
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
    `repuestos-rep-dense-${Date.now()}-${Math.random().toString(16).slice(2)}`,
  );
  fs.mkdirSync(directory, { recursive: true });
  return directory;
}

function createDenseSource({ inputPath, outputPath, extraRows, label }) {
  const script = `
from copy import copy
from pathlib import Path
import warnings
from openpyxl import load_workbook

input_path = Path(${JSON.stringify(inputPath)})
output_path = Path(${JSON.stringify(outputPath)})
extra_rows = ${extraRows}
label = ${JSON.stringify(label)}

with warnings.catch_warnings():
    warnings.filterwarnings("ignore", message="Workbook contains no default style.*", category=UserWarning)
    workbook = load_workbook(input_path, data_only=False, keep_links=True)

worksheet = workbook["RepLibroVentasGeneral"]

def find_total_row():
    for row in range(1, max(worksheet.max_row, 200) + 1):
        for column in range(1, 42):
            value = worksheet.cell(row=row, column=column).value
            if value is not None and "TOTAL GENERAL" in str(value).upper():
                return row
    raise RuntimeError(f"TOTAL GENERAL no encontrado en la fuente {label}.")

total_row = find_total_row()
insert_at = max(${DETAIL_START_ROW + 1}, total_row - 1)
sample_rows = list(range(max(${DETAIL_START_ROW}, insert_at - 3), insert_at))
if not sample_rows:
    sample_rows = [${DETAIL_START_ROW}]

worksheet.insert_rows(insert_at, amount=extra_rows)
for offset in range(extra_rows):
    source_row = sample_rows[offset % len(sample_rows)]
    target_row = insert_at + offset
    for column in range(1, 42):
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
    throw new Error(`No se pudo crear la fuente densa ${label}: ${result.stderr || result.stdout || result.error?.message}`);
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

  const payload = JSON.parse((result.stdout || "").trim());
  assertCondition(payload && payload.success === true, "rep_stage no reporto exito.");
  return payload;
}

function main() {
  const tempDir = buildTempDir();
  const outputPath = path.join(tempDir, "rep_dense_guard.xlsx");
  const manifestPath = path.join(tempDir, "manifest.json");
  const fixturesRoot = fixturePath("resources/cxp/repuestos_tytserv/fixtures");

  const savedInputs = {};
  for (const config of CONFIGS) {
    const inputPath = path.join(fixturesRoot, config.sourceFile);
    const densePath = path.join(tempDir, `${config.field}_dense.xlsx`);
    createDenseSource({ inputPath, outputPath: densePath, extraRows: config.extraRows, label: config.label });
    savedInputs[config.field] = { path: densePath };
  }

  fs.writeFileSync(
    manifestPath,
    JSON.stringify(
      {
        processor: "repuestos_tytserv.rep_stage",
        input_paths: Object.values(savedInputs).map((entry) => entry.path),
        output_path: outputPath,
        template_path: fixturePath("resources/cxp/repuestos_tytserv/templates/FACTURACION REPUESTOS TYTSERV FEBRERO 2026.xlsx"),
        options: { saved_inputs: savedInputs },
      },
      null,
      2,
    ),
  );

  const payload = runProcessor(manifestPath);

  for (const config of CONFIGS) {
    const stageMetadata = payload.metadata && payload.metadata.rep_stage ? payload.metadata.rep_stage[config.field.replace("excel_", "")] : null;
    assertCondition(!!stageMetadata, `${config.targetSheet}: no existe metadata REP para la marca.`);
    assertCondition(
      Number(stageMetadata.actual_mayor_row) === Number(stageMetadata.source_total_row) + 1,
      `${config.targetSheet}: MAYOR no quedo en TOTAL+1 para la fuente densa.`,
    );
  }

  console.log("OK: REP procesa fuentes densas sin romper el hash interno ni la posicion de MAYOR.");
}

main();
