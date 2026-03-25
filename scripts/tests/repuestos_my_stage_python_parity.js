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

function buildOutputPath(prefix) {
  const dir = ensureDir("storage/verify_runs");
  return path.join(dir, `${prefix}_${Date.now()}.xlsx`);
}

function buildManifestPath(name) {
  const dir = ensureDir("storage/verify_runs");
  return path.join(dir, name);
}

function buildSavedInputs() {
  const fixturesRoot = fixturePath("resources/cxp/repuestos_tytserv/fixtures");
  const ncRoot = path.join(fixturesRoot, "archivosnc_rep");
  return {
    excel_tyt: { path: path.join(fixturesRoot, "RepLibroVentasGeneral.xlsx") },
    excel_peug: { path: path.join(fixturesRoot, "RepLibroVentasGeneral (1).xlsx") },
    excel_chgn: { path: path.join(fixturesRoot, "RepLibroVentasGeneral (2).xlsx") },
    excel_szk: { path: path.join(fixturesRoot, "RepLibroVentasGeneral (3).xlsx") },
    excel_nc_tyt: { path: path.join(ncRoot, "RepLibroDevolucionesGeneral.xlsx") },
    excel_nc_peug: { path: path.join(ncRoot, "RepLibroDevolucionesGeneral (1).xlsx") },
    excel_nc_chgn: { path: path.join(ncRoot, "RepLibroDevolucionesGeneral (2).xlsx") },
    excel_nc_szk: { path: path.join(ncRoot, "RepLibroDevolucionesGeneral (3).xlsx") },
  };
}

function buildLegacyManifest(outputPath) {
  return {
    processor: "repuestos_tytserv.process_legacy",
    input_paths: [],
    output_path: outputPath,
    template_path: fixturePath("resources/cxp/repuestos_tytserv/templates/FACTURACION REPUESTOS TYTSERV FEBRERO 2026.xlsx"),
    options: {
      script_path: fixturePath("scripts/cxp/repuestos_tytserv/process.js"),
      cwd: ROOT,
      timeout_seconds: 600,
      file_fields: [
        { field: "excel_tyt", script_flag: "--input-tyt" },
        { field: "excel_peug", script_flag: "--input-peug" },
        { field: "excel_chgn", script_flag: "--input-chgn" },
        { field: "excel_szk", script_flag: "--input-szk" },
        { field: "excel_nc_tyt", script_flag: "--input-nc-tyt" },
        { field: "excel_nc_peug", script_flag: "--input-nc-peug" },
        { field: "excel_nc_chgn", script_flag: "--input-nc-chgn" },
        { field: "excel_nc_szk", script_flag: "--input-nc-szk" },
      ],
      saved_inputs: buildSavedInputs(),
    },
  };
}

function buildPythonManifest(outputPath) {
  return {
    processor: "repuestos_tytserv.my_stage",
    input_paths: [],
    output_path: outputPath,
    template_path: fixturePath("resources/cxp/repuestos_tytserv/templates/FACTURACION REPUESTOS TYTSERV FEBRERO 2026.xlsx"),
    options: {
      saved_inputs: buildSavedInputs(),
    },
  };
}

function runProcessor(manifest, manifestPath, expectedRuntime = null) {
  fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2));
  const result = spawnSync(PYTHON, [path.join(ROOT, "python_services", "cli.py"), manifestPath], {
    cwd: ROOT,
    encoding: "utf8",
  });

  if (result.status !== 0) {
    throw new Error(result.stderr || result.stdout || `Fallo ${manifest.processor}`);
  }

  const payload = JSON.parse((result.stdout || "").trim());
  assertCondition(payload && payload.success === true, `${manifest.processor} no reporto exito.`);
  if (expectedRuntime) {
    assertCondition(
      payload.metadata && payload.metadata.runtime === expectedRuntime,
      `${manifest.processor} no reporto runtime ${expectedRuntime}.`,
    );
  }
  return payload;
}

function cellText(sheet, row, column) {
  const cell = sheet[XLSX.utils.encode_cell({ r: row - 1, c: column - 1 })];
  const value = cell ? (cell.w != null && String(cell.w).trim() !== "" ? cell.w : cell.v) : null;
  return value == null ? "" : String(value).trim();
}

function sheetSignature(sheet, maxColumns = 12) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  const rows = [];
  for (let row = 1; row <= range.e.r + 1; row += 1) {
    const values = [];
    let nonEmpty = false;
    for (let column = 1; column <= maxColumns; column += 1) {
      const value = cellText(sheet, row, column);
      if (value !== "") {
        nonEmpty = true;
      }
      values.push(value);
    }

    if (nonEmpty) {
      rows.push(`${row}:${values.join("|")}`);
    }
  }
  return rows;
}

function compareSheetText(workbookA, workbookB, sheetName) {
  const left = workbookA.Sheets[sheetName];
  const right = workbookB.Sheets[sheetName];
  assertCondition(!!left, `No existe hoja ${sheetName} en workbook legado.`);
  assertCondition(!!right, `No existe hoja ${sheetName} en workbook Python.`);

  const leftSignature = sheetSignature(left);
  const rightSignature = sheetSignature(right);
  assertCondition(
    leftSignature.length === rightSignature.length,
    `${sheetName}: cantidad de filas visibles distinta (${leftSignature.length} != ${rightSignature.length}).`,
  );

  for (let index = 0; index < leftSignature.length; index += 1) {
    assertCondition(
      leftSignature[index] === rightSignature[index],
      `${sheetName}: mismatch visible en ${leftSignature[index]} vs ${rightSignature[index]}.`,
    );
  }
}

function main() {
  const legacyOutput = buildOutputPath("repuestos_legacy_for_my_stage");
  const pythonOutput = buildOutputPath("repuestos_my_stage");
  const legacyManifestPath = buildManifestPath("repuestos_my_stage_legacy_manifest.json");
  const pythonManifestPath = buildManifestPath("repuestos_my_stage_manifest.json");

  runProcessor(buildLegacyManifest(legacyOutput), legacyManifestPath);
  runProcessor(buildPythonManifest(pythonOutput), pythonManifestPath, "python-native-my-stage");

  const legacyWorkbook = XLSX.readFile(legacyOutput, { cellFormula: false, cellText: true, sheetStubs: true });
  const pythonWorkbook = XLSX.readFile(pythonOutput, { cellFormula: false, cellText: true, sheetStubs: true });

  for (const sheetName of ["MY REP TYT", "MY REP PEUG", "MY REP CHGN", "MY REP SZK"]) {
    compareSheetText(legacyWorkbook, pythonWorkbook, sheetName);
  }

  console.log("OK: etapa MY de Repuestos validada contra salida legacy.");
}

main();
