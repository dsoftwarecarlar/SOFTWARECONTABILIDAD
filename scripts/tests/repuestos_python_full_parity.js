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

function buildManifest(processor, outputPath) {
  return {
    processor,
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
  if (!cell) {
    return "";
  }

  const value = cell.w != null && String(cell.w).trim() !== "" ? cell.w : cell.v;
  return value == null ? "" : String(value).trim();
}

function normalizedVisibleSheetSignature(sheet, ignoredHeaderCells = []) {
  const crypto = require("crypto");
  const ref = sheet["!ref"] || "A1";
  const range = XLSX.utils.decode_range(ref);
  const rows = [];

  for (let row = range.s.r + 1; row <= range.e.r + 1; row += 1) {
    const values = [];
    let nonEmpty = false;
    for (let column = range.s.c + 1; column <= range.e.c + 1; column += 1) {
      let value = cellText(sheet, row, column);
      if (row === 2 && ignoredHeaderCells.includes(column)) {
        value = "";
      }
      if (value !== "") {
        nonEmpty = true;
      }
      values.push(value);
    }

    if (nonEmpty) {
      rows.push(values.join("|"));
    }
  }

  return {
    visibleNonEmptyRows: rows.length,
    hash: crypto.createHash("sha256").update(rows.join("\n")).digest("hex"),
  };
}

function loadContract() {
  return JSON.parse(
    fs.readFileSync(
      fixturePath("resources/cxp/repuestos_tytserv/contracts/repuestos_tytserv_fixture_contract.json"),
      "utf8",
    ),
  );
}

function main() {
  const contract = loadContract();
  const legacyOutput = buildOutputPath("repuestos_legacy_full");
  const pythonOutput = buildOutputPath("repuestos_python_full");
  const legacyManifestPath = buildManifestPath("repuestos_python_full_legacy_manifest.json");
  const pythonManifestPath = buildManifestPath("repuestos_python_full_manifest.json");

  runProcessor(buildManifest("repuestos_tytserv.process_legacy", legacyOutput), legacyManifestPath, "python-legacy-node-wrapper");
  const pythonPayload = runProcessor(buildManifest("repuestos_tytserv.process", pythonOutput), pythonManifestPath, "python-native");

  const legacyWorkbook = XLSX.readFile(legacyOutput, { cellFormula: false, cellText: true, sheetStubs: true });
  const pythonWorkbook = XLSX.readFile(pythonPayload.output_path, { cellFormula: false, cellText: true, sheetStubs: true });

  assertCondition(
    JSON.stringify(pythonWorkbook.SheetNames) === JSON.stringify(contract.sheet_order),
    "El workbook Python nativo cambio el orden de hojas del contrato.",
  );
  assertCondition(
    JSON.stringify(legacyWorkbook.SheetNames) === JSON.stringify(pythonWorkbook.SheetNames),
    "El workbook Python nativo cambio el orden de hojas respecto al legado.",
  );

  for (const [sheetName, expectedSignature] of Object.entries(contract.sheets || {})) {
    const legacySheet = legacyWorkbook.Sheets[sheetName];
    const pythonSheet = pythonWorkbook.Sheets[sheetName];
    assertCondition(!!legacySheet, `No existe hoja ${sheetName} en salida legacy.`);
    assertCondition(!!pythonSheet, `No existe hoja ${sheetName} en salida Python.`);

    const ignoredHeaderCells = Array.isArray(contract.visible_signature_rules?.ignored_header_cells?.[sheetName])
      ? contract.visible_signature_rules.ignored_header_cells[sheetName]
      : [];
    const legacySignature = normalizedVisibleSheetSignature(legacySheet, ignoredHeaderCells);
    const pythonSignature = normalizedVisibleSheetSignature(pythonSheet, ignoredHeaderCells);

    assertCondition(
      pythonSignature.visibleNonEmptyRows === expectedSignature.visible_non_empty_rows,
      `${sheetName}: filas visibles inesperadas en Python (${pythonSignature.visibleNonEmptyRows} != ${expectedSignature.visible_non_empty_rows}).`,
    );
    assertCondition(
      pythonSignature.hash === expectedSignature.visible_hash,
      `${sheetName}: hash visible Python no coincide con el contrato actual.`,
    );
    assertCondition(
      legacySignature.visibleNonEmptyRows === pythonSignature.visibleNonEmptyRows
      && legacySignature.hash === pythonSignature.hash,
      `${sheetName}: Python nativo no coincide con la salida legacy.`,
    );
  }

  console.log("OK: Repuestos TYTSERV completo validado con Python nativo contra contrato y salida legacy.");
}

main();
