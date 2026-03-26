const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");

const XLSX = require("xlsx");

const ROOT = path.resolve(__dirname, "..", "..");
const PYTHON = process.env.PYTHON_BINARY || "python";
const STRUCTURE_SHEETS = [
  "REP TYT",
  "REP PEUGT",
  "REP CHGN",
  "REP SZK",
  "NC REP TYT",
  "NC REP PEUG",
  "NC REP SZK",
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

function mergeSignature(sheet) {
  return ((sheet && sheet["!merges"]) || [])
    .map((range) => XLSX.utils.encode_range(range))
    .sort();
}

function main() {
  const legacyOutput = buildOutputPath("repuestos_legacy_layout");
  const pythonOutput = buildOutputPath("repuestos_python_layout");
  const legacyManifestPath = buildManifestPath("repuestos_merge_layout_legacy_manifest.json");
  const pythonManifestPath = buildManifestPath("repuestos_merge_layout_python_manifest.json");

  runProcessor(buildManifest("repuestos_tytserv.process_legacy", legacyOutput), legacyManifestPath, "python-legacy-node-wrapper");
  const pythonPayload = runProcessor(buildManifest("repuestos_tytserv.process", pythonOutput), pythonManifestPath, "python-native");

  const legacyWorkbook = XLSX.readFile(legacyOutput, { cellFormula: true, cellText: true, sheetStubs: true });
  const pythonWorkbook = XLSX.readFile(pythonPayload.output_path, { cellFormula: true, cellText: true, sheetStubs: true });

  for (const sheetName of STRUCTURE_SHEETS) {
    const legacySheet = legacyWorkbook.Sheets[sheetName];
    const pythonSheet = pythonWorkbook.Sheets[sheetName];
    assertCondition(!!legacySheet, `No existe hoja ${sheetName} en salida legacy.`);
    assertCondition(!!pythonSheet, `No existe hoja ${sheetName} en salida Python.`);

    const legacyMerges = mergeSignature(legacySheet);
    const pythonMerges = mergeSignature(pythonSheet);
    assertCondition(legacyMerges.length > 0, `${sheetName}: la referencia legacy no tiene merges para validar.`);
    assertCondition(
      JSON.stringify(legacyMerges) === JSON.stringify(pythonMerges),
      `${sheetName}: la estructura de celdas fusionadas no coincide con la salida legacy.`,
    );
  }

  console.log("OK: la estructura de merges REP/NC coincide entre Python nativo y legacy.");
}

main();
