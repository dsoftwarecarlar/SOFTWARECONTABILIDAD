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

function loadContract() {
  return JSON.parse(
    fs.readFileSync(
      fixturePath("resources/cxp/repuestos_tytserv/contracts/repuestos_tytserv_nc_stage_contract.json"),
      "utf8",
    ),
  );
}

function buildManifestPath() {
  const dir = fixturePath("storage/verify_runs");
  fs.mkdirSync(dir, { recursive: true });
  return path.join(dir, "repuestos_nc_stage_manifest.json");
}

function buildOutputPath() {
  const dir = fixturePath("storage/verify_runs");
  fs.mkdirSync(dir, { recursive: true });
  return path.join(dir, `repuestos_nc_stage_${Date.now()}.xlsx`);
}

function buildSourceEntries() {
  const fixturesRoot = fixturePath("resources/cxp/repuestos_tytserv/fixtures/archivosnc_rep");
  return [
    {
      field: "excel_nc_tyt",
      sourcePath: path.join(fixturesRoot, "RepLibroDevolucionesGeneral.xlsx"),
      targetSheet: "NC REP TYT",
    },
    {
      field: "excel_nc_peug",
      sourcePath: path.join(fixturesRoot, "RepLibroDevolucionesGeneral (1).xlsx"),
      targetSheet: "NC REP PEUG",
    },
    {
      field: "excel_nc_szk",
      sourcePath: path.join(fixturesRoot, "RepLibroDevolucionesGeneral (3).xlsx"),
      targetSheet: "NC REP SZK",
    },
  ];
}

function buildManifest(outputPath, sources) {
  const templatePath = fixturePath("resources/cxp/repuestos_tytserv/templates/FACTURACION REPUESTOS TYTSERV FEBRERO 2026.xlsx");
  const savedInputs = {};
  for (const source of sources) {
    savedInputs[source.field] = {
      path: source.sourcePath,
      original_name: path.basename(source.sourcePath),
    };
  }

  return {
    processor: "repuestos_tytserv.nc_stage",
    input_paths: sources.map((item) => item.sourcePath),
    output_path: outputPath,
    template_path: templatePath,
    options: {
      saved_inputs: savedInputs,
    },
  };
}

function runProcessor(manifestPath) {
  const result = spawnSync(PYTHON, [path.join(ROOT, "python_services", "cli.py"), manifestPath], {
    cwd: ROOT,
    encoding: "utf8",
  });

  if (result.status !== 0) {
    throw new Error(`repuestos_nc_stage fallo: ${result.stderr || result.stdout}`);
  }

  let payload = null;
  try {
    payload = JSON.parse((result.stdout || "").trim());
  } catch (error) {
    throw new Error(`Respuesta JSON invalida de nc_stage: ${result.stdout}`);
  }

  assertCondition(payload && payload.success === true, "nc_stage no reporto exito.");
  return payload;
}

function cellText(sheet, row, column) {
  const cell = sheet[XLSX.utils.encode_cell({ r: row - 1, c: column - 1 })];
  const value = cell ? (cell.w != null && String(cell.w).trim() !== "" ? cell.w : cell.v) : null;
  return value == null ? "" : String(value).trim();
}

function detailSignature(sheet, totalRow) {
  const rows = [];
  for (let row = 8; row <= totalRow; row += 1) {
    const values = [];
    let nonEmpty = false;
    for (let column = 1; column <= 42; column += 1) {
      const value = cellText(sheet, row, column);
      if (value !== "") {
        nonEmpty = true;
      }
      values.push(value);
    }

    if (nonEmpty) {
      rows.push(values.join("|"));
    }
  }

  const crypto = require("crypto");
  return {
    rows: rows.length,
    hash: crypto.createHash("sha256").update(rows.join("\n")).digest("hex"),
  };
}

function main() {
  const contract = loadContract();
  const sources = buildSourceEntries();
  const outputPath = buildOutputPath();
  const manifestPath = buildManifestPath();
  fs.writeFileSync(manifestPath, JSON.stringify(buildManifest(outputPath, sources), null, 2));

  const payload = runProcessor(manifestPath);
  const workbook = XLSX.readFile(payload.output_path, { cellFormula: false, cellText: true, sheetStubs: true });
  assertCondition(payload.metadata && payload.metadata.runtime === "python-native-nc-stage", "nc_stage no reporto runtime esperado.");

  for (const source of sources) {
    const sourceWorkbook = XLSX.readFile(source.sourcePath, { cellFormula: false, cellText: true, sheetStubs: true });
    const sourceSheet = sourceWorkbook.Sheets.RepLibroDevolucionesGeneral;
    const outputSheet = workbook.Sheets[source.targetSheet];
    assertCondition(!!outputSheet, `No existe hoja ${source.targetSheet} en salida nc_stage.`);

    const sheetContract = contract.sheets[source.targetSheet];
    assertCondition(!!sheetContract, `No existe contrato NC para ${source.targetSheet}.`);

    const sourceSignature = detailSignature(sourceSheet, sheetContract.total_row);
    const outputSignature = detailSignature(outputSheet, sheetContract.total_row);
    assertCondition(
      sourceSignature.rows === outputSignature.rows && sourceSignature.hash === outputSignature.hash,
      `${source.targetSheet}: el detalle NC no coincide con la fuente subida.`,
    );

    for (const [column, expectedValue] of Object.entries(sheetContract.mayor_values || {})) {
      const actualValue = cellText(outputSheet, sheetContract.mayor_row, Number(column));
      assertCondition(
        actualValue === String(expectedValue),
        `${source.targetSheet}: valor MAYOR inesperado en columna ${column} (${actualValue} != ${expectedValue}).`,
      );
    }
  }

  console.log("OK: etapa NC de Repuestos validada contra fuente y contrato visible.");
}

main();
