const fs = require("fs");
const os = require("os");
const path = require("path");
const crypto = require("crypto");
const { spawnSync } = require("child_process");

const XLSX = require("xlsx");

const ROOT = path.resolve(__dirname, "..", "..");
const PYTHON = process.env.PYTHON_BINARY || "python";

function assertCondition(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

function firstExistingPath(candidates) {
  for (const candidate of candidates) {
    if (candidate && fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return candidates[0];
}

function fixturePath(relativePath) {
  return firstExistingPath([path.join(ROOT, relativePath)]);
}

function loadContract() {
  const contractPath = fixturePath("resources/cxp/repuestos_tytserv/contracts/repuestos_tytserv_fixture_contract.json");
  return JSON.parse(fs.readFileSync(contractPath, "utf8"));
}

function buildManifestPath() {
  const dir = path.join(ROOT, "storage", "verify_runs");
  fs.mkdirSync(dir, { recursive: true });
  return path.join(dir, "repuestos_rep_stage_manifest.json");
}

function buildOutputPath() {
  const dir = path.join(ROOT, "storage", "verify_runs");
  fs.mkdirSync(dir, { recursive: true });
  return path.join(dir, `repuestos_rep_stage_${Date.now()}.xlsx`);
}

function buildSourceEntries() {
  const fixturesRoot = fixturePath("resources/cxp/repuestos_tytserv/fixtures");
  return [
    {
      field: "excel_tyt",
      sourcePath: path.join(fixturesRoot, "RepLibroVentasGeneral.xlsx"),
      targetSheet: "REP TYT",
    },
    {
      field: "excel_peug",
      sourcePath: path.join(fixturesRoot, "RepLibroVentasGeneral (1).xlsx"),
      targetSheet: "REP PEUGT",
    },
    {
      field: "excel_chgn",
      sourcePath: path.join(fixturesRoot, "RepLibroVentasGeneral (2).xlsx"),
      targetSheet: "REP CHGN",
    },
    {
      field: "excel_szk",
      sourcePath: path.join(fixturesRoot, "RepLibroVentasGeneral (3).xlsx"),
      targetSheet: "REP SZK",
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
    processor: "repuestos_tytserv.rep_stage",
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
    throw new Error(`repuestos_rep_stage fallo: ${result.stderr || result.stdout}`);
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

function sheetText(sheet, row, column) {
  const cell = sheet[XLSX.utils.encode_cell({ r: row - 1, c: column - 1 })];
  if (!cell) {
    return "";
  }

  const value = cell.w != null && String(cell.w).trim() !== "" ? cell.w : cell.v;
  return value == null ? "" : String(value).trim();
}

function findRowContaining(sheet, needle, lastColumn = 41) {
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

function payloadSignature(sheet) {
  const payloadColumns = Array.from({ length: 41 }, (_value, index) => index + 1);
  const totalRow = findRowContaining(sheet, "TOTAL GENERAL");
  assertCondition(!!totalRow, "No se encontro TOTAL GENERAL en la hoja.");

  const rows = [];
  for (let row = 11; row < totalRow; row += 1) {
    const document = sheetText(sheet, row, 5);
    if (document === "" || document.toUpperCase().startsWith("ANULAD")) {
      continue;
    }

    rows.push(payloadColumns.map((column) => sheetText(sheet, row, column)).join("|"));
  }

  return {
    rows: rows.length,
    hash: crypto.createHash("sha256").update(rows.join("\n")).digest("hex"),
    totalRow,
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

  assertCondition(
    JSON.stringify(workbook.SheetNames) === JSON.stringify(contract.sheet_order),
    "rep_stage cambio el orden de hojas respecto al contrato actual.",
  );
  assertCondition(
    payload.metadata && payload.metadata.runtime === "python-native-rep-stage",
    "rep_stage no reporto runtime esperado.",
  );

  for (const source of sources) {
    const sourceWorkbook = XLSX.readFile(source.sourcePath, { cellFormula: false, cellText: true, sheetStubs: true });
    const sourceSheet = sourceWorkbook.Sheets.RepLibroVentasGeneral;
    const outputSheet = workbook.Sheets[source.targetSheet];
    assertCondition(!!outputSheet, `No existe hoja ${source.targetSheet} en salida rep_stage.`);

    const sourceSignature = payloadSignature(sourceSheet);
    const outputSignature = payloadSignature(outputSheet);
    assertCondition(
      sourceSignature.rows === outputSignature.rows && sourceSignature.hash === outputSignature.hash,
      `${source.targetSheet}: el payload REP no coincide con la fuente subida.`,
    );

    const mayorRow = findRowContaining(outputSheet, "MAYOR");
    assertCondition(mayorRow === outputSignature.totalRow + 1, `${source.targetSheet}: fila MAYOR inesperada.`);

  }

  console.log("OK: etapa REP de Repuestos validada contra fuente y contrato estructural.");
}

main();
