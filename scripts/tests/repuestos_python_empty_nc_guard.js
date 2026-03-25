const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");

const XLSX = require("xlsx");

const ROOT = path.resolve(__dirname, "..", "..");
const PYTHON = process.env.PYTHON_BINARY || "python";
const NC_LAST_COLUMN = 42;

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
    `repuestos-empty-nc-${Date.now()}-${Math.random().toString(16).slice(2)}`,
  );
  fs.mkdirSync(directory, { recursive: true });
  return directory;
}

function buildManifestPath(tempDir) {
  return path.join(tempDir, "manifest.json");
}

function buildOutputPath(tempDir) {
  return path.join(tempDir, "repuestos_empty_nc_guard.xlsx");
}

function runProcessor(manifestPath) {
  const result = spawnSync(PYTHON, [path.join(ROOT, "python_services", "cli.py"), manifestPath], {
    cwd: ROOT,
    encoding: "utf8",
  });

  if (result.status !== 0) {
    throw new Error(`repuestos_tytserv.process fallo: ${result.stderr || result.stdout || result.error?.message}`);
  }

  let payload = null;
  try {
    payload = JSON.parse((result.stdout || "").trim());
  } catch (error) {
    throw new Error(`Respuesta JSON invalida del proceso completo: ${result.stdout}`);
  }

  assertCondition(payload && payload.success === true, "El proceso completo no reporto exito.");
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

function cellNumber(sheet, row, column) {
  const text = cellText(sheet, row, column).replace(/,/g, "");
  if (text === "") {
    return 0;
  }
  const value = Number(text);
  return Number.isFinite(value) ? value : 0;
}

function findRowContaining(sheet, needle, lastColumn = NC_LAST_COLUMN) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  const expected = String(needle || "").trim().toUpperCase();
  for (let row = 1; row <= range.e.r + 1; row += 1) {
    for (let column = 1; column <= lastColumn; column += 1) {
      if (cellText(sheet, row, column).toUpperCase().includes(expected)) {
        return row;
      }
    }
  }
  return null;
}

function detailRowCount(sheet, startRow, endRow, lastColumn = NC_LAST_COLUMN) {
  let rows = 0;
  for (let row = startRow; row <= endRow; row += 1) {
    let nonEmpty = false;
    for (let column = 1; column <= lastColumn; column += 1) {
      if (cellText(sheet, row, column) !== "") {
        nonEmpty = true;
        break;
      }
    }
    if (nonEmpty) {
      rows += 1;
    }
  }
  return rows;
}

function main() {
  const tempDir = buildTempDir();
  const outputPath = buildOutputPath(tempDir);
  const manifestPath = buildManifestPath(tempDir);
  const repRoot = fixturePath("resources/cxp/repuestos_tytserv/fixtures/otro_mes");
  const ncRoot = fixturePath("resources/cxp/repuestos_tytserv/fixtures/archivosnc_rep");

  const savedInputs = {
    excel_tyt: { path: path.join(repRoot, "RepLibroVentasGeneral (12).xlsx") },
    excel_nc_tyt: { path: path.join(ncRoot, "RepLibroDevolucionesGeneral.xlsx") },
    excel_peug: { path: path.join(repRoot, "RepLibroVentasGeneral (13).xlsx") },
    excel_nc_peug: { path: path.join(ncRoot, "RepLibroDevolucionesGeneral (2).xlsx") },
    excel_chgn: { path: path.join(repRoot, "RepLibroVentasGeneral (14).xlsx") },
    excel_nc_chgn: { path: path.join(ncRoot, "RepLibroDevolucionesGeneral (2).xlsx") },
    excel_szk: { path: path.join(repRoot, "RepLibroVentasGeneral (15).xlsx") },
    excel_nc_szk: { path: path.join(ncRoot, "RepLibroDevolucionesGeneral (3).xlsx") },
  };

  fs.writeFileSync(
    manifestPath,
    JSON.stringify(
      {
        processor: "repuestos_tytserv.process",
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
  const chgnSections = payload.metadata?.stages?.my?.sections?.chgn || {};
  assertCondition(
    chgnSections.sales_total_row === 10,
    `MY REP CHGN no expandio la seccion de ventas como se esperaba: ${JSON.stringify(chgnSections)}`,
  );
  const workbook = XLSX.readFile(payload.output_path, { cellFormula: false, cellText: true, sheetStubs: true });
  const peugSheet = workbook.Sheets["NC REP PEUG"];

  assertCondition(!!peugSheet, "No existe la hoja NC REP PEUG en la salida.");

  const totalRow = findRowContaining(peugSheet, "TOTAL GENERAL");
  const mayorRow = findRowContaining(peugSheet, "MAYOR");
  assertCondition(totalRow === 9, `NC REP PEUG movio TOTAL GENERAL a una fila inesperada: ${totalRow}.`);
  assertCondition(mayorRow === 10, `NC REP PEUG movio MAYOR a una fila inesperada: ${mayorRow}.`);
  assertCondition(
    detailRowCount(peugSheet, 8, totalRow - 1) === 0,
    "NC REP PEUG conserva detalle visible cuando la fuente NC esta vacia.",
  );

  for (const column of [22, 23, 24, 25, 26, 27, 28, 30, 33, 34, 36]) {
    assertCondition(
      cellNumber(peugSheet, totalRow, column) === 0,
      `NC REP PEUG conserva residuo en TOTAL GENERAL columna ${column}.`,
    );
  }

  assertCondition(cellText(peugSheet, mayorRow, 22) === "MAYOR", "NC REP PEUG perdio la etiqueta MAYOR.");
  assertCondition(cellNumber(peugSheet, mayorRow, 23) === 0, "NC REP PEUG conserva base NC en MAYOR columna 23.");
  assertCondition(cellNumber(peugSheet, mayorRow, 24) === 0, "NC REP PEUG conserva descuento NC en MAYOR columna 24.");

  console.log("OK: el proceso completo acepta NC vacio, expande MY CHGN y limpia PEUGEOT sin arrastrar residuos.");
}

main();
