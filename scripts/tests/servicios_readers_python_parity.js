const fs = require("fs");
const os = require("os");
const path = require("path");
const crypto = require("crypto");
const { spawnSync } = require("child_process");

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

function serviciosFixturePath(fileKey) {
  const candidatesByKey = {
    px: [
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "detalle-vtas-xliquidar (2).xlsx"),
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "detalle-vtas-xliquidar.xlsx"),
    ],
    mayor: [
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "archivosasubirtoy", "CON_MAYORGEN2TOY.TXT"),
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "CON_MAYORGEN2TOY.TXT"),
    ],
    repventas: [
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "RepFacturacionServContabilidad (3).xls"),
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "RepFacturacionServContabilidad.xls"),
    ],
  };

  return firstExistingPath(candidatesByKey[fileKey] || []);
}

function tempJsonPath(prefix) {
  return path.join(
    os.tmpdir(),
    `${prefix}_${Date.now()}_${crypto.randomBytes(4).toString("hex")}.json`,
  );
}

function runReader(command, args, outputPath) {
  const result = spawnSync(command, args, {
    cwd: ROOT,
    encoding: "utf8",
  });

  if (result.status !== 0) {
    throw new Error(
      [
        `Fallo lector ${command}.`,
        `Args: ${args.join(" ")}`,
        `STDOUT: ${(result.stdout || "").trim()}`,
        `STDERR: ${(result.stderr || "").trim()}`,
      ].join("\n"),
    );
  }

  assertCondition(fs.existsSync(outputPath), `El lector ${command} no genero ${outputPath}.`);
  return JSON.parse(fs.readFileSync(outputPath, "utf8"));
}

function normalizeScalar(value) {
  if (typeof value === "number") {
    const normalized = Number.isFinite(value) ? Number(value.toFixed(6)) : value;
    return Object.is(normalized, -0) ? 0 : normalized;
  }

  if (typeof value === "string") {
    return value.trim();
  }

  return value == null ? null : value;
}

function normalizeRow(row) {
  const ordered = {};
  for (const key of Object.keys(row).sort()) {
    ordered[key] = normalizeScalar(row[key]);
  }
  return ordered;
}

function normalizePxRows(rows) {
  return rows.map((row) => {
    const cells = Array.isArray(row) ? row.map((value) => String(value || "").trim()) : [];
    while (cells.length > 0 && cells[cells.length - 1] === "") {
      cells.pop();
    }
    return cells;
  });
}

function normalizeMayorRows(rows) {
  return rows.map(normalizeRow);
}

function digest(value) {
  return crypto.createHash("sha1").update(JSON.stringify(value)).digest("hex");
}

function main() {
  const sourceInput = serviciosFixturePath("repventas");
  const pxInput = serviciosFixturePath("px");
  const mayorInput = serviciosFixturePath("mayor");
  assertCondition(fs.existsSync(sourceInput), "No existe el fixture de REP VENTAS para paridad.");
  assertCondition(fs.existsSync(pxInput), "No existe el fixture PX para paridad.");
  assertCondition(fs.existsSync(mayorInput), "No existe el fixture del mayor TXT para paridad.");

  const sourceJsOut = tempJsonPath("servicios_source_js");
  const sourcePyOut = tempJsonPath("servicios_source_py");
  const pxJsOut = tempJsonPath("servicios_px_js");
  const pxPyOut = tempJsonPath("servicios_px_py");
  const mayorJsOut = tempJsonPath("servicios_mayor_js");
  const mayorPyOut = tempJsonPath("servicios_mayor_py");

  try {
    const sourceJs = runReader(
      "node",
      [path.join(ROOT, "scripts", "cxp", "servicios_marcas", "read_source.js"), "--input", sourceInput, "--output-json", sourceJsOut],
      sourceJsOut,
    );
    const sourcePy = runReader(
      PYTHON,
      [path.join(ROOT, "python_services", "processors", "servicios_marcas", "readers.py"), "source", "--input", sourceInput, "--output-json", sourcePyOut],
      sourcePyOut,
    );

    const sourceJsRows = Array.isArray(sourceJs.rows) ? sourceJs.rows.map(normalizeRow) : [];
    const sourcePyRows = Array.isArray(sourcePy.rows) ? sourcePy.rows.map(normalizeRow) : [];

    assertCondition(sourceJsRows.length === sourcePyRows.length, "Servicios source: cantidad de filas distinta entre JS y Python.");
    assertCondition(
      digest(sourceJsRows) === digest(sourcePyRows),
      "Servicios source: la salida Python no coincide con la salida JS.",
    );

    const pxJs = runReader(
      "node",
      [path.join(ROOT, "scripts", "cxp", "servicios_marcas", "read_px.js"), "--input", pxInput, "--brand", "tyt", "--output-json", pxJsOut],
      pxJsOut,
    );
    const pxPy = runReader(
      PYTHON,
      [path.join(ROOT, "python_services", "processors", "servicios_marcas", "readers.py"), "px", "--input", pxInput, "--brand", "tyt", "--output-json", pxPyOut],
      pxPyOut,
    );

    const pxJsRows = normalizePxRows(Array.isArray(pxJs.rows) ? pxJs.rows : []);
    const pxPyRows = normalizePxRows(Array.isArray(pxPy.rows) ? pxPy.rows : []);

    assertCondition(pxJsRows.length === pxPyRows.length, "Servicios PX: cantidad de filas distinta entre JS y Python.");
    assertCondition(
      digest(pxJsRows) === digest(pxPyRows),
      "Servicios PX: la salida Python no coincide con la salida JS.",
    );

    const mayorJs = runReader(
      "node",
      [path.join(ROOT, "scripts", "cxp", "servicios_marcas", "read_mayor_txt.js"), "--input", mayorInput, "--output-json", mayorJsOut],
      mayorJsOut,
    );
    const mayorPy = runReader(
      PYTHON,
      [path.join(ROOT, "python_services", "processors", "servicios_marcas", "readers.py"), "mayor", "--input", mayorInput, "--output-json", mayorPyOut],
      mayorPyOut,
    );

    const mayorJsRows = normalizeMayorRows(Array.isArray(mayorJs.rows) ? mayorJs.rows : []);
    const mayorPyRows = normalizeMayorRows(Array.isArray(mayorPy.rows) ? mayorPy.rows : []);

    assertCondition(mayorJsRows.length === mayorPyRows.length, "Servicios mayor: cantidad de filas distinta entre JS y Python.");
    assertCondition(
      digest(mayorJsRows) === digest(mayorPyRows),
      "Servicios mayor: la salida Python no coincide con la salida JS.",
    );

    console.log("OK: lectores Python de Servicios por Marca mantienen paridad con los lectores JS.");
  } finally {
    for (const filePath of [sourceJsOut, sourcePyOut, pxJsOut, pxPyOut, mayorJsOut, mayorPyOut]) {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    }
  }
}

main();
