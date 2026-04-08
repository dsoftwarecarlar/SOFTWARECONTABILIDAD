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

function tempPath(prefix, extension) {
  return path.join(
    os.tmpdir(),
    `${prefix}_${Date.now()}_${crypto.randomBytes(4).toString("hex")}.${extension}`,
  );
}

function runReader(inputPath, outputPath) {
  return spawnSync(
    PYTHON,
    [
      path.join(ROOT, "python_services", "processors", "servicios_marcas", "readers.py"),
      "source",
      "--input",
      inputPath,
      "--output-json",
      outputPath,
    ],
    {
      cwd: ROOT,
      encoding: "utf8",
    },
  );
}

function createLegacyLayoutFixture(targetPath) {
  const headers = [
    "AGENCIA",
    "CENTRO",
    "No. ORDEN",
    "LINEA NEGOCIO",
    "CEDULA",
    "FACTURADO A",
    "No. FACTURA / NOTA CREDITO",
    "F. FACT.",
    "F. NOTA CRE.",
    "NOTA CREDITO",
    "TOTAL MANO OBRA",
    "TOTAL SUBCONTRATOS",
    "TOTAL INSUMOS",
    "TOTAL SERVICIO",
    "TOTAL ACCESORIOS",
    "TOTAL REPUESTOS",
    "INTERES",
    "VALOR IVA",
    "TOTAL",
    "C. COSTO ENDEREZADA",
    "C. COSTO MANTENIMIENTO",
    "C. COSTO PINTURA",
    "C. COSTO REPARACIONES",
    "C. COSTO ALINEACION",
    "C. COSTO LAVADA",
  ];

  const invoiceRow = [
    "CHANGAN",
    "03",
    "T0001",
    "24",
    "0999999999",
    "CLIENTE PRUEBA",
    "000001",
    "25/03/2026",
    "",
    "1",
    "100",
    "0",
    "0",
    "100",
    "0",
    "20",
    "0",
    "15",
    "135",
    "10",
    "5",
    "2",
    "3",
    "4",
    "1",
  ];

  const noteRow = [
    "CHANGAN",
    "03",
    "T0001A",
    "24",
    "0999999999",
    "CLIENTE PRUEBA",
    "000099",
    "",
    "30/03/2026",
    "-1",
    "-40",
    "0",
    "0",
    "-40",
    "0",
    "-10",
    "0",
    "-6",
    "-56",
    "1",
    "2",
    "3",
    "4",
    "5",
    "6",
  ];

  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([headers, invoiceRow, noteRow]);
  XLSX.utils.book_append_sheet(workbook, sheet, "Legacy");
  XLSX.writeFile(workbook, targetPath);
}

function main() {
  const modernFixture = path.join(
    ROOT,
    "resources",
    "cxp",
    "servicios_marcas",
    "fixtures",
    "RepFacturacionServContabilidad (3).xls",
  );
  assertCondition(fs.existsSync(modernFixture), "No existe fixture moderno de Ventana 2.");

  const modernOut = tempPath("servicios_layout_modern", "json");
  const legacyFixture = tempPath("servicios_layout_legacy", "xlsx");
  const legacyOut = tempPath("servicios_layout_legacy", "json");

  try {
    const modern = runReader(modernFixture, modernOut);
    assertCondition(modern.status === 0, `La fuente moderna no paso el lector Python.\n${modern.stderr || modern.stdout}`);
    assertCondition(fs.existsSync(modernOut), "La fuente moderna no genero salida JSON.");

    createLegacyLayoutFixture(legacyFixture);
    const legacy = runReader(legacyFixture, legacyOut);
    assertCondition(legacy.status === 0, `El layout legacy debe procesarse.\n${legacy.stderr || legacy.stdout}`);
    assertCondition(fs.existsSync(legacyOut), "El layout legacy no genero salida JSON.");
    const payload = JSON.parse(fs.readFileSync(legacyOut, "utf8"));
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    assertCondition(rows.length === 2, `El layout legacy debe generar 2 filas y devolvio ${rows.length}.`);
    assertCondition(rows[0].DocType === "FA", `La fila legacy de factura debe quedar FA y llego ${rows[0].DocType}.`);
    assertCondition(rows[1].DocType === "DC", `La fila legacy de nota debe quedar DC y llego ${rows[1].DocType}.`);
    assertCondition(
      rows[1].AffectedDocumentTrim === rows[0].DocumentTrim,
      `La nota legacy debe enlazarse con la factura base. Esperado ${rows[0].DocumentTrim} y llego ${rows[1].AffectedDocumentTrim}.`,
    );
    assertCondition(rows[0].Costo === 25, `La factura legacy debe preservar el costo total 25 y llego ${rows[0].Costo}.`);
    assertCondition(rows[1].Costo === 21, `La nota legacy debe preservar el costo total 21 y llego ${rows[1].Costo}.`);

    console.log("OK: guardrail de layouts de Ventana 2 validado.");
  } finally {
    for (const filePath of [modernOut, legacyFixture, legacyOut]) {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    }
  }
}

main();
