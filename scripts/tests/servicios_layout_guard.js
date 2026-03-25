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
    "",
    "ASESOR",
    "",
    "FACTURA",
    "F. FACT",
    "F. NOTA",
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
  ];

  const row = [
    "MATRIZ",
    "03",
    "T0001",
    "",
    "ASESOR",
    "",
    "000001",
    "25/03/2026",
    "",
    "0",
    "100",
    "0",
    "0",
    "100",
    "0",
    "0",
    "0",
    "15",
    "115",
    "10",
  ];

  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([headers, row]);
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
    assertCondition(legacy.status !== 0, "El layout legacy no debe aceptarse como fuente moderna.");
    const stderr = `${legacy.stderr || ""}\n${legacy.stdout || ""}`;
    assertCondition(
      stderr.includes("FacturacionServContabilidadDetallado"),
      "El lector no devolvio el mensaje explicito para layout legacy.",
    );

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
