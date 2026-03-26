const fs = require("fs");
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
    repventas: [
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "RepFacturacionServContabilidad (3).xls"),
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "RepFacturacionServContabilidad.xls"),
    ],
    facturaTyt: [
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "archivosasubirtoy", "SERREP_FACTURAS_NAFTOY.TXT"),
    ],
    notaTyt: [
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "archivosasubirtoy", "SERREP_NOTACRED_NAFTOY.TXT"),
    ],
    templateDir: [
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "templates"),
    ],
  };

  return firstExistingPath(candidatesByKey[fileKey] || []);
}

function buildInvalidMayorFile(targetPath) {
  const columns = new Array(30).fill("");
  columns[6] = "02.01.06.01.0018";
  columns[7] = "15% IVA";
  columns[21] = "N";
  columns[22] = "31-JAN-26";
  columns[23] = "AGCM";
  columns[24] = "485";
  columns[26] = "LIQUIDACION DE IMPUESTOS ENERO 2026";
  columns[27] = "11,321.05";
  columns[28] = "0.00";
  columns[29] = "11,321.05";
  fs.writeFileSync(targetPath, `${columns.join("\t")}\r\n`, "utf8");
}

function main() {
  const pxPath = serviciosFixturePath("px");
  const repVtasPath = serviciosFixturePath("repventas");
  const facturaTytPath = serviciosFixturePath("facturaTyt");
  const notaTytPath = serviciosFixturePath("notaTyt");
  const templateDir = serviciosFixturePath("templateDir");
  const outputDir = path.join(ROOT, "storage", "outputs");
  const runStamp = `badmayor_${Date.now()}_${crypto.randomBytes(2).toString("hex")}`;
  const badMayorPath = path.join(outputDir, `servicios_bad_mayor_${runStamp}.txt`);
  const expectedOutput = path.join(outputDir, `servicios_tyt_${runStamp}.xls`);

  assertCondition(fs.existsSync(pxPath), "No existe fixture PX para la prueba de guardrail del mayor.");
  assertCondition(fs.existsSync(repVtasPath), "No existe fixture REP VENTAS para la prueba de guardrail del mayor.");
  assertCondition(fs.existsSync(facturaTytPath), "No existe fixture FACTURAS TOYOTA para la prueba de guardrail del mayor.");
  assertCondition(fs.existsSync(notaTytPath), "No existe fixture NOTAS TOYOTA para la prueba de guardrail del mayor.");
  assertCondition(fs.existsSync(templateDir), "No existe directorio de plantillas para la prueba de guardrail del mayor.");

  buildInvalidMayorFile(badMayorPath);

  try {
    const result = spawnSync(
      PYTHON,
      [
        path.join(ROOT, "python_services", "processors", "servicios_marcas", "runtime.py"),
        "--input",
        repVtasPath,
        "--output-dir",
        outputDir,
        "--template-dir",
        templateDir,
        "--run-stamp",
        runStamp,
        "--brand-key",
        "tyt",
        "--px-path",
        pxPath,
        "--repvtas-path",
        repVtasPath,
        "--factura-tyt-path",
        facturaTytPath,
        "--nota-tyt-path",
        notaTytPath,
        "--mayor-tyt-path",
        badMayorPath,
      ],
      {
        cwd: ROOT,
        encoding: "utf8",
        timeout: 180000,
      },
    );

    const combinedOutput = `${result.stdout || ""}\n${result.stderr || ""}`;
    assertCondition(result.status !== 0, "El runtime no debe aceptar un MAYOR que no corresponde a ventas.");
    assertCondition(
      combinedOutput.includes("no corresponde al mayor de ventas de la plantilla"),
      `El runtime no devolvio el mensaje claro esperado.\n${combinedOutput}`,
    );
    assertCondition(
      !fs.existsSync(expectedOutput),
      "El runtime dejo una salida parcial guardada pese a rechazar el MAYOR invalido.",
    );

    console.log("OK: guardrail de MAYOR invalido en Ventana 2 validado.");
  } finally {
    for (const filePath of [badMayorPath, expectedOutput]) {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    }
  }
}

main();
