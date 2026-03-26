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

function fixturePath(...parts) {
  return path.join(ROOT, ...parts);
}

function buildMixedMayorFile(sourcePath, targetPath) {
  const sourceLines = fs.readFileSync(sourcePath, "utf8").split(/\r?\n/).filter(Boolean);
  const injected = new Array(30).fill("");
  injected[6] = "02.01.06.01.0018";
  injected[7] = "15% IVA";
  injected[21] = "N";
  injected[22] = "28-FEB-26";
  injected[23] = "AGCM";
  injected[24] = "485";
  injected[26] = "LIQUIDACION DE IMPUESTOS FEBRERO 2026";
  injected[27] = "321.05";
  injected[28] = "0.00";
  injected[29] = "321.05";
  fs.writeFileSync(targetPath, [injected.join("\t"), ...sourceLines].join("\r\n") + "\r\n", "utf8");
}

function main() {
  const outputDir = fixturePath("storage", "outputs");
  const runStamp = `mixedmayor_${Date.now()}_${crypto.randomBytes(2).toString("hex")}`;
  const sourceMayor = fixturePath("resources", "cxp", "servicios_marcas", "fixtures", "archivosasubirtoy", "CON_MAYORGEN2TOY.TXT");
  const mixedMayor = path.join(outputDir, `servicios_mixed_mayor_${runStamp}.txt`);
  const expectedOutput = path.join(outputDir, `servicios_tyt_${runStamp}.xls`);

  buildMixedMayorFile(sourceMayor, mixedMayor);

  try {
    const result = spawnSync(
      PYTHON,
      [
        fixturePath("python_services", "processors", "servicios_marcas", "runtime.py"),
        "--input",
        fixturePath("resources", "cxp", "servicios_marcas", "fixtures", "RepFacturacionServContabilidad (3).xls"),
        "--output-dir",
        outputDir,
        "--template-dir",
        fixturePath("resources", "cxp", "servicios_marcas", "templates"),
        "--run-stamp",
        runStamp,
        "--brand-key",
        "tyt",
        "--px-path",
        fixturePath("resources", "cxp", "servicios_marcas", "fixtures", "detalle-vtas-xliquidar (2).xlsx"),
        "--repvtas-path",
        fixturePath("resources", "cxp", "servicios_marcas", "fixtures", "RepFacturacionServContabilidad (3).xls"),
        "--factura-tyt-path",
        fixturePath("resources", "cxp", "servicios_marcas", "fixtures", "archivosasubirtoy", "SERREP_FACTURAS_NAFTOY.TXT"),
        "--nota-tyt-path",
        fixturePath("resources", "cxp", "servicios_marcas", "fixtures", "archivosasubirtoy", "SERREP_NOTACRED_NAFTOY.TXT"),
        "--mayor-tyt-path",
        mixedMayor,
      ],
      {
        cwd: ROOT,
        encoding: "utf8",
        timeout: 180000,
      },
    );

    const combinedOutput = `${result.stdout || ""}\n${result.stderr || ""}`;
    assertCondition(result.status === 0, `El runtime debe aceptar un MAYOR con filas utiles y cuentas extra.\n${combinedOutput}`);
    assertCondition(
      combinedOutput.includes("WARN|mayor_accounts_ignored|tyt|accounts=02.01.06.01.0018 (15% IVA)"),
      `El runtime no reporto el descarte de la cuenta extra.\n${combinedOutput}`,
    );
    assertCondition(fs.existsSync(expectedOutput), "El runtime no genero la salida esperada con el MAYOR mixto.");

    console.log("OK: guardrail de MAYOR mixto en Ventana 2 validado.");
  } finally {
    for (const filePath of [mixedMayor, expectedOutput]) {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    }
  }
}

main();
