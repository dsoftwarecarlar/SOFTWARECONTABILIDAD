const fs = require("fs");
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

function fixturePath(...parts) {
  return path.join(ROOT, ...parts);
}

function buildFlexibleMayorFile(sourcePath, targetPath) {
  const lines = fs.readFileSync(sourcePath, "utf8").split(/\r?\n/);
  const rewritten = lines.map((line) => {
    if (!line.trim()) {
      return line;
    }

    const columns = line.split("\t");
    const account = String(columns[6] || "").trim();
    if (/^04\.01\.01\.11\.\d{4}$/.test(account)) {
      columns[6] = account.replace("04.01.01.11.", "04.01.01.91.");
    }
    return columns.join("\t");
  });

  fs.writeFileSync(targetPath, `${rewritten.join("\r\n")}\r\n`, "utf8");
}

function getWorkbookAccountSet(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  assertCondition(sheet, `No existe la hoja ${sheetName} en la salida generada.`);
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, blankrows: false });
  return new Set(
    rows
      .map((row) => String((Array.isArray(row) ? row[0] : "") || "").trim())
      .filter(Boolean),
  );
}

function main() {
  const outputDir = fixturePath("storage", "outputs");
  const runStamp = `flexmayortyt_${Date.now()}_${crypto.randomBytes(2).toString("hex")}`;
  const sourceMayor = fixturePath("resources", "cxp", "servicios_marcas", "fixtures", "archivosasubirtoy", "CON_MAYORGEN2TOY.TXT");
  const flexibleMayor = path.join(outputDir, `servicios_flexible_mayor_tyt_${runStamp}.txt`);
  const expectedOutput = path.join(outputDir, `servicios_tyt_${runStamp}.xls`);

  buildFlexibleMayorFile(sourceMayor, flexibleMayor);

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
        flexibleMayor,
      ],
      {
        cwd: ROOT,
        encoding: "utf8",
        timeout: 180000,
      },
    );

    const combinedOutput = `${result.stdout || ""}\n${result.stderr || ""}`;
    assertCondition(result.status === 0, `MATRIZ debe aceptar MAYOR con cuenta alternativa.\n${combinedOutput}`);
    assertCondition(
      combinedOutput.includes("WARN|mayor_account_compatible_section|tyt|04.01.01.91.0001"),
      `El runtime no reporto el mapeo compatible esperado para MATRIZ.\n${combinedOutput}`,
    );
    assertCondition(fs.existsSync(expectedOutput), "No se genero la salida esperada para MATRIZ con MAYOR flexible.");

    const workbook = XLSX.readFile(expectedOutput, { cellFormula: false, cellNF: true, cellText: true });
    const mayVtasSheetName = workbook.SheetNames.includes("MAY VTAS") ? "MAY VTAS" : "VENTAS";
    const accountSet = getWorkbookAccountSet(workbook, mayVtasSheetName);

    for (const expectedAccount of ["04.01.01.91.0001", "04.01.01.91.0003", "04.01.01.91.0010"]) {
      assertCondition(
        accountSet.has(expectedAccount),
        `La hoja ${mayVtasSheetName} no contiene la cuenta flexible ${expectedAccount}.`,
      );
    }

    console.log("OK: MATRIZ acepta y escribe MAYOR VENTAS con numeracion flexible.");
  } finally {
    for (const filePath of [flexibleMayor, expectedOutput]) {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    }
  }
}

main();
