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

function createModernFixture(targetPath) {
  const headers = new Array(39).fill("");
  headers[0] = "AGENCIA";
  headers[1] = "CENTRO";
  headers[2] = "No. ORDEN";
  headers[7] = "TIPO DOC";
  headers[8] = "CEDULA";
  headers[9] = "FACTURADO A";
  headers[11] = "DOCUMENTO";
  headers[14] = "F. FACT";
  headers[17] = "F. NOTA";
  headers[35] = "ANULADA";

  const changanRow = new Array(39).fill("");
  changanRow[0] = "CHANGAN";
  changanRow[1] = "05";
  changanRow[2] = "C10001";
  changanRow[7] = "FA";
  changanRow[8] = "0999999999";
  changanRow[9] = "CLIENTE CHANGAN";
  changanRow[11] = "000001";
  changanRow[14] = "30/03/2026";
  changanRow[18] = 0;
  changanRow[19] = 100;
  changanRow[22] = 100;
  changanRow[27] = 115;
  changanRow[28] = 80;

  const suzukiRiobambaRow = new Array(39).fill("");
  suzukiRiobambaRow[0] = "SUZUKI RIOBAMBA";
  suzukiRiobambaRow[1] = "09";
  suzukiRiobambaRow[2] = "D00168";
  suzukiRiobambaRow[7] = "FA";
  suzukiRiobambaRow[8] = "1799999999001";
  suzukiRiobambaRow[9] = "CLIENTE SUZUKI";
  suzukiRiobambaRow[11] = "000002";
  suzukiRiobambaRow[14] = "30/03/2026";
  suzukiRiobambaRow[18] = 0;
  suzukiRiobambaRow[19] = 40;
  suzukiRiobambaRow[22] = 40;
  suzukiRiobambaRow[27] = 46;
  suzukiRiobambaRow[28] = 12;

  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([headers, changanRow, suzukiRiobambaRow]);
  XLSX.utils.book_append_sheet(workbook, sheet, "Rep");
  XLSX.writeFile(workbook, targetPath);
}

function main() {
  const inputPath = tempPath("servicios_suzuki_riobamba_brand_guard", "xlsx");
  const outputPath = tempPath("servicios_suzuki_riobamba_brand_guard", "json");

  try {
    createModernFixture(inputPath);
    const result = runReader(inputPath, outputPath);
    assertCondition(result.status === 0, `El lector Python fallo.\n${result.stderr || result.stdout}`);
    assertCondition(fs.existsSync(outputPath), "El lector no genero salida JSON.");

    const payload = JSON.parse(fs.readFileSync(outputPath, "utf8"));
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    assertCondition(rows.length === 2, `La fuente de prueba debe generar 2 filas y devolvio ${rows.length}.`);

    const changanRow = rows.find((row) => row.Order === "C10001");
    const suzukiRow = rows.find((row) => row.Order === "D00168");

    assertCondition(changanRow && changanRow.TemplateKey === "changan", "La fila CHANGAN debe seguir clasificando como changan.");
    assertCondition(
      suzukiRow && suzukiRow.TemplateKey === "changan",
      `SUZUKI RIOBAMBA con orden D debe quedar en changan y llego ${suzukiRow ? suzukiRow.TemplateKey : "sin fila"}.`,
    );

    console.log("OK: guardrail de marca para SUZUKI RIOBAMBA validado.");
  } finally {
    for (const filePath of [inputPath, outputPath]) {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    }
  }
}

main();
