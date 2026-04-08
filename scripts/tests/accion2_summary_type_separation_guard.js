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

function round2(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

function tempPath(prefix, extension) {
  return path.join(
    os.tmpdir(),
    `${prefix}_${Date.now()}_${crypto.randomBytes(4).toString("hex")}.${extension}`,
  );
}

function runProcessor(manifestPath) {
  return spawnSync(PYTHON, [path.join(ROOT, "python_services", "cli.py"), manifestPath], {
    cwd: ROOT,
    encoding: "utf8",
  });
}

function cellNumber(sheet, address) {
  const cell = sheet[address];
  if (!cell || cell.v === undefined || cell.v === null || cell.v === "") {
    return 0;
  }
  return round2(Number(cell.v));
}

function main() {
  const inputPath = tempPath("accion2_summary_type_separation", "txt");
  const outputPath = tempPath("accion2_summary_type_separation", "xlsx");
  const manifestPath = tempPath("accion2_summary_type_separation", "json");
  const templatePath = path.join(ROOT, "resources", "cxp", "acciones", "templates", "ACCION2.xlsx");

  try {
    fs.writeFileSync(
      inputPath,
      [
        "NUM RT\tPROVEEDOR\tFECHA\tFECHA CONT\tTIPO\tCOD\tFACT\t%\tBASE\tRETENCION",
        "10001\tPROVEEDOR IVA\t01/01/2026\t01/01/2026\tIVA\t343\tF001\t30\t50.00\t15.00",
        "10002\tPROVEEDOR RENTA\t01/01/2026\t01/01/2026\tRENTA\t332\tF002\t10\t100.00\t10.00",
      ].join("\n"),
      "utf8",
    );

    fs.writeFileSync(
      manifestPath,
      JSON.stringify(
        {
          processor: "cxp_actions.accion2",
          input_paths: [inputPath],
          output_path: outputPath,
          template_path: templatePath,
          options: {},
        },
        null,
        2,
      ),
      "utf8",
    );

    const result = runProcessor(manifestPath);
    assertCondition(result.status === 0, `Accion 2 fallo.\n${result.stderr || result.stdout}`);
    assertCondition(fs.existsSync(outputPath), "Accion 2 no genero salida XLSX.");

    const workbook = XLSX.readFile(outputPath, { cellFormula: true });
    const sheet = workbook.Sheets["RET PROV"];
    assertCondition(!!sheet, "La salida no contiene la hoja RET PROV.");

    assertCondition(cellNumber(sheet, "M2") === 50 && cellNumber(sheet, "N2") === 15, "IVA total debe quedar en la fila IVA.");
    assertCondition(cellNumber(sheet, "M3") === 0 && cellNumber(sheet, "N3") === 0, "RENTA 10% no debe escribirse en IVA 10%.");
    assertCondition(cellNumber(sheet, "M5") === 50 && cellNumber(sheet, "N5") === 15, "IVA 30% debe quedar en su slot.");
    assertCondition(cellNumber(sheet, "M8") === 100 && cellNumber(sheet, "N8") === 10, "RENTA total debe quedar en la fila RENTA.");
    assertCondition(cellNumber(sheet, "M15") === 100 && cellNumber(sheet, "N15") === 10, "RENTA 10% debe quedar en su slot.");
    assertCondition(cellNumber(sheet, "M16") === 150 && cellNumber(sheet, "N16") === 25, "Total general debe sumar ambos bloques.");

    console.log("OK: Accion 2 separa correctamente resumen IVA y RENTA aunque repitan etiquetas.");
  } finally {
    for (const filePath of [inputPath, outputPath, manifestPath]) {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    }
  }
}

main();
