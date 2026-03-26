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

function assertClose(actual, expected, label, tolerance = 0.5) {
  const actualRounded = round2(actual);
  const expectedRounded = round2(expected);
  if (Math.abs(actualRounded - expectedRounded) > tolerance) {
    throw new Error(`${label} esperado=${expectedRounded} actual=${actualRounded}`);
  }
}

function parseDecimalLike(raw) {
  let normalized = String(raw ?? "").replace(/[^\d,.\-]/g, "");
  if (!normalized) {
    return 0;
  }

  const hasComma = normalized.includes(",");
  const hasDot = normalized.includes(".");
  if (hasComma && hasDot) {
    normalized = normalized.lastIndexOf(".") > normalized.lastIndexOf(",")
      ? normalized.replace(/,/g, "")
      : normalized.replace(/\./g, "").replace(",", ".");
  } else if (hasComma) {
    const parts = normalized.split(",");
    normalized = parts.length > 2
      ? normalized.replace(/,/g, "")
      : (parts[1].length === 3 ? parts.join("") : `${parts[0]}.${parts[1]}`);
  } else if ((normalized.match(/\./g) || []).length > 1) {
    const parts = normalized.split(".");
    const frac = parts.pop();
    normalized = frac.length === 3 ? `${parts.join("")}${frac}` : `${parts.join("")}.${frac}`;
  }

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function cellNumber(sheet, address) {
  const cell = sheet?.[address];
  if (!cell) {
    return 0;
  }
  if (typeof cell.v === "number") {
    return round2(cell.v);
  }
  return round2(parseDecimalLike(cell.v));
}

function buildIncompleteMayorFile(sourcePath) {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "servicios-incomplete-mayor-"));
  const outputPath = path.join(tmpDir, "CON_MAYORGEN2TOY_INCOMPLETE.TXT");
  const lines = fs.readFileSync(sourcePath, "utf8").split(/\r?\n/).filter(Boolean);
  const kept = lines.filter((line) => line.includes("\t04.01.01.11.0001\t"));
  assertCondition(kept.length > 0, "No se pudieron extraer filas 0001 del mayor base.");
  fs.writeFileSync(outputPath, `${kept.join(os.EOL)}${os.EOL}`, "utf8");
  return { tmpDir, outputPath };
}

function computeSourceMetricsFromWorkbook(workbook) {
  const repName = workbook.SheetNames.find((name) => name.toUpperCase().includes("REP FACT"));
  const repSheet = workbook.Sheets[repName];
  const noteSheet = workbook.Sheets["NOTA DE CREDITO"];
  const pxSheet = workbook.Sheets["PX"];
  assertCondition(repSheet && noteSheet && pxSheet, "La salida no contiene REP FACTURACION, NOTA DE CREDITO y PX.");

  let invoiceSales = 0;
  let invoiceDiscounts = 0;
  for (let row = 17; row <= 2000; row += 1) {
    if (!repSheet[`C${row}`]) {
      continue;
    }
    invoiceSales += cellNumber(repSheet, `H${row}`);
    invoiceDiscounts += cellNumber(repSheet, `I${row}`);
  }

  let noteSales = 0;
  let noteDiscounts = 0;
  for (let row = 11; row <= 1000; row += 1) {
    if (!noteSheet[`B${row}`]) {
      continue;
    }
    noteSales += cellNumber(noteSheet, `K${row}`) + cellNumber(noteSheet, `L${row}`) + cellNumber(noteSheet, `M${row}`);
    noteDiscounts += cellNumber(noteSheet, `K${row}`);
  }

  let pxGross = 0;
  let pxDiscount = 0;
  for (let row = 1; row <= 1000; row += 1) {
    if (!pxSheet[`E${row}`]) {
      continue;
    }
    pxGross += cellNumber(pxSheet, `L${row}`);
    pxDiscount += cellNumber(pxSheet, `N${row}`);
  }

  invoiceSales = round2(invoiceSales - pxGross);
  invoiceDiscounts = round2(invoiceDiscounts - pxDiscount);

  return {
    InvoiceSales: invoiceSales,
    InvoiceDiscounts: invoiceDiscounts,
    NoteSales: round2(noteSales),
    NoteDiscounts: round2(noteDiscounts),
    NetSales: round2(invoiceSales - invoiceDiscounts - noteSales + noteDiscounts),
  };
}

function run() {
  const fixtureMayor = path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "archivosasubirtoy", "CON_MAYORGEN2TOY.TXT");
  const { tmpDir, outputPath: incompleteMayorPath } = buildIncompleteMayorFile(fixtureMayor);
  const outputDir = path.join(ROOT, "storage", "outputs");
  const runStamp = `incomplete_mayor_${Date.now()}_${crypto.randomBytes(2).toString("hex")}`;
  const outputPath = path.join(outputDir, `servicios_tyt_${runStamp}.xls`);

  try {
    const result = spawnSync(
      PYTHON,
      [
        path.join(ROOT, "python_services", "processors", "servicios_marcas", "runtime.py"),
        "--input",
        path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "RepFacturacionServContabilidad (3).xls"),
        "--output-dir",
        outputDir,
        "--template-dir",
        path.join(ROOT, "resources", "cxp", "servicios_marcas", "templates"),
        "--run-stamp",
        runStamp,
        "--brand-key",
        "tyt",
        "--px-path",
        path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "detalle-vtas-xliquidar (2).xlsx"),
        "--repvtas-path",
        path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "RepFacturacionServContabilidad (3).xls"),
        "--factura-tyt-path",
        path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "archivosasubirtoy", "SERREP_FACTURAS_NAFTOY.TXT"),
        "--nota-tyt-path",
        path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "archivosasubirtoy", "SERREP_NOTACRED_NAFTOY.TXT"),
        "--mayor-tyt-path",
        incompleteMayorPath,
      ],
      {
        cwd: ROOT,
        encoding: "utf8",
        timeout: 600000,
      },
    );

    const output = `${result.stdout || ""}\n${result.stderr || ""}`;
    assertCondition(result.status === 0, `El runtime fallo con mayor incompleto.\n${output}`);
    assertCondition(fs.existsSync(outputPath), `No se genero ${outputPath}`);

    const workbook = XLSX.readFile(outputPath, { cellFormula: true });
    const repName = workbook.SheetNames.find((name) => name.toUpperCase().includes("REP FACT"));
    const repSheet = workbook.Sheets[repName];
    const noteSheet = workbook.Sheets["NOTA DE CREDITO"];
    const repVtasSheet = workbook.Sheets["REP VTAS"];
    const mayorSheet = workbook.Sheets["MAY VTAS"];
    const sourceMetrics = computeSourceMetricsFromWorkbook(workbook);

    assertClose(cellNumber(repSheet, "D9"), sourceMetrics.InvoiceSales, "REP FACT D9");
    assertClose(cellNumber(repSheet, "E9"), sourceMetrics.InvoiceDiscounts, "REP FACT E9");
    assertClose(cellNumber(repSheet, "J9"), sourceMetrics.InvoiceSales, "REP FACT J9");
    assertClose(cellNumber(repSheet, "K9"), sourceMetrics.InvoiceDiscounts, "REP FACT K9");
    assertClose(cellNumber(noteSheet, "D4"), sourceMetrics.NoteSales, "NOTA D4");
    assertClose(cellNumber(noteSheet, "E4"), sourceMetrics.NoteDiscounts, "NOTA E4");
    assertClose(cellNumber(noteSheet, "F4"), sourceMetrics.NoteSales, "NOTA F4");
    assertClose(cellNumber(noteSheet, "G4"), sourceMetrics.NoteDiscounts, "NOTA G4");
    assertClose(cellNumber(repVtasSheet, "D6"), sourceMetrics.NetSales, "REP VTAS D6");
    assertCondition(Math.abs(cellNumber(mayorSheet, "J4") - sourceMetrics.InvoiceSales) > 1, "La prueba no forzo una diferencia real entre mayor y control visible.");

    console.log("OK: controles visibles usan fallback transaccional cuando el MAYOR viene incompleto.");
  } finally {
    if (fs.existsSync(outputPath)) {
      fs.unlinkSync(outputPath);
    }
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
}

run();
