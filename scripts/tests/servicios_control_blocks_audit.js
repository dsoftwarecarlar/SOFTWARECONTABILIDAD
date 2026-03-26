const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { spawnSync } = require("child_process");
const XLSX = require("xlsx");

const ROOT = path.resolve(__dirname, "..", "..");
const PYTHON = process.env.PYTHON_BINARY || "python";

const BRANDS = [
  {
    key: "changan",
    prefix: "servicios_changan_",
    fixtureDir: path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "archivosasubirchan"),
    facturaFlag: "--factura-changan-path",
    notaFlag: "--nota-changan-path",
    mayorFlag: "--mayor-changan-path",
  },
  {
    key: "peug",
    prefix: "servicios_peug_",
    fixtureDir: path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "archivosasubirpeu"),
    facturaFlag: "--factura-peug-path",
    notaFlag: "--nota-peug-path",
    mayorFlag: "--mayor-peug-path",
  },
  {
    key: "szk",
    prefix: "servicios_szk_",
    fixtureDir: path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "archivosasubirsuz"),
    facturaFlag: "--factura-szk-path",
    notaFlag: "--nota-szk-path",
    mayorFlag: "--mayor-szk-path",
  },
  {
    key: "tyt",
    prefix: "servicios_tyt_",
    fixtureDir: path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "archivosasubirtoy"),
    facturaFlag: "--factura-tyt-path",
    notaFlag: "--nota-tyt-path",
    mayorFlag: "--mayor-tyt-path",
  },
];

function assertCondition(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

function round2(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

function assertClose(actual, expected, label, tolerance = 0.02) {
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
    if (normalized.lastIndexOf(".") > normalized.lastIndexOf(",")) {
      normalized = normalized.replace(/,/g, "");
    } else {
      normalized = normalized.replace(/\./g, "").replace(",", ".");
    }
  } else if (hasComma) {
    const commaCount = (normalized.match(/,/g) || []).length;
    if (commaCount > 1) {
      normalized = normalized.replace(/,/g, "");
    } else {
      const [intPart, fracPart = ""] = normalized.split(",");
      normalized = fracPart.length === 3 ? `${intPart}${fracPart}` : `${intPart}.${fracPart}`;
    }
  } else if (hasDot) {
    const dotCount = (normalized.match(/\./g) || []).length;
    if (dotCount > 1) {
      const lastDot = normalized.lastIndexOf(".");
      const intPart = normalized.slice(0, lastDot).replace(/\./g, "");
      const fracPart = normalized.slice(lastDot + 1);
      normalized = fracPart.length === 3 ? `${intPart}${fracPart}` : `${intPart}.${fracPart}`;
    }
  }

  const parsed = Number(normalized);
  if (!Number.isFinite(parsed)) {
    return 0;
  }
  return parsed;
}

function readMayorRows(filePath) {
  return fs
    .readFileSync(filePath, "utf8")
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => line.split("\t"))
    .filter((columns) => columns.length >= 30 && /^\d{2}\.\d{2}\.\d{2}\.\d{2}\.\d{4}$/.test((columns[6] || "").trim()))
    .map((columns) => ({
      account: (columns[6] || "").trim(),
      name: (columns[7] || "").trim(),
      origin: (columns[23] || "").trim().toUpperCase(),
      seat: (columns[24] || "").trim(),
      detail: (columns[26] || "").trim().toUpperCase(),
      debit: parseDecimalLike(columns[27]),
      credit: parseDecimalLike(columns[28]),
    }));
}

function compactAccountCode(account) {
  const digits = String(account || "").replace(/\D/g, "");
  if (digits.length > 0 && digits.length < 12) {
    return digits.padStart(12, "0");
  }
  return digits;
}

function isMayorPxAdjustmentRow(row) {
  const account = compactAccountCode(row.account);
  if (!/^040101\d{2}(0003|0012)$/.test(account)) {
    return false;
  }
  if (String(row.detail || "").includes("REGISTRO DE PX AJUSTE DE EGRESO")) {
    return true;
  }
  return String(row.origin || "") === "AGCM" && String(row.seat || "") === "435";
}

function filterMayorRowsForWorkbook(rows) {
  return rows.filter((row) => !isMayorPxAdjustmentRow(row));
}

function classifyControlBucket(accountValue, nameValue) {
  const account = String(accountValue || "").trim();
  const name = String(nameValue || "").trim().toUpperCase();
  const compact = compactAccountCode(account);
  if (!account && !name) {
    return "";
  }
  if (/^010105\d{2}\d{4}$/.test(compact) || name.includes("GARANT")) {
    return "guarantee";
  }
  if (!/^040101\d{2}\d{4}$/.test(compact)) {
    return "";
  }
  const suffix = compact.slice(-4);
  if (suffix === "0014" || name.includes("DEVOL")) {
    return "return";
  }
  if (["0010", "0011", "0012"].includes(suffix) || name.includes("DESC")) {
    return "discount";
  }
  if (["0001", "0002", "0003"].includes(suffix) || name.includes("VTAS")) {
    return "sales";
  }
  return "";
}

function getControlMetrics(rows, { accountKey, nameKey, debitKey, creditKey }) {
  const metrics = {
    InvoiceSales: 0,
    InvoiceDiscounts: 0,
    NoteSales: 0,
    NoteDiscounts: 0,
    NetSales: 0,
  };

  for (const row of rows) {
    const bucket = classifyControlBucket(row[accountKey], row[nameKey]);
    const debit = Number(row[debitKey] || 0);
    const credit = Number(row[creditKey] || 0);
    if (bucket === "sales") {
      metrics.InvoiceSales += credit;
      metrics.NetSales += credit - debit;
      continue;
    }
    if (bucket === "discount") {
      metrics.InvoiceDiscounts += debit;
      metrics.NoteDiscounts += credit;
      metrics.NetSales += credit - debit;
      continue;
    }
    if (bucket === "return") {
      metrics.NoteSales += debit;
      metrics.NetSales += credit - debit;
    }
  }

  return Object.fromEntries(Object.entries(metrics).map(([key, value]) => [key, round2(value)]));
}

function getMayorControlMetrics(rows) {
  return getControlMetrics(rows, {
    accountKey: "account",
    nameKey: "name",
    debitKey: "debit",
    creditKey: "credit",
  });
}

function pickSheet(workbook, candidates) {
  for (const name of candidates) {
    if (workbook.Sheets[name]) {
      return workbook.Sheets[name];
    }
  }
  return null;
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

function getSheetRows(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, blankrows: false });
}

function readWorkbookMayorMetrics(workbook) {
  const mayorSheet = pickSheet(workbook, ["MAY VTAS", "VENTAS"]);
  assertCondition(!!mayorSheet, "La salida no contiene MAY VTAS/VENTAS para auditar controles.");
  const rows = getSheetRows(mayorSheet)
    .map((row) => (Array.isArray(row) ? row : []))
    .map((row) => ({
      account: String(row[0] || "").trim(),
      name: String(row[1] || "").trim(),
      debit: parseDecimalLike(row[8]),
      credit: parseDecimalLike(row[9]),
    }))
    .filter((row) => row.account !== "");
  return getMayorControlMetrics(rows);
}

function readWorkbookPrecontMetrics(workbook) {
  const precontSheet = pickSheet(workbook, ["PrecontabilizacionVentas"]);
  assertCondition(!!precontSheet, "La salida no contiene PrecontabilizacionVentas para auditar controles.");
  const rows = getSheetRows(precontSheet)
    .map((row) => (Array.isArray(row) ? row : []))
    .map((row) => ({
      Account: String(row[4] || "").trim(),
      Description: String(row[5] || "").trim(),
      Debit: parseDecimalLike(row[7]),
      Credit: parseDecimalLike(row[8]),
    }))
    .filter((row) => row.Account !== "");
  return getControlMetrics(rows, {
    accountKey: "Account",
    nameKey: "Description",
    debitKey: "Debit",
    creditKey: "Credit",
  });
}

function readWorkbookSourceMetrics(workbook) {
  const repSheet = pickSheet(workbook, ["REP FACTURACIÓN", "REP FACTURACION"]);
  const noteSheet = pickSheet(workbook, ["NOTA DE CREDITO"]);
  const pxSheet = pickSheet(workbook, ["PX"]);
  assertCondition(!!repSheet && !!noteSheet && !!pxSheet, "La salida no contiene hojas suficientes para validar fallback transaccional.");

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

function resolveExpectedControlMetrics(workbook) {
  const mayorMetrics = readWorkbookMayorMetrics(workbook);
  const sourceMetrics = readWorkbookSourceMetrics(workbook);
  const useSource = ["InvoiceSales", "InvoiceDiscounts", "NoteSales", "NoteDiscounts"].some(
    (key) => Math.abs(round2(sourceMetrics[key]) - round2(mayorMetrics[key])) > 1.0,
  );
  return useSource ? sourceMetrics : mayorMetrics;
}

function findFixtureFile(directory, prefix) {
  const match = fs
    .readdirSync(directory)
    .find((fileName) => fileName.toUpperCase().startsWith(prefix) && fileName.toUpperCase().endsWith(".TXT"));
  assertCondition(match, `No se encontro fixture ${prefix} en ${directory}`);
  return path.join(directory, match);
}

function runBrandAudit(brand) {
  const outputDir = path.join(ROOT, "storage", "outputs");
  const runStamp = `ctrl_${brand.key}_${Date.now()}_${crypto.randomBytes(2).toString("hex")}`;
  const expectedOutput = path.join(outputDir, `${brand.prefix}${runStamp}.xls`);
  const repventasPath = path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "RepFacturacionServContabilidad (3).xls");
  const pxPath = path.join(ROOT, "resources", "cxp", "servicios_marcas", "fixtures", "detalle-vtas-xliquidar (2).xlsx");
  const facturaPath = findFixtureFile(brand.fixtureDir, "SERREP_FACTURAS");
  const notaPath = findFixtureFile(brand.fixtureDir, "SERREP_NOTACRED");
  const mayorPath = findFixtureFile(brand.fixtureDir, "CON_MAYORGEN2");

  const result = spawnSync(
    PYTHON,
    [
      path.join(ROOT, "python_services", "processors", "servicios_marcas", "runtime.py"),
      "--input",
      repventasPath,
      "--output-dir",
      outputDir,
      "--template-dir",
      path.join(ROOT, "resources", "cxp", "servicios_marcas", "templates"),
      "--run-stamp",
      runStamp,
      "--brand-key",
      brand.key,
      "--px-path",
      pxPath,
      "--repvtas-path",
      repventasPath,
      brand.facturaFlag,
      facturaPath,
      brand.notaFlag,
      notaPath,
      brand.mayorFlag,
      mayorPath,
    ],
    {
      cwd: ROOT,
      encoding: "utf8",
      timeout: 240000,
    },
  );

  const combinedOutput = `${result.stdout || ""}\n${result.stderr || ""}`;
  assertCondition(result.status === 0, `La auditoria de control fallo para ${brand.key}.\n${combinedOutput}`);
  assertCondition(fs.existsSync(expectedOutput), `No se genero salida para ${brand.key}: ${expectedOutput}`);

  try {
    const workbook = XLSX.readFile(expectedOutput, { cellFormula: false, cellNF: false, cellStyles: false });
    const repSheet = pickSheet(workbook, ["REP FACTURACIÓN", "REP FACTURACION"]);
    const noteSheet = pickSheet(workbook, ["NOTA DE CREDITO"]);
    const repVtasSheet = pickSheet(workbook, ["REP VTAS"]);
    const costoSheet = pickSheet(workbook, ["COSTO"]);
    assertCondition(repSheet && noteSheet && repVtasSheet && costoSheet, `Faltan hojas de control en la salida de ${brand.key}.`);

    const expectedMetrics = resolveExpectedControlMetrics(workbook);

    assertClose(cellNumber(repSheet, "D9"), expectedMetrics.InvoiceSales, `${brand.key} REP FACTURACION D9`, 0.5);
    assertClose(cellNumber(repSheet, "E9"), expectedMetrics.InvoiceDiscounts, `${brand.key} REP FACTURACION E9`, 0.5);
    assertClose(cellNumber(repSheet, "J9"), expectedMetrics.InvoiceSales, `${brand.key} REP FACTURACION J9`, 0.5);
    assertClose(cellNumber(repSheet, "K9"), expectedMetrics.InvoiceDiscounts, `${brand.key} REP FACTURACION K9`, 0.5);

    assertClose(cellNumber(noteSheet, "D4"), expectedMetrics.NoteSales, `${brand.key} NOTA D4`, 0.5);
    assertClose(cellNumber(noteSheet, "E4"), expectedMetrics.NoteDiscounts, `${brand.key} NOTA E4`, 0.5);
    assertClose(cellNumber(noteSheet, "F4"), expectedMetrics.NoteSales, `${brand.key} NOTA F4`, 0.5);
    assertClose(cellNumber(noteSheet, "G4"), expectedMetrics.NoteDiscounts, `${brand.key} NOTA G4`, 0.5);

    assertClose(cellNumber(repVtasSheet, "D6"), expectedMetrics.NetSales, `${brand.key} REP VTAS D6`, 0.5);
    assertClose(cellNumber(repVtasSheet, "E6"), cellNumber(costoSheet, "J4"), `${brand.key} REP VTAS E6`);
  } finally {
    if (fs.existsSync(expectedOutput)) {
      fs.unlinkSync(expectedOutput);
    }
  }
}

function main() {
  for (const brand of BRANDS) {
    runBrandAudit(brand);
  }
  console.log("OK: bloques SEGUN MAYOR validados en las 4 plantillas de Ventana 2.");
}

main();
