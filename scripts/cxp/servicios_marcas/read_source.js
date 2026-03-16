const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const { parseDecimalLike, sanitizeText } = require("../shared/core-utils");

function parseArgs(argv) {
  const args = {};
  for (let index = 2; index < argv.length; index += 1) {
    const current = String(argv[index] || "");
    if (!current.startsWith("-")) {
      continue;
    }

    const key = current.replace(/^-+/, "").toLowerCase();
    const next = argv[index + 1];
    if (typeof next === "string" && !next.startsWith("-")) {
      args[key] = next;
      index += 1;
      continue;
    }

    args[key] = "";
  }

  return args;
}

function excelDateSerial(date) {
  const utcMillis = Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
  return utcMillis / 86400000 + 25569;
}

function parseDateText(value) {
  const text = sanitizeText(value);
  if (!text) {
    return null;
  }

  const match = text.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/);
  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]);
    const year = Number(match[3].length === 2 ? `20${match[3]}` : match[3]);
    const parsed = new Date(Date.UTC(year, month - 1, day));
    if (!Number.isNaN(parsed.getTime())) {
      return excelDateSerial(parsed);
    }
  }

  const parsed = new Date(text);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }

  return excelDateSerial(parsed);
}

function cellAddress(row, column) {
  return XLSX.utils.encode_cell({ r: row - 1, c: column - 1 });
}

function getCell(sheet, row, column) {
  return sheet[cellAddress(row, column)] || null;
}

function cellText(sheet, row, column) {
  const cell = getCell(sheet, row, column);
  if (!cell) {
    return "";
  }

  if (cell.w != null) {
    return sanitizeText(cell.w);
  }

  if (cell.v == null) {
    return "";
  }

  return sanitizeText(cell.v);
}

function cellNumber(sheet, row, column) {
  const cell = getCell(sheet, row, column);
  if (!cell || cell.v == null || cell.v === "") {
    return 0;
  }

  if (typeof cell.v === "number" && Number.isFinite(cell.v)) {
    return cell.v;
  }

  return parseDecimalLike(cell.w != null ? cell.w : cell.v);
}

function cellDateValue(sheet, row, column) {
  const cell = getCell(sheet, row, column);
  if (!cell || cell.v == null || cell.v === "") {
    return null;
  }

  if (typeof cell.v === "number" && Number.isFinite(cell.v)) {
    return cell.v;
  }

  if (cell.v instanceof Date && !Number.isNaN(cell.v.getTime())) {
    return excelDateSerial(cell.v);
  }

  return parseDateText(cell.w != null ? cell.w : cell.v);
}

function trimDocument(value) {
  const text = sanitizeText(value);
  if (!text) {
    return "";
  }

  const digits = text.replace(/[^\d]/g, "");
  if (!digits) {
    return text;
  }

  const trimmed = digits.replace(/^0+/, "");
  return trimmed || "0";
}

function getTemplateKey(agency, order) {
  const agencyText = sanitizeText(agency).toUpperCase();
  const orderText = sanitizeText(order).toUpperCase();

  switch (agencyText) {
    case "CHANGAN":
      return "changan";
    case "PEUGEOT":
      return "peug";
    case "MATRIZ":
      return "tyt";
    case "SUZUKI AMBATO":
      return "szk";
    case "SUZUKI RIOBAMBA":
      return /^D\d+/.test(orderText) ? "changan" : "szk";
    default:
      return null;
  }
}

function validateHeaders(sheet) {
  const checks = [
    { row: 1, column: 1, needle: "AGENCIA" },
    { row: 1, column: 2, needle: "CENTRO" },
    { row: 1, column: 3, needle: "No. ORDEN" },
    { row: 1, column: 8, needle: "TIPO DOC" },
    { row: 1, column: 9, needle: "CEDULA" },
    { row: 1, column: 10, needle: "FACTURADO A" },
    { row: 1, column: 12, needle: "DOCUMENTO" },
    { row: 1, column: 15, needle: "F. FACT" },
    { row: 1, column: 18, needle: "F. NOTA" },
    { row: 1, column: 36, needle: "ANULADA" },
  ];

  for (const check of checks) {
    const actual = cellText(sheet, check.row, check.column).toUpperCase();
    if (!actual.includes(check.needle.toUpperCase())) {
      throw new Error(
        `El archivo fuente no coincide con la estructura esperada en fila ${check.row} columna ${check.column}. Esperado contiene '${check.needle}' y llego '${actual}'.`,
      );
    }
  }
}

function readRows(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const lastRow = range.e.r + 1;
  const rows = [];

  for (let row = 2; row <= lastRow; row += 1) {
    const agency = cellText(sheet, row, 1);
    if (!agency) {
      continue;
    }

    const order = cellText(sheet, row, 3);
    let series = cellText(sheet, row, 14);
    if (!series) {
      series = cellText(sheet, row, 13);
    }

    const templateKey = getTemplateKey(agency, order);
    if (!templateKey) {
      continue;
    }

    const docType = cellText(sheet, row, 8).toUpperCase();
    if (!["FA", "FC", "DC", "DE"].includes(docType)) {
      continue;
    }

    const anulada = cellText(sheet, row, 36).toUpperCase();
    if (["SI", "S", "YES", "Y", "ANULADA"].includes(anulada)) {
      continue;
    }

    const documentRaw = cellText(sheet, row, 12);
    if (!documentRaw) {
      continue;
    }

    const affectedRaw = cellText(sheet, row, 37);

    rows.push({
      RowIndex: row,
      TemplateKey: templateKey,
      Agency: agency,
      AgencyRaw: cellText(sheet, row, 1),
      Center: cellText(sheet, row, 2),
      CenterRaw: cellText(sheet, row, 2),
      Order: order,
      OrderRaw: cellText(sheet, row, 3),
      Advisor: cellText(sheet, row, 5),
      AdvisorRaw: cellText(sheet, row, 5),
      Line: cellText(sheet, row, 7),
      LineRaw: cellText(sheet, row, 7),
      DocType: docType,
      Cedula: cellText(sheet, row, 9),
      CedulaRaw: cellText(sheet, row, 9),
      Customer: cellText(sheet, row, 10),
      CustomerRaw: cellText(sheet, row, 10),
      DocumentRaw: documentRaw,
      DocumentTrim: trimDocument(documentRaw),
      Series: series,
      SeriesRaw: series,
      FormaPago: cellText(sheet, row, 16),
      Authorization: cellText(sheet, row, 17),
      DateFactValue: cellDateValue(sheet, row, 15),
      DateNoteValue: cellDateValue(sheet, row, 18),
      NoteCredit: cellNumber(sheet, row, 19),
      TotalManoObra: cellNumber(sheet, row, 20),
      TotalSubcontratos: cellNumber(sheet, row, 21),
      TotalInsumos: cellNumber(sheet, row, 22),
      TotalServicio: cellNumber(sheet, row, 23),
      TotalAccesorios: cellNumber(sheet, row, 24),
      TotalRepuestos: cellNumber(sheet, row, 25),
      Interes: cellNumber(sheet, row, 26),
      Iva: cellNumber(sheet, row, 27),
      Total: cellNumber(sheet, row, 28),
      Costo: cellNumber(sheet, row, 29),
      CostoLubricantes: cellNumber(sheet, row, 30),
      CostoAccesorios: cellNumber(sheet, row, 31),
      CostoRepuestos: cellNumber(sheet, row, 32),
      CostoPintura: cellNumber(sheet, row, 33),
      CostoSubconNc: cellNumber(sheet, row, 34),
      GarExt: cellText(sheet, row, 35),
      GarExtRaw: cellText(sheet, row, 35),
      Anulada: anulada,
      AffectedDocumentTrim: trimDocument(affectedRaw),
      AffectedDocumentRaw: affectedRaw,
      MotivoNc: cellText(sheet, row, 38),
      ObservacionNc: cellText(sheet, row, 39),
    });
  }

  return rows;
}

function main() {
  const args = parseArgs(process.argv);
  const inputPath = path.resolve(process.cwd(), args.input || "");
  const outputJsonPath = path.resolve(process.cwd(), args["output-json"] || "");

  if (!inputPath || !fs.existsSync(inputPath)) {
    throw new Error(`No existe InputPath: ${inputPath}`);
  }

  if (!outputJsonPath) {
    throw new Error("Falta --output-json para exportar filas.");
  }

  const workbook = XLSX.readFile(inputPath, {
    cellDates: false,
    cellFormula: false,
    cellNF: true,
    cellText: true,
    sheetStubs: true,
  });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("El archivo fuente no contiene hojas.");
  }

  const sheet = workbook.Sheets[sheetName];
  validateHeaders(sheet);
  const rows = readRows(sheet);

  fs.mkdirSync(path.dirname(outputJsonPath), { recursive: true });
  fs.writeFileSync(
    outputJsonPath,
    `${JSON.stringify({ sheet_name: sheetName, rows }, null, 2)}\n`,
    "utf8",
  );

  console.log(`INFO|source_read|rows=${rows.length}`);
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(error && error.message ? error.message : String(error));
    process.exit(1);
  }
}
