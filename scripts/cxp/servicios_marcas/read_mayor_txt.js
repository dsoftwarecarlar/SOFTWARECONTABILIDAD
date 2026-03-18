const fs = require("fs");
const path = require("path");

function parseArgs(argv) {
  const args = {};
  for (let index = 2; index < argv.length; index += 1) {
    const key = String(argv[index] || "");
    if (!key.startsWith("-")) {
      continue;
    }

    const name = key.replace(/^-+/, "").toLowerCase();
    const next = argv[index + 1];
    if (next && !String(next).startsWith("-")) {
      args[name] = next;
      index += 1;
    } else {
      args[name] = "";
    }
  }

  return args;
}

function sanitize(value) {
  return String(value ?? "").replace(/\uFEFF/g, "").trim();
}

function parseDecimalLike(value) {
  let text = sanitize(value).replace(/\s+/g, "");
  if (text === "") {
    return 0;
  }

  const hasComma = text.includes(",");
  const hasDot = text.includes(".");
  if (hasComma && hasDot) {
    if (text.lastIndexOf(".") > text.lastIndexOf(",")) {
      text = text.replace(/,/g, "");
    } else {
      text = text.replace(/\./g, "").replace(",", ".");
    }
  } else if (hasComma) {
    text = /,\d{2}$/.test(text) ? text.replace(",", ".") : text.replace(/,/g, "");
  } else if (hasDot) {
    text = /\.\d{2}$/.test(text) ? text : text.replace(/\./g, "");
  }

  const numeric = Number(text);
  return Number.isFinite(numeric) ? numeric : 0;
}

function parseReportDateToExcelSerial(value) {
  const text = sanitize(value).toUpperCase();
  const match = /^(\d{2})-([A-Z]{3})-(\d{2})$/.exec(text);
  if (!match) {
    return null;
  }

  const monthMap = {
    JAN: 0,
    ENE: 0,
    FEB: 1,
    MAR: 2,
    APR: 3,
    ABR: 3,
    MAY: 4,
    JUN: 5,
    JUL: 6,
    AUG: 7,
    AGO: 7,
    SEP: 8,
    SET: 8,
    OCT: 9,
    NOV: 10,
    DEC: 11,
    DIC: 11,
  };

  const day = Number(match[1]);
  const month = monthMap[match[2]];
  const year = 2000 + Number(match[3]);
  if (!Number.isInteger(day) || month == null) {
    return null;
  }

  const date = new Date(Date.UTC(year, month, day));
  if (!Number.isFinite(date.getTime())) {
    return null;
  }

  return (date.getTime() / 86400000) + 25569;
}

function parseMayorRows(inputPath) {
  const lines = fs.readFileSync(inputPath, "utf8")
    .split(/\r?\n/)
    .map((line) => line.replace(/\uFEFF/g, "").trimEnd())
    .filter((line) => sanitize(line) !== "");

  const rows = [];
  for (const line of lines) {
    const cols = line.split("\t");
    if (cols.length < 30) {
      continue;
    }

    const account = sanitize(cols[6]);
    const dateText = sanitize(cols[22]).toUpperCase();
    if (!/^\d{2}\.\d{2}\.\d{2}\.\d{2}\.\d{4}$/.test(account) || !/^\d{2}-[A-Z]{3}-\d{2}$/.test(dateText)) {
      continue;
    }

    rows.push({
      account,
      name: sanitize(cols[7]),
      ext: sanitize(cols[21]) || "N",
      date_text: dateText,
      date_value: parseReportDateToExcelSerial(dateText),
      origin: sanitize(cols[23]).toUpperCase(),
      seat: sanitize(cols[24]),
      reference: sanitize(cols[25]),
      detail: sanitize(cols[26]),
      debit: parseDecimalLike(cols[27]),
      credit: parseDecimalLike(cols[28]),
      balance: parseDecimalLike(cols[29]),
    });
  }

  return rows;
}

function main() {
  const args = parseArgs(process.argv);
  const inputPath = path.resolve(process.cwd(), args.input || "");
  const outputJson = path.resolve(process.cwd(), args["output-json"] || "");

  if (!inputPath || !fs.existsSync(inputPath)) {
    throw new Error(`No existe el mayor TXT: ${inputPath}`);
  }
  if (!outputJson) {
    throw new Error("Falta --output-json.");
  }

  const rows = parseMayorRows(inputPath);
  fs.mkdirSync(path.dirname(outputJson), { recursive: true });
  fs.writeFileSync(outputJson, `${JSON.stringify({ rows }, null, 2)}\n`, "utf8");
  console.log(`INFO|mayor_read|rows=${rows.length}|file=${path.basename(inputPath)}`);
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(error && error.message ? error.message : String(error));
    process.exit(1);
  }
}
