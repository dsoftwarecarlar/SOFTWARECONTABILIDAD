const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

function parseArgs(argv) {
  const args = {};
  for (let i = 2; i < argv.length; i += 1) {
    const key = String(argv[i] || "");
    if (!key.startsWith("-")) continue;
    const name = key.replace(/^-+/, "").toLowerCase();
    const next = argv[i + 1];
    if (next && !String(next).startsWith("-")) {
      args[name] = next;
      i += 1;
    } else {
      args[name] = "";
    }
  }
  return args;
}

function sanitize(text) {
  return String(text ?? "").trim();
}

function normalizeBrandKey(key) {
  switch (sanitize(key).toUpperCase()) {
    case "CHANGAN":
      return "changan";
    case "PEUGEOT":
      return "peug";
    case "SUZUKI":
      return "szk";
    case "TOYOTA":
    case "MATRIZ":
      return "tyt";
    default:
      return "";
  }
}

function readPxRows(sheet, brandKey) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  if (!Array.isArray(rows) || !rows.length) return [];

  let currentBrand = "";
  let capture = false;
  const collected = [];

  for (const row of rows) {
    const cells = Array.isArray(row) ? row : [];
    const marker = sanitize(cells[1]).toUpperCase();
    if (marker === "MARCA:") {
      currentBrand = normalizeBrandKey(cells[3]);
      capture = currentBrand === brandKey;
      continue;
    }

    if (!capture) continue;
    collected.push(cells);
  }

  return collected;
}

function main() {
  const args = parseArgs(process.argv);
  const inputPath = path.resolve(process.cwd(), args.input || "");
  const brandKey = sanitize(args.brand || "").toLowerCase();
  const outputJson = path.resolve(process.cwd(), args["output-json"] || "");

  if (!inputPath || !fs.existsSync(inputPath)) {
    throw new Error(`No existe PX: ${inputPath}`);
  }
  if (!brandKey) {
    throw new Error("Falta --brand para filtrar la seccion correspondiente.");
  }
  if (!outputJson) {
    throw new Error("Falta --output-json para exportar filas PX.");
  }

  const wb = XLSX.readFile(inputPath, { cellDates: false, cellNF: true, cellText: true });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) {
    throw new Error("El PX no tiene hojas.");
  }
  const sheet = wb.Sheets[sheetName];
  const filtered = readPxRows(sheet, brandKey);

  fs.mkdirSync(path.dirname(outputJson), { recursive: true });
  fs.writeFileSync(outputJson, `${JSON.stringify({ sheet_name: sheetName, rows: filtered }, null, 2)}\n`, "utf8");

  console.log(`INFO|px_read|rows=${filtered.length}|brand=${brandKey}`);
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(error && error.message ? error.message : String(error));
    process.exit(1);
  }
}
