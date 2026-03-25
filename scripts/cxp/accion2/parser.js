const fs = require("fs");

const XLSX = require("xlsx");

const {
  sanitizeText,
  round2,
  parseIntLike,
  parseDecimalLike,
} = require("../shared/core-utils");
const {
  HEADER_ALIASES,
  SHEET_NAME,
} = require("./constants");

function normalizeHeader(value) {
  return sanitizeText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9%]/g, "");
}

function parseDateFlexible(value) {
  if (value instanceof Date && Number.isFinite(value.getTime())) {
    return value;
  }

  const clean = sanitizeText(value);
  if (!clean) {
    return null;
  }

  let day;
  let month;
  let year;

  let match = /^(\d{2})[\/\-](\d{2})[\/\-](\d{4})$/.exec(clean);
  if (match) {
    day = Number(match[1]);
    month = Number(match[2]);
    year = Number(match[3]);
  } else {
    match = /^(\d{4})[\/\-](\d{2})[\/\-](\d{2})$/.exec(clean);
    if (!match) {
      return null;
    }
    year = Number(match[1]);
    month = Number(match[2]);
    day = Number(match[3]);
  }

  const date = new Date(year, month - 1, day);
  if (!Number.isFinite(date.getTime())) {
    return null;
  }
  return date;
}

function splitDelimitedLine(line, delimiter) {
  if (delimiter === "\t") {
    return line.split("\t").map((part) => part.trim());
  }

  const result = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    const next = line[i + 1];

    if (char === "\"") {
      if (inQuotes && next === "\"") {
        current += "\"";
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === delimiter && !inQuotes) {
      result.push(current.trim());
      current = "";
      continue;
    }

    current += char;
  }

  result.push(current.trim());
  return result;
}

function mapHeaders(tokens) {
  return tokens.map((token) => {
    const normalized = normalizeHeader(token);
    return HEADER_ALIASES[normalized] || null;
  });
}

function detectHeaderConfig(lines) {
  const delimiters = ["\t", ";", "|", ","];
  let best = null;

  const maxLines = Math.min(lines.length, 60);
  for (let i = 0; i < maxLines; i += 1) {
    const line = sanitizeText(lines[i]);
    if (!line) {
      continue;
    }

    for (const delimiter of delimiters) {
      const tokens = splitDelimitedLine(lines[i], delimiter);
      if (tokens.length < 5) {
        continue;
      }
      const mapped = mapHeaders(tokens);
      const recognized = mapped.filter(Boolean).length;
      if (recognized < 5) {
        continue;
      }

      const score = recognized * 100 + tokens.length;
      if (!best || score > best.score) {
        best = {
          score,
          headerIndex: i,
          delimiter,
          mappedHeaders: mapped,
          tokenCount: tokens.length,
        };
      }
    }
  }

  return best;
}

function parseRowsFromHeader(lines, config) {
  const rows = [];
  for (let i = config.headerIndex + 1; i < lines.length; i += 1) {
    const rawLine = lines[i];
    if (!sanitizeText(rawLine)) {
      continue;
    }

    const tokens = splitDelimitedLine(rawLine, config.delimiter);
    if (tokens.every((token) => !sanitizeText(token))) {
      continue;
    }

    const row = {};
    for (let c = 0; c < config.mappedHeaders.length; c += 1) {
      const header = config.mappedHeaders[c];
      if (!header) {
        continue;
      }
      row[header] = tokens[c] || "";
    }

    const nonEmptyMapped = Object.values(row).some((value) => sanitizeText(value));
    if (nonEmptyMapped) {
      rows.push(row);
    }
  }

  return rows;
}

function parseRowsByPattern(lines) {
  const rows = [];
  const hardPattern = /^(\d+)\s+(.+?)\s+(\d{2}[\/\-]\d{2}[\/\-]\d{4})\s+(\d{2}[\/\-]\d{2}[\/\-]\d{4})\s+(IVA|RENTA)\s+([A-Z0-9]+)\s+([A-Z0-9\-\/]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$/i;

  for (const line of lines) {
    const cleanLine = line.trim();
    if (!cleanLine) {
      continue;
    }

    const match = hardPattern.exec(cleanLine);
    if (match) {
      rows.push({
        "NUM RT": match[1],
        PROVEEDOR: match[2],
        FECHA: match[3],
        "FECHA CONT": match[4],
        TIPO: match[5],
        COD: match[6],
        FACT: match[7],
        "%": match[8],
        BASE: match[9],
        RETENCION: match[10],
      });
      continue;
    }

    const pieces = cleanLine.split(/\s{2,}/).map((part) => part.trim()).filter(Boolean);
    if (pieces.length >= 10) {
      rows.push({
        "NUM RT": pieces[0],
        PROVEEDOR: pieces[1],
        FECHA: pieces[2],
        "FECHA CONT": pieces[3],
        TIPO: pieces[4],
        COD: pieces[5],
        FACT: pieces[6],
        "%": pieces[7],
        BASE: pieces[8],
        RETENCION: pieces[9],
      });
    }
  }

  return rows;
}

function parseRowsFromEmbeddedRetReport(lines) {
  const rows = [];
  let matched = 0;

  for (const line of lines) {
    if (!line || line.indexOf("\t") === -1) {
      continue;
    }

    const tokens = line
      .split("\t")
      .map((token) => sanitizeText(token));

    if (tokens.length < 20) {
      continue;
    }
    if (normalizeHeader(tokens[1]) !== "RETENCION") {
      continue;
    }
    if (normalizeHeader(tokens[2]) !== "FECHA") {
      continue;
    }
    if (normalizeHeader(tokens[3]) !== "FECHADOC") {
      continue;
    }

    const numRt = tokens[11];
    const fecha = tokens[12];
    const fechaCont = tokens[13];
    const cod = tokens[14];
    const fact = tokens[15];
    const percent = tokens[16];
    const base = tokens[17];
    const ret = tokens[18];
    const proveedor = tokens[19];

    if (!/^\d+$/.test(numRt)) {
      continue;
    }
    if (!/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/.test(fecha)) {
      continue;
    }
    if (!/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/.test(fechaCont)) {
      continue;
    }
    if (!sanitizeText(proveedor)) {
      continue;
    }

    matched += 1;
    rows.push({
      "NUM RT": numRt,
      PROVEEDOR: proveedor,
      FECHA: fecha,
      "FECHA CONT": fechaCont,
      COD: cod,
      FACT: fact,
      "%": percent,
      BASE: base,
      RETENCION: ret,
      TIPO: "",
    });
  }

  const minExpected = Math.max(20, Math.floor(lines.length * 0.6));
  if (matched >= minExpected) {
    return rows;
  }

  return [];
}

function loadBestText(buffer) {
  const utf8 = buffer.toString("utf8").replace(/^\uFEFF/, "");
  const latin1 = buffer.toString("latin1").replace(/^\uFEFF/, "");
  const utf8Bad = (utf8.match(/ï¿½/g) || []).length;
  const latin1Bad = (latin1.match(/ï¿½/g) || []).length;
  return utf8Bad <= latin1Bad ? utf8 : latin1;
}

function normalizeTipo(value) {
  const clean = sanitizeText(value).toUpperCase();
  if (clean === "IVA") {
    return "IVA";
  }
  if (clean === "RENTA") {
    return "RENTA";
  }
  return clean;
}

function formatDateKey(value) {
  const date = parseDateFlexible(value);
  if (!date) {
    return "";
  }
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = date.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function normalizeFactKey(value) {
  const clean = sanitizeText(value);
  if (/^\d+$/.test(clean)) {
    return String(Number(clean));
  }
  return clean.toUpperCase();
}

function normalizeProviderKey(value) {
  return sanitizeText(value).toUpperCase();
}

function buildTipoMatchKey(fields) {
  return [
    fields.numRt == null ? "" : String(fields.numRt),
    formatDateKey(fields.fecha),
    formatDateKey(fields.fechaCont || fields.fecha),
    fields.cod == null ? "" : String(fields.cod),
    normalizeFactKey(fields.fact),
    round2(fields.percent).toFixed(4),
    round2(fields.base).toFixed(2),
    round2(fields.retencion).toFixed(2),
    normalizeProviderKey(fields.proveedor),
  ].join("|");
}

function getEmptyTipoHints() {
  return {
    exact: new Map(),
    byCodePercent: new Map(),
  };
}

function loadTemplateTipoHints(templatePath) {
  if (!fs.existsSync(templatePath)) {
    return getEmptyTipoHints();
  }

  const wb = XLSX.readFile(templatePath, { cellDates: true });
  const ws = wb.Sheets[SHEET_NAME];
  if (!ws || !ws["!ref"]) {
    return getEmptyTipoHints();
  }

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
  const hints = getEmptyTipoHints();

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const tipo = normalizeTipo(row[4]);
    if (!["IVA", "RENTA"].includes(tipo)) {
      continue;
    }

    const fields = {
      numRt: parseIntLike(row[0]),
      proveedor: row[1],
      fecha: row[2],
      fechaCont: row[3],
      cod: parseIntLike(row[5]),
      fact: row[6],
      percent: parseDecimalLike(row[7]),
      base: parseDecimalLike(row[8]),
      retencion: parseDecimalLike(row[9]),
    };

    if (fields.numRt == null || fields.cod == null) {
      continue;
    }

    const key = buildTipoMatchKey(fields);
    hints.exact.set(key, tipo);

    const cpKey = `${fields.cod}|${round2(fields.percent).toFixed(4)}`;
    if (!hints.byCodePercent.has(cpKey)) {
      hints.byCodePercent.set(cpKey, { IVA: 0, RENTA: 0 });
    }
    hints.byCodePercent.get(cpKey)[tipo] += 1;
  }

  return hints;
}

function inferTipoFromHints(fields, hints) {
  const exactKey = buildTipoMatchKey(fields);
  if (hints.exact.has(exactKey)) {
    return hints.exact.get(exactKey);
  }

  const rate = round2(fields.percent);
  const cpKey = `${fields.cod}|${rate.toFixed(4)}`;
  const cpHint = hints.byCodePercent.get(cpKey);
  if (cpHint) {
    if (cpHint.IVA > cpHint.RENTA) {
      return "IVA";
    }
    if (cpHint.RENTA > cpHint.IVA) {
      return "RENTA";
    }
  }

  if (rate === 0) {
    return "RENTA";
  }
  if (new Set([20, 30, 70, 100]).has(rate)) {
    return "IVA";
  }
  if (new Set([1, 1.75, 2, 2.75, 3, 5, 8]).has(rate)) {
    return "RENTA";
  }

  if (rate >= 20) {
    return "IVA";
  }
  return "RENTA";
}

function normalizeParsedRows(rawRows, hints) {
  const rows = [];
  const problems = [];

  for (let i = 0; i < rawRows.length; i += 1) {
    const source = rawRows[i];
    const rowNo = i + 1;

    const numRt = parseIntLike(source["NUM RT"]);
    const proveedor = sanitizeText(source.PROVEEDOR);
    const fecha = parseDateFlexible(source.FECHA);
    const fechaCont = parseDateFlexible(source["FECHA CONT"]) || fecha;
    const cod = parseIntLike(source.COD);
    const factRaw = sanitizeText(source.FACT);
    const factNum = /^\d+$/.test(factRaw) ? Number(factRaw) : factRaw;
    const percent = parseDecimalLike(source["%"]);
    const base = parseDecimalLike(source.BASE);
    const retencion = parseDecimalLike(source.RETENCION);
    let tipo = normalizeTipo(source.TIPO);
    if (!["IVA", "RENTA"].includes(tipo)) {
      tipo = inferTipoFromHints(
        {
          numRt,
          proveedor,
          fecha,
          fechaCont,
          cod,
          fact: factNum,
          percent,
          base,
          retencion,
        },
        hints,
      );
    }

    if (numRt == null) {
      problems.push(`Fila ${rowNo}: NUM RT invalido.`);
    }
    if (!proveedor) {
      problems.push(`Fila ${rowNo}: PROVEEDOR vacio.`);
    }
    if (!fecha) {
      problems.push(`Fila ${rowNo}: FECHA invalida.`);
    }
    if (!fechaCont) {
      problems.push(`Fila ${rowNo}: FECHA CONT invalida.`);
    }
    if (!["IVA", "RENTA"].includes(tipo)) {
      problems.push(`Fila ${rowNo}: TIPO invalido (${tipo || "vacio"}).`);
    }
    if (cod == null) {
      problems.push(`Fila ${rowNo}: COD invalido.`);
    }
    if (!factRaw) {
      problems.push(`Fila ${rowNo}: FACT vacio.`);
    }

    rows.push({
      numRt,
      proveedor,
      fecha,
      fechaCont,
      tipo,
      cod,
      fact: factNum,
      percent,
      base,
      retencion,
    });
  }

  if (problems.length > 0) {
    const preview = problems.slice(0, 8).join(" | ");
    throw new Error(`Validacion de TXT fallida (${problems.length} problemas). ${preview}`);
  }

  return rows;
}

function extractRowsFromTxt(inputPath) {
  const buffer = fs.readFileSync(inputPath);
  const text = loadBestText(buffer);
  const lines = text
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .map((line) => line.replace(/\u00A0/g, " "));

  let rawRows = parseRowsFromEmbeddedRetReport(lines);

  if (rawRows.length === 0) {
    const headerConfig = detectHeaderConfig(lines);
    if (headerConfig) {
      rawRows = parseRowsFromHeader(lines, headerConfig);
    }
  }

  if (rawRows.length === 0) {
    rawRows = parseRowsByPattern(lines);
  }

  if (rawRows.length === 0) {
    throw new Error("No se detectaron filas validas en el TXT. Verifica encabezados o delimitadores.");
  }

  return rawRows;
}

module.exports = {
  extractRowsFromTxt,
  loadTemplateTipoHints,
  normalizeParsedRows,
};
