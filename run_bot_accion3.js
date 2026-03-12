const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");
const ExcelJS = require("exceljs");
const { XMLBuilder, XMLParser } = require("fast-xml-parser");
const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.mjs");
const XLSX = require("xlsx");

const DEFAULT_INPUT_SOURCE = "CON_MAYORGEN2ACCION3.txt";
const DEFAULT_OUTPUT_XLSX = "mayor_ret_accion3.xlsx";
const DEFAULT_TEMPLATE_XLSX = path.join(__dirname, "outputs", "EJEMPLOSAMANO", "MAYOR RET_ACCION3.xlsx");

const SHEET_NAME = "MAYOR RET";
const EXPECTED_HEADERS = [
  "COD",
  "CUENTA",
  "EXT",
  "FECHA",
  "ORIGEN",
  "ASIENTO",
  "DOCU",
  "DETALLE",
  "DEBE",
  "HABER",
  "SALDO",
];

const INVALID_XML_CONTROL_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g;

const MOVEMENT_BOUNDARIES = [
  { name: "FECHA", left: Number.NEGATIVE_INFINITY, right: 56 },
  { name: "ORIGEN", left: 56, right: 86 },
  { name: "ASIENTO", left: 86, right: 118 },
  { name: "EXT", left: 118, right: 146 },
  { name: "DOCU", left: 146, right: 208 },
  { name: "DETALLE", left: 208, right: 372 },
  { name: "DEBE", left: 372, right: 449 },
  { name: "HABER", left: 449, right: 523 },
  { name: "SALDO", left: 523, right: Number.POSITIVE_INFINITY },
];

const XML_PARSER = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  allowBooleanAttributes: true,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
});

const XML_BUILDER = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  suppressBooleanAttributes: false,
  format: false,
});

function sanitizeText(value) {
  return String(value || "")
    .replace(INVALID_XML_CONTROL_CHARS, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function deepClone(value) {
  return value ? JSON.parse(JSON.stringify(value)) : {};
}

function round2(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

function toExcelDateSerial(value) {
  const date = value instanceof Date ? value : null;
  if (!date || !Number.isFinite(date.getTime())) {
    return null;
  }
  const utcMillis = Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
  return Math.floor(utcMillis / 86400000 + 25569);
}

function parseIntLike(value) {
  const normalized = sanitizeText(value).replace(/[^\d]/g, "");
  if (!normalized) {
    return null;
  }
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function parseDecimalLike(value) {
  let normalized = sanitizeText(value).replace(/[^\d,.\-]/g, "");
  if (!normalized) {
    return 0;
  }

  const hasComma = normalized.includes(",");
  const hasDot = normalized.includes(".");

  if (hasComma && hasDot) {
    if (normalized.lastIndexOf(".") > normalized.lastIndexOf(",")) {
      normalized = normalized.replace(/,/g, "");
    } else {
      normalized = normalized.replace(/\./g, "").replace(/,/g, ".");
    }
  } else if (hasComma) {
    const commaCount = (normalized.match(/,/g) || []).length;
    if (commaCount > 1) {
      normalized = normalized.replace(/,/g, "");
    } else {
      const [intPart, fracPart = ""] = normalized.split(",");
      if (fracPart.length === 3) {
        normalized = `${intPart}${fracPart}`;
      } else {
        normalized = `${intPart}.${fracPart}`;
      }
    }
  } else if (hasDot) {
    const dotCount = (normalized.match(/\./g) || []).length;
    if (dotCount > 1) {
      const lastDot = normalized.lastIndexOf(".");
      const intPart = normalized.slice(0, lastDot).replace(/\./g, "");
      const fracPart = normalized.slice(lastDot + 1);
      if (fracPart.length === 3) {
        normalized = `${intPart}${fracPart}`;
      } else {
        normalized = `${intPart}.${fracPart}`;
      }
    }
  }

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeReference(value) {
  const clean = sanitizeText(value).replace(/\s+/g, "");
  return clean;
}

function parseDateFromReport(value) {
  const clean = sanitizeText(value).toUpperCase();
  const match = /^(\d{2})-([A-Z]{3})-(\d{2})$/.exec(clean);
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

  if (!Number.isInteger(day) || month == null || !Number.isInteger(year)) {
    return null;
  }

  // Use UTC midnight to avoid timezone offsets shifting Excel serial dates.
  const date = new Date(Date.UTC(year, month, day));
  if (
    !Number.isFinite(date.getTime()) ||
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month ||
    date.getUTCDate() !== day
  ) {
    return null;
  }
  return date;
}

function parseAccountHeaderTotalsFromPdfRow(rowItems) {
  const rightSideTokens = rowItems
    .filter((item) => Number(item.x) >= 300)
    .map((item) => sanitizeText(item.str))
    .filter((text) => /^-?\d[\d.,]*$/.test(text) && /[.,]\d{2}$/.test(text));

  if (rightSideTokens.length < 4) {
    return null;
  }

  const last4 = rightSideTokens.slice(-4).map((token) => round2(parseDecimalLike(token)));
  return {
    saldo_inicial: last4[0],
    total_debe: last4[1],
    total_haber: last4[2],
    saldo_final: last4[3],
  };
}

function assignToBoundary(x, boundaries) {
  for (const boundary of boundaries) {
    if (x >= boundary.left && x < boundary.right) {
      return boundary.name;
    }
  }
  return null;
}

function groupItemsByRow(items, tolerance = 0.9) {
  const rows = [];
  const sorted = [...items].sort((a, b) => b.y - a.y || a.x - b.x);

  for (const item of sorted) {
    let row = null;
    for (const candidate of rows) {
      if (Math.abs(candidate.y - item.y) <= tolerance) {
        row = candidate;
        break;
      }
    }
    if (!row) {
      row = { y: item.y, items: [] };
      rows.push(row);
    }
    row.items.push(item);
  }

  for (const row of rows) {
    row.items.sort((a, b) => a.x - b.x);
  }
  rows.sort((a, b) => b.y - a.y);
  return rows;
}

function extractFieldsFromRow(rowItems) {
  const buckets = {};
  for (const boundary of MOVEMENT_BOUNDARIES) {
    buckets[boundary.name] = [];
  }

  for (const item of rowItems) {
    const name = assignToBoundary(item.x, MOVEMENT_BOUNDARIES);
    if (!name) {
      continue;
    }
    const text = sanitizeText(item.str);
    if (!text) {
      continue;
    }
    buckets[name].push(text);
  }

  const result = {};
  for (const boundary of MOVEMENT_BOUNDARIES) {
    const name = boundary.name;
    const pieces = buckets[name] || [];
    if (["DEBE", "HABER", "SALDO"].includes(name)) {
      result[name] = pieces.join("");
    } else {
      result[name] = sanitizeText(pieces.join(" "));
    }
  }

  return result;
}

function detectAccountHeader(rowItems) {
  const fullText = sanitizeText(rowItems.map((item) => item.str).join(" "));
  const codeMatch = /\b\d{2}\.\d{2}\.\d{2}\.\d{2}\.\d{4}\b/.exec(fullText);
  if (!codeMatch) {
    return null;
  }

  const nameParts = rowItems
    .filter((item) => item.x >= 95 && item.x < 315)
    .map((item) => sanitizeText(item.str))
    .filter(Boolean);

  const name = sanitizeText(nameParts.join(" ")).replace(/\s+0[.,]00$/g, "").trim();
  if (!name || /^CUENTA$/i.test(name)) {
    return null;
  }

  return {
    code: codeMatch[0],
    name,
    totals: parseAccountHeaderTotalsFromPdfRow(rowItems),
  };
}

function validateAccountTotals(rows, accountHeaderTotals) {
  if (accountHeaderTotals.size === 0) {
    return { checked_accounts: 0, mismatches: [] };
  }

  const actualByCode = new Map();
  for (const row of rows) {
    if (!actualByCode.has(row.COD)) {
      actualByCode.set(row.COD, {
        total_debe: 0,
        total_haber: 0,
        saldo_final: 0,
      });
    }
    const item = actualByCode.get(row.COD);
    item.total_debe += round2(row.DEBE);
    item.total_haber += round2(row.HABER);
    item.saldo_final = round2(row.SALDO);
  }

  const mismatches = [];
  for (const [code, expected] of accountHeaderTotals.entries()) {
    const actual = actualByCode.get(code) || {
      total_debe: 0,
      total_haber: 0,
      saldo_final: 0,
    };
    const diffDebe = round2(round2(actual.total_debe) - round2(expected.total_debe));
    const diffHaber = round2(round2(actual.total_haber) - round2(expected.total_haber));
    const diffSaldo = round2(round2(actual.saldo_final) - round2(expected.saldo_final));

    if (Math.abs(diffDebe) > 0.05 || Math.abs(diffHaber) > 0.05 || Math.abs(diffSaldo) > 0.05) {
      mismatches.push({
        code,
        name: expected.name,
        expected: {
          total_debe: round2(expected.total_debe),
          total_haber: round2(expected.total_haber),
          saldo_final: round2(expected.saldo_final),
        },
        actual: {
          total_debe: round2(actual.total_debe),
          total_haber: round2(actual.total_haber),
          saldo_final: round2(actual.saldo_final),
        },
        diff: {
          debe: diffDebe,
          haber: diffHaber,
          saldo: diffSaldo,
        },
      });
    }
  }

  return {
    checked_accounts: accountHeaderTotals.size,
    mismatches,
  };
}

function isMovementRow(fields) {
  const fecha = sanitizeText(fields.FECHA).toUpperCase();
  const origen = sanitizeText(fields.ORIGEN).toUpperCase();
  const saldoText = sanitizeText(fields.SALDO);

  if (!/^\d{2}-[A-Z]{3}-\d{2}$/.test(fecha)) {
    return false;
  }
  if (!/^[A-Z]{2,6}$/.test(origen)) {
    return false;
  }
  if (!/[0-9]/.test(saldoText)) {
    return false;
  }

  return true;
}

async function parseTxtRows(inputTxtPath) {
  const rawText = fs.readFileSync(inputTxtPath, "utf8");
  const lines = rawText
    .split(/\r?\n/)
    .map((line) => line.replace(/\uFEFF/g, "").trimEnd())
    .filter((line) => sanitizeText(line) !== "");

  const rows = [];
  const accountHeaderTotals = new Map();
  const skippedDateRows = [];
  let dateRowsDetected = 0;

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    const cols = line.split("\t");

    // The first metadata line does not contain movement payload.
    if (cols.length < 30) {
      continue;
    }

    const dateToken = sanitizeText(cols[22]).toUpperCase();
    const hasDateToken = /^\d{2}-[A-Z]{3}-\d{2}$/.test(dateToken);
    if (hasDateToken) {
      dateRowsDetected += 1;
    }

    const code = sanitizeText(cols[6]);
    const accountName = sanitizeText(cols[7]);
    const fecha = parseDateFromReport(dateToken);
    const origen = sanitizeText(cols[23]).toUpperCase();
    const asiento = parseIntLike(cols[24]) ?? 0;
    const ext = sanitizeText(cols[21]).toUpperCase() || "N";
    const docu = normalizeReference(cols[25]);
    const detalle = sanitizeText(cols[26]);
    const debe = round2(parseDecimalLike(cols[27]));
    const haber = round2(parseDecimalLike(cols[28]));
    const saldo = round2(parseDecimalLike(cols[29]));

    if (!fecha) {
      if (hasDateToken) {
        skippedDateRows.push({
          line: index + 1,
          reason: "invalid_date_value",
          text: sanitizeText(line),
        });
      }
      continue;
    }

    if (!/^\d{2}\.\d{2}\.\d{2}\.\d{2}\.\d{4}$/.test(code)) {
      skippedDateRows.push({
        line: index + 1,
        reason: "invalid_account_code",
        text: sanitizeText(line),
      });
      continue;
    }
    if (!accountName) {
      skippedDateRows.push({
        line: index + 1,
        reason: "missing_account_name",
        text: sanitizeText(line),
      });
      continue;
    }
    if (!/^[A-Z]{2,6}$/.test(origen)) {
      skippedDateRows.push({
        line: index + 1,
        reason: "invalid_origin",
        text: sanitizeText(line),
      });
      continue;
    }

    accountHeaderTotals.set(code, {
      name: accountName,
      saldo_inicial: round2(parseDecimalLike(cols[8])),
      total_debe: round2(parseDecimalLike(cols[9])),
      total_haber: round2(parseDecimalLike(cols[10])),
      saldo_final: round2(parseDecimalLike(cols[11])),
    });

    rows.push({
      COD: code,
      CUENTA: accountName,
      EXT: ext,
      FECHA: fecha,
      ORIGEN: origen,
      ASIENTO: asiento,
      DOCU: docu,
      DETALLE: detalle,
      DEBE: debe,
      HABER: haber,
      SALDO: saldo,
    });
  }

  if (rows.length === 0) {
    throw new Error("No se detectaron movimientos validos en el TXT de Accion 3.");
  }

  if (skippedDateRows.length > 0) {
    const preview = skippedDateRows
      .slice(0, 5)
      .map((item) => `linea ${item.line} (${item.reason})`)
      .join(" | ");
    throw new Error(
      `Se detectaron ${skippedDateRows.length} filas con fecha que no pudieron mapearse en TXT Accion 3. ${preview}`,
    );
  }

  const totalsValidation = validateAccountTotals(rows, accountHeaderTotals);
  if (totalsValidation.mismatches.length > 0) {
    const preview = totalsValidation.mismatches
      .slice(0, 4)
      .map(
        (item) =>
          `${item.code} (${item.name}): ` +
          `debe TXT=${item.expected.total_debe} / Excel=${item.actual.total_debe}, ` +
          `haber TXT=${item.expected.total_haber} / Excel=${item.actual.total_haber}, ` +
          `saldo TXT=${item.expected.saldo_final} / Excel=${item.actual.saldo_final}`,
      )
      .join(" | ");
    throw new Error(
      `Validacion contable fallida en TXT Accion 3 (${totalsValidation.mismatches.length} cuentas). ${preview}`,
    );
  }

  return {
    rows,
    diagnostics: {
      source_type: "txt",
      lines_total: lines.length,
      date_rows_detected: dateRowsDetected,
      movement_rows_extracted: rows.length,
      skipped_date_rows: skippedDateRows.length,
      account_totals_checked: totalsValidation.checked_accounts,
      account_total_mismatches: totalsValidation.mismatches.length,
    },
  };
}

async function parsePdfRows(inputPdfPath) {
  const data = new Uint8Array(fs.readFileSync(inputPdfPath));
  const loadingTask = pdfjsLib.getDocument({
    data,
    disableWorker: true,
    verbosity: pdfjsLib.VerbosityLevel?.ERRORS,
  });
  const pdf = await loadingTask.promise;

  const rows = [];
  const accountHeaderTotals = new Map();
  const skippedDateRows = [];
  let dateRowsDetected = 0;
  let currentCode = "";
  let currentName = "";

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const textContent = await page.getTextContent();
    const items = textContent.items.map((item) => ({
      str: sanitizeText(item.str),
      x: item.transform[4],
      y: item.transform[5],
    }));

    const groupedRows = groupItemsByRow(items);
    for (const row of groupedRows) {
      const rowText = sanitizeText(row.items.map((item) => item.str).join(" "));
      const hasDateToken = /\b\d{2}-[A-Z]{3}-\d{2}\b/.test(rowText.toUpperCase());
      if (hasDateToken) {
        dateRowsDetected += 1;
      }

      const accountHeader = detectAccountHeader(row.items);
      if (accountHeader) {
        currentCode = accountHeader.code;
        currentName = accountHeader.name;
        if (accountHeader.totals) {
          accountHeaderTotals.set(accountHeader.code, {
            ...accountHeader.totals,
            name: accountHeader.name,
          });
        }
        continue;
      }

      const fields = extractFieldsFromRow(row.items);
      if (!isMovementRow(fields)) {
        if (hasDateToken && !/^PROCESADO:/i.test(rowText)) {
          skippedDateRows.push({
            page: pageNumber,
            y: row.y,
            reason: "not_movement_row",
            text: rowText,
          });
        }
        continue;
      }
      if (!currentCode || !currentName) {
        skippedDateRows.push({
          page: pageNumber,
          y: row.y,
          reason: "missing_account_context",
          text: rowText,
        });
        continue;
      }

      const fecha = parseDateFromReport(fields.FECHA);
      if (!fecha) {
        skippedDateRows.push({
          page: pageNumber,
          y: row.y,
          reason: "invalid_date_value",
          text: rowText,
        });
        continue;
      }

      rows.push({
        COD: currentCode,
        CUENTA: currentName,
        EXT: sanitizeText(fields.EXT).toUpperCase() || "N",
        FECHA: fecha,
        ORIGEN: sanitizeText(fields.ORIGEN).toUpperCase(),
        ASIENTO: parseIntLike(fields.ASIENTO) ?? 0,
        DOCU: normalizeReference(fields.DOCU),
        DETALLE: sanitizeText(fields.DETALLE),
        DEBE: round2(parseDecimalLike(fields.DEBE)),
        HABER: round2(parseDecimalLike(fields.HABER)),
        SALDO: round2(parseDecimalLike(fields.SALDO)),
      });
    }
  }

  if (rows.length === 0) {
    throw new Error("No se detectaron movimientos validos en el PDF de Accion 3.");
  }

  if (skippedDateRows.length > 0) {
    const preview = skippedDateRows
      .slice(0, 5)
      .map((item) => `pag ${item.page} (${item.reason}): ${item.text}`)
      .join(" | ");
    throw new Error(
      `Se detectaron ${skippedDateRows.length} filas con fecha que no pudieron mapearse en Accion 3. ${preview}`,
    );
  }

  const totalsValidation = validateAccountTotals(rows, accountHeaderTotals);
  if (totalsValidation.mismatches.length > 0) {
    const preview = totalsValidation.mismatches
      .slice(0, 4)
      .map(
        (item) =>
          `${item.code} (${item.name}): ` +
          `debe PDF=${item.expected.total_debe} / Excel=${item.actual.total_debe}, ` +
          `haber PDF=${item.expected.total_haber} / Excel=${item.actual.total_haber}, ` +
          `saldo PDF=${item.expected.saldo_final} / Excel=${item.actual.saldo_final}`,
      )
      .join(" | ");
    throw new Error(
      `Validacion contable fallida en Accion 3 (${totalsValidation.mismatches.length} cuentas). ${preview}`,
    );
  }

  return {
    rows,
    diagnostics: {
      source_type: "pdf",
      pages: pdf.numPages,
      date_rows_detected: dateRowsDetected,
      movement_rows_extracted: rows.length,
      skipped_date_rows: skippedDateRows.length,
      account_totals_checked: totalsValidation.checked_accounts,
      account_total_mismatches: totalsValidation.mismatches.length,
    },
  };
}

function validateRows(rows) {
  const problems = [];
  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    const rowNo = i + 1;
    if (!sanitizeText(row.COD)) {
      problems.push(`Fila ${rowNo}: COD vacio.`);
    }
    if (!sanitizeText(row.CUENTA)) {
      problems.push(`Fila ${rowNo}: CUENTA vacia.`);
    }
    if (!(row.FECHA instanceof Date) || !Number.isFinite(row.FECHA.getTime())) {
      problems.push(`Fila ${rowNo}: FECHA invalida.`);
    }
    if (!sanitizeText(row.ORIGEN)) {
      problems.push(`Fila ${rowNo}: ORIGEN vacio.`);
    }
    if (!sanitizeText(row.DETALLE)) {
      problems.push(`Fila ${rowNo}: DETALLE vacio.`);
    }
  }

  if (problems.length > 0) {
    const preview = problems.slice(0, 8).join(" | ");
    throw new Error(`Validacion de filas fallida (${problems.length} problemas). ${preview}`);
  }
}

function compareTextNatural(left, right) {
  return sanitizeText(left).localeCompare(sanitizeText(right), "es", {
    numeric: true,
    sensitivity: "base",
  });
}

function compareMovementRows(left, right) {
  const byCode = compareTextNatural(left.COD, right.COD);
  if (byCode !== 0) {
    return byCode;
  }

  const byDate = (left.FECHA?.getTime() || 0) - (right.FECHA?.getTime() || 0);
  if (byDate !== 0) {
    return byDate;
  }

  const byAsiento = (Number(left.ASIENTO) || 0) - (Number(right.ASIENTO) || 0);
  if (byAsiento !== 0) {
    return byAsiento;
  }

  const byOrigin = compareTextNatural(left.ORIGEN, right.ORIGEN);
  if (byOrigin !== 0) {
    return byOrigin;
  }

  const byDocu = compareTextNatural(left.DOCU, right.DOCU);
  if (byDocu !== 0) {
    return byDocu;
  }

  const byDetail = compareTextNatural(left.DETALLE, right.DETALLE);
  if (byDetail !== 0) {
    return byDetail;
  }

  const byDebe = round2(left.DEBE) - round2(right.DEBE);
  if (byDebe !== 0) {
    return byDebe;
  }

  const byHaber = round2(left.HABER) - round2(right.HABER);
  if (byHaber !== 0) {
    return byHaber;
  }

  const bySaldo = round2(left.SALDO) - round2(right.SALDO);
  if (bySaldo !== 0) {
    return bySaldo;
  }

  const bySource = (left.__sourceIndex || 0) - (right.__sourceIndex || 0);
  if (bySource !== 0) {
    return bySource;
  }

  return (left.__rowIndex || 0) - (right.__rowIndex || 0);
}

function sortRowsForWorkbook(rows) {
  return [...rows].sort(compareMovementRows);
}

function stripInternalSortMetadata(row) {
  const { __sourceIndex, __rowIndex, ...cleanRow } = row;
  return cleanRow;
}

function ensureTemplateCapacity(ws, requiredRowCount) {
  const currentRowCount = Math.max(ws.rowCount, 1);
  if (requiredRowCount <= currentRowCount) {
    return;
  }

  const templateDataRow = Math.max(currentRowCount, 2);
  ws.duplicateRow(templateDataRow, requiredRowCount - currentRowCount, true);
}

async function parseInputSources(inputSources) {
  const resolvedInputs = inputSources.map((input) => path.resolve(process.cwd(), input));
  if (resolvedInputs.length === 0) {
    throw new Error("No se recibieron archivos de entrada para Accion 3.");
  }

  for (const inputPath of resolvedInputs) {
    if (!fs.existsSync(inputPath)) {
      throw new Error(`No se encontro el archivo de entrada: ${inputPath}`);
    }
  }

  const extensions = [...new Set(resolvedInputs.map((inputPath) => path.extname(inputPath).toLowerCase()))];
  if (extensions.length !== 1) {
    throw new Error("No se pueden mezclar TXT y PDF en una sola ejecucion de Accion 3.");
  }

  const inputExt = extensions[0];
  if (inputExt === ".pdf") {
    if (resolvedInputs.length > 1) {
      throw new Error("La consolidacion multiple de Accion 3 solo admite archivos TXT.");
    }
    const parsed = await parsePdfRows(resolvedInputs[0]);
    return {
      ...parsed,
      inputExt,
      inputPaths: resolvedInputs,
      sourceFiles: [
        {
          input_source: resolvedInputs[0],
          file_name: path.basename(resolvedInputs[0]),
          movimientos_extraidos: parsed.rows.length,
          ...parsed.diagnostics,
        },
      ],
    };
  }

  if (inputExt !== ".txt") {
    throw new Error("Solo se permiten archivos TXT o PDF para Accion 3.");
  }

  if (resolvedInputs.length === 1) {
    const parsed = await parseTxtRows(resolvedInputs[0]);
    return {
      ...parsed,
      inputExt,
      inputPaths: resolvedInputs,
      sourceFiles: [
        {
          input_source: resolvedInputs[0],
          file_name: path.basename(resolvedInputs[0]),
          movimientos_extraidos: parsed.rows.length,
          ...parsed.diagnostics,
        },
      ],
    };
  }

  const diagnostics = {
    source_type: "txt_multi",
    files_total: resolvedInputs.length,
    lines_total: 0,
    date_rows_detected: 0,
    movement_rows_extracted: 0,
    skipped_date_rows: 0,
    account_totals_checked: 0,
    account_total_mismatches: 0,
  };
  const rows = [];
  const sourceFiles = [];

  for (let sourceIndex = 0; sourceIndex < resolvedInputs.length; sourceIndex += 1) {
    const inputPath = resolvedInputs[sourceIndex];
    const parsed = await parseTxtRows(inputPath);

    diagnostics.lines_total += parsed.diagnostics.lines_total || 0;
    diagnostics.date_rows_detected += parsed.diagnostics.date_rows_detected || 0;
    diagnostics.movement_rows_extracted += parsed.rows.length;
    diagnostics.skipped_date_rows += parsed.diagnostics.skipped_date_rows || 0;
    diagnostics.account_totals_checked += parsed.diagnostics.account_totals_checked || 0;
    diagnostics.account_total_mismatches += parsed.diagnostics.account_total_mismatches || 0;

    sourceFiles.push({
      input_source: inputPath,
      file_name: path.basename(inputPath),
      movimientos_extraidos: parsed.rows.length,
      ...parsed.diagnostics,
    });

    for (let rowIndex = 0; rowIndex < parsed.rows.length; rowIndex += 1) {
      rows.push({
        ...parsed.rows[rowIndex],
        __sourceIndex: sourceIndex,
        __rowIndex: rowIndex,
      });
    }
  }

  return {
    rows: sortRowsForWorkbook(rows).map(stripInternalSortMetadata),
    diagnostics,
    inputExt,
    inputPaths: resolvedInputs,
    sourceFiles,
  };
}

function parseCliArguments(argv = process.argv.slice(2)) {
  const hasFlags = argv.includes("--output") || argv.includes("--template");
  if (!hasFlags) {
    return {
      inputSources: [argv[0] || DEFAULT_INPUT_SOURCE],
      outputXlsx: argv[1] || DEFAULT_OUTPUT_XLSX,
      templateXlsx: argv[2] ? path.resolve(process.cwd(), argv[2]) : DEFAULT_TEMPLATE_XLSX,
    };
  }

  const inputSources = [];
  let outputXlsx = DEFAULT_OUTPUT_XLSX;
  let templateXlsx = DEFAULT_TEMPLATE_XLSX;

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--output") {
      if (!argv[index + 1]) {
        throw new Error("Falta el valor de --output.");
      }
      outputXlsx = argv[index + 1];
      index += 1;
      continue;
    }

    if (arg === "--template") {
      if (!argv[index + 1]) {
        throw new Error("Falta el valor de --template.");
      }
      templateXlsx = path.resolve(process.cwd(), argv[index + 1]);
      index += 1;
      continue;
    }

    inputSources.push(arg);
  }

  if (inputSources.length === 0) {
    inputSources.push(DEFAULT_INPUT_SOURCE);
  }

  return {
    inputSources,
    outputXlsx,
    templateXlsx,
  };
}

function captureStyleMatrix(ws, maxRows, maxCols) {
  const matrix = Array.from({ length: maxRows + 1 }, () => Array(maxCols + 1).fill(null));
  for (let row = 1; row <= maxRows; row += 1) {
    for (let col = 1; col <= maxCols; col += 1) {
      matrix[row][col] = deepClone(ws.getCell(row, col).style);
    }
  }
  return matrix;
}

function applyStyleFromMatrix(ws, matrix, row, col) {
  const style = matrix[row] && matrix[row][col] ? matrix[row][col] : {};
  ws.getCell(row, col).style = deepClone(style);
}

function clearRangeValues(ws, startRow, endRow, startCol, endCol) {
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      ws.getCell(row, col).value = null;
    }
  }
}

function readTemplateSummaryLabels(ws) {
  const labels = [];
  for (let row = 2; row <= 120; row += 1) {
    const label = sanitizeText(ws.getCell(row, 13).value);
    if (!label) {
      if (labels.length > 0) {
        break;
      }
      continue;
    }
    labels.push({ row, label });
    if (label.toUpperCase() === "TOTAL GENERAL") {
      break;
    }
  }
  return labels;
}

function buildSummary(rows, templateLabels) {
  const bucket = new Map();
  let totalDebe = 0;
  let totalHaber = 0;

  for (const row of rows) {
    const label = sanitizeText(row.CUENTA);
    if (!bucket.has(label)) {
      bucket.set(label, { debe: 0, haber: 0 });
    }
    const item = bucket.get(label);
    item.debe += row.DEBE;
    item.haber += row.HABER;
    totalDebe += row.DEBE;
    totalHaber += row.HABER;
  }

  const result = [];
  for (const labelItem of templateLabels) {
    const upper = labelItem.label.toUpperCase();
    if (upper === "TOTAL GENERAL") {
      result.push({
        row: labelItem.row,
        label: labelItem.label,
        debe: round2(totalDebe),
        haber: round2(totalHaber),
      });
      continue;
    }
    const values = bucket.get(labelItem.label) || { debe: 0, haber: 0 };
    result.push({
      row: labelItem.row,
      label: labelItem.label,
      debe: round2(values.debe),
      haber: round2(values.haber),
    });
  }

  return result;
}

async function buildWorkbookFromTemplate(templatePath, rows) {
  if (!fs.existsSync(templatePath)) {
    throw new Error(`No se encontro plantilla de Accion 3: ${templatePath}`);
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(templatePath);
  const ws = wb.getWorksheet(SHEET_NAME);
  if (!ws) {
    throw new Error(`La plantilla de Accion 3 no contiene la hoja ${SHEET_NAME}.`);
  }

  const requiredRowCount = Math.max(rows.length + 1, ws.rowCount);
  ensureTemplateCapacity(ws, requiredRowCount);
  const templateRowCount = ws.rowCount;

  const templateStyles = captureStyleMatrix(ws, templateRowCount, 16);
  clearRangeValues(ws, 2, templateRowCount, 1, 11);
  clearRangeValues(ws, 2, templateRowCount, 14, 15);

  for (let row = 1; row <= templateRowCount; row += 1) {
    for (let col = 1; col <= 16; col += 1) {
      applyStyleFromMatrix(ws, templateStyles, row, col);
    }
  }

  for (let i = 0; i < rows.length; i += 1) {
    const rowNum = i + 2;
    const row = rows[i];

    ws.getCell(rowNum, 1).value = row.COD;
    ws.getCell(rowNum, 2).value = row.CUENTA;
    ws.getCell(rowNum, 3).value = row.EXT;
    ws.getCell(rowNum, 4).value = toExcelDateSerial(row.FECHA);
    ws.getCell(rowNum, 5).value = row.ORIGEN;
    ws.getCell(rowNum, 6).value = row.ASIENTO;
    ws.getCell(rowNum, 7).value = row.DOCU;
    ws.getCell(rowNum, 8).value = row.DETALLE;
    ws.getCell(rowNum, 9).value = round2(row.DEBE);
    ws.getCell(rowNum, 10).value = round2(row.HABER);
    ws.getCell(rowNum, 11).value = round2(row.SALDO);

    for (let col = 1; col <= 11; col += 1) {
      applyStyleFromMatrix(ws, templateStyles, rowNum, col);
    }
  }

  const summaryLabels = readTemplateSummaryLabels(ws);
  const summary = buildSummary(rows, summaryLabels);
  for (const item of summary) {
    ws.getCell(item.row, 14).value = round2(item.debe);
    ws.getCell(item.row, 15).value = round2(item.haber);
    for (let col = 13; col <= 16; col += 1) {
      applyStyleFromMatrix(ws, templateStyles, item.row, col);
    }
  }

  return { workbook: wb, summary };
}

async function writeWorkbookWithRetries(workbook, preferredPath, maxAttempts = 20) {
  const parsed = path.parse(preferredPath);
  fs.mkdirSync(parsed.dir || ".", { recursive: true });

  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    const candidate =
      attempt === 0
        ? preferredPath
        : path.join(parsed.dir, `${parsed.name}_nuevo${attempt === 1 ? "" : `_${attempt}`}${parsed.ext}`);

    try {
      await workbook.xlsx.writeFile(candidate);
      return candidate;
    } catch (error) {
      const isLocked = error && (error.code === "EBUSY" || error.code === "EPERM");
      if (!isLocked) {
        throw error;
      }
    }
  }

  throw new Error("No se pudo guardar el Excel de Accion 3. Cierra archivos abiertos e intenta de nuevo.");
}

function toArray(value) {
  if (value === null || value === undefined) {
    return [];
  }
  return Array.isArray(value) ? value : [value];
}

function normalizeZipPath(inputPath) {
  return String(inputPath || "").replace(/\\/g, "/").replace(/^\/+/, "");
}

function parseXml(xmlText) {
  return XML_PARSER.parse(xmlText);
}

function buildXml(xmlObject) {
  const body = XML_BUILDER.build(xmlObject);
  if (body.startsWith("<?xml")) {
    return body;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body}`;
}

function readZipText(zip, entryPath) {
  const entry = zip.getEntry(normalizeZipPath(entryPath));
  if (!entry) {
    return null;
  }
  return entry.getData().toString("utf8");
}

function updateZipText(zip, entryPath, xmlObject) {
  zip.updateFile(normalizeZipPath(entryPath), Buffer.from(buildXml(xmlObject), "utf8"));
}

function getWorksheetEntryPath(zip, sheetName) {
  const workbookXml = readZipText(zip, "xl/workbook.xml");
  const relsXml = readZipText(zip, "xl/_rels/workbook.xml.rels");
  if (!workbookXml || !relsXml) {
    throw new Error("No se pudo leer la estructura interna del Excel.");
  }

  const workbook = parseXml(workbookXml);
  const rels = parseXml(relsXml);
  const sheets = toArray(workbook.workbook?.sheets?.sheet);
  const sheet = sheets.find((item) => sanitizeText(item?.["@_name"]).toUpperCase() === sanitizeText(sheetName).toUpperCase());
  if (!sheet) {
    throw new Error(`No se encontro la hoja ${sheetName} en el Excel.`);
  }

  const sheetRid = sheet["@_r:id"];
  const relationships = toArray(rels.Relationships?.Relationship);
  const rel = relationships.find((item) => item?.["@_Id"] === sheetRid);
  if (!rel) {
    throw new Error(`No se pudo resolver la ruta XML de la hoja ${sheetName}.`);
  }

  const target = normalizeZipPath(rel["@_Target"]);
  return target.startsWith("xl/") ? target : normalizeZipPath(path.posix.join("xl", target));
}

function getRowsArray(sheetXmlObject) {
  return toArray(sheetXmlObject.worksheet?.sheetData?.row);
}

function setRowsArray(sheetXmlObject, rows) {
  if (!sheetXmlObject.worksheet.sheetData) {
    sheetXmlObject.worksheet.sheetData = {};
  }
  sheetXmlObject.worksheet.sheetData.row = rows;
}

function getCellsArray(rowObject) {
  return toArray(rowObject?.c);
}

function setCellsArray(rowObject, cells) {
  if (cells.length === 0) {
    delete rowObject.c;
    return;
  }
  rowObject.c = cells;
}

function parseCellRef(ref) {
  const match = /^([A-Z]+)(\d+)$/i.exec(String(ref || ""));
  if (!match) {
    return null;
  }
  let col = 0;
  const letters = match[1].toUpperCase();
  for (let i = 0; i < letters.length; i += 1) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return { col, row: Number(match[2]) };
}

function columnNumberToName(colNumber) {
  let n = Number(colNumber);
  let result = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function cellRef(colNumber, rowNumber) {
  return `${columnNumberToName(colNumber)}${rowNumber}`;
}

function copyCellPayload(targetCell, sourceCell) {
  if (sourceCell["@_t"] !== undefined) {
    targetCell["@_t"] = sourceCell["@_t"];
  } else {
    delete targetCell["@_t"];
  }
  if (sourceCell.f !== undefined) {
    targetCell.f = deepClone(sourceCell.f);
  } else {
    delete targetCell.f;
  }
  if (sourceCell.v !== undefined) {
    targetCell.v = deepClone(sourceCell.v);
  } else {
    delete targetCell.v;
  }
  if (sourceCell.is !== undefined) {
    targetCell.is = deepClone(sourceCell.is);
  } else {
    delete targetCell.is;
  }
}

function cloneTemplateRow(rowTemplate, targetRowNumber) {
  const clonedRow = deepClone(rowTemplate);
  clonedRow["@_r"] = String(targetRowNumber);

  const cells = getCellsArray(clonedRow);
  for (const cell of cells) {
    const ref = parseCellRef(cell?.["@_r"]);
    if (!ref) {
      continue;
    }
    cell["@_r"] = cellRef(ref.col, targetRowNumber);
  }

  return clonedRow;
}

function clearCellPayload(targetCell) {
  delete targetCell["@_t"];
  delete targetCell.f;
  delete targetCell.v;
  delete targetCell.is;
}

function hasCellPayload(cell) {
  return cell.v !== undefined || cell.f !== undefined || cell.is !== undefined;
}

function stripCalcChain(zip) {
  const calcChainPath = "xl/calcChain.xml";
  if (zip.getEntry(calcChainPath)) {
    zip.deleteFile(calcChainPath);
  }

  const relsXml = readZipText(zip, "xl/_rels/workbook.xml.rels");
  if (relsXml) {
    const relsObj = parseXml(relsXml);
    const rels = toArray(relsObj.Relationships?.Relationship);
    const kept = rels.filter((item) => !String(item?.["@_Type"] || "").endsWith("/calcChain"));
    relsObj.Relationships.Relationship = kept;
    updateZipText(zip, "xl/_rels/workbook.xml.rels", relsObj);
  }

  const contentTypesXml = readZipText(zip, "[Content_Types].xml");
  if (contentTypesXml) {
    const typesObj = parseXml(contentTypesXml);
    const overrides = toArray(typesObj.Types?.Override);
    const kept = overrides.filter((item) => item?.["@_PartName"] !== "/xl/calcChain.xml");
    typesObj.Types.Override = kept;
    updateZipText(zip, "[Content_Types].xml", typesObj);
  }
}

function preserveTemplateVisualWorkbook(templatePath, generatedPath, sheetName = SHEET_NAME) {
  const templateZip = new AdmZip(templatePath);
  const generatedZip = new AdmZip(generatedPath);

  const templateSheetPath = getWorksheetEntryPath(templateZip, sheetName);
  const generatedSheetPath = getWorksheetEntryPath(generatedZip, sheetName);
  const templateSheetXml = readZipText(templateZip, templateSheetPath);
  const generatedSheetXml = readZipText(generatedZip, generatedSheetPath);
  if (!templateSheetXml || !generatedSheetXml) {
    throw new Error("No se pudo comparar hoja XML entre plantilla y generado.");
  }

  const templateSheet = parseXml(templateSheetXml);
  const generatedSheet = parseXml(generatedSheetXml);
  const templateRows = getRowsArray(templateSheet);
  const generatedRows = getRowsArray(generatedSheet);

  const templateRowMap = new Map();
  const generatedRowMap = new Map();
  const generatedCellMap = new Map();
  let templateMaxRow = 1;
  let generatedMaxRow = 1;
  let templatePrototypeRow = null;

  for (const row of templateRows) {
    const rowNumber = Number(row?.["@_r"] || 0);
    if (rowNumber > 0) {
      templateRowMap.set(rowNumber, row);
      templateMaxRow = Math.max(templateMaxRow, rowNumber);
    }
  }

  if (templateRowMap.has(templateMaxRow)) {
    templatePrototypeRow = cloneTemplateRow(templateRowMap.get(templateMaxRow), templateMaxRow);
  }

  for (const row of generatedRows) {
    const rowNumber = Number(row?.["@_r"] || 0);
    if (rowNumber <= 0) {
      continue;
    }
    generatedRowMap.set(rowNumber, row);
    const cells = getCellsArray(row);
    for (const cell of cells) {
      const ref = parseCellRef(cell?.["@_r"]);
      if (!ref || ref.col < 1 || ref.col > 16) {
        continue;
      }
      generatedCellMap.set(`${ref.row}:${ref.col}`, cell);
      if (hasCellPayload(cell)) {
        generatedMaxRow = Math.max(generatedMaxRow, ref.row);
      }
    }
  }

  const maxRow = Math.max(templateMaxRow, generatedMaxRow);
  for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
    let row = templateRowMap.get(rowNumber);
    if (!row) {
      row = templatePrototypeRow ? cloneTemplateRow(templatePrototypeRow, rowNumber) : { "@_r": String(rowNumber) };
      templateRows.push(row);
      templateRowMap.set(rowNumber, row);
    }

    const rowCells = getCellsArray(row);
    const rowCellMap = new Map();
    for (const cell of rowCells) {
      const ref = parseCellRef(cell?.["@_r"]);
      if (!ref) {
        continue;
      }
      rowCellMap.set(ref.col, cell);
    }

    for (let col = 1; col <= 16; col += 1) {
      const key = `${rowNumber}:${col}`;
      const generatedCell = generatedCellMap.get(key);
      let targetCell = rowCellMap.get(col);

      if (!targetCell) {
        if (!generatedCell) {
          continue;
        }
        targetCell = { "@_r": cellRef(col, rowNumber) };
        rowCellMap.set(col, targetCell);
        rowCells.push(targetCell);
      }

      if (generatedCell) {
        copyCellPayload(targetCell, generatedCell);
      } else {
        clearCellPayload(targetCell);
      }
    }

    rowCells.sort((a, b) => {
      const ca = parseCellRef(a?.["@_r"]);
      const cb = parseCellRef(b?.["@_r"]);
      return (ca?.col || 0) - (cb?.col || 0);
    });

    const filtered = rowCells.filter((cell) => cell["@_s"] !== undefined || hasCellPayload(cell));
    setCellsArray(row, filtered);
  }

  templateRows.sort((a, b) => Number(a?.["@_r"] || 0) - Number(b?.["@_r"] || 0));
  setRowsArray(templateSheet, templateRows);
  if (generatedSheet.worksheet?.dimension) {
    templateSheet.worksheet.dimension = deepClone(generatedSheet.worksheet.dimension);
  }
  updateZipText(templateZip, templateSheetPath, templateSheet);

  const generatedSharedStrings = generatedZip.getEntry("xl/sharedStrings.xml");
  if (generatedSharedStrings) {
    templateZip.updateFile("xl/sharedStrings.xml", generatedSharedStrings.getData());
  }

  stripCalcChain(templateZip);
  templateZip.writeZip(generatedPath);
}

function stableStringify(value) {
  if (value === null || value === undefined) {
    return String(value);
  }
  if (typeof value !== "object") {
    return JSON.stringify(value);
  }
  if (Array.isArray(value)) {
    return `[${value.map((item) => stableStringify(item)).join(",")}]`;
  }
  const keys = Object.keys(value).sort();
  const parts = keys.map((key) => `${JSON.stringify(key)}:${stableStringify(value[key])}`);
  return `{${parts.join(",")}}`;
}

function styleSignature(style) {
  return stableStringify(style || {});
}

async function verifyOutputWorkbook(outputPath, templatePath, rowsCount) {
  const templateZip = new AdmZip(templatePath);
  const outputZip = new AdmZip(outputPath);
  const criticalEntries = [
    "xl/pivotTables/pivotTable1.xml",
    "xl/worksheets/_rels/sheet1.xml.rels",
    "xl/pivotCache/pivotCacheDefinition1.xml",
  ];
  for (const entryPath of criticalEntries) {
    const expected = templateZip.getEntry(entryPath);
    const actual = outputZip.getEntry(entryPath);
    if (expected && !actual) {
      throw new Error(`Validacion final: falta artefacto visual de plantilla (${entryPath}).`);
    }
  }

  const wb = XLSX.readFile(outputPath, { cellFormula: true });
  if (wb.SheetNames.length !== 1 || wb.SheetNames[0] !== SHEET_NAME) {
    throw new Error(`Validacion final: el archivo debe tener una sola hoja llamada ${SHEET_NAME}.`);
  }

  const ws = wb.Sheets[SHEET_NAME];
  const headers = EXPECTED_HEADERS.map((_, idx) => ws[XLSX.utils.encode_cell({ r: 0, c: idx })]?.v || "");
  for (let i = 0; i < EXPECTED_HEADERS.length; i += 1) {
    if (sanitizeText(headers[i]).toUpperCase() !== EXPECTED_HEADERS[i]) {
      throw new Error(`Validacion final: encabezado incorrecto en columna ${i + 1}.`);
    }
  }

  for (let row = 2; row <= rowsCount + 1; row += 1) {
    const dateCell = ws[`D${row}`];
    if (!dateCell || typeof dateCell.v !== "number") {
      throw new Error(`Validacion final: FECHA invalida en fila ${row}.`);
    }
    const fractionalPart = Math.abs(dateCell.v - Math.trunc(dateCell.v));
    if (fractionalPart > 1e-9) {
      throw new Error(`Validacion final: FECHA con fraccion horaria en fila ${row}.`);
    }
  }

  const templateWb = new ExcelJS.Workbook();
  await templateWb.xlsx.readFile(templatePath);
  const outWb = new ExcelJS.Workbook();
  await outWb.xlsx.readFile(outputPath);

  const templateWs = templateWb.getWorksheet(SHEET_NAME);
  const outWs = outWb.getWorksheet(SHEET_NAME);
  if (!templateWs || !outWs) {
    throw new Error(`Validacion final: no se pudo comparar formato de hoja ${SHEET_NAME}.`);
  }

  for (let col = 1; col <= 16; col += 1) {
    const expected = styleSignature(templateWs.getCell(1, col).style);
    const actual = styleSignature(outWs.getCell(1, col).style);
    if (expected !== actual) {
      throw new Error(`Validacion final: estilo de encabezado alterado en fila 1, columna ${col}.`);
    }
  }

  const lastDataRow = rowsCount + 1;
  if (lastDataRow >= 2) {
    const expectedDataStyleRow = Math.min(lastDataRow, templateWs.rowCount);
    for (let col = 1; col <= 11; col += 1) {
      const expected = styleSignature(templateWs.getCell(expectedDataStyleRow, col).style);
      const actual = styleSignature(outWs.getCell(lastDataRow, col).style);
      if (expected !== actual) {
        throw new Error(`Validacion final: estilo de datos alterado en ${lastDataRow}:${col}.`);
      }
    }
  }

  const summaryLabels = readTemplateSummaryLabels(templateWs);
  for (const item of summaryLabels) {
    for (let col = 13; col <= 16; col += 1) {
      const expected = styleSignature(templateWs.getCell(item.row, col).style);
      const actual = styleSignature(outWs.getCell(item.row, col).style);
      if (expected !== actual) {
        throw new Error(`Validacion final: estilo lateral alterado en ${item.row}:${col}.`);
      }
    }
  }
}

function writeAuditReport(auditPath, payload) {
  fs.mkdirSync(path.dirname(auditPath), { recursive: true });
  fs.writeFileSync(auditPath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
}

async function main() {
  const cli = parseCliArguments();
  const outputPath = path.resolve(process.cwd(), cli.outputXlsx);
  const templatePath = cli.templateXlsx;
  const parsed = await parseInputSources(cli.inputSources);
  const rows = parsed.rows;

  validateRows(rows);

  const { workbook, summary } = await buildWorkbookFromTemplate(templatePath, rows);
  const finalOutputPath = await writeWorkbookWithRetries(workbook, outputPath);
  preserveTemplateVisualWorkbook(templatePath, finalOutputPath, SHEET_NAME);
  await verifyOutputWorkbook(finalOutputPath, templatePath, rows.length);

  const auditPath = path.join(
    path.dirname(finalOutputPath),
    `${path.parse(finalOutputPath).name}_auditoria.json`,
  );

  const auditPayload = {
    fecha_proceso: new Date().toISOString(),
    input_source: parsed.inputPaths.length === 1 ? parsed.inputPaths[0] : parsed.inputPaths,
    input_sources: parsed.inputPaths,
    input_pdf: parsed.inputExt === ".pdf" ? parsed.inputPaths[0] : null,
    input_tipo: parsed.diagnostics.source_type || parsed.inputExt.replace(".", ""),
    archivos_origen: parsed.sourceFiles,
    total_archivos_origen: parsed.inputPaths.length,
    output_xlsx: finalOutputPath,
    hoja_salida: SHEET_NAME,
    movimientos_extraidos: rows.length,
    filas_fecha_detectadas: parsed.diagnostics.date_rows_detected,
    filas_fecha_omitidas: parsed.diagnostics.skipped_date_rows,
    cuentas_validadas_fuente: parsed.diagnostics.account_totals_checked,
    cuentas_descuadradas_fuente: parsed.diagnostics.account_total_mismatches,
    cuentas_validadas_pdf: parsed.diagnostics.account_totals_checked,
    cuentas_descuadradas_pdf: parsed.diagnostics.account_total_mismatches,
    resumen_filtrado: summary.length,
    verificacion_final_ok: true,
  };
  writeAuditReport(auditPath, auditPayload);

  if (parsed.inputPaths.length === 1) {
    console.log(`Archivo leido: ${parsed.inputPaths[0]}`);
  } else {
    console.log(`Archivos leidos: ${parsed.inputPaths.length}`);
    for (const inputPath of parsed.inputPaths) {
      console.log(`- ${inputPath}`);
    }
  }
  console.log(`Movimientos extraidos: ${rows.length}`);
  console.log(`Resumen lateral: ${summary.length} filas`);
  console.log(`Excel generado (una sola hoja): ${finalOutputPath}`);
  console.log(`Auditoria JSON: ${auditPath}`);
}

main().catch((error) => {
  console.error("Error:", error.message);
  process.exit(1);
});
