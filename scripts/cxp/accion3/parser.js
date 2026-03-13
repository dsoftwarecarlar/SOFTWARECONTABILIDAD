const fs = require("fs");
const path = require("path");

const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.mjs");

const {
  sanitizeText,
  round2,
  parseIntLike,
  parseDecimalLike,
} = require("../shared/core-utils");
const {
  MOVEMENT_BOUNDARIES,
} = require("./constants");

function normalizeReference(value) {
  return sanitizeText(value).replace(/\s+/g, "");
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
    result[name] = ["DEBE", "HABER", "SALDO"].includes(name)
      ? pieces.join("")
      : sanitizeText(pieces.join(" "));
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

  return /^\d{2}-[A-Z]{3}-\d{2}$/.test(fecha)
    && /^[A-Z]{2,6}$/.test(origen)
    && /[0-9]/.test(saldoText);
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
        skippedDateRows.push({ line: index + 1, reason: "invalid_date_value", text: sanitizeText(line) });
      }
      continue;
    }
    if (!/^\d{2}\.\d{2}\.\d{2}\.\d{2}\.\d{4}$/.test(code)) {
      skippedDateRows.push({ line: index + 1, reason: "invalid_account_code", text: sanitizeText(line) });
      continue;
    }
    if (!accountName) {
      skippedDateRows.push({ line: index + 1, reason: "missing_account_name", text: sanitizeText(line) });
      continue;
    }
    if (!/^[A-Z]{2,6}$/.test(origen)) {
      skippedDateRows.push({ line: index + 1, reason: "invalid_origin", text: sanitizeText(line) });
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
    const preview = skippedDateRows.slice(0, 5).map((item) => `linea ${item.line} (${item.reason})`).join(" | ");
    throw new Error(`Se detectaron ${skippedDateRows.length} filas con fecha que no pudieron mapearse en TXT Accion 3. ${preview}`);
  }

  const totalsValidation = validateAccountTotals(rows, accountHeaderTotals);
  if (totalsValidation.mismatches.length > 0) {
    const preview = totalsValidation.mismatches
      .slice(0, 4)
      .map(
        (item) => `${item.code} (${item.name}): `
          + `debe TXT=${item.expected.total_debe} / Excel=${item.actual.total_debe}, `
          + `haber TXT=${item.expected.total_haber} / Excel=${item.actual.total_haber}, `
          + `saldo TXT=${item.expected.saldo_final} / Excel=${item.actual.saldo_final}`,
      )
      .join(" | ");
    throw new Error(`Validacion contable fallida en TXT Accion 3 (${totalsValidation.mismatches.length} cuentas). ${preview}`);
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
          skippedDateRows.push({ page: pageNumber, y: row.y, reason: "not_movement_row", text: rowText });
        }
        continue;
      }
      if (!currentCode || !currentName) {
        skippedDateRows.push({ page: pageNumber, y: row.y, reason: "missing_account_context", text: rowText });
        continue;
      }

      const fecha = parseDateFromReport(fields.FECHA);
      if (!fecha) {
        skippedDateRows.push({ page: pageNumber, y: row.y, reason: "invalid_date_value", text: rowText });
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
    const preview = skippedDateRows.slice(0, 5).map((item) => `pag ${item.page} (${item.reason}): ${item.text}`).join(" | ");
    throw new Error(`Se detectaron ${skippedDateRows.length} filas con fecha que no pudieron mapearse en Accion 3. ${preview}`);
  }

  const totalsValidation = validateAccountTotals(rows, accountHeaderTotals);
  if (totalsValidation.mismatches.length > 0) {
    const preview = totalsValidation.mismatches
      .slice(0, 4)
      .map(
        (item) => `${item.code} (${item.name}): `
          + `debe PDF=${item.expected.total_debe} / Excel=${item.actual.total_debe}, `
          + `haber PDF=${item.expected.total_haber} / Excel=${item.actual.total_haber}, `
          + `saldo PDF=${item.expected.saldo_final} / Excel=${item.actual.saldo_final}`,
      )
      .join(" | ");
    throw new Error(`Validacion contable fallida en Accion 3 (${totalsValidation.mismatches.length} cuentas). ${preview}`);
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
      sourceFiles: [{
        input_source: resolvedInputs[0],
        file_name: path.basename(resolvedInputs[0]),
        movimientos_extraidos: parsed.rows.length,
        ...parsed.diagnostics,
      }],
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
      sourceFiles: [{
        input_source: resolvedInputs[0],
        file_name: path.basename(resolvedInputs[0]),
        movimientos_extraidos: parsed.rows.length,
        ...parsed.diagnostics,
      }],
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

module.exports = {
  parseInputSources,
  validateRows,
};
