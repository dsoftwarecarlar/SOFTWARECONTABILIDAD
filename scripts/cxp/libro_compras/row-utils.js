const {
  COLUMN_ORDER,
  SPECIAL_PLAN,
  SPECIAL_ACTIVO,
  SPECIAL_FE_IN_ND,
} = require("./constants");

const INVALID_XML_CONTROL_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g;

function sanitizeText(text) {
  return String(text || "")
    .replace(INVALID_XML_CONTROL_CHARS, " ")
    .replace(/\s+/g, " ")
    .trim();
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

  const number = Number(normalized);
  return Number.isFinite(number) ? number : 0;
}

function parseIntLike(value) {
  const normalized = String(value || "").replace(/[^\d]/g, "");
  if (!normalized) {
    return null;
  }
  const number = Number(normalized);
  return Number.isFinite(number) ? number : null;
}

function normalizeNumericText(value) {
  const clean = sanitizeText(value);
  if (!clean) {
    return "";
  }
  if (!/^\d+$/.test(clean)) {
    return clean;
  }
  return String(Number(clean));
}

function normalizeDocument(value) {
  const clean = sanitizeText(value);
  if (!clean) {
    return "";
  }
  if (/^\d+$/.test(clean)) {
    return String(Number(clean));
  }
  return clean;
}

function parseDateToExcelSerial(dateText) {
  const clean = sanitizeText(dateText);
  const match = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(clean);
  if (!match) {
    return null;
  }

  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3]);
  const utcMillis = Date.UTC(year, month - 1, day);
  return Math.floor(utcMillis / 86400000) + 25569;
}

function isLikelyNumberText(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return true;
  }

  const clean = sanitizeText(value);
  if (!clean) {
    return false;
  }

  return /^-?(?:\d{1,3}(?:[.,]\d{3})+|\d+)(?:[.,]\d+)?$/.test(clean);
}

function validateRows(rows, options = {}) {
  const { strict = true, autofillNumericBlanks = true } = options;
  const requiredCols = ["CODIGO", "CEDULA", "NOMBRE", "FECHA", "TIPO", "DOCUMENTO"];
  const numericCols = ["MONTO", "BASE IVA", "BASE 0", "IMPUESTOS", "RETENCION", "SALDO"];
  const problems = [];

  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    const rowNo = i + 1;

    for (const col of requiredCols) {
      if (!sanitizeText(row[col])) {
        problems.push(`Fila ${rowNo}: ${col} vacio.`);
      }
    }

    const fechaAsSerial = parseDateToExcelSerial(row.FECHA);
    const isExcelDateNumber = typeof row.FECHA === "number" && Number.isFinite(row.FECHA) && row.FECHA > 30000;
    if (!fechaAsSerial && !isExcelDateNumber) {
      problems.push(`Fila ${rowNo}: FECHA invalida (${row.FECHA}).`);
    }

    if (!/^[A-Z]{2,3}$/.test(sanitizeText(row.TIPO))) {
      problems.push(`Fila ${rowNo}: TIPO invalido (${row.TIPO}).`);
    }

    for (const col of numericCols) {
      const raw = row[col];
      const rawText = sanitizeText(raw);
      if (!rawText) {
        if (autofillNumericBlanks) {
          row[col] = 0;
        } else {
          problems.push(`Fila ${rowNo}: ${col} vacio.`);
        }
        continue;
      }
      if (!isLikelyNumberText(raw)) {
        problems.push(`Fila ${rowNo}: ${col} no parece numerico (${rawText}).`);
      }
    }
  }

  if (strict && problems.length > 0) {
    const preview = problems.slice(0, 8).join(" | ");
    throw new Error(`Validacion fallida (${problems.length} problemas). ${preview}`);
  }

  return problems;
}

function deepClone(obj) {
  return obj ? JSON.parse(JSON.stringify(obj)) : {};
}

function rowKey(row) {
  return [
    normalizeNumericText(row.CODIGO),
    normalizeNumericText(row.CEDULA),
    sanitizeText(row.TIPO),
    normalizeDocument(row.DOCUMENTO),
    parseDateToExcelSerial(row.FECHA) || "",
  ].join("|");
}

function rowSignature(row) {
  const fecha = parseDateToExcelSerial(row.FECHA) || row.FECHA || "";
  const amounts = ["MONTO", "BASE IVA", "BASE 0", "IMPUESTOS", "RETENCION", "SALDO"].map((field) =>
    parseDecimalLike(row[field]).toFixed(2),
  );

  return [
    normalizeNumericText(row.CODIGO),
    normalizeNumericText(row.CEDULA),
    sanitizeText(row.NOMBRE),
    fecha,
    sanitizeText(row.TIPO),
    normalizeDocument(row.DOCUMENTO),
    ...amounts,
  ].join("|");
}

function buildMultiset(rows) {
  const map = new Map();
  for (const row of rows) {
    const signature = rowSignature(row);
    map.set(signature, (map.get(signature) || 0) + 1);
  }
  return map;
}

function multisetDiff(left, right) {
  const diff = [];
  for (const [signature, countLeft] of left) {
    const countRight = right.get(signature) || 0;
    if (countLeft > countRight) {
      diff.push({ signature, count: countLeft - countRight });
    }
  }
  return diff;
}

function auditRowsConsistency(generatedRows, templateRows) {
  if (!templateRows || templateRows.length === 0) {
    return {
      enabled: false,
      ok: true,
      generatedCount: generatedRows.length,
      templateCount: 0,
      extraGenerated: [],
      missingGenerated: [],
    };
  }

  const generatedSet = buildMultiset(generatedRows);
  const templateSet = buildMultiset(templateRows);
  const extraGenerated = multisetDiff(generatedSet, templateSet);
  const missingGenerated = multisetDiff(templateSet, generatedSet);

  return {
    enabled: true,
    ok: extraGenerated.length === 0 && missingGenerated.length === 0,
    generatedCount: generatedRows.length,
    templateCount: templateRows.length,
    extraGenerated,
    missingGenerated,
  };
}

function toExcelDataRow(row, note = "") {
  const codigo = parseIntLike(row.CODIGO);
  const cedula = parseIntLike(row.CEDULA);
  const fecha = parseDateToExcelSerial(row.FECHA);
  const docClean = normalizeDocument(row.DOCUMENTO);
  const docNumeric = /^\d+$/.test(docClean) ? Number(docClean) : docClean;

  return [
    codigo == null ? sanitizeText(row.CODIGO) : codigo,
    cedula == null ? sanitizeText(row.CEDULA) : cedula,
    sanitizeText(String(row.NOMBRE || "").replace(/\r?\n/g, " ")),
    fecha == null ? sanitizeText(row.FECHA) : fecha,
    sanitizeText(row.TIPO),
    docNumeric,
    parseDecimalLike(row.MONTO),
    parseDecimalLike(row["BASE IVA"]),
    parseDecimalLike(row["BASE 0"]),
    parseDecimalLike(row.IMPUESTOS),
    parseDecimalLike(row.RETENCION),
    parseDecimalLike(row.SALDO),
    sanitizeText(note),
  ];
}

function headerRow() {
  return [...COLUMN_ORDER, ""];
}

function insertAfterLastMatching(rows, rowToInsert, predicate) {
  if (!rowToInsert) {
    return rows;
  }

  let index = -1;
  for (let i = 0; i < rows.length; i += 1) {
    if (predicate(rows[i])) {
      index = i;
    }
  }

  if (index === -1) {
    return [rowToInsert, ...rows];
  }

  const clone = [...rows];
  clone.splice(index + 1, 0, rowToInsert);
  return clone;
}

function sumField(rows, fieldName) {
  return rows.reduce((sum, row) => sum + parseDecimalLike(row[fieldName]), 0);
}

function buildSingleSheetRows(rows) {
  const pending = rows.map((row, idx) => ({ ...row, _idx: idx }));

  const takeSpecial = (rule) => {
    const idx = pending.findIndex(
      (row) =>
        normalizeNumericText(row.CODIGO) === rule.codigo &&
        normalizeDocument(row.DOCUMENTO) === rule.doc &&
        sanitizeText(row.TIPO) === rule.tipo,
    );
    if (idx === -1) {
      return null;
    }
    const [picked] = pending.splice(idx, 1);
    return picked;
  };

  const specialPlan = takeSpecial(SPECIAL_PLAN);
  const specialActivo = takeSpecial(SPECIAL_ACTIVO);
  const specialFeInNd = takeSpecial(SPECIAL_FE_IN_ND);

  const mainTypeOrder = { FE: 1, LC: 2, OT: 3 };
  const ndtrTypeOrder = { ND: 1, TR: 2 };

  const mainRows = pending
    .filter((row) => ["FE", "LC", "OT"].includes(sanitizeText(row.TIPO)))
    .sort(
      (a, b) =>
        (mainTypeOrder[sanitizeText(a.TIPO)] || 99) - (mainTypeOrder[sanitizeText(b.TIPO)] || 99) ||
        a._idx - b._idx,
    );

  const rimpeRows = pending
    .filter((row) => sanitizeText(row.TIPO) === "NV")
    .sort((a, b) => a._idx - b._idx);

  let ndtrRows = pending
    .filter((row) => ["ND", "TR"].includes(sanitizeText(row.TIPO)))
    .sort(
      (a, b) =>
        (ndtrTypeOrder[sanitizeText(a.TIPO)] || 99) - (ndtrTypeOrder[sanitizeText(b.TIPO)] || 99) ||
        a._idx - b._idx,
    );

  ndtrRows = insertAfterLastMatching(
    ndtrRows,
    specialFeInNd,
    (row) =>
      sanitizeText(row.TIPO) === "ND" &&
      normalizeNumericText(row.CODIGO) === normalizeNumericText(specialFeInNd?.CODIGO) &&
      normalizeNumericText(row.CEDULA) === normalizeNumericText(specialFeInNd?.CEDULA),
  );

  const aoa = [];
  aoa.push(headerRow());
  for (const row of mainRows) {
    aoa.push(toExcelDataRow(row));
  }

  const mainStartRow = 2;
  const mainEndRow = mainStartRow + mainRows.length - 1;
  const total1Row = mainEndRow + 1;
  const mayorIvaRow = total1Row + 1;
  const specialPlanRow = mayorIvaRow + 2;
  const specialActivoRow = mayorIvaRow + 3;
  const totalAtsRow = mayorIvaRow + 5;
  const rimpeLabelRow = mayorIvaRow + 7;
  const rimpeHeaderRow = mayorIvaRow + 8;
  const rimpeStartRow = mayorIvaRow + 9;
  const rimpeEndRow = rimpeStartRow + rimpeRows.length - 1;
  const rimpeSubtotalRow = rimpeEndRow + 1;
  const ndtrLabelRow = rimpeSubtotalRow + 4;
  const ndtrHeaderRow = rimpeSubtotalRow + 5;

  aoa.push(["", "", "", "", "", "", "", null, null, null, "", "", ""]);
  aoa.push(["", "", "", "", "", "", "", "", "", null, null, "MAYOR IVA", ""]);
  aoa.push([]);
  aoa.push(specialPlan ? toExcelDataRow(specialPlan, SPECIAL_PLAN.note) : []);
  aoa.push(specialActivo ? toExcelDataRow(specialActivo, SPECIAL_ACTIVO.note) : []);
  aoa.push([]);
  aoa.push(["", "", "", "", "", "", "", null, null, null, "IVA ATS", "", ""]);
  aoa.push([]);
  aoa.push(["RIMPE NEGOCIO POPULAR"]);
  aoa.push(headerRow());

  for (const row of rimpeRows) {
    aoa.push(toExcelDataRow(row));
  }

  aoa.push(["", "", "", "", "", "", "", "", null, "", "", "", ""]);
  aoa.push([]);
  aoa.push([]);
  aoa.push([]);
  aoa.push(["NDS, TR, ANULACIONES"]);
  aoa.push(headerRow());

  for (const row of ndtrRows) {
    aoa.push(toExcelDataRow(row));
  }

  const sums = {
    mainH: sumField(mainRows, "BASE IVA"),
    mainI: sumField(mainRows, "BASE 0"),
    mainJ: sumField(mainRows, "IMPUESTOS"),
    planH: specialPlan ? parseDecimalLike(specialPlan["BASE IVA"]) : 0,
    planJ: specialPlan ? parseDecimalLike(specialPlan.IMPUESTOS) : 0,
    activoH: specialActivo ? parseDecimalLike(specialActivo["BASE IVA"]) : 0,
    activoJ: specialActivo ? parseDecimalLike(specialActivo.IMPUESTOS) : 0,
    rimpeI: sumField(rimpeRows, "BASE 0"),
  };

  return {
    aoa,
    meta: {
      mainStartRow,
      mainEndRow,
      total1Row,
      mayorIvaRow,
      specialPlanRow,
      specialActivoRow,
      totalAtsRow,
      rimpeLabelRow,
      rimpeHeaderRow,
      rimpeStartRow,
      rimpeEndRow,
      rimpeSubtotalRow,
      ndtrLabelRow,
      ndtrHeaderRow,
      sums,
      counts: {
        total: rows.length,
        main: mainRows.length,
        rimpe: rimpeRows.length,
        ndtr: ndtrRows.length,
      },
    },
  };
}

function roundCurrency(value) {
  return Math.round((Number(value || 0) + Number.EPSILON) * 100) / 100;
}

function normalizeOutputCellValue(value) {
  if (value == null || value === "") {
    return null;
  }
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }
  const clean = sanitizeText(value);
  return clean || null;
}

module.exports = {
  sanitizeText,
  parseDecimalLike,
  parseIntLike,
  normalizeNumericText,
  normalizeDocument,
  parseDateToExcelSerial,
  isLikelyNumberText,
  validateRows,
  deepClone,
  rowKey,
  auditRowsConsistency,
  toExcelDataRow,
  buildSingleSheetRows,
  roundCurrency,
  normalizeOutputCellValue,
};
