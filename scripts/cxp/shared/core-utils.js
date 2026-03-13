const fs = require("fs");
const path = require("path");

const INVALID_XML_CONTROL_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g;

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

  throw new Error("No se pudo guardar el Excel de salida. Cierra los archivos abiertos y vuelve a intentar.");
}

function writeAuditReport(auditPath, payload) {
  fs.mkdirSync(path.dirname(auditPath), { recursive: true });
  fs.writeFileSync(auditPath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
}

module.exports = {
  sanitizeText,
  deepClone,
  round2,
  toExcelDateSerial,
  parseIntLike,
  parseDecimalLike,
  writeWorkbookWithRetries,
  writeAuditReport,
};
