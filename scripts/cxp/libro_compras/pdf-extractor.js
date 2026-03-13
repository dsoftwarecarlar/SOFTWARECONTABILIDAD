const fs = require("fs");
const path = require("path");
const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.mjs");

const {
  COLUMN_ORDER,
  HEADER_ALIASES,
  FIXED_BOUNDARIES,
  LEFT_BOUNDARIES,
  NUMERIC_COLUMNS,
} = require("./constants");
const { sanitizeText } = require("./row-utils");

function assignToColumn(x, boundaries) {
  for (const column of boundaries) {
    if (x >= column.left && x < column.right) {
      return column.name;
    }
  }
  return null;
}

function groupItemsByRow(items, tolerance = 0.8) {
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

function isDataRow(rowItems, boundaries) {
  const codeColumn = boundaries.find((boundary) => boundary.name === "CODIGO");
  if (!codeColumn) {
    return false;
  }

  const codeItem = rowItems.find((item) => item.x >= codeColumn.left && item.x < codeColumn.right);
  if (!codeItem) {
    return false;
  }

  return /^[A-Z0-9\-]{3,}$/.test(codeItem.str.replace(/\s+/g, ""));
}

async function extractRowsFromPdf(pdfPath) {
  const data = new Uint8Array(fs.readFileSync(pdfPath));
  const standardFontDir = path.join(__dirname, "..", "..", "..", "node_modules", "pdfjs-dist", "standard_fonts");
  const standardFontDataUrl = `${standardFontDir.replace(/\\/g, "/")}/`;

  const loadingTask = pdfjsLib.getDocument({
    data,
    standardFontDataUrl,
    verbosity: pdfjsLib.VerbosityLevel.ERRORS,
  });

  const pdf = await loadingTask.promise;
  const allRows = [];

  try {
    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum += 1) {
      const page = await pdf.getPage(pageNum);
      const content = await page.getTextContent();

      const items = content.items
        .map((item) => ({
          str: sanitizeText(item.str),
          x: Number(item.transform[4]),
          y: Number(item.transform[5]),
        }))
        .filter((item) => item.str.length > 0);

      const headerCandidates = items.filter((item) => HEADER_ALIASES[item.str]);
      if (headerCandidates.length < 6) {
        continue;
      }

      const headerMap = new Map();
      for (const candidate of headerCandidates) {
        const canonical = HEADER_ALIASES[candidate.str];
        if (!headerMap.has(canonical)) {
          headerMap.set(canonical, candidate.x);
        }
      }

      const hasEnoughHeaders = COLUMN_ORDER.filter((name) => headerMap.has(name)).length >= 10;
      if (!hasEnoughHeaders) {
        continue;
      }

      const headerY = headerCandidates[0].y;
      const belowHeader = items.filter((item) => item.y < headerY - 1);
      const groupedRows = groupItemsByRow(belowHeader);

      for (const groupedRow of groupedRows) {
        if (!isDataRow(groupedRow.items, FIXED_BOUNDARIES)) {
          continue;
        }

        const row = {};
        for (const col of COLUMN_ORDER) {
          row[col] = "";
        }

        const leftItems = groupedRow.items.filter((item) => item.x < 400);
        const numericItems = groupedRow.items.filter((item) => item.x >= 400).sort((a, b) => a.x - b.x);

        for (const item of leftItems) {
          const colName = assignToColumn(item.x, LEFT_BOUNDARIES);
          if (!colName) {
            continue;
          }
          row[colName] = row[colName] ? `${row[colName]} ${item.str}` : item.str;
        }

        if (numericItems.length >= NUMERIC_COLUMNS.length) {
          const selected = numericItems.slice(-NUMERIC_COLUMNS.length);
          for (let i = 0; i < NUMERIC_COLUMNS.length; i += 1) {
            row[NUMERIC_COLUMNS[i]] = selected[i].str;
          }
        } else {
          for (const item of numericItems) {
            const colName = assignToColumn(item.x, FIXED_BOUNDARIES);
            if (!colName) {
              continue;
            }
            row[colName] = row[colName] ? `${row[colName]} ${item.str}` : item.str;
          }
        }

        const hasDate = /^\d{2}\/\d{2}\/\d{4}$/.test(row.FECHA);
        const hasTipo = /^[A-Z]{2,3}$/.test(row.TIPO);
        const hasDoc = row.DOCUMENTO.length > 0;
        if (hasDate && hasTipo && hasDoc) {
          allRows.push(row);
        }
      }
    }
  } finally {
    await loadingTask.destroy();
  }

  return allRows;
}

module.exports = {
  extractRowsFromPdf,
};
