const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const XLSX = require("xlsx");

// --- UTILITIES ---

function parseArgs(argv) {
  const args = {};
  for (let i = 2; i < argv.length; i += 2) {
    const key = argv[i].replace(/^-+/, "");
    const value = argv[i + 1];
    args[key] = value;
  }
  return args;
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(`❌ ERROR DE AUDITORÍA: ${message}`);
  }
}

// --- LOGIC DE LECTURA Y NORMALIZACIÓN ---

function loadWorkbook(filePath) {
  assert(fs.existsSync(filePath), `El archivo no existe: ${filePath}`);
  return XLSX.readFile(filePath, { cellFormula: true, cellDates: true });
}

function getSheet(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  assert(!!sheet, `La hoja '${sheetName}' no se encontró en el libro.`);
  return sheet;
}

function getNormalizedCellText(sheet, r, c) {
  const cellAddress = XLSX.utils.encode_cell({ r, c });
  const cell = sheet[cellAddress];
  if (!cell || cell.v === null || cell.v === undefined) return "";

  let value = cell.w !== undefined ? cell.w : cell.v;

  if (value instanceof Date) {
    value.setHours(0, 0, 0, 0);
    return value.toISOString().split("T")[0];
  }

  if (typeof value === "number") {
    return value.toFixed(4);
  }

  // Corrección final: colapsar espacios Y quitar otros caracteres invisibles como el non-breaking space (160)
  return String(value).trim().replace(/\s+/g, ' ').replace(/\\u00A0/g, ' ');
}

function getSheetSignature(sheet, startRow, endRow, columns) {
  const rows = [];
  for (let r = startRow; r <= endRow; ++r) {
    const rowData = [];
    let hasData = false;
    for (const c of columns) {
      const cellValue = getNormalizedCellText(sheet, r, c);
      rowData.push(cellValue);
      if (cellValue !== "") hasData = true;
    }
    if (hasData) {
      rows.push(rowData.join("|"));
    }
  }
  const finalString = rows.join("\\n");
  return {
    rowCount: rows.length,
    hash: crypto.createHash("sha256").update(finalString).digest("hex"),
  };
}

function findRow(sheet, text) {
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    for (let r = range.s.r; r <= range.e.r; ++r) {
        for (let c = range.s.c; c <= range.e.c; ++c) {
            if (String(getNormalizedCellText(sheet, r, c)).includes(text)) {
                return r;
            }
        }
    }
    return -1;
}


// --- VERIFICACIONES DE AUDITORÍA ---

function auditRepSheet(outputWb, sourceWb, mapping) {
  console.log(`  🔎 Auditando hoja de datos: ${mapping.out}...`);
  const outSheet = getSheet(outputWb, mapping.out);
  const srcSheet = getSheet(sourceWb, sourceWb.SheetNames[0]);

  const REP_PAYLOAD_COLUMNS_IDX = [
    0, 1, 2, 4, 5, 6, 7, 8, 9, 10, 12, 15, 17, 19, 21, 22, 23, 24, 25,
    26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39
  ];
  
  const srcEndRow = findRow(srcSheet, 'TOTAL GENERAL');
  const outEndRow = findRow(outSheet, 'TOTAL GENERAL');
  assert(srcEndRow !== -1, `No se encontró 'TOTAL GENERAL' en la hoja origen de ${mapping.out}`);
  assert(outEndRow !== -1, `No se encontró 'TOTAL GENERAL' en la hoja de salida ${mapping.out}`);

  const srcSig = getSheetSignature(srcSheet, 10, srcEndRow - 1, REP_PAYLOAD_COLUMNS_IDX);
  const outSig = getSheetSignature(outSheet, 10, outEndRow - 1, REP_PAYLOAD_COLUMNS_IDX);

  assert(srcSig.rowCount === outSig.rowCount, `Discrepancia en el número de filas para '${mapping.out}'. Esperado: ${srcSig.rowCount}, Encontrado: ${outSig.rowCount}.`);
  
  if (srcSig.hash !== outSig.hash) {
    for (let r = 10; r < srcEndRow; ++r) {
      for (const c of REP_PAYLOAD_COLUMNS_IDX) {
        const srcVal = getNormalizedCellText(srcSheet, r, c);
        const outVal = getNormalizedCellText(outSheet, r, c);
        if (srcVal !== outVal) {
          const cellAddress = XLSX.utils.encode_cell({ r, c });
          console.error('DEBUG FINAL: Source Char Codes:', JSON.stringify(srcVal.split('').map(char => char.charCodeAt(0))));
          console.error('DEBUG FINAL: Output Char Codes:', JSON.stringify(outVal.split('').map(char => char.charCodeAt(0))));
          assert(false, `Discrepancia de contenido en '${mapping.out}' celda ${cellAddress}. Origen: '${srcVal}', Salida: '${outVal}'.`);
        }
      }
    }
    assert(false, `Discrepancia en el contenido (hash) para '${mapping.out}', pero no se encontró diferencia celda por celda.`);
  }

  console.log(`    ✅ OK: Datos y número de filas coinciden.`);
}

function auditStaticSheet(outputWb, templateWb, sheetName) {
  console.log(`  🔎 Auditando hoja estática: ${sheetName}...`);
  const outSheet = getSheet(outputWb, sheetName);
  const tplSheet = getSheet(templateWb, sheetName);
  
  const range = XLSX.utils.decode_range(tplSheet['!ref']);

  for (let r = range.s.r; r <= range.e.r; ++r) {
    for (let c = range.s.c; c <= range.e.c; ++c) {
      const tplVal = getNormalizedCellText(tplSheet, r, c);
      const outVal = getNormalizedCellText(outSheet, r, c);
      if (tplVal !== outVal) {
        const cellAddress = XLSX.utils.encode_cell({r, c});
        assert(false, `La hoja estática '${sheetName}' fue modificada en la celda ${cellAddress}. Esperado: '${tplVal}', Encontrado: '${outVal}'.`);
      }
    }
  }

  console.log(`    ✅ OK: La hoja se ha conservado intacta.`);
}


// --- SCRIPT PRINCIPAL ---

async function main() {
  console.log('--- Iniciando Auditoría de Integridad para Tarea 3 ---');
  const args = parseArgs(process.argv);
  
  const requiredArgs = ['output', 'template', 'source-tyt', 'source-peug', 'source-chgn', 'source-szk'];
  for (const arg of requiredArgs) {
    assert(args[arg], `Falta el argumento requerido: --${arg}`);
  }

  const outputWb = loadWorkbook(args.output);
  const templateWb = loadWorkbook(args.template);
  const sources = {
    tyt: loadWorkbook(args['source-tyt']),
    peug: loadWorkbook(args['source-peug']),
    chgn: loadWorkbook(args['source-chgn']),
    szk: loadWorkbook(args['source-szk']),
  };

  const mappings = [
    { key: 'tyt', out: 'REP TYT', my: 'MY REP TYT', nc: 'NC REP TYT' },
    { key: 'peug', out: 'REP PEUGT', my: 'MY REP PEUG', nc: 'NC REP PEUG' },
    { key: 'chgn', out: 'REP CHGN', my: 'MY REP CHGN', nc: null },
    { key: 'szk', out: 'REP SZK', my: 'MY REP SZK', nc: 'NC REP SZK' },
  ];

  for (const m of mappings) {
    auditRepSheet(outputWb, sources[m.key], m);
    if (m.nc) {
      auditStaticSheet(outputWb, templateWb, m.nc);
    }
  }

  auditStaticSheet(outputWb, templateWb, 'MAYOR IVA');

  console.log('\\n✅ AUDITORÍA APROBADA: El archivo generado es íntegro y correcto.');
}

main().catch(err => {
  console.error(err.message);
  process.exit(1);
});
