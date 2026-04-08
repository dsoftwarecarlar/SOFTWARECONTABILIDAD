const { spawnSync } = require("child_process");
const path = require("path");

const ROOT = path.resolve(__dirname, "..", "..");
const PYTHON = process.env.PYTHON_BINARY || "python";

function assertCondition(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

const pythonCode = `
from python_services.processors.servicios_marcas.runtime import compact_mayor_rows_for_capacity

rows = [
    {
        "account": "04.01.01.12.0001",
        "name": "VTAS SERV CHANGAN - CONTADO CON IVA",
        "ext": "M",
        "date_value": 46022.0,
        "date_text": "31-DEC-25",
        "origin": "VENSE",
        "seat": "392",
        "reference": "",
        "detail": "CONT. VENTAS - CENTRO05 PERIODO 2025 - 12",
        "debit": 0.0,
        "credit": value,
        "balance": -100000.0 - index,
        "effective_balance": -100000.0 - index,
    }
    for index, value in enumerate([5231.06, 19.5, 103.0, 1710.26], start=1)
] + [
    {
        "account": "04.01.01.12.0001",
        "name": "VTAS SERV CHANGAN - CONTADO CON IVA",
        "ext": "M",
        "date_value": 46022.0,
        "date_text": "31-DEC-25",
        "origin": "VENSE",
        "seat": "394",
        "reference": "",
        "detail": "CONT. VENTAS - CENTRO08 PERIODO 2025 - 12",
        "debit": 0.0,
        "credit": value,
        "balance": -100010.0 - index,
        "effective_balance": -100010.0 - index,
    }
    for index, value in enumerate([46.46, 456.69, 105.35], start=1)
]

compacted = compact_mayor_rows_for_capacity(rows, 6)
assert len(compacted) == 2, f"expected 2 compacted rows, got {len(compacted)}"

credits = sorted(round(float(row["credit"]), 2) for row in compacted)
assert credits == [608.5, 7063.82], f"unexpected compacted credits: {credits}"
print("OK")
`;

const result = spawnSync(PYTHON, ["-c", pythonCode], {
  cwd: ROOT,
  encoding: "utf8",
});

assertCondition(result.status === 0, `La compactacion del mayor fallo.\n${result.stderr || result.stdout}`);
assertCondition((result.stdout || "").includes("OK"), "La prueba de compactacion del mayor no confirmo OK.");
console.log("OK: compactacion del mayor validada.");
