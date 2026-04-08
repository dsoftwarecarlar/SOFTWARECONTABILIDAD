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
from python_services.processors.servicios_marcas.runtime import resolve_mayor_compatible_layout

layouts = [
    {
        "Account": "04.01.01.12.0001",
        "Name": "VTAS SERV CHANGAN - CONTADO CON IVA",
        "StartRow": 6,
        "EndRow": 11,
        "ParentStartRow": 6,
        "ParentEndRow": 11,
        "BucketHint": "sales",
    },
    {
        "Account": "04.01.01.12.0003",
        "Name": "VTAS SERV CHANGAN - CREDITO CON IVA",
        "StartRow": 12,
        "EndRow": 17,
        "ParentStartRow": 12,
        "ParentEndRow": 17,
        "BucketHint": "sales",
    },
    {
        "Account": "04.01.01.12.0010",
        "Name": "DESC VTAS SERV CHANGAN CONTADO CON IVA",
        "StartRow": 174,
        "EndRow": 177,
        "ParentStartRow": 174,
        "ParentEndRow": 177,
        "BucketHint": "discount",
    },
    {
        "Account": "04.01.01.12.0012",
        "Name": "DESC VTAS SERV CHANGAN CREDITO CON IVA",
        "StartRow": 247,
        "EndRow": 251,
        "ParentStartRow": 247,
        "ParentEndRow": 251,
        "BucketHint": "discount",
    },
    {
        "Account": "__AUTO_RETURN__355_434",
        "Name": "AUTO DEVOLUCIONES",
        "StartRow": 355,
        "EndRow": 434,
        "ParentStartRow": 355,
        "ParentEndRow": 434,
        "BucketHint": "return",
    },
]

selected = resolve_mayor_compatible_layout(
    "04.01.01.12.0014",
    "DEVOL VTAS SERV CHANGAN CONTADO",
    layouts,
    {},
    allow_cross_prefix_family=False,
)
assert selected is not None, "expected a compatible layout for 0014"
assert selected["Account"] == "__AUTO_RETURN__355_434", selected
print("OK")
`;

const result = spawnSync(PYTHON, ["-c", pythonCode], {
  cwd: ROOT,
  encoding: "utf8",
});

assertCondition(result.status === 0, `El mapeo placeholder de devoluciones fallo.\n${result.stderr || result.stdout}`);
assertCondition((result.stdout || "").includes("OK"), "La prueba de placeholder de devoluciones no confirmo OK.");
console.log("OK: placeholder de devoluciones del mayor validado.");
