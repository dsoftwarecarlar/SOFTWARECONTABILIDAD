const path = require("path");

const DEFAULT_INPUT_SOURCE = "CON_MAYORGEN2IVAACCION4.txt";
const DEFAULT_OUTPUT_XLSX = "mayor_iva_accion4.xlsx";
const DEFAULT_TEMPLATE_XLSX = path.resolve(__dirname, "..", "..", "..", "outputs", "EJEMPLOSAMANO", "MAYORIVAACCION4.xlsx");
const SHEET_NAME = "MAYOR IVA";

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

module.exports = {
  DEFAULT_INPUT_SOURCE,
  DEFAULT_OUTPUT_XLSX,
  DEFAULT_TEMPLATE_XLSX,
  SHEET_NAME,
  EXPECTED_HEADERS,
  MOVEMENT_BOUNDARIES,
};
