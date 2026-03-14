const path = require("path");
const fs = require("fs");

const DEFAULT_INPUT_SOURCE = "ACCION2.txt";
const DEFAULT_OUTPUT_XLSX = "retenciones_proveedor.xlsx";
const TEMPLATE_CANDIDATES = [
  path.resolve(__dirname, "..", "..", "..", "outputs", "EJEMPLOSAMANO", "ACCION2.xlsx"),
  path.resolve(__dirname, "..", "..", "..", "outputs", "EJEMPLOSAMANO1", "ACCION2.xlsx"),
];
const DEFAULT_TEMPLATE_XLSX = TEMPLATE_CANDIDATES.find((candidate) => fs.existsSync(candidate)) || TEMPLATE_CANDIDATES[0];
const SHEET_NAME = "RET PROV";

const EXPECTED_COLUMNS = [
  "NUM RT",
  "PROVEEDOR",
  "FECHA",
  "FECHA CONT",
  "TIPO",
  "COD",
  "FACT",
  "%",
  "BASE",
  "RETENCION",
];

const HEADER_ALIASES = {
  NUMRT: "NUM RT",
  NUMERORT: "NUM RT",
  NUMRET: "NUM RT",
  NUMRETENCION: "NUM RT",
  PROVEEDOR: "PROVEEDOR",
  RAZONSOCIAL: "PROVEEDOR",
  NOMBREPROVEEDOR: "PROVEEDOR",
  FECHA: "FECHA",
  FECHADOC: "FECHA CONT",
  FECHADOCUMENTO: "FECHA CONT",
  FECHADOCTO: "FECHA CONT",
  FECHADOCU: "FECHA CONT",
  FECHACONT: "FECHA CONT",
  FECHACONTABLE: "FECHA CONT",
  TIPO: "TIPO",
  COD: "COD",
  CODIGO: "COD",
  TRANS: "COD",
  TRANSACCION: "COD",
  FACT: "FACT",
  FACTURA: "FACT",
  PORCENTAJE: "%",
  "%": "%",
  BASE: "BASE",
  BASERET: "BASE",
  BASEIVA: "BASE",
  VALORRETEN: "RETENCION",
  VALORRETENCION: "RETENCION",
  RETENCION: "RETENCION",
};

module.exports = {
  DEFAULT_INPUT_SOURCE,
  DEFAULT_OUTPUT_XLSX,
  DEFAULT_TEMPLATE_XLSX,
  SHEET_NAME,
  EXPECTED_COLUMNS,
  HEADER_ALIASES,
};
