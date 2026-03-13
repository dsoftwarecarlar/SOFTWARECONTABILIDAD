const COLUMN_ORDER = [
  "CODIGO",
  "CEDULA",
  "NOMBRE",
  "FECHA",
  "TIPO",
  "DOCUMENTO",
  "MONTO",
  "BASE IVA",
  "BASE 0",
  "IMPUESTOS",
  "RETENCION",
  "SALDO",
];

const HEADER_ALIASES = {
  CODIGO: "CODIGO",
  CEDULA: "CEDULA",
  NOMBRE: "NOMBRE",
  FECHA: "FECHA",
  TIPO: "TIPO",
  DOCUMENTO: "DOCUMENTO",
  MONTO: "MONTO",
  "BASE IVA": "BASE IVA",
  "BASE 0": "BASE 0",
  IMPUESTOS: "IMPUESTOS",
  RETENCION: "RETENCION",
  SALDO: "SALDO",
};

const FIXED_BOUNDARIES = [
  { name: "CODIGO", left: Number.NEGATIVE_INFINITY, right: 45 },
  { name: "CEDULA", left: 45, right: 140 },
  { name: "NOMBRE", left: 140, right: 245 },
  { name: "FECHA", left: 245, right: 307 },
  { name: "TIPO", left: 307, right: 334 },
  { name: "DOCUMENTO", left: 334, right: 400 },
  { name: "MONTO", left: 400, right: 475 },
  { name: "BASE IVA", left: 475, right: 545 },
  { name: "BASE 0", left: 545, right: 610 },
  { name: "IMPUESTOS", left: 610, right: 675 },
  { name: "RETENCION", left: 675, right: 739 },
  { name: "SALDO", left: 739, right: Number.POSITIVE_INFINITY },
];

const LEFT_BOUNDARIES = FIXED_BOUNDARIES.filter((boundary) =>
  ["CODIGO", "CEDULA", "NOMBRE", "FECHA", "TIPO", "DOCUMENTO"].includes(boundary.name),
);

const NUMERIC_COLUMNS = ["MONTO", "BASE IVA", "BASE 0", "IMPUESTOS", "RETENCION", "SALDO"];

const SPECIAL_PLAN = { codigo: "306", doc: "108173481", tipo: "FE", note: "PLAN EMPLEADOS" };
const SPECIAL_ACTIVO = { codigo: "3300", doc: "413", tipo: "FE", note: "ACTIVO FIJO-MUEBLES Y ENSERES" };
const SPECIAL_FE_IN_ND = { codigo: "150", doc: "A017835055", tipo: "FE" };

module.exports = {
  COLUMN_ORDER,
  HEADER_ALIASES,
  FIXED_BOUNDARIES,
  LEFT_BOUNDARIES,
  NUMERIC_COLUMNS,
  SPECIAL_PLAN,
  SPECIAL_ACTIVO,
  SPECIAL_FE_IN_ND,
};
