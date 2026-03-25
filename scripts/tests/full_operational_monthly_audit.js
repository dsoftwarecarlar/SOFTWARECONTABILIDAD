const path = require("path");
const { spawnSync } = require("child_process");

const ROOT = path.resolve(__dirname, "..", "..");
const NODE = process.execPath;

const STEPS = [
  { label: "Smoke UI", args: ["scripts/tests/http_resources_smoke.js"] },
  { label: "Contrato Accion 1 y 4", args: ["scripts/tests/e2e_action1_action4_contract.js"] },
  { label: "Contrato Accion 2 y 3", args: ["scripts/tests/e2e_action2_action3_contract.js"] },
  { label: "Auditoria mensual ventana 1", args: ["scripts/tests/window1_monthly_template_audit.js"] },
  { label: "Contrato Repuestos", args: ["scripts/tests/e2e_repuestos_contract.js"] },
  { label: "Paridad completa Repuestos", args: ["scripts/tests/repuestos_python_full_parity.js"] },
  { label: "Guardrail layouts Servicios", args: ["scripts/tests/servicios_layout_guard.js"] },
  { label: "Contrato Servicios", args: ["scripts/tests/e2e_servicios_marcas_contract.js"] },
  { label: "Guardrail rendimiento Servicios", args: ["scripts/tests/servicios_performance_guard.js"] },
  { label: "Paridad readers Servicios", args: ["scripts/tests/servicios_readers_python_parity.js"] },
  { label: "Quality gate contable", args: ["scripts/tests/contable_quality_gate.js"] },
];

for (const step of STEPS) {
  const completed = spawnSync(NODE, step.args, {
    cwd: ROOT,
    stdio: "inherit",
    windowsHide: true,
    timeout: 60 * 60 * 1000,
  });

  if (completed.error) {
    throw completed.error;
  }
  if (completed.status !== 0) {
    throw new Error(`Fallo la etapa: ${step.label}`);
  }
}

console.log("OK: auditoria operativa mensual completa validada.");
