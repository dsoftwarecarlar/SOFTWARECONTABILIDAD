const path = require("path");
const { spawnSync } = require("child_process");

const ROOT = path.resolve(__dirname, "..", "..");
const NODE = process.execPath;
const POWERSHELL = "powershell.exe";

function runStep(step) {
  const command = step.kind === "node" ? NODE : POWERSHELL;
  const args = step.kind === "node"
    ? [step.script]
    : ["-ExecutionPolicy", "Bypass", "-File", step.script];

  console.log(`\n==> ${step.label}`);
  const completed = spawnSync(command, args, {
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

const STEPS = [
  { label: "Smoke UI", kind: "node", script: "scripts/tests/http_resources_smoke.js" },
  { label: "Smoke Laravel portal", kind: "ps1", script: "scripts/dev/test_laravel_app.ps1" },
  { label: "Ventana 1 Laravel", kind: "ps1", script: "scripts/dev/test_laravel_window1.ps1" },
  { label: "Repuestos Laravel", kind: "ps1", script: "scripts/dev/test_laravel_repuestos.ps1" },
  { label: "Servicios Laravel", kind: "ps1", script: "scripts/dev/test_laravel_servicios.ps1" },
  { label: "Contrato Accion 1 y 4", kind: "node", script: "scripts/tests/e2e_action1_action4_contract.js" },
  { label: "Contrato Accion 2 y 3", kind: "node", script: "scripts/tests/e2e_action2_action3_contract.js" },
  { label: "Auditoria plantilla Ventana 1", kind: "node", script: "scripts/tests/window1_monthly_template_audit.js" },
  { label: "Contrato Repuestos", kind: "node", script: "scripts/tests/e2e_repuestos_contract.js" },
  { label: "Paridad Repuestos REP", kind: "node", script: "scripts/tests/repuestos_rep_stage_python_parity.js" },
  { label: "Paridad Repuestos NC", kind: "node", script: "scripts/tests/repuestos_nc_stage_python_parity.js" },
  { label: "Paridad Repuestos MY", kind: "node", script: "scripts/tests/repuestos_my_stage_python_parity.js" },
  { label: "Paridad Repuestos MAYOR IVA", kind: "node", script: "scripts/tests/repuestos_mayor_iva_stage_python_parity.js" },
  { label: "Paridad total Repuestos", kind: "node", script: "scripts/tests/repuestos_python_full_parity.js" },
  { label: "Guardrail layouts Servicios", kind: "node", script: "scripts/tests/servicios_layout_guard.js" },
  { label: "Contrato Servicios", kind: "node", script: "scripts/tests/e2e_servicios_marcas_contract.js" },
  { label: "Guardrail rendimiento Servicios", kind: "node", script: "scripts/tests/servicios_performance_guard.js" },
  { label: "Paridad readers Servicios", kind: "node", script: "scripts/tests/servicios_readers_python_parity.js" },
  { label: "Quality gate contable", kind: "node", script: "scripts/tests/contable_quality_gate.js" },
  { label: "Auditoria operativa mensual completa", kind: "node", script: "scripts/tests/full_operational_monthly_audit.js" },
];

for (const step of STEPS) {
  runStep(step);
}

console.log("\nOK: auditoria definitiva total contable validada.");
