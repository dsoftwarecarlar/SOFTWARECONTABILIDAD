const { spawn } = require("child_process");

const ROOT = "C:/xampp/htdocs/SOFTWARECONTABILIDAD";

const STEPS = [
  {
    name: "UI smoke",
    command: "node",
    args: ["scripts/tests/http_resources_smoke.js"],
  },
  {
    name: "E2E accion2-3",
    command: "node",
    args: ["scripts/tests/e2e_action2_action3_contract.js"],
  },
  {
    name: "E2E accion1-4",
    command: "node",
    args: ["scripts/tests/e2e_action1_action4_contract.js"],
  },
  {
    name: "E2E repuestos",
    command: "node",
    args: ["scripts/tests/e2e_repuestos_contract.js"],
  },
];

function runStep(step) {
  return new Promise((resolve, reject) => {
    const child = spawn(step.command, step.args, {
      cwd: ROOT,
      stdio: "inherit",
      shell: false,
    });

    child.on("error", reject);
    child.on("exit", (code) => {
      if (code === 0) {
        resolve();
        return;
      }
      reject(new Error(`${step.name} fallo con exit code ${code}.`));
    });
  });
}

async function main() {
  const startedAt = Date.now();
  for (const step of STEPS) {
    console.log(`\n==> ${step.name}`);
    await runStep(step);
  }
  const elapsedMs = Date.now() - startedAt;
  console.log(`\nOK: quality gate contable completado (${elapsedMs} ms).`);
}

main().catch((error) => {
  console.error(`FAIL: ${error.message}`);
  process.exit(1);
});
