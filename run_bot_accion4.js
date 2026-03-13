const { main } = require("./scripts/cxp/accion4/process");

module.exports = { main };

if (require.main === module) {
  main().catch((error) => {
    console.error("Error:", error.message);
    process.exit(1);
  });
}
