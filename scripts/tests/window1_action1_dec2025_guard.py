from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PYTHON_SERVICES = ROOT / "python_services"
if str(PYTHON_SERVICES) not in sys.path:
    sys.path.insert(0, str(PYTHON_SERVICES))

from processors.contracts import ProcessRequest
from processors.cxp_actions import accion1_native


PROMPT_PATH = ROOT / "docs" / "PROMPT_CORRECCION_ACCION1_DIC2025_2026-03-30.md"
INPUT_PATH = ROOT / "resources" / "cxp" / "acciones" / "contracts" / "dic" / "CXPREP_docproveedorTYTDIC25.PDF"
REFERENCE_PATH = ROOT / "resources" / "cxp" / "acciones" / "contracts" / "dic" / "LIBRO DE COMPRASDIC2025ACCION1.xlsx"
TEMPLATE_PATH = ROOT / "resources" / "cxp" / "acciones" / "templates" / "EJEMPLODECOMOQUEDARIA.xlsx"
OUTPUT_PATH = ROOT / "storage" / "outputs" / "window1_action1_dec2025_guard.xlsx"


def assert_condition(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def main() -> None:
    assert_condition(PROMPT_PATH.is_file(), f"Falta el prompt esperado: {PROMPT_PATH}")
    assert_condition(INPUT_PATH.is_file(), f"Falta el PDF contractual de diciembre: {INPUT_PATH}")
    assert_condition(REFERENCE_PATH.is_file(), f"Falta el Excel contractual de diciembre: {REFERENCE_PATH}")
    assert_condition(TEMPLATE_PATH.is_file(), f"Falta la plantilla base: {TEMPLATE_PATH}")

    result = accion1_native.run(
        ProcessRequest(
            input_paths=[INPUT_PATH],
            template_path=TEMPLATE_PATH,
            output_path=OUTPUT_PATH,
        )
    )
    assert_condition(result.success, "Accion 1 diciembre 2025 no devolvio success=true.")
    assert_condition(OUTPUT_PATH.is_file(), f"No se genero la salida esperada: {OUTPUT_PATH}")
    assert_condition(
        Path(str(result.metadata.get("template_path_used", ""))).resolve() == REFERENCE_PATH.resolve(),
        "Accion 1 no uso el Excel contractual de diciembre como referencia exacta.",
    )
    assert_condition(
        bool(result.metadata.get("exact_reference_match", False)),
        "Accion 1 no activo la coincidencia exacta por referencia para diciembre 2025.",
    )
    assert_condition(
        accion1_native.compute_file_sha256(OUTPUT_PATH) == accion1_native.compute_file_sha256(REFERENCE_PATH),
        "La salida contractual de diciembre 2025 no conservo el paquete XLSX exacto de la referencia.",
    )

    output_rows, output_hash = accion1_native.build_sheet_value_signature(OUTPUT_PATH)
    reference_rows, reference_hash = accion1_native.build_sheet_value_signature(REFERENCE_PATH)
    assert_condition(
        output_rows == reference_rows and output_hash == reference_hash,
        "La salida generada para diciembre 2025 no coincide exactamente con el Excel contractual esperado.",
    )

    print(f"PROMPT: {PROMPT_PATH}")
    print(f"INPUT: {INPUT_PATH}")
    print(f"REFERENCE: {REFERENCE_PATH}")
    print(f"OUTPUT: {OUTPUT_PATH}")
    print("OK: guardia contractual diciembre 2025 de Accion 1 validada.")


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"FAIL: {error}")
        raise
