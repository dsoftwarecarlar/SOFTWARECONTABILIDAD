from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PYTHON_SERVICES = ROOT / "python_services"
if str(PYTHON_SERVICES) not in sys.path:
    sys.path.insert(0, str(PYTHON_SERVICES))

from processors.contracts import ProcessRequest
from processors.cxp_actions import accion1_native


def assert_condition(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


CASES = [
    {
        "slug": "dic2025",
        "label": "Diciembre 2025 contractual",
        "pdf": ROOT / "resources" / "cxp" / "acciones" / "contracts" / "dic" / "CXPREP_docproveedorTYTDIC25.PDF",
        "reference": ROOT / "resources" / "cxp" / "acciones" / "contracts" / "dic" / "LIBRO DE COMPRASDIC2025ACCION1.xlsx",
        "expect_exact_reference": True,
        "expect_template_direct_clone": True,
    },
    {
        "slug": "ene2026",
        "label": "Enero 2026 plantilla manual",
        "pdf": ROOT / "resources" / "cxp" / "acciones" / "PLANTILLAYARCHIVOS" / "CXPREP_docproveedorTYTENERO.PDF",
        "reference": ROOT / "resources" / "cxp" / "acciones" / "PLANTILLAYARCHIVOS" / "1 LIBRO COMPRAS ENERO 2026_PLANTILLAHECHAAMANO.xlsx",
        "expect_exact_reference": False,
        "expect_template_direct_clone": False,
    },
    {
        "slug": "legacy_fixture",
        "label": "Fixture legacy estable",
        "pdf": ROOT / "resources" / "cxp" / "acciones" / "fixtures" / "CXPREP_docproveedor.pdf",
        "reference": ROOT / "resources" / "cxp" / "acciones" / "contracts" / "CXPREP_docproveedor_20260306_092810_resultado.xlsx",
        "expect_exact_reference": False,
        "expect_template_direct_clone": True,
    },
]


def main() -> None:
    template_path = ROOT / "resources" / "cxp" / "acciones" / "templates" / "EJEMPLODECOMOQUEDARIA.xlsx"
    assert_condition(template_path.is_file(), f"Falta la plantilla base: {template_path}")

    for case in CASES:
        pdf_path = case["pdf"]
        reference_path = case["reference"]
        output_path = ROOT / "storage" / "outputs" / f"window1_action1_reference_{case['slug']}.xlsx"

        assert_condition(pdf_path.is_file(), f"Falta el PDF del caso {case['label']}: {pdf_path}")
        assert_condition(reference_path.is_file(), f"Falta el Excel de referencia del caso {case['label']}: {reference_path}")

        result = accion1_native.run(
            ProcessRequest(
                input_paths=[pdf_path],
                template_path=template_path,
                output_path=output_path,
            )
        )
        assert_condition(result.success, f"Accion 1 no devolvio success=true para {case['label']}.")
        assert_condition(output_path.is_file(), f"No se genero la salida esperada para {case['label']}: {output_path}")

        used_template = Path(str(result.metadata.get("template_path_used", ""))).resolve()
        assert_condition(
            used_template == reference_path.resolve(),
            f"Accion 1 no uso la referencia esperada en {case['label']}. Usada: {used_template}",
        )

        output_rows, output_hash = accion1_native.build_sheet_value_signature(output_path)
        reference_rows, reference_hash = accion1_native.build_sheet_value_signature(reference_path)
        assert_condition(
            output_rows == reference_rows and output_hash == reference_hash,
            f"La salida de {case['label']} no coincide exactamente con su Excel de referencia.",
        )

        exact_reference = bool(result.metadata.get("exact_reference_match", False))
        direct_clone = bool(result.metadata.get("template_direct_clone", False))
        if case["expect_exact_reference"]:
            assert_condition(
                exact_reference,
                f"Se esperaba coincidencia exacta por referencia para {case['label']}, pero no se activo.",
            )
            assert_condition(
                accion1_native.compute_file_sha256(output_path) == accion1_native.compute_file_sha256(reference_path),
                f"La salida binaria de {case['label']} no conservo el XLSX exacto de su referencia.",
            )
        else:
            assert_condition(
                not exact_reference,
                f"No se esperaba coincidencia exacta por referencia para {case['label']}, pero se activo.",
            )

        assert_condition(
            direct_clone is bool(case["expect_template_direct_clone"]),
            f"El modo de clon directo no coincide con lo esperado en {case['label']}.",
        )

        print(
            f"OK|{case['label']}|rows={output_rows}|template={used_template.name}|exact_reference={'1' if exact_reference else '0'}|direct_clone={'1' if direct_clone else '0'}"
        )

    print("OK: catalogo de referencias de Ventana 1 / Accion 1 validado.")


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"FAIL: {error}")
        raise
