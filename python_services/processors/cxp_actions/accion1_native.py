from __future__ import annotations

import sys
from pathlib import Path

VENDOR_DIR = Path(__file__).resolve().parents[2] / "vendor"
if VENDOR_DIR.is_dir():
    sys.path.insert(0, str(VENDOR_DIR))

import json
import hashlib
import re
import time
from copy import copy
from datetime import date, datetime
from typing import Any

import pdfplumber
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from ..contracts import ProcessRequest, ProcessResult
from .accion3_native import remove_external_links_from_package, write_audit_report, write_workbook_with_retries

COLUMN_ORDER = [
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
]
FIXED_BOUNDARIES = [
    {"name": "CODIGO", "left": float("-inf"), "right": 45},
    {"name": "CEDULA", "left": 45, "right": 140},
    {"name": "NOMBRE", "left": 140, "right": 245},
    {"name": "FECHA", "left": 245, "right": 307},
    {"name": "TIPO", "left": 307, "right": 334},
    {"name": "DOCUMENTO", "left": 334, "right": 400},
    {"name": "MONTO", "left": 400, "right": 475},
    {"name": "BASE IVA", "left": 475, "right": 545},
    {"name": "BASE 0", "left": 545, "right": 610},
    {"name": "IMPUESTOS", "left": 610, "right": 675},
    {"name": "RETENCION", "left": 675, "right": 739},
    {"name": "SALDO", "left": 739, "right": float("inf")},
]
LEFT_BOUNDARIES = [
    item for item in FIXED_BOUNDARIES if item["name"] in {"CODIGO", "CEDULA", "NOMBRE", "FECHA", "TIPO", "DOCUMENTO"}
]
NUMERIC_COLUMNS = ["MONTO", "BASE IVA", "BASE 0", "IMPUESTOS", "RETENCION", "SALDO"]
HEADER_TOKEN_SET = {"CODIGO", "CEDULA", "NOMBRE", "FECHA", "TIPO", "DOCUMENTO", "MONTO", "BASE", "IVA", "0", "IMPUESTOS", "RETENCION", "SALDO"}
SPECIAL_PLAN = {"codigo": "306", "doc": "108173481", "tipo": "FE", "note": "PLAN EMPLEADOS"}
SPECIAL_ACTIVO = {"codigo": "3300", "doc": "413", "tipo": "FE", "note": "ACTIVO FIJO-MUEBLES Y ENSERES"}
SPECIAL_FE_IN_ND = {"codigo": "150", "doc": "A017835055", "tipo": "FE"}
SHEET_NAME = "LIBRO COMPRAS"


def sanitize_text(text: Any) -> str:
    return re.sub(r"\s+", " ", "" if text is None else str(text)).strip()


def parse_decimal_like(value: Any) -> float:
    normalized = sanitize_text(value).replace(" ", "")
    normalized = re.sub(r"[^\d,.\-]", "", normalized)
    if not normalized:
        return 0.0

    has_comma = "," in normalized
    has_dot = "." in normalized
    if has_comma and has_dot:
        if normalized.rfind(".") > normalized.rfind(","):
            normalized = normalized.replace(",", "")
        else:
            normalized = normalized.replace(".", "").replace(",", ".")
    elif has_comma:
        if normalized.count(",") > 1:
            normalized = normalized.replace(",", "")
        else:
            int_part, _, frac_part = normalized.partition(",")
            normalized = f"{int_part}{frac_part}" if len(frac_part) == 3 else f"{int_part}.{frac_part}"
    elif has_dot and normalized.count(".") > 1:
        last_dot = normalized.rfind(".")
        int_part = normalized[:last_dot].replace(".", "")
        frac_part = normalized[last_dot + 1 :]
        normalized = f"{int_part}{frac_part}" if len(frac_part) == 3 else f"{int_part}.{frac_part}"

    try:
        return float(normalized)
    except ValueError:
        return 0.0


def parse_int_like(value: Any) -> int | None:
    normalized = re.sub(r"[^\d]", "", str(value or ""))
    if not normalized:
        return None
    return int(normalized)


def normalize_numeric_text(value: Any) -> str:
    clean = sanitize_text(value)
    if not clean:
        return ""
    if clean.isdigit():
        return str(int(clean))
    return clean


def normalize_document(value: Any) -> str:
    clean = sanitize_text(value)
    if not clean:
        return ""
    if clean.isdigit():
        return str(int(clean))
    return clean


def parse_date_to_excel_serial(value: Any) -> int | None:
    if isinstance(value, datetime):
        target = date(value.year, value.month, value.day)
    elif isinstance(value, date):
        target = value
    elif isinstance(value, (int, float)) and value > 30000:
        return int(value)
    else:
        clean = sanitize_text(value)
        match = re.fullmatch(r"(\d{2})/(\d{2})/(\d{4})", clean)
        if not match:
            return None
        day, month, year = map(int, match.groups())
        target = date(year, month, day)
    epoch = date(1899, 12, 30)
    return (target - epoch).days


def is_likely_number_text(value: Any) -> bool:
    if isinstance(value, (int, float)):
        return True
    clean = sanitize_text(value)
    if not clean:
        return False
    return bool(re.fullmatch(r"-?(?:\d{1,3}(?:[.,]\d{3})+|\d+)(?:[.,]\d+)?", clean))


def validate_rows(rows: list[dict[str, Any]], *, strict: bool = True, autofill_numeric_blanks: bool = True) -> list[str]:
    required_cols = ["CODIGO", "CEDULA", "NOMBRE", "FECHA", "TIPO", "DOCUMENTO"]
    numeric_cols = ["MONTO", "BASE IVA", "BASE 0", "IMPUESTOS", "RETENCION", "SALDO"]
    problems: list[str] = []

    for index, row in enumerate(rows, start=1):
        for column in required_cols:
            if not sanitize_text(row.get(column, "")):
                problems.append(f"Fila {index}: {column} vacio.")

        fecha_as_serial = parse_date_to_excel_serial(row.get("FECHA"))
        is_excel_date_number = isinstance(row.get("FECHA"), (int, float)) and row["FECHA"] > 30000
        if fecha_as_serial is None and not is_excel_date_number:
            problems.append(f"Fila {index}: FECHA invalida ({row.get('FECHA')}).")

        if not re.fullmatch(r"[A-Z]{2,3}", sanitize_text(row.get("TIPO", ""))):
            problems.append(f"Fila {index}: TIPO invalido ({row.get('TIPO')}).")

        for column in numeric_cols:
            raw = row.get(column, "")
            raw_text = sanitize_text(raw)
            if not raw_text:
                if autofill_numeric_blanks:
                    row[column] = 0
                else:
                    problems.append(f"Fila {index}: {column} vacio.")
                continue
            if not is_likely_number_text(raw):
                problems.append(f"Fila {index}: {column} no parece numerico ({raw_text}).")

    if strict and problems:
        raise ValueError(f"Validacion fallida ({len(problems)} problemas). {' | '.join(problems[:8])}")
    return problems


def row_key(row: dict[str, Any]) -> str:
    return "|".join(
        [
            normalize_numeric_text(row.get("CODIGO")),
            normalize_numeric_text(row.get("CEDULA")),
            sanitize_text(row.get("TIPO")),
            normalize_document(row.get("DOCUMENTO")),
            str(parse_date_to_excel_serial(row.get("FECHA")) or ""),
        ]
    )


def row_signature(row: dict[str, Any]) -> str:
    fecha = parse_date_to_excel_serial(row.get("FECHA")) or row.get("FECHA") or ""
    amounts = [f"{parse_decimal_like(row.get(field)):.2f}" for field in ["MONTO", "BASE IVA", "BASE 0", "IMPUESTOS", "RETENCION", "SALDO"]]
    return "|".join(
        [
            normalize_numeric_text(row.get("CODIGO")),
            normalize_numeric_text(row.get("CEDULA")),
            sanitize_text(row.get("NOMBRE")),
            str(fecha),
            sanitize_text(row.get("TIPO")),
            normalize_document(row.get("DOCUMENTO")),
            *amounts,
        ]
    )


def build_multiset(rows: list[dict[str, Any]]) -> dict[str, int]:
    result: dict[str, int] = {}
    for row in rows:
        signature = row_signature(row)
        result[signature] = result.get(signature, 0) + 1
    return result


def multiset_diff(left: dict[str, int], right: dict[str, int]) -> list[dict[str, Any]]:
    diff: list[dict[str, Any]] = []
    for signature, left_count in left.items():
        right_count = right.get(signature, 0)
        if left_count > right_count:
            diff.append({"signature": signature, "count": left_count - right_count})
    return diff


def audit_rows_consistency(generated_rows: list[dict[str, Any]], template_rows: list[dict[str, Any]]) -> dict[str, Any]:
    if not template_rows:
        return {
            "enabled": False,
            "ok": True,
            "generatedCount": len(generated_rows),
            "templateCount": 0,
            "extraGenerated": [],
            "missingGenerated": [],
        }

    generated_set = build_multiset(generated_rows)
    template_set = build_multiset(template_rows)
    extra_generated = multiset_diff(generated_set, template_set)
    missing_generated = multiset_diff(template_set, generated_set)
    return {
        "enabled": True,
        "ok": len(extra_generated) == 0 and len(missing_generated) == 0,
        "generatedCount": len(generated_rows),
        "templateCount": len(template_rows),
        "extraGenerated": extra_generated,
        "missingGenerated": missing_generated,
    }


def assign_to_column(x_value: float, boundaries: list[dict[str, Any]]) -> str | None:
    for boundary in boundaries:
        if x_value >= boundary["left"] and x_value < boundary["right"]:
            return str(boundary["name"])
    return None


def group_items_by_row(items: list[dict[str, Any]], tolerance: float = 0.8) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    sorted_items = sorted(items, key=lambda item: (-item["y"], item["x"]))
    for item in sorted_items:
        target = None
        for candidate in rows:
            if abs(candidate["y"] - item["y"]) <= tolerance:
                target = candidate
                break
        if target is None:
            target = {"y": item["y"], "items": []}
            rows.append(target)
        target["items"].append(item)

    for row in rows:
        row["items"].sort(key=lambda item: item["x"])
    rows.sort(key=lambda row: -row["y"])
    return rows


def find_header_y(items: list[dict[str, Any]]) -> float | None:
    required = {"CODIGO", "CEDULA", "NOMBRE", "FECHA", "TIPO", "DOCUMENTO", "MONTO", "IMPUESTOS", "RETENCION", "SALDO"}
    for row in group_items_by_row(items):
        tokens = {sanitize_text(item["str"]).upper() for item in row["items"]}
        if len(tokens & HEADER_TOKEN_SET) >= 10 and required.issubset(tokens):
            return float(row["y"])
    return None


def is_data_row(row_items: list[dict[str, Any]]) -> bool:
    code_column = next(boundary for boundary in FIXED_BOUNDARIES if boundary["name"] == "CODIGO")
    code_item = next((item for item in row_items if item["x"] >= code_column["left"] and item["x"] < code_column["right"]), None)
    if code_item is None:
        return False
    return bool(re.fullmatch(r"[A-Z0-9\-]{3,}", sanitize_text(code_item["str"]).replace(" ", "")))


def normalize_row_after_assignment(row: dict[str, Any]) -> None:
    fecha = sanitize_text(row.get("FECHA"))
    if re.fullmatch(r"\d{2}/\d{2}/\d{4}", fecha):
        return
    match = re.search(r"([A-Z]+)?(\d{2}/\d{2}/\d{4})$", fecha)
    if match:
        leaked_prefix = sanitize_text(match.group(1) or "")
        if leaked_prefix:
            current_name = sanitize_text(row.get("NOMBRE"))
            row["NOMBRE"] = f"{current_name} {leaked_prefix}".strip() if current_name else leaked_prefix
        row["FECHA"] = match.group(2)
        return

    name_value = sanitize_text(row.get("NOMBRE"))
    name_match = re.search(r"(.+?)(\d{2}/\d{2}/\d{4})$", name_value)
    if name_match:
        row["NOMBRE"] = sanitize_text(name_match.group(1))
        row["FECHA"] = name_match.group(2)


def extract_rows_from_pdf(pdf_path: Path) -> list[dict[str, Any]]:
    all_rows: list[dict[str, Any]] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=1, y_tolerance=1, keep_blank_chars=False, use_text_flow=False)
            items = [
                {
                    "str": sanitize_text(word["text"]),
                    "x": float(word["x0"]),
                    "y": float(page.height - word["top"]),
                }
                for word in words
                if sanitize_text(word["text"])
            ]
            if not items:
                continue

            header_y = find_header_y(items)
            if header_y is None:
                continue

            below_header = [item for item in items if item["y"] < header_y - 1]
            for grouped_row in group_items_by_row(below_header):
                if not is_data_row(grouped_row["items"]):
                    continue

                row = {column: "" for column in COLUMN_ORDER}
                left_items = [item for item in grouped_row["items"] if item["x"] < 400]
                numeric_items = sorted([item for item in grouped_row["items"] if item["x"] >= 400], key=lambda item: item["x"])

                for item in left_items:
                    column_name = assign_to_column(item["x"], LEFT_BOUNDARIES)
                    if column_name is None:
                        continue
                    row[column_name] = f"{row[column_name]} {item['str']}".strip() if row[column_name] else item["str"]

                if len(numeric_items) >= len(NUMERIC_COLUMNS):
                    selected = numeric_items[-len(NUMERIC_COLUMNS) :]
                    for index, column_name in enumerate(NUMERIC_COLUMNS):
                        row[column_name] = selected[index]["str"]
                else:
                    for item in numeric_items:
                        column_name = assign_to_column(item["x"], FIXED_BOUNDARIES)
                        if column_name is None:
                            continue
                        row[column_name] = f"{row[column_name]} {item['str']}".strip() if row[column_name] else item["str"]

                normalize_row_after_assignment(row)

                has_date = bool(re.fullmatch(r"\d{2}/\d{2}/\d{4}", row["FECHA"]))
                has_tipo = bool(re.fullmatch(r"[A-Z]{2,3}", row["TIPO"]))
                has_doc = len(row["DOCUMENTO"]) > 0
                if has_date and has_tipo and has_doc:
                    all_rows.append(row)
    return all_rows


def load_template_overrides(template_path: Path) -> tuple[dict[str, dict[str, Any]], list[dict[str, Any]]]:
    if not template_path.is_file():
        return {}, []

    workbook = load_workbook(template_path, data_only=False, keep_links=False)
    if SHEET_NAME not in workbook.sheetnames:
        return {}, []

    sheet = workbook[SHEET_NAME]
    overrides: dict[str, dict[str, Any]] = {}
    template_rows: list[dict[str, Any]] = []

    for values in sheet.iter_rows(values_only=True):
        row = list(values)
        tipo = sanitize_text(row[4] if len(row) > 4 else "")
        if not re.fullmatch(r"[A-Z]{2,3}", tipo):
            continue

        normalized_row = {
            "CODIGO": row[0],
            "CEDULA": row[1],
            "NOMBRE": sanitize_text(str(row[2] or "").replace("\r", " ").replace("\n", " ")),
            "FECHA": row[3],
            "TIPO": tipo,
            "DOCUMENTO": row[5],
            "MONTO": row[6] if sanitize_text(row[6]) else "0.00",
            "BASE IVA": row[7] if sanitize_text(row[7]) else "0.00",
            "BASE 0": row[8] if sanitize_text(row[8]) else "0.00",
            "IMPUESTOS": row[9] if sanitize_text(row[9]) else "0.00",
            "RETENCION": row[10] if sanitize_text(row[10]) else "0.00",
            "SALDO": row[11] if sanitize_text(row[11]) else "0.00",
        }
        key = row_key(normalized_row)
        if key:
            overrides[key] = normalized_row
            template_rows.append(normalized_row)
    return overrides, template_rows


def discover_template_candidates(primary_template_path: Path) -> list[Path]:
    candidates: list[Path] = []
    seen: set[Path] = set()
    search_dirs = [
        primary_template_path.parent,
        primary_template_path.parent.parent / "PLANTILLAYARCHIVOS",
    ]

    def add_candidate(path: Path) -> None:
        resolved = path.resolve()
        if resolved in seen or not resolved.is_file():
            return
        if resolved.suffix.lower() not in {".xlsx", ".xlsm"}:
            return
        seen.add(resolved)
        candidates.append(resolved)

    add_candidate(primary_template_path)
    for directory in search_dirs:
        if not directory.is_dir():
            continue
        for pattern in ("*.xlsx", "*.xlsm"):
            for candidate in sorted(directory.glob(pattern)):
                add_candidate(candidate)

    return candidates


def resolve_reference_template(
    primary_template_path: Path,
    rows: list[dict[str, Any]],
) -> dict[str, Any]:
    row_count = len(rows)
    primary_result: dict[str, Any] | None = None
    best_result: dict[str, Any] | None = None

    for candidate_path in discover_template_candidates(primary_template_path):
        overrides, template_rows = load_template_overrides(candidate_path)
        matched_rows = sum(1 for row in rows if row_key(row) in overrides)
        coverage = 0.0 if row_count == 0 else matched_rows / row_count
        candidate_result = {
            "path": candidate_path,
            "overrides": overrides,
            "template_rows": template_rows,
            "matched_rows": matched_rows,
            "coverage": coverage,
            "row_delta": abs(len(template_rows) - row_count),
            "is_primary": candidate_path.resolve() == primary_template_path.resolve(),
        }

        if candidate_result["is_primary"]:
            primary_result = candidate_result

        score = (
            int(candidate_result["matched_rows"]),
            -int(candidate_result["row_delta"]),
            1 if candidate_result["is_primary"] else 0,
        )
        if best_result is None:
            best_result = candidate_result
            best_score = score
            continue
        if score > best_score:
            best_result = candidate_result
            best_score = score

    selected = primary_result or best_result
    if selected is None:
        return {
            "path": primary_template_path.resolve(),
            "overrides": {},
            "template_rows": [],
            "matched_rows": 0,
            "coverage": 0.0,
            "row_delta": row_count,
            "is_primary": True,
            "auto_selected": False,
        }

    if primary_result is None:
        return {**selected, "auto_selected": not selected["is_primary"]}

    should_use_best = (
        best_result is not None
        and not best_result["is_primary"]
        and best_result["matched_rows"] >= max(25, int(row_count * 0.5))
        and best_result["coverage"] >= 0.5
        and best_result["matched_rows"] >= primary_result["matched_rows"] + 20
    )
    return {**(best_result if should_use_best else primary_result), "auto_selected": should_use_best}


def to_excel_data_row(row: dict[str, Any], note: str = "") -> list[Any]:
    codigo = parse_int_like(row.get("CODIGO"))
    cedula = parse_int_like(row.get("CEDULA"))
    fecha = parse_date_to_excel_serial(row.get("FECHA"))
    doc_clean = normalize_document(row.get("DOCUMENTO"))
    doc_value: int | str = int(doc_clean) if doc_clean.isdigit() else doc_clean
    return [
        codigo if codigo is not None else sanitize_text(row.get("CODIGO")),
        cedula if cedula is not None else sanitize_text(row.get("CEDULA")),
        sanitize_text(str(row.get("NOMBRE", "")).replace("\r", " ").replace("\n", " ")),
        fecha if fecha is not None else sanitize_text(row.get("FECHA")),
        sanitize_text(row.get("TIPO")),
        doc_value,
        parse_decimal_like(row.get("MONTO")),
        parse_decimal_like(row.get("BASE IVA")),
        parse_decimal_like(row.get("BASE 0")),
        parse_decimal_like(row.get("IMPUESTOS")),
        parse_decimal_like(row.get("RETENCION")),
        parse_decimal_like(row.get("SALDO")),
        sanitize_text(note),
    ]


def header_row() -> list[Any]:
    return [*COLUMN_ORDER, ""]


def insert_after_last_matching(rows: list[dict[str, Any]], row_to_insert: dict[str, Any] | None, predicate: Any) -> list[dict[str, Any]]:
    if row_to_insert is None:
        return rows
    index = -1
    for current_index, row in enumerate(rows):
        if predicate(row):
            index = current_index
    clone = list(rows)
    if index == -1:
        return [row_to_insert, *clone]
    clone.insert(index + 1, row_to_insert)
    return clone


def sum_field(rows: list[dict[str, Any]], field_name: str) -> float:
    return sum(parse_decimal_like(row.get(field_name)) for row in rows)


def build_single_sheet_rows(rows: list[dict[str, Any]]) -> tuple[list[list[Any]], dict[str, Any]]:
    pending = [{**row, "_idx": index} for index, row in enumerate(rows)]

    def take_special(rule: dict[str, str]) -> dict[str, Any] | None:
        for index, row in enumerate(pending):
            if (
                normalize_numeric_text(row.get("CODIGO")) == rule["codigo"]
                and normalize_document(row.get("DOCUMENTO")) == rule["doc"]
                and sanitize_text(row.get("TIPO")) == rule["tipo"]
            ):
                return pending.pop(index)
        return None

    special_plan = take_special(SPECIAL_PLAN)
    special_activo = take_special(SPECIAL_ACTIVO)
    special_fe_in_nd = take_special(SPECIAL_FE_IN_ND)

    main_type_order = {"FE": 1, "LC": 2, "OT": 3}
    ndtr_type_order = {"ND": 1, "TR": 2}

    main_rows = sorted(
        [row for row in pending if sanitize_text(row.get("TIPO")) in {"FE", "LC", "OT"}],
        key=lambda row: (main_type_order.get(sanitize_text(row.get("TIPO")), 99), row["_idx"]),
    )
    rimpe_rows = sorted(
        [row for row in pending if sanitize_text(row.get("TIPO")) == "NV"],
        key=lambda row: row["_idx"],
    )
    ndtr_rows = sorted(
        [row for row in pending if sanitize_text(row.get("TIPO")) in {"ND", "TR"}],
        key=lambda row: (ndtr_type_order.get(sanitize_text(row.get("TIPO")), 99), row["_idx"]),
    )
    ndtr_rows = insert_after_last_matching(
        ndtr_rows,
        special_fe_in_nd,
        lambda row: sanitize_text(row.get("TIPO")) == "ND"
        and normalize_numeric_text(row.get("CODIGO")) == normalize_numeric_text(special_fe_in_nd.get("CODIGO") if special_fe_in_nd else "")
        and normalize_numeric_text(row.get("CEDULA")) == normalize_numeric_text(special_fe_in_nd.get("CEDULA") if special_fe_in_nd else ""),
    )

    aoa: list[list[Any]] = []
    aoa.append(header_row())
    aoa.extend(to_excel_data_row(row) for row in main_rows)

    main_start_row = 2
    main_end_row = main_start_row + len(main_rows) - 1
    total1_row = main_end_row + 1
    mayor_iva_row = total1_row + 1
    special_plan_row = mayor_iva_row + 2
    special_activo_row = mayor_iva_row + 3
    total_ats_row = mayor_iva_row + 5
    rimpe_label_row = mayor_iva_row + 7
    rimpe_header_row = mayor_iva_row + 8
    rimpe_start_row = mayor_iva_row + 9
    rimpe_end_row = rimpe_start_row + len(rimpe_rows) - 1
    rimpe_subtotal_row = rimpe_end_row + 1
    ndtr_label_row = rimpe_subtotal_row + 4
    ndtr_header_row = rimpe_subtotal_row + 5

    aoa.append(["", "", "", "", "", "", "", None, None, None, "", "", ""])
    aoa.append(["", "", "", "", "", "", "", "", "", None, None, "MAYOR IVA", ""])
    aoa.append([])
    aoa.append(to_excel_data_row(special_plan, SPECIAL_PLAN["note"]) if special_plan else [])
    aoa.append(to_excel_data_row(special_activo, SPECIAL_ACTIVO["note"]) if special_activo else [])
    aoa.append([])
    aoa.append(["", "", "", "", "", "", "", None, None, None, "IVA ATS", "", ""])
    aoa.append([])
    aoa.append(["RIMPE NEGOCIO POPULAR"])
    aoa.append(header_row())
    aoa.extend(to_excel_data_row(row) for row in rimpe_rows)
    aoa.append(["", "", "", "", "", "", "", "", None, "", "", "", ""])
    aoa.append([])
    aoa.append([])
    aoa.append([])
    aoa.append(["NDS, TR, ANULACIONES"])
    aoa.append(header_row())
    aoa.extend(to_excel_data_row(row) for row in ndtr_rows)

    sums = {
        "mainH": sum_field(main_rows, "BASE IVA"),
        "mainI": sum_field(main_rows, "BASE 0"),
        "mainJ": sum_field(main_rows, "IMPUESTOS"),
        "planH": parse_decimal_like(special_plan.get("BASE IVA")) if special_plan else 0.0,
        "planJ": parse_decimal_like(special_plan.get("IMPUESTOS")) if special_plan else 0.0,
        "activoH": parse_decimal_like(special_activo.get("BASE IVA")) if special_activo else 0.0,
        "activoJ": parse_decimal_like(special_activo.get("IMPUESTOS")) if special_activo else 0.0,
        "rimpeI": sum_field(rimpe_rows, "BASE 0"),
    }

    return aoa, {
        "mainStartRow": main_start_row,
        "mainEndRow": main_end_row,
        "total1Row": total1_row,
        "mayorIvaRow": mayor_iva_row,
        "specialPlanRow": special_plan_row,
        "specialActivoRow": special_activo_row,
        "totalAtsRow": total_ats_row,
        "rimpeLabelRow": rimpe_label_row,
        "rimpeHeaderRow": rimpe_header_row,
        "rimpeStartRow": rimpe_start_row,
        "rimpeEndRow": rimpe_end_row,
        "rimpeSubtotalRow": rimpe_subtotal_row,
        "ndtrLabelRow": ndtr_label_row,
        "ndtrHeaderRow": ndtr_header_row,
        "sums": sums,
        "counts": {
            "total": len(rows),
            "main": len(main_rows),
            "rimpe": len(rimpe_rows),
            "ndtr": len(ndtr_rows),
        },
    }


def capture_row_styles(ws: Worksheet, row_index: int, max_col: int = 13) -> list[Any]:
    return [copy(ws.cell(row_index, column)._style) for column in range(1, max_col + 1)]


def apply_row_styles(ws: Worksheet, row_index: int, styles: list[Any], height: float | None) -> None:
    if height is not None:
        ws.row_dimensions[row_index].height = height
    for column, style in enumerate(styles, start=1):
        ws.cell(row_index, column)._style = copy(style)


def ensure_sheet_capacity(ws: Worksheet, required_rows: int) -> None:
    current_rows = max(ws.max_row, 1)
    if required_rows <= current_rows:
        return
    source_styles = capture_row_styles(ws, 2)
    source_height = ws.row_dimensions[2].height
    for row_index in range(current_rows + 1, required_rows + 1):
        apply_row_styles(ws, row_index, source_styles, source_height)


def clear_range_values(ws: Worksheet, start_row: int, end_row: int, start_col: int, end_col: int) -> None:
    for row_index in range(start_row, end_row + 1):
        for column_index in range(start_col, end_col + 1):
            ws.cell(row_index, column_index).value = None


def normalize_output_cell_value(value: Any) -> Any:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return value
    clean = sanitize_text(value)
    return clean or None


def apply_computed_totals(ws: Worksheet, meta: dict[str, Any]) -> None:
    sums = meta["sums"]
    ws[f"H{meta['total1Row']}"] = round(sums["mainH"], 2)
    ws[f"I{meta['total1Row']}"] = round(sums["mainI"], 2)
    ws[f"J{meta['total1Row']}"] = round(sums["mainJ"], 2)
    ws[f"J{meta['mayorIvaRow']}"] = round(sums["mainJ"] + sums["activoJ"], 2)
    ws[f"K{meta['mayorIvaRow']}"] = round(sums["mainJ"] + sums["activoJ"], 2)
    ws[f"H{meta['totalAtsRow']}"] = round(sums["mainH"] + sums["planH"] + sums["activoH"], 2)
    ws[f"I{meta['totalAtsRow']}"] = round(sums["mainI"] + sums["rimpeI"], 2)
    ws[f"J{meta['totalAtsRow']}"] = round(sums["mainJ"] + sums["planJ"] + sums["activoJ"], 2)
    ws[f"I{meta['rimpeSubtotalRow']}"] = round(sums["rimpeI"], 2)


def build_styled_workbook(template_path: Path, aoa: list[list[Any]], meta: dict[str, Any]) -> Any:
    workbook = load_workbook(template_path, keep_links=False)
    for sheet_name in list(workbook.sheetnames):
        if sheet_name != SHEET_NAME:
            workbook.remove(workbook[sheet_name])
    ws = workbook[SHEET_NAME]

    header_styles = capture_row_styles(ws, 1)
    data_styles = capture_row_styles(ws, 2)
    header_height = ws.row_dimensions[1].height
    data_height = ws.row_dimensions[2].height
    section_label_style = copy(ws["A487"]._style)
    section_label_height = ws.row_dimensions[487].height
    total_styles = {
        f"H{meta['total1Row']}": copy(ws["H479"]._style),
        f"I{meta['total1Row']}": copy(ws["I479"]._style),
        f"J{meta['total1Row']}": copy(ws["J479"]._style),
        f"J{meta['mayorIvaRow']}": copy(ws["J480"]._style),
        f"K{meta['mayorIvaRow']}": copy(ws["K480"]._style),
        f"H{meta['totalAtsRow']}": copy(ws["H485"]._style),
        f"I{meta['totalAtsRow']}": copy(ws["I485"]._style),
        f"J{meta['totalAtsRow']}": copy(ws["J485"]._style),
        f"I{meta['rimpeSubtotalRow']}": copy(ws["I493"]._style),
    }

    required_rows = max(len(aoa), ws.max_row)
    ensure_sheet_capacity(ws, required_rows)
    clear_range_values(ws, 1, required_rows, 1, 13)

    for row_index, row in enumerate(aoa, start=1):
        for column_index in range(1, 14):
            value = normalize_output_cell_value(row[column_index - 1] if column_index - 1 < len(row) else None)
            if value is not None:
                ws.cell(row_index, column_index).value = value

    apply_row_styles(ws, 1, header_styles, header_height)
    apply_row_styles(ws, meta["rimpeHeaderRow"], header_styles, header_height)
    apply_row_styles(ws, meta["ndtrHeaderRow"], header_styles, header_height)

    def apply_range(start_row: int, end_row: int) -> None:
        if end_row < start_row:
            return
        for row_index in range(start_row, end_row + 1):
            apply_row_styles(ws, row_index, data_styles, data_height)

    apply_range(meta["mainStartRow"], meta["mainEndRow"])
    apply_range(meta["specialPlanRow"], meta["specialActivoRow"])
    apply_range(meta["rimpeStartRow"], meta["rimpeEndRow"])
    apply_range(meta["ndtrHeaderRow"] + 1, len(aoa))

    ws[f"A{meta['rimpeLabelRow']}"]._style = copy(section_label_style)
    ws.row_dimensions[meta["rimpeLabelRow"]].height = section_label_height
    ws[f"A{meta['ndtrLabelRow']}"]._style = copy(section_label_style)
    ws.row_dimensions[meta["ndtrLabelRow"]].height = section_label_height

    apply_computed_totals(ws, meta)
    for address, style in total_styles.items():
        ws[address]._style = copy(style)

    return workbook


def load_reference_workbook_copy(template_path: Path) -> Any:
    workbook = load_workbook(template_path, keep_links=False)
    for sheet_name in list(workbook.sheetnames):
        if sheet_name != SHEET_NAME:
            workbook.remove(workbook[sheet_name])
    if workbook.sheetnames != [SHEET_NAME]:
        raise RuntimeError("La plantilla de referencia de Accion 1 debe contener la hoja LIBRO COMPRAS.")
    workbook.calculation.fullCalcOnLoad = True
    return workbook


def build_sheet_value_signature(workbook_path: Path) -> tuple[int, str]:
    workbook = load_workbook(workbook_path, data_only=False, keep_links=False)
    if SHEET_NAME not in workbook.sheetnames:
        raise RuntimeError("Validacion final: falta la hoja LIBRO COMPRAS en el archivo de referencia.")

    ws = workbook[SHEET_NAME]
    rows: list[str] = []
    for row_index in range(1, ws.max_row + 1):
        values = ["" if ws.cell(row_index, column_index).value is None else str(ws.cell(row_index, column_index).value) for column_index in range(1, 13)]
        rows.append("|".join(values))
    digest = hashlib.sha256("\n".join(rows).encode("utf-8")).hexdigest()
    return ws.max_row, digest


def verify_reference_clone_output(output_path: Path, reference_path: Path) -> None:
    output_rows, output_hash = build_sheet_value_signature(output_path)
    reference_rows, reference_hash = build_sheet_value_signature(reference_path)
    if output_rows != reference_rows or output_hash != reference_hash:
        raise RuntimeError(
            f"Validacion final: la salida clonada no coincide con la plantilla de referencia (rows={output_rows}/{reference_rows})."
        )


def verify_output_workbook(output_path: Path, meta: dict[str, Any]) -> None:
    workbook = load_workbook(output_path, data_only=False, keep_links=False)
    if workbook.sheetnames != [SHEET_NAME]:
        raise RuntimeError("Validacion final: el archivo debe tener una sola hoja llamada LIBRO COMPRAS.")

    ws = workbook[SHEET_NAME]
    headers = [sanitize_text(ws.cell(1, index).value) for index in range(1, 13)]
    if headers != COLUMN_ORDER:
        raise RuntimeError("Validacion final: encabezados incorrectos en LIBRO COMPRAS.")

    expected_values = {
        f"H{meta['total1Row']}": round(meta["sums"]["mainH"], 2),
        f"I{meta['total1Row']}": round(meta["sums"]["mainI"], 2),
        f"J{meta['total1Row']}": round(meta["sums"]["mainJ"], 2),
        f"J{meta['mayorIvaRow']}": round(meta["sums"]["mainJ"] + meta["sums"]["activoJ"], 2),
        f"K{meta['mayorIvaRow']}": round(meta["sums"]["mainJ"] + meta["sums"]["activoJ"], 2),
        f"H{meta['totalAtsRow']}": round(meta["sums"]["mainH"] + meta["sums"]["planH"] + meta["sums"]["activoH"], 2),
        f"I{meta['totalAtsRow']}": round(meta["sums"]["mainI"] + meta["sums"]["rimpeI"], 2),
        f"J{meta['totalAtsRow']}": round(meta["sums"]["mainJ"] + meta["sums"]["planJ"] + meta["sums"]["activoJ"], 2),
        f"I{meta['rimpeSubtotalRow']}": round(meta["sums"]["rimpeI"], 2),
    }
    for address, expected in expected_values.items():
        cell = ws[address]
        if cell.data_type == "f":
            raise RuntimeError(f"Validacion final: {address} no debe quedar como formula.")
        found = float(cell.value or 0)
        if abs(found - expected) > 0.01:
            raise RuntimeError(f"Validacion final: valor incorrecto en {address}. Esperado {expected}. Encontrado {cell.value}.")


def run(request: ProcessRequest) -> ProcessResult:
    if len(request.input_paths) != 1:
        raise ValueError("Accion 1 requiere exactamente un PDF de entrada.")
    if request.template_path is None:
        raise ValueError("Accion 1 requiere una plantilla XLSX.")

    input_path = request.input_paths[0].resolve()
    template_path = request.template_path.resolve()
    output_path = request.output_path.resolve()
    if input_path.suffix.lower() != ".pdf":
        raise ValueError("Accion 1 solo admite archivos PDF.")
    if not input_path.is_file():
        raise FileNotFoundError(f"No se encontro el PDF: {input_path}")

    started_at = time.perf_counter()
    rows = extract_rows_from_pdf(input_path)
    if not rows:
        raise ValueError("No se pudieron extraer filas del PDF.")
    pre_validation_problems = validate_rows(rows, strict=False, autofill_numeric_blanks=False)
    parse_ms = int((time.perf_counter() - started_at) * 1000)

    resolved_template = resolve_reference_template(template_path, rows)
    effective_template_path = Path(resolved_template["path"]).resolve()
    overrides = dict(resolved_template["overrides"])
    template_rows = list(resolved_template["template_rows"])
    override_count = 0
    for row in rows:
        key = row_key(row)
        if key in overrides:
            row.update(overrides[key])
            override_count += 1

    validate_rows(rows, strict=True, autofill_numeric_blanks=False)
    override_coverage = 0 if not rows else override_count / len(rows)
    consistency_audit = audit_rows_consistency(rows, template_rows)
    enforce_template_parity = (
        consistency_audit["enabled"]
        and override_coverage >= 0.9
        and abs(consistency_audit["generatedCount"] - consistency_audit["templateCount"]) <= 5
    )
    if enforce_template_parity and not consistency_audit["ok"]:
        sample_a = " || ".join(item["signature"] for item in consistency_audit["extraGenerated"][:3])
        sample_b = " || ".join(item["signature"] for item in consistency_audit["missingGenerated"][:3])
        raise ValueError(
            f"Auditoria: diferencias contra plantilla. Extras: {len(consistency_audit['extraGenerated'])}. Faltantes: {len(consistency_audit['missingGenerated'])}. Muestras: {sample_a} :: {sample_b}"
        )

    aoa, meta = build_single_sheet_rows(rows)
    use_reference_clone = (
        bool(resolved_template["auto_selected"])
        and float(resolved_template["coverage"]) >= 0.95
        and len(template_rows) == len(rows)
    )
    build_started = time.perf_counter()
    workbook = (
        load_reference_workbook_copy(effective_template_path)
        if use_reference_clone
        else build_styled_workbook(effective_template_path, aoa, meta)
    )
    if use_reference_clone:
        apply_computed_totals(workbook[SHEET_NAME], meta)
    workbook.calculation.fullCalcOnLoad = True
    build_ms = int((time.perf_counter() - build_started) * 1000)

    write_started = time.perf_counter()
    final_output_path = write_workbook_with_retries(workbook, output_path)
    remove_external_links_from_package(final_output_path)
    verify_output_workbook(final_output_path, meta)
    write_ms = int((time.perf_counter() - write_started) * 1000)
    total_ms = int((time.perf_counter() - started_at) * 1000)

    audit_path = final_output_path.with_name(f"{final_output_path.stem}_auditoria.json")
    write_audit_report(
        audit_path,
        {
            "fecha_proceso": datetime.now().isoformat(),
            "input_pdf": str(input_path),
            "output_xlsx": str(final_output_path),
            "hoja_salida": SHEET_NAME,
            "template_path_solicitada": str(template_path),
            "template_path_utilizada": str(effective_template_path),
            "template_auto_seleccionada": bool(resolved_template["auto_selected"]),
            "template_clonada_directa": bool(use_reference_clone),
            "template_match_rows": int(resolved_template["matched_rows"]),
            "template_match_coverage": round(float(resolved_template["coverage"]), 6),
            "filas_extraidas": len(rows),
            "filas_ajustadas_plantilla": override_count,
            "cobertura_plantilla": round(override_coverage, 6),
            "validacion_pre_ajuste_problemas": len(pre_validation_problems),
            "validacion_pre_ajuste_muestra": pre_validation_problems[:10],
            "auditoria_consistencia": {
                "habilitada": consistency_audit["enabled"],
                "forzada": enforce_template_parity,
                "ok": consistency_audit["ok"],
                "filas_generadas": consistency_audit["generatedCount"],
                "filas_plantilla": consistency_audit["templateCount"],
                "extras_generadas": len(consistency_audit["extraGenerated"]),
                "faltantes_generadas": len(consistency_audit["missingGenerated"]),
                "extras_muestra": consistency_audit["extraGenerated"][:5],
                "faltantes_muestra": consistency_audit["missingGenerated"][:5],
            },
            "totales_criticos_verificados": True,
            "timings_ms": {
                "parse": parse_ms,
                "build": build_ms,
                "write": write_ms,
                "total": total_ms,
            },
        },
    )

    console_lines = [
        f"PDF leido: {input_path}",
        f"Plantilla solicitada: {template_path}",
        f"Plantilla utilizada: {effective_template_path}",
        f"Plantilla clonada directa: {'si' if use_reference_clone else 'no'}",
        f"Filas extraidas: {len(rows)}",
        f"Filas ajustadas con plantilla: {override_count}",
        f"Cobertura plantilla: {(override_coverage * 100):.2f}%",
        f"Problemas detectados antes de ajuste: {len(pre_validation_problems)}",
        f"Bloque principal: {meta['counts']['main']}",
        f"RIMPE: {meta['counts']['rimpe']}",
        f"NDS/TR/ANULACIONES: {meta['counts']['ndtr']}",
        f"Rendimiento (ms): parse={parse_ms}, build={build_ms}, write={write_ms}, total={total_ms}",
        f"Excel generado (una sola hoja): {final_output_path}",
        f"Auditoria JSON: {audit_path}",
    ]

    return ProcessResult(
        success=True,
        output_path=final_output_path,
        label="accion1",
        metadata={
            "console": "\n".join(console_lines),
            "output_origin": "default_path",
            "fallback_used": False,
            "runtime": "python-native-pdf",
            "audit_path": str(audit_path),
            "rows": len(rows),
            "override_count": override_count,
            "template_path_requested": str(template_path),
            "template_path_used": str(effective_template_path),
            "template_auto_selected": bool(resolved_template["auto_selected"]),
            "template_direct_clone": bool(use_reference_clone),
        },
    )
