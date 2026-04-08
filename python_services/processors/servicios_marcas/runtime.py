from __future__ import annotations

import argparse
import csv
import gc
import math
import os
import re
import shutil
import sys
import time
import unicodedata
from dataclasses import dataclass
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[3]
PYTHON_SERVICES_ROOT = Path(__file__).resolve().parents[2]
if str(PYTHON_SERVICES_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_SERVICES_ROOT))

from bootstrap import bootstrap_vendor

bootstrap_vendor()

import pythoncom
import pywintypes
import xlrd
from win32com.client import DispatchEx

from processors.servicios_marcas.readers import (
    parse_date_text,
    parse_decimal_like,
    parse_report_date_to_excel_serial,
    read_mayor_rows,
    read_px_rows,
    read_source_rows,
)

OUTPUT_CONFIG: dict[str, dict[str, str]] = {
    "changan": {"label": "CHANGAN", "prefix": "servicios_changan_"},
    "peug": {"label": "PEUGEOT", "prefix": "servicios_peug_"},
    "szk": {"label": "SUZUKI", "prefix": "servicios_szk_"},
    "tyt": {"label": "MATRIZ", "prefix": "servicios_tyt_"},
}

BRAND_INPUT_KEYS: dict[str, dict[str, str]] = {
    "changan": {
        "factura": "factura_changan_path",
        "nota": "nota_changan_path",
        "mayor": "mayor_changan_path",
    },
    "peug": {
        "factura": "factura_peug_path",
        "nota": "nota_peug_path",
        "mayor": "mayor_peug_path",
    },
    "szk": {
        "factura": "factura_szk_path",
        "nota": "nota_szk_path",
        "mayor": "mayor_szk_path",
    },
    "tyt": {
        "factura": "factura_tyt_path",
        "nota": "nota_tyt_path",
        "mayor": "mayor_tyt_path",
    },
}

TEMPLATE_CONFIGS: dict[str, dict[str, str]] = {
    "changan": {"label": "CHANGAN", "file": "11. Concili. Servicios CHANGAN  2026.xls"},
    "peug": {"label": "PEUGEOT", "file": "11. Concili. Servicios PEUG  2026.xls"},
    "szk": {"label": "SUZUKI", "file": "11. Concili. Servicios SZK  2026.xls"},
    "tyt": {"label": "MATRIZ", "file": "11. Concili. Servicios TYT 2026.xls"},
}

ALLOW_TEMPLATE_DATA_FALLBACK = os.environ.get("SERVICIOS_MARCAS_ALLOW_TEMPLATE_FALLBACK") == "1"
PIVOT_REFRESH_ENABLED = os.environ.get("SERVICIOS_MARCAS_REFRESH_PIVOTS") == "1"

XLCELLTYPE_CONSTANTS = 2
XL_DIRECTION_UP = -4162
XL_CALCULATION_MANUAL = -4135


class CancelRequestedError(RuntimeError):
    pass


@dataclass(frozen=True)
class RuntimeArgs:
    input_path: Path
    output_dir: Path
    template_dir: Path
    run_stamp: str
    cancel_path: Path | None
    brand_key: str
    px_path: Path | None
    rep_vtas_path: Path | None
    factura_paths: dict[str, Path | None]
    nota_paths: dict[str, Path | None]
    mayor_paths: dict[str, Path | None]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_sheet_name(value: Any) -> str:
    text = normalize_text(value).upper()
    normalized = unicodedata.normalize("NFD", text)
    pieces: list[str] = []
    for char in normalized:
        if unicodedata.category(char) == "Mn":
            continue
        if char.isalnum() or char == " ":
            pieces.append(char)
    return "".join(pieces)


def to_number(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = normalize_text(value)
    if text == "":
        return 0.0
    return float(parse_decimal_like(text))


def round_amount(value: Any) -> float:
    return float(Decimal(str(to_number(value))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def trim_document(value: Any) -> str:
    text = normalize_text(value)
    if text == "":
        return ""
    digits = "".join(ch for ch in text if ch.isdigit())
    if digits == "":
        return text
    trimmed = digits.lstrip("0")
    return trimmed or "0"


def strip_order_suffix(value: Any) -> str:
    text = normalize_text(value)
    if text and text[-1:].isalpha():
        return text[:-1]
    return text


def get_document_sort_value(value: Any) -> float:
    text = trim_document(value)
    if text.isdigit():
        return float(text)
    return 0.0


def get_template_key(agency: Any, series: Any, order: Any = "") -> str | None:
    agency_text = normalize_text(agency).upper()
    order_text = normalize_text(order).upper()
    if agency_text == "CHANGAN":
        return "changan"
    if agency_text == "PEUGEOT":
        return "peug"
    if agency_text == "MATRIZ":
        return "tyt"
    if agency_text == "SUZUKI AMBATO":
        return "szk"
    if agency_text == "SUZUKI RIOBAMBA":
        if order_text.startswith("D"):
            return "changan"
        return "szk"
    return None


def get_compact_account_code(value: Any) -> str:
    text = normalize_text(value)
    if text == "":
        return ""
    digits = "".join(ch for ch in text if ch.isdigit())
    if 0 < len(digits) < 12:
        return digits.rjust(12, "0")
    return digits


def format_account_code(value: Any) -> str:
    compact = get_compact_account_code(value)
    if re.fullmatch(r"\d{12}", compact or "") is not None:
        return f"{compact[0:2]}.{compact[2:4]}.{compact[4:6]}.{compact[6:8]}.{compact[8:12]}"
    return normalize_text(value)


def get_normalized_center_code(value: Any) -> str:
    text = normalize_text(value)
    if text == "":
        return ""
    if text.isdigit():
        return text.zfill(2)
    return text


def get_brand_display_label(brand_key: str) -> str:
    return {
        "changan": "CHANGAN",
        "peug": "PEUGEOT",
        "szk": "SUZUKI",
        "tyt": "TOYOTA",
    }.get(normalize_text(brand_key).lower(), normalize_text(brand_key).upper())


def get_excel_text_literal(value: Any) -> str:
    if value is None:
        return ""
    text = str(value)
    if text == "":
        return ""
    if text.startswith("'"):
        return text
    if text[0] in ("=", "+", "-", "@"):
        return "'" + text
    if text.isdigit():
        return "'" + text
    return text


def get_invoice_asiento(row: dict[str, Any]) -> str:
    payment_method = normalize_text(row.get("FormaPago")).upper()
    if "CRED" in payment_method:
        return "C"
    if "CONT" in payment_method or "EFEC" in payment_method:
        return "E"
    if normalize_text(row.get("DocType")).upper() == "FC":
        return "C"
    return "E"


def get_preferred_source_text(source_raw: Any, source_normalized: Any = "", lookup_value: Any = "") -> Any:
    if normalize_text(source_raw) != "":
        return source_raw
    if normalize_text(source_normalized) != "":
        return source_normalized
    if not ALLOW_TEMPLATE_DATA_FALLBACK:
        return source_normalized
    if normalize_text(lookup_value) != "":
        return lookup_value
    return source_normalized


def get_preferred_lookup_text(lookup_value: Any, source_raw: Any = "", source_normalized: Any = "") -> Any:
    return get_preferred_source_text(source_raw, source_normalized, lookup_value)


def get_preferred_source_date(source_value: Any, lookup_value: Any = None) -> Any:
    if source_value not in (None, ""):
        return source_value
    if not ALLOW_TEMPLATE_DATA_FALLBACK:
        return None
    return lookup_value


def get_lookup_default_text(lookups: dict[str, Any], section: str, field: str) -> str:
    if not ALLOW_TEMPLATE_DATA_FALLBACK:
        return ""
    defaults = lookups.get("Defaults", {})
    section_defaults = defaults.get(section, {})
    return normalize_text(section_defaults.get(field, ""))


def resolve_template_garext(lookup_value: Any, source_raw: Any = "", source_normalized: Any = "", template_default: Any = "") -> Any:
    source_value = get_preferred_source_text(source_raw, source_normalized)
    if not ALLOW_TEMPLATE_DATA_FALLBACK:
        return source_value

    source_text = normalize_text(source_value).upper()
    default_text = normalize_text(template_default)
    if source_text != "":
        if source_text in {"N", "NO", "0", "FALSE"}:
            return default_text
        return source_value

    lookup_text = normalize_text(lookup_value).upper()
    if lookup_text == "":
        return default_text
    if lookup_text in {"N", "NO", "0", "FALSE"}:
        return default_text
    return lookup_value


def get_iva_buckets(row: dict[str, Any], net_base: float) -> dict[str, float]:
    iva_total = round_amount(abs(to_number(row.get("Iva"))))
    iva12 = round_amount(abs(to_number(row.get("Iva12")))) if "Iva12" in row else 0.0
    iva15 = round_amount(abs(to_number(row.get("Iva15")))) if "Iva15" in row else 0.0

    if (iva12 + iva15) > 0:
        if iva_total == 0:
            iva_total = round_amount(iva12 + iva15)
    elif net_base > 0 and iva_total > 0:
        rate = iva_total / net_base
        if rate >= 0.14:
            iva15 = iva_total
        elif rate >= 0.105:
            iva12 = iva_total
        else:
            iva12 = iva_total

    if (iva12 + iva15) == 0 and iva_total > 0:
        iva12 = iva_total

    return {"Total": float(iva_total), "Iva12": float(iva12), "Iva15": float(iva15)}


def get_invoice_source_amounts(row: dict[str, Any]) -> dict[str, float]:
    subtotal_source = abs(to_number(row.get("Subtotal")))
    discount_source = abs(to_number(row.get("Discount"))) or abs(to_number(row.get("NoteCredit")))
    subtotal_raw = (
        abs(to_number(row.get("TotalManoObra")))
        + abs(to_number(row.get("TotalSubcontratos")))
        + abs(to_number(row.get("TotalInsumos")))
        + abs(to_number(row.get("TotalAccesorios")))
        + abs(to_number(row.get("TotalRepuestos")))
    )
    subtotal = round_amount(subtotal_source if subtotal_source > 0 else subtotal_raw)
    discount = round_amount(discount_source)
    neto_con_iva_source = abs(to_number(row.get("NetoConIva")))
    neto_con_iva = round_amount(neto_con_iva_source if neto_con_iva_source > 0 else subtotal - discount)
    iva_buckets = get_iva_buckets(row, neto_con_iva)
    iva_amount = float(iva_buckets["Total"])
    interest_amount = round_amount(abs(to_number(row.get("Interest"))) or abs(to_number(row.get("Interes"))))
    total_source = abs(to_number(row.get("Total")))
    total_amount = round_amount(total_source if total_source > 0 else neto_con_iva + iva_amount + interest_amount)
    neto_iva0_source = abs(to_number(row.get("NetoIva0")))
    neto_iva0 = round_amount(neto_iva0_source if neto_iva0_source > 0 else (neto_con_iva if round_amount(abs(iva_amount)) == 0 else 0.0))
    return {
        "Total": float(total_amount),
        "Iva": float(iva_amount),
        "Iva12": float(iva_buckets["Iva12"]),
        "Iva15": float(iva_buckets["Iva15"]),
        "Interest": float(interest_amount),
        "NetoConIva": float(neto_con_iva),
        "Discount": float(discount),
        "Subtotal": float(subtotal),
        "NetoIva0": float(neto_iva0),
    }


def get_note_source_amounts(row: dict[str, Any]) -> dict[str, float]:
    subtotal_source = abs(to_number(row.get("Subtotal")))
    discount_source = abs(to_number(row.get("Discount"))) or abs(to_number(row.get("NoteCredit")))
    subtotal_raw = (
        abs(to_number(row.get("TotalManoObra")))
        + abs(to_number(row.get("TotalSubcontratos")))
        + abs(to_number(row.get("TotalInsumos")))
        + abs(to_number(row.get("TotalAccesorios")))
        + abs(to_number(row.get("TotalRepuestos")))
    )
    subtotal = round_amount(subtotal_source if subtotal_source > 0 else subtotal_raw)
    discount = round_amount(discount_source)
    neto_con_iva_source = abs(to_number(row.get("NetoConIva")))
    neto_con_iva = round_amount(neto_con_iva_source if neto_con_iva_source > 0 else subtotal - discount)
    iva_buckets = get_iva_buckets(row, neto_con_iva)
    iva_amount = float(iva_buckets["Total"])
    interest_amount = round_amount(abs(to_number(row.get("Interest"))) or abs(to_number(row.get("Interes"))))
    total_source = abs(to_number(row.get("Total")))
    total_amount = round_amount(total_source if total_source > 0 else neto_con_iva + iva_amount + interest_amount)
    neto_sin_iva_source = abs(to_number(row.get("NetoSinIva")))
    neto_sin_iva = round_amount(neto_sin_iva_source if neto_sin_iva_source > 0 else (neto_con_iva if round_amount(abs(iva_amount)) == 0 else 0.0))
    anticipo = round_amount(abs(to_number(row.get("Anticipo"))))
    neto = round_amount(abs(to_number(row.get("Neto"))))
    if neto == 0.0:
        neto = round_amount(total_amount - anticipo)
    return {
        "Total": float(total_amount),
        "Iva": float(iva_amount),
        "Iva12": float(iva_buckets["Iva12"]),
        "Iva15": float(iva_buckets["Iva15"]),
        "Interest": float(interest_amount),
        "NetoConIva": float(neto_con_iva),
        "Discount": float(discount),
        "Subtotal": float(subtotal),
        "NetoSinIva": float(neto_sin_iva),
        "Anticipo": float(anticipo),
        "Neto": float(neto),
    }


def assert_not_cancelled(cancel_path: Path | None, context: str = "proceso") -> None:
    if cancel_path is not None and cancel_path.is_file():
        raise CancelRequestedError("Proceso detenido por el usuario.")


def parse_tab_file(path: Path | None, kind: str = "") -> list[dict[str, Any]]:
    if path is None or not path.is_file():
        return []

    lines = path.read_text(encoding="utf-8-sig", errors="ignore").splitlines()
    if lines:
        sample = lines[0].split("\t")
        if kind == "factura" and len(sample) >= 36 and normalize_text(sample[0]).upper() == "AGENCIA :" and normalize_text(sample[2]).upper() == "SERIE":
            pass
        elif kind == "nota" and len(sample) >= 40 and normalize_text(sample[0]).upper() == "AGENCIA :" and normalize_text(sample[2]).upper() == "NOTA CRED.":
            pass
        else:
            try:
                with path.open("r", encoding="utf-8-sig", newline="") as handle:
                    reader = csv.DictReader(handle, delimiter="\t")
                    rows = [row for row in reader]
                filtered = [
                    row
                    for row in rows
                    if normalize_text(row.get("Agencia", "")) != ""
                    or normalize_text(row.get("Agencia :", "")) != ""
                    or normalize_text(row.get("Factura", "")) != ""
                    or normalize_text(row.get("Nota Cred.", "")) != ""
                ]
                if filtered:
                    return filtered
            except Exception:
                pass

    fallback_rows: list[dict[str, Any]] = []
    for raw_line in lines:
        columns = raw_line.split("\t")
        if not columns:
            continue

        if kind == "factura":
            if len(columns) < 36:
                continue
            if normalize_text(columns[0]).upper() != "AGENCIA :" or normalize_text(columns[2]).upper() != "SERIE":
                continue
            fallback_rows.append(
                {
                    "Agencia :": normalize_text(columns[1]),
                    "Serie": normalize_text(columns[19]),
                    "Factura": normalize_text(columns[20]),
                    "Fecha": normalize_text(columns[21]),
                    "Orden": normalize_text(columns[22]),
                    "C.I.": normalize_text(columns[23]),
                    "Cliente": normalize_text(columns[24]),
                    "Sub total": normalize_text(columns[25]),
                    "Des-cuento": normalize_text(columns[26]),
                    "Neto con IVA": normalize_text(columns[27]),
                    "Neto Iva 0": normalize_text(columns[28]),
                    "Iva 12%": normalize_text(columns[29]),
                    "Iva 15%": normalize_text(columns[30]),
                    "Inte-reses": normalize_text(columns[31]),
                    "Total": normalize_text(columns[32]),
                    "Asiento": normalize_text(columns[33]),
                    "Gar.Ext.": normalize_text(columns[34]),
                    "T.V.": normalize_text(columns[35]),
                }
            )
            continue

        if kind == "nota":
            if len(columns) < 40:
                continue
            if normalize_text(columns[0]).upper() != "AGENCIA :" or normalize_text(columns[2]).upper() != "NOTA CRED.":
                continue
            fallback_rows.append(
                {
                    "Agencia :": normalize_text(columns[1]),
                    "Nota Cred.": normalize_text(columns[20]),
                    "Fecha": normalize_text(columns[21]),
                    "Tipo": normalize_text(columns[22]),
                    "Serie": normalize_text(columns[23]),
                    "Factura": normalize_text(columns[24]),
                    "Orden": normalize_text(columns[25]),
                    "Cedula": normalize_text(columns[26]),
                    "Cliente": normalize_text(columns[27]),
                    "Sub total": normalize_text(columns[28]),
                    "Des-cuento": normalize_text(columns[29]),
                    "Netosin Iva": normalize_text(columns[30]),
                    "Netocon Iva": normalize_text(columns[31]),
                    "Iva 15%": normalize_text(columns[32]),
                    "Iva 12 %": normalize_text(columns[33]),
                    "Interes": normalize_text(columns[34]),
                    "Total": normalize_text(columns[35]),
                    "Anticipo": normalize_text(columns[36]),
                    "NETO": normalize_text(columns[37]),
                    "Asiento": normalize_text(columns[38]),
                    "Gar.Ext.": normalize_text(columns[39]),
                }
            )

    return fallback_rows


def build_source_rows_from_brand_inputs(brand_file_map: dict[str, dict[str, Path | None]], brand_key: str = "") -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    brand_order = ["changan", "peug", "szk", "tyt"]
    if brand_key:
        brand_order = [item for item in brand_order if item == brand_key]

    def append_rows(items: list[dict[str, Any]], doc_type: str, template_key: str) -> None:
        index = len(rows) + 1
        for item in items:
            agency_raw = item.get("Agencia :", "") or item.get("Agencia", "")
            agency = normalize_text(agency_raw)
            if agency == "" or template_key == "":
                continue

            series_raw = normalize_text(item.get("Serie", ""))
            document_raw = item.get("Nota Cred.", "") if doc_type == "DC" else item.get("Factura", "")
            document_raw = document_raw or item.get("Factura", "") or item.get("Nota Cred.", "")
            document = normalize_text(document_raw)
            if document == "":
                continue

            order_raw = normalize_text(item.get("Orden", ""))
            cedula_raw = item.get("Cedula", "") or item.get("C.I.", "")
            cedula = normalize_text(cedula_raw)
            customer_raw = normalize_text(item.get("Cliente", ""))
            affected_document_raw = normalize_text(item.get("Factura", "")) if doc_type == "DC" else ""

            fecha_text = normalize_text(item.get("Fecha", ""))
            date_fact_value = None
            date_note_value = None
            if fecha_text:
                excel_serial = parse_date_text(fecha_text)
                if excel_serial is None:
                    excel_serial = parse_report_date_to_excel_serial(fecha_text)
                if excel_serial is not None:
                    if doc_type == "DC":
                        date_note_value = float(excel_serial)
                    else:
                        date_fact_value = float(excel_serial)

            subtotal = float(parse_decimal_like(item.get("Sub total", "")))
            discount = float(parse_decimal_like(item.get("Des-cuento", "")))
            iva12 = float(parse_decimal_like(item.get("Iva 12%", "")))
            if iva12 == 0.0:
                iva12 = float(parse_decimal_like(item.get("Iva 12 %", "")))
            iva15 = float(parse_decimal_like(item.get("Iva 15%", "")))
            iva = float(Decimal(str(iva12 + iva15)))
            interest = float(parse_decimal_like(item.get("Interes", "")))
            total = float(parse_decimal_like(item.get("Total", "")))
            gar_ext = normalize_text(item.get("Gar.Ext.", ""))
            tv = normalize_text(item.get("T.V.", ""))
            anticipo = float(parse_decimal_like(item.get("Anticipo", "")))
            neto = float(parse_decimal_like(item.get("NETO", "")))
            neto_con_iva = float(parse_decimal_like(item.get("Netocon Iva", "")))
            neto_sin_iva = float(parse_decimal_like(item.get("Netosin Iva", "")))
            neto_iva0 = float(parse_decimal_like(item.get("Neto Iva 0", "")))

            note_credit = discount
            mano_obra = subtotal
            if not mano_obra:
                if doc_type == "DC":
                    mano_obra = neto_con_iva + neto_sin_iva + discount
                else:
                    mano_obra = neto_con_iva + neto_iva0 + discount

            rows.append(
                {
                    "RowIndex": index,
                    "TemplateKey": template_key,
                    "Agency": agency,
                    "AgencyRaw": agency_raw,
                    "Center": "",
                    "CenterRaw": "",
                    "Order": order_raw,
                    "OrderRaw": order_raw,
                    "Advisor": "",
                    "AdvisorRaw": "",
                    "Line": "",
                    "LineRaw": "",
                    "DocType": doc_type,
                    "Cedula": cedula,
                    "CedulaRaw": cedula_raw,
                    "Customer": customer_raw,
                    "CustomerRaw": customer_raw,
                    "DocumentRaw": document,
                    "DocumentTrim": trim_document(document),
                    "Series": series_raw,
                    "SeriesRaw": series_raw,
                    "FormaPago": tv,
                    "Authorization": "",
                    "DateFactValue": date_fact_value,
                    "DateNoteValue": date_note_value,
                    "Subtotal": subtotal or 0.0,
                    "Discount": discount or 0.0,
                    "NetoConIva": neto_con_iva or 0.0,
                    "NetoSinIva": neto_sin_iva or 0.0,
                    "NetoIva0": neto_iva0 or 0.0,
                    "NoteCredit": note_credit or 0.0,
                    "TotalManoObra": mano_obra or subtotal or 0.0,
                    "TotalSubcontratos": 0.0,
                    "TotalInsumos": 0.0,
                    "TotalServicio": 0.0,
                    "TotalAccesorios": 0.0,
                    "TotalRepuestos": 0.0,
                    "Interes": interest or 0.0,
                    "Iva": iva or 0.0,
                    "Iva12": iva12 or 0.0,
                    "Iva15": iva15 or 0.0,
                    "Total": total or 0.0,
                    "Anticipo": anticipo or 0.0,
                    "Neto": neto or 0.0,
                    "Costo": 0.0,
                    "CostoLubricantes": 0.0,
                    "CostoAccesorios": 0.0,
                    "CostoRepuestos": 0.0,
                    "CostoPintura": 0.0,
                    "CostoSubconNc": 0.0,
                    "GarExt": gar_ext,
                    "GarExtRaw": gar_ext,
                    "AffectedDocumentTrim": trim_document(affected_document_raw),
                    "AffectedDocumentRaw": affected_document_raw,
                    "MotivoNc": "",
                    "ObservacionNc": "",
                }
            )
            index += 1

    for target_brand in brand_order:
        paths = brand_file_map.get(target_brand)
        if not paths:
            continue
        facturas = parse_tab_file(paths.get("FacturaPath"), "factura")
        notas = parse_tab_file(paths.get("NotaPath"), "nota")
        append_rows(facturas, "FA", target_brand)
        append_rows(notas, "DC", target_brand)

    return rows


def last_non_empty_value(group: list[dict[str, Any]], *keys: str) -> Any:
    for item in reversed(group):
        for key in keys:
            value = item.get(key)
            if normalize_text(value) != "":
                return value
    return ""


def normalize_source_rows(rows: list[dict[str, Any]], consolidate_invoice_documents: bool = False) -> list[dict[str, Any]]:
    sorted_rows = sorted(rows, key=lambda item: int(item.get("RowIndex", 0)))
    normalized: list[dict[str, Any]] = []
    note_seen: set[str] = set()
    invoice_groups: dict[str, list[dict[str, Any]]] = {}

    for row in sorted_rows:
        doc_type = normalize_text(row.get("DocType")).upper()
        if doc_type in {"DC", "DE"}:
            dedupe_key = "|".join(
                [
                    normalize_text(row.get("TemplateKey")),
                    doc_type,
                    normalize_text(row.get("DocumentTrim")),
                    normalize_text(row.get("Order")),
                    str(get_date_write_value(row.get("DateNoteValue"))),
                    str(round_amount(row.get("NoteCredit"))),
                    str(round_amount(row.get("TotalManoObra"))),
                    str(round_amount(row.get("TotalSubcontratos"))),
                    str(round_amount(row.get("TotalInsumos"))),
                    str(round_amount(row.get("TotalServicio"))),
                    str(round_amount(row.get("TotalAccesorios"))),
                    str(round_amount(row.get("TotalRepuestos"))),
                    str(round_amount(row.get("Total"))),
                    str(round_amount(row.get("Iva"))),
                    str(round_amount(row.get("Interes"))),
                    str(round_amount(row.get("Anticipo", 0.0))),
                    str(round_amount(row.get("Neto", 0.0))),
                    normalize_text(row.get("AffectedDocumentTrim")),
                    normalize_text(row.get("MotivoNc")),
                    normalize_text(row.get("ObservacionNc")),
                ]
            )
            if dedupe_key in note_seen:
                continue
            note_seen.add(dedupe_key)
            normalized.append(dict(row))
            continue

        if not consolidate_invoice_documents:
            normalized.append(dict(row))
            continue

        invoice_group_key = "|".join(
            [
                normalize_text(row.get("TemplateKey")),
                doc_type,
                normalize_text(row.get("DocumentTrim")),
                normalize_text(row.get("Series")),
                str(get_date_write_value(row.get("DateFactValue"))),
            ]
        )
        invoice_groups.setdefault(invoice_group_key, []).append(row)

    if consolidate_invoice_documents:
        sum_fields = [
            "Subtotal",
            "Discount",
            "NetoConIva",
            "NetoSinIva",
            "NetoIva0",
            "Anticipo",
            "Neto",
            "NoteCredit",
            "TotalManoObra",
            "TotalSubcontratos",
            "TotalInsumos",
            "TotalServicio",
            "TotalAccesorios",
            "TotalRepuestos",
            "Interes",
            "Iva12",
            "Iva15",
            "Iva",
            "Total",
            "Costo",
            "CostoLubricantes",
            "CostoAccesorios",
            "CostoRepuestos",
            "CostoPintura",
            "CostoSubconNc",
        ]
        for group in invoice_groups.values():
            if len(group) == 1:
                normalized.append(dict(group[0]))
                continue
            distinct_orders = {normalize_text(item.get("Order")) for item in group if normalize_text(item.get("Order")) != ""}
            if len(distinct_orders) <= 1:
                normalized.append(dict(group[0]))
                continue
            first = sorted(group, key=lambda item: int(item.get("RowIndex", 0)))[0]
            sum_values = {field: 0.0 for field in sum_fields}
            for item in group:
                for field in sum_fields:
                    sum_values[field] += float(to_number(item.get(field, 0.0)))
            merged = dict(first)
            merged["Order"] = ""
            merged["OrderRaw"] = ""
            merged["Agency"] = last_non_empty_value(group, "AgencyRaw", "Agency")
            merged["AgencyRaw"] = last_non_empty_value(group, "AgencyRaw", "Agency")
            merged["Center"] = last_non_empty_value(group, "CenterRaw", "Center")
            merged["CenterRaw"] = last_non_empty_value(group, "CenterRaw", "Center")
            merged["Advisor"] = last_non_empty_value(group, "AdvisorRaw", "Advisor")
            merged["AdvisorRaw"] = last_non_empty_value(group, "AdvisorRaw", "Advisor")
            merged["Line"] = last_non_empty_value(group, "LineRaw", "Line")
            merged["LineRaw"] = last_non_empty_value(group, "LineRaw", "Line")
            merged["Cedula"] = last_non_empty_value(group, "CedulaRaw", "Cedula")
            merged["CedulaRaw"] = last_non_empty_value(group, "CedulaRaw", "Cedula")
            merged["Customer"] = last_non_empty_value(group, "CustomerRaw", "Customer")
            merged["CustomerRaw"] = last_non_empty_value(group, "CustomerRaw", "Customer")
            merged["DocumentRaw"] = last_non_empty_value(group, "DocumentRaw", "DocumentTrim")
            merged["Series"] = last_non_empty_value(group, "SeriesRaw", "Series")
            merged["SeriesRaw"] = last_non_empty_value(group, "SeriesRaw", "Series")
            merged["FormaPago"] = last_non_empty_value(group, "FormaPago")
            merged["Authorization"] = last_non_empty_value(group, "Authorization")
            merged["GarExt"] = last_non_empty_value(group, "GarExtRaw", "GarExt")
            merged["GarExtRaw"] = last_non_empty_value(group, "GarExtRaw", "GarExt")
            merged["AffectedDocumentTrim"] = last_non_empty_value(group, "AffectedDocumentRaw", "AffectedDocumentTrim")
            merged["AffectedDocumentRaw"] = last_non_empty_value(group, "AffectedDocumentRaw", "AffectedDocumentTrim")
            merged["MotivoNc"] = last_non_empty_value(group, "MotivoNc")
            merged["ObservacionNc"] = last_non_empty_value(group, "ObservacionNc")
            for field in sum_fields:
                merged[field] = sum_values[field]
            normalized.append(merged)

    return sorted(
        normalized,
        key=lambda item: (
            int(item.get("RowIndex", 0)),
            normalize_text(item.get("TemplateKey")),
            normalize_text(item.get("DocType")),
            get_document_sort_value(item.get("DocumentTrim")),
            normalize_text(item.get("Order")),
        ),
    )


def new_lookup_store() -> dict[str, dict[str, Any]]:
    return {"ByDocOrder": {}, "ByDoc": {}, "ByOrder": {}}


def add_lookup_entry(store: dict[str, dict[str, Any]], doc_key: Any, order_key: Any, entry: dict[str, Any]) -> None:
    doc_key_text = normalize_text(doc_key)
    order_key_text = normalize_text(order_key)
    if doc_key_text and order_key_text:
        store["ByDocOrder"].setdefault(f"{doc_key_text}|{order_key_text}", entry)
    if doc_key_text:
        store["ByDoc"].setdefault(doc_key_text, entry)
    if order_key_text:
        store["ByOrder"].setdefault(order_key_text, entry)


def find_lookup_entry(store: dict[str, dict[str, Any]], doc_key: Any, order_key: Any) -> dict[str, Any] | None:
    doc_key_text = normalize_text(doc_key)
    order_key_text = normalize_text(order_key)
    if doc_key_text and order_key_text:
        item = store["ByDocOrder"].get(f"{doc_key_text}|{order_key_text}")
        if item is not None:
            return item
    if doc_key_text:
        item = store["ByDoc"].get(doc_key_text)
        if item is not None:
            return item
    if order_key_text:
        return store["ByOrder"].get(order_key_text)
    return None


def find_rep_vtas_entry(store: dict[str, dict[str, Any]], doc_key: Any, order_key: Any) -> dict[str, Any] | None:
    doc_key_text = normalize_text(doc_key)
    order_key_text = normalize_text(order_key)
    if doc_key_text and order_key_text:
        item = store["ByDocOrder"].get(f"{doc_key_text}|{order_key_text}")
        if item is not None:
            return item
    if order_key_text:
        item = store["ByOrder"].get(order_key_text)
        if item is not None:
            return item
    if doc_key_text:
        return store["ByDoc"].get(doc_key_text)
    return None


def open_template_book(template_path: Path) -> xlrd.book.Book:
    return xlrd.open_workbook(str(template_path), on_demand=True)


def get_xls_sheet(book: xlrd.book.Book, candidate_names: list[str]) -> xlrd.sheet.Sheet:
    normalized_candidates = {normalize_sheet_name(name) for name in candidate_names}
    for index in range(book.nsheets):
        sheet = book.sheet_by_index(index)
        if normalize_sheet_name(sheet.name) in normalized_candidates:
            return sheet
    raise RuntimeError(f"No se encontro la hoja requerida: {', '.join(candidate_names)}")


def xls_text(book: xlrd.book.Book, sheet: xlrd.sheet.Sheet, row: int, column: int) -> str:
    cell_value = sheet.cell(row - 1, column - 1)
    if cell_value.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return ""
    if cell_value.ctype == xlrd.XL_CELL_NUMBER and float(cell_value.value).is_integer():
        return str(int(cell_value.value))
    if cell_value.ctype == xlrd.XL_CELL_DATE:
        return xlrd.xldate.xldate_as_datetime(cell_value.value, book.datemode).strftime("%d/%m/%Y")
    return normalize_text(cell_value.value)


def xls_number(book: xlrd.book.Book, sheet: xlrd.sheet.Sheet, row: int, column: int) -> float:
    cell_value = sheet.cell(row - 1, column - 1)
    if cell_value.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return 0.0
    return to_number(cell_value.value)


def xls_last_row(book: xlrd.book.Book, sheet: xlrd.sheet.Sheet, key_column: int) -> int:
    for row in range(sheet.nrows, 0, -1):
        if xls_text(book, sheet, row, key_column) != "":
            return row
    return 1


def read_template_lookups_from_xls(template_path: Path) -> dict[str, Any]:
    book = open_template_book(template_path)
    invoice_store = new_lookup_store()
    note_store = new_lookup_store()
    rep_vtas_store = new_lookup_store()
    invoice_sheet = get_xls_sheet(book, ["REP FACTURACION", "REP FACTURACION"])
    note_sheet = get_xls_sheet(book, ["NOTA DE CREDITO"])
    rep_vtas_sheet = get_xls_sheet(book, ["REP VTAS"])

    defaults = {
        "Invoice": {"GarExt": next((xls_text(book, invoice_sheet, row, 17) for row in range(17, xls_last_row(book, invoice_sheet, 3) + 1) if xls_text(book, invoice_sheet, row, 3) != "" and xls_text(book, invoice_sheet, row, 17) != ""), "")},
        "Note": {"GarExt": next((xls_text(book, note_sheet, row, 21) for row in range(11, xls_last_row(book, note_sheet, 2) + 1) if xls_text(book, note_sheet, row, 2) != "" and xls_text(book, note_sheet, row, 21) != ""), "")},
        "RepVtas": {"GarExt": next((xls_text(book, rep_vtas_sheet, row, 27) for row in range(15, xls_last_row(book, rep_vtas_sheet, 8) + 1) if xls_text(book, rep_vtas_sheet, row, 8) != "" and xls_text(book, rep_vtas_sheet, row, 27) != ""), "")},
    }

    for row in range(17, xls_last_row(book, invoice_sheet, 3) + 1):
        doc_key = trim_document(xls_text(book, invoice_sheet, row, 3))
        if doc_key == "":
            continue
        order_key = xls_text(book, invoice_sheet, row, 5)
        add_lookup_entry(invoice_store, doc_key, order_key, {"Agency": xls_text(book, invoice_sheet, row, 1), "Series": xls_text(book, invoice_sheet, row, 2), "Order": order_key, "Cedula": xls_text(book, invoice_sheet, row, 6), "Customer": xls_text(book, invoice_sheet, row, 7), "Asiento": xls_text(book, invoice_sheet, row, 16), "GarExt": xls_text(book, invoice_sheet, row, 17), "Tv": xls_text(book, invoice_sheet, row, 18)})
    for row in range(11, xls_last_row(book, note_sheet, 2) + 1):
        doc_key = trim_document(xls_text(book, note_sheet, row, 2))
        if doc_key == "":
            continue
        order_key = xls_text(book, note_sheet, row, 7)
        add_lookup_entry(note_store, doc_key, order_key, {"Agency": xls_text(book, note_sheet, row, 1), "Kind": xls_text(book, note_sheet, row, 4), "Series": xls_text(book, note_sheet, row, 5), "Invoice": trim_document(xls_text(book, note_sheet, row, 6)), "Order": order_key, "Cedula": xls_text(book, note_sheet, row, 8), "Customer": xls_text(book, note_sheet, row, 9), "GarExt": xls_text(book, note_sheet, row, 21)})
    for row in range(15, xls_last_row(book, rep_vtas_sheet, 3) + 1):
        doc_key = trim_document(xls_text(book, rep_vtas_sheet, row, 8))
        order_key = xls_text(book, rep_vtas_sheet, row, 3)
        if doc_key == "" and order_key == "":
            continue
        add_lookup_entry(rep_vtas_store, doc_key, order_key, {"RowOrder": row, "Agency": xls_text(book, rep_vtas_sheet, row, 1), "Center": xls_text(book, rep_vtas_sheet, row, 2), "Order": order_key, "Advisor": xls_text(book, rep_vtas_sheet, row, 4), "Line": xls_text(book, rep_vtas_sheet, row, 5), "Cedula": xls_text(book, rep_vtas_sheet, row, 6), "Customer": xls_text(book, rep_vtas_sheet, row, 7), "DocumentRaw": xls_text(book, rep_vtas_sheet, row, 8), "DateFactValue": xls_number(book, rep_vtas_sheet, row, 9), "DateNoteValue": xls_number(book, rep_vtas_sheet, row, 10), "GarExt": xls_text(book, rep_vtas_sheet, row, 27)})

    return {"Invoice": invoice_store, "Note": note_store, "RepVtas": rep_vtas_store, "Defaults": defaults}


def read_precont_ventas_prototypes_from_xls(template_path: Path) -> list[dict[str, Any]]:
    book = open_template_book(template_path)
    sheet = get_xls_sheet(book, ["PrecontabilizacionVentas"])
    prototypes: list[dict[str, Any]] = []
    for row in range(2, sheet.nrows + 1):
        account_text = xls_text(book, sheet, row, 5)
        account_digits = "".join(ch for ch in account_text if ch.isdigit())
        account = account_digits.rjust(12, "0") if len(account_digits) >= 10 else account_text
        doc = xls_text(book, sheet, row, 3)
        if account == "" or doc == "":
            continue
        prototypes.append({"TemplateRow": row, "Ag": xls_text(book, sheet, row, 2), "Doc": doc, "Line": xls_text(book, sheet, row, 4), "Account": account, "Description": xls_text(book, sheet, row, 6), "CostCenter": xls_text(book, sheet, row, 7), "Asiento": xls_text(book, sheet, row, 10)})
    return prototypes


def read_precont_costos2_prototypes_from_xls(template_path: Path) -> list[dict[str, Any]]:
    book = open_template_book(template_path)
    sheet = get_xls_sheet(book, ["PrecontabilizacionCostos (2)"])
    prototypes: list[dict[str, Any]] = []
    for row in range(2, sheet.nrows + 1):
        account_text = xls_text(book, sheet, row, 4)
        account_digits = "".join(ch for ch in account_text if ch.isdigit())
        account = account_digits.rjust(12, "0") if len(account_digits) >= 10 else account_text
        if account == "":
            continue
        prototypes.append({"TemplateRow": row, "Ag": xls_text(book, sheet, row, 2), "Line": xls_text(book, sheet, row, 3), "Account": account, "Number": xls_text(book, sheet, row, 5), "Description": xls_text(book, sheet, row, 6), "CostCenter": xls_text(book, sheet, row, 7), "Asiento": xls_text(book, sheet, row, 10)})
    return prototypes


def cell(sheet: Any, row: int, column: int) -> Any:
    return sheet.Cells(row, column)


def get_worksheet_safe(workbook: Any, candidate_names: list[str]) -> Any:
    normalized_candidates = {normalize_sheet_name(name) for name in candidate_names}
    for index in range(1, int(workbook.Worksheets.Count) + 1):
        worksheet = workbook.Worksheets(index)
        if normalize_sheet_name(worksheet.Name) in normalized_candidates:
            return worksheet
    raise RuntimeError(f"No se encontro la hoja requerida: {', '.join(candidate_names)}")


def get_date_value(worksheet: Any, row: int, column: int) -> float | None:
    raw = cell(worksheet, row, column).Value2
    if isinstance(raw, (int, float)):
        return float(raw)
    text = normalize_text(cell(worksheet, row, column).Text)
    if text == "":
        return None
    for parser in ("%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            parsed = datetime.strptime(text, parser)
            return float((parsed - datetime(1899, 12, 30)).days)
        except ValueError:
            continue
    return None


def get_date_write_value(value: Any) -> float | None:
    if value in (None, ""):
        return None
    return float(math.floor(float(value) + 0.000001))


def get_worksheet_column_default_text(worksheet: Any, start_row: int, key_column: int, value_column: int) -> str:
    last_row = int(cell(worksheet, worksheet.Rows.Count, key_column).End(XL_DIRECTION_UP).Row)
    for row in range(start_row, last_row + 1):
        key_text = normalize_text(cell(worksheet, row, key_column).Text)
        if key_text == "":
            continue
        value_text = normalize_text(cell(worksheet, row, value_column).Text)
        if value_text != "":
            return value_text
    return ""


def read_template_lookups(workbook: Any) -> dict[str, Any]:
    invoice_store = new_lookup_store()
    note_store = new_lookup_store()
    rep_vtas_store = new_lookup_store()

    invoice_sheet = get_worksheet_safe(workbook, ["REP FACTURACION", "REP FACTURACION"])
    note_sheet = get_worksheet_safe(workbook, ["NOTA DE CREDITO"])
    rep_vtas_sheet = get_worksheet_safe(workbook, ["REP VTAS"])

    defaults = {
        "Invoice": {"GarExt": get_worksheet_column_default_text(invoice_sheet, 17, 3, 17)},
        "Note": {"GarExt": get_worksheet_column_default_text(note_sheet, 11, 2, 21)},
        "RepVtas": {"GarExt": get_worksheet_column_default_text(rep_vtas_sheet, 15, 8, 27)},
    }

    last_invoice_row = int(cell(invoice_sheet, invoice_sheet.Rows.Count, 3).End(XL_DIRECTION_UP).Row)
    for row in range(17, last_invoice_row + 1):
        doc_key = trim_document(cell(invoice_sheet, row, 3).Text)
        if doc_key == "":
            continue
        order_key = normalize_text(cell(invoice_sheet, row, 5).Text)
        entry = {
            "Agency": normalize_text(cell(invoice_sheet, row, 1).Text),
            "Series": normalize_text(cell(invoice_sheet, row, 2).Text),
            "Document": doc_key,
            "Order": order_key,
            "Cedula": normalize_text(cell(invoice_sheet, row, 6).Text),
            "Customer": normalize_text(cell(invoice_sheet, row, 7).Text),
            "Subtotal": to_number(cell(invoice_sheet, row, 8).Value2),
            "Discount": to_number(cell(invoice_sheet, row, 9).Value2),
            "NetoConIva": to_number(cell(invoice_sheet, row, 10).Value2),
            "NetoIva0": to_number(cell(invoice_sheet, row, 11).Value2),
            "Iva12": to_number(cell(invoice_sheet, row, 12).Value2),
            "Iva": to_number(cell(invoice_sheet, row, 13).Value2),
            "Interest": to_number(cell(invoice_sheet, row, 14).Value2),
            "Total": to_number(cell(invoice_sheet, row, 15).Value2),
            "IvaText": normalize_text(cell(invoice_sheet, row, 13).Text),
            "Asiento": normalize_text(cell(invoice_sheet, row, 16).Text),
            "GarExt": normalize_text(cell(invoice_sheet, row, 17).Text),
            "Tv": normalize_text(cell(invoice_sheet, row, 18).Text),
            "Marker": normalize_text(cell(invoice_sheet, row, 19).Text),
        }
        add_lookup_entry(invoice_store, doc_key, order_key, entry)

    last_note_row = int(cell(note_sheet, note_sheet.Rows.Count, 2).End(XL_DIRECTION_UP).Row)
    for row in range(11, last_note_row + 1):
        doc_key = trim_document(cell(note_sheet, row, 2).Text)
        if doc_key == "":
            continue
        order_key = normalize_text(cell(note_sheet, row, 7).Text)
        entry = {
            "Agency": normalize_text(cell(note_sheet, row, 1).Text),
            "Document": doc_key,
            "Kind": normalize_text(cell(note_sheet, row, 4).Text),
            "Series": normalize_text(cell(note_sheet, row, 5).Text),
            "Invoice": trim_document(cell(note_sheet, row, 6).Text),
            "Order": order_key,
            "Cedula": normalize_text(cell(note_sheet, row, 8).Text),
            "Customer": normalize_text(cell(note_sheet, row, 9).Text),
            "Subtotal": to_number(cell(note_sheet, row, 10).Value2),
            "Discount": to_number(cell(note_sheet, row, 11).Value2),
            "NetoSinIva": to_number(cell(note_sheet, row, 12).Value2),
            "NetoConIva": to_number(cell(note_sheet, row, 13).Value2),
            "Iva": to_number(cell(note_sheet, row, 14).Value2),
            "Iva12": to_number(cell(note_sheet, row, 15).Value2),
            "Interest": to_number(cell(note_sheet, row, 16).Value2),
            "Total": to_number(cell(note_sheet, row, 17).Value2),
            "Anticipo": to_number(cell(note_sheet, row, 18).Value2),
            "Neto": to_number(cell(note_sheet, row, 19).Value2),
            "Asiento": normalize_text(cell(note_sheet, row, 20).Text),
            "GarExt": normalize_text(cell(note_sheet, row, 21).Text),
        }
        add_lookup_entry(note_store, doc_key, order_key, entry)

    last_rep_vtas_row = int(cell(rep_vtas_sheet, rep_vtas_sheet.Rows.Count, 3).End(XL_DIRECTION_UP).Row)
    for row in range(15, last_rep_vtas_row + 1):
        doc_key = trim_document(cell(rep_vtas_sheet, row, 8).Text)
        order_key = normalize_text(cell(rep_vtas_sheet, row, 3).Text)
        if doc_key == "" and order_key == "":
            continue
        entry = {
            "RowOrder": row,
            "Agency": normalize_text(cell(rep_vtas_sheet, row, 1).Text),
            "Center": normalize_text(cell(rep_vtas_sheet, row, 2).Text),
            "Order": order_key,
            "Advisor": normalize_text(cell(rep_vtas_sheet, row, 4).Text),
            "Line": normalize_text(cell(rep_vtas_sheet, row, 5).Text),
            "Cedula": normalize_text(cell(rep_vtas_sheet, row, 6).Text),
            "Customer": normalize_text(cell(rep_vtas_sheet, row, 7).Text),
            "DocumentRaw": normalize_text(cell(rep_vtas_sheet, row, 8).Text),
            "DateFactValue": get_date_value(rep_vtas_sheet, row, 9),
            "DateNoteValue": get_date_value(rep_vtas_sheet, row, 10),
            "GarExt": normalize_text(cell(rep_vtas_sheet, row, 27).Text),
        }
        add_lookup_entry(rep_vtas_store, doc_key, order_key, entry)

    return {"Invoice": invoice_store, "Note": note_store, "RepVtas": rep_vtas_store, "Defaults": defaults}


def read_precont_ventas_prototypes(worksheet: Any) -> list[dict[str, Any]]:
    used_range = worksheet.UsedRange
    last_row = int(used_range.Row + used_range.Rows.Count - 1)
    prototypes: list[dict[str, Any]] = []
    for row in range(2, last_row + 1):
        account_text = normalize_text(cell(worksheet, row, 5).Text)
        account_digits = "".join(ch for ch in account_text if ch.isdigit())
        account = account_digits.rjust(12, "0") if len(account_digits) >= 10 else account_text
        doc = normalize_text(cell(worksheet, row, 3).Text)
        if account == "" or doc == "":
            continue
        prototypes.append(
            {
                "TemplateRow": row,
                "Ag": normalize_text(cell(worksheet, row, 2).Text),
                "Doc": doc,
                "Line": normalize_text(cell(worksheet, row, 4).Text),
                "Account": account,
                "Description": normalize_text(cell(worksheet, row, 6).Text),
                "CostCenter": normalize_text(cell(worksheet, row, 7).Text),
                "Asiento": normalize_text(cell(worksheet, row, 10).Text),
            }
        )
    return prototypes


def read_precont_costos2_prototypes(worksheet: Any) -> list[dict[str, Any]]:
    used_range = worksheet.UsedRange
    last_row = int(used_range.Row + used_range.Rows.Count - 1)
    prototypes: list[dict[str, Any]] = []
    for row in range(2, last_row + 1):
        account_text = normalize_text(cell(worksheet, row, 4).Text)
        account_digits = "".join(ch for ch in account_text if ch.isdigit())
        account = account_digits.rjust(12, "0") if len(account_digits) >= 10 else account_text
        if account == "":
            continue
        prototypes.append(
            {
                "TemplateRow": row,
                "Ag": normalize_text(cell(worksheet, row, 2).Text),
                "Line": normalize_text(cell(worksheet, row, 3).Text),
                "Account": account,
                "Number": normalize_text(cell(worksheet, row, 5).Text),
                "Description": normalize_text(cell(worksheet, row, 6).Text),
                "CostCenter": normalize_text(cell(worksheet, row, 7).Text),
                "Asiento": normalize_text(cell(worksheet, row, 10).Text),
            }
        )
    return prototypes


def clear_output_sheet(worksheet: Any, start_row: int, last_column: str) -> None:
    try:
        constants = worksheet.Range(f"A{start_row}:{last_column}65536").SpecialCells(XLCELLTYPE_CONSTANTS)
        constants.ClearContents()
    except pywintypes.com_error:
        return


def clear_worksheet_range_contents(worksheet: Any, start_column: str, start_row: int, end_column: str, end_row: int) -> None:
    if worksheet is None or start_row > end_row:
        return
    worksheet.Range(f"{start_column}{start_row}:{end_column}{end_row}").ClearContents()


def write_rows_to_worksheet(worksheet: Any, rows: list[list[Any]], start_row: int = 1, start_column: int = 1) -> None:
    if not rows:
        return
    max_columns = max(len(row) for row in rows)
    normalized_rows = [tuple(list(row) + [None] * (max_columns - len(row))) for row in rows]
    start_cell = worksheet.Cells(start_row, start_column)
    end_cell = worksheet.Cells(start_row + len(normalized_rows) - 1, start_column + max_columns - 1)
    worksheet.Range(start_cell, end_cell).Value = tuple(normalized_rows)


def get_date_column_write_mode(worksheet: Any, row: int, column: int) -> str:
    number_format = normalize_text(cell(worksheet, row, column).NumberFormat)
    if number_format in {"", "General"}:
        return "text"
    return "serial"


def get_date_matrix_value(value: Any, mode: str = "serial", text_format: str = "%d/%m/%Y") -> Any:
    if value in (None, ""):
        return None
    date_serial = float(value)
    if mode == "text":
        base_date = datetime(1899, 12, 30)
        return "'" + (base_date + timedelta(days=date_serial)).strftime(text_format)
    return date_serial


def numeric_date_text(value: float, text_format: str = "%d/%m/%Y") -> str:
    base_date = datetime(1899, 12, 30)
    return (base_date + timedelta(days=float(value))).strftime(text_format)


def get_numeric_matrix_value(value: Any, blank_if_zero: bool = False) -> Any:
    rounded = float(round_amount(to_number(value)))
    if blank_if_zero and round_amount(abs(rounded)) == 0:
        return None
    return rounded


def set_date_cell_value(worksheet: Any, row: int, column: int, value: Any, context: str = "") -> None:
    target_cell = cell(worksheet, row, column)
    if value in (None, ""):
        target_cell.ClearContents()
        return
    date_serial = float(value)
    current_format = normalize_text(target_cell.NumberFormat)
    base_date = datetime(1899, 12, 30)
    date_text = (base_date + timedelta(days=date_serial)).strftime("%d/%m/%Y")
    if current_format in {"", "General"}:
        target_cell.NumberFormat = "General"
        target_cell.Value2 = "'" + date_text
    else:
        target_cell.Value = base_date + timedelta(days=date_serial)


def set_numeric_cell_safe(worksheet: Any, row: int, column: int, value: float, blank_if_zero: bool = False) -> None:
    rounded = float(round_amount(value))
    target_cell = cell(worksheet, row, column)
    if blank_if_zero and round_amount(abs(rounded)) == 0:
        target_cell.ClearContents()
        return
    try:
        target_cell.Value2 = rounded
    except pywintypes.com_error:
        target_cell.Value2 = str(rounded)


def get_indexed_row_value(row: list[Any], index: int) -> str:
    if row is None or index >= len(row):
        return ""
    return normalize_text(row[index])


def convert_px_rows_to_detail_rows(rows: list[list[Any]]) -> list[dict[str, Any]]:
    detail_rows: list[dict[str, Any]] = []
    for row in rows:
        ag = get_indexed_row_value(row, 1)
        factura = get_indexed_row_value(row, 4)
        item = get_indexed_row_value(row, 15)
        if ag == "" or factura == "" or item == "":
            continue
        if not ag.isdigit():
            continue
        detail_rows.append(
            {
                "Agencia": ag,
                "Estado": get_indexed_row_value(row, 2),
                "Orden": get_indexed_row_value(row, 3),
                "Factura": factura,
                "Fecha": get_indexed_row_value(row, 6),
                "PxNo": get_indexed_row_value(row, 8),
                "Cuenta": get_indexed_row_value(row, 9),
                "Fr": get_indexed_row_value(row, 11),
                "Codigo": get_indexed_row_value(row, 13),
                "Item": item,
                "PvpBruto": to_number(get_indexed_row_value(row, 19)),
                "DescPct": to_number(get_indexed_row_value(row, 21)),
                "DescValor": to_number(get_indexed_row_value(row, 22)),
                "PvpNeto": to_number(get_indexed_row_value(row, 23)),
                "Costo": to_number(get_indexed_row_value(row, 24)),
                "Origen": get_indexed_row_value(row, 25),
            }
        )
    return detail_rows


def get_px_detail_ranges(worksheet: Any) -> list[dict[str, int]]:
    ranges: list[dict[str, int]] = []
    seen: set[str] = set()
    used_range = worksheet.UsedRange
    last_row = int(used_range.Row + used_range.Rows.Count - 1)
    for row in range(1, last_row + 1):
        formula_text = normalize_text(cell(worksheet, row, 12).Formula)
        if formula_text == "":
            continue
        normalized_formula = formula_text.upper().replace("$", "")
        match = re.search(r"SUBTOTAL\(9,L(\d+):L(\d+)\)", normalized_formula)
        if match is None:
            continue
        start_row = int(match.group(1))
        end_row = int(match.group(2))
        key = f"{start_row}:{end_row}"
        if key in seen:
            continue
        seen.add(key)
        ranges.append({"StartRow": start_row, "EndRow": end_row})
    return sorted(ranges, key=lambda item: item["StartRow"])


def mayor_brand_allows_flexible_account_mapping(brand_key: str) -> bool:
    return normalize_text(brand_key).lower() in {"tyt"}


def get_sales_account_family_key(account: Any, ignore_prefix: bool = False) -> str:
    compact = get_compact_account_code(account)
    match = re.fullmatch(r"(040101\d{2})(\d{4})", compact)
    if match is None:
        return normalize_text(account)
    prefix, suffix = match.groups()
    family_prefix = "*" if ignore_prefix else prefix
    if suffix in {"0001", "0002", "0003"}:
        return f"{family_prefix}|VENTA"
    if suffix in {"0010", "0011", "0012"}:
        return f"{family_prefix}|DESCUENTO"
    if suffix == "0014":
        return f"{family_prefix}|DEVOLUCION"
    return f"{family_prefix}|{suffix}"


def get_mayor_account_family_key(account: str, ignore_prefix: bool = False) -> str:
    return get_sales_account_family_key(account, ignore_prefix=ignore_prefix)


def get_sales_account_suffix_key(account: Any, ignore_prefix: bool = False) -> str:
    compact = get_compact_account_code(account)
    match = re.fullmatch(r"(040101\d{2})(\d{4})", compact)
    if match is None:
        return normalize_text(account)
    prefix, suffix = match.groups()
    return f"{'*' if ignore_prefix else prefix}|{suffix}"


def get_mayor_account_suffix_key(account: str, ignore_prefix: bool = False) -> str:
    return get_sales_account_suffix_key(account, ignore_prefix=ignore_prefix)


def classify_sales_control_bucket(account: Any, name: Any) -> str:
    account_text = normalize_text(account)
    name_text = normalize_text(name).upper()
    compact = get_compact_account_code(account_text)
    if account_text == "" and name_text == "":
        return ""
    if re.fullmatch(r"010105\d{2}\d{4}", compact) or "GARANT" in name_text:
        return "guarantee"
    if re.fullmatch(r"040101\d{2}\d{4}", compact) is None:
        return ""
    suffix = compact[-4:]
    if suffix == "0014" or "DEVOL" in name_text:
        return "return"
    if suffix in {"0010", "0011", "0012"} or "DESC" in name_text:
        return "discount"
    if suffix in {"0001", "0002", "0003"} or "VTAS" in name_text:
        return "sales"
    return ""


def get_mayor_account_layout_bucket(account: Any, name: Any) -> str:
    bucket = classify_sales_control_bucket(account, name)
    return bucket if bucket in {"sales", "discount", "return"} else ""


def get_control_metrics_from_rows(
    rows: list[dict[str, Any]],
    *,
    account_key: str,
    name_key: str,
    debit_key: str,
    credit_key: str,
) -> dict[str, float]:
    metrics = {
        "InvoiceSales": 0.0,
        "InvoiceDiscounts": 0.0,
        "NoteSales": 0.0,
        "NoteDiscounts": 0.0,
        "NetSales": 0.0,
    }
    for row in rows:
        bucket = classify_sales_control_bucket(row.get(account_key), row.get(name_key))
        debit = to_number(row.get(debit_key))
        credit = to_number(row.get(credit_key))
        if bucket == "sales":
            metrics["InvoiceSales"] += credit
            metrics["NetSales"] += credit - debit
            continue
        if bucket == "discount":
            metrics["InvoiceDiscounts"] += debit
            metrics["NoteDiscounts"] += credit
            metrics["NetSales"] += credit - debit
            continue
        if bucket == "return":
            metrics["NoteSales"] += debit
            metrics["NetSales"] += credit - debit
    return {key: float(round_amount(value)) for key, value in metrics.items()}


def select_mayor_layout_candidate(
    candidates: list[dict[str, Any]],
    rows_by_layout_key: dict[str, list[dict[str, Any]]],
) -> dict[str, Any] | None:
    candidates.sort(key=lambda item: (-(int(item["EndRow"]) - int(item["StartRow"]) + 1), int(item["StartRow"])))
    if not candidates:
        return None
    for candidate in candidates:
        layout_key = f"{candidate['Account']}:{int(candidate['StartRow'])}-{int(candidate['EndRow'])}"
        existing = len(rows_by_layout_key.get(layout_key, []))
        capacity = int(candidate["EndRow"]) - int(candidate["StartRow"]) + 1
        if existing < capacity:
            return candidate
    return candidates[0]


def resolve_mayor_compatible_layout(
    account: str,
    name: Any,
    layouts: list[dict[str, Any]],
    rows_by_layout_key: dict[str, list[dict[str, Any]]],
    allow_cross_prefix_family: bool = False,
) -> dict[str, Any] | None:
    family_key = get_mayor_account_family_key(account)
    if family_key != "":
        exact_family_candidates = [
            layout
            for layout in layouts
            if get_mayor_account_family_key(str(layout["Account"])) == family_key
        ]
        selected = select_mayor_layout_candidate(exact_family_candidates, rows_by_layout_key)
        if selected is not None:
            return selected

    if allow_cross_prefix_family:
        suffix_key = get_mayor_account_suffix_key(account, ignore_prefix=True)
        if suffix_key != "":
            cross_prefix_suffix_candidates = [
                layout
                for layout in layouts
                if get_mayor_account_suffix_key(str(layout["Account"]), ignore_prefix=True) == suffix_key
            ]
            selected = select_mayor_layout_candidate(cross_prefix_suffix_candidates, rows_by_layout_key)
            if selected is not None:
                return selected

    layout_bucket = get_mayor_account_layout_bucket(account, name)
    if layout_bucket == "":
        return None
    cross_prefix_bucket_candidates = [
        layout
        for layout in layouts
        if (normalize_text(layout.get("BucketHint")) or get_mayor_account_layout_bucket(layout.get("Account"), layout.get("Name"))) == layout_bucket
    ]
    return select_mayor_layout_candidate(cross_prefix_bucket_candidates, rows_by_layout_key)


def test_mayor_px_adjustment_row(row: dict[str, Any]) -> bool:
    account = get_compact_account_code(row.get("account"))
    origin = normalize_text(row.get("origin")).upper()
    seat = normalize_text(row.get("seat"))
    detail = normalize_text(row.get("detail")).upper()
    if re.fullmatch(r"040101\d{2}(0003|0012)", account or "") is None:
        return False
    if "REGISTRO DE PX AJUSTE DE EGRESO" in detail:
        return True
    return origin == "AGCM" and seat == "435"


def filter_mayor_rows_for_workbook(rows: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    kept: list[dict[str, Any]] = []
    removed: list[dict[str, Any]] = []
    balance_adjustments: dict[str, float] = {}
    for row in rows:
        account = get_compact_account_code(row.get("account"))
        balance_adjustments.setdefault(account, 0.0)
        if test_mayor_px_adjustment_row(row):
            balance_adjustments[account] += float(to_number(row.get("credit"))) - float(to_number(row.get("debit")))
            removed.append(row)
            continue
        clone = dict(row)
        adjusted_balance = float(to_number(row.get("balance"))) + float(balance_adjustments[account])
        clone["effective_balance"] = float(round_amount(adjusted_balance))
        kept.append(clone)
    return {"Rows": kept, "Removed": removed}


def get_mayor_sheet_section_layouts(worksheet: Any) -> list[dict[str, Any]]:
    layouts: list[dict[str, Any]] = []
    summary_ranges: list[dict[str, int]] = []
    used_range = worksheet.UsedRange
    last_row = int(used_range.Row + used_range.Rows.Count - 1)
    for row in range(1, last_row + 1):
        formula_text = normalize_text(cell(worksheet, row, 9).Formula)
        if formula_text == "":
            continue
        normalized_formula = formula_text.upper().replace("$", "")
        match = re.search(r"SUBTOTAL\(9,I(\d+):I(\d+)\)", normalized_formula)
        if match is None:
            continue
        summary_ranges.append({"StartRow": int(match.group(1)), "EndRow": int(match.group(2))})

    if not summary_ranges:
        return []

    account_starts: list[dict[str, Any]] = []
    seen_accounts: set[str] = set()
    for row in range(1, last_row + 1):
        account = normalize_text(cell(worksheet, row, 1).Text)
        if re.fullmatch(r"\d{2}\.\d{2}\.\d{2}\.\d{2}\.\d{4}", account or "") is None or account in seen_accounts:
            continue
        parent_range = next((item for item in summary_ranges if item["StartRow"] <= row <= item["EndRow"]), None)
        if parent_range is None:
            continue
        seen_accounts.add(account)
        account_starts.append(
            {
                "Account": account,
                "Name": normalize_text(cell(worksheet, row, 2).Text),
                "StartRow": row,
                "ParentStartRow": parent_range["StartRow"],
                "ParentEndRow": parent_range["EndRow"],
            }
        )

    ordered_starts = sorted(account_starts, key=lambda item: int(item["StartRow"]))
    for index, current in enumerate(ordered_starts):
        end_row = int(current["ParentEndRow"])
        for next_item in ordered_starts[index + 1 :]:
            if int(next_item["ParentStartRow"]) != int(current["ParentStartRow"]):
                break
            end_row = int(next_item["StartRow"]) - 1
            break
        layouts.append(
            {
                "Account": current["Account"],
                "Name": current["Name"],
                "StartRow": int(current["StartRow"]),
                "EndRow": end_row,
                "ParentStartRow": int(current["ParentStartRow"]),
                "ParentEndRow": int(current["ParentEndRow"]),
                "BucketHint": get_mayor_account_layout_bucket(current["Account"], current["Name"]),
            }
        )

    occupied_parent_ranges = {
        (int(item["ParentStartRow"]), int(item["ParentEndRow"]))
        for item in ordered_starts
    }
    for summary_range in summary_ranges:
        parent_key = (int(summary_range["StartRow"]), int(summary_range["EndRow"]))
        if parent_key in occupied_parent_ranges:
            continue
        layouts.append(
            {
                "Account": f"__AUTO_RETURN__{int(summary_range['StartRow'])}_{int(summary_range['EndRow'])}",
                "Name": "AUTO DEVOLUCIONES",
                "StartRow": int(summary_range["StartRow"]),
                "EndRow": int(summary_range["EndRow"]),
                "ParentStartRow": int(summary_range["StartRow"]),
                "ParentEndRow": int(summary_range["EndRow"]),
                "BucketHint": "return",
            }
        )
    return sorted(layouts, key=lambda item: int(item["StartRow"]))


def get_brand_period_date_value(rows: list[dict[str, Any]], mayor_rows: list[dict[str, Any]]) -> float:
    max_date = None
    for row in rows:
        for candidate in (row.get("DateFactValue"), row.get("DateNoteValue")):
            if candidate in (None, ""):
                continue
            candidate_value = float(candidate)
            if max_date is None or candidate_value > max_date:
                max_date = candidate_value
    for row in mayor_rows:
        candidate = row.get("date_value")
        if candidate in (None, ""):
            continue
        candidate_value = float(candidate)
        if max_date is None or candidate_value > max_date:
            max_date = candidate_value
    if max_date is None:
        return float((datetime.now() - datetime(1899, 12, 30)).days)
    return float(max_date)


def get_period_year_month(period_date_value: Any) -> tuple[int, int]:
    try:
        period_date = datetime(1899, 12, 30) + timedelta(days=float(period_date_value))
    except Exception:
        period_date = datetime.now()
    return period_date.year, period_date.month


def build_posting_document_meta(rows_posting: list[dict[str, Any]]) -> dict[str, dict[str, str]]:
    metadata: dict[str, dict[str, str]] = {}
    for row in rows_posting:
        doc_type = normalize_text(row.get("DocType")).upper()
        if doc_type not in {"FA", "FC"}:
            continue
        document_key = trim_document(row.get("DocumentTrim"))
        if document_key == "" or document_key in metadata:
            continue
        metadata[document_key] = {
            "DocType": doc_type,
            "Center": get_normalized_center_code(row.get("Center")) or "00",
        }
    return metadata


def build_supplemental_mayor_rows_from_source(
    layouts: list[dict[str, Any]],
    mayor_rows: list[dict[str, Any]],
    rows_display: list[dict[str, Any]],
    rows_posting: list[dict[str, Any]],
    px_rows: list[list[Any]],
    period_date_value: Any,
) -> list[dict[str, Any]]:
    existing_accounts = {get_compact_account_code(row.get("account")) for row in mayor_rows}
    target_layouts: dict[str, dict[str, Any]] = {}
    for layout in layouts:
        compact = get_compact_account_code(layout.get("Account"))
        suffix = compact[-4:] if len(compact) >= 4 else ""
        if suffix in {"0010", "0012", "0014"} and compact not in existing_accounts:
            target_layouts[suffix] = layout
    if not target_layouts:
        return []

    invoice_meta_by_document = build_posting_document_meta(rows_posting)
    invoice_discount_totals: dict[str, dict[str, float]] = {"0010": {}, "0012": {}}
    note_discount_credit_totals: dict[str, dict[str, float]] = {"0010": {}, "0012": {}}
    note_sales_totals: dict[str, float] = {}

    for row in rows_display:
        doc_type = normalize_text(row.get("DocType")).upper()
        if doc_type in {"FA", "FC"}:
            document_key = trim_document(row.get("DocumentTrim"))
            metadata = invoice_meta_by_document.get(document_key, {})
            actual_doc_type = normalize_text(metadata.get("DocType")).upper() or doc_type
            center = get_normalized_center_code(metadata.get("Center")) or get_normalized_center_code(row.get("Center")) or "00"
            account_suffix = "0010" if actual_doc_type == "FA" else "0012" if actual_doc_type == "FC" else ""
            if account_suffix not in target_layouts:
                continue
            discount_amount = float(round_amount(get_invoice_source_amounts(row)["Discount"]))
            if discount_amount == 0.0:
                continue
            invoice_discount_totals[account_suffix][center] = float(invoice_discount_totals[account_suffix].get(center, 0.0)) + discount_amount
            continue

        if doc_type not in {"DC", "DE"}:
            continue
        affected_key = trim_document(row.get("AffectedDocumentTrim"))
        metadata = invoice_meta_by_document.get(affected_key, {})
        affected_doc_type = normalize_text(metadata.get("DocType")).upper()
        center = get_normalized_center_code(metadata.get("Center")) or get_normalized_center_code(row.get("Center")) or "00"
        amounts = get_note_source_amounts(row)
        note_sales_amount = float(round_amount(amounts["Discount"] + amounts["NetoSinIva"] + amounts["NetoConIva"]))
        if note_sales_amount != 0.0 and "0014" in target_layouts:
            note_sales_totals[center] = float(note_sales_totals.get(center, 0.0)) + note_sales_amount
        discount_suffix = "0010" if affected_doc_type == "FA" else "0012"
        if discount_suffix in target_layouts:
            discount_credit = float(round_amount(amounts["Discount"]))
            if discount_credit != 0.0:
                note_discount_credit_totals[discount_suffix][center] = float(note_discount_credit_totals[discount_suffix].get(center, 0.0)) + discount_credit

    px_discount_totals: dict[str, dict[str, float]] = {"FA": {}, "FC": {}}
    for px_row in convert_px_rows_to_detail_rows(px_rows):
        document_key = trim_document(px_row.get("Factura"))
        metadata = invoice_meta_by_document.get(document_key)
        if not metadata:
            continue
        actual_doc_type = normalize_text(metadata.get("DocType")).upper()
        if actual_doc_type not in {"FA", "FC"}:
            continue
        center = get_normalized_center_code(metadata.get("Center")) or "00"
        px_discount_totals[actual_doc_type][center] = float(px_discount_totals[actual_doc_type].get(center, 0.0)) + float(to_number(px_row.get("DescValor")))

    for suffix, doc_type in (("0010", "FA"), ("0012", "FC")):
        if suffix not in target_layouts:
            continue
        for center, discount_total in px_discount_totals[doc_type].items():
            adjusted_total = float(invoice_discount_totals[suffix].get(center, 0.0)) - float(discount_total)
            invoice_discount_totals[suffix][center] = float(round_amount(adjusted_total))

    year_value, month_value = get_period_year_month(period_date_value)
    generated: list[dict[str, Any]] = []
    for suffix in ("0010", "0012", "0014"):
        layout = target_layouts.get(suffix)
        if layout is None:
            continue
        account_rows: list[dict[str, Any]] = []
        if suffix in {"0010", "0012"}:
            for center in sorted(invoice_discount_totals[suffix].keys()):
                amount = float(round_amount(invoice_discount_totals[suffix][center]))
                if amount == 0.0:
                    continue
                account_rows.append(
                    {
                        "account": format_account_code(layout.get("Account")),
                        "name": normalize_text(layout.get("Name")),
                        "ext": "N",
                        "date_value": float(period_date_value),
                        "origin": "VENSE",
                        "seat": "400",
                        "reference": "",
                        "detail": f"CONTA VENTAS - CENTRO{center} PERIODO {year_value} - {month_value:02d}",
                        "debit": amount,
                        "credit": 0.0,
                    }
                )
            for center in sorted(note_discount_credit_totals[suffix].keys()):
                amount = float(round_amount(note_discount_credit_totals[suffix][center]))
                if amount == 0.0:
                    continue
                account_rows.append(
                    {
                        "account": format_account_code(layout.get("Account")),
                        "name": normalize_text(layout.get("Name")),
                        "ext": "N",
                        "date_value": float(period_date_value),
                        "origin": "VENSE",
                        "seat": "401",
                        "reference": "",
                        "detail": f"CONTABILIZACION VENTAS - CENTRO{center} PERIODO {year_value} - {month_value:02d}",
                        "debit": 0.0,
                        "credit": amount,
                    }
                )
        else:
            for center in sorted(note_sales_totals.keys()):
                amount = float(round_amount(note_sales_totals[center]))
                if amount == 0.0:
                    continue
                account_rows.append(
                    {
                        "account": format_account_code(layout.get("Account")),
                        "name": normalize_text(layout.get("Name")),
                        "ext": "N",
                        "date_value": float(period_date_value),
                        "origin": "VENSE",
                        "seat": "401",
                        "reference": "",
                        "detail": f"CONTABILIZACION VENTAS - CENTRO{center} PERIODO {year_value} - {month_value:02d}",
                        "debit": amount,
                        "credit": 0.0,
                    }
                )
        running_balance = 0.0
        for item in account_rows:
            running_balance = float(round_amount(running_balance + float(to_number(item.get("debit"))) - float(to_number(item.get("credit")))))
            item["balance"] = running_balance
            generated.append(item)
    return generated


def add_generated_template_row(target: list[dict[str, Any]], prototype: dict[str, Any] | None, debit: float = 0.0, credit: float = 0.0) -> None:
    if prototype is None:
        return
    target.append(
        {
            "Ag": normalize_text(prototype.get("Ag")),
            "Doc": normalize_text(prototype.get("Doc")),
            "Line": normalize_text(prototype.get("Line")),
            "Account": normalize_text(prototype.get("Account")),
            "Description": normalize_text(prototype.get("Description")),
            "CostCenter": normalize_text(prototype.get("CostCenter")),
            "Debit": float(round_amount(debit)),
            "Credit": float(round_amount(credit)),
            "Asiento": normalize_text(prototype.get("Asiento")),
        }
    )


def add_template_aggregate_rows(target: list[dict[str, Any]], prototypes: list[dict[str, Any]], totals_by_center: dict[str, float], amount_side: str = "Debit", ensure_at_least_one_row: bool = False) -> None:
    if not prototypes:
        return
    written = False
    for center in sorted(totals_by_center.keys()):
        prototype = next((item for item in prototypes if normalize_text(item.get("Ag")) == center), prototypes[0])
        amount = float(round_amount(totals_by_center[center]))
        if amount_side == "Credit":
            add_generated_template_row(target, prototype, 0.0, amount)
        else:
            add_generated_template_row(target, prototype, amount, 0.0)
        written = True
    if not written and ensure_at_least_one_row:
        add_generated_template_row(target, prototypes[0], 0.0, 0.0)


def add_template_sequential_rows_from_mayor(target: list[dict[str, Any]], prototypes: list[dict[str, Any]], mayor_rows: list[dict[str, Any]], ensure_at_least_one_row: bool = False) -> None:
    if not prototypes:
        return
    items = list(mayor_rows)
    if not items and ensure_at_least_one_row:
        add_generated_template_row(target, prototypes[0], 0.0, 0.0)
        return
    for index, item in enumerate(items):
        prototype = prototypes[index] if index < len(prototypes) else prototypes[-1]
        add_generated_template_row(target, prototype, to_number(item.get("debit")), to_number(item.get("credit")))


def select_precont_target_account(
    candidates: list[str],
    assigned_group_counts: dict[str, int],
    assigned_row_counts: dict[str, int],
) -> str | None:
    if not candidates:
        return None
    ordered = sorted(
        candidates,
        key=lambda account: (
            int(assigned_group_counts.get(account, 0)),
            int(assigned_row_counts.get(account, 0)),
            account,
        ),
    )
    return ordered[0]


def resolve_precont_mayor_target_account(
    source_account: str,
    source_name: Any,
    critical_accounts: list[str],
    assigned_group_counts: dict[str, int],
    assigned_row_counts: dict[str, int],
    allow_cross_prefix_family: bool = False,
) -> str | None:
    if source_account in critical_accounts:
        return source_account

    exact_family_key = get_sales_account_family_key(source_account)
    if exact_family_key != "":
        exact_family_candidates = [
            account
            for account in critical_accounts
            if get_sales_account_family_key(account) == exact_family_key
        ]
        selected = select_precont_target_account(exact_family_candidates, assigned_group_counts, assigned_row_counts)
        if selected is not None:
            return selected

    if allow_cross_prefix_family:
        suffix_key = get_sales_account_suffix_key(source_account, ignore_prefix=True)
        if suffix_key != "":
            suffix_candidates = [
                account
                for account in critical_accounts
                if get_sales_account_suffix_key(account, ignore_prefix=True) == suffix_key
            ]
            selected = select_precont_target_account(suffix_candidates, assigned_group_counts, assigned_row_counts)
            if selected is not None:
                return selected

        family_key = get_sales_account_family_key(source_account, ignore_prefix=True)
        if family_key != "":
            family_candidates = [
                account
                for account in critical_accounts
                if get_sales_account_family_key(account, ignore_prefix=True) == family_key
            ]
            selected = select_precont_target_account(family_candidates, assigned_group_counts, assigned_row_counts)
            if selected is not None:
                return selected

    source_bucket = classify_sales_control_bucket(source_account, source_name)
    if source_bucket == "":
        return None
    bucket_candidates = [
        account
        for account in critical_accounts
        if classify_sales_control_bucket(account, "") == source_bucket
    ]
    return select_precont_target_account(bucket_candidates, assigned_group_counts, assigned_row_counts)


def map_mayor_rows_to_precont_critical_accounts(
    critical_accounts: list[str],
    mayor_rows: list[dict[str, Any]],
    brand_key: str = "",
) -> dict[str, Any]:
    rows_by_source_account: dict[str, list[dict[str, Any]]] = {}
    for mayor_row in mayor_rows:
        account = get_compact_account_code(mayor_row.get("account"))
        if account == "":
            continue
        rows_by_source_account.setdefault(account, []).append(mayor_row)

    target_rows: dict[str, list[dict[str, Any]]] = {account: [] for account in critical_accounts}
    assigned_group_counts: dict[str, int] = {account: 0 for account in critical_accounts}
    assigned_row_counts: dict[str, int] = {account: 0 for account in critical_accounts}
    compatible_mapped_accounts: list[str] = []
    unmapped_accounts: list[str] = []
    allow_cross_prefix_family = mayor_brand_allows_flexible_account_mapping(brand_key)

    for source_account in sorted(rows_by_source_account.keys()):
        source_rows = rows_by_source_account.get(source_account, [])
        source_name = source_rows[0].get("name") if source_rows else ""
        target_account = resolve_precont_mayor_target_account(
            source_account,
            source_name,
            critical_accounts,
            assigned_group_counts,
            assigned_row_counts,
            allow_cross_prefix_family=allow_cross_prefix_family,
        )
        if target_account is None:
            unmapped_accounts.append(source_account)
            continue
        if target_account != source_account:
            compatible_mapped_accounts.append(source_account)
        target_rows.setdefault(target_account, []).extend(source_rows)
        assigned_group_counts[target_account] = int(assigned_group_counts.get(target_account, 0)) + 1
        assigned_row_counts[target_account] = int(assigned_row_counts.get(target_account, 0)) + len(source_rows)

    return {
        "RowsByAccount": target_rows,
        "CompatibleMappedAccounts": compatible_mapped_accounts,
        "UnmappedAccounts": unmapped_accounts,
    }


def new_precont_ventas_generated_rows(
    prototypes: list[dict[str, Any]],
    rows_posting: list[dict[str, Any]],
    mayor_rows: list[dict[str, Any]],
    px_rows: list[list[Any]],
    brand_key: str = "",
) -> list[dict[str, Any]]:
    generated: list[dict[str, Any]] = []
    invoice_totals_by_doc_center: dict[str, dict[str, float]] = {"FA": {}, "FC": {}}
    iva_totals_by_doc_center: dict[str, dict[str, float]] = {"FA": {}, "FC": {}}
    cash_notes_by_center: dict[str, float] = {}
    invoice_doc_type_by_document: dict[str, str] = {}
    invoice_info_by_document: dict[str, dict[str, str]] = {}
    px_gross_by_doc_center: dict[str, dict[str, float]] = {"FA": {}, "FC": {}}
    px_discount_by_doc_center: dict[str, dict[str, float]] = {"FA": {}, "FC": {}}

    for row in [item for item in rows_posting if normalize_text(item.get("DocType")).upper() in {"FA", "FC"}]:
        doc_type = normalize_text(row.get("DocType")).upper()
        center = get_normalized_center_code(row.get("Center")) or "00"
        amounts = get_invoice_source_amounts(row)
        invoice_totals_by_doc_center[doc_type][center] = float(invoice_totals_by_doc_center[doc_type].get(center, 0.0)) + float(amounts["Total"])
        iva_totals_by_doc_center[doc_type][center] = float(iva_totals_by_doc_center[doc_type].get(center, 0.0)) + float(amounts["Iva"])
        document_key = trim_document(row.get("DocumentTrim"))
        if document_key and document_key not in invoice_doc_type_by_document:
            invoice_doc_type_by_document[document_key] = doc_type
        if document_key and document_key not in invoice_info_by_document:
            invoice_info_by_document[document_key] = {"DocType": doc_type, "Center": center}

    for row in [item for item in rows_posting if normalize_text(item.get("DocType")).upper() in {"DC", "DE"}]:
        center = get_normalized_center_code(row.get("Center")) or "00"
        affected_key = trim_document(row.get("AffectedDocumentTrim"))
        if affected_key == "" or invoice_doc_type_by_document.get(affected_key) != "FA":
            continue
        amounts = get_note_source_amounts(row)
        cash_notes_by_center[center] = float(cash_notes_by_center.get(center, 0.0)) + float(amounts["Total"])

    for px_row in convert_px_rows_to_detail_rows(px_rows):
        document_key = trim_document(px_row.get("Factura"))
        invoice_info = invoice_info_by_document.get(document_key)
        if not invoice_info:
            continue
        doc_type = normalize_text(invoice_info.get("DocType")).upper()
        if doc_type not in {"FA", "FC"}:
            continue
        center = get_normalized_center_code(invoice_info.get("Center")) or "00"
        px_gross_by_doc_center[doc_type][center] = float(px_gross_by_doc_center[doc_type].get(center, 0.0)) + float(to_number(px_row.get("PvpBruto")))
        px_discount_by_doc_center[doc_type][center] = float(px_discount_by_doc_center[doc_type].get(center, 0.0)) + float(to_number(px_row.get("DescValor")))

    client_fa_prototypes = [item for item in prototypes if item["Doc"] == "FA" and re.search(r"^CLIENTES SERV", item["Description"] or "")]
    client_fc_prototypes = [item for item in prototypes if item["Doc"] == "FC" and re.search(r"^CLIENTES SERV", item["Description"] or "")]
    client_ca_prototypes = [item for item in prototypes if item["Doc"] == "CA" and re.search(r"^CLIENTES SERV", item["Description"] or "")]
    iva_fa_prototypes = [item for item in prototypes if item["Doc"] == "FA" and re.search(r"IVA", item["Description"] or "")]
    iva_fc_prototypes = [item for item in prototypes if item["Doc"] == "FC" and re.search(r"IVA", item["Description"] or "")]
    reserve_ca_prototypes = [item for item in prototypes if item["Doc"] == "CA" and re.search(r"RESERVA", item["Description"] or "")]
    liquidar_fa_credit_prototypes = [item for item in prototypes if item["Doc"] == "FA" and re.fullmatch(r"020120\d{2}0002", item["Account"] or "")]
    liquidar_fa_discount_prototypes = [item for item in prototypes if item["Doc"] == "FA" and re.fullmatch(r"020120\d{2}0004", item["Account"] or "")]
    liquidar_fc_credit_prototypes = [item for item in prototypes if item["Doc"] == "FC" and re.fullmatch(r"020120\d{2}0002", item["Account"] or "")]
    liquidar_fc_discount_prototypes = [item for item in prototypes if item["Doc"] == "FC" and re.fullmatch(r"020120\d{2}0004", item["Account"] or "")]

    add_template_aggregate_rows(generated, client_fa_prototypes, invoice_totals_by_doc_center["FA"], "Debit")
    add_template_aggregate_rows(generated, iva_fa_prototypes, iva_totals_by_doc_center["FA"], "Credit")
    add_template_aggregate_rows(generated, client_fc_prototypes, invoice_totals_by_doc_center["FC"], "Debit")
    add_template_aggregate_rows(generated, iva_fc_prototypes, iva_totals_by_doc_center["FC"], "Credit")
    add_template_aggregate_rows(generated, client_ca_prototypes, cash_notes_by_center, "Credit")
    add_template_aggregate_rows(generated, reserve_ca_prototypes, cash_notes_by_center, "Debit", ensure_at_least_one_row=True)
    add_template_aggregate_rows(generated, liquidar_fa_credit_prototypes, px_gross_by_doc_center["FA"], "Credit", ensure_at_least_one_row=True)
    add_template_aggregate_rows(generated, liquidar_fa_discount_prototypes, px_discount_by_doc_center["FA"], "Debit", ensure_at_least_one_row=True)
    add_template_aggregate_rows(generated, liquidar_fc_credit_prototypes, px_gross_by_doc_center["FC"], "Credit", ensure_at_least_one_row=True)
    add_template_aggregate_rows(generated, liquidar_fc_discount_prototypes, px_discount_by_doc_center["FC"], "Debit", ensure_at_least_one_row=True)

    critical_accounts = sorted(
        {
            item["Account"]
            for item in prototypes
            if re.fullmatch(r"040101\d{2}(0001|0003|0010|0012|0014)", item.get("Account", "") or "")
        }
    )
    mayor_mapping = map_mayor_rows_to_precont_critical_accounts(critical_accounts, mayor_rows, brand_key)
    for source_account in mayor_mapping["CompatibleMappedAccounts"]:
        print(f"WARN|precont_ventas_mayor_compatible_account|{brand_key}|source={source_account}")
    for source_account in mayor_mapping["UnmappedAccounts"]:
        print(f"WARN|precont_ventas_mayor_unmapped_account|{brand_key}|source={source_account}")

    for account in critical_accounts:
        account_prototypes = sorted([item for item in prototypes if item["Account"] == account], key=lambda item: int(item["TemplateRow"]))
        mayor_group = list((mayor_mapping["RowsByAccount"] or {}).get(account, []))
        add_template_sequential_rows_from_mayor(generated, account_prototypes, mayor_group, ensure_at_least_one_row=True)

    return generated


def get_brand_cost_metrics(rows: list[dict[str, Any]]) -> dict[str, float]:
    metrics = {
        "Costo": 0.0,
        "CostoLubricantes": 0.0,
        "CostoAccesorios": 0.0,
        "CostoRepuestos": 0.0,
        "CostoPintura": 0.0,
        "CostoSubconNc": 0.0,
    }
    for row in rows:
        for key in metrics.keys():
            metrics[key] += float(to_number(row.get(key)))
    return {key: float(round_amount(value)) for key, value in metrics.items()}


def write_precont_ventas_generated_rows(worksheet: Any, rows: list[dict[str, Any]]) -> int:
    clear_worksheet_range_contents(worksheet, "B", 2, "J", 1412)
    sheet_rows: list[list[Any]] = []
    for row in rows:
        sheet_rows.append(
            [
                get_excel_text_literal(row.get("Ag")),
                get_excel_text_literal(row.get("Doc")),
                get_excel_text_literal(row.get("Line")),
                get_excel_text_literal(row.get("Account")),
                get_excel_text_literal(row.get("Description")),
                get_excel_text_literal(row.get("CostCenter")),
                get_numeric_matrix_value(row.get("Debit"), blank_if_zero=True),
                get_numeric_matrix_value(row.get("Credit"), blank_if_zero=True),
                get_excel_text_literal(row.get("Asiento")),
            ]
        )
    if sheet_rows:
        write_rows_to_worksheet(worksheet, sheet_rows, 2, 2)
    return len(rows)


def write_precont_costos2_generated_rows(worksheet: Any, prototypes: list[dict[str, Any]], metrics: dict[str, float]) -> int:
    row_definitions = [
        {"Account": "050201010001", "Amount": float(metrics["CostoRepuestos"])},
        {"Account": "050201010002", "Amount": float(metrics["CostoLubricantes"])},
        {"Account": "050201010003", "Amount": float(metrics["Costo"] + metrics["CostoSubconNc"])},
        {"Account": "050201010004", "Amount": float(metrics["CostoPintura"])},
        {"Account": "050201010005", "Amount": 0.0},
        {"Account": "050201010008", "Amount": float(metrics["CostoAccesorios"])},
    ]
    clear_worksheet_range_contents(worksheet, "B", 2, "J", 51)
    sheet_rows: list[list[Any]] = []
    for definition in row_definitions:
        prototype = next((item for item in prototypes if item["Account"] == definition["Account"]), None)
        if prototype is None:
            continue
        sheet_rows.append(
            [
                get_excel_text_literal(prototype.get("Ag")),
                get_excel_text_literal(prototype.get("Line")),
                get_excel_text_literal(prototype.get("Account")),
                get_excel_text_literal(prototype.get("Number")),
                get_excel_text_literal(prototype.get("Description")),
                get_excel_text_literal(prototype.get("CostCenter")),
                get_numeric_matrix_value(definition["Amount"], blank_if_zero=True),
                None,
                get_excel_text_literal(prototype.get("Asiento")),
            ]
        )
    if sheet_rows:
        write_rows_to_worksheet(worksheet, sheet_rows, 2, 2)
    return len(sheet_rows)


def read_estadisticas_seed_rows(worksheet: Any, row_numbers: list[int]) -> dict[int, dict[str, str]]:
    seed_rows: dict[int, dict[str, str]] = {}
    for row_index in row_numbers:
        seed_rows[row_index] = {
            "Account": normalize_text(cell(worksheet, row_index, 1).Text),
            "Description": normalize_text(cell(worksheet, row_index, 2).Text),
            "Mod": normalize_text(cell(worksheet, row_index, 5).Text),
            "Asiento": normalize_text(cell(worksheet, row_index, 6).Text),
            "Detalle": normalize_text(cell(worksheet, row_index, 8).Text),
        }
    return seed_rows


def read_costo_seed_rows(worksheet: Any, row_numbers: list[int]) -> dict[int, dict[str, str]]:
    seed_rows: dict[int, dict[str, str]] = {}
    for row_index in row_numbers:
        seed_rows[row_index] = {
            "Account": normalize_text(cell(worksheet, row_index, 1).Text),
            "Description": normalize_text(cell(worksheet, row_index, 2).Text),
            "Mod": normalize_text(cell(worksheet, row_index, 4).Text),
            "Asiento": normalize_text(cell(worksheet, row_index, 5).Text),
            "Detalle": normalize_text(cell(worksheet, row_index, 7).Text),
        }
    return seed_rows


def write_estadisticas_generated_rows(worksheet: Any, metrics: dict[str, float], period_date_value: float) -> None:
    row_numbers = [6, 131, 222, 290, 365]
    seed_rows = read_estadisticas_seed_rows(worksheet, row_numbers)
    for start_row, end_row, end_column in [(6, 128, "J"), (131, 218, "J"), (222, 287, "J"), (290, 359, "J"), (365, 442, "J"), (446, 586, "K")]:
        try:
            constants = worksheet.Range(f"A{start_row}:{end_column}{end_row}").SpecialCells(XLCELLTYPE_CONSTANTS)
            constants.ClearContents()
        except pywintypes.com_error:
            pass

    block_rows = [
        {"Row": 6, "Amount": float(metrics["CostoRepuestos"])},
        {"Row": 131, "Amount": float(metrics["CostoLubricantes"])},
        {"Row": 222, "Amount": float(metrics["Costo"] + metrics["CostoSubconNc"])},
        {"Row": 290, "Amount": float(metrics["CostoPintura"])},
        {"Row": 365, "Amount": float(metrics["CostoAccesorios"])},
    ]
    for definition in block_rows:
        row_index = int(definition["Row"])
        seed = seed_rows.get(row_index, {})
        date_mode = get_date_column_write_mode(worksheet, row_index, 4)
        write_rows_to_worksheet(
            worksheet,
            [
                [
                    get_excel_text_literal(seed.get("Account", "")),
                    get_excel_text_literal(seed.get("Description", "")),
                    "N",
                    get_date_matrix_value(period_date_value, date_mode),
                    get_excel_text_literal(seed.get("Mod") or "COSSE"),
                    get_excel_text_literal(seed.get("Asiento", "")),
                    None,
                    get_excel_text_literal(seed.get("Detalle", "")),
                    get_numeric_matrix_value(definition["Amount"], blank_if_zero=True),
                    None,
                ]
            ],
            row_index,
            1,
        )


def write_costo_generated_rows(worksheet: Any, metrics: dict[str, float], period_date_value: float) -> None:
    row_numbers = [6, 7, 8, 9]
    seed_rows = read_costo_seed_rows(worksheet, row_numbers)
    clear_worksheet_range_contents(worksheet, "A", 6, "J", 141)
    rows = [
        {"Row": 6, "Detail": "REPUESTOS", "Amount": float(metrics["CostoRepuestos"])},
        {"Row": 7, "Detail": "INSUMOS Y LUBRICANTES", "Amount": float(metrics["CostoLubricantes"])},
        {"Row": 8, "Detail": "SUBCONTRATOS", "Amount": float(metrics["Costo"] + metrics["CostoSubconNc"])},
        {"Row": 9, "Detail": "ACCESORIOS", "Amount": float(metrics["CostoAccesorios"])},
    ]
    for definition in rows:
        row_index = int(definition["Row"])
        seed = seed_rows.get(row_index, {})
        date_mode = get_date_column_write_mode(worksheet, row_index, 3)
        write_rows_to_worksheet(
            worksheet,
            [
                [
                    get_excel_text_literal(seed.get("Account") or "05.01.01.01.0005"),
                    get_excel_text_literal(seed.get("Description") or "COSTO DE VENTAS-SERVICIO"),
                    get_date_matrix_value(period_date_value, date_mode),
                    get_excel_text_literal(seed.get("Mod") or "COSSE"),
                    get_excel_text_literal(seed.get("Asiento", "")),
                    None,
                    get_excel_text_literal(seed.get("Detalle") or definition["Detail"]),
                    get_numeric_matrix_value(definition["Amount"], blank_if_zero=True),
                    None,
                ]
            ],
            row_index,
            1,
        )


def clear_precont_costos_worksheet(worksheet: Any) -> None:
    clear_worksheet_range_contents(worksheet, "B", 2, "J", 922)
    clear_worksheet_range_contents(worksheet, "P", 1, "T", 25)


def new_sumif_account_formula(sheet_name: str, sum_column: str, account: str, sign: str = "+") -> str:
    normalized_sign = "-" if sign == "-" else "+"
    safe_sheet_name = sheet_name.replace("'", "''")
    safe_account = account.replace('"', '""')
    return f"""{normalized_sign}SUMIFS('{safe_sheet_name}'!${sum_column}:${sum_column},'{safe_sheet_name}'!$E:$E,"{safe_account}")"""


def first_generated_account(rows: list[dict[str, Any]], pattern: str) -> str:
    for row in rows:
        account = normalize_text(row.get("Account"))
        if re.fullmatch(pattern, account or ""):
            return account
    return ""


def update_precont_ventas_control_formulas(worksheet: Any, rows: list[dict[str, Any]]) -> None:
    credit_client_account = first_generated_account(rows, r"010104\d{2}0002")
    reserve_account = first_generated_account(rows, r"020105\d{2}000\d")
    liquidar_credit_account = first_generated_account(rows, r"020120\d{2}0002")
    liquidar_discount_account = first_generated_account(rows, r"020120\d{2}0004")
    worksheet.Range("V7").Formula = (
        "=" + new_sumif_account_formula("PrecontabilizacionVentas", "I", credit_client_account, "+") + new_sumif_account_formula("PrecontabilizacionVentas", "H", reserve_account, "-")
        if credit_client_account and reserve_account
        else "=0"
    )
    worksheet.Range("V10").Formula = (
        "=" + new_sumif_account_formula("PrecontabilizacionVentas", "I", liquidar_credit_account, "+") + new_sumif_account_formula("PrecontabilizacionVentas", "H", liquidar_discount_account, "-")
        if liquidar_credit_account and liquidar_discount_account
        else "=0"
    )


def classify_mayor_control_bucket(account: Any, name: Any) -> str:
    return classify_sales_control_bucket(account, name)


def get_mayor_control_metrics(rows: list[dict[str, Any]]) -> dict[str, float]:
    return get_control_metrics_from_rows(
        rows,
        account_key="account",
        name_key="name",
        debit_key="debit",
        credit_key="credit",
    )


def get_precont_ventas_control_metrics(rows: list[dict[str, Any]]) -> dict[str, float]:
    return get_control_metrics_from_rows(
        rows,
        account_key="Account",
        name_key="Description",
        debit_key="Debit",
        credit_key="Credit",
    )


def get_rep_vtas_control_cost_total(metrics: dict[str, float]) -> float:
    return float(
        round_amount(
            to_number(metrics.get("CostoRepuestos"))
            + to_number(metrics.get("CostoLubricantes"))
            + to_number(metrics.get("Costo"))
            + to_number(metrics.get("CostoSubconNc"))
            + to_number(metrics.get("CostoAccesorios"))
        )
    )


def get_rows_numeric_total(worksheet: Any, rows: list[dict[str, Any]], column: int) -> float:
    total = 0.0
    if worksheet is None:
        return total
    for item in rows:
        row_number = int(item.get("RowNumber", 0) or 0)
        if row_number <= 0:
            continue
        total += to_number(cell(worksheet, row_number, column).Value2)
    return float(round_amount(total))


def get_note_visible_sales_total(worksheet: Any, rows: list[dict[str, Any]]) -> float:
    total = 0.0
    if worksheet is None:
        return total
    for item in rows or []:
        row_number = int(item.get("RowNumber") or 0)
        if row_number <= 0:
            continue
        total += (
            to_number(cell(worksheet, row_number, 11).Value2)
            + to_number(cell(worksheet, row_number, 12).Value2)
            + to_number(cell(worksheet, row_number, 13).Value2)
        )
    return float(round_amount(total))


def get_px_visible_adjustment_totals(worksheet: Any) -> dict[str, float]:
    if worksheet is None:
        return {"Gross": 0.0, "Discount": 0.0}
    ranges = get_px_detail_ranges(worksheet)
    if len(ranges) < 2:
        return {"Gross": 0.0, "Discount": 0.0}
    bottom_range = ranges[1]
    gross_total = 0.0
    discount_total = 0.0
    for row in range(int(bottom_range["StartRow"]), int(bottom_range["EndRow"]) + 1):
        invoice_text = normalize_text(cell(worksheet, row, 5).Text)
        if invoice_text == "":
            continue
        gross_total += to_number(cell(worksheet, row, 12).Value2)
        discount_total += to_number(cell(worksheet, row, 14).Value2)
    return {
        "Gross": float(round_amount(gross_total)),
        "Discount": float(round_amount(discount_total)),
    }


def get_source_control_metrics_from_visible_sheets(
    rep_worksheet: Any,
    invoice_rows: list[dict[str, Any]],
    note_worksheet: Any,
    note_rows: list[dict[str, Any]],
    px_worksheet: Any,
) -> dict[str, float]:
    invoice_sales = get_rows_numeric_total(rep_worksheet, invoice_rows, 8)
    invoice_discounts = get_rows_numeric_total(rep_worksheet, invoice_rows, 9)
    note_sales = get_note_visible_sales_total(note_worksheet, note_rows)
    note_discounts = get_rows_numeric_total(note_worksheet, note_rows, 11)
    px_totals = get_px_visible_adjustment_totals(px_worksheet)
    invoice_sales = float(round_amount(invoice_sales - px_totals["Gross"]))
    invoice_discounts = float(round_amount(invoice_discounts - px_totals["Discount"]))
    return {
        "InvoiceSales": invoice_sales,
        "InvoiceDiscounts": invoice_discounts,
        "NoteSales": float(round_amount(note_sales)),
        "NoteDiscounts": float(round_amount(note_discounts)),
        "NetSales": float(round_amount(invoice_sales - invoice_discounts - note_sales + note_discounts)),
    }


def should_use_source_control_metrics(source_metrics: dict[str, float], mayor_metrics: dict[str, float]) -> bool:
    keys = ("InvoiceSales", "InvoiceDiscounts", "NoteSales", "NoteDiscounts")
    for key in keys:
        if abs(to_number(source_metrics.get(key)) - to_number(mayor_metrics.get(key))) > 1.0:
            return True
    return False


def log_control_metrics(tag: str, brand_key: str, metrics: dict[str, float]) -> None:
    print(
        "INFO|{tag}|{brand}|invoice_sales={invoice_sales}|invoice_discounts={invoice_discounts}|"
        "note_sales={note_sales}|note_discounts={note_discounts}|net_sales={net_sales}".format(
            tag=tag,
            brand=brand_key,
            invoice_sales=float(round_amount(to_number(metrics.get("InvoiceSales")))),
            invoice_discounts=float(round_amount(to_number(metrics.get("InvoiceDiscounts")))),
            note_sales=float(round_amount(to_number(metrics.get("NoteSales")))),
            note_discounts=float(round_amount(to_number(metrics.get("NoteDiscounts")))),
            net_sales=float(round_amount(to_number(metrics.get("NetSales")))),
        )
    )


def set_control_numeric_value(worksheet: Any, row: int, column: int, value: Any) -> None:
    if worksheet is None:
        return
    set_numeric_cell_safe(worksheet, row, column, to_number(value), blank_if_zero=True)


def update_sales_control_blocks(
    rep_worksheet: Any,
    note_worksheet: Any,
    rep_vtas_worksheet: Any,
    mayor_metrics: dict[str, float],
    precont_metrics: dict[str, float],
    cost_metrics: dict[str, float],
    preserve_rep_vtas_cost_formula: bool = False,
) -> None:
    if rep_worksheet is not None:
        set_control_numeric_value(rep_worksheet, 9, 4, precont_metrics["InvoiceSales"])
        set_control_numeric_value(rep_worksheet, 9, 5, precont_metrics["InvoiceDiscounts"])
        set_control_numeric_value(rep_worksheet, 9, 10, mayor_metrics["InvoiceSales"])
        set_control_numeric_value(rep_worksheet, 9, 11, mayor_metrics["InvoiceDiscounts"])

    if note_worksheet is not None:
        set_control_numeric_value(note_worksheet, 4, 4, mayor_metrics["NoteSales"])
        set_control_numeric_value(note_worksheet, 4, 5, mayor_metrics["NoteDiscounts"])
        set_control_numeric_value(note_worksheet, 4, 6, precont_metrics["NoteSales"])
        set_control_numeric_value(note_worksheet, 4, 7, precont_metrics["NoteDiscounts"])

    if rep_vtas_worksheet is not None:
        set_control_numeric_value(rep_vtas_worksheet, 6, 4, mayor_metrics["NetSales"])
        if not preserve_rep_vtas_cost_formula:
            set_control_numeric_value(rep_vtas_worksheet, 6, 5, get_rep_vtas_control_cost_total(cost_metrics))


def should_preserve_template_costo_sheet(worksheet: Any) -> bool:
    if worksheet is None:
        return False
    try:
        formula_h4 = normalize_text(cell(worksheet, 4, 8).Formula)
        formula_i4 = normalize_text(cell(worksheet, 4, 9).Formula)
        formula_j4 = normalize_text(cell(worksheet, 4, 10).Formula)
    except Exception:
        return False
    return "SUBTOTAL" in formula_h4.upper() and "SUBTOTAL" in formula_i4.upper() and formula_j4 != ""


def get_costo_control_total(worksheet: Any) -> float:
    if worksheet is None:
        return 0.0
    return float(round_amount(to_number(cell(worksheet, 4, 10).Value2)))


def write_px_rows_to_worksheet(worksheet: Any, rows: list[list[Any]], brand_key: str = "") -> dict[str, Any]:
    ranges = get_px_detail_ranges(worksheet)
    if len(ranges) < 2:
        raise RuntimeError("La hoja PX no contiene los bloques esperados para notas de credito y ventas por liquidar.")
    top_range, bottom_range = ranges[0], ranges[1]
    for px_range in (top_range, bottom_range):
        try:
            constants = worksheet.Range(f"B{px_range['StartRow']}:Q{px_range['EndRow']}").SpecialCells(XLCELLTYPE_CONSTANTS)
            constants.ClearContents()
        except pywintypes.com_error:
            pass
    detail_rows = convert_px_rows_to_detail_rows(rows)
    capacity = int(bottom_range["EndRow"]) - int(bottom_range["StartRow"]) + 1
    if len(detail_rows) > capacity:
        raise RuntimeError(f"Las filas PX de {brand_key} exceden la capacidad de la plantilla. Capacidad={capacity}, filas={len(detail_rows)}")
    if detail_rows:
        sheet_rows = [
            [
                get_excel_text_literal(row.get("Agencia")),
                get_excel_text_literal(row.get("Estado")),
                get_excel_text_literal(row.get("Orden")),
                get_excel_text_literal(row.get("Factura")),
                get_excel_text_literal(row.get("Fecha")),
                get_excel_text_literal(row.get("PxNo")),
                get_excel_text_literal(row.get("Cuenta")),
                get_excel_text_literal(row.get("Fr")),
                get_excel_text_literal(row.get("Codigo")),
                get_excel_text_literal(row.get("Item")),
                get_numeric_matrix_value(row.get("PvpBruto"), blank_if_zero=True),
                get_numeric_matrix_value(row.get("DescPct"), blank_if_zero=True),
                get_numeric_matrix_value(row.get("DescValor"), blank_if_zero=True),
                get_numeric_matrix_value(row.get("PvpNeto"), blank_if_zero=True),
                get_numeric_matrix_value(row.get("Costo"), blank_if_zero=True),
                get_excel_text_literal(row.get("Origen")),
            ]
            for row in detail_rows
        ]
        write_rows_to_worksheet(worksheet, sheet_rows, int(bottom_range["StartRow"]), 2)
    cell(worksheet, 3, 4).Value2 = get_excel_text_literal(datetime.now().strftime("%d/%m/%Y %H.%M.%S"))
    cell(worksheet, 5, 4).Value2 = get_excel_text_literal(get_brand_display_label(brand_key))
    return {
        "RowCount": len(detail_rows),
        "TopCapacity": int(top_range["EndRow"]) - int(top_range["StartRow"]) + 1,
        "BottomCapacity": capacity,
    }


def aggregate_mayor_rows(rows: list[dict[str, Any]], key_fields: tuple[str, ...]) -> list[dict[str, Any]]:
    aggregated: dict[tuple[Any, ...], dict[str, Any]] = {}
    ordered_keys: list[tuple[Any, ...]] = []

    for row in rows:
        key = tuple(
            get_date_write_value(row.get(field)) if field == "date_value" else normalize_text(row.get(field))
            for field in key_fields
        )
        current = aggregated.get(key)
        if current is None:
            current = dict(row)
            current["debit"] = 0.0
            current["credit"] = 0.0
            aggregated[key] = current
            ordered_keys.append(key)

        current["debit"] = float(round_amount(to_number(current.get("debit")) + to_number(row.get("debit"))))
        current["credit"] = float(round_amount(to_number(current.get("credit")) + to_number(row.get("credit"))))
        current["balance"] = row.get("balance")
        current["effective_balance"] = row.get("effective_balance", row.get("balance"))
        if row.get("date_value") not in (None, ""):
            current["date_value"] = row.get("date_value")
        if normalize_text(row.get("date_text")) != "":
            current["date_text"] = row.get("date_text")
        for field in ("account", "name", "ext", "origin", "seat", "reference", "detail"):
            if normalize_text(current.get(field)) == "" and normalize_text(row.get(field)) != "":
                current[field] = row.get(field)

    return [aggregated[key] for key in ordered_keys]


def compact_mayor_rows_for_capacity(rows: list[dict[str, Any]], capacity: int) -> list[dict[str, Any]]:
    compacted = list(rows)
    if len(compacted) <= capacity:
        return compacted

    grouping_levels = [
        ("account", "name", "ext", "date_value", "origin", "seat", "reference", "detail"),
        ("account", "name", "ext", "date_value", "origin", "seat", "detail"),
        ("account", "name", "ext", "date_value", "origin", "detail"),
        ("account", "name", "ext", "date_value", "origin"),
        ("account", "name", "ext"),
    ]

    for level in grouping_levels:
        next_rows = aggregate_mayor_rows(compacted, level)
        if len(next_rows) < len(compacted):
            compacted = next_rows
        if len(compacted) <= capacity:
            return compacted

    return compacted


def write_mayor_rows_to_worksheet(worksheet: Any, rows: list[dict[str, Any]], brand_key: str = "") -> dict[str, Any]:
    layouts = get_mayor_sheet_section_layouts(worksheet)
    if not layouts:
        raise RuntimeError(f"La hoja {worksheet.Name} no tiene secciones SUBTOTAL reconocibles para cargar el mayor.")
    rows_by_account: dict[str, list[dict[str, Any]]] = {}
    for row in rows:
        account = normalize_text(row.get("account"))
        if account == "":
            continue
        rows_by_account.setdefault(account, []).append(row)
    rows_by_layout_key: dict[str, list[dict[str, Any]]] = {}
    compatible_mapped_accounts: list[str] = []
    unmapped_accounts: list[str] = []
    allow_cross_prefix_family = mayor_brand_allows_flexible_account_mapping(brand_key)
    written_rows_for_metrics: list[dict[str, Any]] = []
    for account in rows_by_account.keys():
        layout = next((item for item in layouts if item["Account"] == account), None)
        if layout is None:
            account_rows = rows_by_account.get(account, [])
            sample_name = account_rows[0].get("name") if account_rows else ""
            layout = resolve_mayor_compatible_layout(
                account,
                sample_name,
                layouts,
                rows_by_layout_key,
                allow_cross_prefix_family=allow_cross_prefix_family,
            )
            if layout is None:
                print(f"WARN|mayor_account_unmapped|{brand_key}|{account}|sheet={worksheet.Name}")
                unmapped_accounts.append(account)
                continue
            compatible_mapped_accounts.append(account)
            print(f"WARN|mayor_account_compatible_section|{brand_key}|{account}|section={layout['Account']}|sheet={worksheet.Name}")
        layout_key = f"{layout['Account']}:{int(layout['StartRow'])}-{int(layout['EndRow'])}"
        rows_by_layout_key.setdefault(layout_key, []).extend(rows_by_account[account])

    written_count = 0
    for layout in layouts:
        clear_worksheet_range_contents(worksheet, "A", int(layout["StartRow"]), "M", int(layout["EndRow"]))
        layout_key = f"{layout['Account']}:{int(layout['StartRow'])}-{int(layout['EndRow'])}"
        account_rows = rows_by_layout_key.get(layout_key, [])
        capacity = int(layout["EndRow"]) - int(layout["StartRow"]) + 1
        if len(account_rows) > capacity:
            compacted_rows = compact_mayor_rows_for_capacity(account_rows, capacity)
            if len(compacted_rows) < len(account_rows):
                print(f"WARN|mayor_account_compacted|{brand_key}|{layout['Account']}|sheet={worksheet.Name}|before={len(account_rows)}|after={len(compacted_rows)}")
                account_rows = compacted_rows
        if len(account_rows) > capacity:
            raise RuntimeError(f"El mayor para {brand_key} excede la capacidad de la seccion {layout['Account']} en {worksheet.Name}. Capacidad={capacity}, filas={len(account_rows)}")
        if account_rows:
            date_mode = get_date_column_write_mode(worksheet, int(layout["StartRow"]), 4)
            sheet_rows = []
            for account_row in account_rows:
                balance_value = account_row.get("effective_balance", account_row.get("balance"))
                sheet_rows.append(
                    [
                        get_excel_text_literal(normalize_text(account_row.get("account"))),
                        get_excel_text_literal(normalize_text(account_row.get("name"))),
                        get_excel_text_literal(normalize_text(account_row.get("ext"))),
                        get_date_matrix_value(account_row.get("date_value"), date_mode),
                        get_excel_text_literal(normalize_text(account_row.get("origin"))),
                        get_excel_text_literal(normalize_text(account_row.get("seat"))),
                        get_excel_text_literal(normalize_text(account_row.get("reference"))),
                        get_excel_text_literal(normalize_text(account_row.get("detail"))),
                        get_numeric_matrix_value(account_row.get("debit")),
                        get_numeric_matrix_value(account_row.get("credit")),
                        get_numeric_matrix_value(balance_value),
                    ]
                )
                written_rows_for_metrics.append(account_row)
            write_rows_to_worksheet(worksheet, sheet_rows, int(layout["StartRow"]), 1)
            written_count += len(account_rows)
    return {
        "RowCount": written_count,
        "SectionCount": len(layouts),
        "UnmappedAccounts": unmapped_accounts,
        "CompatibleMappedAccounts": compatible_mapped_accounts,
        "ControlMetrics": get_mayor_control_metrics(written_rows_for_metrics),
    }


def split_mayor_unmapped_accounts(accounts: list[str]) -> dict[str, list[str]]:
    warning: list[str] = []
    fatal: list[str] = []
    for account in sorted({normalize_text(item) for item in accounts if normalize_text(item) != ""}):
        if re.fullmatch(r"04\.01\.01\.\d{2}\.0002", account):
            warning.append(account)
            continue
        fatal.append(account)
    return {"Warning": warning, "Fatal": fatal}


def describe_mayor_accounts(rows: list[dict[str, Any]], accounts: list[str]) -> list[str]:
    names_by_account: dict[str, str] = {}
    for row in rows:
        account = normalize_text(row.get("account"))
        if account == "" or account in names_by_account:
            continue
        names_by_account[account] = normalize_text(row.get("name"))

    descriptions: list[str] = []
    for account in accounts:
        normalized_account = normalize_text(account)
        if normalized_account == "":
            continue
        account_name = names_by_account.get(normalized_account, "")
        descriptions.append(f"{normalized_account} ({account_name})" if account_name else normalized_account)
    return descriptions


def assert_mayor_matches_template(
    brand_key: str,
    mayor_path: Path | None,
    mayor_rows: list[dict[str, Any]],
    mayor_result: dict[str, Any] | None,
    worksheet_name: str,
) -> None:
    if mayor_path is None or not mayor_path.is_file():
        return

    brand_label = get_brand_display_label(brand_key)
    if not mayor_rows:
        raise RuntimeError(
            f"El archivo MAYOR de {brand_label} no contiene movimientos validos para poblar {worksheet_name}. "
            "Verifica que corresponda al mayor de ventas del periodo."
        )

    if mayor_result is None:
        raise RuntimeError(f"No se pudo validar el archivo MAYOR de {brand_label} para {worksheet_name}.")

    unmapped_split = split_mayor_unmapped_accounts(list(mayor_result.get("UnmappedAccounts") or []))
    if unmapped_split["Warning"]:
        accounts_text = ", ".join(unmapped_split["Warning"])
        print(f"WARN|mayor_unmapped_compatible|{brand_key}|accounts={accounts_text}")
    if int(mayor_result.get("RowCount") or 0) <= 0:
        ignored_accounts = unmapped_split["Fatal"] or list(mayor_result.get("UnmappedAccounts") or [])
        accounts_text = ", ".join(describe_mayor_accounts(mayor_rows, ignored_accounts))
        raise RuntimeError(
            f"El archivo MAYOR de {brand_label} no genero filas utiles en {worksheet_name}. "
            f"Cuentas detectadas sin seccion util: {accounts_text}. "
            "Debes subir el MAYOR VENTAS con cuentas 04.01.01.xx.xxxx."
        )
    if unmapped_split["Fatal"]:
        accounts_text = ", ".join(describe_mayor_accounts(mayor_rows, unmapped_split["Fatal"]))
        print(f"WARN|mayor_accounts_ignored|{brand_key}|accounts={accounts_text}")


def fill_rep_vtas(worksheet: Any, rows: list[dict[str, Any]], lookups: dict[str, Any]) -> dict[str, Any]:
    clear_output_sheet(worksheet, 15, "AL")
    fact_date_mode = get_date_column_write_mode(worksheet, 15, 9)
    note_date_mode = get_date_column_write_mode(worksheet, 15, 10)
    sorted_rows = sorted(
        rows,
        key=lambda row: (
            int(find_rep_vtas_entry(lookups["RepVtas"], row.get("DocumentTrim"), row.get("Order")).get("RowOrder"))
            if find_rep_vtas_entry(lookups["RepVtas"], row.get("DocumentTrim"), row.get("Order")) is not None
            else 100000 + int(row.get("RowIndex", 0))
        ),
    )
    sheet_rows: list[list[Any]] = []
    written_rows: list[dict[str, Any]] = []
    target_row = 15
    for row in sorted_rows:
        rep_lookup = find_rep_vtas_entry(lookups["RepVtas"], row.get("DocumentTrim"), row.get("Order"))
        gar_ext_default = get_lookup_default_text(lookups, "RepVtas", "GarExt")
        agency_value = get_preferred_source_text(row.get("AgencyRaw"), row.get("Agency"), rep_lookup.get("Agency", "") if rep_lookup else "")
        center_value = get_preferred_source_text(row.get("CenterRaw"), row.get("Center"), rep_lookup.get("Center", "") if rep_lookup else "")
        order_value = get_preferred_source_text(row.get("OrderRaw"), row.get("Order"), rep_lookup.get("Order", "") if rep_lookup else "")
        advisor_value = get_preferred_source_text(row.get("AdvisorRaw"), row.get("Advisor"), rep_lookup.get("Advisor", "") if rep_lookup else "")
        line_value = get_preferred_source_text(row.get("LineRaw"), row.get("Line"), rep_lookup.get("Line", "") if rep_lookup else "")
        cedula_value = get_preferred_source_text(row.get("CedulaRaw"), row.get("Cedula"), rep_lookup.get("Cedula", "") if rep_lookup else "")
        customer_value = get_preferred_source_text(row.get("CustomerRaw"), row.get("Customer"), rep_lookup.get("Customer", "") if rep_lookup else "")
        document_raw_value = get_preferred_source_text(row.get("DocumentRaw"), row.get("DocumentTrim"), rep_lookup.get("DocumentRaw", "") if rep_lookup else "")
        fact_date_value = get_preferred_source_date(row.get("DateFactValue"), rep_lookup.get("DateFactValue") if rep_lookup else None)
        note_date_value = get_preferred_source_date(row.get("DateNoteValue"), rep_lookup.get("DateNoteValue") if rep_lookup else None)
        if normalize_text(row.get("DocType")).upper() in {"DC", "DE"} and note_date_value in (None, "") and fact_date_value not in (None, ""):
            note_date_value = fact_date_value
        gar_ext_value = resolve_template_garext(rep_lookup.get("GarExt", "") if rep_lookup else "", row.get("GarExtRaw"), row.get("GarExt"), gar_ext_default)
        values = [
            get_excel_text_literal(agency_value),
            "'" + normalize_text(center_value),
            get_excel_text_literal(order_value),
            get_excel_text_literal(advisor_value),
            get_excel_text_literal(line_value),
            "'" + normalize_text(cedula_value),
            get_excel_text_literal(customer_value),
            "'" + normalize_text(document_raw_value),
            get_date_matrix_value(fact_date_value, fact_date_mode),
            get_date_matrix_value(note_date_value, note_date_mode),
            get_numeric_matrix_value(row.get("NoteCredit")),
            get_numeric_matrix_value(row.get("TotalManoObra")),
            get_numeric_matrix_value(row.get("TotalSubcontratos")),
            get_numeric_matrix_value(row.get("TotalInsumos")),
            get_numeric_matrix_value(row.get("TotalServicio")),
            get_numeric_matrix_value(row.get("TotalAccesorios")),
            get_numeric_matrix_value(row.get("TotalRepuestos")),
            get_numeric_matrix_value(row.get("Interes"), blank_if_zero=True),
            get_numeric_matrix_value(row.get("Iva")),
            get_numeric_matrix_value(row.get("Total")),
            get_numeric_matrix_value(row.get("Costo")),
            get_numeric_matrix_value(row.get("CostoLubricantes")),
            get_numeric_matrix_value(row.get("CostoAccesorios")),
            get_numeric_matrix_value(row.get("CostoRepuestos")),
            get_numeric_matrix_value(row.get("CostoPintura")),
            get_numeric_matrix_value(row.get("CostoSubconNc")),
            get_excel_text_literal(gar_ext_value),
        ]
        sheet_rows.append(values)
        written_rows.append({"RowNumber": target_row})
        target_row += 1
    if sheet_rows:
        write_rows_to_worksheet(worksheet, sheet_rows, 15, 1)
    return {"RowCount": len(written_rows), "Rows": written_rows}


def fill_invoices(worksheet: Any, rows: list[dict[str, Any]], lookups: dict[str, Any]) -> dict[str, Any]:
    clear_output_sheet(worksheet, 17, "S")
    invoice_date_mode = get_date_column_write_mode(worksheet, 17, 4)
    sorted_rows = sorted(
        [row for row in rows if normalize_text(row.get("DocType")).upper() in {"FA", "FC"}],
        key=lambda row: (normalize_text(row.get("Agency")), get_document_sort_value(row.get("DocumentTrim")), normalize_text(row.get("Order"))),
    )
    fallback_count = 0
    sheet_rows: list[list[Any]] = []
    written_rows: list[dict[str, Any]] = []
    target_row = 17
    for row in sorted_rows:
        lookup = find_lookup_entry(lookups["Invoice"], row.get("DocumentTrim"), row.get("Order"))
        if lookup is None:
            fallback_count += 1
        source_amounts = get_invoice_source_amounts(row)
        gar_ext_default = get_lookup_default_text(lookups, "Invoice", "GarExt")
        agency_value = get_preferred_lookup_text(lookup.get("Agency", "") if lookup else "", row.get("AgencyRaw"), row.get("Agency"))
        series_value = get_preferred_lookup_text(lookup.get("Series", "") if lookup else "", row.get("SeriesRaw"), row.get("Series"))
        order_value = get_preferred_lookup_text(lookup.get("Order", "") if lookup else "", row.get("OrderRaw"), row.get("Order"))
        cedula_value = get_preferred_lookup_text(lookup.get("Cedula", "") if lookup else "", row.get("CedulaRaw"), row.get("Cedula"))
        customer_value = get_preferred_lookup_text(lookup.get("Customer", "") if lookup else "", row.get("CustomerRaw"), row.get("Customer"))
        asiento_value = get_preferred_source_text(get_invoice_asiento(row), "", lookup.get("Asiento", "") if lookup else "")
        gar_ext_value = resolve_template_garext(lookup.get("GarExt", "") if lookup else "", row.get("GarExtRaw"), row.get("GarExt"), gar_ext_default)
        tv_value = get_preferred_lookup_text(lookup.get("Tv", "") if lookup else "", row.get("FormaPago"), row.get("FormaPago"))
        values = [
            get_excel_text_literal(agency_value),
            get_excel_text_literal(series_value),
            row.get("DocumentTrim"),
            get_date_matrix_value(get_date_write_value(row.get("DateFactValue")), invoice_date_mode),
            get_excel_text_literal(order_value),
            "'" + normalize_text(cedula_value),
            get_excel_text_literal(customer_value),
            get_numeric_matrix_value(source_amounts["Subtotal"]),
            get_numeric_matrix_value(source_amounts["Discount"]),
            get_numeric_matrix_value(source_amounts["NetoConIva"]),
            get_numeric_matrix_value(source_amounts["NetoIva0"]),
            get_numeric_matrix_value(source_amounts["Iva12"]),
            get_numeric_matrix_value(source_amounts["Iva15"]),
            get_numeric_matrix_value(source_amounts["Interest"]),
            get_numeric_matrix_value(source_amounts["Total"]),
            get_excel_text_literal(asiento_value),
            get_excel_text_literal(gar_ext_value),
            get_excel_text_literal(tv_value),
            get_excel_text_literal("N"),
        ]
        sheet_rows.append(values)
        written_rows.append({"RowNumber": target_row})
        target_row += 1
    if sheet_rows:
        write_rows_to_worksheet(worksheet, sheet_rows, 17, 1)
    return {"FallbackCount": fallback_count, "RowCount": len(written_rows), "Rows": written_rows}


def fill_notes(worksheet: Any, rows: list[dict[str, Any]], lookups: dict[str, Any]) -> dict[str, Any]:
    clear_output_sheet(worksheet, 11, "U")
    credit_note_date_mode = get_date_column_write_mode(worksheet, 11, 3)
    sorted_rows = sorted(
        [row for row in rows if normalize_text(row.get("DocType")).upper() in {"DC", "DE"}],
        key=lambda row: (normalize_text(row.get("Agency")), get_document_sort_value(row.get("DocumentTrim")), normalize_text(row.get("Order"))),
    )
    fallback_count = 0
    sheet_rows: list[list[Any]] = []
    written_rows: list[dict[str, Any]] = []
    target_row = 11
    for row in sorted_rows:
        order_key = strip_order_suffix(row.get("Order"))
        lookup = find_lookup_entry(lookups["Note"], row.get("DocumentTrim"), order_key)
        if lookup is None:
            fallback_count += 1
        source_amounts = get_note_source_amounts(row)
        gar_ext_default = get_lookup_default_text(lookups, "Note", "GarExt")
        agency_value = get_preferred_lookup_text(lookup.get("Agency", "") if lookup else "", row.get("AgencyRaw"), row.get("Agency"))
        kind_value = "CON" if normalize_text(row.get("DocType")).upper() == "DE" else get_preferred_source_text("CRE", "", lookup.get("Kind", "") if lookup else "")
        series_value = get_preferred_lookup_text(lookup.get("Series", "") if lookup else "", row.get("SeriesRaw"), row.get("Series"))
        invoice_value = get_preferred_source_text(trim_document(row.get("AffectedDocumentRaw")), trim_document(row.get("AffectedDocumentTrim")), lookup.get("Invoice", "") if lookup else "")
        order_value = get_preferred_source_text(row.get("OrderRaw"), order_key or normalize_text(row.get("Order")), lookup.get("Order", "") if lookup else "")
        cedula_value = get_preferred_lookup_text(lookup.get("Cedula", "") if lookup else "", row.get("CedulaRaw"), row.get("Cedula"))
        customer_value = get_preferred_lookup_text(lookup.get("Customer", "") if lookup else "", row.get("CustomerRaw"), row.get("Customer"))
        gar_ext_value = resolve_template_garext(lookup.get("GarExt", "") if lookup else "", row.get("GarExtRaw"), row.get("GarExt"), gar_ext_default)
        values = [
            get_excel_text_literal(agency_value),
            row.get("DocumentTrim"),
            get_date_matrix_value(get_date_write_value(row.get("DateNoteValue")), credit_note_date_mode),
            get_excel_text_literal(kind_value),
            get_excel_text_literal(series_value),
            get_excel_text_literal(invoice_value),
            get_excel_text_literal(order_value),
            "'" + normalize_text(cedula_value),
            get_excel_text_literal(customer_value),
            get_numeric_matrix_value(source_amounts["Subtotal"]),
            get_numeric_matrix_value(source_amounts["Discount"]),
            get_numeric_matrix_value(source_amounts["NetoSinIva"]),
            get_numeric_matrix_value(source_amounts["NetoConIva"]),
            get_numeric_matrix_value(source_amounts["Iva15"]),
            get_numeric_matrix_value(source_amounts["Iva12"]),
            get_numeric_matrix_value(source_amounts["Interest"]),
            get_numeric_matrix_value(source_amounts["Total"]),
            get_numeric_matrix_value(source_amounts["Anticipo"]),
            get_numeric_matrix_value(source_amounts["Neto"]),
            get_excel_text_literal(""),
            get_excel_text_literal(gar_ext_value),
        ]
        sheet_rows.append(values)
        written_rows.append({"RowNumber": target_row})
        target_row += 1
    if sheet_rows:
        write_rows_to_worksheet(worksheet, sheet_rows, 11, 1)
    return {"FallbackCount": fallback_count, "RowCount": len(written_rows), "Rows": written_rows}


def refresh_worksheet_pivot_tables_safe(worksheet: Any, label: str) -> None:
    if not PIVOT_REFRESH_ENABLED:
        print(f"INFO|pivot_refresh_skipped|{label}|mode=fast")
        return
    try:
        pivot_tables = worksheet.PivotTables()
        count = int(pivot_tables.Count)
        for index in range(1, count + 1):
            pivot_table = pivot_tables.Item(index)
            pivot_table.RefreshTable()
    except pywintypes.com_error as exc:
        raise RuntimeError(f"No se pudo refrescar la tabla dinamica de {label}: {exc}") from exc


def update_px_period_anchor(worksheet: Any, period_date_value: Any) -> None:
    if worksheet is None or period_date_value in (None, ""):
        return
    period_date = datetime(1899, 12, 30) + timedelta(days=float(period_date_value))
    period_start = datetime(period_date.year, period_date.month, 1)
    set_date_cell_value(worksheet, 3, 8, float((period_start - datetime(1899, 12, 30)).days), "PX_PERIODO_INICIO")


def open_workbook_with_retry(excel: Any, path: Path, read_only: bool = True, attempts: int = 20) -> Any:
    last_error = ""
    for attempt in range(1, attempts + 1):
        try:
            return excel.Workbooks.Open(str(path), 0, read_only)
        except pywintypes.com_error as exc:
            last_error = str(exc)
            if attempt < attempts:
                time.sleep(0.9 * attempt if any(token in last_error for token in ("RPC_E_CALL_REJECTED", "0x80010001", "0x800AC472")) else 0.7 * attempt)
    raise RuntimeError(f"No se pudo abrir el libro de Excel en {path}. Detalle: {last_error}")


def save_workbook_with_retry(workbook: Any, path_for_error: Path, attempts: int = 6) -> None:
    last_error = ""
    for attempt in range(1, attempts + 1):
        try:
            workbook.Save()
            return
        except pywintypes.com_error as exc:
            last_error = str(exc)
            if attempt < attempts:
                time.sleep(0.6 * attempt)
    raise RuntimeError(f"No se pudo guardar el libro de Excel en {path_for_error}. Detalle: {last_error}")


def parse_args(argv: list[str]) -> RuntimeArgs:
    parser = argparse.ArgumentParser(description="Runtime Python final para Servicios por Marca")
    parser.add_argument("--input", required=True)
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--template-dir", required=True)
    parser.add_argument("--run-stamp", default="")
    parser.add_argument("--cancel-path", default="")
    parser.add_argument("--brand-key", default="")
    parser.add_argument("--px-path", default="")
    parser.add_argument("--repvtas-path", default="")
    for brand_key, mapping in BRAND_INPUT_KEYS.items():
        parser.add_argument(f"--{mapping['factura'].replace('_', '-')}", default="")
        parser.add_argument(f"--{mapping['nota'].replace('_', '-')}", default="")
        parser.add_argument(f"--{mapping['mayor'].replace('_', '-')}", default="")
    args = parser.parse_args(argv)

    factura_paths = {brand_key: (Path(getattr(args, mapping["factura"])).resolve() if getattr(args, mapping["factura"]) else None) for brand_key, mapping in BRAND_INPUT_KEYS.items()}
    nota_paths = {brand_key: (Path(getattr(args, mapping["nota"])).resolve() if getattr(args, mapping["nota"]) else None) for brand_key, mapping in BRAND_INPUT_KEYS.items()}
    mayor_paths = {brand_key: (Path(getattr(args, mapping["mayor"])).resolve() if getattr(args, mapping["mayor"]) else None) for brand_key, mapping in BRAND_INPUT_KEYS.items()}
    return RuntimeArgs(
        input_path=Path(args.input).resolve(),
        output_dir=Path(args.output_dir).resolve(),
        template_dir=Path(args.template_dir).resolve(),
        run_stamp=str(args.run_stamp or "").strip(),
        cancel_path=Path(args.cancel_path).resolve() if str(args.cancel_path or "").strip() else None,
        brand_key=str(args.brand_key or "").strip(),
        px_path=Path(args.px_path).resolve() if str(args.px_path or "").strip() else None,
        rep_vtas_path=Path(args.repvtas_path).resolve() if str(args.repvtas_path or "").strip() else None,
        factura_paths=factura_paths,
        nota_paths=nota_paths,
        mayor_paths=mayor_paths,
    )


def run_runtime(args: RuntimeArgs) -> int:
    if not args.input_path.is_file():
        raise RuntimeError("No se encontro el archivo Excel subido.")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    if not args.template_dir.is_dir():
        raise RuntimeError("No se encontro la carpeta base de plantillas mensuales.")

    brand_file_map = {
        brand_key: {"FacturaPath": args.factura_paths.get(brand_key), "NotaPath": args.nota_paths.get(brand_key)}
        for brand_key in BRAND_INPUT_KEYS.keys()
    }
    display_requested = any((path is not None and path.is_file()) for path in list(args.factura_paths.values()) + list(args.nota_paths.values()))
    display_source_rows = build_source_rows_from_brand_inputs(brand_file_map, args.brand_key) if display_requested else []
    if display_source_rows:
        print(f"INFO|custom_source|rows={len(display_source_rows)}")
    source_sheet_name, source_rows = read_source_rows(args.rep_vtas_path or args.input_path)
    if source_rows and display_source_rows:
        print(f"INFO|source_priority|excel_primary|custom_support={len(display_source_rows)}")
    if not source_rows and display_source_rows:
        source_rows = display_source_rows
    if not source_rows:
        raise RuntimeError("Los archivos de entrada no contienen filas validas para procesar.")

    display_rows = normalize_source_rows(display_source_rows or source_rows, consolidate_invoice_documents=True)
    rep_vtas_rows = normalize_source_rows(source_rows)
    posting_rows = normalize_source_rows(source_rows, consolidate_invoice_documents=True)

    timestamp = args.run_stamp or datetime.now().strftime("%Y%m%d_%H%M%S")
    brand_order = ["changan", "peug", "szk", "tyt"]
    if args.brand_key:
        brand_order = [item for item in brand_order if item == args.brand_key]

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        try:
            excel.Calculation = XL_CALCULATION_MANUAL
        except Exception:
            pass

        for template_key in brand_order:
            assert_not_cancelled(args.cancel_path, "marcas")
            rows_display = [row for row in display_rows if row.get("TemplateKey") == template_key]
            rows_rep_vtas = [row for row in rep_vtas_rows if row.get("TemplateKey") == template_key]
            if not rows_rep_vtas and not rows_display:
                continue
            rows_posting = [row for row in posting_rows if row.get("TemplateKey") == template_key]
            print(f"INFO|processing|{template_key}|rows={len(rows_rep_vtas)}")

            template_path = args.template_dir / TEMPLATE_CONFIGS[template_key]["file"]
            output_name = f"servicios_{template_key}_{timestamp}.xls"
            output_path = args.output_dir / output_name
            shutil.copyfile(template_path, output_path)

            brand_start = time.perf_counter()
            output_workbook = open_workbook_with_retry(excel, output_path, False)
            brand_succeeded = False
            try:
                lookups = read_template_lookups_from_xls(template_path)
                precont_ventas_prototypes = read_precont_ventas_prototypes_from_xls(template_path)
                precont_costos2_prototypes = read_precont_costos2_prototypes_from_xls(template_path)
                rep_sheet = get_worksheet_safe(output_workbook, ["REP FACTURACION", "REP FACTURACION"])
                note_sheet = get_worksheet_safe(output_workbook, ["NOTA DE CREDITO"])
                rep_vtas_sheet = get_worksheet_safe(output_workbook, ["REP VTAS"])

                fill_start = time.perf_counter()
                invoice_result = fill_invoices(rep_sheet, rows_display, lookups)
                print(f"INFO|fill_invoices_ms|{template_key}|{int((time.perf_counter() - fill_start) * 1000)}")

                fill_start = time.perf_counter()
                note_result = fill_notes(note_sheet, rows_display, lookups)
                print(f"INFO|fill_notes_ms|{template_key}|{int((time.perf_counter() - fill_start) * 1000)}")

                fill_start = time.perf_counter()
                rep_vtas_result = fill_rep_vtas(rep_vtas_sheet, rows_rep_vtas, lookups)
                print(f"INFO|fill_repvtas_ms|{template_key}|{int((time.perf_counter() - fill_start) * 1000)}")

                px_sheet = None
                px_rows: list[list[Any]] = []
                try:
                    px_sheet = get_worksheet_safe(output_workbook, ["PX"])
                except RuntimeError:
                    px_sheet = None
                if px_sheet is not None:
                    fill_start = time.perf_counter()
                    _, px_rows = read_px_rows(args.px_path, template_key) if args.px_path and args.px_path.is_file() else ("", [])
                    px_result = write_px_rows_to_worksheet(px_sheet, px_rows, template_key)
                    print(f"INFO|fill_px_ms|{template_key}|{int((time.perf_counter() - fill_start) * 1000)}")
                    print(f"INFO|px|{template_key}|rows={px_result['RowCount']}|bottom_capacity={px_result['BottomCapacity']}")

                mayor_rows: list[dict[str, Any]] = []
                mayor_sheet = None
                mayor_layouts: list[dict[str, Any]] = []
                try:
                    mayor_sheet = get_worksheet_safe(output_workbook, ["VENTAS", "MAY VTAS"])
                    mayor_layouts = get_mayor_sheet_section_layouts(mayor_sheet)
                except RuntimeError:
                    mayor_sheet = None
                    mayor_layouts = []
                if mayor_sheet is not None:
                    fill_start = time.perf_counter()
                    mayor_path = args.mayor_paths.get(template_key)
                    if mayor_path is not None and mayor_path.is_file():
                        mayor_rows = read_mayor_rows(mayor_path)
                        mayor_filter = filter_mayor_rows_for_workbook(mayor_rows)
                        mayor_rows = mayor_filter["Rows"]
                        if mayor_filter["Removed"]:
                            print(f"INFO|mayor_px_adjustments_filtered|{template_key}|rows={len(mayor_filter['Removed'])}")
                    mayor_result = write_mayor_rows_to_worksheet(mayor_sheet, mayor_rows, template_key)
                    assert_mayor_matches_template(template_key, mayor_path, mayor_rows, mayor_result, str(mayor_sheet.Name))
                    print(f"INFO|fill_mayor_ms|{template_key}|{int((time.perf_counter() - fill_start) * 1000)}")
                    print(f"INFO|mayor|{template_key}|rows={mayor_result['RowCount']}|sections={mayor_result['SectionCount']}")
                mayor_control_metrics = dict((mayor_result or {}).get("ControlMetrics") or get_mayor_control_metrics(mayor_rows))

                cost_metrics = get_brand_cost_metrics(rows_rep_vtas)
                period_date_value = get_brand_period_date_value(rows_rep_vtas, mayor_rows)
                if mayor_sheet is not None:
                    supplemental_mayor_rows = build_supplemental_mayor_rows_from_source(
                        mayor_layouts,
                        mayor_rows,
                        rows_display,
                        rows_posting,
                        px_rows,
                        period_date_value,
                    )
                    if supplemental_mayor_rows:
                        mayor_rows = list(mayor_rows) + supplemental_mayor_rows
                        mayor_result = write_mayor_rows_to_worksheet(mayor_sheet, mayor_rows, template_key)
                        mayor_control_metrics = dict((mayor_result or {}).get("ControlMetrics") or get_mayor_control_metrics(mayor_rows))
                        print(f"INFO|mayor_source_supplement|{template_key}|rows={len(supplemental_mayor_rows)}")
                precont_ventas_sheet = get_worksheet_safe(output_workbook, ["PrecontabilizacionVentas"])
                precont_costos2_sheet = get_worksheet_safe(output_workbook, ["PrecontabilizacionCostos (2)"])
                precont_costos_sheet = get_worksheet_safe(output_workbook, ["PrecontabilizacionCostos"])
                costo_sheet = get_worksheet_safe(output_workbook, ["COSTO"])
                estadisticas_sheet = get_worksheet_safe(output_workbook, ["ESTADISTICAS"])
                source_control_metrics = get_source_control_metrics_from_visible_sheets(
                    rep_sheet,
                    invoice_result.get("Rows", []),
                    note_sheet,
                    note_result.get("Rows", []),
                    px_sheet,
                )
                precont_ventas_control_metrics = {
                    "InvoiceSales": 0.0,
                    "InvoiceDiscounts": 0.0,
                    "NoteSales": 0.0,
                    "NoteDiscounts": 0.0,
                    "NetSales": 0.0,
                }

                fill_start = time.perf_counter()
                precont_ventas_rows = new_precont_ventas_generated_rows(precont_ventas_prototypes, rows_posting, mayor_rows, px_rows, template_key)
                precont_ventas_control_metrics = get_precont_ventas_control_metrics(precont_ventas_rows)
                precont_ventas_count = write_precont_ventas_generated_rows(precont_ventas_sheet, precont_ventas_rows)
                refresh_worksheet_pivot_tables_safe(precont_ventas_sheet, "PrecontabilizacionVentas")
                update_precont_ventas_control_formulas(precont_ventas_sheet, precont_ventas_rows)
                print(f"INFO|fill_precont_ventas_ms|{template_key}|{int((time.perf_counter() - fill_start) * 1000)}")
                print(f"INFO|precont_ventas|{template_key}|rows={precont_ventas_count}")

                fill_start = time.perf_counter()
                precont_costos2_count = write_precont_costos2_generated_rows(precont_costos2_sheet, precont_costos2_prototypes, cost_metrics)
                refresh_worksheet_pivot_tables_safe(precont_costos2_sheet, "PrecontabilizacionCostos (2)")
                print(f"INFO|fill_precont_costos2_ms|{template_key}|{int((time.perf_counter() - fill_start) * 1000)}")
                print(f"INFO|precont_costos2|{template_key}|rows={precont_costos2_count}")

                clear_precont_costos_worksheet(precont_costos_sheet)
                print(f"INFO|precont_costos_legacy_neutralized|{template_key}")
                write_estadisticas_generated_rows(estadisticas_sheet, cost_metrics, period_date_value)
                preserve_rep_vtas_cost_formula = should_preserve_template_costo_sheet(costo_sheet)
                if preserve_rep_vtas_cost_formula:
                    print(f"INFO|costo_mode|{template_key}|template_preserved")
                else:
                    write_costo_generated_rows(costo_sheet, cost_metrics, period_date_value)
                    refresh_worksheet_pivot_tables_safe(costo_sheet, "COSTO")
                    print(f"INFO|costo_mode|{template_key}|generated")
                update_px_period_anchor(px_sheet, period_date_value)
                effective_control_metrics = source_control_metrics if should_use_source_control_metrics(source_control_metrics, mayor_control_metrics) else mayor_control_metrics
                mode_label = "source_fallback" if effective_control_metrics is source_control_metrics else "mayor_aligned"
                log_control_metrics("mayor_control_metrics", template_key, mayor_control_metrics)
                log_control_metrics("source_control_metrics", template_key, source_control_metrics)
                log_control_metrics("precont_control_metrics", template_key, precont_ventas_control_metrics)
                print(f"INFO|sales_control_mode|{template_key}|mode={mode_label}")
                rep_vtas_cost_total = get_costo_control_total(costo_sheet) if preserve_rep_vtas_cost_formula else get_rep_vtas_control_cost_total(cost_metrics)
                print(f"INFO|rep_vtas_cost_total|{template_key}|value={float(round_amount(rep_vtas_cost_total))}")
                update_sales_control_blocks(
                    rep_sheet,
                    note_sheet,
                    rep_vtas_sheet,
                    effective_control_metrics,
                    effective_control_metrics,
                    cost_metrics,
                    preserve_rep_vtas_cost_formula=preserve_rep_vtas_cost_formula,
                )

                try:
                    output_workbook.Application.CalculateFull()
                except Exception:
                    try:
                        output_workbook.Calculate()
                    except Exception as exc:
                        print(f"WARN|recalc_full_failed|{exc}")

                save_workbook_with_retry(output_workbook, output_path)
                brand_succeeded = True
                print(f"OUTPUT|{output_name}|{TEMPLATE_CONFIGS[template_key]['label']}")
                print(f"INFO|{template_key}|invoice_fallbacks={int(invoice_result['FallbackCount'])}|note_fallbacks={int(note_result['FallbackCount'])}")
                print(f"INFO|total_brand_ms|{template_key}|{int((time.perf_counter() - brand_start) * 1000)}")
            finally:
                try:
                    output_workbook.Close(bool(brand_succeeded))
                finally:
                    if not brand_succeeded and output_path.exists():
                        try:
                            output_path.unlink()
                        except OSError:
                            pass
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        gc.collect()
        pythoncom.CoUninitialize()
    return 0


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    try:
        return run_runtime(args)
    except CancelRequestedError as exc:
        print(f"CANCELLED|{exc}")
        return 130
    except Exception as exc:  # noqa: BLE001
        print(str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
