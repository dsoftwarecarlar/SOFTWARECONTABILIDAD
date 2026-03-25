from __future__ import annotations

import subprocess
from pathlib import Path

from ..contracts import ProcessRequest, ProcessResult
from ..legacy_node import extract_lines, resolve_node_binary


def run(request: ProcessRequest) -> ProcessResult:
    options = request.options
    script_path = Path(str(options.get("script_path", ""))).resolve()
    if not script_path.is_file():
        raise FileNotFoundError(f"Legacy script not found: {script_path}")

    template_path = request.template_path.resolve() if request.template_path else None
    if template_path is None or not template_path.is_file():
        raise FileNotFoundError(f"Template not found: {template_path}")

    file_fields = [item for item in options.get("file_fields", []) if isinstance(item, dict)]
    saved_inputs = options.get("saved_inputs", {})
    if not isinstance(saved_inputs, dict):
        raise ValueError("saved_inputs must be a mapping.")

    command = [resolve_node_binary(request), str(script_path)]
    for field_config in file_fields:
        field = str(field_config.get("field", "")).strip()
        script_flag = normalize_flag(str(field_config.get("script_flag", "")).strip())
        if not field or not script_flag:
            continue

        input_payload = saved_inputs.get(field)
        if isinstance(input_payload, dict):
            input_path = Path(str(input_payload.get("path", ""))).resolve()
        else:
            input_path = Path(str(input_payload or "")).resolve()

        if not input_path.is_file():
            raise FileNotFoundError(f"Input not found for {field}: {input_path}")

        command.extend([script_flag, str(input_path)])

    command.extend(["--template-path", str(template_path), "--output-path", str(request.output_path.resolve())])

    cwd = Path(str(options.get("cwd", script_path.parent))).resolve()
    timeout = int(options.get("timeout_seconds", 600))
    completed = subprocess.run(
        command,
        cwd=str(cwd),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        check=False,
    )
    if completed.returncode != 0:
        combined = "\n".join(
            part.strip() for part in (completed.stdout, completed.stderr) if part.strip()
        ).strip()
        raise RuntimeError(combined or "Repuestos TYTSERV worker failed.")

    output_path = request.output_path.resolve()
    if not output_path.is_file():
        raise FileNotFoundError(f"Generated artifact not found: {output_path}")

    lines = extract_lines(completed.stdout, completed.stderr)
    return ProcessResult(
        success=True,
        output_path=output_path,
        label="repuestos_tytserv",
        metadata={
            "console": "\n".join(lines).strip(),
            "summary": parse_summary(lines, file_fields),
            "output_origin": "default_path",
            "fallback_used": False,
            "command": command,
            "runtime": "python-legacy-node-wrapper",
        },
    )


def normalize_flag(flag: str) -> str:
    mapping = {
        "-inputtyt": "--input-tyt",
        "-inputpeug": "--input-peug",
        "-inputchgn": "--input-chgn",
        "-inputszk": "--input-szk",
        "-inputnctyt": "--input-nc-tyt",
        "-inputncpeug": "--input-nc-peug",
        "-inputncchgn": "--input-nc-chgn",
        "-inputncszk": "--input-nc-szk",
    }
    return mapping.get(flag.lower(), flag)


def parse_summary(lines: list[str], file_fields: list[dict]) -> list[dict[str, int | str]]:
    summary_map: dict[str, int] = {}
    for line in lines:
        if not line.startswith("INFO|") or "|rows=" not in line:
            continue

        _, key, rows_part = line.split("|", 2)
        rows_token = rows_part.split("rows=", 1)[1]
        try:
            summary_map[key.lower()] = int(rows_token)
        except ValueError:
            continue

    ordered_summary: list[dict[str, int | str]] = []
    for field_config in file_fields:
        summary_key = str(field_config.get("summary_key", "")).strip().lower()
        if not summary_key or summary_key not in summary_map:
            continue

        ordered_summary.append(
            {
                "label": str(field_config.get("summary_label", summary_key.upper())),
                "rows": summary_map[summary_key],
            }
        )

    return ordered_summary
