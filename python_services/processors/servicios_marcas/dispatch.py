from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from ..contracts import ProcessRequest, ProcessResult


def run(request: ProcessRequest) -> ProcessResult:
    options = request.options if isinstance(request.options, dict) else {}
    job_id = required_text_option(options, "job_id")
    worker_path = required_path_option(options, "worker_path")
    input_path = required_path_option(options, "input_path")
    output_dir = required_path_option(options, "output_dir")
    template_dir = required_path_option(options, "template_dir")
    jobs_dir = required_path_option(options, "jobs_dir")
    job_path = jobs_dir / f"servicios_marcas_{job_id}.json"

    if not worker_path.is_file():
        raise FileNotFoundError(f"Servicios worker not found: {worker_path}")
    if not input_path.is_file():
        raise FileNotFoundError(f"Servicios input file not found: {input_path}")
    if not output_dir.is_dir():
        raise FileNotFoundError(f"Servicios output directory not found: {output_dir}")
    if not template_dir.is_dir():
        raise FileNotFoundError(f"Servicios template directory not found: {template_dir}")
    if not jobs_dir.is_dir():
        raise FileNotFoundError(f"Servicios jobs directory not found: {jobs_dir}")
    if not job_path.is_file():
        raise FileNotFoundError(f"Servicios job snapshot not found: {job_path}")

    marker_path = request.output_path.resolve()
    marker_path.parent.mkdir(parents=True, exist_ok=True)
    marker_payload = {
        "job_id": job_id,
        "worker_path": str(worker_path),
        "input_path": str(input_path),
        "output_dir": str(output_dir),
        "template_dir": str(template_dir),
        "job_path": str(job_path),
        "bridge_mode": "python-preflight",
    }
    marker_path.write_text(json.dumps(marker_payload, indent=2, ensure_ascii=True), encoding="utf-8")

    return ProcessResult(
        success=True,
        output_path=marker_path,
        label="servicios_marcas.dispatch",
        metadata={
            "runtime": "python-preflight",
            "job_id": job_id,
            "dispatch_method": "python-preflight",
            "output_origin": "default_path",
            "fallback_used": False,
        },
    )


def required_text_option(options: dict[str, Any], key: str) -> str:
    value = str(options.get(key, "")).strip()
    if value == "":
        raise ValueError(f"Missing required option: {key}")
    return value


def required_path_option(options: dict[str, Any], key: str) -> Path:
    return Path(required_text_option(options, key)).resolve()
