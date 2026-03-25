from __future__ import annotations

import subprocess
import os
import shutil
from pathlib import Path

from .contracts import ProcessRequest, ProcessResult


def run_legacy_node_processor(
    request: ProcessRequest,
    *,
    label: str,
    command_mode: str,
    generated_path_prefix: str,
    ignore_prefixes: list[str],
) -> ProcessResult:
    script_path = Path(str(request.options.get("script_path", ""))).resolve()
    if not script_path.is_file():
        raise FileNotFoundError(f"Legacy script not found: {script_path}")

    output_path = request.output_path.resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    command = build_command(command_mode, script_path, request)
    cwd = Path(str(request.options.get("cwd", script_path.parent))).resolve()
    timeout = int(request.options.get("timeout_seconds", 300))
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
        raise RuntimeError(combined or "Legacy worker failed.")

    lines = extract_lines(completed.stdout, completed.stderr)
    resolved_path, origin, fallback_used = resolve_generated_artifact(
        output_path=output_path,
        lines=lines,
        prefix=generated_path_prefix,
    )
    if not resolved_path.is_file():
        raise FileNotFoundError(f"Generated artifact not found: {resolved_path}")

    return ProcessResult(
        success=True,
        output_path=resolved_path,
        label=label,
        metadata={
            "console": filter_console_lines(lines, ignore_prefixes),
            "output_origin": origin,
            "fallback_used": fallback_used,
            "command": command,
        },
    )


def build_command(command_mode: str, script_path: Path, request: ProcessRequest) -> list[str]:
    command = [resolve_node_binary(request), str(script_path)]

    if command_mode == "positional_action1":
        if len(request.input_paths) != 1:
            raise ValueError("Accion 1 requiere exactamente un archivo de entrada.")
        return command + [str(request.input_paths[0]), str(request.output_path)]

    if command_mode == "positional_with_template":
        if len(request.input_paths) != 1:
            raise ValueError("Este procesador requiere exactamente un archivo de entrada.")
        if request.template_path is None:
            raise ValueError("La plantilla es obligatoria para este procesador.")
        return command + [str(request.input_paths[0]), str(request.output_path), str(request.template_path)]

    if command_mode == "flagged_with_template":
        if request.template_path is None:
            raise ValueError("La plantilla es obligatoria para este procesador.")
        if not request.input_paths:
            raise ValueError("Este procesador requiere al menos un archivo de entrada.")
        return command + [str(path) for path in request.input_paths] + [
            "--output",
            str(request.output_path),
            "--template",
            str(request.template_path),
        ]

    if command_mode == "bundle_output_flag":
        return command + ["--output", str(request.output_path)]

    raise ValueError(f"Unsupported command mode: {command_mode}")


def resolve_node_binary(request: ProcessRequest) -> str:
    configured = str(request.options.get("node_binary", "")).strip()
    if configured:
        configured_path = Path(configured)
        if configured_path.is_file():
            return str(configured_path.resolve())
        return configured

    discovered = shutil.which("node")
    if discovered:
        return discovered

    common_locations = [
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "nodejs" / "node.exe",
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "nodejs" / "node.exe",
    ]
    for location in common_locations:
        if location.is_file():
            return str(location.resolve())

    raise FileNotFoundError("Node binary not found for Python legacy bridge.")


def extract_lines(stdout: str, stderr: str) -> list[str]:
    combined = "\n".join(part for part in (stdout, stderr) if part)
    return [line.strip() for line in combined.splitlines() if line.strip()]


def filter_console_lines(lines: list[str], ignore_prefixes: list[str]) -> str:
    filtered: list[str] = []
    for line in lines:
        if not line:
            continue
        ignored = False
        for prefix in ignore_prefixes:
            if prefix and line.lower().startswith(prefix.lower()):
                ignored = True
                break
        if not ignored:
            filtered.append(line)
    return "\n".join(filtered).strip()


def resolve_generated_artifact(
    *,
    output_path: Path,
    lines: list[str],
    prefix: str,
) -> tuple[Path, str, bool]:
    default_path = output_path.resolve()
    if not prefix:
        return default_path, "default_path", False

    for line in lines:
        if not line.lower().startswith(prefix.lower()):
            continue
        reported = line[len(prefix) :].strip()
        if not reported:
            continue
        reported_path = Path(reported).resolve()
        if reported_path.is_file():
            return reported_path, "console_path", False

    return default_path, "default_path", True
