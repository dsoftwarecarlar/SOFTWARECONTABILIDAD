from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[3]
OUTPUT_RETENTION_LIMIT = 3
UPLOAD_RETENTION_LIMIT = 20
JOB_RETENTION_LIMIT = 24
ACTIVE_STATUSES = {"queued", "running", "cancel_requested"}
OUTPUT_CONFIG: dict[str, dict[str, str]] = {
    "changan": {"label": "CHANGAN", "prefix": "servicios_changan_"},
    "peug": {"label": "PEUGEOT", "prefix": "servicios_peug_"},
    "szk": {"label": "SUZUKI", "prefix": "servicios_szk_"},
    "tyt": {"label": "MATRIZ", "prefix": "servicios_tyt_"},
}


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    try:
        run_job(args)
        return 0
    except Exception as exc:  # noqa: BLE001
        try:
            mark_job_error_from_args(args, str(exc))
        except Exception:
            pass
        sys.stderr.write(f"{exc}\n")
        return 1


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Worker Python para Servicios por Marca")
    parser.add_argument("--job", required=True)
    parser.add_argument("--input", dest="input_path", required=True)
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--template-dir", required=True)
    parser.add_argument("--jobs-dir", required=True)
    parser.add_argument("--worker-timeout-seconds", type=int, default=2700)
    parser.add_argument("--cancel-grace-seconds", type=int, default=120)
    parser.add_argument("--queued-timeout-seconds", type=int, default=300)
    parser.add_argument("--dispatch-boot-timeout-seconds", type=int, default=20)
    return parser.parse_args(argv)


def run_job(args: argparse.Namespace) -> None:
    job_id = str(args.job).strip()
    if re.fullmatch(r"[A-Za-z0-9_-]+", job_id) is None:
        raise RuntimeError("Job invalido.")

    input_path = Path(args.input_path).resolve()
    output_dir = Path(args.output_dir).resolve()
    template_dir = Path(args.template_dir).resolve()
    jobs_dir = Path(args.jobs_dir).resolve()
    job_path = jobs_dir / f"servicios_marcas_{job_id}.json"
    cancel_path = jobs_dir / f"servicios_marcas_{job_id}.stop"
    console_log_path = jobs_dir / f"servicios_marcas_{job_id}.log"

    if not input_path.is_file():
        raise RuntimeError("No se encontro el archivo Excel subido.")
    if not output_dir.is_dir():
        output_dir.mkdir(parents=True, exist_ok=True)
    if not output_dir.is_dir():
        raise RuntimeError("No se encontro la carpeta de salidas.")
    if not template_dir.is_dir():
        raise RuntimeError("No se encontro la carpeta de plantillas.")

    runtime_path = ROOT / "python_services" / "processors" / "servicios_marcas" / "runtime.py"
    if not runtime_path.is_file():
        raise RuntimeError(
            "No existe el runtime Python de servicios por marca en python_services/processors/servicios_marcas/runtime.py."
        )

    existing_job = read_job_snapshot(job_path)
    source_name = str(existing_job.get("source_name") or input_path.name).strip()
    run_stamp = str(existing_job.get("run_stamp") or "").strip()
    if run_stamp == "":
        run_stamp = time.strftime("%Y%m%d_%H%M%S") + "_" + os.urandom(2).hex()

    if cancel_path.is_file() or str(existing_job.get("status") or "") in {"cancel_requested", "cancelled"}:
        if cancel_path.is_file():
            cancel_path.unlink(missing_ok=True)
        write_status(
            job_path,
            existing_job,
            job_id,
            source_name,
            run_stamp,
            "cancelled",
            {
                "message": "Proceso detenido por solicitud del usuario.",
                "downloads": [],
                "summary": [],
                "console": "",
                "completed_at": now_string(),
            },
        )
        return

    write_status(
        job_path,
        existing_job,
        job_id,
        source_name,
        run_stamp,
        "running",
        {
            "message": "Procesando plantillas en segundo plano. Este paso puede tardar varios minutos.",
            "started_at": now_string(),
        },
    )

    uploads = existing_job.get("uploads") if isinstance(existing_job.get("uploads"), dict) else {}
    brand_key = str(existing_job.get("brand_key") or "").strip()
    command = build_worker_command(
        python_binary=resolve_python(),
        runtime_path=runtime_path,
        input_path=input_path,
        output_dir=output_dir,
        template_dir=template_dir,
        run_stamp=run_stamp,
        cancel_path=cancel_path,
        brand_key=brand_key,
        uploads=uploads,
    )

    console_text = ""
    started_ts = int(time.time())
    last_heartbeat_at = started_ts
    timed_out = False

    if console_log_path.is_file():
        console_log_path.unlink(missing_ok=True)

    exit_code = 0
    with console_log_path.open("ab") as log_stream:
        process = subprocess.Popen(command, stdout=log_stream, stderr=log_stream)  # noqa: S603
        try:
            while True:
                polled_exit_code = process.poll()
                if polled_exit_code is not None:
                    exit_code = int(polled_exit_code)
                    break

                if int(time.time()) - started_ts >= int(args.worker_timeout_seconds):
                    terminate_process_tree(process.pid)
                    timed_out = True
                    break

                if int(time.time()) - last_heartbeat_at >= 20:
                    elapsed = max(1, int(time.time()) - started_ts)
                    minutes, seconds = divmod(elapsed, 60)
                    current_job = read_job_snapshot(job_path)
                    if str(current_job.get("status") or "") != "cancel_requested":
                        write_status(
                            job_path,
                            current_job,
                            job_id,
                            source_name,
                            run_stamp,
                            "running",
                            {
                                "message": f"Procesando plantillas en segundo plano. Tiempo transcurrido: {minutes:02d}:{seconds:02d}.",
                            },
                        )
                    last_heartbeat_at = int(time.time())

                time.sleep(0.5)
        finally:
            if process.poll() is None:
                terminate_process_tree(process.pid)
            try:
                exit_code = int(process.wait(timeout=10))
            except subprocess.TimeoutExpired:
                terminate_process_tree(process.pid)

    output_lines = read_console_lines(console_log_path)
    console_log_path.unlink(missing_ok=True)
    console_text = "\n".join(output_lines).strip()

    if timed_out:
        raise RuntimeError(
            f"El worker de servicios supero el tiempo maximo permitido ({timeout_label(int(args.worker_timeout_seconds))}) y fue detenido para evitar bloqueos."
        )

    downloads: dict[str, dict[str, str]] = {}
    summary_by_key: dict[str, dict[str, Any]] = {}
    cancel_message = ""

    for line in output_lines:
        text = line.strip()
        cancel_pos = text.find("CANCELLED|")
        if cancel_pos != -1:
            cancel_message = text[cancel_pos + len("CANCELLED|") :].strip()
            continue

        info_pos = text.find("INFO|")
        if info_pos != -1:
            text = text[info_pos:]

        processing_match = re.fullmatch(r"INFO\|processing\|([a-z0-9_]+)\|rows=(\d+)", text, re.IGNORECASE)
        if processing_match:
            key = processing_match.group(1).lower()
            item = summary_by_key.setdefault(key, {"key": key})
            item["rows"] = int(processing_match.group(2))
            continue

        fallback_match = re.fullmatch(
            r"INFO\|([a-z0-9_]+)\|invoice_fallbacks=(\d+)\|note_fallbacks=(\d+)",
            text,
            re.IGNORECASE,
        )
        if fallback_match:
            key = fallback_match.group(1).lower()
            item = summary_by_key.setdefault(key, {"key": key})
            item["invoice_fallbacks"] = int(fallback_match.group(2))
            item["note_fallbacks"] = int(fallback_match.group(3))
            continue

        output_pos = text.find("OUTPUT|")
        if output_pos == -1:
            continue

        text = text[output_pos:]
        parts = text.split("|", 2)
        if len(parts) < 3:
            continue

        _, file_name, label = parts
        file_name = file_name.strip()
        key = Path(file_name).stem.lower()
        if key in downloads:
            continue

        downloads[key] = {
            "label": label.strip(),
            "name": file_name,
        }

    if not downloads:
        collect_generated_downloads(downloads, output_dir, run_stamp, started_ts)

    summary_list = []
    for key in OUTPUT_CONFIG.keys():
        item = summary_by_key.get(key)
        if item is None:
            continue
        summary_list.append(
            {
                "key": key,
                "label": OUTPUT_CONFIG[key]["label"],
                "rows": int(item.get("rows") or 0),
                "invoice_fallbacks": int(item.get("invoice_fallbacks") or 0),
                "note_fallbacks": int(item.get("note_fallbacks") or 0),
            }
        )

    if cancel_message != "":
        delete_generated_outputs(output_dir, run_stamp)
        cancel_path.unlink(missing_ok=True)
        write_status(
            job_path,
            read_job_snapshot(job_path),
            job_id,
            source_name,
            run_stamp,
            "cancelled",
            {
                "message": cancel_message,
                "downloads": [],
                "summary": summary_list,
                "console": console_text,
                "completed_at": now_string(),
            },
        )
        return

    if exit_code != 0:
        raise RuntimeError(console_text or "El script de servicios termino con error.")

    if not downloads:
        raise RuntimeError("El proceso termino sin generar archivos de salida.")

    cleanup_outputs(output_dir)
    cleanup_uploads(ROOT / "storage" / "uploads")
    cleanup_jobs(jobs_dir)
    cancel_path.unlink(missing_ok=True)

    write_status(
        job_path,
        read_job_snapshot(job_path),
        job_id,
        source_name,
        run_stamp,
        "complete",
        {
            "message": "Proceso terminado. Ya puedes descargar las plantillas generadas.",
            "downloads": list(downloads.values()),
            "summary": summary_list,
            "console": console_text,
            "completed_at": now_string(),
        },
    )


def mark_job_error_from_args(args: argparse.Namespace, error_message: str) -> None:
    job_id = str(getattr(args, "job", "") or "").strip()
    jobs_dir_arg = str(getattr(args, "jobs_dir", "") or "").strip()
    if job_id == "" or jobs_dir_arg == "":
        return

    jobs_dir = Path(jobs_dir_arg).resolve()
    job_path = jobs_dir / f"servicios_marcas_{job_id}.json"
    current = read_job_snapshot(job_path)
    if current == {}:
        return

    input_name = Path(str(getattr(args, "input_path", "") or "")).name
    source_name = str(current.get("source_name") or input_name or "servicios_marcas.xlsx").strip()
    run_stamp = str(current.get("run_stamp") or time.strftime("%Y%m%d_%H%M%S")).strip()
    console_log_path = jobs_dir / f"servicios_marcas_{job_id}.log"
    console_text = "\n".join(read_console_lines(console_log_path)).strip()
    cancel_path = jobs_dir / f"servicios_marcas_{job_id}.stop"
    cancel_path.unlink(missing_ok=True)
    console_log_path.unlink(missing_ok=True)

    write_status(
        job_path,
        current,
        job_id,
        source_name,
        run_stamp,
        "error",
        {
            "message": "El proceso termino con error.",
            "error": error_message,
            "downloads": [],
            "summary": current.get("summary") if isinstance(current.get("summary"), list) else [],
            "console": console_text,
            "completed_at": now_string(),
        },
    )


def build_worker_command(
    *,
    python_binary: str,
    runtime_path: Path,
    input_path: Path,
    output_dir: Path,
    template_dir: Path,
    run_stamp: str,
    cancel_path: Path,
    brand_key: str,
    uploads: dict[str, Any],
) -> list[str]:
    def uploaded_path(key: str) -> str:
        value = str(uploads.get(key) or "").strip()
        return value if value and Path(value).is_file() else ""

    return [
        python_binary,
        str(runtime_path),
        "--input",
        str(input_path),
        "--output-dir",
        str(output_dir),
        "--template-dir",
        str(template_dir),
        "--run-stamp",
        run_stamp,
        "--cancel-path",
        str(cancel_path),
        "--brand-key",
        brand_key,
        "--px-path",
        uploaded_path("px_file"),
        "--repvtas-path",
        uploaded_path("repventas_file"),
        "--factura-changan-path",
        uploaded_path("factura_changan_file"),
        "--nota-changan-path",
        uploaded_path("nota_changan_file"),
        "--mayor-changan-path",
        uploaded_path("mayor_changan_file"),
        "--factura-peug-path",
        uploaded_path("factura_peug_file"),
        "--nota-peug-path",
        uploaded_path("nota_peug_file"),
        "--mayor-peug-path",
        uploaded_path("mayor_peug_file"),
        "--factura-szk-path",
        uploaded_path("factura_szk_file"),
        "--nota-szk-path",
        uploaded_path("nota_szk_file"),
        "--mayor-szk-path",
        uploaded_path("mayor_szk_file"),
        "--factura-tyt-path",
        uploaded_path("factura_tyt_file"),
        "--nota-tyt-path",
        uploaded_path("nota_tyt_file"),
        "--mayor-tyt-path",
        uploaded_path("mayor_tyt_file"),
    ]


def resolve_python() -> str:
    candidate = Path(sys.executable)
    if candidate.is_file():
        return str(candidate)
    return "python"


def timeout_label(seconds: int) -> str:
    if seconds <= 0:
        return "0 minutos"
    if seconds % 60 == 0:
        minutes = seconds // 60
        return f"{minutes} minuto" + ("" if minutes == 1 else "s")
    minutes, remaining_seconds = divmod(seconds, 60)
    if minutes <= 0:
        return f"{remaining_seconds} segundo" + ("" if remaining_seconds == 1 else "s")
    return (
        f"{minutes} minuto" + ("" if minutes == 1 else "s")
        + f" {remaining_seconds} segundo"
        + ("" if remaining_seconds == 1 else "s")
    )


def now_string() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


def read_job_snapshot(path: Path) -> dict[str, Any]:
    if not path.is_file():
        return {}
    raw = path.read_text(encoding="utf-8", errors="ignore").strip()
    if raw == "":
        return {}
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return {}
    return data if isinstance(data, dict) else {}


def write_job_snapshot(path: Path, payload: dict[str, Any]) -> None:
    path.write_text(
        json.dumps(payload, indent=2, ensure_ascii=True),
        encoding="utf-8",
    )


def write_status(
    job_path: Path,
    current: dict[str, Any],
    job_id: str,
    source_name: str,
    run_stamp: str,
    status: str,
    extra: dict[str, Any],
) -> None:
    payload = dict(current)
    payload.update(
        {
            "job_id": job_id,
            "status": status,
            "source_name": source_name,
            "run_stamp": run_stamp,
            "updated_at": now_string(),
        }
    )
    payload.update(extra)
    write_job_snapshot(job_path, payload)


def read_console_lines(path: Path) -> list[str]:
    if not path.is_file():
        return []
    raw = path.read_bytes()
    raw = raw.replace(b"\x00", b"")
    if raw.startswith(b"\xef\xbb\xbf"):
        raw = raw[3:]
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        raw = raw[2:]
    text = raw.decode("utf-8", errors="ignore")
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    return [line.strip() for line in normalized.split("\n") if line.strip()]


def terminate_process_tree(pid: int) -> None:
    if pid <= 0:
        return
    if os.name == "nt":
        subprocess.run(  # noqa: S603
            ["cmd", "/C", "taskkill", "/PID", str(pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
        return
    subprocess.run(["kill", "-TERM", str(pid)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)  # noqa: S603


def collect_generated_downloads(
    downloads: dict[str, dict[str, str]],
    output_dir: Path,
    run_stamp: str,
    started_ts: int,
) -> None:
    for generated_path in output_dir.glob(f"servicios_*_{run_stamp}.xls"):
        if not generated_path.is_file():
            continue
        generated_name = generated_path.name
        key = generated_path.stem.lower()
        if key in downloads:
            continue
        downloads[key] = {
            "label": detect_output_label(generated_name),
            "name": generated_name,
        }

    if downloads:
        return

    for generated_path in output_dir.glob("servicios_*.xls"):
        if not generated_path.is_file():
            continue
        timestamp = safe_file_timestamp(generated_path)
        if timestamp < (started_ts - 10):
            continue
        generated_name = generated_path.name
        key = generated_path.stem.lower()
        if key in downloads:
            continue
        downloads[key] = {
            "label": detect_output_label(generated_name),
            "name": generated_name,
        }


def detect_output_label(file_name: str) -> str:
    lowered = file_name.lower()
    for brand_key, config in OUTPUT_CONFIG.items():
        if lowered.startswith(f"servicios_{brand_key}_"):
            return config["label"]
    return "SALIDA"


def delete_generated_outputs(output_dir: Path, run_stamp: str) -> None:
    for path in output_dir.glob(f"servicios_*_{run_stamp}.xls"):
        if path.is_file():
            path.unlink(missing_ok=True)


def cleanup_outputs(output_dir: Path) -> None:
    for config in OUTPUT_CONFIG.values():
        prefix = config["prefix"].lower()
        files = [
            path
            for path in output_dir.glob(f"{prefix}*.xls")
            if path.is_file()
        ]
        files.sort(key=safe_file_timestamp, reverse=True)
        for stale in files[OUTPUT_RETENTION_LIMIT:]:
            stale.unlink(missing_ok=True)


def cleanup_uploads(uploads_dir: Path) -> None:
    if not uploads_dir.is_dir():
        return
    files = [path for path in uploads_dir.iterdir() if path.is_file()]
    files.sort(key=safe_file_timestamp, reverse=True)
    for stale in files[UPLOAD_RETENTION_LIMIT:]:
        stale.unlink(missing_ok=True)


def cleanup_jobs(jobs_dir: Path) -> None:
    if not jobs_dir.is_dir():
        return

    files = [path for path in jobs_dir.glob("servicios_marcas_*.json") if path.is_file()]
    files.sort(key=safe_file_timestamp, reverse=True)
    for stale in files[JOB_RETENTION_LIMIT:]:
        stale_job_id = re.sub(r"^servicios_marcas_|\.json$", "", stale.name)
        cancel_path = jobs_dir / f"servicios_marcas_{stale_job_id}.stop"
        stale.unlink(missing_ok=True)
        if cancel_path.is_file():
            cancel_path.unlink(missing_ok=True)

    for cancel_file in jobs_dir.glob("servicios_marcas_*.stop"):
        job_id = re.sub(r"^servicios_marcas_|\.stop$", "", cancel_file.name)
        job_path = jobs_dir / f"servicios_marcas_{job_id}.json"
        if not job_path.is_file():
            cancel_file.unlink(missing_ok=True)


def safe_file_timestamp(path: Path) -> int:
    try:
        stat = path.stat()
    except OSError:
        return 0
    return int(stat.st_ctime or stat.st_mtime or 0)


if __name__ == "__main__":
    raise SystemExit(main())
