from __future__ import annotations

from .contracts import ProcessRequest, ProcessResult


def run(request: ProcessRequest) -> ProcessResult:
    return ProcessResult(
        success=True,
        output_path=request.output_path,
        label="probe",
        metadata={
            "console": "python scaffold placeholder",
            "output_origin": "probe",
            "fallback_used": False,
        },
    )

