from __future__ import annotations

from collections.abc import Callable

from .contracts import ProcessRequest, ProcessResult
from .cxp_actions import accion1, accion2, accion3, accion4, export_all
from .probe import run as probe_run
from .repuestos_tytserv import (
    mayor_iva_stage as repuestos_mayor_iva_stage,
    my_stage as repuestos_my_stage,
    nc_stage as repuestos_nc_stage,
    process as repuestos_process,
    process_legacy as repuestos_process_legacy,
    rep_stage as repuestos_rep_stage,
)
from .servicios_marcas import dispatch as servicios_marcas_dispatch

Processor = Callable[[ProcessRequest], ProcessResult]

PROCESSORS: dict[str, Processor] = {
    "probe": probe_run,
    "cxp_actions.accion1": accion1.run,
    "cxp_actions.accion2": accion2.run,
    "cxp_actions.accion3": accion3.run,
    "cxp_actions.accion4": accion4.run,
    "cxp_actions.export_all": export_all.run,
    "repuestos_tytserv.process": repuestos_process.run,
    "repuestos_tytserv.process_legacy": repuestos_process_legacy.run,
    "repuestos_tytserv.rep_stage": repuestos_rep_stage.run,
    "repuestos_tytserv.nc_stage": repuestos_nc_stage.run,
    "repuestos_tytserv.my_stage": repuestos_my_stage.run,
    "repuestos_tytserv.mayor_iva_stage": repuestos_mayor_iva_stage.run,
    "servicios_marcas.dispatch": servicios_marcas_dispatch.run,
}


def dispatch(processor_name: str, request: ProcessRequest) -> ProcessResult:
    processor = PROCESSORS.get(processor_name)
    if processor is None:
        raise ValueError(f"Processor not registered: {processor_name}")

    return processor(request)
