# Creado por Francisco Grandón Vergara
"""
Modelos de datos y TypedDict para solicitudes enviadas al Daemon y comandos de Antigravity Excel Engine.
"""

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Union


@dataclass
class DaemonRequest:
    """Solicitud enviada al Daemon residente de Excel."""
    action: Optional[str] = None
    command: Optional[str] = None
    args: Dict[str, Any] = field(default_factory=dict)
    request_id: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {}
        if self.action is not None:
            data["action"] = self.action
        if self.command is not None:
            data["command"] = self.command
        data["args"] = self.args
        if self.request_id is not None:
            data["request_id"] = self.request_id
        return data


@dataclass
class PipeCommandRequest:
    """Solicitud IPC simplificada de ejecución de comando a través de Named Pipe."""
    command: str
    args: Dict[str, Any] = field(default_factory=dict)
    file_path: Optional[str] = None
    mode: Optional[str] = None

    def to_daemon_request(self) -> DaemonRequest:
        merged_args = dict(self.args)
        if self.file_path and "file" not in merged_args:
            merged_args["file"] = self.file_path
        if self.mode and "mode" not in merged_args:
            merged_args["mode"] = self.mode
        return DaemonRequest(command=self.command, args=merged_args)


@dataclass
class CellWriteRequest:
    """Solicitud de escritura individual o por lote en celdas de Excel."""
    cell: str
    value: Any
    sheet: Optional[str] = None
    number_format: Optional[str] = None
    formula: bool = False


@dataclass
class CuratePipelineRequest:
    """Solicitud para ejecutar un pipeline compuesto de curación de datos."""
    target_range_or_table: str
    pipeline_spec: Dict[str, Any]
    sheet: Optional[str] = None
