from __future__ import annotations

import json
from dataclasses import asdict, is_dataclass
from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path
from typing import Any

from xl_agent_core.core.contracts import AgentResponseEnvelope, SourceRef, WarningMessage


def serialize(value: Any) -> Any:
    if is_dataclass(value):
        return serialize(asdict(value))
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, time):
        return value.isoformat()
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, dict):
        return {str(key): serialize(item) for key, item in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [serialize(item) for item in value]
    return value


def make_envelope(
    operation: str,
    data: Any,
    *,
    status: str = "ok",
    warnings: list[WarningMessage] | None = None,
    sources: list[SourceRef] | None = None,
    refusal_reason: str | None = None,
    metadata: dict[str, Any] | None = None,
) -> AgentResponseEnvelope:
    envelope_warnings = list(warnings or [])
    envelope_sources = list(sources or [])

    data_warnings = getattr(data, "warnings", None)
    if data_warnings:
        envelope_warnings.extend(data_warnings)

    workbook_path = getattr(data, "workbook_path", None)
    sheet = getattr(data, "sheet", None)
    range_ref = getattr(data, "region", None) or getattr(data, "range_ref", None)
    if workbook_path:
        envelope_sources.append(
            SourceRef(workbook_path=workbook_path, sheet=sheet, range_ref=range_ref),
        )

    return AgentResponseEnvelope(
        operation=operation,
        status=status,
        data=data,
        warnings=envelope_warnings,
        sources=envelope_sources,
        refusal_reason=refusal_reason,
        metadata=metadata or {},
    )


def render_json(value: Any, *, pretty: bool = False) -> str:
    indent = 2 if pretty else None
    separators = None if pretty else (",", ":")
    return json.dumps(serialize(value), indent=indent, sort_keys=True, separators=separators)
