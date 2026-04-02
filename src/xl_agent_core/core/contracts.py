from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(slots=True)
class WarningMessage:
    code: str
    message: str
    severity: str = "warning"


@dataclass(slots=True)
class SourceRef:
    workbook_path: str
    sheet: str | None = None
    range_ref: str | None = None
    named_item: str | None = None


@dataclass(slots=True)
class SheetOverview:
    name: str
    state: str
    used_range: str | None
    non_empty_cells: int
    max_row: int
    max_column: int
    hidden_row_count: int
    hidden_column_count: int
    merged_range_count: int
    table_count: int


@dataclass(slots=True)
class WorkbookOverview:
    workbook_path: str
    sheet_count: int
    active_sheet: str
    sheet_order: list[str]
    calculation_mode: str | None
    defined_name_count: int
    sheets: list[SheetOverview]


@dataclass(slots=True)
class RegionCandidate:
    sheet: str
    label: str
    range_ref: str
    header_row: int | None
    non_empty_cells: int
    density: float
    width: int
    height: int
    notes: list[str] = field(default_factory=list)


@dataclass(slots=True)
class NamedItem:
    name: str
    scope: str
    refers_to: str
    destinations: list[str]
    is_formula_like: bool


@dataclass(slots=True)
class LayoutCell:
    sheet: str
    ref: str
    row: int
    column: int
    value: Any
    displayed_value: Any
    formula: str | None
    data_type: str
    is_merged: bool
    merged_range: str | None
    row_hidden: bool
    column_hidden: bool
    style: dict[str, Any]


@dataclass(slots=True)
class LayoutReadResult:
    workbook_path: str
    sheet: str
    range_ref: str
    cells: list[LayoutCell]
    merged_ranges: list[str]
    hidden_rows: list[int]
    hidden_columns: list[str]
    filter_ref: str | None = None
    warnings: list[WarningMessage] = field(default_factory=list)


@dataclass(slots=True)
class CellInspection:
    sheet: str
    ref: str
    value: Any
    displayed_value: Any
    formula: str | None
    data_type: str
    is_merged: bool
    merged_range: str | None
    row_hidden: bool
    column_hidden: bool


@dataclass(slots=True)
class TableSchemaColumn:
    name: str
    ref: str
    index: int


@dataclass(slots=True)
class TableReadResult:
    workbook_path: str
    sheet: str
    region: str
    header: list[str]
    schema: list[TableSchemaColumn]
    rows: list[dict[str, Any]]
    grid: list[list[Any]]
    row_count: int
    column_count: int
    warnings: list[WarningMessage] = field(default_factory=list)


@dataclass(slots=True)
class FormulaTraceLink:
    sheet: str
    ref: str
    kind: str
    formula_cell: str
    formula: str | None = None


@dataclass(slots=True)
class FormulaTraceResult:
    workbook_path: str
    sheet: str
    cell: str
    direction: str
    formula: str | None
    links: list[FormulaTraceLink]
    warnings: list[WarningMessage] = field(default_factory=list)


@dataclass(slots=True)
class MutationPlan:
    operation: str
    workbook_path: str
    sheet: str | None
    output_path: str | None
    targets: list[str]
    edits: list["MutationEdit"] = field(default_factory=list)
    dry_run: bool = True
    copy_on_write: bool = True
    overwrite: bool = False
    status: str = "planned"
    warnings: list[WarningMessage] = field(default_factory=list)


@dataclass(slots=True)
class MutationEdit:
    sheet: str
    ref: str
    kind: str
    before_value: Any
    after_value: Any
    before_formula: str | None
    after_formula: str | None
    note: str | None = None


@dataclass(slots=True)
class WorkbookCopyResult:
    workbook_path: str
    source_path: str
    output_path: str
    copied: bool
    overwrite: bool = False
    warnings: list[WarningMessage] = field(default_factory=list)


@dataclass(slots=True)
class MutationResult:
    workbook_path: str
    output_path: str
    operation: str
    saved: bool
    edit_count: int
    sheet: str | None
    edits: list[MutationEdit] = field(default_factory=list)
    recalc: "RecalcResult | None" = None
    warnings: list[WarningMessage] = field(default_factory=list)


@dataclass(slots=True)
class ProofResult:
    baseline_path: str
    current_path: str
    targets: list["ProofTargetResult"]
    compared_cell_count: int
    changed_cell_count: int
    status: str = "ok"
    warnings: list[WarningMessage] = field(default_factory=list)


@dataclass(slots=True)
class DiffResult:
    original_path: str
    modified_path: str
    changed_cells: int
    summary: dict[str, int]
    entries: list["DiffEntry"] = field(default_factory=list)
    added_sheets: list[str] = field(default_factory=list)
    removed_sheets: list[str] = field(default_factory=list)
    status: str = "ok"
    warnings: list[WarningMessage] = field(default_factory=list)


@dataclass(slots=True)
class RecalcResult:
    workbook_path: str
    backend_requested: str
    backend_used: str
    recalculated: bool
    calculation_mode: str | None
    full_rebuild: bool = False
    warnings: list[WarningMessage] = field(default_factory=list)


@dataclass(slots=True)
class DiffEntry:
    sheet: str
    ref: str
    classification: str
    before_value: Any
    after_value: Any
    before_formula: str | None
    after_formula: str | None


@dataclass(slots=True)
class ProofCellResult:
    sheet: str
    ref: str
    classification: str | None
    before_value: Any
    after_value: Any
    before_formula: str | None
    after_formula: str | None
    changed: bool


@dataclass(slots=True)
class ProofTargetResult:
    target: str
    cells: list[ProofCellResult]
    changed_cell_count: int
    warnings: list[WarningMessage] = field(default_factory=list)


@dataclass(slots=True)
class AgentResponseEnvelope:
    operation: str
    status: str
    data: Any
    warnings: list[WarningMessage] = field(default_factory=list)
    sources: list[SourceRef] = field(default_factory=list)
    refusal_reason: str | None = None
    metadata: dict[str, Any] = field(default_factory=dict)
