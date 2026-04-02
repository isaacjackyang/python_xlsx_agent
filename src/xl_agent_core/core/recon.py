from __future__ import annotations

from xl_agent_core.core.contracts import SheetOverview, WorkbookOverview
from xl_agent_core.core.loader import is_populated, load_workbook_bundle
from xl_agent_core.core.names import read_named_items
from xl_agent_core.core.regions import RegionDetector


class WorkbookReconService:
    def __init__(self) -> None:
        self._regions = RegionDetector()

    def probe(self, path: str) -> WorkbookOverview:
        bundle = load_workbook_bundle(path)
        sheets = self.list_sheets(path)
        return WorkbookOverview(
            workbook_path=str(bundle.path),
            sheet_count=len(bundle.formulas.sheetnames),
            active_sheet=bundle.formulas.active.title,
            sheet_order=list(bundle.formulas.sheetnames),
            calculation_mode=getattr(bundle.formulas.calculation, "calcMode", None),
            defined_name_count=len(bundle.formulas.defined_names),
            sheets=sheets,
        )

    def list_sheets(self, path: str) -> list[SheetOverview]:
        bundle = load_workbook_bundle(path)
        items: list[SheetOverview] = []

        for worksheet in bundle.formulas.worksheets:
            non_empty_cells = sum(
                1
                for row in worksheet.iter_rows()
                for cell in row
                if is_populated(cell.value)
            )
            hidden_row_count = sum(1 for row in worksheet.row_dimensions.values() if row.hidden)
            hidden_column_count = sum(1 for column in worksheet.column_dimensions.values() if column.hidden)
            used_range = worksheet.calculate_dimension()
            if used_range == "A1:A1" and not is_populated(worksheet["A1"].value):
                used_range = None

            items.append(
                SheetOverview(
                    name=worksheet.title,
                    state=worksheet.sheet_state,
                    used_range=used_range,
                    non_empty_cells=non_empty_cells,
                    max_row=worksheet.max_row,
                    max_column=worksheet.max_column,
                    hidden_row_count=hidden_row_count,
                    hidden_column_count=hidden_column_count,
                    merged_range_count=len(worksheet.merged_cells.ranges),
                    table_count=len(getattr(worksheet, "tables", {})),
                ),
            )

        return items

    def read_names(self, path: str):
        return read_named_items(path)

    def detect_regions(self, path: str, sheet: str):
        return self._regions.detect(path, sheet)
