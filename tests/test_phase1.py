from __future__ import annotations

import io
import json
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from xl_agent_core.cli.main import main
from xl_agent_core.core.formulas import FormulaService
from xl_agent_core.core.readers import ReadService
from xl_agent_core.core.recon import WorkbookReconService


class Phase1Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.temp_dir = tempfile.TemporaryDirectory()
        cls.workbook_path = Path(cls.temp_dir.name) / "sample.xlsx"
        cls._build_workbook(cls.workbook_path)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.temp_dir.cleanup()

    @classmethod
    def _build_workbook(cls, path: Path) -> None:
        wb = Workbook()
        model = wb.active
        model.title = "Model"
        model["A1"] = "Item"
        model["B1"] = "Qty"
        model["C1"] = "Total"
        model["A2"] = "Apple"
        model["B2"] = 2
        model["C2"] = "=B2*10"
        model["A3"] = "Orange"
        model["B3"] = 3
        model["C3"] = "=B3*10"
        model["A4"] = "Pear"
        model["B4"] = 4
        model["C4"] = "=B4*10"
        model.row_dimensions[4].hidden = True
        model["E1"] = "Rate"
        model["F1"] = "Value"
        model["E2"] = "VAT"
        model["F2"] = 0.05
        model.auto_filter.ref = "A1:C4"

        summary = wb.create_sheet("Summary")
        summary["A1"] = "Metric"
        summary["B1"] = "Value"
        summary["A2"] = "Grand Total"
        summary["B2"] = "=SUM(Model!C2:C4)"

        wb.defined_names.add(DefinedName("ModelInputs", attr_text="Model!$A$1:$C$4"))
        wb.save(path)

    def test_probe_lists_sheets_and_names(self) -> None:
        recon = WorkbookReconService()
        overview = recon.probe(str(self.workbook_path))

        self.assertEqual(overview.sheet_count, 2)
        self.assertEqual(overview.active_sheet, "Model")
        self.assertEqual(overview.defined_name_count, 1)
        self.assertEqual([sheet.name for sheet in overview.sheets], ["Model", "Summary"])

    def test_detect_regions_and_read_table(self) -> None:
        recon = WorkbookReconService()
        reader = ReadService()

        regions = recon.detect_regions(str(self.workbook_path), "Model")
        self.assertTrue(any(region.range_ref == "A1:C4" for region in regions))

        table = reader.read_table(str(self.workbook_path), "Model", "auto")
        self.assertEqual(table.region, "A1:C4")
        self.assertEqual(table.header, ["Item", "Qty", "Total"])
        self.assertEqual(table.row_count, 3)
        self.assertTrue(any(warning.code == "hidden_rows_in_range" for warning in table.warnings))

    def test_formula_trace_precedents_and_dependents(self) -> None:
        service = FormulaService()

        precedents = service.formula_trace(str(self.workbook_path), "Summary", "B2", "precedents")
        self.assertTrue(any(link.sheet == "Model" and link.ref == "C2:C4" for link in precedents.links))

        dependents = service.formula_trace(str(self.workbook_path), "Model", "C3", "dependents")
        self.assertTrue(any(link.sheet == "Summary" and link.ref == "B2" for link in dependents.links))

    def test_cli_outputs_json(self) -> None:
        buffer = io.StringIO()
        with redirect_stdout(buffer):
            exit_code = main(["read", "table", str(self.workbook_path), "--sheet", "Model", "--region", "auto"])

        payload = json.loads(buffer.getvalue())
        self.assertEqual(exit_code, 0)
        self.assertEqual(payload["operation"], "read.table")
        self.assertEqual(payload["status"], "ok")
        self.assertEqual(payload["data"]["region"], "A1:C4")


if __name__ == "__main__":
    unittest.main()
