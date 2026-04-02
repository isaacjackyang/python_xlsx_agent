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

from xl_agent_core.cli.main import main
from xl_agent_core.core.mutate import MutationService
from xl_agent_core.core.recalc import RecalcService
from xl_agent_core.core.verify import VerifyService


def build_sample_workbook(path: Path) -> None:
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

    summary = wb.create_sheet("Summary")
    summary["A1"] = "Metric"
    summary["B1"] = "Value"
    summary["A2"] = "Grand Total"
    summary["B2"] = "=SUM(Model!C2:C4)"
    wb.save(path)


class Phase2Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.temp_dir = tempfile.TemporaryDirectory()
        cls.base_path = Path(cls.temp_dir.name) / "baseline.xlsx"
        build_sample_workbook(cls.base_path)
        cls.mutate = MutationService()
        cls.verify = VerifyService()
        cls.recalc = RecalcService()
        cls.has_excel_com = "excel_com" in cls.recalc.available_backends()
        if cls.has_excel_com:
            cls.recalc.recalculate(str(cls.base_path), backend="excel_com")

    @classmethod
    def tearDownClass(cls) -> None:
        cls.temp_dir.cleanup()

    def test_copy_and_dry_run_plan(self) -> None:
        copied_path = Path(self.temp_dir.name) / "copied.xlsx"
        result = self.mutate.copy_workbook(str(self.base_path), str(copied_path), overwrite=True)
        self.assertTrue(copied_path.exists())
        self.assertTrue(result.copied)

        plan = self.mutate.plan_write(
            str(self.base_path),
            "Model",
            {"B2": 5},
            output_path=str(copied_path),
            overwrite=True,
            dry_run=True,
        )
        self.assertEqual(plan.operation, "write_cells")
        self.assertEqual(plan.targets, ["B2"])
        self.assertTrue(plan.copy_on_write)

    def test_apply_write_and_diff(self) -> None:
        modified_path = Path(self.temp_dir.name) / "modified_values.xlsx"
        result = self.mutate.apply_write(
            str(self.base_path),
            "Model",
            {"B2": 5},
            output_path=str(modified_path),
            overwrite=True,
        )
        self.assertTrue(result.saved)
        diff = self.verify.diff(str(self.base_path), str(modified_path))
        self.assertTrue(any(entry.sheet == "Model" and entry.ref == "B2" and entry.classification == "value_changed" for entry in diff.entries))

    def test_replace_formula_and_diff(self) -> None:
        modified_path = Path(self.temp_dir.name) / "modified_formula.xlsx"
        result = self.mutate.replace_formulas(
            str(self.base_path),
            "Model",
            {"C2": "=B2*11"},
            output_path=str(modified_path),
            overwrite=True,
        )
        self.assertTrue(result.saved)
        diff = self.verify.diff(str(self.base_path), str(modified_path))
        self.assertTrue(any(entry.sheet == "Model" and entry.ref == "C2" and entry.classification == "formula_changed" for entry in diff.entries))

    def test_cli_phase2_dry_run(self) -> None:
        output_path = Path(self.temp_dir.name) / "cli_plan.xlsx"
        buffer = io.StringIO()
        with redirect_stdout(buffer):
            exit_code = main(
                [
                    "mutate",
                    "write-cells",
                    str(self.base_path),
                    "--sheet",
                    "Model",
                    "--set",
                    "B2=5",
                    "--output",
                    str(output_path),
                    "--dry-run",
                ]
            )

        payload = json.loads(buffer.getvalue())
        self.assertEqual(exit_code, 0)
        self.assertEqual(payload["operation"], "mutate.write_cells.plan")
        self.assertEqual(payload["data"]["targets"], ["B2"])

    @unittest.skipUnless("excel_com" in RecalcService().available_backends(), "Excel COM backend is not available.")
    def test_recalc_and_proof_with_excel_com(self) -> None:
        modified_path = Path(self.temp_dir.name) / "modified_recalc.xlsx"
        self.mutate.apply_write(
            str(self.base_path),
            "Model",
            {"B2": 5},
            output_path=str(modified_path),
            overwrite=True,
            recalc_backend="excel_com",
        )

        diff = self.verify.diff(str(self.base_path), str(modified_path))
        self.assertTrue(any(entry.sheet == "Model" and entry.ref == "C2" and entry.classification == "recalc_impact" for entry in diff.entries))
        self.assertTrue(any(entry.sheet == "Summary" and entry.ref == "B2" and entry.classification == "recalc_impact" for entry in diff.entries))

        proof = self.verify.proof(str(self.base_path), str(modified_path), ["Model!C2", "Summary!B2"])
        self.assertEqual(proof.changed_cell_count, 2)
        self.assertTrue(any(target.target == "Summary!B2" and target.cells[0].after_value == 120 for target in proof.targets))


if __name__ == "__main__":
    unittest.main()
