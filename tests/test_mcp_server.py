from __future__ import annotations

import asyncio
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from openpyxl import Workbook


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
    summary = wb.create_sheet("Summary")
    summary["A1"] = "Metric"
    summary["B1"] = "Value"
    summary["A2"] = "Grand Total"
    summary["B2"] = "=Model!C2"
    wb.save(path)


class MCPServerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.temp_dir = tempfile.TemporaryDirectory()
        cls.workbook_path = Path(cls.temp_dir.name) / "mcp_sample.xlsx"
        build_sample_workbook(cls.workbook_path)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.temp_dir.cleanup()

    def test_stdio_server_lists_tools_and_executes_workflow(self) -> None:
        asyncio.run(self._run_stdio_flow())

    async def _run_stdio_flow(self) -> None:
        server_params = StdioServerParameters(
            command=sys.executable,
            args=["-m", "xl_agent_core.mcp.server", "--transport", "stdio"],
            cwd=ROOT,
        )

        async with stdio_client(server_params) as (read_stream, write_stream):
            async with ClientSession(read_stream, write_stream) as session:
                await session.initialize()

                tools_response = await session.list_tools()
                tool_names = {tool.name for tool in tools_response.tools}
                self.assertIn("server_info", tool_names)
                self.assertIn("probe_workbook", tool_names)
                self.assertIn("write_cells", tool_names)
                self.assertIn("verify_diff", tool_names)

                info_result = await session.call_tool("server_info")
                info = info_result.structuredContent
                self.assertEqual(info["name"], "xl-agent-core")
                self.assertIn("probe_workbook", info["supports"]["phase_1"])

                probe_result = await session.call_tool(
                    "probe_workbook",
                    arguments={"path": str(self.workbook_path)},
                )
                probe = probe_result.structuredContent
                self.assertEqual(probe["operation"], "probe")
                self.assertEqual(probe["status"], "ok")
                self.assertEqual(probe["data"]["sheet_count"], 2)

                read_result = await session.call_tool(
                    "read_table",
                    arguments={"path": str(self.workbook_path), "sheet": "Model", "region": "auto"},
                )
                read_table = read_result.structuredContent
                self.assertEqual(read_table["operation"], "read.table")
                self.assertEqual(read_table["data"]["header"], ["Item", "Qty", "Total"])

                modified_path = Path(self.temp_dir.name) / "mcp_modified.xlsx"
                write_result = await session.call_tool(
                    "write_cells",
                    arguments={
                        "path": str(self.workbook_path),
                        "sheet": "Model",
                        "edits": {"B2": 5},
                        "output_path": str(modified_path),
                        "overwrite": True,
                        "recalc_backend": "none",
                    },
                )
                write_payload = write_result.structuredContent
                self.assertEqual(write_payload["operation"], "mutate.write_cells")
                self.assertEqual(write_payload["data"]["edit_count"], 1)
                self.assertTrue(modified_path.exists())

                diff_result = await session.call_tool(
                    "verify_diff",
                    arguments={"original": str(self.workbook_path), "modified": str(modified_path)},
                )
                diff_payload = diff_result.structuredContent
                self.assertEqual(diff_payload["operation"], "verify.diff")
                self.assertTrue(any(entry["ref"] == "B2" for entry in diff_payload["data"]["entries"]))


if __name__ == "__main__":
    unittest.main()
