from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

from xl_agent_core import __version__
from xl_agent_core.core.exporters import make_envelope, serialize
from xl_agent_core.core.formulas import FormulaService
from xl_agent_core.core.mutate import MutationService
from xl_agent_core.core.readers import ReadService
from xl_agent_core.core.recalc import RecalcService
from xl_agent_core.core.recon import WorkbookReconService
from xl_agent_core.core.verify import VerifyService

SERVER_NAME = "xl-agent-core"
SERVER_INSTRUCTIONS = (
    "Use progressive disclosure for spreadsheets. Probe first, detect regions before deep reads, "
    "treat warnings as real risk signals, and prefer copy-on-write mutation workflows followed by "
    "recalculation and proof/diff verification."
)


def _resolve_path(path: str) -> str:
    return str(Path(path).expanduser().resolve())


def _envelope(operation: str, data: Any, *, metadata: dict[str, Any] | None = None) -> dict[str, Any]:
    return serialize(make_envelope(operation, data, metadata=metadata or {}))


def create_mcp_server(*, transport: str = "stdio", streamable_http_path: str = "/mcp") -> FastMCP:
    recon = WorkbookReconService()
    reader = ReadService()
    formulas = FormulaService()
    mutate = MutationService()
    recalc = RecalcService()
    verify = VerifyService()

    mcp = FastMCP(
        SERVER_NAME,
        instructions=SERVER_INSTRUCTIONS,
        json_response=True,
        log_level="WARNING",
        stateless_http=transport == "streamable-http",
        streamable_http_path=streamable_http_path,
    )

    @mcp.tool()
    def server_info() -> dict[str, Any]:
        """Describe the server's purpose, transports, and supported spreadsheet workflows."""
        return {
            "name": SERVER_NAME,
            "version": __version__,
            "instructions": SERVER_INSTRUCTIONS,
            "supports": {
                "phase_1": [
                    "probe_workbook",
                    "list_sheets",
                    "read_names",
                    "detect_regions",
                    "read_table",
                    "read_layout",
                    "read_cells",
                    "formula_trace",
                ],
                "phase_2": [
                    "copy_workbook",
                    "plan_write_cells",
                    "write_cells",
                    "plan_replace_formulas",
                    "replace_formulas",
                    "recalculate_workbook",
                    "verify_proof",
                    "verify_diff",
                ],
            },
            "recalc_backends": recalc.available_backends(),
        }

    @mcp.tool()
    def probe_workbook(path: str) -> dict[str, Any]:
        """Inspect workbook structure before reading exact regions."""
        return _envelope("probe", recon.probe(path))

    @mcp.tool()
    def list_sheets(path: str) -> dict[str, Any]:
        """List workbook sheets and high-level sheet statistics."""
        return _envelope("read.sheets", recon.list_sheets(path), metadata={"workbook_path": _resolve_path(path)})

    @mcp.tool()
    def read_names(path: str) -> dict[str, Any]:
        """List workbook defined names and their destinations."""
        return _envelope("read.names", recon.read_names(path), metadata={"workbook_path": _resolve_path(path)})

    @mcp.tool()
    def detect_regions(path: str, sheet: str) -> dict[str, Any]:
        """Detect candidate data regions on a given sheet."""
        return _envelope(
            "detect.regions",
            recon.detect_regions(path, sheet),
            metadata={"workbook_path": _resolve_path(path), "sheet": sheet},
        )

    @mcp.tool()
    def read_table(path: str, sheet: str, region: str = "auto") -> dict[str, Any]:
        """Read an exact table region or auto-select the best detected region."""
        return _envelope("read.table", reader.read_table(path, sheet, region))

    @mcp.tool()
    def read_layout(path: str, sheet: str, range_ref: str) -> dict[str, Any]:
        """Read layout-aware cell details for a range."""
        return _envelope("read.layout", reader.read_layout(path, sheet, range_ref))

    @mcp.tool()
    def read_cells(path: str, sheet: str, refs: list[str]) -> dict[str, Any]:
        """Read exact cells or ranges expanded into per-cell inspections."""
        return _envelope(
            "inspect.cells",
            reader.read_cells(path, sheet, refs),
            metadata={"workbook_path": _resolve_path(path), "sheet": sheet, "requested_refs": refs},
        )

    @mcp.tool()
    def formula_trace(path: str, sheet: str, cell: str, direction: str = "precedents") -> dict[str, Any]:
        """Trace formula precedents or dependents for an exact cell."""
        return _envelope("analyze.formula_trace", formulas.formula_trace(path, sheet, cell, direction))

    @mcp.tool()
    def copy_workbook(source_path: str, output_path: str, overwrite: bool = False) -> dict[str, Any]:
        """Copy a workbook to start a non-destructive mutation workflow."""
        return _envelope("workbook.copy", mutate.copy_workbook(source_path, output_path, overwrite=overwrite))

    @mcp.tool()
    def plan_write_cells(
        path: str,
        sheet: str,
        edits: dict[str, Any],
        output_path: str,
        overwrite: bool = False,
    ) -> dict[str, Any]:
        """Dry-run a literal cell write plan without writing a workbook."""
        return _envelope(
            "mutate.write_cells.plan",
            mutate.plan_write(
                path,
                sheet,
                edits,
                output_path=output_path,
                overwrite=overwrite,
                operation="write_cells",
                edit_kind="value",
                dry_run=True,
            ),
        )

    @mcp.tool()
    def write_cells(
        path: str,
        sheet: str,
        edits: dict[str, Any],
        output_path: str,
        overwrite: bool = False,
        recalc_backend: str = "none",
    ) -> dict[str, Any]:
        """Write literal values to a copied workbook and optionally recalculate it."""
        return _envelope(
            "mutate.write_cells",
            mutate.apply_write(
                path,
                sheet,
                edits,
                output_path=output_path,
                overwrite=overwrite,
                recalc_backend=recalc_backend,
            ),
        )

    @mcp.tool()
    def plan_replace_formulas(
        path: str,
        sheet: str,
        edits: dict[str, str],
        output_path: str,
        overwrite: bool = False,
    ) -> dict[str, Any]:
        """Dry-run a formula replacement plan without writing a workbook."""
        return _envelope(
            "mutate.replace_formulas.plan",
            mutate.plan_write(
                path,
                sheet,
                edits,
                output_path=output_path,
                overwrite=overwrite,
                operation="replace_formulas",
                edit_kind="formula",
                dry_run=True,
            ),
        )

    @mcp.tool()
    def replace_formulas(
        path: str,
        sheet: str,
        edits: dict[str, str],
        output_path: str,
        overwrite: bool = False,
        recalc_backend: str = "none",
    ) -> dict[str, Any]:
        """Replace formulas in a copied workbook and optionally recalculate it."""
        return _envelope(
            "mutate.replace_formulas",
            mutate.replace_formulas(
                path,
                sheet,
                edits,
                output_path=output_path,
                overwrite=overwrite,
                recalc_backend=recalc_backend,
            ),
        )

    @mcp.tool()
    def recalculate_workbook(path: str, backend: str = "auto") -> dict[str, Any]:
        """Run a recalculation backend for an existing workbook."""
        return _envelope("workbook.recalc", recalc.recalculate(path, backend=backend))

    @mcp.tool()
    def verify_proof(baseline: str, current: str, targets: list[str]) -> dict[str, Any]:
        """Compare explicit proof targets between baseline and current workbooks."""
        return _envelope("verify.proof", verify.proof(baseline, current, targets))

    @mcp.tool()
    def verify_diff(original: str, modified: str) -> dict[str, Any]:
        """Compare workbook-wide changes between an original and modified workbook."""
        return _envelope("verify.diff", verify.diff(original, modified))

    return mcp


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="xl-mcp", description="Run the xl-agent-core MCP server.")
    parser.add_argument("--transport", choices=("stdio", "streamable-http"), default="stdio")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--path", default="/mcp", help="Streamable HTTP mount path.")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    mcp = create_mcp_server(transport=args.transport, streamable_http_path=args.path)
    if hasattr(mcp, "settings"):
        if hasattr(mcp.settings, "host"):
            mcp.settings.host = args.host
        if hasattr(mcp.settings, "port"):
            mcp.settings.port = args.port
        if hasattr(mcp.settings, "streamable_http_path"):
            mcp.settings.streamable_http_path = args.path

    mcp.run(transport=args.transport)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
