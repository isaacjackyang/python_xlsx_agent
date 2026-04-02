from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from xl_agent_core.core.contracts import WarningMessage
from xl_agent_core.core.exporters import make_envelope, render_json
from xl_agent_core.core.formulas import FormulaService
from xl_agent_core.core.mutate import MutationService
from xl_agent_core.core.readers import ReadService
from xl_agent_core.core.recalc import RecalcService
from xl_agent_core.core.recon import WorkbookReconService
from xl_agent_core.core.verify import VerifyService


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="xl", description="Spreadsheet agent CLI.")
    parser.add_argument("--pretty", action="store_true", help="Pretty-print JSON output.")

    subparsers = parser.add_subparsers(dest="command", required=True)

    probe_parser = subparsers.add_parser("probe", help="Inspect workbook structure.")
    probe_parser.add_argument("path")

    read_parser = subparsers.add_parser("read", help="Read workbook structures.")
    read_subparsers = read_parser.add_subparsers(dest="read_command", required=True)

    read_sheets = read_subparsers.add_parser("sheets", help="List workbook sheets.")
    read_sheets.add_argument("path")

    read_names = read_subparsers.add_parser("names", help="List workbook named items.")
    read_names.add_argument("path")

    read_table = read_subparsers.add_parser("table", help="Read a detected or explicit table region.")
    read_table.add_argument("path")
    read_table.add_argument("--sheet", required=True)
    read_table.add_argument("--region", default="auto")

    read_layout = read_subparsers.add_parser("layout", help="Read layout-aware cell information for a range.")
    read_layout.add_argument("path")
    read_layout.add_argument("--sheet", required=True)
    read_layout.add_argument("--range", dest="range_ref", required=True)

    detect_parser = subparsers.add_parser("detect", help="Detect workbook structures.")
    detect_subparsers = detect_parser.add_subparsers(dest="detect_command", required=True)
    detect_regions = detect_subparsers.add_parser("regions", help="Detect region candidates on a sheet.")
    detect_regions.add_argument("path")
    detect_regions.add_argument("--sheet", required=True)

    inspect_parser = subparsers.add_parser("inspect", help="Inspect workbook elements.")
    inspect_subparsers = inspect_parser.add_subparsers(dest="inspect_command", required=True)
    inspect_cells = inspect_subparsers.add_parser("cells", help="Inspect exact cells or ranges.")
    inspect_cells.add_argument("path")
    inspect_cells.add_argument("--sheet", required=True)
    inspect_cells.add_argument("--refs", nargs="+", required=True)

    analyze_parser = subparsers.add_parser("analyze", help="Analyze workbook relationships.")
    analyze_subparsers = analyze_parser.add_subparsers(dest="analyze_command", required=True)
    formula_trace = analyze_subparsers.add_parser("formula-trace", help="Trace formula precedents or dependents.")
    formula_trace.add_argument("path")
    formula_trace.add_argument("--sheet", required=True)
    formula_trace.add_argument("--cell", required=True)
    formula_trace.add_argument(
        "--direction",
        choices=("precedents", "dependents"),
        default="precedents",
    )

    workbook_parser = subparsers.add_parser("workbook", help="Workbook-level Phase 2 operations.")
    workbook_subparsers = workbook_parser.add_subparsers(dest="workbook_command", required=True)

    workbook_copy = workbook_subparsers.add_parser("copy", help="Copy a workbook for copy-on-write workflows.")
    workbook_copy.add_argument("source_path")
    workbook_copy.add_argument("output_path")
    workbook_copy.add_argument("--overwrite", action="store_true")

    workbook_recalc = workbook_subparsers.add_parser("recalc", help="Recalculate a workbook.")
    workbook_recalc.add_argument("path")
    workbook_recalc.add_argument("--backend", choices=("auto", "excel_com", "openpyxl"), default="auto")

    mutate_parser = subparsers.add_parser("mutate", help="Safe mutation workflows.")
    mutate_subparsers = mutate_parser.add_subparsers(dest="mutate_command", required=True)

    mutate_write = mutate_subparsers.add_parser("write-cells", help="Write literal cell values to a copied workbook.")
    mutate_write.add_argument("path")
    mutate_write.add_argument("--sheet", required=True)
    mutate_write.add_argument("--set", dest="assignments", action="append", required=True)
    mutate_write.add_argument("--output", required=True)
    mutate_write.add_argument("--overwrite", action="store_true")
    mutate_write.add_argument("--dry-run", action="store_true")
    mutate_write.add_argument("--recalc-backend", choices=("none", "auto", "excel_com", "openpyxl"), default="none")

    mutate_formulas = mutate_subparsers.add_parser("replace-formulas", help="Replace formulas in a copied workbook.")
    mutate_formulas.add_argument("path")
    mutate_formulas.add_argument("--sheet", required=True)
    mutate_formulas.add_argument("--set", dest="assignments", action="append", required=True)
    mutate_formulas.add_argument("--output", required=True)
    mutate_formulas.add_argument("--overwrite", action="store_true")
    mutate_formulas.add_argument("--dry-run", action="store_true")
    mutate_formulas.add_argument("--recalc-backend", choices=("none", "auto", "excel_com", "openpyxl"), default="none")

    verify_parser = subparsers.add_parser("verify", help="Proof and diff verification workflows.")
    verify_subparsers = verify_parser.add_subparsers(dest="verify_command", required=True)

    verify_proof = verify_subparsers.add_parser("proof", help="Compare explicit proof targets between workbooks.")
    verify_proof.add_argument("baseline")
    verify_proof.add_argument("current")
    verify_proof.add_argument("--targets", nargs="+", required=True)

    verify_diff = verify_subparsers.add_parser("diff", help="Compare workbook-wide cell/formula differences.")
    verify_diff.add_argument("original")
    verify_diff.add_argument("modified")

    return parser


def execute(args: argparse.Namespace):
    recon = WorkbookReconService()
    reader = ReadService()
    formulas = FormulaService()
    mutate = MutationService()
    recalc = RecalcService()
    verify = VerifyService()

    path = getattr(args, "path", None)
    resolved_path = str(Path(path).expanduser().resolve()) if path else None

    if args.command == "probe":
        return make_envelope("probe", recon.probe(args.path))

    if args.command == "read":
        if args.read_command == "sheets":
            return make_envelope(
                "read.sheets",
                recon.list_sheets(args.path),
                metadata={"workbook_path": resolved_path},
            )
        if args.read_command == "names":
            return make_envelope(
                "read.names",
                recon.read_names(args.path),
                metadata={"workbook_path": resolved_path},
            )
        if args.read_command == "table":
            return make_envelope("read.table", reader.read_table(args.path, args.sheet, args.region))
        if args.read_command == "layout":
            return make_envelope("read.layout", reader.read_layout(args.path, args.sheet, args.range_ref))

    if args.command == "detect" and args.detect_command == "regions":
        return make_envelope(
            "detect.regions",
            recon.detect_regions(args.path, args.sheet),
            metadata={"workbook_path": resolved_path, "sheet": args.sheet},
        )

    if args.command == "inspect" and args.inspect_command == "cells":
        return make_envelope(
            "inspect.cells",
            reader.read_cells(args.path, args.sheet, args.refs),
            metadata={"workbook_path": resolved_path, "sheet": args.sheet, "requested_refs": args.refs},
        )

    if args.command == "analyze" and args.analyze_command == "formula-trace":
        return make_envelope(
            "analyze.formula_trace",
            formulas.formula_trace(args.path, args.sheet, args.cell, args.direction),
        )

    if args.command == "workbook":
        if args.workbook_command == "copy":
            return make_envelope(
                "workbook.copy",
                mutate.copy_workbook(args.source_path, args.output_path, overwrite=args.overwrite),
            )
        if args.workbook_command == "recalc":
            return make_envelope(
                "workbook.recalc",
                recalc.recalculate(args.path, backend=args.backend),
            )

    if args.command == "mutate":
        if args.mutate_command == "write-cells":
            assignments = _parse_value_assignments(args.assignments)
            if args.dry_run:
                return make_envelope(
                    "mutate.write_cells.plan",
                    mutate.plan_write(
                        args.path,
                        args.sheet,
                        assignments,
                        output_path=args.output,
                        overwrite=args.overwrite,
                        operation="write_cells",
                        edit_kind="value",
                        dry_run=True,
                    ),
                )
            return make_envelope(
                "mutate.write_cells",
                mutate.apply_write(
                    args.path,
                    args.sheet,
                    assignments,
                    output_path=args.output,
                    overwrite=args.overwrite,
                    recalc_backend=args.recalc_backend,
                ),
            )
        if args.mutate_command == "replace-formulas":
            assignments = _parse_formula_assignments(args.assignments)
            if args.dry_run:
                return make_envelope(
                    "mutate.replace_formulas.plan",
                    mutate.plan_write(
                        args.path,
                        args.sheet,
                        assignments,
                        output_path=args.output,
                        overwrite=args.overwrite,
                        operation="replace_formulas",
                        edit_kind="formula",
                        dry_run=True,
                    ),
                )
            return make_envelope(
                "mutate.replace_formulas",
                mutate.replace_formulas(
                    args.path,
                    args.sheet,
                    assignments,
                    output_path=args.output,
                    overwrite=args.overwrite,
                    recalc_backend=args.recalc_backend,
                ),
            )

    if args.command == "verify":
        if args.verify_command == "proof":
            return make_envelope(
                "verify.proof",
                verify.proof(args.baseline, args.current, args.targets),
            )
        if args.verify_command == "diff":
            return make_envelope(
                "verify.diff",
                verify.diff(args.original, args.modified),
            )

    raise ValueError("Unsupported command.")


def _parse_value_assignments(assignments: list[str]) -> dict[str, object]:
    parsed: dict[str, object] = {}
    for assignment in assignments:
        ref, raw_value = _split_assignment(assignment)
        parsed[ref] = _parse_literal_value(raw_value)
    return parsed


def _parse_formula_assignments(assignments: list[str]) -> dict[str, str]:
    parsed: dict[str, str] = {}
    for assignment in assignments:
        ref, raw_value = _split_assignment(assignment)
        formula = raw_value.strip()
        parsed[ref] = formula if formula.startswith("=") else f"={formula}"
    return parsed


def _split_assignment(assignment: str) -> tuple[str, str]:
    if "=" not in assignment:
        raise ValueError(f"Assignment '{assignment}' must use REF=VALUE syntax.")
    ref, raw_value = assignment.split("=", 1)
    ref = ref.strip().replace("$", "").upper()
    if not ref:
        raise ValueError(f"Assignment '{assignment}' is missing a cell reference.")
    return ref, raw_value


def _parse_literal_value(raw_value: str) -> object:
    candidate = raw_value.strip()
    if candidate == "":
        return ""
    try:
        return json.loads(candidate)
    except json.JSONDecodeError:
        return candidate


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        envelope = execute(args)
        print(render_json(envelope, pretty=args.pretty))
        return 0
    except Exception as exc:  # pragma: no cover - CLI guard path
        error_envelope = make_envelope(
            operation="error",
            data={},
            status="error",
            warnings=[WarningMessage(code=type(exc).__name__, message=str(exc), severity="error")],
        )
        print(render_json(error_envelope, pretty=True), file=sys.stderr)
        return 1
