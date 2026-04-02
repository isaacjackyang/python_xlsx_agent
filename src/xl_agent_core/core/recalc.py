from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from xl_agent_core.core.contracts import RecalcResult, WarningMessage


class RecalcService:
    """Pluggable recalculation backends for Phase 2 workflows."""

    _excel_com_available: bool | None = None

    def available_backends(self) -> list[str]:
        backends = ["openpyxl"]
        if self._has_excel_com():
            backends.insert(0, "excel_com")
        return backends

    def recalculate(self, path: str, backend: str = "auto") -> RecalcResult:
        workbook_path = Path(path).expanduser().resolve()
        if not workbook_path.exists():
            raise FileNotFoundError(f"Workbook was not found: {workbook_path}")

        selected_backend = self._select_backend(backend)
        if selected_backend == "excel_com":
            return self._recalc_with_excel_com(workbook_path, backend)
        if selected_backend == "openpyxl":
            return self._recalc_with_openpyxl(workbook_path, backend)
        raise ValueError(f"Unsupported recalculation backend '{backend}'.")

    def _select_backend(self, requested_backend: str) -> str:
        normalized = requested_backend.lower()
        if normalized == "auto":
            return "excel_com" if self._has_excel_com() else "openpyxl"
        if normalized in {"excel_com", "openpyxl"}:
            if normalized == "excel_com" and not self._has_excel_com():
                raise RuntimeError("Excel COM is not available on this machine.")
            return normalized
        raise ValueError(f"Unsupported recalculation backend '{requested_backend}'.")

    def _has_excel_com(self) -> bool:
        if self.__class__._excel_com_available is not None:
            return self.__class__._excel_com_available
        try:
            import win32com.client  # type: ignore

            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Quit()
            self.__class__._excel_com_available = True
        except Exception:
            self.__class__._excel_com_available = False
        return self.__class__._excel_com_available

    def _recalc_with_openpyxl(self, workbook_path: Path, backend_requested: str) -> RecalcResult:
        workbook = load_workbook(workbook_path)
        workbook.calculation.calcMode = "auto"
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcOnSave = True
        workbook.save(workbook_path)

        return RecalcResult(
            workbook_path=str(workbook_path),
            backend_requested=backend_requested,
            backend_used="openpyxl",
            recalculated=False,
            calculation_mode=getattr(workbook.calculation, "calcMode", None),
            full_rebuild=True,
            warnings=[
                WarningMessage(
                    code="external_recalc_required",
                    message=(
                        "openpyxl marked the workbook for recalculation, but it did not recompute formula results. "
                        "Open the workbook in Excel or use the excel_com backend to refresh cached values."
                    ),
                ),
            ],
        )

    def _recalc_with_excel_com(self, workbook_path: Path, backend_requested: str) -> RecalcResult:
        try:
            import win32com.client  # type: ignore
        except Exception as exc:  # pragma: no cover - import guard
            raise RuntimeError("Excel COM backend is unavailable.") from exc

        excel = None
        workbook = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False
            workbook = excel.Workbooks.Open(str(workbook_path))
            excel.CalculateFullRebuild()
            workbook.Save()
            calculation_mode = str(getattr(workbook, "CalculationVersion", "")) or "excel_com"
            return RecalcResult(
                workbook_path=str(workbook_path),
                backend_requested=backend_requested,
                backend_used="excel_com",
                recalculated=True,
                calculation_mode=calculation_mode,
                full_rebuild=True,
                warnings=[],
            )
        finally:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
            if excel is not None:
                excel.Quit()
