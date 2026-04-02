from __future__ import annotations

from xl_agent_core.core.contracts import NamedItem
from xl_agent_core.core.loader import load_workbook_bundle


def read_named_items(path: str) -> list[NamedItem]:
    bundle = load_workbook_bundle(path)
    items: list[NamedItem] = []

    for name, defined_name in bundle.formulas.defined_names.items():
        destinations: list[str] = []
        try:
            for sheet, ref in defined_name.destinations:
                destinations.append(f"{sheet}!{ref}")
        except Exception:
            destinations = []

        scope = "workbook"
        if getattr(defined_name, "localSheetId", None) is not None:
            scope_index = defined_name.localSheetId
            if 0 <= scope_index < len(bundle.formulas.sheetnames):
                scope = bundle.formulas.sheetnames[scope_index]

        items.append(
            NamedItem(
                name=name,
                scope=scope,
                refers_to=defined_name.attr_text or "",
                destinations=destinations,
                is_formula_like=not destinations,
            ),
        )

    items.sort(key=lambda item: (item.scope, item.name.lower()))
    return items
