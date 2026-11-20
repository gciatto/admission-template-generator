"""Reduce a workbook to a comparable structure, for layout parity tests."""

from __future__ import annotations

from typing import Any

import openpyxl


def _color(c: Any) -> str | None:
    if c is None:
        return None
    if c.type == "rgb":
        return c.rgb
    return f"{c.type}:{c.value}:{c.tint}"


def _style(cell: Any) -> tuple:
    f, fill, a, b = cell.font, cell.fill, cell.alignment, cell.border
    return (
        (f.name, f.sz, bool(f.b), bool(f.i), _color(f.color)),
        (fill.fill_type, _color(fill.fgColor) if fill.fill_type else None),
        cell.number_format,
        (bool(a.wrap_text), a.horizontal, a.vertical),
        tuple(
            (s.style, _color(s.color)) if s else (None, None)
            for s in (b.left, b.right, b.top, b.bottom)
        ),
        (bool(cell.protection.locked), bool(cell.protection.hidden)),
    )


def snapshot(path: Any) -> dict[str, Any]:
    wb = openpyxl.load_workbook(path)
    out: dict[str, Any] = {
        "sheets": wb.sheetnames,
        "names": {k: v.attr_text for k, v in wb.defined_names.items()},
        "lock_structure": bool(wb.security and wb.security.lockStructure),
        "calc": (
            wb.calculation.calcMode,
            wb.calculation.fullCalcOnLoad,
            wb.calculation.forceFullCalc,
        ),
        "default_font": (wb._fonts[0].name, wb._fonts[0].sz),  # type: ignore[attr-defined]
    }
    for ws in wb.worksheets:
        view, prot = ws.sheet_view, ws.protection
        pane = view.pane
        sheet: dict[str, Any] = {
            "state": ws.sheet_state,
            "tab": _color(ws.sheet_properties.tabColor),
            "view": (
                view.showGridLines,
                view.zoomScale,
                pane and (pane.xSplit, pane.ySplit, pane.topLeftCell, pane.activePane),
            ),
            "protection": tuple(
                bool(getattr(prot, k))
                for k in [
                    "sheet",
                    "objects",
                    "scenarios",
                    "formatCells",
                    "formatColumns",
                    "formatRows",
                    "insertColumns",
                    "insertRows",
                    "deleteColumns",
                    "deleteRows",
                    "sort",
                    "autoFilter",
                    "selectLockedCells",
                    "selectUnlockedCells",
                ]
            ),
            "merges": sorted(str(r) for r in ws.merged_cells.ranges),
            "widths": {k: d.width for k, d in ws.column_dimensions.items() if d.customWidth},
            "heights": {k: d.height for k, d in ws.row_dimensions.items() if d.height},
            "tables": sorted(
                (
                    t.name,
                    t.ref,
                    t.tableStyleInfo.name,
                    bool(t.tableStyleInfo.showRowStripes),
                    t.autoFilter is not None,
                    tuple(c.name for c in t.tableColumns),
                )
                for t in ws.tables.values()
            ),
            "validations": sorted(
                (
                    str(dv.sqref),
                    dv.type,
                    dv.operator,
                    dv.formula1,
                    dv.formula2,
                    dv.prompt,
                    dv.promptTitle,
                    dv.error,
                    dv.errorTitle,
                    bool(dv.allow_blank),
                    (dv.showInputMessage, dv.showErrorMessage, dv.errorStyle),
                )
                for dv in ws.data_validations.dataValidation
            ),
            "conditional": sorted(
                (
                    str(cf.sqref),
                    r.type,
                    tuple(r.formula or ()),
                    r.dxf and r.dxf.fill and (r.dxf.fill.fill_type, _color(r.dxf.fill.bgColor)),
                    r.dxf and r.dxf.font and (_color(r.dxf.font.color), bool(r.dxf.font.b)),
                )
                for cf in ws.conditional_formatting
                for r in cf.rules
            ),
            "cells": {
                c.coordinate: (c.value, _style(c))
                for row in ws.iter_rows()
                for c in row
                if c.value is not None or c.has_style
            },
        }
        out[ws.title] = sheet
    return out


def diff(a: dict[str, Any], b: dict[str, Any], limit: int = 40) -> list[str]:
    """Human-readable differences between two snapshots (empty when equal)."""
    found: list[str] = []

    def walk(x: Any, y: Any, where: str) -> None:
        if len(found) >= limit:
            return
        if isinstance(x, dict) and isinstance(y, dict):
            for k in sorted(set(x) | set(y), key=str):
                if k not in x or k not in y:
                    found.append(f"{where}/{k}: only in {'expected' if k in x else 'actual'}")
                else:
                    walk(x[k], y[k], f"{where}/{k}")
        elif x != y:
            found.append(f"{where}: expected {x!r}, got {y!r}")

    walk(a, b, "")
    return found
