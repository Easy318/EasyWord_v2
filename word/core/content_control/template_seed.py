"""从 Word COM 提取控件模板种子（text/chart/table）。"""

from __future__ import annotations

from typing import Any

from word.core.content_control.chart_workbook import (
    map_chart_com_error,
    chart_sheet_session,
)
from word.core.content_control.constants import ControlType
from word.core.content_control.selection_guard import find_control_by_id
from word.core.preview.number_format import clean_float_str
from word.runtime.document_binding import document_binding
from word.runtime.errors import EasyWordError


def _normalize_cell_text(text: str) -> str:
    return str(text or "").replace("\r\x07", "").replace("\a", "").replace("\r", "").strip()


def _normalize_cc_text(text: str) -> str:
    return str(text or "").replace("\r\n", "\n").replace("\r", "\n").replace("\a", "").rstrip("\n")


def _cell_value_to_str(value: Any) -> str:
    return clean_float_str(value)


def _read_table_grid(rng: Any) -> tuple[list[list[str]], str | None]:
    if int(rng.Tables.Count) == 0:
        return [], None
    table = rng.Tables.Item(1)
    rows: list[list[str]] = []
    row_count = int(table.Rows.Count)
    col_count = int(table.Columns.Count)
    for r in range(1, row_count + 1):
        row: list[str] = []
        for c in range(1, col_count + 1):
            cell_text = _normalize_cell_text(str(table.Cell(r, c).Range.Text or ""))
            row.append(cell_text)
        rows.append(row)
    return rows, None


def _find_chart_in_range(rng: Any) -> Any | None:
    try:
        inlines = rng.InlineShapes
        for i in range(1, int(inlines.Count) + 1):
            shape = inlines.Item(i)
            try:
                return shape.Chart
            except Exception:  # noqa: BLE001
                if int(shape.Type) == 12:
                    return shape.Chart
    except Exception:  # noqa: BLE001
        pass
    try:
        doc = rng.Document
        for i in range(1, int(doc.Shapes.Count) + 1):
            sh = doc.Shapes.Item(i)
            try:
                anchor_start = int(sh.Anchor.Start)
                if int(rng.Start) <= anchor_start < int(rng.End):
                    return sh.Chart
            except Exception:  # noqa: BLE001
                continue
    except Exception:  # noqa: BLE001
        pass
    return None


def _read_chart_grid(rng: Any) -> tuple[list[list[str]], str | None]:
    chart = _find_chart_in_range(rng)
    if chart is None:
        raise EasyWordError("no_chart", "控件内未找到图表，无法读取数据网格")

    try:
        with chart_sheet_session(chart, sync_visual=False) as (_chart_data, _wb, ws, used):
            row_count = int(used.Rows.Count)
            col_count = int(used.Columns.Count)
            rows: list[list[str]] = []
            for r in range(1, row_count + 1):
                row: list[str] = []
                for c in range(1, col_count + 1):
                    val = used.Cells(r, c).Value
                    row.append(_cell_value_to_str(val))
                rows.append(row)
            sheet_name = str(ws.Name or "")
            return rows, sheet_name or None
    except EasyWordError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_chart_com_error(exc, action="读取图表数据") from exc


def extract_template_seed(info: Any) -> dict[str, object]:
    """根据 CcInfo 提取模板种子字段。"""
    cc = info.com
    rng = cc.Range
    bt = info.business_type
    if bt is None:
        raise ValueError("非业务内容控件")

    if bt == ControlType.text:
        return {"jinjaTemplate": _normalize_cc_text(str(rng.Text or ""))}

    if bt == ControlType.table:
        rows, _ = _read_table_grid(rng)
        return {
            "gridSnapshot": {"rows": rows, "sheetName": None},
            "cellBindings": [],
        }

    if bt == ControlType.chart:
        rows, sheet_name = _read_chart_grid(rng)
        snapshot: dict[str, object] = {"rows": rows}
        if sheet_name:
            snapshot["sheetName"] = sheet_name
        return {
            "gridSnapshot": snapshot,
            "cellBindings": [],
        }

    if bt == ControlType.picture:
        return {"placeholder": True}

    raise ValueError(f"不支持的控件类型: {bt.value}")


def get_template_seed(control_id: int) -> dict[str, object]:
    doc = document_binding.get_bound_document()
    info = find_control_by_id(doc, control_id)
    seed = extract_template_seed(info)
    seed["controlId"] = int(info.id)
    seed["type"] = info.business_type.value if info.business_type else None
    seed["controlTag"] = str(info.tag or "")
    return seed
