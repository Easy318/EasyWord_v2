"""从 Word COM 提取控件模板种子（text/chart/table）。"""

from __future__ import annotations

import re
from typing import Any
from xml.etree import ElementTree as ET

from word.core.content_control.chart_workbook import (
    map_chart_com_error,
    chart_sheet_session,
)
from word.core.content_control.constants import ControlType
from word.core.content_control.selection_guard import find_control_by_id
from word.core.preview.number_format import clean_float_str
from word.runtime.document_binding import document_binding
from word.runtime.errors import EasyWordError

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _normalize_cell_text(text: str) -> str:
    return str(text or "").replace("\r\x07", "").replace("\a", "").replace("\r", "").strip()


def _normalize_cc_text(text: str) -> str:
    return str(text or "").replace("\r\n", "\n").replace("\r", "\n").replace("\a", "").rstrip("\n")


def _cell_value_to_str(value: Any) -> str:
    return clean_float_str(value)


def _merges_from_rects(
    rects: list[tuple[int, int, int, int]],
) -> list[dict[str, int]]:
    """rects 为 0-based (row, col, rowSpan, colSpan)。"""
    merges: list[dict[str, int]] = []
    for row, col, row_span, col_span in rects:
        if row_span <= 1 and col_span <= 1:
            continue
        merges.append(
            {
                "row": row,
                "col": col,
                "rowSpan": row_span,
                "colSpan": col_span,
            }
        )
    merges.sort(key=lambda m: (m["row"], m["col"]))
    return merges


def _infer_vertical_rowspans(
    rows: list[list[str]],
    merges: list[dict[str, int]],
) -> list[dict[str, int]]:
    """Word 有时竖合不写 vMerge（下行是空独立 tc）。把整带空行收进 rowSpan。"""
    if not rows:
        return merges
    row_count = len(rows)
    col_count = max((len(r) for r in rows), default=0)
    if col_count <= 0:
        return merges

    # 现有合并覆盖（含锚点）
    def build_cover(ms: list[dict[str, int]]) -> dict[tuple[int, int], tuple[int, int]]:
        cover: dict[tuple[int, int], tuple[int, int]] = {}
        for m in ms:
            r0, c0 = m["row"], m["col"]
            for dr in range(m["rowSpan"]):
                for dc in range(m["colSpan"]):
                    cover[(r0 + dr, c0 + dc)] = (r0, c0)
        return cover

    by_anchor: dict[tuple[int, int], dict[str, int]] = {
        (m["row"], m["col"]): dict(m) for m in merges
    }

    for r, row in enumerate(rows):
        for c, raw in enumerate(row):
            text = str(raw or "").strip()
            if not text:
                continue
            cover = build_cover(list(by_anchor.values()))
            owner = cover.get((r, c))
            if owner is not None and owner != (r, c):
                continue
            m = by_anchor.get((r, c))
            if m is None:
                m = {"row": r, "col": c, "rowSpan": 1, "colSpan": 1}
            rs, cs = m["rowSpan"], m["colSpan"]
            while r + rs < row_count:
                ok = True
                cover = build_cover(list(by_anchor.values()))
                for dc in range(cs):
                    cc = c + dc
                    if cc >= col_count:
                        ok = False
                        break
                    rr = r + rs
                    if str(rows[rr][cc] if cc < len(rows[rr]) else "").strip():
                        ok = False
                        break
                    other = cover.get((rr, cc))
                    if other is not None and other != (r, c):
                        ok = False
                        break
                if not ok:
                    break
                rs += 1
            m["rowSpan"] = rs
            m["colSpan"] = cs
            if rs > 1 or cs > 1:
                by_anchor[(r, c)] = m
            elif (r, c) in by_anchor and rs <= 1 and cs <= 1:
                by_anchor.pop((r, c), None)

    out = [m for m in by_anchor.values() if m["rowSpan"] > 1 or m["colSpan"] > 1]
    out.sort(key=lambda m: (m["row"], m["col"]))
    return out


def _finalize_grid(
    rows: list[list[str]], merges: list[dict[str, int]]
) -> tuple[list[list[str]], list[dict[str, int]]]:
    return rows, _infer_vertical_rowspans(rows, merges)

def _local(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[-1]
    if ":" in tag:
        return tag.split(":", 1)[-1]
    return tag


def _attr(el: ET.Element, name: str) -> str | None:
    for key, val in el.attrib.items():
        if _local(key) == name:
            return val
    return None


def _find_child(parent: ET.Element, name: str) -> ET.Element | None:
    for child in parent:
        if _local(child.tag) == name:
            return child
    return None


def _tc_text(tc: ET.Element) -> str:
    parts: list[str] = []
    for node in tc.iter():
        if _local(node.tag) != "t":
            continue
        if node.text:
            parts.append(node.text)
        if node.tail:
            parts.append(node.tail)
    return _normalize_cell_text("".join(parts))


def _tc_grid_span(tc: ET.Element) -> int:
    tc_pr = _find_child(tc, "tcPr")
    if tc_pr is None:
        return 1
    gs = _find_child(tc_pr, "gridSpan")
    if gs is None:
        return 1
    try:
        return max(1, int(_attr(gs, "val") or "1"))
    except ValueError:
        return 1


def _tc_vmerge(tc: ET.Element) -> str | None:
    """返回 'restart' | 'continue' | None。"""
    tc_pr = _find_child(tc, "tcPr")
    if tc_pr is None:
        return None
    vm = _find_child(tc_pr, "vMerge")
    if vm is None:
        return None
    val = _attr(vm, "val")
    if val is None or val == "continue":
        return "continue"
    return "restart"


def _find_tbl_element(root: ET.Element) -> ET.Element | None:
    if _local(root.tag) == "tbl":
        return root
    for node in root.iter():
        if _local(node.tag) == "tbl":
            return node
    return None


def _child_by_local(parent: ET.Element, name: str) -> list[ET.Element]:
    return [c for c in list(parent) if _local(c.tag) == name]


def _grid_from_ooxml_tbl(
    tbl: ET.Element,
) -> tuple[list[list[str]], list[dict[str, int]]] | None:
    """从 w:tbl 解析逻辑网格与 merges（gridSpan + vMerge）。"""
    trs = _child_by_local(tbl, "tr")
    if not trs:
        return None

    row_specs: list[list[tuple[ET.Element, int, str | None]]] = []
    col_count = 0
    for tr in trs:
        specs: list[tuple[ET.Element, int, str | None]] = []
        logical = 0
        for tc in _child_by_local(tr, "tc"):
            span = _tc_grid_span(tc)
            specs.append((tc, span, _tc_vmerge(tc)))
            logical += span
        col_count = max(col_count, logical)
        row_specs.append(specs)

    if col_count <= 0:
        return None

    row_count = len(row_specs)
    rows: list[list[str]] = [["" for _ in range(col_count)] for _ in range(row_count)]
    owner: list[list[tuple[int, int] | None]] = [
        [None for _ in range(col_count)] for _ in range(row_count)
    ]
    anchors: dict[tuple[int, int], dict[str, int]] = {}

    for r, specs in enumerate(row_specs):
        c = 0
        for tc, col_span, vmerge in specs:
            while c < col_count and owner[r][c] is not None:
                c += 1
            if c >= col_count:
                break

            if vmerge == "continue":
                up = owner[r - 1][c] if r > 0 else None
                if up is not None:
                    ar, ac = up
                    for dr in range(r - ar + 1):
                        for dc in range(col_span):
                            rr, cc = ar + dr, ac + dc
                            if rr < row_count and cc < col_count:
                                owner[rr][cc] = (ar, ac)
                    info = anchors[(ar, ac)]
                    info["rowSpan"] = max(info["rowSpan"], r - ar + 1)
                    info["colSpan"] = max(info["colSpan"], col_span)
                c += col_span
                continue

            text = _tc_text(tc)
            rows[r][c] = text
            anchors[(r, c)] = {"row": r, "col": c, "rowSpan": 1, "colSpan": col_span}
            for dc in range(col_span):
                if c + dc < col_count:
                    owner[r][c + dc] = (r, c)
            c += col_span

    rects = [
        (a["row"], a["col"], a["rowSpan"], a["colSpan"]) for a in anchors.values()
    ]
    return _finalize_grid(rows, _merges_from_rects(rects))


def _table_xml_text(table: Any) -> str | None:
    """尽量取出表格 OOXML 片段。"""
    for attr in ("XML", "WordOpenXML"):
        try:
            raw = getattr(table.Range, attr)
            if raw and str(raw).strip():
                return str(raw)
        except Exception:  # noqa: BLE001
            continue
    return None


def _parse_table_ooxml(
    xml_text: str,
) -> tuple[list[list[str]], list[dict[str, int]]] | None:
    text = xml_text.strip()
    if not text:
        return None
    root: ET.Element | None = None
    try:
        root = ET.fromstring(text)
    except ET.ParseError:
        try:
            root = ET.fromstring(f'<root xmlns:w="{_W_NS}">{text}</root>')
        except ET.ParseError:
            root = None
    tbl = _find_tbl_element(root) if root is not None else None
    if tbl is None:
        m = re.search(r"<w:tbl\b.*</w:tbl>", text, re.DOTALL)
        if not m:
            return None
        frag = m.group(0)
        try:
            tbl = ET.fromstring(frag)
        except ET.ParseError:
            try:
                tbl = ET.fromstring(f'<wrapper xmlns:w="{_W_NS}">{frag}</wrapper>')
                tbl = _find_tbl_element(tbl)
            except ET.ParseError:
                return None
    if tbl is None:
        return None
    return _grid_from_ooxml_tbl(tbl)


def _colspan_from_widths(
    cell_width: float, start_col: int, widths: list[float]
) -> int:
    if start_col >= len(widths):
        return 1
    tol = max(1.5, 0.02 * max(cell_width, 1.0))
    acc = 0.0
    span = 0
    for c in range(start_col, len(widths)):
        acc += widths[c]
        span += 1
        if abs(acc - cell_width) <= tol or acc >= cell_width - tol:
            break
    return max(1, span)


def _read_table_grid_by_width(
    table: Any, row_count: int, col_count: int
) -> tuple[list[list[str]], list[dict[str, int]]]:
    """回退：行内物理格 + 列宽推 colspan；occupancy 推 rowspan。"""
    widths: list[float] = []
    for c in range(1, col_count + 1):
        try:
            widths.append(float(table.Columns.Item(c).Width))
        except Exception:  # noqa: BLE001
            widths.append(0.0)

    rows: list[list[str]] = [["" for _ in range(col_count)] for _ in range(row_count)]
    occupied = [[False] * col_count for _ in range(row_count)]
    rects: list[tuple[int, int, int, int]] = []

    for r in range(1, row_count + 1):
        try:
            row_cells = table.Rows.Item(r).Cells
            n = int(row_cells.Count)
        except Exception:  # noqa: BLE001
            continue
        logical_c = 0
        for i in range(1, n + 1):
            while logical_c < col_count and occupied[r - 1][logical_c]:
                logical_c += 1
            if logical_c >= col_count:
                break
            try:
                cell = row_cells.Item(i)
                text = _normalize_cell_text(str(cell.Range.Text or ""))
                cw = float(cell.Width)
            except Exception:  # noqa: BLE001
                continue
            col_span = min(
                _colspan_from_widths(cw, logical_c, widths), col_count - logical_c
            )

            row_span = 1
            for next_0 in range(r, row_count):
                if any(occupied[next_0][logical_c + dc] for dc in range(col_span)):
                    break
                try:
                    next_cells = table.Rows.Item(next_0 + 1).Cells
                    nn = int(next_cells.Count)
                except Exception:  # noqa: BLE001
                    break
                trial = [row[:] for row in occupied]
                for dc in range(col_span):
                    trial[next_0][logical_c + dc] = True
                cursor = 0
                ok = True
                for j in range(1, nn + 1):
                    while cursor < col_count and trial[next_0][cursor]:
                        cursor += 1
                    if cursor >= col_count:
                        ok = False
                        break
                    try:
                        jw = float(next_cells.Item(j).Width)
                    except Exception:  # noqa: BLE001
                        ok = False
                        break
                    js = min(
                        _colspan_from_widths(jw, cursor, widths), col_count - cursor
                    )
                    cursor += js
                if not ok:
                    break
                row_span += 1
                for dc in range(col_span):
                    occupied[next_0][logical_c + dc] = True

            rows[r - 1][logical_c] = text
            rects.append((r - 1, logical_c, row_span, col_span))
            for dr in range(row_span):
                for dc in range(col_span):
                    occupied[r - 1 + dr][logical_c + dc] = True
            logical_c += col_span

    return _finalize_grid(rows, _merges_from_rects(rects))


def _read_table_grid(
    rng: Any,
) -> tuple[list[list[str]], list[dict[str, int]], str | None]:
    """读逻辑网格 + merges。优先 OOXML gridSpan/vMerge；失败再列宽回退。"""
    if int(rng.Tables.Count) == 0:
        return [], [], None
    table = rng.Tables.Item(1)
    row_count = int(table.Rows.Count)
    col_count = int(table.Columns.Count)

    xml_text = _table_xml_text(table)
    if xml_text:
        parsed = _parse_table_ooxml(xml_text)
        if parsed is not None:
            rows, merges = parsed
            if len(rows) != row_count or (rows and len(rows[0]) != col_count):
                norm = [["" for _ in range(col_count)] for _ in range(row_count)]
                for ri, row in enumerate(rows[:row_count]):
                    for ci, val in enumerate(row[:col_count]):
                        norm[ri][ci] = val
                rows = norm
                merges = [
                    m
                    for m in merges
                    if m["row"] < row_count
                    and m["col"] < col_count
                    and m["row"] + m["rowSpan"] <= row_count
                    and m["col"] + m["colSpan"] <= col_count
                ]
                rows, merges = _finalize_grid(rows, merges)
            return rows, merges, None

    rows, merges = _read_table_grid_by_width(table, row_count, col_count)
    return rows, merges, None

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
        rows, merges, _ = _read_table_grid(rng)
        snap: dict[str, object] = {"rows": rows, "sheetName": None}
        if merges:
            snap["merges"] = merges
        return {
            "gridSnapshot": snap,
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
