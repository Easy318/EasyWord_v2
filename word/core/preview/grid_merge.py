"""表格合并单元格：锚点 / 覆盖位判定（与 QHub、客户端语义对齐；禁止 import QHub）。"""

from __future__ import annotations

from typing import Any, Iterable, Mapping, Sequence


def _as_merge(m: Any) -> tuple[int, int, int, int] | None:
    if m is None:
        return None
    if isinstance(m, Mapping):
        row = m.get("row")
        col = m.get("col")
        row_span = m.get("rowSpan", m.get("row_span", 1))
        col_span = m.get("colSpan", m.get("col_span", 1))
    else:
        row = getattr(m, "row", None)
        col = getattr(m, "col", None)
        row_span = getattr(m, "row_span", None)
        if row_span is None:
            row_span = getattr(m, "rowSpan", 1)
        col_span = getattr(m, "col_span", None)
        if col_span is None:
            col_span = getattr(m, "colSpan", 1)
    try:
        r, c = int(row), int(col)  # type: ignore[arg-type]
        rs, cs = int(row_span or 1), int(col_span or 1)
    except (TypeError, ValueError):
        return None
    if r < 0 or c < 0 or rs < 1 or cs < 1:
        return None
    return r, c, rs, cs


def iter_merges(merges: Sequence[Any] | None) -> Iterable[tuple[int, int, int, int]]:
    for m in merges or []:
        parsed = _as_merge(m)
        if parsed is not None:
            yield parsed


def anchor_of(
    row: int, col: int, merges: Sequence[Any] | None
) -> tuple[int, int] | None:
    for r, c, rs, cs in iter_merges(merges):
        if r <= row < r + rs and c <= col < c + cs:
            return r, c
    return None


def is_covered(row: int, col: int, merges: Sequence[Any] | None) -> bool:
    anchor = anchor_of(row, col, merges)
    return anchor is not None and anchor != (row, col)


def is_anchor_or_plain(row: int, col: int, merges: Sequence[Any] | None) -> bool:
    return not is_covered(row, col, merges)


def col_span_at(row: int, col: int, merges: Sequence[Any] | None) -> int:
    """逻辑格 (row,col) 作为锚点时的列跨度；独立格为 1。"""
    for r, c, _rs, cs in iter_merges(merges):
        if r == row and c == col:
            return cs
    return 1


def iter_physical_anchors(
    row_count: int,
    col_count: int,
    merges: Sequence[Any] | None,
) -> Iterable[tuple[int, int]]:
    """阅读序物理格锚点（0-based），与 Word Table.Range.Cells 顺序对齐。"""
    for r in range(row_count):
        c = 0
        while c < col_count:
            if is_covered(r, c, merges):
                c += 1
                continue
            yield r, c
            c += col_span_at(r, c, merges)


def physical_cell_index(
    row: int,
    col: int,
    row_count: int,
    col_count: int,
    merges: Sequence[Any] | None,
) -> int:
    """逻辑锚点 → Range.Cells 的 1-based 序号。"""
    if is_covered(row, col, merges):
        raise ValueError(f"({row},{col}) 为合并覆盖位")
    idx = 0
    for r, c in iter_physical_anchors(row_count, col_count, merges):
        idx += 1
        if r == row and c == col:
            return idx
    raise ValueError(f"({row},{col}) 无法映射到物理单元格")
