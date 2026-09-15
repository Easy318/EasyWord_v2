"""Word COM 预览写回（text Jinja + chart/table cell 值）。"""

from __future__ import annotations

import re
from typing import Any

from jinja2 import Environment, TemplateSyntaxError, UndefinedError

from word.core.content_control.chart_workbook import (
    chart_sheet_session,
    map_chart_com_error,
)
from word.core.content_control.constants import ControlType
from word.core.content_control.selection_guard import find_control_by_id
from word.core.content_control.template_seed import (
    _cell_value_to_str,
    _find_chart_in_range,
)
from word.core.preview.grid_merge import is_covered, physical_cell_index
from word.core.preview.number_format import clean_float_str, fmt_number
from word.core.preview.schema.common import JINJA_RESERVED_NAMES
from word.core.preview.schema.control_template import (
    ChartControlTemplate,
    ControlTemplate,
    PictureControlTemplate,
    TableControlTemplate,
    TextControlTemplate,
)
from word.core.preview.schema.preview import PreviewApplyIn, PreviewApplyOut, PreviewAppliedItem, PreviewErrorItem
from word.core.preview.schema.render_context import RenderContext, RenderValue
from word.runtime.document_binding import document_binding
from word.runtime.errors import EasyWordError


def _jinja_finalize(value: Any) -> Any:
    """写出前去掉 float 二进制尾巴；其它类型交 Jinja 默认 stringify。"""
    if isinstance(value, float) and not isinstance(value, bool):
        return clean_float_str(value)
    return value


def _jinja_fmt(value: Any, places: int = 2) -> str:
    return fmt_number(value, places)


_JINJA_ENV = Environment(
    autoescape=False,
    finalize=_jinja_finalize,
    trim_blocks=True,
    lstrip_blocks=True,
)
_JINJA_ENV.filters["fmt"] = _jinja_fmt


def _render_value(item: RenderValue) -> Any:
    if item.status != "active":
        return None
    return item.value


def _build_jinja_context(ctx: RenderContext) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for name, item in ctx.entries.items():
        out[name] = _render_value(item)
    if ctx.burst is not None:
        out["burst_var"] = ctx.burst.var_name
        out["burst_value"] = ctx.burst.value
        alias = ctx.burst.var_name
        if alias and alias not in JINJA_RESERVED_NAMES:
            out[alias] = ctx.burst.value
    return out


def _resolve_ref_value(ctx: RenderContext, ref_name: str) -> Any:
    item = ctx.entries.get(ref_name)
    if item is None:
        raise EasyWordError("ref_missing", f"引用变量 {ref_name} 不存在")
    if item.status == "tombstone":
        raise EasyWordError("ref_tombstone", f"引用变量 {ref_name} 已失效")
    if item.status == "missing":
        raise EasyWordError("ref_missing", f"引用变量 {ref_name} 缺失")
    return item.value


def _normalize_rendered_text(text: str) -> str:
    """统一换行、去掉行尾空白；trim_blocks 之外再压掉多余空行。"""
    t = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
    t = "\n".join(line.rstrip() for line in t.split("\n"))
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip("\n")


def _render_jinja(source: str, ctx: RenderContext) -> str:
    try:
        raw = _JINJA_ENV.from_string(source).render(**_build_jinja_context(ctx))
    except TemplateSyntaxError as exc:
        raise EasyWordError("jinja_syntax", f"模板语法错误: {exc}") from exc
    except UndefinedError as exc:
        raise EasyWordError("jinja_undefined", f"模板变量未定义: {exc}") from exc
    return _normalize_rendered_text(raw)


def _write_text_control(cc: Any, text: str) -> None:
    # 纯文本 SDT 只能单段：多行必须用软换行 Chr(11)=\v（对齐 OOXML <w:br/>）。
    # 用 \r 会插入非法多段，常导致数字旁缺字形方框，保存后才被 Word 纠正。
    normalized = _normalize_rendered_text(text).replace("\n", "\v")
    try:
        cc.MultiLine = True
    except Exception:
        pass

    rng = cc.Range
    prev_fe = ""
    prev_ascii = ""
    try:
        prev_fe = str(rng.Font.NameFarEast or "").strip()
        prev_ascii = str(rng.Font.NameAscii or "").strip()
    except Exception:
        pass

    rng.Text = normalized

    # Text 赋值会打乱 run 字体；按写前字体恢复，缺省用宋体/TNR
    try:
        rng = cc.Range
        font = rng.Font
        font.NameFarEast = prev_fe or "宋体"
        font.NameAscii = prev_ascii or "Times New Roman"
        try:
            font.NameOther = prev_ascii or "Times New Roman"
        except Exception:
            pass
    except Exception:
        pass
    try:
        rng.LanguageID = 2052  # wdSimplifiedChinese
        rng.LanguageIDFarEast = 2052
    except Exception:
        pass


def _collect_grid_updates(
    rows: list[list[str]],
    bindings: list,
    ctx: RenderContext,
    merges: list | None = None,
) -> list[tuple[int, int, Any]]:
    """收集待写单元格（0-based row/col）。

    - 合并覆盖位跳过（不写）
    - 有 ref 绑定：写解析值（可为空）
    - 无绑定：写 snapshot（含空串，表示清空；覆盖位空串仍跳过）
    """
    binding_map = {(b.row, b.col): b.ref_name for b in bindings if b.ref_name}
    updates: list[tuple[int, int, Any]] = []
    for row_idx, row in enumerate(rows):
        for col_idx, snapshot_val in enumerate(row):
            if is_covered(row_idx, col_idx, merges):
                if binding_map.get((row_idx, col_idx)):
                    raise EasyWordError(
                        "merged_cell_covered",
                        f"cellBindings 不能落在合并覆盖位 ({row_idx},{col_idx})",
                    )
                continue
            ref_name = binding_map.get((row_idx, col_idx))
            if ref_name:
                updates.append((row_idx, col_idx, _resolve_ref_value(ctx, ref_name)))
            else:
                # None 视为未提供该列（短行），跳过；空串表示用户清空，要写回
                if snapshot_val is None:
                    continue
                updates.append((row_idx, col_idx, snapshot_val))
    return updates


def _set_word_cell_text(cell: Any, value: Any) -> None:
    """写入单元格文本；空串清空内容但保留单元格结束标记。"""
    text = _cell_value_to_str(value)
    rng = cell.Range
    try:
        # Range 含末尾 \r\a；缩一格再赋值，避免拆坏表格结构
        if int(rng.End) > int(rng.Start):
            rng.MoveEnd(1, -1)  # wdCharacter
        rng.Text = text
    except Exception:  # noqa: BLE001
        cell.Range.Text = text


def _write_table_cell(
    rng: Any,
    row: int,
    col: int,
    value: Any,
    merges: list | None = None,
    row_count: int | None = None,
    col_count: int | None = None,
) -> None:
    """按逻辑格写入。横合表里 Word.Cell(r,c) 的 c 是物理序号，必须映射到 Range.Cells。"""
    if int(rng.Tables.Count) == 0:
        raise EasyWordError("no_table", "控件内未找到表格")
    table = rng.Tables.Item(1)
    n_rows = row_count if row_count is not None else int(table.Rows.Count)
    n_cols = col_count if col_count is not None else int(table.Columns.Count)
    if row < 0 or col < 0 or row >= n_rows or col >= n_cols:
        raise EasyWordError("cell_out_of_range", f"单元格 ({row},{col}) 超出表格范围")
    if is_covered(row, col, merges):
        raise EasyWordError(
            "merged_cell_covered",
            f"单元格 ({row},{col}) 为合并覆盖位，无法写入；请写合并锚点",
        )
    try:
        idx = physical_cell_index(row, col, n_rows, n_cols, merges)
        cell = table.Range.Cells.Item(idx)
    except EasyWordError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise EasyWordError(
            "merged_cell_covered",
            f"单元格 ({row},{col}) 无法定位物理格，无法写入",
        ) from exc
    _set_word_cell_text(cell, value)

def _write_chart_grid(rng: Any, updates: list[tuple[int, int, Any]]) -> None:
    """单次打开 ChartData.Workbook，批量写格，关簿前重绑 Word 图表面缓存。"""
    if not updates:
        return
    chart = _find_chart_in_range(rng)
    if chart is None:
        raise EasyWordError("no_chart", "控件内未找到图表")

    try:
        with chart_sheet_session(chart) as (_chart_data, _wb, _ws, used):
            max_r = int(used.Rows.Count)
            max_c = int(used.Columns.Count)
            for row_idx, col_idx, value in updates:
                r, c = row_idx + 1, col_idx + 1
                if r > max_r or c > max_c:
                    raise EasyWordError(
                        "cell_out_of_range",
                        f"单元格 ({row_idx},{col_idx}) 超出图表数据范围",
                    )
                used.Cells(r, c).Value = value
    except EasyWordError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_chart_com_error(exc, action="写入图表数据") from exc


def _apply_grid_cells(
    rng: Any,
    rows: list[list[str]],
    bindings: list,
    ctx: RenderContext,
    write_fn,
    merges: list | None = None,
) -> None:
    """表格：逐格写入手写/引用值。"""
    n_rows = len(rows)
    n_cols = max((len(r) for r in rows), default=0)
    for row_idx, col_idx, value in _collect_grid_updates(
        rows, bindings, ctx, merges=merges
    ):
        write_fn(
            rng,
            row_idx,
            col_idx,
            value,
            merges=merges,
            row_count=n_rows,
            col_count=n_cols,
        )


def _apply_template(template: ControlTemplate, ctx: RenderContext) -> None:
    doc = document_binding.get_bound_document()
    info = find_control_by_id(doc, template.control_id)
    cc = info.com
    rng = cc.Range

    if isinstance(template, TextControlTemplate):
        rendered = _render_jinja(template.jinja_template, ctx)
        _write_text_control(cc, rendered)
        return

    if isinstance(template, TableControlTemplate):
        merges = list(template.grid_snapshot.merges or [])
        _apply_grid_cells(
            rng,
            template.grid_snapshot.rows,
            template.cell_bindings,
            ctx,
            _write_table_cell,
            merges=merges,
        )
        return

    if isinstance(template, ChartControlTemplate):
        updates = _collect_grid_updates(
            template.grid_snapshot.rows,
            template.cell_bindings,
            ctx,
        )
        _write_chart_grid(rng, updates)
        return

    if isinstance(template, PictureControlTemplate):
        raise EasyWordError("not_implemented", "图片控件预览写回尚未实现")


def apply_preview_com(body: PreviewApplyIn) -> PreviewApplyOut:
    applied: list[PreviewAppliedItem] = []
    errors: list[PreviewErrorItem] = []

    for template in body.control_templates:
        try:
            _apply_template(template, body.render_context)
            applied.append(
                PreviewAppliedItem(controlId=template.control_id, type=template.type)
            )
        except EasyWordError as exc:
            errors.append(
                PreviewErrorItem(controlId=template.control_id, message=str(exc.message))
            )
        except Exception as exc:  # noqa: BLE001
            mapped = map_chart_com_error(exc)
            errors.append(
                PreviewErrorItem(controlId=template.control_id, message=str(mapped.message))
            )

    return PreviewApplyOut(ok=len(errors) == 0, applied=applied, errors=errors)
