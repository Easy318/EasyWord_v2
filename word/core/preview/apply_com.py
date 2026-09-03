"""Word COM 预览写回（text Jinja + chart/table cell 值）。"""

from __future__ import annotations

from typing import Any

from jinja2 import Environment, TemplateSyntaxError, UndefinedError

from word.core.content_control.chart_workbook import (
    map_chart_com_error,
    open_chart_sheet,
    release_chart_workbook,
)
from word.core.content_control.constants import ControlType
from word.core.content_control.selection_guard import find_control_by_id
from word.core.content_control.template_seed import (
    _cell_value_to_str,
    _find_chart_in_range,
)
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


_JINJA_ENV = Environment(autoescape=False, finalize=_jinja_finalize)
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


def _render_jinja(source: str, ctx: RenderContext) -> str:
    try:
        return _JINJA_ENV.from_string(source).render(**_build_jinja_context(ctx))
    except TemplateSyntaxError as exc:
        raise EasyWordError("jinja_syntax", f"模板语法错误: {exc}") from exc
    except UndefinedError as exc:
        raise EasyWordError("jinja_undefined", f"模板变量未定义: {exc}") from exc


def _write_text_control(cc: Any, text: str) -> None:
    cc.Range.Text = text


def _write_table_cell(rng: Any, row: int, col: int, value: Any) -> None:
    if int(rng.Tables.Count) == 0:
        raise EasyWordError("no_table", "控件内未找到表格")
    table = rng.Tables.Item(1)
    r, c = row + 1, col + 1
    if r > int(table.Rows.Count) or c > int(table.Columns.Count):
        raise EasyWordError("cell_out_of_range", f"单元格 ({row},{col}) 超出表格范围")
    table.Cell(r, c).Range.Text = _cell_value_to_str(value)


def _collect_grid_updates(
    rows: list[list[str]],
    bindings: list,
    ctx: RenderContext,
) -> list[tuple[int, int, Any]]:
    """收集待写单元格（0-based row/col）。"""
    binding_map = {(b.row, b.col): b.ref_name for b in bindings if b.ref_name}
    updates: list[tuple[int, int, Any]] = []
    for row_idx, row in enumerate(rows):
        for col_idx, snapshot_val in enumerate(row):
            ref_name = binding_map.get((row_idx, col_idx))
            if ref_name:
                updates.append((row_idx, col_idx, _resolve_ref_value(ctx, ref_name)))
            else:
                if snapshot_val is None or str(snapshot_val).strip() == "":
                    continue
                updates.append((row_idx, col_idx, snapshot_val))
    return updates


def _write_chart_grid(rng: Any, updates: list[tuple[int, int, Any]]) -> None:
    """单次打开 ChartData.Workbook，批量写格后尽量释放。"""
    if not updates:
        return
    chart = _find_chart_in_range(rng)
    if chart is None:
        raise EasyWordError("no_chart", "控件内未找到图表")

    wb: Any = None
    try:
        _chart_data, wb, _ws, used = open_chart_sheet(chart)
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
    finally:
        release_chart_workbook(wb, chart)


def _apply_grid_cells(
    rng: Any,
    rows: list[list[str]],
    bindings: list,
    ctx: RenderContext,
    write_fn,
) -> None:
    """表格：逐格写入手写/引用值。"""
    for row_idx, col_idx, value in _collect_grid_updates(rows, bindings, ctx):
        write_fn(rng, row_idx, col_idx, value)


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
        _apply_grid_cells(
            rng,
            template.grid_snapshot.rows,
            template.cell_bindings,
            ctx,
            _write_table_cell,
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
