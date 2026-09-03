"""Selection 与 ContentControl Range 关系校验。"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Any

from word.core.content_control.constants import (
    ControlType,
    parse_business_type,
)
from word.runtime.errors import EasyWordError
from word.runtime.paths import normalize_path
from word.runtime.word_app import word_app


class RangeRelation(str, Enum):
    disjoint = "disjoint"
    equal = "equal"
    sel_inside_cc = "sel_inside_cc"
    cc_inside_sel = "cc_inside_sel"
    cross = "cross"


@dataclass
class CcInfo:
    id: int
    title: str
    tag: str
    word_type: int
    business_type: ControlType | None
    start: int
    end: int
    com: Any


def range_relation(sel_start: int, sel_end: int, cc_start: int, cc_end: int) -> RangeRelation:
    if sel_end <= cc_start or sel_start >= cc_end:
        return RangeRelation.disjoint
    if sel_start == cc_start and sel_end == cc_end:
        return RangeRelation.equal
    if sel_start >= cc_start and sel_end <= cc_end:
        return RangeRelation.sel_inside_cc
    if cc_start >= sel_start and cc_end <= sel_end:
        return RangeRelation.cc_inside_sel
    return RangeRelation.cross


def _selection_in_document(selection: Any, doc: Any) -> bool:
    try:
        sel_doc = selection.Document
        return normalize_path(str(sel_doc.FullName)) == normalize_path(str(doc.FullName))
    except Exception:  # noqa: BLE001
        return False


def list_content_controls(doc: Any) -> list[CcInfo]:
    result: list[CcInfo] = []
    ccs = doc.ContentControls
    count = int(ccs.Count)
    for i in range(1, count + 1):
        cc = ccs.Item(i)
        try:
            rng = cc.Range
            start, end = int(rng.Start), int(rng.End)
            tag = str(cc.Tag or "")
            title = str(cc.Title or "")
            word_type = int(cc.Type)
            cid = int(cc.ID)
        except Exception:  # noqa: BLE001
            continue
        result.append(
            CcInfo(
                id=cid,
                title=title,
                tag=tag,
                word_type=word_type,
                business_type=parse_business_type(tag, word_type),
                start=start,
                end=end,
                com=cc,
            )
        )
    return result


def require_non_empty_selection(doc: Any) -> tuple[Any, int, int]:
    app = word_app.ensure_app()
    selection = app.Selection
    if not _selection_in_document(selection, doc):
        raise EasyWordError("selection_wrong_doc", "当前选区不在已绑定的项目模板文档中", status_code=409)
    try:
        rng = selection.Range
        start, end = int(rng.Start), int(rng.End)
    except Exception as exc:  # noqa: BLE001
        raise EasyWordError("selection_invalid", f"无法读取选区: {exc}") from exc
    if start == end:
        raise EasyWordError("selection_empty", "请先在 Word 中选中目标内容（选区不能为空）")
    return selection, start, end


def intersecting_controls(ccs: list[CcInfo], sel_start: int, sel_end: int) -> list[tuple[CcInfo, RangeRelation]]:
    hits: list[tuple[CcInfo, RangeRelation]] = []
    for cc in ccs:
        rel = range_relation(sel_start, sel_end, cc.start, cc.end)
        if rel != RangeRelation.disjoint:
            hits.append((cc, rel))
    return hits


def assert_create_selection_ok(doc: Any) -> tuple[Any, ControlType]:
    """校验 create 选区并自动识别类型；返回 (Range, ControlType)。"""
    selection, sel_start, sel_end = require_non_empty_selection(doc)
    ccs = list_content_controls(doc)
    hits = intersecting_controls(ccs, sel_start, sel_end)
    if hits:
        for cc, rel in hits:
            if rel == RangeRelation.sel_inside_cc or rel == RangeRelation.equal:
                raise EasyWordError("nested", "选区位于已有内容控件内部，禁止嵌套创建")
            if rel == RangeRelation.cc_inside_sel:
                raise EasyWordError("nested", "选区包含已有内容控件，禁止嵌套创建")
            if rel == RangeRelation.cross:
                raise EasyWordError("overlap", "选区与已有内容控件交叉覆盖，无法创建")
        raise EasyWordError("overlap", "选区与已有内容控件冲突，无法创建")

    # 同文档允许多个同类型业务控件；限制只在「选区不嵌套/不交叉」与「单次选区只识别一种类型」。
    control_type = detect_selection_control_type(selection)
    return selection.Range, control_type


def detect_selection_control_type(selection: Any) -> ControlType:
    """根据选区内容自动识别唯一业务类型。"""
    rng = selection.Range
    has_chart = _selection_has_chart(rng)
    has_picture = _selection_has_picture(rng)
    has_table = int(rng.Tables.Count) > 0
    text = str(rng.Text or "").replace("\r", "").replace("\a", "").strip()

    # 图表优先于图片误判
    object_kinds: list[ControlType] = []
    if has_chart:
        object_kinds.append(ControlType.chart)
    elif has_picture:
        object_kinds.append(ControlType.picture)
    if has_table:
        object_kinds.append(ControlType.table)

    if len(object_kinds) > 1:
        names = "、".join(k.value for k in object_kinds)
        raise EasyWordError("mixed_types", f"选区包含多种内容类型（{names}），请只选中一种")

    if object_kinds:
        return object_kinds[0]

    if text:
        return ControlType.text

    raise EasyWordError("type_unrecognized", "无法识别选区内容类型，请选中文字、图片、表格或图表")


def find_control_by_id(doc: Any, control_id: int) -> CcInfo:
    """按 Word ContentControl ID 查找业务控件（无需选区）。"""
    for cc in list_content_controls(doc):
        if int(cc.id) == int(control_id):
            if cc.business_type is None:
                raise EasyWordError("not_business_control", f"控件 {control_id} 不是业务内容控件")
            return cc
    raise EasyWordError("control_not_found", f"未找到 ID 为 {control_id} 的内容控件", status_code=404)


def resolve_target_control(doc: Any, control_id: int | None = None) -> CcInfo:
    """update/delete：选区须唯一命中恰好一个业务 CC；若传入 control_id 则交叉校验。"""
    _selection, sel_start, sel_end = require_non_empty_selection(doc)
    ccs = list_content_controls(doc)
    hits = intersecting_controls(ccs, sel_start, sel_end)
    # 只认业务 Tag 控件
    business_hits = [(cc, rel) for cc, rel in hits if cc.business_type is not None]
    if not business_hits:
        if hits:
            raise EasyWordError("selection_miss", "选区命中的不是业务内容控件")
        raise EasyWordError("selection_miss", "选区未命中任何内容控件")
    if len(business_hits) > 1:
        types = {h[0].business_type.value for h in business_hits if h[0].business_type}
        if len(types) > 1:
            raise EasyWordError("mixed_types", "选区包含多种控件类型，请缩小选区到单一控件")
        raise EasyWordError("ambiguous", "选区命中多个内容控件，请精确选中目标控件")

    cc, rel = business_hits[0]
    if rel == RangeRelation.cross:
        raise EasyWordError("overlap", "选区与目标控件交叉覆盖，请完整选中该控件或选中其内部")

    if control_id is not None and int(cc.id) != int(control_id):
        raise EasyWordError(
            "id_mismatch",
            f"选区命中的控件 ID 为 {cc.id}，与请求的 {control_id} 不一致",
        )
    return cc


def _selection_has_picture(rng: Any) -> bool:
    try:
        inlines = rng.InlineShapes
        for i in range(1, int(inlines.Count) + 1):
            shape = inlines.Item(i)
            # wdInlineShapePicture=3, wdInlineShapeLinkedPicture=4, chart=12
            t = int(shape.Type)
            if t in (3, 4):
                return True
            if t == 12:
                continue
            # 无 Chart 属性则可能是图
            try:
                _ = shape.Chart
            except Exception:  # noqa: BLE001
                if t not in (12,):
                    # 保守：非 chart 的 inline 当作可能的图
                    if t in (1, 3, 4, 5, 7):
                        return True
    except Exception:  # noqa: BLE001
        pass
    return False


def _selection_has_chart(rng: Any) -> bool:
    try:
        inlines = rng.InlineShapes
        for i in range(1, int(inlines.Count) + 1):
            shape = inlines.Item(i)
            try:
                _ = shape.Chart
                return True
            except Exception:  # noqa: BLE001
                if int(shape.Type) == 12:  # wdInlineShapeChart
                    return True
    except Exception:  # noqa: BLE001
        pass
    try:
        # Shape 范围内的图表（浮动）
        doc = rng.Document
        for i in range(1, int(doc.Shapes.Count) + 1):
            sh = doc.Shapes.Item(i)
            try:
                anchor_start = int(sh.Anchor.Start)
                if rng.Start <= anchor_start < rng.End:
                    _ = sh.Chart
                    return True
            except Exception:  # noqa: BLE001
                continue
    except Exception:  # noqa: BLE001
        pass
    return False
