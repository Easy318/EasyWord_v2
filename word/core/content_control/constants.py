"""ContentControl 业务类型与 Word 常量。"""

from __future__ import annotations

from enum import Enum

# WdContentControlType
WD_CC_RICH_TEXT = 0
WD_CC_TEXT = 1
WD_CC_PICTURE = 2

TAG_PREFIX = "eai:"


class ControlType(str, Enum):
    text = "text"
    picture = "picture"
    table = "table"
    chart = "chart"


def default_tag(control_type: ControlType) -> str:
    return f"{TAG_PREFIX}{control_type.value}"


def word_cc_type_for_create(control_type: ControlType) -> int:
    if control_type == ControlType.text:
        return WD_CC_TEXT
    if control_type == ControlType.picture:
        return WD_CC_PICTURE
    # table / chart：RichText 包裹已有对象
    return WD_CC_RICH_TEXT


def parse_business_type(tag: str, word_type: int) -> ControlType | None:
    """根据 Tag（优先）与 Word 类型推断业务类型；非本产品控件返回 None。"""
    t = (tag or "").strip().lower()
    if t == default_tag(ControlType.text) or t == "eai:text":
        return ControlType.text
    if t == default_tag(ControlType.picture):
        return ControlType.picture
    if t == default_tag(ControlType.table):
        return ControlType.table
    if t == default_tag(ControlType.chart):
        return ControlType.chart
    # 兼容：无 Tag 但类型为标准 Picture/Text 时不认作业务控件（避免误伤用户自建 CC）
    _ = word_type
    return None
