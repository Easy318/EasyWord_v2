"""预览写回用的数字字符串化（去浮点噪点 / 固定小数位）。"""

from __future__ import annotations

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any

_MAX_FMT_PLACES = 10


def clean_float_str(value: Any) -> str:
    """把值写成适合 Word 展示的字符串；有限 float 去掉二进制尾巴。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not (value == value) or value in (float("inf"), float("-inf")):
            return str(value)
        if value == 0:
            return "0"
        if value == int(value) and abs(value) < 1e15:
            return str(int(value))
        text = format(value, ".12g")
        if text in ("-0", "-0.0"):
            return "0"
        return text
    return str(value)


def fmt_number(value: Any, places: int = 2) -> str:
    """Decimal HALF_UP 固定小数位（对齐计算节点 / ResultDock toFixed）。"""
    n = max(0, min(_MAX_FMT_PLACES, int(places)))
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    try:
        if isinstance(value, Decimal):
            dec = value
        elif isinstance(value, (int, float)):
            if isinstance(value, float) and (
                not (value == value) or value in (float("inf"), float("-inf"))
            ):
                return str(value)
            dec = Decimal(str(value))
        else:
            dec = Decimal(str(value).strip())
    except (InvalidOperation, ValueError, TypeError):
        return str(value)

    quant = Decimal("1") if n == 0 else Decimal(f"1e-{n}")
    quantized = dec.quantize(quant, rounding=ROUND_HALF_UP)
    if n == 0:
        return str(int(quantized))
    return f"{quantized:.{n}f}"
