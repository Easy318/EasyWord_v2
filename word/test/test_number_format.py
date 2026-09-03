"""Jinja / 表格数字展示：去噪与 |fmt。"""

from __future__ import annotations

from word.core.content_control.template_seed import _cell_value_to_str
from word.core.preview.apply_com import _render_jinja
from word.core.preview.number_format import clean_float_str, fmt_number
from word.core.preview.schema.render_context import RenderContext, RenderValue


def _ctx(**vals: float) -> RenderContext:
    entries = {
        name: RenderValue(
            value=value,
            dtype="float",
            shape=None,
            source="flow",
            status="active",
        )
        for name, value in vals.items()
    }
    return RenderContext(schemaVersion="1.0", projectId="p1", burst=None, entries=entries)


def test_clean_float_str_strips_binary_noise() -> None:
    noisy = 2.22 - 1.97
    assert noisy != 0.25  # IEEE artifact present
    assert clean_float_str(noisy) == "0.25"
    assert clean_float_str(1.0) == "1"
    assert clean_float_str(None) == ""


def test_fmt_number_half_up_fixed_places() -> None:
    assert fmt_number(2.22 - 1.97, 2) == "0.25"
    assert fmt_number(1.975, 2) == "1.98"
    assert fmt_number(1.5, 2) == "1.50"
    assert fmt_number(3, 0) == "3"


def test_cell_value_to_str_uses_clean_float() -> None:
    assert _cell_value_to_str(2.22 - 1.97) == "0.25"
    assert _cell_value_to_str(2.0) == "2"


def test_render_jinja_finalize_cleans_subtraction() -> None:
    ctx = _ctx(a1=1.97, b1=2.22)
    out = _render_jinja("低于全区 {{ b1 - a1 }} 分", ctx)
    assert out == "低于全区 0.25 分"
    assert "000000" not in out


def test_render_jinja_fmt_and_round() -> None:
    ctx = _ctx(a1=1.97, b1=2.22)
    assert _render_jinja("{{ (b1 - a1)|fmt(2) }}", ctx) == "0.25"
    assert _render_jinja("{{ (b1 - a1)|round(2) }}", ctx) == "0.25"
    assert _render_jinja("{{ a1|fmt(2) }}", ctx) == "1.97"
