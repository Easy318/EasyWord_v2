"""M0 预览 Schema 镜像测试。

fixture 复制自 EasyAnalytiQHub/qhub/test/pure_flow/fixtures/m0/（禁止跨仓 import）。
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from pydantic import TypeAdapter, ValidationError

from word.core.preview.schema.control_template import ControlTemplate
from word.core.preview.schema.preview import PreviewApplyIn
from word.core.preview.schema.render_context import RenderContext

FIXTURES = Path(__file__).parent / "fixtures" / "m0"
_CONTROL_ADAPTER = TypeAdapter(ControlTemplate)


def _load(name: str) -> dict:
    return json.loads((FIXTURES / name).read_text(encoding="utf-8"))


def test_render_context_valid() -> None:
    ctx = RenderContext.model_validate(_load("render_context.valid.json"))
    assert ctx.project_id == "proj-demo"
    assert ctx.burst is not None
    assert ctx.burst.var_name == "province"
    assert "a1" in ctx.entries


def test_render_context_allows_null_burst() -> None:
    ctx = RenderContext.model_validate(_load("render_context.no_burst.json"))
    assert ctx.burst is None
    assert "a1" in ctx.entries


def test_render_context_rejects_reserved_entry_keys() -> None:
    data = _load("render_context.valid.json")
    data["entries"]["burst_value"] = {
        "value": "x",
        "dtype": "str",
        "shape": None,
        "source": "burst",
        "status": "active",
    }
    with pytest.raises(ValidationError):
        RenderContext.model_validate(data)


def test_control_templates_valid() -> None:
    text = _CONTROL_ADAPTER.validate_python(_load("control_template.text.valid.json"))
    chart = _CONTROL_ADAPTER.validate_python(_load("control_template.chart.valid.json"))
    assert text.type == "text"
    assert chart.cell_bindings[0].ref_name == "a1"


def test_preview_apply_in_roundtrip() -> None:
    body = PreviewApplyIn.model_validate(
        {
            "renderContext": _load("render_context.valid.json"),
            "controlTemplates": [
                _load("control_template.text.valid.json"),
                _load("control_template.chart.valid.json"),
            ],
        }
    )
    assert body.render_context.project_id == "proj-demo"
    assert len(body.control_templates) == 2
