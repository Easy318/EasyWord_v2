"""预览 API Schema。"""

from __future__ import annotations

from pydantic import BaseModel, Field

from word.core.preview.schema.common import CAMEL_CONFIG, ControlType
from word.core.preview.schema.control_template import ControlTemplate
from word.core.preview.schema.render_context import RenderContext


class PreviewApplyIn(BaseModel):
    render_context: RenderContext = Field(..., alias="renderContext")
    control_templates: list[ControlTemplate] = Field(..., alias="controlTemplates")

    model_config = CAMEL_CONFIG


class PreviewAppliedItem(BaseModel):
    control_id: int = Field(..., alias="controlId")
    type: ControlType

    model_config = CAMEL_CONFIG


class PreviewErrorItem(BaseModel):
    control_id: int | None = Field(None, alias="controlId")
    message: str

    model_config = CAMEL_CONFIG


class PreviewApplyOut(BaseModel):
    ok: bool = True
    applied: list[PreviewAppliedItem] = Field(default_factory=list)
    errors: list[PreviewErrorItem] = Field(default_factory=list)

    model_config = CAMEL_CONFIG
