"""模板种子 schema。"""

from __future__ import annotations

from typing import Any

from pydantic import BaseModel, Field

from word.core.content_control.schema.content_control import ControlTypeLiteral
from word.core.preview.schema.control_template import CellBinding, GridSnapshot


class TemplateSeedOut(BaseModel):
    control_id: int = Field(..., alias="controlId")
    type: ControlTypeLiteral | None = None
    control_tag: str = Field("", alias="controlTag")
    jinja_template: str | None = Field(None, alias="jinjaTemplate")
    grid_snapshot: GridSnapshot | None = Field(None, alias="gridSnapshot")
    cell_bindings: list[CellBinding] | None = Field(None, alias="cellBindings")
    placeholder: bool | None = None

    model_config = {"populate_by_name": True}


class ContentControlCreateOut(BaseModel):
    id: int
    type: ControlTypeLiteral | None = None
    title: str = ""
    tag: str = ""
    initial_text: str | None = Field(None, alias="initialText")

    model_config = {"populate_by_name": True}
