"""ContentControl schema。"""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field

ControlTypeLiteral = Literal["text", "picture", "table", "chart"]


class ContentControlOut(BaseModel):
    id: int
    type: ControlTypeLiteral | None = None
    title: str = ""
    tag: str = ""
    initial_text: str | None = Field(None, alias="initialText")

    model_config = {"populate_by_name": True}


class ContentControlListOut(BaseModel):
    items: list[ContentControlOut]


class ContentControlCreateIn(BaseModel):
    """类型由选区自动识别；type 仅作可选校验。"""

    type: ControlTypeLiteral | None = Field(default=None, description="可选；若传则须与选区识别一致")
    title: str | None = Field(default=None, description="显示标题")
    tag: str | None = Field(default=None, description="默认 eai:{type}")


class ContentControlUpdateIn(BaseModel):
    title: str | None = None
    tag: str | None = None


class OkOut(BaseModel):
    ok: bool = True
