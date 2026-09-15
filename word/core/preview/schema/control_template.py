"""控件模板流镜像（与 QHub 线格式一致，禁止 import QHub）。"""

from __future__ import annotations

from typing import Annotated, Literal, Union

from pydantic import BaseModel, Field

from word.core.preview.schema.common import CAMEL_CONFIG, REF_NAME_PATTERN, SCHEMA_VERSION


class _ControlBase(BaseModel):
    schema_version: str = Field(SCHEMA_VERSION, alias="schemaVersion")
    control_id: int = Field(..., alias="controlId")
    control_tag: str = Field(..., alias="controlTag")

    model_config = CAMEL_CONFIG


class TextControlTemplate(_ControlBase):
    type: Literal["text"] = "text"
    jinja_template: str = Field(..., alias="jinjaTemplate")


class GridMerge(BaseModel):
    """逻辑网格上的合并锚点（0-based）；仅 table 使用。"""

    row: int = Field(..., ge=0)
    col: int = Field(..., ge=0)
    row_span: int = Field(..., ge=1, alias="rowSpan")
    col_span: int = Field(..., ge=1, alias="colSpan")

    model_config = CAMEL_CONFIG


class GridSnapshot(BaseModel):
    rows: list[list[str]] = Field(default_factory=list)
    sheet_name: str | None = Field(None, alias="sheetName")
    merges: list[GridMerge] = Field(default_factory=list)

    model_config = CAMEL_CONFIG


class CellBinding(BaseModel):
    row: int = Field(..., ge=0)
    col: int = Field(..., ge=0)
    ref_name: str | None = Field(None, alias="refName", pattern=REF_NAME_PATTERN)

    model_config = CAMEL_CONFIG


class RegionMapping(BaseModel):
    region_key: str = Field(..., alias="regionKey")
    ref_name: str = Field(..., alias="refName", pattern=REF_NAME_PATTERN)
    role: str | None = None

    model_config = CAMEL_CONFIG


class ChartControlTemplate(_ControlBase):
    type: Literal["chart"] = "chart"
    grid_snapshot: GridSnapshot = Field(default_factory=GridSnapshot, alias="gridSnapshot")
    cell_bindings: list[CellBinding] = Field(default_factory=list, alias="cellBindings")
    region_mappings: list[RegionMapping] = Field(default_factory=list, alias="regionMappings")


class TableControlTemplate(_ControlBase):
    type: Literal["table"] = "table"
    grid_snapshot: GridSnapshot = Field(default_factory=GridSnapshot, alias="gridSnapshot")
    cell_bindings: list[CellBinding] = Field(default_factory=list, alias="cellBindings")
    region_mappings: list[RegionMapping] = Field(default_factory=list, alias="regionMappings")


class PictureControlTemplate(_ControlBase):
    type: Literal["picture"] = "picture"
    source_ref: str | None = Field(None, alias="sourceRef")


ControlTemplate = Annotated[
    Union[
        TextControlTemplate,
        ChartControlTemplate,
        TableControlTemplate,
        PictureControlTemplate,
    ],
    Field(discriminator="type"),
]
