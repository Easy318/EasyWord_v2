"""ContentControl 路由。"""

from __future__ import annotations

from fastapi import APIRouter

from word.core.content_control import service
from word.core.content_control.constants import ControlType
from word.core.content_control.schema import (
    ContentControlCreateIn,
    ContentControlListOut,
    ContentControlOut,
    ContentControlUpdateIn,
    OkOut,
    TemplateSeedOut,
)

content_control_api = APIRouter(prefix="/content-controls", tags=["content-controls"])


@content_control_api.get("", response_model=ContentControlListOut)
def list_content_controls() -> ContentControlListOut:
    items = service.list_controls()
    return ContentControlListOut(items=[ContentControlOut.model_validate(i) for i in items])


@content_control_api.get("/selection", response_model=ContentControlOut)
def get_selection_content_control() -> ContentControlOut:
    return ContentControlOut.model_validate(service.get_selection_control())


@content_control_api.post("", response_model=ContentControlOut)
def create_content_control(body: ContentControlCreateIn | None = None) -> ContentControlOut:
    payload = body or ContentControlCreateIn()
    ct = ControlType(payload.type) if payload.type else None
    result = service.create_control(ct, title=payload.title, tag=payload.tag)
    return ContentControlOut.model_validate(result)


@content_control_api.patch("", response_model=ContentControlOut)
def update_content_control_by_selection(body: ContentControlUpdateIn) -> ContentControlOut:
    result = service.update_control(title=body.title, tag=body.tag, control_id=None)
    return ContentControlOut.model_validate(result)


@content_control_api.delete("", response_model=ContentControlOut)
def delete_content_control_by_selection() -> ContentControlOut:
    return ContentControlOut.model_validate(service.delete_control(control_id=None))


@content_control_api.patch("/{control_id}", response_model=ContentControlOut)
def update_content_control(control_id: int, body: ContentControlUpdateIn) -> ContentControlOut:
    result = service.update_control(title=body.title, tag=body.tag, control_id=control_id)
    return ContentControlOut.model_validate(result)


@content_control_api.delete("/{control_id}", response_model=ContentControlOut)
def delete_content_control_by_id(control_id: int) -> ContentControlOut:
    result = service.delete_control(control_id)
    return ContentControlOut.model_validate(result)


@content_control_api.post("/{control_id}/locate", response_model=ContentControlOut)
def locate_content_control(control_id: int) -> ContentControlOut:
    result = service.locate_control(control_id)
    return ContentControlOut.model_validate(result)


@content_control_api.get("/{control_id}/template-seed", response_model=TemplateSeedOut)
def get_template_seed(control_id: int) -> TemplateSeedOut:
    result = service.get_template_seed(control_id)
    return TemplateSeedOut.model_validate(result)
