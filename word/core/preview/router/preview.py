"""预览路由（M2）。"""

from __future__ import annotations

from fastapi import APIRouter

from word.core.preview import service
from word.core.preview.schema import PreviewApplyIn, PreviewApplyOut

preview_api = APIRouter(prefix="/preview", tags=["preview"])


@preview_api.post("/apply", response_model=PreviewApplyOut)
def apply_preview(body: PreviewApplyIn) -> PreviewApplyOut:
    return service.apply_preview(body)
