"""预览服务。"""

from __future__ import annotations

from word.core.preview.apply_com import apply_preview_com
from word.core.preview.schema import PreviewApplyIn, PreviewApplyOut
from word.runtime.com_host import com_host


def apply_preview(body: PreviewApplyIn) -> PreviewApplyOut:
    def _run() -> PreviewApplyOut:
        return apply_preview_com(body)

    return com_host.submit(_run)
