"""ContentControl CRUD（COM 线程内逻辑）。"""

from __future__ import annotations

from typing import Any

from word.core.content_control.constants import ControlType, default_tag, word_cc_type_for_create
from word.core.content_control.selection_guard import (
    assert_create_selection_ok,
    list_content_controls,
    resolve_target_control,
)
from word.runtime.com_host import com_host
from word.runtime.document_binding import document_binding
from word.runtime.errors import EasyWordError
from word.runtime.word_app import word_app


def _cc_to_dict(cc_info: Any) -> dict[str, object]:
    bt = cc_info.business_type
    return {
        "id": cc_info.id,
        "type": bt.value if bt is not None else None,
        "title": cc_info.title,
        "tag": cc_info.tag,
    }


def list_controls() -> list[dict[str, object]]:
    def _run() -> list[dict[str, object]]:
        doc = document_binding.get_bound_document()
        items = list_content_controls(doc)
        out: list[dict[str, object]] = []
        for info in items:
            if info.business_type is None:
                continue
            out.append(_cc_to_dict(info))
        return out

    return com_host.submit(_run)


def get_selection_control() -> dict[str, object]:
    def _run() -> dict[str, object]:
        doc = document_binding.get_bound_document()
        info = resolve_target_control(doc, control_id=None)
        return _cc_to_dict(info)

    return com_host.submit(_run)


def create_control(
    control_type: ControlType | None = None,
    title: str | None = None,
    tag: str | None = None,
) -> dict[str, object]:
    def _run() -> dict[str, object]:
        doc = document_binding.get_bound_document()
        rng, detected = assert_create_selection_ok(doc)
        if control_type is not None and control_type != detected:
            raise EasyWordError(
                "type_mismatch",
                f"选区识别为 {detected.value}，与请求类型 {control_type.value} 不一致",
            )
        resolved = detected

        text = str(rng.Text or "")
        if len(text) == 0:
            raise EasyWordError("empty_control", "选区为空，禁止创建空控件")

        wd_type = word_cc_type_for_create(resolved)
        try:
            cc = doc.ContentControls.Add(wd_type, rng)
        except Exception as exc:  # noqa: BLE001
            raise EasyWordError("create_failed", f"创建内容控件失败: {exc}", status_code=500) from exc

        use_tag = (tag or "").strip() or default_tag(resolved)
        use_title = (title or "").strip() or resolved.value
        try:
            cc.Tag = use_tag
            cc.Title = use_title
        except Exception as exc:  # noqa: BLE001
            try:
                cc.Delete(False)
            except Exception:  # noqa: BLE001
                pass
            raise EasyWordError("create_failed", f"设置控件属性失败: {exc}", status_code=500) from exc

        try:
            after = str(cc.Range.Text or "")
            if len(after) == 0 and resolved == ControlType.text:
                cc.Delete(False)
                raise EasyWordError("empty_control", "创建结果为空控件，已撤销")
        except EasyWordError:
            raise
        except Exception:  # noqa: BLE001
            pass

        payload: dict[str, object] = {
            "id": int(cc.ID),
            "type": resolved.value,
            "title": str(cc.Title or ""),
            "tag": str(cc.Tag or ""),
        }
        if resolved == ControlType.text:
            from word.core.content_control.template_seed import _normalize_cc_text

            payload["initialText"] = _normalize_cc_text(text)
        return payload

    return com_host.submit(_run)


def update_control(
    title: str | None = None,
    tag: str | None = None,
    control_id: int | None = None,
) -> dict[str, object]:
    def _run() -> dict[str, object]:
        doc = document_binding.get_bound_document()
        info = resolve_target_control(doc, control_id)
        cc = info.com
        if title is not None:
            cc.Title = title
        if tag is not None:
            cc.Tag = tag
        return {
            "id": int(cc.ID),
            "type": info.business_type.value if info.business_type else None,
            "title": str(cc.Title or ""),
            "tag": str(cc.Tag or ""),
        }

    return com_host.submit(_run)


def delete_control(control_id: int | None = None) -> dict[str, object]:
    def _run() -> dict[str, object]:
        from word.core.content_control.selection_guard import find_control_by_id

        doc = document_binding.get_bound_document()
        if control_id is not None:
            info = find_control_by_id(doc, control_id)
        else:
            info = resolve_target_control(doc, None)
        payload = _cc_to_dict(info)
        try:
            info.com.Delete(False)
        except Exception as exc:  # noqa: BLE001
            raise EasyWordError("delete_failed", f"删除控件失败: {exc}", status_code=500) from exc
        return payload

    return com_host.submit(_run)


def get_template_seed(control_id: int) -> dict[str, object]:
    from word.core.content_control import template_seed

    def _run() -> dict[str, object]:
        return template_seed.get_template_seed(control_id)

    return com_host.submit(_run)


def locate_control(control_id: int) -> dict[str, object]:
    """定位到指定控件（选中并滚动到可见区域，不修改应用会话状态）。"""

    def _run() -> dict[str, object]:
        from word.core.content_control.selection_guard import find_control_by_id

        doc = document_binding.get_bound_document()
        info = find_control_by_id(doc, control_id)
        rng = info.com.Range
        try:
            rng.Select()
        except Exception as exc:  # noqa: BLE001
            raise EasyWordError("locate_failed", f"定位控件失败: {exc}", status_code=500) from exc
        try:
            app = word_app.ensure_app()
            app.ActiveWindow.ScrollIntoView(rng, True)
        except Exception:  # noqa: BLE001
            pass
        try:
            from word.runtime.window_snap import focus_word_window

            focus_word_window(app)
        except Exception:  # noqa: BLE001
            pass
        return _cc_to_dict(info)

    return com_host.submit(_run)
