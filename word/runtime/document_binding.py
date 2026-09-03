"""当前项目模板文档绑定（按规范化路径）。"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from word.logger import get_logger
from word.runtime.errors import EasyWordError
from word.runtime.paths import normalize_path, paths_equal
from word.runtime.window_snap import snap_word_to_left_half
from word.runtime.word_app import word_app

_log = get_logger("runtime")

# wdFormatXMLDocument — .docx
_WD_FORMAT_XML_DOCUMENT = 12


@dataclass
class BindingState:
    project_id: str
    template_path: str


class DocumentBinding:
    def __init__(self) -> None:
        self._state: BindingState | None = None

    @property
    def state(self) -> BindingState | None:
        return self._state

    def is_bound(self) -> bool:
        return self._state is not None

    def require_bound(self) -> BindingState:
        if self._state is None:
            raise EasyWordError("not_bound", "尚未打开项目模板文档", status_code=409)
        return self._state

    def find_document_by_path(self, app: Any, path: str) -> Any | None:
        target = normalize_path(path)
        try:
            count = int(app.Documents.Count)
        except Exception as exc:  # noqa: BLE001
            raise EasyWordError("word_error", f"无法枚举 Word 文档: {exc}", status_code=500) from exc
        for i in range(1, count + 1):
            doc = app.Documents.Item(i)
            try:
                full = str(doc.FullName)
            except Exception:  # noqa: BLE001
                continue
            if paths_equal(full, path) or normalize_path(full) == target:
                return doc
        return None

    def _resolve_bound_document(self, app: Any, state: BindingState) -> Any:
        doc = self.find_document_by_path(app, state.template_path)
        if doc is not None:
            return doc
        try:
            active = app.ActiveDocument
            if active is not None and paths_equal(str(active.FullName), state.template_path):
                return active
        except Exception:  # noqa: BLE001
            pass
        raise EasyWordError("document_missing", "绑定文档已不在 Word 中打开", status_code=409)

    def _save_document_to_template(self, doc: Any, template_path: str) -> None:
        """强制落盘到绑定模板路径，确保客户端可读本地 template.docx。"""
        target = normalize_path(template_path)
        try:
            doc.SaveAs2(target, _WD_FORMAT_XML_DOCUMENT)
            _log.info(f"Saved template via SaveAs2: {target}")
            return
        except Exception as exc_save_as:  # noqa: BLE001
            _log.warning(f"SaveAs2 failed ({exc_save_as}), trying Save()")
        try:
            doc.Save()
            full = normalize_path(str(doc.FullName))
            if full != target:
                raise EasyWordError(
                    "save_path_mismatch",
                    f"文档保存路径与模板不一致：{doc.FullName}",
                    status_code=500,
                )
            _log.info(f"Saved template via Save: {target}")
        except EasyWordError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise EasyWordError("save_failed", f"保存失败: {exc}", status_code=500) from exc

    def open(self, project_id: str, template_path: str) -> dict[str, object]:
        """打开或重绑模板；切换项目时先保存并关闭旧绑定。"""
        from pathlib import Path

        path = normalize_path(template_path)
        if not Path(path).is_file():
            raise EasyWordError("file_not_found", f"模板文件不存在: {template_path}", status_code=404)

        app = word_app.ensure_app()
        rebound = False

        if self._state is not None:
            old = self._state
            if old.project_id == project_id and normalize_path(old.template_path) == path:
                doc = self.find_document_by_path(app, path)
                if doc is not None:
                    doc.Activate()
                    snap_word_to_left_half(app)
                    rebound = True
                    return {
                        "projectId": project_id,
                        "templatePath": path,
                        "rebound": rebound,
                    }
            else:
                self._close_bound(save=True)

        doc = self.find_document_by_path(app, path)
        if doc is not None:
            doc.Activate()
            rebound = True
        else:
            try:
                doc = app.Documents.Open(path, ReadOnly=False, AddToRecentFiles=False)
            except Exception as exc:  # noqa: BLE001
                raise EasyWordError("open_failed", f"打开文档失败: {exc}", status_code=500) from exc
            doc.Activate()

        snap_word_to_left_half(app)

        self._state = BindingState(project_id=project_id, template_path=path)
        _log.info(f"Bound document projectId={project_id} path={path} rebound={rebound}")
        return {
            "projectId": project_id,
            "templatePath": path,
            "rebound": rebound,
        }

    def save(self) -> None:
        state = self.require_bound()
        app = word_app.ensure_app()
        doc = self._resolve_bound_document(app, state)
        self._save_document_to_template(doc, state.template_path)

    def close(self, save: bool = True) -> None:
        if self._state is None:
            return
        self._close_bound(save=save)

    def _close_bound(self, save: bool) -> None:
        state = self._state
        if state is None:
            return
        app = word_app.try_app()
        if app is not None:
            doc = self.find_document_by_path(app, state.template_path)
            if doc is None:
                try:
                    active = app.ActiveDocument
                    if active is not None and paths_equal(str(active.FullName), state.template_path):
                        doc = active
                except Exception:  # noqa: BLE001
                    pass
            if doc is not None:
                if save:
                    self._save_document_to_template(doc, state.template_path)
                try:
                    # 已显式 SaveAs2 后不再重复保存
                    doc.Close(SaveChanges=0)
                except Exception as exc:  # noqa: BLE001
                    raise EasyWordError("close_failed", f"关闭文档失败: {exc}", status_code=500) from exc
            elif save:
                raise EasyWordError("document_missing", "绑定文档已不在 Word 中打开，无法保存关闭", status_code=409)
            word_app.quit_if_no_documents()
        self._state = None
        _log.info("Document unbound")

    def get_bound_document(self) -> Any:
        state = self.require_bound()
        app = word_app.ensure_app()
        return self._resolve_bound_document(app, state)

    def status(self) -> dict[str, object]:
        app = word_app.try_app()
        word_available = app is not None
        bound = self._state is not None
        document_open = False
        project_id = None
        template_path = None
        if self._state is not None:
            project_id = self._state.project_id
            template_path = self._state.template_path
            if app is not None:
                document_open = self.find_document_by_path(app, self._state.template_path) is not None
        return {
            "bound": bound,
            "projectId": project_id,
            "templatePath": template_path,
            "wordAvailable": word_available,
            "documentOpen": document_open,
        }

    def finalize_by_path(self, template_path: str) -> None:
        """按模板路径保存并关闭 Word 文档（不依赖内存绑定，用于退出/热重载后落盘）。"""
        path = normalize_path(template_path)
        app = word_app.try_app()
        if app is None:
            if self._state is not None and normalize_path(self._state.template_path) == path:
                self._state = None
            return

        doc = self.find_document_by_path(app, path)
        if doc is None:
            try:
                active = app.ActiveDocument
                if active is not None and paths_equal(str(active.FullName), path):
                    doc = active
            except Exception:  # noqa: BLE001
                pass

        if doc is not None:
            self._save_document_to_template(doc, path)
            try:
                doc.Close(SaveChanges=0)
            except Exception as exc:  # noqa: BLE001
                raise EasyWordError("close_failed", f"关闭文档失败: {exc}", status_code=500) from exc

        if self._state is not None and normalize_path(self._state.template_path) == path:
            self._state = None
        word_app.quit_if_no_documents()
        _log.info(f"Finalized document by path: {path}")

    def clear_without_close(self) -> None:
        self._state = None


document_binding = DocumentBinding()
