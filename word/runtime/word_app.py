"""Word.Application 单例：附着已有实例优先，安全 Quit。"""

from __future__ import annotations

from typing import Any

from word.logger import get_logger

_log = get_logger("runtime")

# wdAlertsNone
_WD_ALERTS_NONE = 0


class WordAppManager:
    def __init__(self) -> None:
        self._app: Any = None
        self._created_by_us = False

    @property
    def created_by_us(self) -> bool:
        return self._created_by_us

    def ensure_app(self) -> Any:
        """获取或创建 Word.Application；须在 COM STA 线程调用。"""
        if self._app is not None:
            try:
                _ = self._app.Name
                return self._app
            except Exception:  # noqa: BLE001
                _log.warning("Word.Application 句柄失效，将重新获取")
                self._app = None
                self._created_by_us = False

        import win32com.client

        try:
            self._app = win32com.client.GetActiveObject("Word.Application")
            self._created_by_us = False
            _log.info("Attached to existing Word.Application")
        except Exception:  # noqa: BLE001
            self._app = win32com.client.DispatchEx("Word.Application")
            self._created_by_us = True
            _log.info("Created new Word.Application")

        try:
            self._app.Visible = True
        except Exception:  # noqa: BLE001
            pass
        try:
            self._app.DisplayAlerts = _WD_ALERTS_NONE
        except Exception:  # noqa: BLE001
            pass
        return self._app

    def try_app(self) -> Any | None:
        if self._app is None:
            try:
                import win32com.client

                self._app = win32com.client.GetActiveObject("Word.Application")
                self._created_by_us = False
                return self._app
            except Exception:  # noqa: BLE001
                return None
        try:
            _ = self._app.Name
            return self._app
        except Exception:  # noqa: BLE001
            self._app = None
            self._created_by_us = False
            return None

    def document_count(self) -> int:
        app = self.try_app()
        if app is None:
            return 0
        try:
            return int(app.Documents.Count)
        except Exception:  # noqa: BLE001
            return 0

    def quit_if_no_documents(self) -> None:
        """无打开文档时退出 Word（项目关闭模板后使用）。"""
        app = self._app
        if app is None:
            app = self.try_app()
        if app is None:
            return
        try:
            count = int(app.Documents.Count)
        except Exception:  # noqa: BLE001
            return
        if count != 0:
            return
        try:
            app.Quit()
            _log.info("Word.Application Quit (no documents)")
        except Exception as exc:  # noqa: BLE001
            _log.warning(f"Word.Quit failed: {exc}")
        self._app = None
        self._created_by_us = False

    def shutdown(self) -> None:
        """进程退出时：无文档则 Quit。"""
        self.quit_if_no_documents()


word_app = WordAppManager()
