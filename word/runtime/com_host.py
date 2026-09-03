"""专用 STA 线程：所有 win32com 调用经此串行执行。"""

from __future__ import annotations

import queue
import threading
from concurrent.futures import Future
from typing import Callable, TypeVar

from word.logger import get_logger

T = TypeVar("T")
_log = get_logger("runtime")


class ComHost:
    def __init__(self) -> None:
        self._q: queue.Queue[tuple[Callable[[], object], Future[object]] | None] = queue.Queue()
        self._thread: threading.Thread | None = None
        self._started = False
        self._lock = threading.Lock()

    @property
    def started(self) -> bool:
        return self._started

    def start(self) -> None:
        with self._lock:
            if self._started:
                return
            self._thread = threading.Thread(target=self._loop, name="easyword-com-sta", daemon=True)
            self._thread.start()
            self._started = True
            _log.info("COM STA host started")

    def stop(self, timeout: float = 5.0) -> None:
        with self._lock:
            if not self._started:
                return
            self._q.put(None)
            t = self._thread
            self._started = False
            self._thread = None
        if t is not None:
            t.join(timeout=timeout)
            _log.info("COM STA host stopped")

    def submit(self, fn: Callable[[], T], timeout: float = 120.0) -> T:
        if not self._started:
            raise RuntimeError("COM host 未启动")
        fut: Future[object] = Future()
        self._q.put((fn, fut))
        return fut.result(timeout=timeout)  # type: ignore[return-value]

    def _loop(self) -> None:
        import pythoncom

        pythoncom.CoInitialize()
        try:
            while True:
                item = self._q.get()
                if item is None:
                    break
                fn, fut = item
                if fut.set_running_or_notify_cancel():
                    try:
                        fut.set_result(fn())
                    except Exception as exc:  # noqa: BLE001 — 传回调用方
                        fut.set_exception(exc)
        finally:
            pythoncom.CoUninitialize()


com_host = ComHost()
