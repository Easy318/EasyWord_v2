"""图表 COM 辅助：外链拒绝、释放与 HRESULT 映射（不依赖 Word）。"""

from __future__ import annotations

import subprocess

import pytest

from word.core.content_control import chart_workbook as cw
from word.core.content_control.chart_workbook import (
    chart_data_is_linked,
    ensure_chart_embedded,
    map_chart_com_error,
    release_chart_workbook,
)
from word.runtime.errors import EasyWordError


@pytest.fixture(autouse=True)
def _fast_excel_release(monkeypatch) -> None:
    monkeypatch.setattr(cw, "_OLE_FLUSH_SEC", 0.0)
    monkeypatch.setattr(cw, "_QUIT_EXIT_WAIT_SEC", 0.0)


class _FakeLinkedChartData:
    IsLinked = True


class _FakeEmbeddedChartData:
    IsLinked = False


class _FakeWorkbooks:
    def __init__(self, count: int = 1) -> None:
        self.Count = count

    def __call__(self, index: int):  # noqa: ANN001
        raise AssertionError(f"unexpected Workbooks({index})")


class _FakeExcel:
    def __init__(self) -> None:
        self.Workbooks = _FakeWorkbooks(1)
        self.DisplayAlerts = True
        self.Visible = True
        self.quit_called = False
        self.calculate_called = False
        self.Hwnd = 0

    def Calculate(self) -> None:
        self.calculate_called = True

    def Quit(self) -> None:
        self.quit_called = True


class _FakeWorkbook:
    def __init__(self, app: _FakeExcel, events: list[str] | None = None) -> None:
        self.Application = app
        self.closed = False
        self._events = events

    def Close(self, SaveChanges: bool = True) -> None:  # noqa: N803
        self.closed = True
        self.Application.Workbooks.Count = 0
        if self._events is not None:
            self._events.append("close")


class _FakeChart:
    def __init__(self, events: list[str]) -> None:
        self._events = events

    def Refresh(self) -> None:
        self._events.append("refresh")


def test_ensure_chart_embedded_rejects_linked() -> None:
    try:
        ensure_chart_embedded(_FakeLinkedChartData())
        raise AssertionError("expected EasyWordError")
    except EasyWordError as exc:
        assert exc.code == "chart_linked"
        assert "嵌入" in exc.message


def test_ensure_chart_embedded_allows_embedded() -> None:
    ensure_chart_embedded(_FakeEmbeddedChartData())
    assert chart_data_is_linked(_FakeEmbeddedChartData()) is False


def test_map_disp_e_exception() -> None:
    class FakeComError(Exception):
        hresult = -2147352567

    err = map_chart_com_error(FakeComError("发生意外"))
    assert err.code == "chart_com_failed"
    assert "嵌入" in err.message or "Excel" in err.message


def test_map_link_message() -> None:
    err = map_chart_com_error(RuntimeError("链接的文件不可用"))
    assert err.code == "chart_link_broken"


def test_release_closes_and_quits_empty_excel() -> None:
    app = _FakeExcel()
    wb = _FakeWorkbook(app)
    drop = [object(), object()]
    release_chart_workbook(wb, excel=app, drop=drop)
    assert wb.closed is True
    assert app.DisplayAlerts is False
    assert app.Visible is False
    assert app.quit_called is True
    assert app.calculate_called is True
    assert drop == []


def test_release_does_not_quit_when_other_workbooks_open() -> None:
    app = _FakeExcel()
    wb = _FakeWorkbook(app)

    def close_keep_count(*, SaveChanges: bool = True) -> None:  # noqa: N803
        wb.closed = True
        app.Workbooks.Count = 2

    wb.Close = close_keep_count  # type: ignore[method-assign]
    release_chart_workbook(wb, excel=app)
    assert wb.closed is True
    assert app.quit_called is False


def test_release_refreshes_before_and_after_close() -> None:
    events: list[str] = []
    app = _FakeExcel()
    wb = _FakeWorkbook(app, events)
    chart = _FakeChart(events)
    release_chart_workbook(wb, chart, excel=app)
    assert events[:3] == ["refresh", "close", "refresh"]
    assert app.quit_called is True


def test_quit_watchdog_kills_when_quit_hangs(monkeypatch) -> None:
    app = _FakeExcel()
    app.Workbooks.Count = 0
    killed: list[int] = []

    def hang_quit() -> None:
        import time

        time.sleep(3.0)

    app.Quit = hang_quit  # type: ignore[method-assign]
    monkeypatch.setattr(cw, "_QUIT_GRACE_SEC", 0.2)
    monkeypatch.setattr(cw, "_kill_pid", lambda pid: killed.append(pid or -1))
    monkeypatch.setattr(cw, "_pid_alive", lambda _pid: True)

    cw._quit_or_kill_excel(app, pid=4242)
    assert 4242 in killed


def test_quit_does_not_kill_if_process_exits(monkeypatch) -> None:
    app = _FakeExcel()
    app.Workbooks.Count = 0
    killed: list[int] = []
    monkeypatch.setattr(cw, "_kill_pid", lambda pid: killed.append(pid or -1))
    monkeypatch.setattr(cw, "_pid_alive", lambda _pid: False)
    cw._quit_or_kill_excel(app, pid=4242)
    assert app.quit_called is True
    assert killed == []


def test_quit_kills_zombie_after_wait(monkeypatch) -> None:
    app = _FakeExcel()
    app.Workbooks.Count = 0
    killed: list[int] = []
    monkeypatch.setattr(cw, "_QUIT_GRACE_SEC", 5.0)
    monkeypatch.setattr(cw, "_QUIT_EXIT_WAIT_SEC", 0.05)
    monkeypatch.setattr(cw, "_kill_pid", lambda pid: killed.append(pid or -1))
    monkeypatch.setattr(cw, "_pid_alive", lambda _pid: True)
    cw._quit_or_kill_excel(app, pid=4242)
    assert app.quit_called is True
    assert killed == [4242]


def test_kill_pid_uses_terminate_process(monkeypatch) -> None:
    spawned: list[object] = []
    monkeypatch.setattr(cw.subprocess, "run", lambda *a, **k: spawned.append((a, k)))
    monkeypatch.setattr(cw, "_pid_alive", lambda _pid: True)

    class _Kernel:
        def OpenProcess(self, *_a, **_k):  # noqa: ANN001
            return 99

        def TerminateProcess(self, handle, code):  # noqa: ANN001
            assert handle == 99
            assert code == 1
            return 1

        def CloseHandle(self, handle):  # noqa: ANN001
            assert handle == 99
            return 1

    monkeypatch.setattr(cw.ctypes.windll, "kernel32", _Kernel())
    cw._kill_pid(4242)
    assert spawned == []


def test_taskkill_fallback_hides_console_window(monkeypatch) -> None:
    captured: dict[str, object] = {}

    def fake_run(*_args, **kwargs):  # noqa: ANN001
        captured.update(kwargs)

    monkeypatch.setattr(cw.subprocess, "run", fake_run)
    monkeypatch.setattr(cw, "_pid_alive", lambda _pid: True)
    monkeypatch.setattr(cw, "_terminate_via_win32", lambda _pid: False)
    cw._kill_pid(4242)
    assert captured.get("creationflags") == subprocess.CREATE_NO_WINDOW
    startupinfo = captured.get("startupinfo")
    assert startupinfo is not None
    assert startupinfo.dwFlags & subprocess.STARTF_USESHOWWINDOW
    assert startupinfo.wShowWindow == 0


def test_used_range_source_from_address() -> None:
    class _Ws:
        Name = "Sheet1"

    class _Used:
        Address = "$A$1:$D$5"

    assert cw._used_range_source(_Ws(), _Used()) == "'Sheet1'!$A$1:$D$5"


def test_used_range_source_strips_workbook_qualifier() -> None:
    class _Ws:
        Name = "Data"

    class _Used:
        Address = "[Book1]Data!$B$2:$C$10"

    assert cw._used_range_source(_Ws(), _Used()) == "'Data'!$B$2:$C$10"


def test_a1_col() -> None:
    assert cw._a1_col(1) == "A"
    assert cw._a1_col(26) == "Z"
    assert cw._a1_col(27) == "AA"


def test_sync_word_chart_cache_sets_source_and_retouches_formula() -> None:
    class _Series:
        def __init__(self) -> None:
            self.Formula = "=SERIES(Sheet1!$B$1,Sheet1!$A$2:$A$3,Sheet1!$B$2:$B$3,1)"
            self.writes: list[str] = []

        def __setattr__(self, name: str, value: object) -> None:
            if name == "Formula" and "Formula" in self.__dict__:
                self.writes.append(str(value))
            object.__setattr__(self, name, value)

    class _Coll:
        def __init__(self, series: _Series) -> None:
            self.Count = 1
            self._series = series

        def Item(self, index: int) -> _Series:
            assert index == 1
            return self._series

    series = _Series()

    class _Chart:
        def __init__(self) -> None:
            self.PlotBy = 2
            self.sources: list[tuple[str, object]] = []
            self.refreshed = 0

        def SetSourceData(self, Source: str, PlotBy: object = None) -> None:  # noqa: N803
            self.sources.append((Source, PlotBy))

        def SeriesCollection(self) -> _Coll:
            return _Coll(series)

        def Refresh(self) -> None:
            self.refreshed += 1

    class _Used:
        Address = "$A$1:$B$3"

    class _Excel:
        def Calculate(self) -> None:
            return None

    class _Ws:
        Name = "Sheet1"
        Application = _Excel()

    chart = _Chart()
    cw.sync_word_chart_cache(chart, _Ws(), _Used())
    assert chart.sources[0] == ("'Sheet1'!$A$1:$B$3", 2)
    assert series.writes == [
        "=SERIES(Sheet1!$B$1,Sheet1!$A$2:$A$3,Sheet1!$B$2:$B$3,1)"
    ]
    assert chart.refreshed == 1

