"""图表 ChartData 工作簿：激活、外链检测、COM 错误映射、安全释放。"""

from __future__ import annotations

import ctypes
import gc
import subprocess
import threading
import time
from contextlib import contextmanager
from typing import Any, Iterator

from word.logger import get_logger
from word.runtime.errors import EasyWordError

_log = get_logger("chart_workbook")

_DISP_E_EXCEPTION = -2147352567
_E_FAIL = -2147467259
# Close 后 Word 需把 Excel 数据写入 chart.xml 缓存；立刻杀进程会导致图表面不刷新
_OLE_FLUSH_SEC = 0.5
# Quit 被加载项卡住时的强杀超时（须大于 OLE flush，避免打断缓存提交）
_QUIT_GRACE_SEC = 3.0
# Quit 返回后等进程退出；仍存活再静默结束空白壳
_QUIT_EXIT_WAIT_SEC = 1.0
_CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
_PROCESS_TERMINATE = 0x0001
_PROCESS_QUERY_LIMITED_INFORMATION = 0x1000


def map_chart_com_error(exc: BaseException, *, action: str = "访问图表数据") -> EasyWordError:
    """把 pywin32/COM 异常转成可读 EasyWordError。"""
    if isinstance(exc, EasyWordError):
        return exc
    raw = str(exc)
    lower = raw.lower()
    hresult: int | None = getattr(exc, "hresult", None)
    if hresult is None and getattr(exc, "args", None):
        arg0 = exc.args[0]
        if isinstance(arg0, int):
            hresult = arg0

    link_hints = (
        "链接",
        "link",
        "linked",
        "不可用",
        "cannot be accessed",
        "unavailable",
    )
    if any(h in raw or h in lower for h in link_hints):
        return EasyWordError(
            "chart_link_broken",
            "图表数据源不可用（可能是链接已断开，请将图表改为「嵌入数据」后重试）",
        )
    if hresult in (_DISP_E_EXCEPTION, _E_FAIL) or str(_DISP_E_EXCEPTION) in raw:
        return EasyWordError(
            "chart_com_failed",
            "Excel 未响应或图表数据簿无法打开，请确认图表为嵌入数据后重试",
        )
    return EasyWordError("chart_com_failed", f"{action}失败: {exc}")


def chart_data_is_linked(chart_data: Any) -> bool:
    try:
        return bool(chart_data.IsLinked)
    except Exception:  # noqa: BLE001
        return False


def activate_chart_workbook(chart_data: Any) -> None:
    """写前激活数据簿（Word 侧常用 Activate，部分环境为 ActivateWorkbook）。"""
    for name in ("Activate", "ActivateWorkbook"):
        fn = getattr(chart_data, name, None)
        if callable(fn):
            try:
                fn()
                return
            except Exception:  # noqa: BLE001
                continue


def ensure_chart_embedded(chart_data: Any) -> None:
    if chart_data_is_linked(chart_data):
        raise EasyWordError(
            "chart_linked",
            "当前图表使用外部链接数据，预览/写回仅支持「嵌入数据」。"
            "请在 Word 中打开图表「编辑数据」，改为嵌入后再试。",
        )


def _safe_workbook_count(excel: Any) -> int | None:
    try:
        return int(excel.Workbooks.Count)
    except Exception:  # noqa: BLE001
        return None


def _excel_pid(excel: Any) -> int | None:
    """从 Excel.Application.Hwnd 取进程 PID；失败返回 None。"""
    if excel is None:
        return None
    try:
        hwnd = int(excel.Hwnd)
    except Exception:  # noqa: BLE001
        return None
    if hwnd <= 0:
        return None
    try:
        import win32process

        _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
        return int(pid) if pid else None
    except Exception:  # noqa: BLE001
        return None


def _pid_alive(pid: int) -> bool:
    try:
        handle = ctypes.windll.kernel32.OpenProcess(
            _PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid)
        )
        if not handle:
            return False
        ctypes.windll.kernel32.CloseHandle(handle)
        return True
    except Exception:  # noqa: BLE001
        return False


def _pump_com(seconds: float) -> None:
    """在 STA 上泵送消息，让 Word 完成 ChartData OLE 回写（不可用 sleep 代替）。"""
    if seconds <= 0:
        return
    try:
        import pythoncom
    except Exception:  # noqa: BLE001
        time.sleep(seconds)
        return
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:  # noqa: BLE001
            pass
        time.sleep(0.02)


def _wait_pid_exit(pid: int | None, timeout_sec: float) -> bool:
    if pid is None or pid <= 0:
        return True
    deadline = time.monotonic() + max(timeout_sec, 0.0)
    while True:
        if not _pid_alive(pid):
            return True
        if time.monotonic() >= deadline:
            return not _pid_alive(pid)
        _pump_com(0.05)


def _terminate_via_win32(pid: int) -> bool:
    """TerminateProcess 不拉起控制台，避免打包后闪 CMD 黑窗。"""
    handle = ctypes.windll.kernel32.OpenProcess(_PROCESS_TERMINATE, False, int(pid))
    if not handle:
        return False
    try:
        return bool(ctypes.windll.kernel32.TerminateProcess(handle, 1))
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)


def _kill_via_taskkill_hidden(pid: int) -> None:
    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = 0
    subprocess.run(
        ["taskkill", "/F", "/PID", str(pid)],
        check=False,
        capture_output=True,
        text=True,
        timeout=8,
        creationflags=_CREATE_NO_WINDOW,
        startupinfo=si,
    )


def _kill_pid(pid: int | None) -> None:
    """强杀 ChartData 拉起的空 Excel（Quit 被加载项卡住时的兜底）。"""
    if pid is None or pid <= 0:
        return
    if not _pid_alive(pid):
        return
    try:
        if _terminate_via_win32(pid):
            _log.info(f"force-killed ChartData Excel pid={pid}")
            return
        _kill_via_taskkill_hidden(pid)
        _log.info(f"force-killed ChartData Excel pid={pid} via taskkill")
    except Exception as exc:  # noqa: BLE001
        _log.warning(f"kill Excel pid={pid} failed: {exc}")


def _hide_excel(excel: Any) -> None:
    if excel is None:
        return
    for name, value in (
        ("Visible", False),
        ("DisplayAlerts", False),
        ("ScreenUpdating", False),
        ("EnableEvents", False),
        ("AskToUpdateLinks", False),
    ):
        try:
            setattr(excel, name, value)
        except Exception:  # noqa: BLE001
            pass


def _refresh_chart(chart: Any | None) -> None:
    if chart is None:
        return
    try:
        chart.Refresh()
    except Exception:  # noqa: BLE001
        pass


def _a1_col(index: int) -> str:
    if index <= 0:
        return "A"
    out: list[str] = []
    n = int(index)
    while n:
        n, rem = divmod(n - 1, 26)
        out.append(chr(ord("A") + rem))
    return "".join(reversed(out))


def _used_range_source(ws: Any, used: Any) -> str | None:
    """Word.SetSourceData 要的是地址字符串，例如 'Sheet1'!$A$1:$D$5（不是 Excel Range）。"""
    sheet = "Sheet1"
    try:
        sheet = str(ws.Name or sheet)
    except Exception:  # noqa: BLE001
        pass
    addr = ""
    for getter in (
        lambda: str(used.Address or ""),
        lambda: str(used.GetAddress(True, True, 1, False) or ""),
    ):
        try:
            addr = getter().strip()
        except Exception:  # noqa: BLE001
            addr = ""
        if addr:
            break
    if not addr:
        try:
            rows = int(used.Rows.Count)
            cols = int(used.Columns.Count)
            addr = f"$A$1:${_a1_col(cols)}${rows}"
        except Exception:  # noqa: BLE001
            return None
    if "!" in addr:
        sheet_part, addr = addr.rsplit("!", 1)
        if "]" in sheet_part:
            sheet_part = sheet_part.split("]", 1)[-1]
        sheet = sheet_part.strip("'") or sheet
    quoted = sheet.replace("'", "''")
    return f"'{quoted}'!{addr}"


def _set_source_data(chart: Any, source: str) -> bool:
    plot_by = None
    try:
        plot_by = chart.PlotBy
    except Exception:  # noqa: BLE001
        plot_by = None
    sources = [source]
    if source.startswith("="):
        sources.append(source[1:])
    else:
        sources.append("=" + source)
    attempts: list[tuple[str, Any]] = []
    for src in sources:
        if plot_by is not None:
            attempts.append((src, plot_by))
        attempts.append((src, None))
    for src, pb in attempts:
        try:
            if pb is not None:
                chart.SetSourceData(Source=src, PlotBy=pb)
            else:
                chart.SetSourceData(Source=src)
            return True
        except Exception:  # noqa: BLE001
            continue
    return False


def _iter_series(chart: Any) -> Iterator[Any]:
    for name in ("FullSeriesCollection", "SeriesCollection"):
        fn = getattr(chart, name, None)
        if not callable(fn):
            continue
        try:
            coll = fn()
            count = int(coll.Count)
        except Exception:  # noqa: BLE001
            continue
        for i in range(1, count + 1):
            try:
                yield coll.Item(i)
            except Exception:  # noqa: BLE001
                try:
                    yield fn(i)
                except Exception:  # noqa: BLE001
                    continue
        return


def _retouch_series_formulas(chart: Any) -> None:
    """重写 SERIES 公式，迫使 Word 丢掉旧 numCache（Refresh 做不到这一点）。"""
    for series in _iter_series(chart):
        old = None
        try:
            old = series.Formula
        except Exception:  # noqa: BLE001
            continue
        if not old:
            continue
        try:
            series.Formula = old
        except Exception:  # noqa: BLE001
            try:
                series.Formula = "=SERIES(,,1,1)"
                series.Formula = old
            except Exception:  # noqa: BLE001
                try:
                    series.Formula = old
                except Exception:  # noqa: BLE001
                    pass


def sync_word_chart_cache(chart: Any | None, ws: Any, used: Any) -> None:
    """
    在 ChartData 工作簿仍打开时，把 Excel 格子同步进 Word 图表面缓存。

    Word.Chart.Refresh 只按已有 c:numCache 重绘，不会从 Excel 拉数；
    手动「编辑数据」之所以有效，是因为 UI 会走 SetSourceData 同类重绑。
    必须在 Close/Quit 之前调用。
    """
    if chart is None:
        return
    try:
        ws.Application.Calculate()
    except Exception:  # noqa: BLE001
        pass
    source = _used_range_source(ws, used)
    if source:
        if not _set_source_data(chart, source):
            _log.warning(f"SetSourceData failed for {source}")
    _retouch_series_formulas(chart)
    try:
        chart.Application.ScreenUpdating = True
    except Exception:  # noqa: BLE001
        pass
    _refresh_chart(chart)


def _reclose_chart_data_if_reopened(app: Any, count_after_close: int | None) -> None:
    """Close 后再 Refresh 可能把 ChartData 重新打开；仅在原本已空时关掉这本簿。"""
    if app is None or count_after_close != 0:
        return
    if _safe_workbook_count(app) != 1:
        return
    try:
        app.Workbooks(1).Close(SaveChanges=True)
    except Exception:  # noqa: BLE001
        pass
    _hide_excel(app)


def _quit_or_kill_excel(app: Any, pid: int | None) -> None:
    """
    先优雅 Quit，等进程退出；仅空白壳残留或 Quit 被加载项拖死时才强杀。

    不可在 Close 后立刻杀进程：Word 图表显示的是 chart.xml 缓存，OLE 尚未提交完
    就被掐掉时，图表面不更新，直到用户手动「编辑数据」才会重建缓存。
    """
    count = _safe_workbook_count(app)
    if count not in (None, 0):
        # 用户自己的簿还在同一实例里，绝不动 Quit/杀进程
        return

    stop = threading.Event()

    def _watchdog() -> None:
        if not stop.wait(_QUIT_GRACE_SEC):
            _kill_pid(pid)

    watcher = threading.Thread(
        target=_watchdog, name="excel-chart-quit-watchdog", daemon=True
    )
    watcher.start()
    try:
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            pass
    finally:
        stop.set()
        watcher.join(timeout=1.0)

    if _wait_pid_exit(pid, _QUIT_EXIT_WAIT_SEC):
        return
    if _safe_workbook_count(app) not in (None, 0):
        return
    _kill_pid(pid)
    _pump_com(0.1)


def release_chart_workbook(
    wb: Any,
    chart: Any | None = None,
    *,
    excel: Any | None = None,
    drop: list[Any] | None = None,
    excel_pid: int | None = None,
) -> None:
    """
    关闭 ChartData 工作簿并退出为其拉起的 Excel。

    图表面缓存必须在关簿前由 sync_word_chart_cache 写入；此处只负责 Close + Quit。
    强杀只作为空白壳/加载项卡死的兜底，且不得使用会弹黑窗的 taskkill。
    """
    app = excel
    if app is None and wb is not None:
        try:
            app = wb.Application
        except Exception:  # noqa: BLE001
            app = None

    pid = excel_pid if excel_pid is not None else _excel_pid(app)
    _hide_excel(app)

    if app is not None:
        try:
            app.Calculate()
        except Exception:  # noqa: BLE001
            pass

    # 先 Refresh（簿仍打开时更稳），再 Close；图表面缓存须由 sync_word_chart_cache 在关簿前写入
    if chart is not None:
        try:
            chart.Refresh()
        except Exception:  # noqa: BLE001
            pass

    if wb is not None:
        try:
            # Close(SaveChanges=True) 把嵌入 xlsx 写回 Word；图表面缓存已在 sync 时写入
            wb.Close(SaveChanges=True)
        except Exception:  # noqa: BLE001
            pass

    count_after_close = _safe_workbook_count(app)
    # 给 Word STA 处理 OLE OnClose，把内嵌簿同步进 c:numCache / c:strCache
    _pump_com(_OLE_FLUSH_SEC)

    # 关簿后再 Refresh（Word VBA 推荐顺序），缓存才稳定画到图表面
    _hide_excel(app)
    _refresh_chart(chart)
    _pump_com(min(_OLE_FLUSH_SEC, 0.2) if _OLE_FLUSH_SEC > 0 else 0.0)
    _reclose_chart_data_if_reopened(app, count_after_close)

    if drop:
        drop.clear()

    if app is not None:
        _quit_or_kill_excel(app, pid)
    else:
        _kill_pid(pid)

    try:
        gc.collect()
        gc.collect()
    except Exception:  # noqa: BLE001
        pass


def open_chart_sheet(chart: Any) -> tuple[Any, Any, Any, Any]:
    """
    激活并打开图表数据工作表。
    返回 (chart_data, workbook, worksheet, used_range)。
    调用方须在 finally 中 release_chart_workbook(...)；推荐用 chart_sheet_session。
    """
    try:
        chart_data = chart.ChartData
        ensure_chart_embedded(chart_data)
        activate_chart_workbook(chart_data)
        wb = chart_data.Workbook
        try:
            _hide_excel(wb.Application)
        except Exception:  # noqa: BLE001
            pass
        ws = wb.Worksheets(1)
        used = ws.UsedRange
        return chart_data, wb, ws, used
    except EasyWordError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_chart_com_error(exc, action="打开图表数据簿") from exc


@contextmanager
def chart_sheet_session(
    chart: Any, *, sync_visual: bool = True
) -> Iterator[tuple[Any, Any, Any, Any]]:
    """打开图表数据表；写回后先重绑 Word 缓存，再关闭并退出/强杀空 Excel。"""
    chart_data, wb, ws, used = open_chart_sheet(chart)
    excel = None
    excel_pid = None
    try:
        excel = wb.Application
        _hide_excel(excel)
        excel_pid = _excel_pid(excel)
    except Exception:  # noqa: BLE001
        excel = None
    drop: list[Any] = [chart_data, ws, used]
    try:
        yield chart_data, wb, ws, used
    finally:
        if sync_visual:
            try:
                sync_word_chart_cache(chart, ws, used)
            except Exception as exc:  # noqa: BLE001
                _log.warning(f"sync Word chart cache failed: {exc}")
        release_chart_workbook(
            wb, chart, excel=excel, drop=drop, excel_pid=excel_pid
        )
