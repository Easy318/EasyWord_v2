"""将 Word 主窗口贴到当前显示器工作区左半屏（近似 Win+←）。"""

from __future__ import annotations

from typing import Any

from word.logger import get_logger

_log = get_logger("runtime")

_SW_RESTORE = 9
_SWP_NOZORDER = 0x0004
_MONITOR_DEFAULTTONEAREST = 2
# Word 主框架窗口类名
_WORD_MAIN_CLASS = "OpusApp"


def _try_int_hwnd(value: Any) -> int | None:
    try:
        hwnd = int(value)
        return hwnd if hwnd else None
    except Exception:  # noqa: BLE001
        return None


def _resolve_word_hwnd(app: Any) -> int | None:
    """多路径获取 Word 主窗口句柄（兼容无 Application.Hwnd 的环境）。"""
    import win32gui

    # 1) Application.Hwnd（Word 2013+ 部分版本可用）
    try:
        hwnd = _try_int_hwnd(getattr(app, "Hwnd", None))
        if hwnd and win32gui.IsWindow(hwnd):
            return hwnd
    except Exception:  # noqa: BLE001
        pass

    # 2) ActiveWindow.Hwnd
    try:
        aw = app.ActiveWindow
        hwnd = _try_int_hwnd(getattr(aw, "Hwnd", None))
        if hwnd and win32gui.IsWindow(hwnd):
            # ActiveWindow.Hwnd 可能是文档子窗口，向上找到 OpusApp
            root = win32gui.GetAncestor(hwnd, 2)  # GA_ROOT = 2
            if root and win32gui.IsWindow(root):
                class_name = win32gui.GetClassName(root)
                if class_name == _WORD_MAIN_CLASS:
                    return int(root)
            return hwnd
    except Exception:  # noqa: BLE001
        pass

    # 3) 按窗口类名查找前台/可见的 Word 主窗口
    try:
        found: list[int] = []

        def _enum(hwnd: int, _: Any) -> bool:
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                if win32gui.GetClassName(hwnd) != _WORD_MAIN_CLASS:
                    return True
                found.append(int(hwnd))
            except Exception:  # noqa: BLE001
                pass
            return True

        win32gui.EnumWindows(_enum, None)
        if found:
            # 优先前台窗口，否则取第一个可见 OpusApp
            fg = win32gui.GetForegroundWindow()
            if fg in found:
                return int(fg)
            return found[0]
    except Exception as exc:  # noqa: BLE001
        _log.warning(f"EnumWindows 查找 Word 失败: {exc}")

    # 4) FindWindow 兜底
    try:
        hwnd = win32gui.FindWindow(_WORD_MAIN_CLASS, None)
        if hwnd and win32gui.IsWindow(hwnd):
            return int(hwnd)
    except Exception:  # noqa: BLE001
        pass

    return None


def focus_word_window(app: Any) -> None:
    """将 Word 主窗口恢复并置于前台（定位控件后使用）。"""
    try:
        import win32gui
    except ImportError:
        _log.warning("win32gui 不可用，跳过 Word 窗口聚焦")
        return

    hwnd = _resolve_word_hwnd(app)
    if hwnd:
        try:
            win32gui.ShowWindow(hwnd, _SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
        except Exception as exc:  # noqa: BLE001
            _log.warning(f"SetForegroundWindow 失败: {exc}")

    try:
        app.Activate()
    except Exception:  # noqa: BLE001
        pass


def snap_word_to_left_half(app: Any) -> None:
    """把 Word 应用窗口恢复并贴到所在显示器工作区左半侧。失败仅记日志，不影响打开文档。"""
    try:
        import win32api
        import win32gui
    except ImportError:
        _log.warning("win32api/win32gui 不可用，跳过左半屏布局")
        return

    hwnd = _resolve_word_hwnd(app)
    if not hwnd:
        _log.warning("无法定位 Word 主窗口句柄，跳过左半屏布局")
        return

    try:
        win32gui.ShowWindow(hwnd, _SW_RESTORE)
        try:
            app.ActiveWindow.WindowState = 0  # wdWindowStateNormal
        except Exception:  # noqa: BLE001
            pass

        monitor = win32api.MonitorFromWindow(hwnd, _MONITOR_DEFAULTTONEAREST)
        info = win32api.GetMonitorInfo(monitor)
        work_left, work_top, work_right, work_bottom = info["Work"]
        width = max(1, (work_right - work_left) // 2)
        height = max(1, work_bottom - work_top)

        win32gui.SetWindowPos(
            hwnd,
            0,
            int(work_left),
            int(work_top),
            int(width),
            int(height),
            _SWP_NOZORDER,
        )
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception:  # noqa: BLE001
            pass
        _log.info(
            f"Word snapped left-half hwnd={hwnd} "
            f"work=({work_left},{work_top})-({work_right},{work_bottom})"
        )
    except Exception as exc:  # noqa: BLE001
        _log.warning(f"左半屏布局失败: {exc}")
