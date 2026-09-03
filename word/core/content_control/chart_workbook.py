"""图表 ChartData 工作簿：激活、外链检测、COM 错误映射、安全释放。"""

from __future__ import annotations

from typing import Any

from word.runtime.errors import EasyWordError

_DISP_E_EXCEPTION = -2147352567
_E_FAIL = -2147467259


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


def release_chart_workbook(wb: Any, chart: Any | None = None) -> None:
    """关闭数据簿并 Refresh 图表，否则 Word 视图可能仍显示旧图直到手动「编辑数据」。"""
    if wb is not None:
        try:
            wb.Close(SaveChanges=True)
        except Exception:  # noqa: BLE001
            pass
    if chart is not None:
        try:
            chart.Refresh()
        except Exception:  # noqa: BLE001
            pass


def open_chart_sheet(chart: Any) -> tuple[Any, Any, Any, Any]:
    """
    激活并打开图表数据工作表。
    返回 (chart_data, workbook, worksheet, used_range)。
    调用方须在 finally 中 release_chart_workbook(workbook, chart)。
    """
    try:
        chart_data = chart.ChartData
        ensure_chart_embedded(chart_data)
        activate_chart_workbook(chart_data)
        wb = chart_data.Workbook
        ws = wb.Worksheets(1)
        used = ws.UsedRange
        return chart_data, wb, ws, used
    except EasyWordError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_chart_com_error(exc, action="打开图表数据簿") from exc
