"""图表 COM 辅助：外链拒绝与 HRESULT 映射（不依赖 Word）。"""

from __future__ import annotations

from word.core.content_control.chart_workbook import (
    chart_data_is_linked,
    ensure_chart_embedded,
    map_chart_com_error,
)
from word.runtime.errors import EasyWordError


class _FakeLinkedChartData:
    IsLinked = True


class _FakeEmbeddedChartData:
    IsLinked = False


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
