"""
Word 内嵌图表 L1 同步模块（同结构换值）

用法:
    1. 将含内嵌图表的 .docx 放到 word/test/test_data/template.docx
    2. python -m word.test.test

模块会原地修改内嵌 Excel 与 chart.xml 缓存，保留原始 OOXML 命名空间结构。
"""

from __future__ import annotations

import io
import posixpath
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from word.config.path import TEST_DATA_DIR

# ---------------------------------------------------------------------------
# OOXML 命名空间
# ---------------------------------------------------------------------------
NS_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
REL_PACKAGE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
)

ET.register_namespace("c", NS_CHART)
ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace(
    "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)

_CELL_RE = re.compile(r"^([A-Za-z]+)(\d+)$")
_FORMULA_RE = re.compile(
    r"^(?:(?:'([^']+)')|([^!]+))!"
    r"(\$?)([A-Za-z]+)(\$?)(\d+)"
    r"(?::(\$?)([A-Za-z]+)(\$?)(\d+))?$"
)


# ---------------------------------------------------------------------------
# 公式 / 单元格工具
# ---------------------------------------------------------------------------
def _col_to_index(col: str) -> int:
    idx = 0
    for ch in col.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def _index_to_col(index: int) -> str:
    result: list[str] = []
    while index:
        index, rem = divmod(index - 1, 26)
        result.append(chr(rem + ord("A")))
    return "".join(reversed(result))


@dataclass(frozen=True)
class CellRef:
    col: int
    row: int

    @classmethod
    def parse(cls, ref: str) -> CellRef:
        m = _CELL_RE.match(ref.replace("$", ""))
        if not m:
            raise ValueError(f"无效单元格引用: {ref}")
        return cls(_col_to_index(m.group(1)), int(m.group(2)))

    def to_a1(self) -> str:
        return f"{_index_to_col(self.col)}{self.row}"


@dataclass(frozen=True)
class FormulaRange:
    sheet: str
    start: CellRef
    end: CellRef

    @property
    def key(self) -> str:
        s, e = self.start.to_a1(), self.end.to_a1()
        if s == e:
            return f"{self.sheet}!{s}"
        return f"{self.sheet}!{s}:{e}"

    def cells_row_major(self) -> list[CellRef]:
        cells: list[CellRef] = []
        r0, r1 = sorted((self.start.row, self.end.row))
        c0, c1 = sorted((self.start.col, self.end.col))
        for row in range(r0, r1 + 1):
            for col in range(c0, c1 + 1):
                cells.append(CellRef(col, row))
        return cells


def parse_formula(formula: str) -> FormulaRange:
    formula = formula.strip()
    m = _FORMULA_RE.match(formula)
    if not m:
        raise ValueError(f"无法解析图表公式: {formula}")
    sheet = (m.group(1) or m.group(2)).strip()
    start = CellRef.parse(f"{m.group(4)}{m.group(6)}")
    if m.group(8):
        end = CellRef.parse(f"{m.group(8)}{m.group(10)}")
    else:
        end = start
    return FormulaRange(sheet, start, end)


def normalize_formula_key(formula: str) -> str:
    return parse_formula(formula).key


# ---------------------------------------------------------------------------
# 关系 / 包路径
# ---------------------------------------------------------------------------
def _resolve_part_path(base_part: str, target: str) -> str:
    base_dir = posixpath.dirname(base_part)
    return posixpath.normpath(posixpath.join(base_dir, target)).replace("\\", "/")


def _parse_relationships(rels_xml: bytes) -> dict[str, tuple[str, str]]:
    root = ET.fromstring(rels_xml)
    rels: dict[str, tuple[str, str]] = {}
    for rel in root:
        if rel.tag.split("}")[-1] != "Relationship":
            continue
        rel_id = rel.get("Id", "")
        rel_type = rel.get("Type", "")
        target = rel.get("Target", "")
        rels[rel_id] = (rel_type, target)
    return rels


def _find_chart_parts(parts: dict[str, bytes]) -> list[str]:
    rels_path = "word/_rels/document.xml.rels"
    if rels_path not in parts:
        return []
    charts: list[str] = []
    for rel_type, target in _parse_relationships(parts[rels_path]).values():
        if rel_type == REL_CHART:
            charts.append(_resolve_part_path("word/document.xml", target))
    return sorted(set(charts))


def _find_embedded_workbook(parts: dict[str, bytes], chart_part: str) -> str | None:
    chart_name = posixpath.basename(chart_part)
    rels_path = f"word/charts/_rels/{chart_name}.rels"
    if rels_path not in parts:
        return None
    for rel_type, target in _parse_relationships(parts[rels_path]).values():
        if rel_type == REL_PACKAGE:
            return _resolve_part_path(chart_part, target)
    return None


# ---------------------------------------------------------------------------
# Excel 读写（内嵌 xlsx 原地补丁）
# ---------------------------------------------------------------------------
def _format_cache_value(value: Any, is_numeric: bool) -> str:
    if value is None:
        return "0" if is_numeric else ""
    if is_numeric:
        if isinstance(value, bool):
            return "1" if value else "0"
        if isinstance(value, (int, float)):
            if isinstance(value, float) and value.is_integer():
                return str(int(value))
            return str(value)
        return str(value)
    return str(value)


# ---------------------------------------------------------------------------
# 内嵌 xlsx 原地补丁（保留 sharedStrings 等结构，避免 openpyxl 整包重写）
# ---------------------------------------------------------------------------
def _unpack_zip(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def _repack_zip(original: bytes, parts: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(original), "r") as zin:
        with zipfile.ZipFile(buf, "w") as zout:
            for item in zin.infolist():
                zout.writestr(item, parts[item.filename])
    return buf.getvalue()


def _sanitize_workbook_xml(workbook_xml: bytes) -> bytes:
    """移除会导致 Word「编辑数据」时查找外部链接的 Excel 元数据。"""
    text = workbook_xml.decode("utf-8")
    text = re.sub(
        r"<mc:AlternateContent\b[^>]*>\s*"
        r"<mc:Choice\b[^>]*Requires=\"x15\"[^>]*>\s*"
        r"<x15ac:absPath\b[^>]*/>\s*"
        r"</mc:Choice>\s*"
        r"</mc:AlternateContent>",
        "",
        text,
        flags=re.DOTALL,
    )
    text = re.sub(
        r"<extLst>\s*<ext uri=\"\{140A7094-0E35-4892-8432-C4D2E57EDEB5\}\"[^>]*>\s*"
        r"<x15:workbookPr chartTrackingRefBase=\"1\"/>\s*</ext>\s*</extLst>",
        "",
        text,
        flags=re.DOTALL,
    )
    text = re.sub(r"<bookViews>.*?</bookViews>\s*", "", text, flags=re.DOTALL)
    return text.encode("utf-8")


def _sanitize_word_settings(settings_xml: bytes) -> bytes:
    """关闭 Word 图表按外部引用跟踪，避免编辑时查找链接文件。"""
    text = settings_xml.decode("utf-8")
    text = re.sub(r"<w15:chartTrackingRefBased\b[^>]*/>", "", text)
    return text.encode("utf-8")


def _sanitize_embedded_xlsx(xlsx_bytes: bytes) -> bytes:
    parts = _unpack_zip(xlsx_bytes)
    if "xl/workbook.xml" in parts:
        parts["xl/workbook.xml"] = _sanitize_workbook_xml(parts["xl/workbook.xml"])
    return _repack_zip(xlsx_bytes, parts)


def _parse_shared_strings(sst_xml: bytes) -> list[str]:
    if not sst_xml:
        return []
    texts: list[str] = []
    for block in re.findall(r"<si\b[^>]*>.*?</si>", sst_xml.decode("utf-8"), re.DOTALL):
        m = re.search(r"<t(?:\s+xml:space=\"preserve\")?>([^<]*)</t>", block)
        texts.append(m.group(1) if m else "")
    return texts


def _patch_shared_string_at_index(sst_xml: bytes, index: int, text: str) -> bytes:
    xml = sst_xml.decode("utf-8")
    blocks = list(re.finditer(r"(<si\b[^>]*>.*?</si>)", xml, re.DOTALL))
    if index >= len(blocks):
        raise IndexError(f"sharedStrings 索引越界: {index}")
    block = blocks[index].group(1)
    escaped = (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )
    if re.search(r"<t(?:\s+xml:space=\"preserve\")?>", block):
        new_block = re.sub(
            r"<t(?:\s+xml:space=\"preserve\")?>([^<]*)</t>",
            f"<t>{escaped}</t>",
            block,
            count=1,
        )
    else:
        new_block = re.sub(r"(</si>)", f"<t>{escaped}</t>\\1", block, count=1)
    return (
        xml[: blocks[index].start()] + new_block + xml[blocks[index].end() :]
    ).encode("utf-8")


def _load_xlsx_cell_values(xlsx_bytes: bytes) -> dict[str, Any]:
    parts = _unpack_zip(xlsx_bytes)
    strings = _parse_shared_strings(parts.get("xl/sharedStrings.xml", b""))
    sheet = parts["xl/worksheets/sheet1.xml"].decode("utf-8")
    values: dict[str, Any] = {}
    for m in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*)>(.*?)</c>', sheet, re.DOTALL):
        ref, attrs, inner = m.group(1), m.group(2), m.group(3)
        v_match = re.search(r"<v>([^<]*)</v>", inner)
        if not v_match:
            continue
        raw = v_match.group(1)
        t_match = re.search(r't="([^"]+)"', attrs)
        if t_match and t_match.group(1) == "s":
            values[ref] = strings[int(raw)]
        elif t_match and t_match.group(1) == "inlineStr":
            is_match = re.search(r"<t>([^<]*)</t>", inner)
            values[ref] = is_match.group(1) if is_match else ""
        else:
            try:
                values[ref] = float(raw) if "." in raw else int(raw)
            except ValueError:
                values[ref] = raw
    return values


def _read_range_values_from_xlsx(xlsx_bytes: bytes, formula: str) -> list[Any]:
    fr = parse_formula(formula)
    cell_map = _load_xlsx_cell_values(xlsx_bytes)
    return [cell_map.get(cell.to_a1()) for cell in fr.cells_row_major()]


def _patch_sheet_cell(sheet_xml: str, ref: str, value: Any) -> tuple[str, int | None]:
    """返回 (更新后的 sheet, 需同步的 sharedString 索引或 None)。"""
    pattern = re.compile(
        rf'(<c r="{ref}"([^>]*)>)(.*?)(</c>)',
        re.DOTALL,
    )
    m = pattern.search(sheet_xml)
    if not m:
        raise ValueError(f"sheet1 中找不到单元格: {ref}")

    open_tag, attrs, inner, close_tag = m.group(1), m.group(2), m.group(3), m.group(4)
    t_match = re.search(r't="([^"]+)"', attrs)

    if isinstance(value, str):
        if t_match and t_match.group(1) == "s":
            v_match = re.search(r"<v>(\d+)</v>", inner)
            if not v_match:
                raise ValueError(f"单元格 {ref} 缺少 sharedString 索引")
            # sheet 仍指向原索引，仅更新 sharedStrings.xml 中对应文本
            return sheet_xml, int(v_match.group(1))
        escaped = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        new_inner = f"<v>{escaped}</v>" if not t_match else inner
        if t_match and t_match.group(1) == "inlineStr":
            new_inner = re.sub(r"<t>([^<]*)</t>", f"<t>{escaped}</t>", inner, count=1)
        return sheet_xml[: m.start()] + open_tag + new_inner + close_tag + sheet_xml[
            m.end() :
        ], None

    num_text = _format_cache_value(value, True)
    if re.search(r"<v>[^<]*</v>", inner):
        new_inner = re.sub(r"<v>[^<]*</v>", f"<v>{num_text}</v>", inner, count=1)
    else:
        new_inner = inner + f"<v>{num_text}</v>"
    return sheet_xml[: m.start()] + open_tag + new_inner + close_tag + sheet_xml[
        m.end() :
    ], None


def _patch_table_column_name(table_xml: str, col_id: int, name: str) -> str:
    escaped = (
        name.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )
    pattern = re.compile(rf'(<tableColumn id="{col_id}"[^>]*\bname=")([^"]*)(")')
    if not pattern.search(table_xml):
        return table_xml
    return pattern.sub(rf"\1{escaped}\3", table_xml, count=1)


def _patch_embedded_xlsx(xlsx_bytes: bytes, values: dict[str, Any]) -> bytes:
    parts = _unpack_zip(xlsx_bytes)
    sheet = parts["xl/worksheets/sheet1.xml"].decode("utf-8")
    sst = parts.get("xl/sharedStrings.xml", b"")
    table = (
        parts.get("xl/tables/table1.xml", b"").decode("utf-8")
        if "xl/tables/table1.xml" in parts
        else ""
    )
    sst_patches: dict[int, str] = {}
    header_col_ids = {"B1": 2, "C1": 3, "D1": 4}

    for key, value in values.items():
        norm_key = normalize_formula_key(key) if "!" in key else f"Sheet1!{key}"
        fr = parse_formula(norm_key)
        if isinstance(value, list):
            cells = fr.cells_row_major()
            if len(value) != len(cells):
                raise ValueError(
                    f"区域 {norm_key} 需要 {len(cells)} 个值，实际 {len(value)} 个"
                )
            for cell, cell_value in zip(cells, value):
                sheet, sst_idx = _patch_sheet_cell(sheet, cell.to_a1(), cell_value)
                if sst_idx is not None and isinstance(cell_value, str):
                    sst_patches[sst_idx] = cell_value
        else:
            ref = fr.start.to_a1()
            sheet, sst_idx = _patch_sheet_cell(sheet, ref, value)
            if sst_idx is not None and isinstance(value, str):
                sst_patches[sst_idx] = value
                if table and ref in header_col_ids:
                    table = _patch_table_column_name(table, header_col_ids[ref], value)

    parts["xl/worksheets/sheet1.xml"] = sheet.encode("utf-8")
    if sst:
        for idx, text in sorted(sst_patches.items()):
            sst = _patch_shared_string_at_index(sst, idx, text)
        parts["xl/sharedStrings.xml"] = sst
    if table:
        parts["xl/tables/table1.xml"] = table.encode("utf-8")
    return _sanitize_embedded_xlsx(_repack_zip(xlsx_bytes, parts))


# ---------------------------------------------------------------------------
# chart.xml cache 更新（仅替换 cache 节点，保留原始命名空间/结构）
# ---------------------------------------------------------------------------
def _escape_xml(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _build_str_cache_inner(values: list[Any]) -> str:
    lines = [f'<c:ptCount val="{len(values)}" />']
    for idx, value in enumerate(values):
        text = _escape_xml("" if value is None else str(value))
        lines.append(f'<c:pt idx="{idx}"><c:v>{text}</c:v></c:pt>')
    return "".join(lines)


def _build_num_cache_inner(values: list[Any], format_code: str) -> str:
    lines = [f"<c:formatCode>{format_code}</c:formatCode>"]
    lines.append(f'<c:ptCount val="{len(values)}" />')
    for idx, value in enumerate(values):
        text = _escape_xml(_format_cache_value(value, True))
        lines.append(f'<c:pt idx="{idx}"><c:v>{text}</c:v></c:pt>')
    return "".join(lines)


_STR_CACHE_RE = re.compile(
    r"(<c:strRef>.*?<c:f>(?P<formula>[^<]+)</c:f>.*?<c:strCache[^>]*>)"
    r"(?P<inner>.*?)"
    r"(?P<close></c:strCache>)",
    re.DOTALL,
)
_NUM_CACHE_RE = re.compile(
    r"(<c:numRef>.*?<c:f>(?P<formula>[^<]+)</c:f>.*?<c:numCache[^>]*>)"
    r"(?P<inner>.*?)"
    r"(?P<close></c:numCache>)",
    re.DOTALL,
)


def _assert_well_formed_chart_xml(chart_xml: bytes) -> None:
    try:
        ET.fromstring(chart_xml)
    except ET.ParseError as exc:
        raise ValueError(f"chart.xml 结构损坏: {exc}") from exc


def _refresh_chart_caches(chart_xml: bytes, xlsx_bytes: bytes) -> bytes:
    text = chart_xml.decode("utf-8")

    def _replace_str(m: re.Match[str]) -> str:
        formula = m.group("formula").strip()
        values = _read_range_values_from_xlsx(xlsx_bytes, formula)
        inner = _build_str_cache_inner(values)
        return m.group(1) + inner + m.group("close")

    def _replace_num(m: re.Match[str]) -> str:
        formula = m.group("formula").strip()
        values = _read_range_values_from_xlsx(xlsx_bytes, formula)
        fc_match = re.search(r"<c:formatCode>([^<]*)</c:formatCode>", m.group("inner"))
        format_code = fc_match.group(1) if fc_match else "General"
        inner = _build_num_cache_inner(values, format_code)
        return m.group(1) + inner + m.group("close")

    text = _STR_CACHE_RE.sub(_replace_str, text)
    text = _NUM_CACHE_RE.sub(_replace_num, text)
    result = text.encode("utf-8")
    _assert_well_formed_chart_xml(result)
    return result


# ---------------------------------------------------------------------------
# 主类
# ---------------------------------------------------------------------------
class WordChartSync:
    """L1：同结构换值 — 写内嵌 Excel 并刷新 chart.xml 缓存。"""

    def __init__(self, docx_path: str | Path):
        self.docx_path = Path(docx_path)
        self._parts: dict[str, bytes] = {}

    def load(self) -> WordChartSync:
        if not self.docx_path.is_file():
            raise FileNotFoundError(f"找不到 docx: {self.docx_path}")
        with zipfile.ZipFile(self.docx_path, "r") as zf:
            self._parts = {name: zf.read(name) for name in zf.namelist()}
        return self

    def list_embedded_charts(self) -> list[tuple[str, str]]:
        """返回 [(chart_part, embedded_xlsx_part), ...]"""
        result: list[tuple[str, str]] = []
        for chart_part in _find_chart_parts(self._parts):
            xlsx_part = _find_embedded_workbook(self._parts, chart_part)
            if xlsx_part:
                result.append((chart_part, xlsx_part))
        return result

    def update_cells(self, values: dict[str, Any]) -> WordChartSync:
        """写入内嵌 Excel。键支持 'Sheet1!B2' 或 'Sheet1!B2:B5'（列表）。"""
        charts = self.list_embedded_charts()
        if not charts:
            raise ValueError("文档中未找到带内嵌 Excel 的图表（L1 不支持外部链接图表）")

        for _chart_part, xlsx_part in charts:
            if xlsx_part not in self._parts:
                raise FileNotFoundError(
                    f"缺少内嵌 Excel 部件: {xlsx_part}，请确认 docx 完整"
                )
            self._parts[xlsx_part] = _patch_embedded_xlsx(
                self._parts[xlsx_part], values
            )
        return self

    def refresh_caches(self) -> WordChartSync:
        """按 chart.xml 中的 c:f 公式，从内嵌 Excel 重建所有 cache。"""
        for chart_part, xlsx_part in self.list_embedded_charts():
            self._parts[chart_part] = _refresh_chart_caches(
                self._parts[chart_part], self._parts[xlsx_part]
            )
        return self

    def apply_l1(self, values: dict[str, Any]) -> WordChartSync:
        """写 Excel + 刷新 chart cache + 清理链接元数据（L1 一次完成）。"""
        return self.update_cells(values).refresh_caches().sanitize_document()

    def sanitize_document(self) -> WordChartSync:
        """清除内嵌 Excel 与 Word 设置中会导致「编辑数据」报链接错误的元数据。"""
        self.sanitize_embedded_workbooks()
        if "word/settings.xml" in self._parts:
            self._parts["word/settings.xml"] = _sanitize_word_settings(
                self._parts["word/settings.xml"]
            )
        return self

    def sanitize_embedded_workbooks(self) -> WordChartSync:
        """清除内嵌 Excel 中的 absPath / chartTrackingRefBase 等元数据。"""
        for _chart_part, xlsx_part in self.list_embedded_charts():
            self._parts[xlsx_part] = _sanitize_embedded_xlsx(self._parts[xlsx_part])
        return self

    def save(self, output_path: str | Path | None = None) -> Path:
        out = Path(output_path) if output_path else self.docx_path
        out.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(self.docx_path, "r") as zin:
            with zipfile.ZipFile(out, "w") as zout:
                for item in zin.infolist():
                    zout.writestr(item, self._parts[item.filename])
        return out


# ---------------------------------------------------------------------------
# 手动测试入口
# ---------------------------------------------------------------------------
def _run_demo() -> None:
    template = TEST_DATA_DIR / "template.docx"
    output = TEST_DATA_DIR / "output_sync.docx"

    if not template.is_file():
        print("未找到测试模板，请按以下步骤准备:")
        print(f"  1. 将含内嵌图表的 Word 文件复制为: {template}")
        print("  2. 再运行: python -m word.test.test")
        print()
        print(
            "说明: 当前 test_data/word/ 是解压后的目录，缺少 word/embeddings/*.xlsx。"
        )
        print("      请直接使用原始 .docx 作为 template.docx。")
        sys.exit(1)

    # chart1 默认结构: A 列类别, B/C/D 列三个系列（可按你的模板修改）
    demo_values: dict[str, Any] = {
        "Sheet1!A2:A5": ["新类别 1", "新类别 2", "新类别 3", "新类别 4"],
        "Sheet1!B1": "系列 1",
        "Sheet1!C1": "系列 2",
        "Sheet1!D1": "系列 3",
        "Sheet1!B2:B5": [10.0, 20.0, 30.0, 40.0],
        "Sheet1!C2:C5": [15.0, 25.0, 35.0, 45.0],
        "Sheet1!D2:D5": [6.0, 6.0, 6.0, 6.0],
    }

    sync = WordChartSync(template)
    sync.load()
    embedded = sync.list_embedded_charts()
    print(f"模板: {template}")
    print(f"发现内嵌图表 {len(embedded)} 个:")
    for chart_part, xlsx_part in embedded:
        print(f"  - {chart_part}  ->  {xlsx_part}")

    sync.apply_l1(demo_values)
    sync.save(output)

    print()
    print(f"已生成: {output}")
    print("测试前请先关闭 Word，再打开 output_sync.docx（不要打开 template.docx）。")
    print("若仍提示链接不可用：文件 -> 信息 -> 编辑指向文件的链接，看是否还有残留项。")
    print(
        "说明：程序生成文档的目标是「直接显示新图表」；编辑数据依赖 Word/Excel 本地行为，"
    )
    print("      已自动清除 absPath、chartTrackingRefBase、chartTrackingRefBased。")
    print("若需自定义数据，修改本文件 _run_demo() 中的 demo_values 后重新运行。")


if __name__ == "__main__":
    _run_demo()
