"""合并单元格：OOXML gridSpan/vMerge 解析。"""

from __future__ import annotations

from word.core.content_control.template_seed import _parse_table_ooxml
from word.core.preview.grid_merge import is_covered, physical_cell_index

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _header_tbl_xml() -> str:
    """区/高中校竖合；中考/高考/匹配后各横合 3 列。"""
    return f"""<?xml version="1.0"?>
<w:tbl xmlns:w="{_W}">
  <w:tr>
    <w:tc>
      <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
      <w:p><w:r><w:t>区/高中校</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:gridSpan w:val="3"/></w:tcPr>
      <w:p><w:r><w:t>中考</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:gridSpan w:val="3"/></w:tcPr>
      <w:p><w:r><w:t>高考</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:gridSpan w:val="3"/></w:tcPr>
      <w:p><w:r><w:t>匹配后</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
  <w:tr>
    <w:tc>
      <w:tcPr><w:vMerge/></w:tcPr>
      <w:p/>
    </w:tc>
    <w:tc><w:p><w:r><w:t>2021</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>2022</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>2023</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>2024</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>2025</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>2026</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>2024</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>2025</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>2026</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p><w:r><w:t>石景山区</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>2173</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>1877</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>2134</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>1171</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>1367</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>1482</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>1053</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>1162</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>1273</w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>
"""


def test_parse_mixed_header_ooxml() -> None:
    parsed = _parse_table_ooxml(_header_tbl_xml())
    assert parsed is not None
    rows, merges = parsed
    assert rows[0][0] == "区/高中校"
    assert rows[0][1] == "中考"
    assert rows[0][4] == "高考"
    assert rows[0][7] == "匹配后"
    assert rows[0][2] == ""
    assert rows[1][0] == ""
    assert rows[1][1] == "2021"
    assert rows[1][9] == "2026"
    assert rows[2][0] == "石景山区"
    assert merges == [
        {"row": 0, "col": 0, "rowSpan": 2, "colSpan": 1},
        {"row": 0, "col": 1, "rowSpan": 1, "colSpan": 3},
        {"row": 0, "col": 4, "rowSpan": 1, "colSpan": 3},
        {"row": 0, "col": 7, "rowSpan": 1, "colSpan": 3},
    ]
    assert is_covered(1, 0, merges)
    assert is_covered(0, 2, merges)
    assert not is_covered(0, 1, merges)


def test_infer_vertical_when_vmerge_omitted() -> None:
    """Word 有时下行是空 tc、无 vMerge；应仍推出 rowSpan=2。"""
    xml = f"""<?xml version="1.0"?>
<w:tbl xmlns:w="{_W}">
  <w:tr>
    <w:tc><w:p><w:r><w:t>区/高中校</w:t></w:r></w:p></w:tc>
    <w:tc>
      <w:tcPr><w:gridSpan w:val="3"/></w:tcPr>
      <w:p><w:r><w:t>中考</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p/></w:tc>
    <w:tc><w:p><w:r><w:t>2021</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>2022</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>2023</w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>
"""
    parsed = _parse_table_ooxml(xml)
    assert parsed is not None
    rows, merges = parsed
    assert rows[0][0] == "区/高中校"
    assert rows[1][0] == ""
    assert {"row": 0, "col": 0, "rowSpan": 2, "colSpan": 1} in merges
    assert {"row": 0, "col": 1, "rowSpan": 1, "colSpan": 3} in merges

    xml = f"""<?xml version="1.0"?>
<w:tbl xmlns:w="{_W}">
  <w:tr>
    <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>
"""
    parsed = _parse_table_ooxml(xml)
    assert parsed is not None
    rows, merges = parsed
    assert rows == [["A", "B"]]
    assert merges == []


def test_physical_index_gaokao_anchor() -> None:
    """逻辑 (0,4) 高考是第 3 个物理格，不是 Cell(1,5)。"""
    merges = [
        {"row": 0, "col": 0, "rowSpan": 2, "colSpan": 1},
        {"row": 0, "col": 1, "rowSpan": 1, "colSpan": 3},
        {"row": 0, "col": 4, "rowSpan": 1, "colSpan": 3},
        {"row": 0, "col": 7, "rowSpan": 1, "colSpan": 3},
    ]
    assert physical_cell_index(0, 0, 11, 10, merges) == 1
    assert physical_cell_index(0, 1, 11, 10, merges) == 2
    assert physical_cell_index(0, 4, 11, 10, merges) == 3
    assert physical_cell_index(0, 7, 11, 10, merges) == 4
    assert physical_cell_index(1, 1, 11, 10, merges) == 5
    assert physical_cell_index(2, 0, 11, 10, merges) == 14
