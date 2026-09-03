"""与 QHub M0 对齐的常量（禁止 import QHub）。JSON 线格式 camelCase。"""

from __future__ import annotations

import re
from typing import Literal

from pydantic import ConfigDict

SCHEMA_VERSION = "1.0"
CAMEL_CONFIG = ConfigDict(from_attributes=True, populate_by_name=True)
REF_NAME_PATTERN = r"^(?P<series>[a-z]+)(?P<index>[1-9]\d*)$"
REF_NAME_RE = re.compile(REF_NAME_PATTERN)
JINJA_RESERVED_NAMES = frozenset({"burst_var", "burst_value"})

ControlType = Literal["text", "chart", "table", "picture"]
RenderSource = Literal["flow", "projectConst", "projectVar", "burst"]
RenderStatus = Literal["active", "tombstone", "missing"]
ValueDtype = Literal["null", "bool", "int", "float", "str", "list", "dict"]
