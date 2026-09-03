"""手工冒烟：document open/status（不进发布契约）。

用法（先启动 EasyWord）::

    python word/test/smoke_document_api.py --template "D:\\path\\template.docx" --project-id demo
"""

from __future__ import annotations

import argparse
import json
import urllib.error
import urllib.request

BASE = "http://127.0.0.1:18765"


def _req(method: str, path: str, body: dict | None = None) -> dict:
    data = None if body is None else json.dumps(body).encode("utf-8")
    request = urllib.request.Request(
        f"{BASE}{path}",
        data=data,
        method=method,
        headers={"Content-Type": "application/json", "Accept": "application/json"},
    )
    try:
        with urllib.request.urlopen(request, timeout=60) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise SystemExit(f"HTTP {exc.code}: {detail}") from exc


def main() -> None:
    parser = argparse.ArgumentParser(description="EasyWord document API smoke")
    parser.add_argument("--template", required=True, help="本机 template.docx 绝对路径")
    parser.add_argument("--project-id", default="smoke")
    args = parser.parse_args()

    health = _req("GET", "/health")
    print("health:", health)

    opened = _req(
        "POST",
        "/document/open",
        {"projectId": args.project_id, "templatePath": args.template},
    )
    print("open:", opened)

    status = _req("GET", "/document/status")
    print("status:", status)

    controls = _req("GET", "/content-controls")
    print("content-controls:", controls)

    print("OK — 可在 Word 中选中内容后用客户端或 curl 测 create/update/delete")


if __name__ == "__main__":
    main()
