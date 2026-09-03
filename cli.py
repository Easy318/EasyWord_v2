import os
import sysconfig
from pathlib import Path

import typer
import uvicorn
from rich import print

from word.config.settings import get_host, get_port

tp = typer.Typer(help="EasyWord 本地 sidecar 开发工具")

# 热重载仅监视业务代码目录，避免扫整个仓库
_RELOAD_DIRS = [str(Path(__file__).resolve().parent / "word")]


@tp.command()
def init_env():
    """将本仓库根目录写入 site-packages/project.pth，便于本地 import word。"""
    workspace_dir = os.path.dirname(os.path.abspath(__file__))
    site_packages_path = sysconfig.get_path("purelib")
    pth_path = Path(site_packages_path) / "project.pth"

    if pth_path.is_file():
        data = pth_path.read_text(encoding="utf-8")
        paths = [Path(path) for path in data.split() if path.strip()]
        if Path(workspace_dir) not in paths and workspace_dir not in {str(p) for p in paths}:
            paths.append(Path(workspace_dir))
        data = "\n".join(str(path) for path in paths)
        pth_path.write_text(data, encoding="utf-8")
    else:
        pth_path.write_text(str(workspace_dir), encoding="utf-8")
    print("[green]开发测试环境初始化完成")


@tp.command()
def run(
    host: str | None = typer.Option(None, help="默认 127.0.0.1，可用 EASYWORD_HOST 覆盖"),
    port: int | None = typer.Option(None, help="默认 18765，可用 EASYWORD_PORT 覆盖"),
    reload: bool = typer.Option(
        True,
        "--reload/--no-reload",
        help="代码改动时自动重载（开发默认开启）",
    ),
):
    """启动 EasyWord HTTP 服务（仅本机 loopback）。"""
    bind_host = host or get_host()
    bind_port = port if port is not None else get_port()
    print(f"[cyan]EasyWord listening on http://{bind_host}:{bind_port}")
    if reload:
        print(f"[dim]热重载已开启，监视: {_RELOAD_DIRS[0]}")
    uvicorn.run(
        "word.app:app",
        host=bind_host,
        port=bind_port,
        reload=reload,
        reload_dirs=_RELOAD_DIRS if reload else None,
        log_level="info",
    )


if __name__ == "__main__":
    tp()
