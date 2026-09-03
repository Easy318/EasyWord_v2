from pathlib import Path


def init_dirpath(dirpath: Path):
    """初始化文件夹,如果文件夹不存在则创建

    Args:
        path ('Path'): 文件夹路径
    """
    if not dirpath.is_dir():
        dirpath.mkdir(parents=True)
    return dirpath
