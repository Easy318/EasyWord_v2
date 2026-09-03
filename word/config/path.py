import os
from pathlib import Path

from word.config.utils import init_dirpath

# 根目录
ROOT_DIR = Path(
    os.path.dirname(os.path.dirname(os.path.dirname((os.path.abspath(__file__)))))
)
# word 目录
PROJECT_DIR = ROOT_DIR / "word"

# 项目数据目录
DATA_DIR = init_dirpath(PROJECT_DIR / "data")
# 日志目录
LOG_DIR = init_dirpath(DATA_DIR / "logs")
# 上传文件存储目录
UPLOAD_ROOT_DIR = init_dirpath(DATA_DIR / "uploads")
# 数据缓存目录
CACHE_ROOT_DIR = init_dirpath(DATA_DIR / "cache")

# 测试目录
TEST_DIR = init_dirpath(PROJECT_DIR / "test")
# 测试数据存放目录(保密,不上传)
TEST_DATA_DIR = init_dirpath(TEST_DIR / "test_data")

if __name__ == "__main__":
    print(ROOT_DIR)
    print(PROJECT_DIR)
