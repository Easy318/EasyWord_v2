# EasyWord

本地 Word / 数据分析 **sidecar** 服务。无用户登录；由净流客户端传入工作区路径或小 payload。鉴权与业务数据权限在 QHub / 客户端侧。

架构对齐 EasyAnalytiQHub：`word/core/{feature}/` 内聚 `router` + `schema` + 业务；`word/router.register_router` 统一挂载。Word COM 经 `word/runtime`（STA 单线程）访问。

## 模块边界

| 路径 | 职责 |
| --- | --- |
| `word/app.py` | FastAPI 应用工厂、lifespan（启停 COM host） |
| `word/runtime/` | COM STA host、Word.Application、文档路径绑定 |
| `word/core/document/` | 打开/保存/关闭/状态 |
| `word/core/content_control/` | ContentControl CRUD + Selection 校验 |
| `word/core/health/` | 健康检查 |
| `word/router.py` | 集中 `register_router` |
| `word/config/` | 常量、settings、路径 |
| `word/test/` | 可行性试验；**不进正式 API** |
| `cli.py` | 开发启动 |

## HTTP 契约（本机 `127.0.0.1:18765`）

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| GET | `/health` | `{ ok, service, version, apiVersion }` |
| POST | `/document/open` | `{ projectId, templatePath }` → 打开/重绑模板 |
| POST | `/document/save` | 保存绑定文档 |
| POST | `/document/close` | `{ save? }` 关闭绑定文档 |
| GET | `/document/status` | 绑定与 Word 状态 |
| GET | `/content-controls` | 列出业务控件（Tag=`eai:*`） |
| GET | `/content-controls/selection` | 当前选区命中的业务控件 |
| POST | `/content-controls` | 按选区**自动识别**类型并创建（可选 title/tag） |
| PATCH | `/content-controls` | 按选区更新命中控件 |
| DELETE | `/content-controls` | 按选区删除命中控件 |

错误体：`{ code, message }`。

## 开发环境

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

依赖统一维护在 `requirements.txt`（开发、运行、打包共用）。

## 开发启动

```bash
# 首次执行项目时(写入 project.pth)
python cli.py init_env

# 默认 http://127.0.0.1:18765（开发默认开启热重载，可用 --no-reload 关闭）
python cli.py run
```

环境变量：`EASYWORD_HOST`、`EASYWORD_PORT`。

## 打包（Windows sidecar）

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_sidecar.ps1
```

脚本会自动创建并使用项目 `.venv`（从 `requirements.txt` 安装依赖），**不要**在 `qhub_v1` 等重型 conda 环境里直接打包，否则 PyInstaller 会扫描无关大包，在 `Looking for dynamic libraries` 阶段极慢或看似卡死。

产物目录 `dist/easyword/`，供净流客户端 `extraResources` 引用。规格见 `easyword.spec`。

### 打包后快速测试（自动启停）

```powershell
# 仅测试已有产物（启动 -> /health -> 自动关闭进程）
powershell -ExecutionPolicy Bypass -File .\scripts\test_packed_sidecar.ps1

# 测试后保持进程运行（需手动 stop_sidecar）
powershell -ExecutionPolicy Bypass -File .\scripts\test_packed_sidecar.ps1 -KeepAlive

# 手动结束 easyword
powershell -ExecutionPolicy Bypass -File .\scripts\stop_sidecar.ps1
```

`build_sidecar.ps1` 在打包成功后会自动执行 `test_packed_sidecar.ps1`。

### COM 冒烟（需本机已装 Word，且服务已启动）

```powershell
# 将 PATH 换成真实 template.docx
python word/test/smoke_document_api.py --template "D:\path\to\template.docx" --project-id demo
```
