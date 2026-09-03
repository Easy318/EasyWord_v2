"""核心业务按功能模块划分（对齐 EasyAnalytiQHub/qhub/core）。

新增能力：在本目录下建 `{feature}/`，内含 `router/`、`schema/` 与业务实现；
在 `word.router.register_router` 中挂载模块导出的 `*_api`。
试验代码放 `word/test/`，勿直接挂正式路由。
"""
