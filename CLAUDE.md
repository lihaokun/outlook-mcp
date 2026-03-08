# outlook-mcp

> MCP server for Outlook — 通过 Windows COM 接口为 AI 助手提供邮件、联系人和日历操作能力

## 项目结构

```
outlook_mcp/       — 核心源码
  server.py        — MCP Server（FastMCP 工具注册与路由）
  outlook.py       — Outlook COM 接口封装（win32com）
docs/              — 设计文档与流程规范
  workflow.md      — 开发工作流程规范（@docs/workflow.md）
  design/          — 架构与细化文档
```

## 技术栈

- 语言：Python 3.10+
- COM 接口：pywin32 (`win32com.client`)
- MCP 框架：mcp[cli] (FastMCP)
- 传输：stdio
- 平台：Windows（需 Outlook 桌面版运行中）

## 常用命令

```bash
# 安装（开发模式）
pip install -e .

# 运行 MCP Server
outlook-mcp

# 或通过 Python 模块运行
python -m outlook_mcp.server
```

## 代码风格

- 命名：Python 内部 snake_case，MCP 工具参数 camelCase（与 thunderbird-mcp 一致）
- 注释语言：英文（docstring），中文（设计文档）

## 参考实现

- thunderbird-mcp：https://github.com/TKasperczyk/thunderbird-mcp — 对标工具接口设计
- mail-assistant：https://github.com/lihaokun/mail-assistant — 上游项目，定义功能规格

## 已知限制与注意事项

- Outlook 桌面客户端必须处于运行状态，否则 COM 调用失败
- 发送/回复/转发调用 `.Display()` 而非 `.Send()`，需用户手动确认
- COM 对象不可跨线程使用
- `EntryID` 在邮件移动到不同 Store 后可能改变
- 中文版 Outlook 文件夹名为 "收件箱" 而非 "Inbox"，代码已做兼容
- outlook.py 中严禁 print() 到 stdout（会破坏 JSON-RPC 通信）

## 工作流程

遵循 @docs/workflow.md 中定义的开发流程规范。
