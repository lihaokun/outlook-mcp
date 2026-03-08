# Outlook MCP Server — 架构文档

> 版本：0.1.0
> 状态：初始版本

---

## 1. 概述

Outlook MCP Server 是一个 Python 实现的 MCP（Model Context Protocol）服务器，通过 Windows COM 接口（`win32com.client`）操控本地 Outlook 桌面客户端，为 AI 助手提供邮件、联系人和日历操作能力。

### 1.1 设计目标

- 完全本地运行，不依赖云端 API 或 OAuth
- 通过 stdio 传输与 MCP 客户端（Claude Code 等）通信
- 发送类操作打开 Outlook 撰写窗口供用户确认，不自动发送
- 与 mail-assistant 项目的 thunderbird-mcp 保持工具接口一致

### 1.2 约束

- **平台限制**：仅支持 Windows（依赖 COM 接口）
- **运行时依赖**：Outlook 桌面客户端必须处于运行状态
- **线程安全**：COM 对象不可跨线程使用，所有操作必须在同一线程
- **安全提示**：Outlook Object Model Guard 会在发送操作时弹出确认对话框

---

## 2. 核心流程

```
MCP Client (Claude Code)
    │
    │  JSON-RPC over stdio
    ▼
┌──────────────────────┐
│   MCP Server Layer   │  server.py — FastMCP 框架
│   (工具注册与路由)     │  接收 MCP tool call → 调用 Outlook 层 → 返回 JSON 结果
└──────────┬───────────┘
           │  Python 函数调用
           ▼
┌──────────────────────┐
│   Outlook COM Layer  │  outlook.py — win32com 封装
│   (业务逻辑实现)      │  操作 Outlook COM 对象，执行邮件/联系人/日历操作
└──────────┬───────────┘
           │  COM 接口调用
           ▼
┌──────────────────────┐
│   Outlook Desktop    │  Windows 本地 Outlook 客户端
│   (COM Server)       │
└──────────────────────┘
```

数据流向：
1. MCP Client 发起 JSON-RPC 请求（tool call）
2. `server.py` 中 FastMCP 框架解析请求，路由到对应的 `@mcp.tool()` 函数
3. Tool 函数调用 `outlook.py` 中的对应业务函数
4. `outlook.py` 通过 `win32com.client` 操作 Outlook COM 对象
5. 结果经 JSON 序列化后通过 stdio 返回给 MCP Client

---

## 3. 模块划分

### 3.1 模块总览

| 模块 | 文件 | 职责 |
|------|------|------|
| MCP Server | `outlook_mcp/server.py` | 工具注册、参数验证、JSON 序列化、MCP 传输 |
| Outlook COM | `outlook_mcp/outlook.py` | COM 接口封装、业务逻辑实现 |

### 3.2 MCP Server 模块（server.py）

```
模块名称：MCP Server

功能描述：注册 MCP 工具，接收外部调用，委托给 Outlook COM 模块执行，返回 JSON 结果。

前置条件（Requires）：
  - Python 环境已安装 mcp[cli] 和 pywin32

后置条件（Ensures）：
  - 所有 14 个 MCP 工具已注册并可通过 stdio 接收调用
  - 每个工具返回格式化的 JSON 字符串

副作用：通过 stdio 与 MCP Client 通信
```

职责边界：
- **负责**：工具声明（名称、参数 schema、描述）、参数类型转换、结果 JSON 序列化、错误包装
- **不负责**：任何 Outlook COM 操作逻辑

### 3.3 Outlook COM 模块（outlook.py）

```
模块名称：Outlook COM

功能描述：封装 Outlook COM 接口，实现邮件、联系人、日历的全部业务操作。

前置条件（Requires）：
  - Windows 操作系统
  - Outlook 桌面客户端正在运行
  - pywin32 已安装

后置条件（Ensures）：
  - 每个公开函数返回 Python dict/list，可直接 JSON 序列化
  - 发送类操作（sendMail, replyToMessage, forwardMessage, createEvent）
    调用 .Display() 而非 .Send()，确保用户确认

不变式（Invariants）：
  - COM 初始化（CoInitialize）在每次获取 Outlook 对象时执行
  - 所有 COM 操作在调用线程内完成

副作用：读写 Outlook 邮箱数据，可能触发 Outlook 安全提示
```

内部分层：
- **COM 基础设施**：`get_outlook()`, `get_namespace()` — COM 对象获取
- **文件夹操作**：`_resolve_folder()`, `_collect_folders()` — 路径解析与遍历
- **邮件转换**：`_mail_item_to_summary()`, `_mail_item_to_full()` — COM 对象到 dict 转换
- **消息定位**：`_get_item_by_entry_id()` — 基于 EntryID 查找邮件
- **公开 API**：14 个对应 MCP 工具的业务函数

---

## 4. 接口规约

### 4.1 MCP Server → Outlook COM

```
接口：server.py → outlook.py

输入数据：Python 原生类型（str, int, bool, list, None），由 FastMCP 从 JSON-RPC 参数解析而来
输出数据：Python dict 或 list[dict]，可直接 JSON 序列化

协议约定：
  - 调用方责任：参数已通过 FastMCP 类型校验，符合工具声明的类型约束
  - 被调用方责任：
    - 正常时返回 dict/list
    - 异常时抛出 ValueError（参数错误）或让 COM 异常自然传播
    - 不向 stdout 输出任何内容（避免破坏 JSON-RPC 通信）
```

### 4.2 Outlook COM → Outlook Desktop

```
接口：outlook.py → Outlook COM Server

输入数据：COM 方法调用及其参数
输出数据：COM 对象属性值

协议约定：
  - 调用方责任：
    - 每次操作前调用 pythoncom.CoInitialize()
    - 使用 EntryID 作为消息唯一标识
    - 搜索使用 Items.Restrict() 而非遍历（性能考虑）
  - 被调用方责任（Outlook）：
    - 提供有效的 COM 接口
    - 发送操作时触发 Object Model Guard 安全确认
```

---

## 5. 工具分类与映射

### 5.1 工具清单

| 分类 | MCP 工具 | Outlook COM 方法 |
|------|----------|-----------------|
| 账户与文件夹 | `listAccounts` | `Namespace.Accounts` |
| | `listFolders` | `Namespace.Folders` 递归遍历 |
| | `createFolder` | `Folder.Folders.Add()` |
| 搜索与读取 | `searchMessages` | `Items.Restrict()` + DASL 过滤 |
| | `getRecentMessages` | `Items.Restrict("[ReceivedTime] >= ...")` |
| | `getMessage` | `Namespace.GetItemFromID()` / 遍历查找 |
| 邮件操作 | `sendMail` | `CreateItem(0)` → `.Display()` |
| | `replyToMessage` | `MailItem.Reply()` / `.ReplyAll()` → `.Display()` |
| | `forwardMessage` | `MailItem.Forward()` → `.Display()` |
| | `updateMessage` | `.UnRead`, `.FlagStatus`, `.Move()`, `.Delete()` |
| | `deleteMessages` | `.Delete()` 批量 |
| 联系人 | `searchContacts` | `GetDefaultFolder(10).Items.Restrict()` |
| 日历 | `listCalendars` | `Namespace.Folders` → 日历类型文件夹 |
| | `createEvent` | `CreateItem(1)` → `.Display()` |

### 5.2 Outlook 文件夹常量

| 常量 | 值 | 说明 |
|------|----|------|
| olFolderInbox | 6 | 收件箱 |
| olFolderOutbox | 4 | 发件箱 |
| olFolderSentMail | 5 | 已发送 |
| olFolderDeletedItems | 3 | 已删除 |
| olFolderDrafts | 16 | 草稿 |
| olFolderCalendar | 9 | 日历 |
| olFolderContacts | 10 | 联系人 |
| olFolderJunk | 23 | 垃圾邮件 |

---

## 6. 关键设计决策

### 6.1 Display 而非 Send

**决策**：所有发送类操作（sendMail, replyToMessage, forwardMessage, createEvent）调用 `.Display()` 而非 `.Send()`。

**理由**：
- 用户安全：AI 不应在无人确认的情况下自动发送邮件
- 与 thunderbird-mcp 行为一致（打开撰写窗口）
- 避免 Object Model Guard 安全警告的干扰

### 6.2 EntryID 作为消息标识

**决策**：使用 Outlook 的 `EntryID` 属性作为消息唯一标识。

**理由**：
- EntryID 是 Outlook COM 中原生的唯一标识符
- 在同一 Outlook profile 内保持稳定
- 注意：跨 profile 或邮件移动后 EntryID 可能改变

### 6.3 每次调用重新获取 COM 对象

**决策**：`get_outlook()` 和 `get_namespace()` 在每次公开 API 调用时重新创建。

**理由**：
- COM 对象可能因 Outlook 重启而失效
- 避免缓存过期的 COM 引用导致不可预期错误
- 性能开销可接受（MCP 调用频率不高）

### 6.4 搜索策略

**决策**：优先使用 `Items.Restrict()` + DASL 过滤器进行搜索。

**理由**：
- `Restrict()` 在服务端过滤，性能远优于 Python 逐条遍历
- DASL 支持主题、发件人、收件人等多字段模糊匹配
- 回退机制：联系人搜索如 DASL 失败，降级为前 500 条遍历匹配

### 6.5 收件箱名称国际化

**决策**：`getRecentMessages` 无指定文件夹时，依次尝试 "Inbox" 和 "收件箱"。

**理由**：
- 中文版 Outlook 的收件箱文件夹名为 "收件箱"
- 英文版为 "Inbox"
- 简单的 fallback 策略覆盖最常见场景

---

## 7. 注意事项

1. **stdout 禁令**：`outlook.py` 中严禁使用 `print()` 输出到 stdout，会破坏 JSON-RPC 通信。调试信息应使用 `logging` 或 `print(..., file=sys.stderr)`
2. **编码处理**：中文邮件主题和正文通过 COM 接口获取时已自动处理编码
3. **大量邮件性能**：搜索结果上限 200 条，避免 COM 遍历超时
4. **附件保存**：保存到 `tempfile.mkdtemp()` 创建的临时目录，由调用方负责清理
