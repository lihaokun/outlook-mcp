# Outlook MCP Server — 细化文档

> 版本：0.1.0
> 状态：初始版本
> 前置：请先阅读 architecture.md

---

## 1. Outlook COM 模块（outlook.py）

### 1.1 COM 基础设施

#### `get_outlook() -> COMObject`

```
功能描述：获取 Outlook Application COM 对象

前置条件（Requires）：
  - Windows 操作系统
  - Outlook 桌面客户端已安装

后置条件（Ensures）：
  - 返回可用的 Outlook.Application COM 对象
  - 当前线程已调用 CoInitialize()

副作用：初始化当前线程的 COM 环境
```

#### `get_namespace(outlook=None) -> COMObject`

```
功能描述：获取 MAPI Namespace 对象

前置条件（Requires）：
  - outlook 参数为 None 或有效的 Outlook.Application 对象

后置条件（Ensures）：
  - 返回 MAPI Namespace 对象，可用于访问文件夹、账户等

副作用：若 outlook=None，则内部调用 get_outlook()
```

### 1.2 内部辅助函数

#### `_resolve_folder(ns, folder_path: str) -> COMObject`

```
功能描述：根据路径字符串解析到对应的 Outlook 文件夹 COM 对象

前置条件（Requires）：
  - ns 为有效的 MAPI Namespace
  - folder_path 格式为 "TopLevelName/SubFolder/..." 或空字符串

后置条件（Ensures）：
  - folder_path 非空时：返回路径对应的文件夹 COM 对象
  - folder_path 为空时：返回默认收件箱（olFolderInbox = 6）
  - 路径中任一层级不存在时：抛出 ValueError

副作用：无
```

解析逻辑：
1. 空路径 → 返回 `ns.GetDefaultFolder(6)`
2. 按 `/` 分割路径
3. 第一段在 `ns.Folders` 中匹配顶级文件夹（通常是账户名）
4. 后续各段在当前文件夹的 `Folders` 中逐层匹配

#### `_collect_folders(folder, prefix="") -> list[dict]`

```
功能描述：递归收集文件夹信息

前置条件（Requires）：
  - folder 为有效的 Outlook Folder COM 对象

后置条件（Ensures）：
  - 返回 list[dict]，每个 dict 包含 name, path, totalCount, unreadCount
  - 包含 folder 本身及其所有子文件夹（递归）

副作用：无
```

#### `_mail_item_to_summary(item, folder_path="") -> dict | None`

```
功能描述：将 MailItem COM 对象转换为摘要 dict

前置条件（Requires）：
  - item 为 Outlook MailItem COM 对象

后置条件（Ensures）：
  - 正常时返回 dict，键：messageId, folderPath, subject, sender, date, read
  - item 属性访问异常时返回 None

副作用：无
```

返回字段说明：
- `messageId`：EntryID，用于后续 getMessage/updateMessage 等操作
- `sender`：格式为 "SenderName <SenderEmail>"（若两者不同）或仅邮箱地址
- `date`：ReceivedTime 的字符串表示
- `read`：布尔值，True=已读

#### `_mail_item_to_full(item, folder_path="", save_attachments=False) -> dict`

```
功能描述：将 MailItem COM 对象转换为完整详情 dict

前置条件（Requires）：
  - item 为 Outlook MailItem COM 对象

后置条件（Ensures）：
  - 返回 dict，包含 summary 的全部字段，外加 to, cc, body, htmlBody, attachments
  - save_attachments=True 时，附件保存到临时目录，dict 中包含 savedPath

副作用：save_attachments=True 时创建临时目录并写入附件文件
```

#### `_get_item_by_entry_id(ns, entry_id: str, folder_path: str) -> COMObject`

```
功能描述：根据 EntryID 定位邮件 COM 对象

前置条件（Requires）：
  - ns 为有效的 MAPI Namespace
  - entry_id 为有效的 Outlook EntryID 字符串
  - folder_path 为有效的文件夹路径

后置条件（Ensures）：
  - 返回对应的 MailItem COM 对象
  - 未找到时抛出 ValueError

副作用：无
```

查找策略：
1. 解析 folder_path 得到文件夹，遍历其 Items 匹配 EntryID
2. 若遍历未找到，回退到 `ns.GetItemFromID(entry_id)`
3. 两种方式均失败则抛出 ValueError

---

### 1.3 公开 API — 账户与文件夹

#### `list_accounts() -> list[dict]`

```
功能描述：列出 Outlook 中配置的所有邮箱账户

前置条件（Requires）：无额外条件

后置条件（Ensures）：
  - 返回 list[dict]，每个 dict 包含 name, email, accountType

副作用：无
```

实现：遍历 `ns.Accounts`，逐个提取 DisplayName、SmtpAddress、AccountType。

#### `list_folders(account_id: str | None = None) -> list[dict]`

```
功能描述：列出所有邮件文件夹及消息统计

前置条件（Requires）：
  - account_id 为 None 或已配置的账户名称

后置条件（Ensures）：
  - 返回 list[dict]，每个 dict 包含 name, path, totalCount, unreadCount
  - account_id 非 None 时仅返回该账户下的文件夹

副作用：无
```

实现：遍历 `ns.Folders` 顶层，调用 `_collect_folders()` 递归收集。

#### `create_folder(parent_folder_path: str, name: str) -> dict`

```
功能描述：在指定父文件夹下创建子文件夹

前置条件（Requires）：
  - parent_folder_path 对应一个已存在的文件夹
  - name 不与现有子文件夹重名

后置条件（Ensures）：
  - 新文件夹已创建
  - 返回 dict 包含 name, path

副作用：在 Outlook 中创建新文件夹
```

---

### 1.4 公开 API — 搜索与读取

#### `search_messages(query, start_date=None, end_date=None, max_results=50, sort_order="desc") -> list[dict]`

```
功能描述：按关键词搜索邮件，支持日期范围过滤

前置条件（Requires）：
  - query 为非空字符串
  - max_results ∈ [1, 200]

后置条件（Ensures）：
  - 返回 list[dict]（摘要格式），长度 ≤ min(max_results, 200)
  - 按 ReceivedTime 排序（sort_order 控制升降序）
  - 搜索范围：所有账户的所有文件夹（递归）

副作用：无
```

搜索实现：
1. 构建 DASL 过滤器：`urn:schemas:httpmail:subject/fromemail/displayto LIKE '%query%'`
2. 可选追加日期过滤：`[ReceivedTime] >= 'start_date'`
3. 对每个文件夹递归调用 `Items.Restrict()` + `Sort()`
4. 累计到 max_results 后停止

#### `get_recent_messages(folder_path=None, days_back=7, max_results=50, unread_only=False) -> list[dict]`

```
功能描述：获取最近 N 天的邮件

前置条件（Requires）：
  - days_back ≥ 1
  - max_results ∈ [1, 200]

后置条件（Ensures）：
  - 返回 list[dict]（摘要格式），按时间倒序
  - folder_path 指定时仅搜索该文件夹
  - folder_path 为 None 时搜索所有账户的收件箱

副作用：无
```

实现要点：
- 日期过滤：`[ReceivedTime] >= 'MM/DD/YYYY HH:MM AM'` 格式（Outlook Restrict 要求）
- 未读过滤：追加 `AND [UnRead] = True`
- 无指定文件夹时尝试 "Inbox" 和 "收件箱" 两种名称

#### `get_message(message_id, folder_path, save_attachments=False) -> dict`

```
功能描述：读取邮件完整内容

前置条件（Requires）：
  - message_id 为有效的 EntryID
  - folder_path 为邮件所在文件夹路径

后置条件（Ensures）：
  - 返回完整邮件 dict（含 body, htmlBody, attachments 等）

副作用：save_attachments=True 时写入临时文件
```

---

### 1.5 公开 API — 邮件操作

#### `send_mail(to, subject, body, cc=None, bcc=None, is_html=False, from_account=None, attachments=None) -> dict`

```
功能描述：创建新邮件并打开 Outlook 撰写窗口

前置条件（Requires）：
  - to 为非空字符串（一个或多个收件人，分号分隔）
  - subject 为非空字符串
  - attachments 中的路径必须指向已存在的文件

后置条件（Ensures）：
  - Outlook 撰写窗口已打开，包含填好的收件人、主题、正文
  - 返回 {"status": "displayed", "message": "..."}

副作用：打开 Outlook 撰写窗口
```

实现要点：
- `CreateItem(0)` 创建 MailItem
- from_account：遍历 `ns.Accounts` 匹配 SmtpAddress 或 DisplayName
- attachments：逐个检查 `os.path.isfile()` 后调用 `Attachments.Add()`
- 最后调用 `.Display()` 而非 `.Send()`

#### `reply_to_message(message_id, folder_path, body, reply_all=False, ...) -> dict`

```
功能描述：回复邮件并打开 Outlook 撰写窗口

前置条件（Requires）：
  - message_id 对应一封有效邮件

后置条件（Ensures）：
  - 回复内容 prepend 到原始邮件内容前
  - Outlook 撰写窗口已打开

副作用：打开 Outlook 撰写窗口
```

实现：`item.Reply()` 或 `item.ReplyAll()` → 设置 Body/HTMLBody → `.Display()`

#### `forward_message(message_id, folder_path, to, body=None, ...) -> dict`

```
功能描述：转发邮件并打开 Outlook 撰写窗口

后置条件（Ensures）：
  - 转发包含原始邮件内容和附件
  - 附加正文 prepend 到原始内容前

副作用：打开 Outlook 撰写窗口
```

实现：`item.Forward()` → 设置 To、Body → `.Display()`

#### `update_message(message_id, folder_path, read=None, flagged=None, move_to=None, trash=False) -> dict`

```
功能描述：更新邮件状态

前置条件（Requires）：
  - move_to 和 trash 不能同时为真
  - move_to 路径必须对应已存在的文件夹

后置条件（Ensures）：
  - read != None 时：item.UnRead = not read，并 Save()
  - flagged != None 时：item.FlagStatus 更新，并 Save()
  - trash=True 时：item.Delete()（移入已删除）
  - move_to != None 时：item.Move(target_folder)

副作用：修改 Outlook 中的邮件状态
```

操作顺序：先设置属性并 Save()，再执行 Move/Delete（因为移动/删除后原引用失效）。

#### `delete_messages(message_ids: list[str], folder_path: str) -> dict`

```
功能描述：批量删除邮件

后置条件（Ensures）：
  - 返回 {"deleted": N, "errors": [...]}
  - 每条消息独立处理，单条失败不影响其余

副作用：将邮件移入 Outlook 已删除文件夹
```

---

### 1.6 公开 API — 联系人

#### `search_contacts(query: str) -> list[dict]`

```
功能描述：在默认联系人文件夹中搜索联系人

前置条件（Requires）：
  - query 为非空字符串

后置条件（Ensures）：
  - 返回 list[dict]，每个 dict 包含 name, email, phone, company

副作用：无
```

实现：
1. 优先尝试 DASL Restrict（`urn:schemas:contacts:cn` / `email1`）
2. 失败时降级为遍历前 500 条联系人，Python 端做字符串匹配

---

### 1.7 公开 API — 日历

#### `list_calendars() -> list[dict]`

```
功能描述：列出所有日历

后置条件（Ensures）：
  - 返回 list[dict]，每个 dict 包含 name, path, itemCount
  - 包含所有账户下的日历文件夹（通过 DefaultItemType == 1 识别）

副作用：无
```

实现：递归遍历所有文件夹，筛选 `DefaultItemType == OL_APPOINTMENT_ITEM (1)` 的文件夹。

#### `create_event(title, start_date, end_date=None, location=None, description=None, calendar_id=None, all_day=False) -> dict`

```
功能描述：创建日历事件并打开 Outlook 确认窗口

前置条件（Requires）：
  - title 为非空字符串
  - start_date 为有效的 ISO 8601 时间字符串

后置条件（Ensures）：
  - end_date 为 None 时默认 start_date + 1 小时
  - Outlook 事件编辑窗口已打开

副作用：打开 Outlook 事件编辑窗口
```

实现：`CreateItem(1)` → 设置属性 → `.Display()`

---

## 2. MCP Server 模块（server.py）

### 2.1 职责

每个 `@mcp.tool()` 函数：
1. 接收 MCP 参数（FastMCP 自动从 JSON-RPC 解析和类型校验）
2. 调用 `outlook.py` 中对应的公开 API
3. 用 `_json()` 将结果序列化为格式化 JSON 字符串返回

### 2.2 工具注册清单

| MCP 工具名 | 调用的 outlook.py 函数 |
|-----------|----------------------|
| `listAccounts` | `outlook.list_accounts()` |
| `listFolders` | `outlook.list_folders(accountId)` |
| `createFolder` | `outlook.create_folder(parentFolderPath, name)` |
| `searchMessages` | `outlook.search_messages(query, startDate, endDate, maxResults, sortOrder)` |
| `getRecentMessages` | `outlook.get_recent_messages(folderPath, daysBack, maxResults, unreadOnly)` |
| `getMessage` | `outlook.get_message(messageId, folderPath, saveAttachments)` |
| `sendMail` | `outlook.send_mail(to, subject, body, cc, bcc, isHtml, fromAccount, attachments)` |
| `replyToMessage` | `outlook.reply_to_message(...)` |
| `forwardMessage` | `outlook.forward_message(...)` |
| `updateMessage` | `outlook.update_message(messageId, folderPath, read, flagged, moveTo, trash)` |
| `deleteMessages` | `outlook.delete_messages(messageIds, folderPath)` |
| `searchContacts` | `outlook.search_contacts(query)` |
| `listCalendars` | `outlook.list_calendars()` |
| `createEvent` | `outlook.create_event(title, startDate, endDate, location, description, calendarId, allDay)` |

### 2.3 参数命名映射

MCP 工具参数使用 camelCase（与 thunderbird-mcp 保持一致），outlook.py 内部使用 snake_case：

| MCP 参数 | Python 参数 |
|----------|------------|
| `accountId` | `account_id` |
| `folderPath` | `folder_path` |
| `parentFolderPath` | `parent_folder_path` |
| `startDate` | `start_date` |
| `endDate` | `end_date` |
| `maxResults` | `max_results` |
| `sortOrder` | `sort_order` |
| `daysBack` | `days_back` |
| `unreadOnly` | `unread_only` |
| `messageId` | `message_id` |
| `saveAttachments` | `save_attachments` |
| `isHtml` | `is_html` |
| `fromAccount` | `from_account` |
| `replyAll` | `reply_all` |
| `moveTo` | `move_to` |
| `messageIds` | `message_ids` |
| `calendarId` | `calendar_id` |
| `allDay` | `all_day` |
| `startDate` | `start_date` |

### 2.4 错误处理策略

- FastMCP 自动处理参数类型校验错误，返回 MCP 标准错误响应
- outlook.py 抛出的 `ValueError` 自然传播，FastMCP 包装为错误响应
- COM 异常（如 Outlook 未运行）同样传播为 MCP 错误
- 不在 server.py 中做额外 try/except，保持错误信息透明

### 2.5 入口点

```python
def main():
    mcp.run(transport="stdio")
```

通过 `pyproject.toml` 注册为 `outlook-mcp` 命令行入口。
