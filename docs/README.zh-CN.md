# Cat5Dev

Cat5Dev 是一款用于在 VSCode 和 CATIA V5 之间同步 VBA 模块的 VSCode 扩展。

---

## 功能特点

### 拉取 / 推送 (Pull / Push) VBA 模块
在 CATIA V5 和本地工作区之间同步 VBA 模块。

- **拉取 (Pull)** — 将目标 CATIA 项目中的所有 VBA 模块导出到本地的 `modules/` 文件夹。
- **推送 (Push)** — 将本地 VBA 模块导入回 CATIA（仅传输有改动的文件，未改动的模块将被跳过以提高速度）。

### 选择目标项目
通过快速选择对话框，选择要同步的 CATIA VBA 项目。
所选项目名称将保存在 `cat5dev.toml` 的 `[project]` 部分中，并在不同的会话间持久保留。

### 模块树形视图 (Module Tree View)
活动栏 (Activity Bar) 中的专用面板将显示工作区的当前状态。

- 显示目标项目名称
- 模块按类型分组：
  - **标准模块** (`.bas_utf`)
  - **类模块** (`.cls_utf`)
  - **用户窗体** (`.frm_utf`)
- 单击任何模块即可在编辑器中将其打开
- 更改目标项目时，树形视图会自动更新
- 右键单击任何模块以访问上下文菜单选项：
  - **重命名 (Rename)** — 重命名模块文件
  - **删除 (Delete)** — 删除模块文件（有确认提示）
  - **复制路径 (Copy Path)** — 将完整的文件路径复制到剪贴板

### 工具栏按钮
面板标题栏中提供了快速访问按钮：

| 图标 | 操作 |
|------|--------|
| ☁️↓ | 从 CATIA 拉取 (Pull) |
| ☁️↑ | 推送到 CATIA (Push) |
| ⚙️ | 选择目标项目 |
| 🔄 | 刷新模块列表 |
| 🌐 | 切换语言 |

### 多语言支持
本扩展支持 **日语 (日本語)**、**英语 (English)** 和 **中文 (简体中文)**。
- 单击面板标题栏中的 🌐 按钮切换语言
- 语言偏好设置保存在 `cat5dev.toml` 的 `[project]` 部分下
- 所有 UI 消息、TreeView 标签和工具提示都会被翻译
- 更改语言后需要重新加载 VSCode 才能生效

### 符号导航 (Symbol Navigation)
VBA 符号会被识别并公开给 VSCode 的导航功能：

- **面包屑导航 (Breadcrumb)** — 随着光标移动显示当前的 `Sub` / `Function` / `Property` 名称
- **大纲视图 (Outline view)** — 列出当前文件中的所有过程 (Procedure) 和属性 (Property)
- **跳转到符号 (Go to Symbol)** (`Ctrl+Shift+O`) — 直接跳转到任何过程

可识别的符号类型：`Sub`, `Function`, `Property Get/Let/Set`, `Type`, `Enum`

### VBA 代码检查器 (Linter)

在您编辑 `.bas_utf`、`.cls_utf` 和 `.frm_utf` 文件时，内置的 VBA 校验器 (Linter) 会实时运行。
诊断结果会以内联波浪线形式显示，并在 **问题 (Problems)** 面板中列出。

**规则 (Rules)：**

| 诊断代码 | 规则 | 严重程度 |
|------|------|----------|
| VBA001 | 未声明 `Option Explicit` | 警告 |
| VBA002 | 使用 `On Error Resume Next` | 警告 |
| VBA003 | 使用 `GoTo`（不包括 `On Error GoTo`） | 警告 |
| VBA004 | 行长度超过最大限制（默认：100） | 警告 |
| VBA005 | 使用 `Dim` 声明了变量但从未被使用 | 警告 |
| VBA006 | 嵌套深度超过阈值（默认：5） | 警告 |
| VBA007 | Sub/Function 超过行数阈值（默认：300） | 警告 |
| VBA008 | 括号不匹配 | 错误 |
| VBA009 | 缺少 `End If` / `End Sub` / `End Function` 等 | 错误 |

**配置 (`cat5dev.toml`)：**

代码检查 (Linting) 默认**被禁用**。要启用它，请在工作区根目录中创建一个 `cat5dev.toml` 文件：

1. 从命令面板运行 **`cat5dev.init`** 生成配置模板，或者
2. 按照以下结构手动创建 `cat5dev.toml`：

```toml
[project]
target_project = ""
language = "zh"

[lint]
enabled = true

[lint.rules]
option_explicit       = true
on_error_resume_next  = false
goto                  = false
max_line_length       = 100   # 0 = 禁用
unused_variables      = true
max_nesting_depth     = 5     # 0 = 禁用
max_function_lines    = 300   # 0 = 禁用
unmatched_parens      = true
unmatched_blocks      = true

[formatter]
enabled = false
# ... (参见下面的格式化器部分)
```

对 `cat5dev.toml` 的更改在保存后立即生效——无需重新加载。

### VBA 格式化器 (Formatter) (`Shift+Alt+F` / 保存时)

内置的 VBA 代码格式化器可按需使用 (`Shift+Alt+F`)，也可在保存时自动运行。

格式化器默认**被禁用**。要启用它，请在 `cat5dev.toml` 中设置 `enabled = true`：

```toml
[formatter]
enabled = true
```

**默认格式化处理（启用时）：**
- 关键字大小写规范（如 `dim` → `Dim`，`end if` → `End If` 等）
- 缩进修正（支持的块包括：`Sub`、`If`、`For`、`With`、`Select Case` 等）
- 删除行尾空白字符
- 换行连字符前添加空格（如 `Show(_` → `Show( _`）
- 换行缩进（以 `_` 结束的行下一行会多加一级缩进；仅包含 `)` 的闭合行除外）
- 空行规范化（连续空行最多保留 2 行；过程与过程之间保证有空行）

**可选格式化选项（默认禁用）：**

| `cat5dev.toml` 键名 | 描述 |
|---------------------|-------------|
| `normalize_operator_spacing` | `x=1+2` → `x = 1 + 2` |
| `normalize_comma_spacing` | `foo(a,b)` → `foo(a, b)` |
| `normalize_comment_space` | `'comment` → `' comment` |
| `expand_type_suffixes` | `Dim x%` → `Dim x As Integer` |

**完整的 `cat5dev.toml` 格式化器配置：**

```toml
[formatter]
enabled = true
indent_size = 4
capitalize_keywords = true
fix_indentation = true
trim_trailing_space = true
ensure_continuation_space = true
indent_continuation_lines = true
max_blank_lines = 2
normalize_operator_spacing = false
normalize_comma_spacing = false
normalize_comment_space = false
expand_type_suffixes = false
format_on_save = false
```

---

## 环境要求

- 必须运行 CATIA V5 并打开一个 VBA 项目
- Windows 系统（通过 `cscript.exe` 使用 COM 自动化操作）

---

## 安装方法

Cat5Dev 以 `.vsix` 文件形式分发。安装步骤如下：

1. **下载 `.vsix` 文件**  
   从发布页面下载 `cat5dev-x.x.x.vsix`（或更晚版本）。

2. **在 VSCode 中安装**  
   在 VSCode 中打开扩展视图（`Ctrl+Shift+X`），单击右上角的 `⋯` 菜单，然后选择 **从 VSIX 安装... (Install from VSIX...)**。

   ![安装步骤](./resources/installation.png)

3. **选择文件**  
   选择已下载的 `.vsix` 文件并确认安装。

### ⚠️ 重要提示：首次使用前请备份 CATIA 设置

在首次使用 Cat5Dev 之前，**请务必备份您的 CATIA V5 设置文件夹**：

- **文件夹路径**: `%APPDATA%\Dassault Systemes\CATSettings`
- **原因**: 在极少数情况下，CATIA 设置可能会损坏或配置错误。备份允许您在需要时恢复设置。

**备份步骤：**
```
1. 按下 Win+R 并输入: %APPDATA%
2. 导航至 Dassault Systemes\CATSettings
3. 右键单击并选择“复制”
4. 将其粘贴到安全的位置（如桌面、OneDrive 等）
5. 重命名备份文件夹（例如 “CATSettings_backup_2026-04-18”）
```

---

## 文件结构

```
<workspace>/
├── .gitignore                      # 由 cat5dev.init 生成
├── cat5dev.toml                    # 配置文件（由 cat5dev.init 生成）
└── modules/
    ├── MyModule.bas_utf            # 标准模块 (Standard Module)
    ├── MyClass.cls_utf             # 类模块 (Class Module)
    └── MyForm.frm_utf              # 用户窗体 (UserForm)
```

> 配置保存在 `cat5dev.toml` 的 `[project]` 部分下：
> - `target_project` — 目标 CATIA VBA 项目的名称
> - `language` — `"ja"` 表示日语，`"en"` 表示英语，`"zh"` 表示中文（默认：`"ja"`）

---

## 快速入门

1. 安装本扩展
2. 在 VSCode 中打开一个工作区文件夹
3. 单击活动栏中的 **Cat5Dev** 图标
4. 单击 ⚙️ 选择要同步的目标 CATIA VBA 项目
5. 单击 ☁️↓ **拉取 (Pull)**，将模块从 CATIA 导入到 `modules/` 文件夹中
6. 在 VSCode 中编辑您的 VBA 代码
7. 单击 ☁️↑ **推送 (Push)**，将更改同步回 CATIA

---

## 重要提示：推送 (Push) 后必须在 CATIA 的 VBA 编辑器中进行保存

在从 VSCode 推送之后，**您必须在 CATIA 的 VBA 编辑器中保存项目**（在 VBA 编辑器内按 `Ctrl+S`，或点击 菜单 > 保存）。

推送操作是将模块代码写入 CATIA 内存中的 VBE。如果在未保存的情况下关闭 CATIA，所有推送的更改都将丢失。

```
VSCode (编辑) → 推送 (Push) → 在 CATIA 的 VBA 编辑器中保存 ✅
                                   ↓
                             更改将持久保留到 CATIA 文档中
```

---

## 注意事项

- VBA 文件以 UTF-8 编码存储，并带有 `_utf` 后缀，以便与 CATIA 的原生 Shift-JIS 导出进行区分。
- 本扩展通过 COM (`MSAPC.Apc`) 来操作 CATIA V5 的 VBE (Visual Basic 编辑器)。
- 目前不支持在 VSCode 内进行调试运行；请使用 CATIA 自身的 VBA IDE 进行调试。

---

## 为什么选择 VBA？

我曾尝试使用多种语言进行 CATIA V5 宏开发，但出于以下几个原因，我最终还是回到了 VBA：

- **原生支持** — VBA 直接内置于 CATIA V5 中。
- **最佳的开发环境** — VBA 编辑器 (VBE) 依然是最完整和最稳定的选择。
- **出乎意料的快速** — 在许多情况下，VBA 的性能表现优于其他替代方案。
- **紧凑的发布** — 不需要外部运行时或额外的依赖关系。

我*喜欢* VBA 这种语言吗？  
嗯……那就是完全另外一个问题了。

---

## 关于编码的注意事项（请提供帮助）

我只使用过日语版的 CATIA V5，所以只熟悉从 VBA 编辑器导出的、以 Shift-JIS 保存的文件。由于这非常不方便，Cat5Dev 在拉取/推送过程中会自动将编码转换为 UTF-8。

如果您在您的环境中使用此扩展时遇到乱码，您可能需要调整编码设置。很遗憾，我没有办法在日语环境之外验证其行为。

如果您遇到任何问题，欢迎随时联系我。

---

## 开源协议 (License)

MIT
