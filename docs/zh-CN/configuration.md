# 配置参考

[← 返回 README](README.zh-CN.md)

Cat5Dev 的配置主要写在放置于工作区根目录的 `cat5dev.toml` 中。

- [创建 cat5dev.toml](#创建-cat5devtoml)
- [cat5dev.toml 的全部选项](#cat5devtoml-的全部选项)
- [显示语言](#显示语言)
- [VSCode 设置与环境变量](#vscode-设置与环境变量)
- [生成的文件和文件夹](#生成的文件和文件夹)

---

## 创建 cat5dev.toml

在命令面板（`Ctrl+Shift+P`）中运行 **`CATIA: Initialize Project (cat5dev.toml)`**，会创建以下文件并在编辑器中打开。

| 文件 | 内容 |
|------|------|
| `cat5dev.toml` | 以默认值写入全部选项的配置文件（Lint 和格式化程序处于禁用状态） |
| `.gitignore` | Git 的忽略设置（见下文） |

- 如果 `cat5dev.toml` 已存在，会确认是否覆盖。即使覆盖，`target_project` 的值也会被保留。
- 如果 `.gitignore` 已存在，会另外确认是否覆盖。即使选择不覆盖，`cat5dev.toml` 也会被创建。
- `cat5dev.toml` 不会被自动创建。**保存目标项目需要 `cat5dev.toml`**，因此请首先运行此命令。

<details>
<summary>生成的 <code>.gitignore</code> 的内容</summary>

```gitignore
Thumbs.db
.vscode/
*.log
logs/
.cache/
tmp/
temp/
*.zip
*.tar
*.gz
*.tmp
*.bak
.claude/
cat5dev.toml
```

`cat5dev.toml` 每次修改设置都会产生差异，因此不纳入 Git 管理。如果它已经被提交，仅写入 `.gitignore` 不会生效，请先运行 `git rm --cached cat5dev.toml`，再提交该变更。
</details>

---

## cat5dev.toml 的全部选项

对 `cat5dev.toml` 的更改在保存后立即生效。

```toml
# Cat5Dev 配置文件

[project]
target_project = ""          # 目标 CATIA VBA 项目名称（通过 ⚙️ 按钮设置）
language = "ja"              # 显示语言 ja / en / zh

[lint]
enabled = false              # true 为启用 Lint

[lint.rules]
option_explicit = true
on_error_resume_next = false
goto = false
max_line_length = 100
unused_variables = true
max_nesting_depth = 5
max_function_lines = 300
unmatched_parens = true
unmatched_blocks = true

[formatter]
enabled = false              # true 为启用格式化程序
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
format_on_pull = false

[editor]
breadcrumbs_sort = "name"    # 面包屑导航的排序方式 position / name / type
```

### [project]
| 键 | 默认值 | 说明 |
|-----|--------|------|
| `target_project` | `""` | 要同步的 CATIA VBA 项目名称。通过 ⚙️ 按钮选择后会写入此处 |
| `language` | `"ja"` | 显示语言（`ja` / `en` / `zh`）。通过 🌐 按钮切换后会写入此处 |

> [!NOTE]
> 旧版本中的 `encoding`（与 CATIA 交换数据时使用的字符编码）已不再使用。Pull / Push 以 Unicode 形式传输正文。旧的 `cat5dev.toml` 中即使保留该项也会被忽略（参见 [字符编码](sync.md#字符编码)）。

### [lint] / [lint.rules]
请参阅 [Lint 与格式化程序](lint-format.md#vba-lint)。

### [formatter]
请参阅 [Lint 与格式化程序](lint-format.md#vba-格式化程序)。

### [editor]
| 键 | 默认值 | 说明 |
|-----|--------|------|
| `breadcrumbs_sort` | （没有该键时不做更改） | 面包屑导航和大纲的排序方式。`position`（代码顺序）/ `name`（按名称）/ `type`（按类型）。会写入工作区设置 `.vscode/settings.json` 的 `breadcrumbs.symbolSortOrder` |

### 书写注意事项
- 从 `#` 到行尾的内容被视为注释。**值的 `"…"` 中不能包含 `#`。**
- `true` / `false` 不区分大小写。

---

## 显示语言

通知和对话框的消息可以用日语、英语、中文（简体）显示。通过面板顶部的 🌐 按钮（`CATIA: Switch Language`）切换。

- 切换后消息会立即使用新语言（无需重新加载）。
- 树视图的显示名称（Modules 等）、命令名称和 Lint 消息不会被翻译。

语言按以下顺序确定。

1. `cat5dev.toml` 中的 `language`
2. VSCode 设置 `cat5dev.language`（仅限用户自行设置时）
3. 如果 VSCode 的显示语言为中文，则为 `zh`
4. 其他情况为 `ja`

通过 🌐 按钮选择的语言，如果存在 `cat5dev.toml` 则保存到其中，否则保存到工作区的 `.vscode/settings.json`（`cat5dev.language`）中。

---

## VSCode 设置与环境变量

### VSCode 设置
| 设置 | 默认值 | 说明 |
|------|--------|------|
| `cat5dev.language` | — | 显示语言。`cat5dev.toml` 中的 `language` 优先 |

> [!NOTE]
> 格式化程序的设置写在 `cat5dev.toml` 的 `[formatter]` 中。为保持兼容，仅当 `cat5dev.toml` 中没有写 `format_on_pull` 时，仍会读取旧的 VSCode 设置 `cat5dev.formatter.formatOnPull`。

### 环境变量（面向高级用户）
| 环境变量 | 说明 |
|----------|------|
| `CAT5DEV_TLB_DB` | CATIA API 类型信息数据库的路径。使用包含说明文字的数据库时设置（参见 [IntelliSense](intellisense.md#catia-api-类型信息)） |
| `CAT5DEV_VBA_LS` | 语言服务器 `vba-ls.exe` 的路径。用于开发 |

更改环境变量后，请重启 VSCode。

---

## 生成的文件和文件夹

### 工作区内
```
<工作区>/
├── cat5dev.toml               配置文件（由 Initialize Project 创建）
├── .gitignore                 由 Initialize Project 创建
├── .vscode/settings.json      保存 breadcrumbs_sort，以及没有 cat5dev.toml 时的显示语言
├── modules/
│   ├── Module1.bas_utf        标准模块
│   ├── Class1.cls_utf         类模块
│   └── UserForm1.frm_utf      用户窗体（仅代码）
└── .cat5dev-pull-XXXXXX/      仅在 Pull 期间创建的临时文件夹（仅在失败且无法恢复时保留）
```

- 带 `_utf` 的扩展名，用于与 CATIA 以 Shift-JIS 等编码导出的文件相区别。其内容为 UTF-8。
- 如果 `.cat5dev-pull-XXXXXX/` 仍然保留，请在从其中的 `backup/` 恢复文件之前不要删除它（参见 [同步](sync.md#失败时的安全措施)）。

### 工作区外
| 位置 | 内容 |
|------|------|
| `%LOCALAPPDATA%\Cat5Dev\operation-lock\` | 同步锁的信息（参见 [同步锁](sync.md#防止并发执行与同步锁)） |
| `%LOCALAPPDATA%\Cat5Dev\bridge\` | 临时加载到 CATIA 中的中介项目 `C5D_Bridge` 的文件。每次处理所用的文件会在 CATIA 退出后的下一次处理时删除（参见 [与 CATIA 通信的机制](sync.md#与-catia-通信的机制)） |
| `%TEMP%\cat5dev-…\` | 与 CATIA 交换数据时使用的临时文件。通常在处理结束时删除 |
