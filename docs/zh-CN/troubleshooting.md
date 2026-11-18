# 故障排除

[← 返回 README](README.zh-CN.md)

## 首先检查：输出面板

错误的详细信息会记录在 VSCode 的输出面板中（菜单“查看”→“输出”，或 `Ctrl+Shift+U`）。请在右上角的下拉列表中选择以下通道。

| 输出通道 | 内容 |
|----------|------|
| **CATIA VBA Sync** | 与 CATIA 的通信（选择项目、Pull、Push），以及 Lint 和格式化程序的日志。出错时会自动打开 |
| **VBA Language Server** | 提供补全、悬停等功能的语言服务器的日志 |

---

## 树视图中没有显示任何内容

- **尚未设置目标项目。** 请通过 ⚙️ 按钮选择项目。
- **没有 `cat5dev.toml`。** 没有 `cat5dev.toml` 时，所选的项目不会被保存，树视图会保持为空。请在命令面板中运行 `CATIA: Initialize Project (cat5dev.toml)`，然后再次通过 ⚙️ 选择项目。
- **`modules/` 中没有文件。** 请通过 ☁️↓ 执行 Pull。

## 无法连接到 CATIA（Pull / Push / 选择项目失败）

- 请确认 CATIA V5 已启动。CATIA 未启动时，会显示“Pull 失败。”等消息。
- 请确认目标 VBA 项目已在 CATIA 中打开。未打开时，日志中会记录“Target project not found”。
- **请确认已启动 CATIA 的 VBA 编辑器（`Alt+F11`）。** 未启动 VBA 编辑器时无法执行 Pull / Push。
- 请确认 `cat5dev.toml` 中的 `target_project` 与 CATIA 中的项目名称一致（区分大小写）。
- 详细信息请查看输出面板的“CATIA VBA Sync”。

## 显示“另一个CATIA操作正在运行…”或“无法确认CATIA操作已完成”

与 CATIA 的通信同时只能执行一个。请确认没有在其他 VSCode 窗口中执行 Pull / Push。

如果处理未在 60 秒内结束，为防止重复写入，锁会保持锁定状态。请按照 [同步锁的恢复步骤](sync.md#显示无法确认完成时) 操作。

## Pull 因“检测到未保存的编辑…”而中止

如果 `modules/` 中有未保存的文件，Pull 将不会执行。请按 `Ctrl+K S`（全部保存）保存后再重新执行。

## Pull 失败后残留了 `.cat5dev-pull-XXXXXX` 文件夹

这是 Pull 中途失败、且恢复处理也未完成的状态。错误消息中显示的文件夹内的 `backup/` 中保存着原文件。请将其恢复到 `modules/` 后，再删除此文件夹。

## Push 后 CATIA 中没有反映／重启 CATIA 后更改消失了

Push 只会写入 CATIA 的内存。请**在 CATIA 的 VBA 编辑器中保存项目**。

## 无法 Push 用户窗体

如果 CATIA 中没有同名的用户窗体，则无法创建，会被跳过。请先在 CATIA 的 VBA 编辑器中创建同名的空用户窗体，然后再执行 Push。

## 出现乱码

请检查 `cat5dev.toml` 中 `[project]` 的 `encoding`。日文版 CATIA 通常使用 `shift_jis`，中文版通常使用 `gbk`。请参阅 [字符编码](sync.md#字符编码)。

## Lint 波浪线不显示／格式化程序不工作

- 请确认 `cat5dev.toml` 中已设置 `[lint] enabled = true`、`[formatter] enabled = true`（两者默认均为 `false`）。
- 请确认文件扩展名为 `.bas_utf` / `.cls_utf` / `.frm_utf`。
- 请确认输出面板的“CATIA VBA Sync”中没有错误。

## 每次保存时代码都会被改写

`[formatter] enabled = true` 时，以下任一情况都会在保存时格式化。如果不需要，请将两者都关闭。

- `cat5dev.toml` 中的 `format_on_save = true`
- 启用了 VSCode 设置 `editor.formatOnSave`（无论 `format_on_save` 如何都会格式化）

## 补全或悬停不显示

- `As Object` / `As Variant` 的变量不会显示补全。请声明类型，或添加 [类型提示 `'@As`](intellisense.md#类型提示-as)。
- `CATIA.ActiveDocument.` 不显示 `PartDocument` 的成员属于设计行为。请声明为 `Dim oDoc As PartDocument`。
- 如果其他模块的过程没有显示，请确认该文件位于 `modules/` 下一级（不包括子文件夹）。
- 请确认输出面板的“VBA Language Server”中没有错误。

## CATIA 的类型不显示说明文字

随附的类型信息中不包含说明文字。请生成并设置 [包含说明文字的数据库](intellisense.md#如需显示说明文字)。
