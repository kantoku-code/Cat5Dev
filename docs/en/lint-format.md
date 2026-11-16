# Lint and Formatter

[← Back to README](../../README.md)

Both are **disabled by default**. Enable them in `cat5dev.toml` at the workspace root (if you do not have `cat5dev.toml`, see [Configuration](configuration.md) to create it).

Changes to `cat5dev.toml` take effect as soon as you save the file (no need to reload VSCode).

---

## VBA Lint

Code is checked when a file is opened, when it is edited (0.5 seconds after you stop typing), and when it is saved. Problems are shown as squiggles in the editor and in the Problems panel. The source shown in the Problems panel is `cat5dev`.

<img src="images/screenshot-diagnostics.svg" alt="Lint squiggles and the Problems panel (image coming soon)" width="560">

### Enabling Lint
```toml
[lint]
enabled = true
```

> [!NOTE]
> The `cat5dev.toml` created by `CATIA: Initialize Project (cat5dev.toml)` contains `enabled = false`. Change it to `true` to use Lint.

### Rules
| Code | Description | Severity | `[lint.rules]` key | Default |
|------|-------------|----------|--------------------|---------|
| VBA001 | `Option Explicit` is not declared | Warning | `option_explicit` | `true` |
| VBA002 | `On Error Resume Next` is used | Warning | `on_error_resume_next` | `false` |
| VBA003 | `GoTo` is used (`On Error GoTo` is excluded) | Warning | `goto` | `false` |
| VBA004 | Line is too long | Warning | `max_line_length` | `100` (characters; `0` disables) |
| VBA005 | Variable declared with `Dim` is not used | Warning | `unused_variables` | `true` |
| VBA006 | Nesting is too deep | Warning | `max_nesting_depth` | `5` (`0` disables) |
| VBA007 | `Sub` / `Function` has too many lines | Warning | `max_function_lines` | `300` (lines; `0` disables) |
| VBA008 | Unmatched parentheses | Error | `unmatched_parens` | `true` |
| VBA009 | Missing `End If` / `End Sub` / `End Function`, etc. | Error | `unmatched_blocks` | `true` |

`true` / `false` rules are disabled with `false`, and numeric rules are disabled with `0`.

```toml
[lint.rules]
option_explicit      = true
on_error_resume_next = false
goto                 = false
max_line_length      = 100
unused_variables     = true
max_nesting_depth    = 5
max_function_lines   = 300
unmatched_parens     = true
unmatched_blocks     = true
```

### Notes
- Lint messages are always shown in Japanese, regardless of the display language setting.
- VBA005 checks only the first variable of a `Dim` statement (only `a` in `Dim a, b`). It covers `Dim` inside procedures; `Private` / `Public` / `Static` declarations are not checked.
- VBA006 counts the procedure itself as one level. `Select Case` counts as two levels.
- Nonexistent member detection (source `vba-ls`) is a separate feature from Lint. See [IntelliSense](intellisense.md#nonexistent-member-detection).

---

## VBA Formatter

Formats your code. It runs on VSCode's "Format Document" (`Shift+Alt+F`) and when a file is saved.

### Enabling the formatter
```toml
[formatter]
enabled = true
```

> [!IMPORTANT]
> Once you set `enabled = true`, **code is formatted every time you save**, because `format_on_save` defaults to `true`. If you do not want formatting on save, add `format_on_save = false`.

### What it does by default
| Operation | Example | Key |
|-----------|---------|-----|
| Capitalize keywords | `dim` → `Dim`, `end if` → `End If` | `capitalize_keywords` |
| Fix indentation | Blocks such as `Sub` / `If` / `For` / `Do` / `With` / `Select Case` | `fix_indentation` |
| Remove trailing whitespace | | `trim_trailing_space` |
| Add a space before the line continuation `_` | `Show(_` → `Show( _` | `ensure_continuation_space` |
| Indent continuation lines | Indents the line after a line ending in `_` by one level | `indent_continuation_lines` |
| Clean up blank lines | Limits consecutive blank lines to 2 | `max_blank_lines` |

- The indent unit is `indent_size` (default: 4 spaces).
- Inside `Select Case`, `Case` is indented one level and its body two levels. A single-line `If … Then …` is not treated as a block.
- Continuation line indentation works only when `fix_indentation = true`. Lines containing only `)` or `]` are not indented.
- Blank line cleanup inserts one blank line when the next procedure directly follows `End Sub`, etc. `max_blank_lines = 0` turns off blank line cleanup entirely.
- Strings and comments are not changed.
- Formatted files use CRLF line endings and end with a single newline.

### Options (disabled by default)
| Key | Example |
|-----|---------|
| `normalize_operator_spacing` | `x=1+2` → `x = 1 + 2` |
| `normalize_comma_spacing` | `foo(a,b)` → `foo(a, b)` |
| `normalize_comment_space` | `'comment` → `' comment` |
| `expand_type_suffixes` | `Dim x%` → `Dim x As Integer` |

### All keys and defaults
```toml
[formatter]
enabled = false                    # true to enable
indent_size = 4
capitalize_keywords = true
fix_indentation = true
trim_trailing_space = true
ensure_continuation_space = true
indent_continuation_lines = true
max_blank_lines = 2                # 0 = disabled
normalize_operator_spacing = false
normalize_comma_spacing = false
normalize_comment_space = false
expand_type_suffixes = false
format_on_save = true              # format on save
```

### When formatting runs
| When | Condition |
|------|-----------|
| `Shift+Alt+F` (Format Document) | `enabled = true` |
| On file save | `enabled = true` and `format_on_save = true` (independent of VSCode's `editor.formatOnSave`) |
| On Pull | `enabled = true` and the VSCode setting `cat5dev.formatter.formatOnPull` is `true` |

- Formatting only a selection ("Format Selection") is not supported. The whole file is always formatted.
- The `cat5dev.formatter.*` VSCode settings (other than `formatOnPull`) are currently not used. Put your settings in `cat5dev.toml`.

### Command-line usage
The formatter binary `vbafmt.exe` (in the `bin/` folder of the extension) can also be used as a standalone command. It takes one file (or standard input) and writes the formatted result to standard output. It does not read `cat5dev.toml`; specify options as arguments (run `vbafmt.exe -h` for the list).

```bat
vbafmt.exe modules\Module1.bas_utf > Module1_formatted.bas_utf
vbafmt.exe --normalize-comma-spacing modules\Module1.bas_utf
vbafmt.exe --check modules\Module1.bas_utf   :: exit code 1 if formatting is needed (for CI)
```
