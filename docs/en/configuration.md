# Configuration Reference

[← Back to README](../../README.md)

Cat5Dev is configured mainly through `cat5dev.toml` at the workspace root.

- [Creating cat5dev.toml](#creating-cat5devtoml)
- [All cat5dev.toml settings](#all-cat5devtoml-settings)
- [Display language](#display-language)
- [VSCode settings and environment variables](#vscode-settings-and-environment-variables)
- [Files and folders that are created](#files-and-folders-that-are-created)

---

## Creating cat5dev.toml

Run **`CATIA: Initialize Project (cat5dev.toml)`** from the Command Palette (`Ctrl+Shift+P`) to create the following files and open them in the editor.

| File | Contents |
|------|----------|
| `cat5dev.toml` | A settings file with every setting at its default value (Lint and the formatter disabled) |
| `.gitignore` | Git ignore rules (below) |

- If `cat5dev.toml` already exists, you are asked whether to overwrite it. The `target_project` value is kept even if you overwrite.
- If `.gitignore` already exists, you are asked separately whether to overwrite it. `cat5dev.toml` is created even if you choose not to overwrite `.gitignore`.
- `cat5dev.toml` is never created automatically. **Saving the target project requires `cat5dev.toml`**, so run this command first.

<details>
<summary>Contents of the generated <code>.gitignore</code></summary>

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
```
</details>

---

## All cat5dev.toml settings

Changes to `cat5dev.toml` take effect as soon as you save the file.

```toml
# Cat5Dev configuration file

[project]
target_project = ""          # Target CATIA VBA project name (set with the ⚙️ button)
language = "ja"              # Display language: ja / en / zh
encoding = "shift_jis"       # Encoding used when exchanging code with CATIA

[lint]
enabled = false              # true to enable Lint

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
enabled = false              # true to enable the formatter
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
breadcrumbs_sort = "name"    # Breadcrumbs sort order: position / name / type
```

### [project]
| Key | Default | Description |
|-----|---------|-------------|
| `target_project` | `""` | Name of the CATIA VBA project to sync. Written when you select it with the ⚙️ button |
| `language` | `"ja"` | Display language (`ja` / `en` / `zh`). Written when you switch with the 🌐 button |
| `encoding` | `"gbk"` if `language` is `zh`, otherwise `"shift_jis"` | Encoding used when exchanging code with CATIA. See [Sync](sync.md#character-encoding) |

### [lint] / [lint.rules]
See [Lint and Formatter](lint-format.md#vba-lint).

### [formatter]
See [Lint and Formatter](lint-format.md#vba-formatter).

### [editor]
| Key | Default | Description |
|-----|---------|-------------|
| `breadcrumbs_sort` | (no change if the key is absent) | Sort order of the breadcrumbs and Outline. `position` (source order) / `name` (by name) / `type` (by kind). Written to `breadcrumbs.symbolSortOrder` in the workspace settings `.vscode/settings.json` |

### Syntax notes
- Everything from `#` to the end of the line is treated as a comment. **A value in `"…"` cannot contain `#`.**
- `true` / `false` can be written in lowercase or uppercase.

---

## Display language

Notifications and dialog messages can be shown in Japanese, English, or Chinese (Simplified). Switch with the 🌐 button at the top of the panel (`CATIA: Switch Language`).

- After switching, messages appear in the new language immediately (no reload needed).
- Tree labels (Modules, etc.), command names, and Lint messages are not translated.

The language is determined in this order:

1. `language` in `cat5dev.toml`
2. The VSCode setting `cat5dev.language` (only if you set it yourself)
3. `zh` if the VSCode display language is Chinese
4. Otherwise `ja`

The language selected with the 🌐 button is saved to `cat5dev.toml` if it exists, or otherwise to `.vscode/settings.json` (`cat5dev.language`) in the workspace.

---

## VSCode settings and environment variables

### VSCode settings
| Setting | Default | Description |
|---------|---------|-------------|
| `cat5dev.language` | — | Display language. `language` in `cat5dev.toml` takes precedence |

> [!NOTE]
> Formatter settings go in the `[formatter]` section of `cat5dev.toml`. For compatibility, the old VSCode setting `cat5dev.formatter.formatOnPull` is still read, but only when `format_on_pull` is not written in `cat5dev.toml`.

### Environment variables (advanced)
| Variable | Description |
|----------|-------------|
| `CAT5DEV_TLB_DB` | Path to the CATIA API type information database. Set it to use a database with descriptions (see [IntelliSense](intellisense.md#catia-api-type-information)) |
| `CAT5DEV_VBA_LS` | Path to the language server `vba-ls.exe`. For development |

Restart VSCode after changing environment variables.

---

## Files and folders that are created

### In the workspace
```
<workspace>/
├── cat5dev.toml               Settings file (created by Initialize Project)
├── .gitignore                 Created by Initialize Project
├── .vscode/settings.json      Stores breadcrumbs_sort, and the display language when there is no cat5dev.toml
├── modules/
│   ├── Module1.bas_utf        Standard module
│   ├── Class1.cls_utf         Class module
│   └── UserForm1.frm_utf      UserForm (code only)
└── .cat5dev-pull-XXXXXX/      Temporary folder created only during Pull (remains only if Pull failed and could not restore)
```

- The `_utf` suffix in the extensions distinguishes these files from files that CATIA exports in Shift-JIS and similar encodings. The content is UTF-8.
- If `.cat5dev-pull-XXXXXX/` remains, do not delete it until you have restored your files from the `backup/` folder inside it (see [Sync](sync.md#safety-net-on-failure)).

### Outside the workspace
| Location | Contents |
|----------|----------|
| `%LOCALAPPDATA%\Cat5Dev\operation-lock\` | Sync lock information (see [Sync lock](sync.md#preventing-concurrent-operations-and-the-sync-lock)) |
| `%TEMP%\cat5dev-…\` | Temporary files used to exchange data with CATIA. Normally deleted when the operation ends |
