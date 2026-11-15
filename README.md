# Cat5Dev

English | [日本語](docs/README.ja.md) | [简体中文](docs/README.zh-CN.md)

Cat5Dev is a VSCode extension that syncs VBA modules between VSCode and CATIA V5.

---

## Features

### Pull / Push VBA Modules
Sync VBA modules between CATIA V5 and your local workspace.

- **Pull** — Export all VBA modules from the target CATIA project to the local `modules/` folder
- **Push** — Import local VBA modules back into CATIA (only changed files are transferred, unchanged modules are skipped for speed)

### Target Project Selection
Select the CATIA VBA project to sync with via a quick-pick dialog.
The selected project name is saved in `cat5dev.toml` under the `[project]` section and persists across sessions.

### Module Tree View
A dedicated panel in the Activity Bar shows the current state of your workspace.

- Displays the target project name
- Modules are grouped by type:
  - **Standard Modules** (`.bas_utf`)
  - **Class Modules** (`.cls_utf`)
  - **UserForms** (`.frm_utf`)
- Click any module to open it in the editor
- The tree updates automatically when the target project is changed
- Right-click any module to access context menu options:
  - **Rename** — Rename the module file
  - **Delete** — Delete the module file (with confirmation)
  - **Copy Path** — Copy the full file path to clipboard

### Toolbar Buttons
Quick-access buttons are available in the panel title bar:

| Icon | Action |
|------|--------|
| ☁️↓ | Pull from CATIA |
| ☁️↑ | Push to CATIA |
| ⚙️ | Select target project |
| 🔄 | Refresh module list |
| 🌐 | Switch language |

### Multi-Language Support
The extension supports **Japanese (日本語)**, **English**, and **Chinese (简体中文)**.
- Click the 🌐 button in the panel title bar to switch languages
- The language preference is saved in `cat5dev.toml` under the `[project]` section
- All UI messages, TreeView labels, and tooltips will be translated
- A reload of VSCode is required to apply the new language setting

### Symbol Navigation
VBA symbols are recognized and exposed to VSCode's navigation features:

- **Breadcrumb** — Shows the current `Sub` / `Function` / `Property` name as you move the cursor
- **Outline view** — Lists all procedures and properties in the current file
- **Go to Symbol** (`Ctrl+Shift+O`) — Jump directly to any procedure

Recognized symbol types: `Sub`, `Function`, `Property Get/Let/Set`, `Type`, `Enum`

**Breadcrumb symbol order (`cat5dev.toml`):**

```toml
[editor]
breadcrumbs_sort = "name"   # position = source order / name = by name / type = by kind
```

The value is written to the workspace settings (`.vscode/settings.json`) as `breadcrumbs.symbolSortOrder`, so it applies to every file in this workspace. If the key is absent, the workspace setting is left untouched.

### VBA Linter

A built-in VBA linter runs in real time as you edit `.bas_utf`, `.cls_utf`, and `.frm_utf` files.
Diagnostics are displayed inline as squiggly underlines and shown in the **Problems** panel.

**Rules:**

| Code | Rule | Severity |
|------|------|----------|
| VBA001 | `Option Explicit` not declared | Warning |
| VBA002 | `On Error Resume Next` usage | Warning |
| VBA003 | `GoTo` usage (excluding `On Error GoTo`) | Warning |
| VBA004 | Line exceeds maximum length (default: 100) | Warning |
| VBA005 | Variable declared with `Dim` but never used | Warning |
| VBA006 | Nesting depth exceeds threshold (default: 5) | Warning |
| VBA007 | Sub/Function exceeds line count threshold (default: 300) | Warning |
| VBA008 | Mismatched parentheses | Error |
| VBA009 | Missing `End If` / `End Sub` / `End Function` etc. | Error |

**Configuration (`cat5dev.toml`):**

Linting is **disabled by default**. To enable it, create a `cat5dev.toml` file in your workspace root:

1. Run **`cat5dev.init`** from the Command Palette to generate a configuration template, or
2. Manually create `cat5dev.toml` with the following structure:

```toml
[project]
target_project = ""
language = "en"

[lint]
enabled = true

[lint.rules]
option_explicit       = true
on_error_resume_next  = false
goto                  = false
max_line_length       = 100   # 0 = disabled
unused_variables      = true
max_nesting_depth     = 5     # 0 = disabled
max_function_lines    = 300   # 0 = disabled
unmatched_parens      = true
unmatched_blocks      = true

[formatter]
enabled = false
# ... (see Formatter section below)
```

Changes to `cat5dev.toml` take effect immediately on save — no reload required.

### VBA Formatter (`Shift+Alt+F` / on save)

A built-in VBA code formatter is available on demand (`Shift+Alt+F`) or automatically on save.

The formatter is **disabled by default**. To enable it, set `enabled = true` in `cat5dev.toml`:

```toml
[formatter]
enabled = true
```

**Default formatting (when enabled):**
- Keyword capitalization (`dim` → `Dim`, `end if` → `End If`, etc.)
- Indentation correction (blocks: `Sub`, `If`, `For`, `With`, `Select Case`, etc.)
- Trailing whitespace removal
- Space before continuation marker (`Show(_` → `Show( _`)
- Continuation line indentation (`_` lines get +1 extra indent; closing-only lines like `)` are exempt)
- Blank line normalization (max 2 consecutive blank lines; blank line guaranteed between procedures)

**Optional formatting (disabled by default):**

| `cat5dev.toml` key | Description |
|---------------------|-------------|
| `normalize_operator_spacing` | `x=1+2` → `x = 1 + 2` |
| `normalize_comma_spacing` | `foo(a,b)` → `foo(a, b)` |
| `normalize_comment_space` | `'comment` → `' comment` |
| `expand_type_suffixes` | `Dim x%` → `Dim x As Integer` |

**Full `cat5dev.toml` formatter configuration:**

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

## Requirements

- CATIA V5 must be running with a VBA project open
- Windows (uses COM automation via `cscript.exe`)

---

## Installation

Cat5Dev is distributed as a `.vsix` file. Install it as follows:

1. **Download the `.vsix` file**  
   Obtain `cat5dev-x.x.x.vsix` (or later) from the release page.

2. **Install in VSCode**  
   In VSCode, open the Extensions view (`Ctrl+Shift+X`), then click the `⋯` menu in the top-right corner and select **Install from VSIX...**

   ![Installation Steps](./resources/installation.png)

3. **Select the file**  
   Choose the downloaded `.vsix` file and confirm installation.

### ⚠️ Important: Backup CATIA Settings Before First Use

Before using Cat5Dev for the first time, **backup your CATIA V5 settings folder**:

- **Folder path**: `%APPDATA%\Dassault Systemes\CATSettings`
- **Why**: In rare cases, CATIA settings can become corrupted or misconfigured. Having a backup allows you to restore settings if needed.

**Backup procedure:**
```
1. Press Win+R and type: %APPDATA%
2. Navigate to Dassault Systemes\CATSettings
3. Right-click and select "Copy"
4. Paste it in a safe location (e.g., Desktop, OneDrive, etc.)
5. Rename the backup folder (e.g., "CATSettings_backup_2026-04-18")
```

---

## File Structure

```
<workspace>/
├── .gitignore                      # Generated by cat5dev.init
├── cat5dev.toml                    # Configuration file (generated by cat5dev.init)
└── modules/
    ├── MyModule.bas_utf            # Standard Module
    ├── MyClass.cls_utf             # Class Module
    └── MyForm.frm_utf              # UserForm
```

> Settings are stored in `cat5dev.toml` under the `[project]` section:
> - `target_project` — The name of the target CATIA VBA project
> - `language` — `"ja"` for Japanese, `"en"` for English, `"zh"` for Chinese (default: `"ja"`)
> - `encoding` — The character encoding of VBA code transfer (e.g., `"gbk"`, `"shift_jis"`, `"utf-8"`). Defaults to `"gbk"` for Chinese (`"zh"`), and `"shift_jis"` for Japanese (`"ja"`) and English (`"en"`).

---

## Getting Started

1. Install the extension
2. Open a workspace folder in VSCode
3. Click the **Cat5Dev** icon in the Activity Bar
4. Click ⚙️ to select the target CATIA VBA project
5. Click ☁️↓ **Pull** to import modules from CATIA into the `modules/` folder
6. Edit your VBA code in VSCode
7. Click ☁️↑ **Push** to sync changes back to CATIA

---

## Important: Save in CATIA's VBA Editor after Push

After pushing from VSCode, **you must save the project in CATIA's VBA Editor** (Tools > Save, or `Ctrl+S` inside the VBA Editor).

Push writes the module code into CATIA's in-memory VBE. If CATIA is closed without saving, all pushed changes will be lost.

```
VSCode (edit) → Push → Save in CATIA's VBA Editor ✅
                         ↓
                       Changes are persisted in the CATIA document
```

---

## Notes

- VBA files are stored in UTF-8 encoding with a `_utf` suffix to distinguish them from CATIA's native Shift-JIS exports
- The extension targets CATIA V5's VBE (Visual Basic Editor) via COM (`MSAPC.Apc`)
- Debug execution within VSCode is not currently supported; use CATIA's own VBA IDE for debugging

---

## Why VBA?

I’ve experimented with CATIA V5 macro development in several languages, but there are a few reasons why I ultimately returned to VBA:

- **Native support** — VBA is built directly into CATIA V5
- **Best development environment** — The VBE is still the most complete and stable option
- **Surprisingly fast** — In many cases, VBA outperformed alternatives
- **Compact distribution** — No external runtimes or dependencies required

Do I *like* VBA as a language?  
Well… that’s a different question entirely.

---

## Notes on Encoding
VBA files are synchronized between VSCode (UTF-8) and CATIA V5 (native ANSI, e.g., Shift-JIS or GBK). 
By default, Cat5Dev automatically detects and sets the transfer encoding based on the configured language:
- `"gbk"` for Chinese (`"zh"`)
- `"shift_jis"` for Japanese (`"ja"`) and English (`"en"`)

You can customize this by setting the `encoding` property under the `[project]` section in `cat5dev.toml` (e.g., `encoding = "gb2312"` or `encoding = "shift_jis"`). If you run into any issues, please feel free to let me know.

---

## License

MIT
