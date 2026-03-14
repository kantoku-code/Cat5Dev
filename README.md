# Cat5Dev

A VSCode extension for synchronizing VBA modules between VSCode and CATIA V5.

---

## Features

### Pull / Push VBA Modules
Sync VBA modules between CATIA V5 and your local workspace.

- **Pull** — Export all VBA modules from the target CATIA project to the local `modules/` folder
- **Push** — Import local VBA modules back into CATIA (only changed files are transferred, unchanged modules are skipped for speed)

### Target Project Selection
Select the CATIA VBA project to sync with via a quick-pick dialog.
The selected project name is saved in `catia-vba.json` and persists across sessions.

### Module Tree View
A dedicated panel in the Activity Bar shows the current state of your workspace.

- Displays the target project name
- Modules are grouped by type:
  - **Standard Modules** (`.bas_utf`)
  - **Class Modules** (`.cls_utf`)
  - **UserForms** (`.frm_utf`)
- Click any module to open it in the editor
- The tree updates automatically when the target project is changed

### Toolbar Buttons
Quick-access buttons are available in the panel title bar:

| Icon | Action |
|------|--------|
| ☁️↓ | Pull from CATIA |
| ☁️↑ | Push to CATIA |
| ⚙️ | Select target project |
| 🔄 | Refresh module list |

### Symbol Navigation
VBA symbols are recognized and exposed to VSCode's navigation features:

- **Breadcrumb** — Shows the current `Sub` / `Function` / `Property` name as you move the cursor
- **Outline view** — Lists all procedures and properties in the current file
- **Go to Symbol** (`Ctrl+Shift+O`) — Jump directly to any procedure

Recognized symbol types: `Sub`, `Function`, `Property Get/Let/Set`, `Type`, `Enum`

---

## Requirements

- CATIA V5 must be running with a VBA project open
- Windows (uses COM automation via `cscript.exe`)

---

## File Structure

```
<workspace>/
├── catia-vba.json                  # Target project settings
├── .catia-vba-push-cache.json      # Push optimization cache (auto-generated)
└── modules/
    ├── MyModule.bas_utf            # Standard Module
    ├── MyClass.cls_utf             # Class Module
    └── MyForm.frm_utf              # UserForm
```

> Both `.catia-vba-push-cache.json` and `modules/` are automatically added to `.gitignore` on first use.

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

## Push Optimization (Cache)

On each Push, a hash of every module is computed and stored in `.catia-vba-push-cache.json`.
On subsequent Pushes, modules whose content has not changed are skipped automatically, significantly reducing sync time for large projects.

The cache is keyed per project name, so switching between projects with **Select Target Project** always triggers a full push for the new project.

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

## License

MIT
