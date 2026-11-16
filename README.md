# Cat5Dev

English | [日本語](docs/ja/README.ja.md) | [简体中文](docs/zh-CN/README.zh-CN.md)

Cat5Dev is a VSCode extension for developing CATIA V5 VBA macros in VSCode.
It transfers modules between a CATIA VBA project and VSCode, so you can edit them in VSCode with completion, lint, and formatting.

<img src="docs/en/images/workflow.png" alt="Edit in VSCode, send to CATIA with Push, then save in the CATIA VBA editor" width="760">

---

## Features

| Feature | Summary | Details |
|---------|---------|---------|
| **Pull / Push** | Sync modules between the CATIA VBA project and the `modules/` folder. Only changed modules are sent, and you are asked before anything is deleted | [Sync](docs/en/sync.md) |
| **Module tree** | Shows modules by type in a side panel. Rename, delete, and copy path | [Sync](docs/en/sync.md#module-tree) |
| **IntelliSense** | Code completion including the CATIA API, hover, Go to Definition, parameter hints, and detection of nonexistent members | [IntelliSense](docs/en/intellisense.md) |
| **Symbol navigation** | Outline, breadcrumbs, and `Ctrl+Shift+O` to jump to a procedure | [IntelliSense](docs/en/intellisense.md#symbol-navigation) |
| **VBA Lint** | 9 rules, such as missing `Option Explicit`, unused variables, and missing `End If` (disabled by default) | [Lint and Formatter](docs/en/lint-format.md#vba-lint) |
| **VBA Formatter** | Keyword capitalization, indentation, and whitespace / blank line cleanup (disabled by default) | [Lint and Formatter](docs/en/lint-format.md#vba-formatter) |
| **Multilingual UI** | 日本語, English, 简体中文 | [Configuration](docs/en/configuration.md#display-language) |

<img src="docs/en/images/screenshot-completion.png" alt="CATIA API completion" width="560">

---

## Requirements

- Windows (64-bit)
- CATIA V5 running, with the VBA project you want to sync open
- VSCode 1.85 or later

Cat5Dev communicates with CATIA over COM through the Windows built-in `cscript.exe` (32-bit).

---

## Installation

Cat5Dev is distributed as a `.vsix` file.

1. **Download the `.vsix` file**\
   Get `cat5dev-x.x.x.vsix` from the [Releases page](https://github.com/kantoku-code/Cat5Dev/releases).

2. **Install it in VSCode**\
   Open the Extensions view (`Ctrl+Shift+X`), and choose **Install from VSIX...** from the `⋯` menu at the top right.

   ![Installation steps](docs/en/images/installation.png)

3. **Select the file**\
   Select the downloaded `.vsix` file to install it.

> [!WARNING]
> **Before using Cat5Dev for the first time, back up your CATIA V5 settings folder.**\
> In rare cases, CATIA settings can become corrupted. With a backup, you can restore them.
>
> 1. Press `Win+R` and enter `%APPDATA%`
> 2. Open the `Dassault Systemes\CATSettings` folder
> 3. Copy the folder and paste it somewhere safe (Desktop, OneDrive, etc.)
> 4. Rename it (e.g. `CATSettings_backup_2026-04-18`)

---

## Getting Started

1. Open a working folder in VSCode.
2. Run **`CATIA: Initialize Project (cat5dev.toml)`** from the Command Palette (`Ctrl+Shift+P`).\
   This creates the settings file `cat5dev.toml` and a `.gitignore`.
3. Start CATIA V5 and open the VBA project you want to sync.
4. Click the **CATIA V5 VBA** icon in the Activity Bar to open the panel.
5. Click the ⚙️ button and select the target VBA project.
6. Click the ☁️↓ button (Pull) to import the modules into the `modules/` folder.
7. Edit the code in VSCode.
8. Click the ☁️↑ button (Push) to send your changes to CATIA.
9. **Save the project in the CATIA VBA editor** (see the note below).

<img src="docs/en/images/screenshot-panel.png" alt="Cat5Dev panel" width="560">

> [!TIP]
> Lint and the formatter are disabled by default. To use them, set `enabled = true` in the `[lint]` and `[formatter]` sections of `cat5dev.toml` (see [Lint and Formatter](docs/en/lint-format.md)).

---

## Important: Save in the CATIA VBA Editor After Pushing

Push only writes to the VBA project in CATIA's memory. **If you close CATIA without saving, all pushed changes are lost.**\
After pushing from VSCode, save in the CATIA VBA editor (`Ctrl+S`).

Push may also ask **whether to delete modules from CATIA that do not exist in VSCode**. Check the list carefully before you choose (see [Sync](docs/en/sync.md#push-vscode--catia)).

---

## Documentation

| Page | Contents |
|------|----------|
| [Syncing with CATIA](docs/en/sync.md) | How Pull / Push work and their confirmation dialogs, the module tree, the sync lock and recovery, character encoding |
| [IntelliSense](docs/en/intellisense.md) | Completion, hover, Go to Definition, parameter hints, the `'@As` type hint, CATIA API type information |
| [Lint and Formatter](docs/en/lint-format.md) | Lint rules, formatter operations and options |
| [Configuration Reference](docs/en/configuration.md) | All `cat5dev.toml` settings, display language, VSCode settings, files that are created |
| [Troubleshooting](docs/en/troubleshooting.md) | Common problems and fixes, reading the Output panel |

---

## Limitations

- You cannot debug in VSCode. Use the CATIA VBA editor for debugging.
- Only the code of UserForms is handled. Edit the form design (control layout) in the CATIA VBA editor.
- Push cannot create new UserForms. Create an empty form with the same name in CATIA first.

---

## Why VBA?

I have tried developing CATIA V5 macros in several languages, but I came back to VBA for a few reasons:

- **Native support** — VBA is built directly into CATIA V5
- **Best development environment** — The VBE is still the most complete and stable option
- **Surprisingly fast** — In many cases VBA outperformed the alternatives
- **Compact distribution** — No external runtime or dependencies required

Do I like VBA as a language?\
That is another story.

---

## License

MIT
