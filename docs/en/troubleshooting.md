# Troubleshooting

[← Back to README](../../README.md)

## Check first: the Output panel

Error details are logged in VSCode's Output panel (**View** → **Output** in the menu, or `Ctrl+Shift+U`). Select one of the following in the drop-down at the top right.

| Output channel | Contents |
|----------------|----------|
| **CATIA VBA Sync** | Communication with CATIA (project selection, Pull, Push), and Lint / formatter logs. Opens automatically on errors |
| **VBA Language Server** | Logs from the language server that provides completion, hover, etc. |

---

## Nothing is shown in the tree

- **No target project is set.** Select a project with the ⚙️ button.
- **There is no `cat5dev.toml`.** Without `cat5dev.toml`, the selected project is not saved and the tree stays empty. Run `CATIA: Initialize Project (cat5dev.toml)` from the Command Palette, then select the project with ⚙️ again.
- **There are no files in `modules/`.** Pull with ☁️↓.

## Cannot connect to CATIA (Pull / Push / project selection fails)

- Make sure CATIA V5 is running. If CATIA is not running, messages such as "Pull failed. Check the Output panel for details." appear.
- Make sure the target VBA project is open in CATIA. If it is not, "Target project not found" is logged.
- Make sure `target_project` in `cat5dev.toml` matches the project name in CATIA (case-sensitive).
- For details, check "CATIA VBA Sync" in the Output panel.

## "Another CATIA operation is running…" or "CATIA completion could not be confirmed" appears

Only one operation can talk to CATIA at a time. Check that you are not running Pull / Push in another VSCode window.

If an operation does not finish within 60 seconds, the lock stays in place to prevent double writes. Follow the [sync lock recovery steps](sync.md#if-completion-cannot-be-confirmed).

## Pull stops because of unsaved edits

Pull does not run if any file in `modules/` is unsaved. Save with `Ctrl+K S` (Save All) and try again.

## Pull failed and left a `.cat5dev-pull-XXXXXX` folder

Pull failed partway and the restore did not complete either. The original files are in the `backup/` folder inside the folder shown in the error message. Restore them to `modules/`, then delete this folder.

## Pushed changes do not appear in CATIA / disappeared after restarting CATIA

Push only writes to CATIA's memory. **Save the project in the CATIA VBA editor.**

## Cannot push a UserForm

If CATIA has no UserForm with the same name, it cannot be created and is skipped. Create an empty UserForm with the same name in the CATIA VBA editor, then push.

## Text is garbled

Check `encoding` in the `[project]` section of `cat5dev.toml`. `shift_jis` is typical for the Japanese version of CATIA, and `gbk` for the Chinese version. See [Character encoding](sync.md#character-encoding).

## No Lint squiggles / the formatter does not run

- Check that `cat5dev.toml` has `[lint] enabled = true` and `[formatter] enabled = true` (both default to `false`).
- Check that the file extension is `.bas_utf` / `.cls_utf` / `.frm_utf`.
- Check "CATIA VBA Sync" in the Output panel for errors.

## Code changes every time I save

When `[formatter] enabled = true`, formatting on save (`format_on_save`) is enabled by default. If you do not want it, add `format_on_save = false` to `cat5dev.toml`.

## Completion or hover does not appear

- Variables declared `As Object` / `As Variant` get no completion. Declare the type or add a [type hint `'@As`](intellisense.md#type-hints-as).
- `CATIA.ActiveDocument.` not showing `PartDocument` members is by design. Declare `Dim oDoc As PartDocument`.
- If procedures from other modules do not appear, check that the file is directly under `modules/` (subfolders are ignored).
- Check "VBA Language Server" in the Output panel for errors.

## No descriptions for CATIA types

The bundled type information does not include descriptions. Build and configure a [database with descriptions](intellisense.md#showing-api-descriptions).
