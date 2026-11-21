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

- If "CATIA is not running. Start CATIA and try again." appears, start CATIA V5 and try again.
- If "Open the CATIA VBA editor (Alt+F11) once, then try again." appears, open the VBA editor in CATIA once and try again (you may close it afterwards). Until the VBA editor has been opened, Cat5Dev cannot see the VBA projects.
- Make sure the target VBA project is open in CATIA. If it is not, "Target project not found" is logged.
- Make sure `target_project` in `cat5dev.toml` matches the project name in CATIA (case-sensitive).
- For details, check "CATIA VBA Sync" in the Output panel.

## "Cat5Dev could not prepare its bridge project (C5D_Bridge)" appears

Cat5Dev could not prepare `C5D_Bridge`, the helper project it loads into CATIA ([how it works](sync.md#how-cat5dev-talks-to-catia)). Nothing has been written to your VBA projects.

1. Check "CATIA VBA Sync" in the Output panel for details.
2. If a dialog (such as a MsgBox) is open in CATIA, close it, wait a moment and try again. This error can appear while the previous operation has not finished yet.
3. Make sure `%LOCALAPPDATA%\Cat5Dev\bridge\` is writable (permissions, antivirus software, and so on).
4. If the problem persists, save your work in CATIA, restart CATIA and try again.

## A project named `C5D_Bridge` appears in the VBA editor

This is the helper project that Cat5Dev loads temporarily while an operation is running. It is removed when the operation ends. It is read-only and is never saved. Do not edit or save it; leave it alone.

It can remain after an operation (for example after a timeout); it is removed automatically at the next project selection, Pull or Push, and also disappears when CATIA exits.

## Files remain or keep increasing in `%LOCALAPPDATA%\Cat5Dev\bridge\`

This is normal. Each time Cat5Dev talks to CATIA, it creates a file named `C5D_Bridge-<8 characters>-<number>-<date and time>.catvba` in this folder and has CATIA load it. CATIA keeps this file in use until CATIA exits, even after the operation ends, so the files cannot be deleted while CATIA is running and a new one is added at each operation.

- After CATIA exits, they are deleted automatically at the next project selection, Pull or Push.
- Most of them are hard links to the original file, so they use almost no disk space (File Explorer shows each of them as about 155 KB).
- To delete them right away, exit CATIA and delete them yourself. You can even delete the whole folder; the files that are needed are created again at the next operation. While CATIA is running, files in use cannot be deleted.
- If they are not deleted even after CATIA exits, check the lines starting with `[bridge]` in "CATIA VBA Sync" in the Output panel. Files recorded as `kept (locked)` were kept because they were still in use.

For details, see [Files in the `bridge` folder](sync.md#files-in-the-bridge-folder).

## "Another CATIA operation is running…" or "CATIA completion could not be confirmed" appears

Only one operation can talk to CATIA at a time. Check that you are not running Pull / Push in another VSCode window.

If the CATIA operation does not finish within 60 seconds (for example, because a MsgBox is left open in CATIA), "CATIA completion could not be confirmed. …" appears and the sync lock is held to prevent double writes.

1. Close any dialog open in CATIA, and make sure the previous operation has finished
2. In the same VSCode window, run project selection, Pull or Push again
3. In the warning "The previous CATIA operation could not be confirmed, so the synchronization lock is held. …", click "Release Lock", then click "Release Lock" again in the confirmation dialog
4. Run the original operation again

Do not release the lock while CATIA may still be working; the changes may be written twice. If "Release Lock" is not shown or the lock cannot be released, see [Releasing the sync lock](sync.md#if-completion-cannot-be-confirmed).

## Pull stops because of unsaved edits

Pull does not run if any file in `modules/` is unsaved. Save with `Ctrl+K S` (Save All) and try again.

## Pull failed and left a `.cat5dev-pull-XXXXXX` folder

Pull failed partway and the restore did not complete either. The original files are in the `backup/` folder inside the folder shown in the error message. Restore them to `modules/`, then delete this folder.

## Pushed changes do not appear in CATIA / disappeared after restarting CATIA

Push only writes to CATIA's memory. **Save the project in the CATIA VBA editor.**

## Cannot push a UserForm

If CATIA has no UserForm with the same name, it cannot be created and is skipped. Create an empty UserForm with the same name in the CATIA VBA editor, then push.

## Text is garbled

Pull / Push carry the text as Unicode, so characters that the system ANSI code page of Windows can represent (for example Japanese on Japanese Windows, where it is Shift-JIS) are not garbled. Characters that the ANSI code page cannot represent, such as emoji (for example `🙂`), are replaced with `?` when VBA stores them in the module. This is how VBA works and the same happens when you type them in the VBA editor, so it cannot be avoided. `encoding` in `cat5dev.toml` is no longer used (it is ignored if present). If files pulled with an earlier version are garbled, pull again. See [Character encoding](sync.md#character-encoding).

## No Lint squiggles / the formatter does not run

- Check that `cat5dev.toml` has `[lint] enabled = true` and `[formatter] enabled = true` (both default to `false`).
- Check that the file extension is `.bas_utf` / `.cls_utf` / `.frm_utf`.
- Check "CATIA VBA Sync" in the Output panel for errors.

## Code changes every time I save

When `[formatter] enabled = true`, files are formatted on save in either of these cases. If you do not want it, turn off both.

- `format_on_save = true` in `cat5dev.toml`
- VSCode's `editor.formatOnSave` is enabled (files are formatted regardless of `format_on_save`)

## Completion or hover does not appear

- Variables declared `As Object` / `As Variant` get no completion. Declare the type or add a [type hint `'@As`](intellisense.md#type-hints-as).
- `CATIA.ActiveDocument.` not showing `PartDocument` members is by design. Declare `Dim oDoc As PartDocument`.
- If procedures from other modules do not appear, check that the file is directly under `modules/` (subfolders are ignored).
- Check "VBA Language Server" in the Output panel for errors.

## No descriptions for CATIA types

The bundled type information does not include descriptions. Build and configure a [database with descriptions](intellisense.md#showing-api-descriptions).
