# Cat5Dev

Cat5Dev is a VSCode extension that syncs VBA modules between VSCode and CATIA V5.

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
The extension supports both **Japanese (日本語)** and **English**.
- Click the 🌐 button in the panel title bar to switch languages
- The language preference is saved in `.vscode/settings.json`
- All UI messages, TreeView labels, and tooltips will be translated
- A reload of VSCode is required to apply the new language setting

### Symbol Navigation
VBA symbols are recognized and exposed to VSCode's navigation features:

- **Breadcrumb** — Shows the current `Sub` / `Function` / `Property` name as you move the cursor
- **Outline view** — Lists all procedures and properties in the current file
- **Go to Symbol** (`Ctrl+Shift+O`) — Jump directly to any procedure

Recognized symbol types: `Sub`, `Function`, `Property Get/Let/Set`, `Type`, `Enum`

### VBA Formatter (on save / `Shift+Alt+F`)

A built-in VBA code formatter runs automatically on save (or via `Shift+Alt+F`).

**Default formatting:**
- Keyword capitalization (`dim` → `Dim`, `end if` → `End If`, etc.)
- Indentation correction (blocks: `Sub`, `If`, `For`, `With`, `Select Case`, etc.)
- Trailing whitespace removal
- Space before continuation marker (`Show(_` → `Show( _`)
- Continuation line indentation (`_` lines get +1 extra indent; closing-only lines like `)` are exempt)
- Blank line normalization (max 2 consecutive blank lines; blank line guaranteed between procedures)

**Optional formatting (disabled by default):**

| Option flag | Description |
|-------------|-------------|
| `-normalize-operator-spacing` | `x=1+2` → `x = 1 + 2` |
| `-normalize-comma-spacing` | `foo(a,b)` → `foo(a, b)` |
| `-normalize-comment-space` | `'comment` → `' comment` |
| `-normalize-then-placement` | `If x _\n    Then` → `If x Then` |
| `-split-colon-statements` | `Dim x: x = 1` → two lines |
| `-expand-type-suffixes` | `Dim x%` → `Dim x As Integer` |
| `-normalize-on-error` | On Error style normalization |

---

## Requirements

- CATIA V5 must be running with a VBA project open
- Windows (uses COM automation via `cscript.exe`)

---

## File Structure

```
<workspace>/
├── .vscode/
│   └── settings.json               # Target project & language settings
├── .gitignore                      # Auto-generated to exclude cache
├── .catia-vba-push-cache.json      # Push optimization cache (auto-generated)
└── modules/
    ├── MyModule.bas_utf            # Standard Module
    ├── MyClass.cls_utf             # Class Module
    └── MyForm.frm_utf              # UserForm
```

> The `.catia-vba-push-cache.json` entry is automatically added to `.gitignore` on first use (pull, push, or project selection).
> Settings are stored in `.vscode/settings.json` with keys:
> - `targetProject` — The name of the target CATIA VBA project
> - `language` — `"ja"` for Japanese, `"en"` for English (default: `"ja"`)

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

## Why VBA?

I’ve experimented with CATIA V5 macro development in several languages, but there are a few reasons why I ultimately returned to VBA:

- **Native support** — VBA is built directly into CATIA V5
- **Best development environment** — The VBE is still the most complete and stable option
- **Surprisingly fast** — In many cases, VBA outperformed alternatives
- **Compact distribution** — No external runtimes or dependencies required

Do I *like* VBA as a language?  
Well… that’s a different question entirely.

---

## Notes on Encoding (Please Help)

I have only ever used the Japanese version of CATIA V5, so I am familiar only with files exported from the VBA Editor being saved in Shift‑JIS. Because this was quite inconvenient, Cat5Dev automatically converts the encoding to UTF‑8 during the Pull/Push process.

If you encounter garbled characters when using this extension in your environment, you may need to adjust the encoding settings. Unfortunately, I have no way to verify its behavior outside of a Japanese environment.

If you run into any issues, please feel free to let me know.

---

## License

MIT

---
---
Cat5Dev は、VSCode と CATIA V5 の間で VBA モジュールを同期する VSCode 拡張機能です。

---

## 機能

### VBA モジュールの Pull / Push
CATIA V5 とローカルワークスペース間で VBA モジュールを同期します。

- **Pull** — 対象 CATIA プロジェクトからすべての VBA モジュールをローカルの `modules/` フォルダにエクスポート
- **Push** — ローカルの VBA モジュールを CATIA にインポート（変更されたファイルのみ転送し、未変更のモジュールはスキップして高速化）

### 対象プロジェクトの選択
クイックピックダイアログから同期する CATIA VBA プロジェクトを選択できます。
選択したプロジェクト名は `catia-vba.json` に保存され、セッションをまたいで維持されます。

### モジュールツリービュー
アクティビティバーに専用パネルを表示し、ワークスペースの現在の状態を確認できます。

- 対象プロジェクト名を表示
- モジュールはタイプ別にグループ化：
  - **標準モジュール** (`.bas_utf`)
  - **クラスモジュール** (`.cls_utf`)
  - **ユーザーフォーム** (`.frm_utf`)
- モジュールをクリックするとエディタで開く
- 対象プロジェクトを変更するとツリーが自動更新
- モジュール上で右クリックするとコンテキストメニューが表示：
  - **名前変更** — モジュールファイルの名前を変更
  - **削除** — モジュールファイルを削除（確認あり）
  - **パスをコピー** — ファイルの完全パスをクリップボードにコピー

### ツールバーボタン
パネルタイトルバーにクイックアクセスボタンが表示されます：

| アイコン | 操作 |
|----------|------|
| ☁️↓ | CATIA から Pull |
| ☁️↑ | CATIA へ Push |
| ⚙️ | 対象プロジェクトを選択 |
| 🔄 | モジュール一覧を更新 |
| 🌐 | 言語を切り替え |

### 多言語対応
拡張機能は**日本語（日本語）**と**英語（English）**に対応しています。
- パネルタイトルバーの 🌐 ボタンをクリックして言語を切り替え
- 言語設定は `.vscode/settings.json` に保存されます
- すべての UI メッセージ、TreeView ラベル、ツールチップが翻訳されます
- 新しい言語設定を適用するには VSCode の再読み込みが必要です

### シンボルナビゲーション
VBA シンボルが認識され、VSCode のナビゲーション機能に公開されます：

- **ブレッドクラム** — カーソル位置の `Sub` / `Function` / `Property` 名を表示
- **アウトラインビュー** — 現在のファイルのすべてのプロシージャとプロパティを一覧表示
- **シンボルへ移動** (`Ctrl+Shift+O`) — 任意のプロシージャに直接ジャンプ

認識されるシンボルタイプ： `Sub`、`Function`、`Property Get/Let/Set`、`Type`、`Enum`

### VBA フォーマッタ（保存時 / `Shift+Alt+F`）

保存時または `Shift+Alt+F` で VBA コードフォーマッタが自動実行されます。

**デフォルトで有効な処理：**
- キーワード大文字化（`dim` → `Dim`、`end if` → `End If` など）
- インデント修正（`Sub`、`If`、`For`、`With`、`Select Case` などのブロック）
- 行末スペース除去
- 継続行マーカー前スペース追加（`Show(_` → `Show( _`）
- 継続行インデント（`_` で終わる行の次行は +1 インデント、ただし `)` のみの行は除外）
- 空行正規化（連続2行まで・プロシージャ間に空行1行を保証）

**オプション（デフォルト無効）：**

| オプションフラグ | 説明 |
|----------------|------|
| `-normalize-operator-spacing` | `x=1+2` → `x = 1 + 2` |
| `-normalize-comma-spacing` | `foo(a,b)` → `foo(a, b)` |
| `-normalize-comment-space` | `'comment` → `' comment` |
| `-normalize-then-placement` | `If x _\n    Then` → `If x Then` |
| `-split-colon-statements` | `Dim x: x = 1` → 2行に分割 |
| `-expand-type-suffixes` | `Dim x%` → `Dim x As Integer` |
| `-normalize-on-error` | On Error スタイル統一 |

---

## 要件

- CATIA V5 が起動しており、VBA プロジェクトが開いていること
- Windows（`cscript.exe` 経由の COM オートメーションを使用）

---

## ファイル構成

```
<workspace>/
├── .vscode/
│   └── settings.json               # 対象プロジェクトと言語の設定
├── .gitignore                      # キャッシュを除外するよう自動生成
├── .catia-vba-push-cache.json      # Push 最適化キャッシュ（自動生成）
└── modules/
    ├── MyModule.bas_utf            # 標準モジュール
    ├── MyClass.cls_utf             # クラスモジュール
    └── MyForm.frm_utf              # ユーザーフォーム
```

> `.catia-vba-push-cache.json` は、初回使用時（pull、push、またはプロジェクト選択）に自動的に `.gitignore` に追加されます。
> 設定は `.vscode/settings.json` に以下のキーで保存されます：
> - `targetProject` — 対象 CATIA VBA プロジェクトの名前
> - `language` — `"ja"` は日本語、`"en"` は英語（デフォルト：`"ja"`）

---

## はじめに

1. 拡張機能をインストール
2. VSCode でワークスペースフォルダを開く
3. アクティビティバーの **Cat5Dev** アイコンをクリック
4. ⚙️ をクリックして対象の CATIA VBA プロジェクトを選択
5. ☁️↓ **Pull** をクリックして CATIA からモジュールを `modules/` フォルダにインポート
6. VSCode で VBA コードを編集
7. ☁️↑ **Push** をクリックして変更を CATIA に同期

---

## Push 最適化（キャッシュ）

Push のたびに各モジュールのハッシュが計算され、`.catia-vba-push-cache.json` に保存されます。
次回以降の Push では、内容が変更されていないモジュールは自動的にスキップされるため、大規模なプロジェクトでも同期時間を大幅に短縮できます。

キャッシュはプロジェクト名ごとに管理されているため、**対象プロジェクトの選択**で別のプロジェクトに切り替えると、新しいプロジェクトに対して常にフル Push が実行されます。

---

## 重要：Push 後は CATIA の VBA エディタで保存してください

VSCode から Push した後は、**CATIA の VBA エディタでプロジェクトを保存**してください（VBA エディタ内で ツール > 保存 または `Ctrl+S`）。

Push は CATIA のインメモリ VBE にモジュールコードを書き込みます。保存せずに CATIA を閉じると、Push した変更はすべて失われます。

```
VSCode（編集）→ Push → CATIA の VBA エディタで保存 ✅
                              ↓
                        変更が CATIA ドキュメントに永続化される
```

---

## 注意事項

- VBA ファイルは UTF-8 エンコーディングで保存され、CATIA のネイティブな Shift-JIS エクスポートと区別するために `_utf` サフィックスが付きます
- 本拡張機能は COM（`MSAPC.Apc`）経由で CATIA V5 の VBE（Visual Basic エディタ）を操作します
- VSCode 内でのデバッグ実行は現在サポートされていません。デバッグには CATIA 付属の VBA IDE をご利用ください

---

## なぜ VBA なのか

CATIA V5 のマクロ開発をいくつかの言語で試してきましたが、最終的に VBA に戻ってきた理由がいくつかあります：

- **ネイティブサポート** — VBA は CATIA V5 に直接組み込まれています
- **最良の開発環境** — VBE は今でも最も充実した安定した選択肢です
- **驚くほど速い** — 多くのケースで VBA が他の選択肢を上回りました
- **コンパクトな配布** — 外部ランタイムや依存関係が不要です

VBA という言語が好きかどうか？
それはまた別の話です。

---


## ライセンス

MIT
