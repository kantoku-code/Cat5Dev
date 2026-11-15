# Cat5Dev

[English](../README.md) | 日本語 | [简体中文](README.zh-CN.md)

Cat5Dev は、VSCode と CATIA V5 の間で VBA モジュールを同期する VSCode 拡張機能です。

---

## 機能

### VBA モジュールの Pull / Push
CATIA V5 とローカルワークスペース間で VBA モジュールを同期します。

- **Pull** — 対象 CATIA プロジェクトからすべての VBA モジュールをローカルの `modules/` フォルダにエクスポート
- **Push** — ローカルの VBA モジュールを CATIA にインポート（変更されたファイルのみ転送し、未変更のモジュールはスキップして高速化）

### 対象プロジェクトの選択
クイックピックダイアログから同期する CATIA VBA プロジェクトを選択できます。
選択したプロジェクト名は `cat5dev.toml` の `[project]` セクションに保存され、セッションをまたいで維持されます。

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
拡張機能は**日本語**、**英語（English）**、**中国語（简体中文）**に対応しています。
- パネルタイトルバーの 🌐 ボタンをクリックして言語を切り替え
- 言語設定は `cat5dev.toml` の `[project]` セクションに保存されます
- すべての UI メッセージ、TreeView ラベル、ツールチップが翻訳されます
- 新しい言語設定を適用するには VSCode の再読み込みが必要です

### シンボルナビゲーション
VBA シンボルが認識され、VSCode のナビゲーション機能に公開されます：

- **ブレッドクラム** — カーソル位置の `Sub` / `Function` / `Property` 名を表示
- **アウトラインビュー** — 現在のファイルのすべてのプロシージャとプロパティを一覧表示
- **シンボルへ移動** (`Ctrl+Shift+O`) — 任意のプロシージャに直接ジャンプ

認識されるシンボルタイプ： `Sub`、`Function`、`Property Get/Let/Set`、`Type`、`Enum`

**ブレッドクラムの並び順（`cat5dev.toml`）：**

```toml
[editor]
breadcrumbs_sort = "name"   # position = 記述順 / name = 名前順 / type = 種類順
```

値はワークスペース設定（`.vscode/settings.json`）の `breadcrumbs.symbolSortOrder` に書き込まれ、このワークスペースの全ファイルに適用されます。キーが無い場合はワークスペース設定を変更しません。

### VBA Lint

`.bas_utf`・`.cls_utf`・`.frm_utf` ファイルの編集中、リアルタイムで VBA Lint が実行されます。
診断結果はエディタ上の波線と **Problems** パネルに表示されます。

**ルール一覧：**

| コード | ルール | 重要度 |
|--------|--------|--------|
| VBA001 | `Option Explicit` が宣言されていない | 警告 |
| VBA002 | `On Error Resume Next` の使用 | 警告 |
| VBA003 | `GoTo` の使用（`On Error GoTo` は除外） | 警告 |
| VBA004 | 最大行長超過（デフォルト: 100文字） | 警告 |
| VBA005 | `Dim` 宣言後に一度も使用されない変数 | 警告 |
| VBA006 | ネスト深さが閾値を超過（デフォルト: 5） | 警告 |
| VBA007 | Sub/Function の行数が閾値を超過（デフォルト: 300行） | 警告 |
| VBA008 | 括弧の不一致 | エラー |
| VBA009 | `End If` / `End Sub` / `End Function` 等の End 忘れ | エラー |

**設定（`cat5dev.toml`）：**

Lint は**デフォルトで無効**です。有効にするには、ワークスペースルートに `cat5dev.toml` ファイルを作成します：

1. コマンドパレットから **`cat5dev.init`** を実行するか、
2. 以下の内容で `cat5dev.toml` を手動作成してください：

```toml
[project]
target_project = ""
language = "ja"

[lint]
enabled = true

[lint.rules]
option_explicit       = true
on_error_resume_next  = false
goto                  = false
max_line_length       = 100   # 0 = 無効
unused_variables      = true
max_nesting_depth     = 5     # 0 = 無効
max_function_lines    = 300   # 0 = 無効
unmatched_parens      = true
unmatched_blocks      = true

[formatter]
enabled = false
# ... （フォーマッタセクション参照）
```

`cat5dev.toml` の変更は保存後すぐに反映されます（再読み込み不要）。

### VBA フォーマッタ（`Shift+Alt+F` / 保存時）

`Shift+Alt+F` またはファイル保存時に VBA コードフォーマッタを実行できます。

フォーマッタは**デフォルトで無効**です。`cat5dev.toml` で `enabled = true` を設定して有効化します：

```toml
[formatter]
enabled = true
```

**有効時のデフォルト処理：**
- キーワード大文字化（`dim` → `Dim`、`end if` → `End If` など）
- インデント修正（`Sub`、`If`、`For`、`With`、`Select Case` などのブロック）
- 行末スペース除去
- 継続行マーカー前スペース追加（`Show(_` → `Show( _`）
- 継続行インデント（`_` で終わる行の次行は +1 インデント、ただし `)` のみの行は除外）
- 空行正規化（連続2行まで・プロシージャ間に空行1行を保証）

**オプション（デフォルト無効）：**

| `cat5dev.toml` キー | 説明 |
|----------------------|------|
| `normalize_operator_spacing` | `x=1+2` → `x = 1 + 2` |
| `normalize_comma_spacing` | `foo(a,b)` → `foo(a, b)` |
| `normalize_comment_space` | `'comment` → `' comment` |
| `expand_type_suffixes` | `Dim x%` → `Dim x As Integer` |

**`cat5dev.toml` フォーマッタ設定の全項目：**

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

## 要件

- CATIA V5 が起動しており、VBA プロジェクトが開いていること
- Windows（`cscript.exe` 経由の COM オートメーションを使用）

---

## インストール

Cat5Dev は `.vsix` ファイルとして配布されています。以下の手順でインストールしてください：

1. **`.vsix` ファイルをダウンロード**  
   `cat5dev-x.x.x.vsix` （またはそれ以降）をリリースページから取得します。

2. **VSCode にインストール**  
   VSCode の拡張機能ビュー（`Ctrl+Shift+X`）を開き、右上の `⋯` メニューをクリックして **VSIX からインストール** を選択します。

   ![インストール手順](../resources/installation.png)

3. **ファイルを選択**  
   ダウンロードした `.vsix` ファイルを選択して、インストールを確認します。

### ⚠️ 重要：初回使用前に CATIA 設定をバックアップしてください

Cat5Dev を初めて使用する前に、**CATIA V5 の設定フォルダをバックアップ**してください：

- **フォルダパス**: `%APPDATA%\Dassault Systemes\CATSettings`
- **理由**: まれに CATIA の設定が破損または不適切な状態になることがあります。バックアップがあれば、必要に応じて設定を復元できます。

**バックアップ手順：**
```
1. Win+R キーを押して「%APPDATA%」と入力
2. Dassault Systemes\CATSettings に移動
3. 右クリックして「コピー」を選択
4. 安全な場所（デスクトップ、OneDrive など）に貼り付け
5. バックアップフォルダの名前を変更（例：「CATSettings_backup_2026-04-18」）
```

---

## ファイル構成

```
<workspace>/
├── .gitignore                      # cat5dev.init で生成
├── cat5dev.toml                    # 設定ファイル（cat5dev.init で生成）
└── modules/
    ├── MyModule.bas_utf            # 標準モジュール
    ├── MyClass.cls_utf             # クラスモジュール
    └── MyForm.frm_utf              # ユーザーフォーム
```

> 設定は `cat5dev.toml` の `[project]` セクションに保存されます：
> - `target_project` — 対象 CATIA VBA プロジェクトの名前
> - `language` — `"ja"` は日本語、`"en"` は英語、`"zh"` は中国語（デフォルト：`"ja"`）
> - `encoding` — VBA コード転送時の文字エンコーディング（例：`"gbk"`、`"shift_jis"`、`"utf-8"`）。中国語（`"zh"`）では `"gbk"`、日本語（`"ja"`）と英語（`"en"`）では `"shift_jis"` がデフォルトです。

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

## エンコーディングに関する注意

VBA ファイルは VSCode（UTF-8）と CATIA V5（Shift-JIS や GBK などのネイティブ ANSI）の間で同期されます。
デフォルトでは、設定された言語に応じて転送エンコーディングが自動的に決まります：
- 中国語（`"zh"`）は `"gbk"`
- 日本語（`"ja"`）と英語（`"en"`）は `"shift_jis"`

`cat5dev.toml` の `[project]` セクションで `encoding` を設定すると変更できます（例：`encoding = "gb2312"`、`encoding = "shift_jis"`）。問題があればお気軽にお知らせください。

---

## ライセンス

MIT
