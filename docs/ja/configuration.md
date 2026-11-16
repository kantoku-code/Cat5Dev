# 設定リファレンス

[← README に戻る](README.ja.md)

Cat5Dev の設定は、主にワークスペースのルートに置く `cat5dev.toml` に書きます。

- [cat5dev.toml の作成](#cat5devtoml-の作成)
- [cat5dev.toml の全項目](#cat5devtoml-の全項目)
- [表示言語](#表示言語)
- [VSCode の設定と環境変数](#vscode-の設定と環境変数)
- [作成されるファイルとフォルダ](#作成されるファイルとフォルダ)

---

## cat5dev.toml の作成

コマンドパレット（`Ctrl+Shift+P`）から **`CATIA: Initialize Project (cat5dev.toml)`** を実行すると、次のファイルを作成してエディタで開きます。

| ファイル | 内容 |
|----------|------|
| `cat5dev.toml` | 全項目を既定値で書いた設定ファイル（Lint とフォーマッタは無効の状態） |
| `.gitignore` | Git 用の除外設定（下記） |

- `cat5dev.toml` がすでにある場合は、上書きするか確認します。上書きしても `target_project` の値は引き継ぎます。
- `.gitignore` がすでにある場合は、別途上書きするか確認します。「上書きしない」を選んでも `cat5dev.toml` は作成されます。
- `cat5dev.toml` が自動で作られることはありません。**対象プロジェクトの保存には `cat5dev.toml` が必要**なので、最初に実行してください。

<details>
<summary>作成される <code>.gitignore</code> の内容</summary>

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

## cat5dev.toml の全項目

`cat5dev.toml` の変更は保存するとすぐに反映されます。

```toml
# Cat5Dev 設定ファイル

[project]
target_project = ""          # 対象の CATIA VBA プロジェクト名（⚙️ ボタンで設定）
language = "ja"              # 表示言語 ja / en / zh
encoding = "shift_jis"       # CATIA とやり取りするときの文字コード

[lint]
enabled = false              # true で Lint を有効にする

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
enabled = false              # true でフォーマッタを有効にする
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
format_on_save = true

[editor]
breadcrumbs_sort = "name"    # パンくずリストの並び順 position / name / type
```

### [project]
| キー | 既定値 | 説明 |
|------|--------|------|
| `target_project` | `""` | 同期する CATIA の VBA プロジェクト名。⚙️ ボタンで選ぶと書き込まれます |
| `language` | `"ja"` | 表示言語（`ja` / `en` / `zh`）。🌐 ボタンで切り替えると書き込まれます |
| `encoding` | `language` が `zh` なら `"gbk"`、それ以外は `"shift_jis"` | CATIA とやり取りするときの文字コード。[同期](sync.md#文字コードエンコーディング) を参照 |

### [lint] / [lint.rules]
[Lint とフォーマッタ](lint-format.md#vba-lint) を参照してください。

### [formatter]
[Lint とフォーマッタ](lint-format.md#vba-フォーマッタ) を参照してください。

### [editor]
| キー | 既定値 | 説明 |
|------|--------|------|
| `breadcrumbs_sort` | （キーが無ければ変更しない） | パンくずリスト・アウトラインの並び順。`position`（記述順）/ `name`（名前順）/ `type`（種類順）。ワークスペース設定 `.vscode/settings.json` の `breadcrumbs.symbolSortOrder` に書き込まれます |

### 書き方の注意
- `#` から行末まではコメントとして扱います。**値の `"…"` の中に `#` を含めることはできません。**
- `true` / `false` は小文字・大文字どちらでも構いません。

---

## 表示言語

通知やダイアログのメッセージを日本語・英語・中国語（簡体字）で表示できます。パネル上部の 🌐 ボタン（`CATIA: Switch Language`）で切り替えます。

- 切り替えるとメッセージはすぐに新しい言語になります（「再読み込みしてください」と表示されますが、再読み込みしなくても反映されます）。
- ツリーの表示名（Modules など）、コマンド名、Lint のメッセージは翻訳されません。

言語は次の順に決まります。

1. `cat5dev.toml` の `language`
2. VSCode の設定 `cat5dev.language`（自分で設定した場合のみ）
3. VSCode の表示言語が中国語なら `zh`
4. それ以外は `ja`

🌐 ボタンで選んだ言語は、`cat5dev.toml` があればそこに、無ければワークスペースの `.vscode/settings.json`（`cat5dev.language`）に保存されます。

---

## VSCode の設定と環境変数

### VSCode の設定
| 設定 | 既定値 | 説明 |
|------|--------|------|
| `cat5dev.language` | — | 表示言語。`cat5dev.toml` の `language` が優先されます |
| `cat5dev.formatter.formatOnPull` | `false` | Pull で取り込んだファイルを整形する（`cat5dev.toml` の `[formatter] enabled = true` も必要） |

> [!NOTE]
> 設定画面には `cat5dev.formatter.indentSize` などのフォーマッタ設定も表示されますが、現在は使われていません。フォーマッタの設定は `cat5dev.toml` の `[formatter]` に書いてください。

### 環境変数（上級者向け）
| 環境変数 | 説明 |
|----------|------|
| `CAT5DEV_TLB_DB` | CATIA API の型情報データベースのパス。説明文付きのデータベースを使う場合に設定します（[IntelliSense](intellisense.md#catia-api-の型情報) を参照） |
| `CAT5DEV_VBA_LS` | 言語サーバー `vba-ls.exe` のパス。開発用です |

環境変数を変えた後は VSCode を再起動してください。

---

## 作成されるファイルとフォルダ

### ワークスペース内
```
<ワークスペース>/
├── cat5dev.toml               設定ファイル（Initialize Project で作成）
├── .gitignore                 Initialize Project で作成
├── .vscode/settings.json      breadcrumbs_sort や、cat5dev.toml が無いときの表示言語を保存
├── modules/
│   ├── Module1.bas_utf        標準モジュール
│   ├── Class1.cls_utf         クラスモジュール
│   └── UserForm1.frm_utf      ユーザーフォーム（コードのみ）
└── .cat5dev-pull-XXXXXX/      Pull 中だけ作られる一時フォルダ（失敗して復元できなかった場合のみ残る）
```

- `_utf` の付いた拡張子は、CATIA が Shift-JIS などで書き出すファイルと区別するためのものです。中身は UTF-8 です。
- `.cat5dev-pull-XXXXXX/` が残っている場合は、中の `backup/` からファイルを復元するまで削除しないでください（[同期](sync.md#失敗したときの安全策) を参照）。

### ワークスペースの外
| 場所 | 内容 |
|------|------|
| `%LOCALAPPDATA%\Cat5Dev\operation-lock\` | 同期ロックの情報（[同期ロック](sync.md#同時実行の防止と同期ロック) を参照） |
| `%TEMP%\cat5dev-…\` | CATIA とのやり取りに使う一時ファイル。通常は処理の終了時に削除されます |
