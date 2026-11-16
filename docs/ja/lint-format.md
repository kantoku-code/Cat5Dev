# Lint とフォーマッタ

[← README に戻る](../README.ja.md)

どちらも**既定では無効**です。ワークスペースのルートにある `cat5dev.toml` で有効にします（`cat5dev.toml` が無い場合は [設定](configuration.md) を参照して作成してください）。

`cat5dev.toml` の変更は保存するとすぐに反映されます（VSCode の再読み込みは不要です）。

---

## VBA Lint

ファイルを開いたとき・編集したとき（入力が止まって 0.5 秒後）・保存したときにコードを検査し、問題をエディタの波線と「問題」パネルに表示します。「問題」パネルの出所（ソース）は `cat5dev` です。

<img src="../images/screenshot-diagnostics.svg" alt="Lint の波線と問題パネル（画像準備中）" width="560">

### 有効にする
```toml
[lint]
enabled = true
```

> [!NOTE]
> `CATIA: Initialize Project (cat5dev.toml)` で作成した `cat5dev.toml` は `enabled = false` になっています。Lint を使うには `true` に書き換えてください。

### ルール一覧
| コード | 内容 | 重要度 | `[lint.rules]` のキー | 既定値 |
|--------|------|--------|------------------------|--------|
| VBA001 | `Option Explicit` が宣言されていない | 警告 | `option_explicit` | `true` |
| VBA002 | `On Error Resume Next` を使っている | 警告 | `on_error_resume_next` | `false` |
| VBA003 | `GoTo` を使っている（`On Error GoTo` は除く） | 警告 | `goto` | `false` |
| VBA004 | 行が長すぎる | 警告 | `max_line_length` | `100`（文字。`0` で無効） |
| VBA005 | `Dim` で宣言した変数を使っていない | 警告 | `unused_variables` | `true` |
| VBA006 | ネスト（入れ子）が深すぎる | 警告 | `max_nesting_depth` | `5`（`0` で無効） |
| VBA007 | `Sub` / `Function` の行数が多すぎる | 警告 | `max_function_lines` | `300`（行。`0` で無効） |
| VBA008 | 括弧の対応が取れていない | エラー | `unmatched_parens` | `true` |
| VBA009 | `End If` / `End Sub` / `End Function` などが足りない | エラー | `unmatched_blocks` | `true` |

`true` / `false` のルールは `false` で無効、数値のルールは `0` で無効になります。

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

### 注意点
- メッセージは表示言語の設定にかかわらず日本語です。
- VBA005 は `Dim` の最初の変数だけを検査します（`Dim a, b` では `a` のみ）。プロシージャ内の `Dim` が対象で、`Private` / `Public` / `Static` の宣言は対象外です。
- VBA006 はプロシージャ自体を 1 段と数えます。`Select Case` は 2 段と数えます。
- 存在しないメンバーの検出（ソース `vba-ls`）は Lint とは別の機能です。[IntelliSense](intellisense.md#存在しないメンバーの検出) を参照してください。

---

## VBA フォーマッタ

コードの書式を整えます。VSCode の「ドキュメントのフォーマット」（`Shift+Alt+F`）と、ファイル保存時に動きます。

### 有効にする
```toml
[formatter]
enabled = true
```

> [!IMPORTANT]
> `enabled = true` にすると、`format_on_save` の既定値が `true` のため、**保存するたびに整形されます**。保存時の整形が不要な場合は `format_on_save = false` を書いてください。

### 既定で行う処理
| 処理 | 例 | キー |
|------|-----|------|
| キーワードの大文字化 | `dim` → `Dim`、`end if` → `End If` | `capitalize_keywords` |
| インデントの修正 | `Sub` / `If` / `For` / `Do` / `With` / `Select Case` などのブロック | `fix_indentation` |
| 行末の空白の削除 | | `trim_trailing_space` |
| 行継続 `_` の前に空白を入れる | `Show(_` → `Show( _` | `ensure_continuation_space` |
| 継続行のインデント | `_` で終わる行の次の行を 1 段下げる | `indent_continuation_lines` |
| 空行の整理 | 連続する空行を 2 行までにする | `max_blank_lines` |

- インデントの単位は `indent_size`（既定 4 スペース）です。
- `Select Case` の中は `Case` が 1 段、その中身が 2 段下がります。1 行で書いた `If … Then …` はブロックとして扱いません。
- 継続行のインデントは `fix_indentation = true` のときだけ働きます。`)` や `]` だけの行は下げません。
- 空行の整理では、`End Sub` などの直後に次のプロシージャが続いている場合に空行を 1 行入れます。`max_blank_lines = 0` にすると空行の整理をすべて行いません。
- 文字列とコメントの中は変更しません。
- 整形後のファイルは改行が CRLF になり、最後に改行が 1 つ付きます。

### オプション（既定では無効）
| キー | 例 |
|------|-----|
| `normalize_operator_spacing` | `x=1+2` → `x = 1 + 2` |
| `normalize_comma_spacing` | `foo(a,b)` → `foo(a, b)` |
| `normalize_comment_space` | `'comment` → `' comment` |
| `expand_type_suffixes` | `Dim x%` → `Dim x As Integer` |

### 全項目と既定値
```toml
[formatter]
enabled = false                    # true で有効
indent_size = 4
capitalize_keywords = true
fix_indentation = true
trim_trailing_space = true
ensure_continuation_space = true
indent_continuation_lines = true
max_blank_lines = 2                # 0 = 無効
normalize_operator_spacing = false
normalize_comma_spacing = false
normalize_comment_space = false
expand_type_suffixes = false
format_on_save = true              # 保存時に整形
```

### 整形のタイミング
| タイミング | 条件 |
|------------|------|
| `Shift+Alt+F`（ドキュメントのフォーマット） | `enabled = true` |
| ファイル保存時 | `enabled = true` かつ `format_on_save = true`（VSCode の `editor.formatOnSave` とは無関係） |
| Pull 時 | `enabled = true` かつ VSCode の設定 `cat5dev.formatter.formatOnPull` が `true` |

- 選択範囲だけの整形（「選択範囲のフォーマット」）には対応していません。常にファイル全体を整形します。
- VSCode の設定にある `cat5dev.formatter.*`（`formatOnPull` 以外）は現在使われていません。設定は `cat5dev.toml` に書いてください。

### コマンドラインで使う
フォーマッタ本体 `vbafmt.exe`（拡張機能フォルダの `bin/` にあります）は単体のコマンドとしても使えます。ファイル（または標準入力）を 1 つ受け取り、整形結果を標準出力に書き出します。`cat5dev.toml` は読まず、オプションは引数で指定します（一覧は `vbafmt.exe -h`）。

```bat
vbafmt.exe modules\Module1.bas_utf > Module1_formatted.bas_utf
vbafmt.exe --normalize-comma-spacing modules\Module1.bas_utf
vbafmt.exe --check modules\Module1.bas_utf   :: 整形が必要なら終了コード 1（CI 向け）
```
