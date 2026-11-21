# C5D_Bridge 契約書（Push / Pull の VBScript 非依存化）

対象: `feature/push-improvement`。プラン `temp/push-bridge-plan-2026-11-21.md` の 2.3・3.1a・3.3〜3.6 を実装するための仕様。
C# 担当（`catia-run.exe`）と TypeScript 担当（拡張機能）は、**この文書だけを見て並行して実装できる**ことを目標にしている。
曖昧な点を見つけたら、実装で決めずにこの文書を直してから進めること。

| 用語 | 意味 |
|---|---|
| 正本（master） | `%LOCALAPPDATA%\Cat5Dev\bridge\C5D_Bridge-<sha8>.catvba`。読み取り専用。**このパスでは読み込まない** |
| リンク（link） | Push / Pull の 1 ステップ（1 回の `catia-run.exe`）ごとに作る、正本へのハードリンク（作れなければ読み取り専用のコピー） |
| ブリッジ | CATIA に読み込まれた VBA プロジェクト `C5D_Bridge`（リンクのパスで自動読み込みされたもの） |
| 自分 | この `catia-run.exe` が作ったリンクのパスで読み込まれたブリッジ |
| 残骸 | 自分以外に読み込まれている `C5D_Bridge`（前回の Close に失敗した、タイムアウトで exe が殺された等） |
| runtime | VBA モジュール `C5D_Bridge_Runtime`（`resources/bridge/C5D_Bridge_Runtime.bas`）。exe に埋め込み、ブリッジに書き込む |

関係するファイル:

| ファイル | 役割 | 担当 |
|---|---|---|
| `resources/bridge/C5D_Bridge.catvba` | 同梱するブリッジ（v2 = `C5D_Bridge_Push` v2 ＋ `C5D_Bridge_Runtime`） | `scripts/build-bridge-catvba.js` で作る |
| `resources/bridge/C5D_Bridge_Push.bas` | ブリッジの起動用モジュール v2 の本文 | 確定（本書 8 章） |
| `resources/bridge/C5D_Bridge_Runtime.bas` | runtime の本文 | 確定（本書 9 章） |
| `tools/catia-run/CatiaRun*.cs` | `catia-run.exe` | C# |
| `scripts/build-bridge.js` | `catia-run.exe` のビルド（runtime の埋め込み） | C# |
| `src/*.ts` | request.json を書き、exe を呼び、result.json を読む | TS |

---

## 1. 全体の流れ

```
VSCode (TS)                             catia-run.exe (C#)                                  CATIA 内（本物の VBA）
───────────                             ──────────────────                                  ─────────────────────
排他ロックを取る
正本を用意（3.2）・古いリンクを掃除（2.3）
request.json を書く ───────────────→   引数・request.json を検証（不正なら 2）
                                        CATIA / SystemService を取得（10 / 11）
                                        リンクを作る（2.2）
                                        ExecuteScript(link, 2, C5D_Bridge_Push, C5D_Version) ─→ 自動読み込み
                                        C5D_Who（自分の特定） / C5D_List（単一の確認・残骸の Close）
                                        runtime の版を確認、違えば C5D_Install で入れる
                                        C5D_RT_ListProjects（0 件なら 14）
                                        要求ごとに ExecuteScript(link, 2, C5D_Bridge_Runtime, C5D_RT_*) ─→ APC / VBE を操作
                                        m<i>.txt（Pull / check）と result.json を書く（tmp → rename）
                                        finally: C5D_Self を guard して Close(False)
result.json を検証（validateCatiaResult、形式は今と同じ）←┘
```

- 1 回の `catia-run.exe` で読み込むブリッジは常に 1 個だけ。Close してから終わる。
- Push は今と同じく `check` → （利用者の確認）→ `push` の 2 ステップ = exe 2 回。読み込みと Close がそれぞれ約 0.5 秒・0.6 秒かかるので、1 ステップあたり約 1.2 秒増える。

---

## 2. ファイル名と配置

### 2.1 正本（TS が用意する）

- パス: `%LOCALAPPDATA%\Cat5Dev\bridge\C5D_Bridge-<sha8>.catvba`
  - `<sha8>` = 同梱の `resources/bridge/C5D_Bridge.catvba` の SHA-256（小文字 16 進）の先頭 8 文字。
- 手順（プラン 3.2 のとおり）:
  1. 同じ名前のファイルがあり、ハッシュが同じなら、そのまま使う（ロックされていてもよい）。
  2. 無ければコピーし、**読み取り専用属性を付け、付いたことを確かめる**。ハッシュも確かめる。
  3. 名前が同じでハッシュが違う場合（壊れている）: ロックされていなければ消して作り直す。ロックされていれば Push を中止する（エラー表示: 一般的なエラー。CATIA を再起動すれば直る）。
- `catia-run.exe` には `--bridge` でこの絶対パスを渡す。

### 2.2 リンク（C# が作る）

- 置き場所: 正本と同じフォルダ。
- 名前: `<正本の拡張子なしの名前>-<CNEXT の pid>-<yyyyMMddHHmmssfff>.catvba`
  - 例: `C5D_Bridge-9d6a56f7-12345-20261121151114123.catvba`
  - `<CNEXT の pid>`: `CNEXT.exe` のプロセスがちょうど 1 個ならその pid、それ以外は `0`（診断用。意味は持たせない）。
  - 時刻はローカル時刻。
  - 同じ名前があれば末尾に `-1`〜`-99` を付ける（それでも重なれば終了コード 13）。**既存のファイルは決して上書きしない**。
- 作り方:
  1. `CreateHardLinkW(link, master, NULL)`（kernel32 の P/Invoke）。
  2. 失敗したら `File.Copy(master, link, false)` → 読み取り専用属性を付ける → 付いたことを確かめる。
  3. どちらも失敗したら終了コード 13。
- ハードリンクは正本と属性を共有するので、正本が読み取り専用ならリンクも読み取り専用になる。**C# は読み込み前に、リンクの読み取り専用属性を確かめる**（無ければ 13。VBA エディタからの保存で正本が書き換わるのを防ぐため）。
- C# はリンクを消さない（Close した後も CATIA が終了するまでロックされるため。掃除は TS の役目）。

### 2.3 古いリンク・コピーの掃除（TS の役目）

- **排他ロック（`SyncOperationManager`）を持っている間に**行う。Push / Pull / Select の各操作で、最初の `catia-run.exe` の前に 1 回。拡張機能の起動時にも行ってよいが、その場合もロックを取ってから。
- 対象: 正本フォルダ直下で、名前が次に合うもの。
  - リンク・コピー: `^C5D_Bridge-[0-9a-f]{8}-\d+-\d{17}(?:-\d{1,2})?\.catvba$`（大文字小文字を区別しない）
  - 古い正本: `^C5D_Bridge-[0-9a-f]{8}\.catvba$` のうち、現在の `<sha8>` でないもの
- ロックの判定: 共有モード 0 で開けるか（`scripts/probe-bridge.js` の `lockProbe` と同じ。`fs.openSync(file, O_RDONLY | 0x10000000)` → 閉じる）。`EBUSY` ならロック中で、消さずに残す（次の機会に消す）。
- 消し方: 読み取り専用を外す（`chmod 0o666`）→ `unlink`。**ハードリンクは正本と属性を共有するので、消した後で正本に読み取り専用属性を付け直し、付いたことを確かめる**（`cleanupProbeFiles` と同じ）。
- 上の名前に合わないファイル（開発用の試験ツールが残したもの等）には触らない。

---

## 3. `catia-run.exe` のコマンドライン

```
catia-run.exe --bridge <正本の絶対パス> --request <dir>\request.json --result <dir>\result.json
catia-run.exe --validate-only --bridge <...> --request <...> --result <...>
```

- 3 つのオプションはどれも必須・1 回だけ・順不同。値は絶対パス（`Path.GetFullPath` しても変わらないこと）。それ以外の引数があれば終了コード 2。
- `--request` と `--result` は**同じフォルダ**にあること。Pull / check の `m<i>.txt` もこのフォルダに書く（= TS の `request.directory`）。
- `--result` のファイルが既にあれば終了コード 2（上書きしない）。
- `--bridge` の名前が `^C5D_Bridge-[0-9a-f]{8}\.catvba$` に合わない、またはファイルが無いときは終了コード 2。
- `--validate-only`: 引数と request.json の検証だけをして、CATIA には接続しない。正しければ 0、不正なら 2。result.json は書かない（TS / C# の CATIA 非接続テスト用）。
- runtime の本文は **exe に埋め込む**（`/resource:resources\bridge\C5D_Bridge_Runtime.bas,C5D_Bridge_Runtime.bas`。マニフェスト リソース名 `C5D_Bridge_Runtime.bas`）。横置きのファイルや `--runtime` オプションは設けない。
  - 理由: exe と runtime の組み合わせがビルド時に固定され、VSIX の中で食い違わない。置き場所を探す処理も要らない。
  - 期待する runtime の版は、埋め込んだ本文の `Private Const RT_VERSION As String = "c5d-runtime:<n>"` から読む（見つからなければ 13）。
- 同時に 2 個動かさない: 名前付き Mutex `Local\Cat5Dev.catia-run` を待ち時間 0 で取る。取れなければ終了コード 13（stderr に `another catia-run.exe is running`）。TS の排他ロックの二重の守り。
- 出力（どちらも UTF-8、BOM なし）:
  - stdout: 処理の段階ごとに 1 行（人が読むログ。例 `load 512ms`、`close ok 604ms`）。TS は解析しない（出力チャネルに出すだけ）。
  - stderr: エラーの詳細。ExecuteScript の例外は今と同じ `0xHRESULT: メッセージ`。
- ビルド: `scripts/build-bridge.js` を変える。コンパイル対象は `tools/catia-run/CatiaRun.cs` と `tools/catia-run/CatiaRun.*.cs`（同じフォルダの `EvalProbe.cs`・`TypeLibDump.cs` は含めない）。JSON に `JavaScriptSerializer` を使う場合は `/r:System.Web.Extensions.dll` を足し、`MaxJsonLength = int.MaxValue` にする。C# 5 の構文に留める。

---

## 4. request.json（TS が書き、C# が読む）

UTF-8（BOM なし。C# は BOM があっても受け付ける）。`JSON.stringify` の出力をそのまま書く。大きさは 64 MiB まで（超えたら 2）。

```json
{
  "version": 2,
  "operationId": "<UUID>",
  "stepId": "<UUID>",
  "kind": "list" | "pull" | "check" | "push",
  "targetProject": "<VBProject.Name>",
  "modules": [ { "name": "Module1", "type": 1, "text": "Option Explicit\r\n..." } ],
  "deletions": [ "OldModule" ]
}
```

| フィールド | 規則 |
|---|---|
| `version` | 数値 `2` |
| `operationId` / `stepId` | 空でない文字列。そのまま result.json に返す |
| `kind` | 4 種のどれか。`pull` と `check` は CATIA 側の処理が同じ（今と同じ） |
| `targetProject` | 文字列。`list` では `""`。それ以外では空でないこと。そのまま result.json に返す |
| `modules` | `push` だけ（他の種類では無いか空配列）。`name`: 空でない、制御文字（U+0000〜U+001F）と `\/:*?"<>\|` を含まない。`type`: 1 / 2 / 3。`text`: 文字列（空でもよい） |
| `deletions` | `push` だけ。名前の規則は `modules[].name` と同じ |

- `modules` と `deletions` を合わせて、名前が大文字小文字を無視して重複していれば 2。
- **`encoding` は無い**（あっても無視する）。本文は JSON の文字列として UTF-16 のまま運ぶ。`cat5dev.toml` の `[project] encoding` は使わなくなる。
- `text` は**今の Push と同じものを送る**: ローカルファイルを UTF-8 で読み、BOM を除き、末尾の空白を削ったもの（`stripBom(...).trimEnd()`）。改行は LF でも CRLF でもそのまま。C# と runtime は一切変えずに `AddFromString` へ渡す。
- TS は今までどおり、新しいユーザーフォーム（CATIA に同じ名前が無い種類 3）を送らない。

---

## 5. result.json（C# が書き、TS が読む。**形式は今と同じ**）

`validateCatiaResult`（`src/catiaTypes.ts`）が受け付ける形をそのまま守る。フィールドの順序は今と同じにする。

```json
{"version":1,"operationId":"..","stepId":"..","kind":"pull","targetProject":"..","expectedCount":2,"succeededCount":2,
 "failures":[],"items":[{"name":"Module1","type":1,"file":"m0.txt"}],"projects":[],"updatedCount":0,"deletedCount":0}
```

- `version` は **1 のまま**（request.json の版とは別）。`operationId`・`stepId`・`kind`・`targetProject` は request.json の値をそのまま返す。
- 書き方: 同じフォルダの `result.json.tmp` に UTF-8（BOM なし）で書き、閉じてから `File.Move(tmp, result)`。JSON の文字列では `"`・`\`・U+0000〜U+001F をエスケープする（`\uXXXX` 可）。それ以外の文字は UTF-8 のままでも `\uXXXX` でもよい。
- 書くのは終了コード 0 と 12 のときだけ（6 章）。それ以外では書かない。
- 終了コード 0 で返すのは、result.json を rename し終えた後。

### 5.1 種類ごとの中身

| kind | 処理 | expectedCount | succeededCount | その他 |
|---|---|---|---|---|
| `list` | `C5D_RT_ListProjects` | 返った名前の数 | 同じ | `projects` = 名前（VBE の順、`C5D_Bridge` を除く、重複はそのまま） |
| `pull` / `check` | `C5D_RT_ListModules` → 種類 1・2・3 の部品ごとに `C5D_RT_ReadModule` | 種類 1・2・3 の部品の数 | 読み出してファイルに書けた数 | `items` = 読めた部品。`file` = `m<i>.txt`（`i` は種類 1・2・3 の部品の中での 0 始まりの順番。失敗した部品も番号を使う） |
| `push` | `modules` を順に `C5D_RT_WriteModule`、続いて `deletions` を順に `C5D_RT_DeleteModule` | `modules.length + deletions.length`（**常に**。途中で止まっても） | `updatedCount + deletedCount` | `updatedCount` = Write の成功数、`deletedCount` = Delete の成功数 |

- 使わないフィールドは空配列・0（`items`・`projects`・`updatedCount`・`deletedCount`）。
- `m<i>.txt` の中身: `C5D_RT_ReadModule` の `ok:` より後の文字列を **UTF-8（BOM なし）** で書く（`new UTF8Encoding(false)`。対になっていないサロゲートは U+FFFD になる）。空のモジュールは 0 バイトのファイル。書けなければ失敗 `export.save`（code = `HResult`、message = 例外のメッセージ）。
  - **TS への影響**: これまで `iconv.decode(..., encoding)` で読んでいた `m<i>.txt` を **UTF-8** で読む（`fs.readFileSync(file, 'utf8')`）。
- 種類 100（ドキュメント モジュール）など 1・2・3 以外の部品は、今と同じく数えず、読まない。

### 5.2 失敗（failures）への変換

runtime の戻り値 `error:<phase>:<number>:<description>`（8.1）を、そのまま 1 件の失敗にする。

| フィールド | 値 |
|---|---|
| `phase` | `<phase>`（例 `push.type`） |
| `target` | `<phase>` が `connect.` か `project.` で始まれば `targetProject`。それ以外: `list` の呼び出しでは `targetProject`（= `""`）、`C5D_RT_ListModules` では `targetProject`、Read / Write / Delete ではそのモジュール名 |
| `code` | `<number>`（32 ビット符号付き整数。負もある） |
| `message` | `<description>`（改行を含むことがある） |

失敗した後の続け方（今の `operation.catvbs` と同じ）:

- `phase` が `connect.` / `project.` で始まる失敗、または `C5D_RT_ListProjects` / `C5D_RT_ListModules` の失敗 → 1 件記録して、**その後の呼び出しをしない**（result.json は書く。終了コード 0）。
- それ以外（モジュール 1 件ごとの失敗）→ 記録して次のモジュールへ進む。
- `ExecuteScript` そのものが例外になった、または戻り値が文字列でない → 失敗を 1 件記録し（`phase` = `list.call` / `export.call` / `push.call`、`target` は上の表のとおり、`code` = 例外の `HResult`（戻り値が文字列でないときは 13）、`message` = 例外のメッセージ（または `Unexpected return value: <型名>`））、**その後の呼び出しをしない**。result.json を書いて**終了コード 12**。

---

## 6. 終了コード

| コード | 意味 | result.json | CATIA で書き込んだか | TS の扱い |
|---|---|---|---|---|
| 0 | 処理を最後まで行った（一部失敗は failures に記録） | ある | 書いた可能性あり | 今と同じ |
| 2 | 引数・request.json が不正、`--result` が既にある | 無い | 何もしていない | **ロックをすぐ解除**。一般的なエラー |
| 10 | CATIA が起動していない（`GetActiveObject` 失敗） | 無い | 何もしていない | **ロックをすぐ解除**。「CATIA を起動してください」（新設） |
| 11 | `SystemService` を取得できない | 無い | 何もしていない | **ロックをすぐ解除**。一般的なエラー |
| 12 | 要求の処理中に `ExecuteScript` が例外になった | 原則ある（書けなかったときだけ無い） | 書いた可能性あり | result があれば今の「一部失敗」。無ければ今と同じ「完了を確認できない」（ロック保留） |
| 13 | ブリッジを準備できない（リンクを作れない、読み取り専用でない、読み込めない、v2 でない、自分を特定できない、残骸を外せない、runtime を入れられない、Mutex を取れない） | 無い | **利用者のプロジェクトには何もしていない** | **ロックをすぐ解除**。「Cat5Dev の仲介プロジェクトを準備できませんでした」（新設） |
| 14 | VBA エディタが未起動（`C5D_RT_ListProjects` が 0 件） | 無い | 何もしていない | **ロックをすぐ解除**。「VBA エディタを一度開いてから再実行してください」（新設） |

- TS の `preDispatch`（`src/catiaExecution.ts`）は `[2, 10, 11, 13, 14, 'ENOENT', 'EACCES', 'EPERM']` にする。
- 13 は「ブリッジにしか書いていない」ことが確かなときだけ使う。利用者のプロジェクトへの呼び出し（`C5D_RT_ReadModule` / `WriteModule` / `DeleteModule`、`ListModules`）を 1 回でも始めた後の異常は 12。
- タイムアウト（TS が exe を止める）は今と同じく「完了を確認できない」。exe が殺されると finally の Close は走らない → 次回の残骸処理（7 章の手順 6）で外れる。
- Close に失敗しても終了コードは変えない（stderr に `warning: bridge close failed: ...`）。

---

## 7. `catia-run.exe` の処理手順

ExecuteScript の呼び方（全呼び出しで共通）:

```csharp
object[] args = new object[] { "a", "b" };          // 必ず要素型 object の配列（string[] だと 0x80020005）
object[] parameters = { libraryPath, 2, moduleName, functionName, args };   // 2 = catScriptLibraryTypeVBAProject
object r = service.GetType().InvokeMember("ExecuteScript", BindingFlags.InvokeMethod, null, service, parameters);
```

- `libraryPath` はリンクのフルパス（残骸を外すときは残骸の DisplayName）。VBA の引数はすべて文字列（モジュールの種類も `"1"` のような文字列）。ByRef の再試行はしない（自動読み込みが絡むので 1 回だけ呼ぶ）。
- 戻り値は VBA の Variant。文字列なら `string`、オブジェクトなら COM プロキシ（`Marshal.IsComObject`）。COM オブジェクトは使い終わったら `Marshal.ReleaseComObject`。

手順:

1. 引数と request.json を検証する（不正なら 2）。`--validate-only` ならここで 0。
2. Mutex を取る（取れなければ 13）。
3. `Marshal.GetActiveObject("CATIA.Application")`（失敗 10）→ `SystemService`（失敗 11）。
4. リンクを作り、読み取り専用を確かめる（2.2。失敗 13）。
5. **読み込み**: `ExecuteScript(link, 2, "C5D_Bridge_Push", "C5D_Version")` → `"c5d-bridge:2"` 以外（例外、v1 の `"c5d-bridge:1"` を含む）なら 13。v1 が返ったときは読み込まれているが `C5D_Self` が無いので Close できない（stderr に `bridge v1 is not supported; restart CATIA` と出す）。
6. **自分の特定**: `C5D_Who()`（8.3）→ 1 件目の Name が `C5D_Bridge`、DisplayName と FileName がリンクのパスと等しい（序数比較、だめなら大文字小文字を無視）こと。違えば 13（このときは Close もしない: 自分を確かめられないので。読み込まれたままのものは次回の手順 7 で残骸として外れる）。
7. **単一の確認**: `C5D_List()`（8.4）。Name が `C5D_Bridge` で自分以外の行（残骸）があれば:
   - DisplayName のフォルダが正本のフォルダと等しく（大文字小文字を無視）、ファイル名が `^C5D_Bridge.*\.catvba$` なら、**残骸自身のパスで** `ExecuteScript(<残骸の DisplayName>, 2, "C5D_Bridge_Push", "C5D_Self")` → 戻ったオブジェクトを guard（Name = `C5D_Bridge`、DisplayName = 残骸の DisplayName）してから `Close(False)`。
   - それ以外（フォルダ外の `C5D_Bridge`、DisplayName が空、`C5D_Self` が無い v1 の残骸、Close の失敗、残骸が 5 個を超える）→ 13。
   - 全部外したら、もう一度 `C5D_List()` で自分だけになったことを確かめる（違えば 13）。
   - **この手順が終わるまで、ブリッジにも書き込まない**（同じ実ファイルを 2 つ読み込んで書き込むと VBA 全体が壊れるため。プラン 2.3）。
8. **runtime**: 手順 6 の部品一覧に `C5D_Bridge_Runtime` があれば `C5D_RT_Version()` を呼び、埋め込みの版と同じなら何もしない。無い、または違えば `C5D_Install("C5D_Bridge_Runtime", <埋め込みの本文>)` → `"ok"` を確かめ、`C5D_RT_Version()` が埋め込みの版を返すことを確かめる。どこかで違えば 13。
9. **VBA エディタの確認**: `C5D_RT_ListProjects()` → `ok` で名前が 0 件なら 14。`error:` なら（kind に関わらず）その失敗を記録し、result.json を書いて 0（5.2）。
10. **要求の処理**（5.1）。`list` は手順 9 の結果をそのまま使う。
11. m<i>.txt と result.json を書く（5 章）。
12. **finally（手順 5 で ExecuteScript を呼んだなら必ず）**: `ExecuteScript(link, 2, "C5D_Bridge_Push", "C5D_Self")` → オブジェクトなら次の 3 つを確かめてから `Close(False)`:
    - `Name` = `C5D_Bridge`（序数比較）
    - `DisplayName` = リンクのパス（大文字小文字を無視）
    - `VBProject.FileName` = リンクのパス（大文字小文字を無視）
    - ひとつでも違えば Close しない（stderr に警告）。手順 6 で自分を特定できていなければ、この finally は行わない。
    - Close の後で同じパスを呼ばない（「マクロ プログラムが不明です」になるだけで、確認にならない）。

呼び出し回数の目安: 読み込み 1 ＋ Who 1 ＋ List 1 ＋ runtime 1〜3 ＋ ListProjects 1 ＋ 要求（list 0 / pull はモジュール数＋1 / push は件数）＋ Close 1。初回の読み込みが約 0.5 秒、Close が約 0.6 秒、それ以外は 1 回数 ms。

---

## 8. 文字列の形式と `C5D_Bridge_Push` v2

### 8.1 共通の形式

VBA の関数はすべて `Variant`（中身は String。`C5D_Self` だけはオブジェクトも返す）を返す。

- **レコード形式**（一覧を返す関数）: 1 行目が `ok`。2 行目から 1 行 1 レコード。行の区切りは LF（`vbLf`）、フィールドの区切りは TAB（`vbTab`）。
  - 各フィールドは次の 4 文字だけを置き換える（VBA 側は `%` を最初に置き換える）: `%` → `%25`、TAB → `%09`、LF → `%0A`、CR → `%0D`。
  - 読む側は `%(09|0A|0D|25)` を**1 回の走査で**戻す（C# なら `Regex.Replace(field, "%(09|0A|0D|25)", m => ((char)Convert.ToInt32(m.Groups[1].Value, 16)).ToString())`）。
  - 一覧が 0 件なら `ok` だけ。
- **本文の形式**（`C5D_RT_ReadModule`）: `ok:` に本文をそのまま続ける（エスケープしない）。
- **エラー**: `error:<種類>:<番号>:<説明>`。`<種類>` は英小文字・数字・`.` で、`:` を含まない。`<番号>` は 10 進の整数（負もある）。`<説明>` は残り全部（`:` や改行を含むことがある）。
  - 正規表現（.NET）: `^error:([a-z0-9.]+):(-?\d+):(.*)$`（`RegexOptions.Singleline`）。
  - 例外: `C5D_Install` の `error:self`（番号と説明が無い。v1 と同じ）。
- 成功・エラーのどちらでもない文字列は、異常として扱う（ブリッジの準備中なら 13、要求の処理中なら 5.2 の `*.call` と同じ扱いで `code` = 13）。

### 8.2 `C5D_Version()` → `"c5d-bridge:2"`

### 8.3 `C5D_Who()`

```
ok
<Name>␉<DisplayName>␉<VBProject.FileName>
<部品の名前>␉<部品の種類>          ← 自分の VBProject の VBComponents ごと（VBE の順）
```

エラー: `error:who.apc:429:..` / `error:who.exec:<n>:..`（`ExecutingProject` が取れない）/ `error:who.project:..` / `error:who.components:..` / `error:who.component:..`。

### 8.4 `C5D_List()`

```
ok
<Name>␉<DisplayName>␉<自分なら 1、それ以外は 0>   ← apc.Projects（IApcProject）ごと
```

- VBA エディタが未起動でも、読み込まれている `C5D_Bridge` は見える（利用者のプロジェクトは見えないことがある。14 の判定には使わない）。
- 読み出しに失敗したフィールドは空文字になる。Name が空の行は利用者のプロジェクトとみなす（残骸の判定に入れない）。DisplayName が空の `C5D_Bridge` 行は残骸として外せない → 13。
- エラー: `error:list.apc:..` / `error:list.projects:..` / `error:list.item:..`。

### 8.5 `C5D_Self()`

- `apc.ExecutingProject`（このコードを実行している `IApcProject`）をオブジェクトとして返す。ExecuteScript が戻った後で C# から `Close(False)` するためのもの。
- エラー: `error:self.apc:429:..` / `error:self.exec:<n>:..`。
- 実行中にしか取れない（ExecuteScript が戻った後の `ExecutingProject` は Empty）ので、必ず呼び出しの戻り値として受け取る。

### 8.6 `C5D_Install(moduleName, source)`

- `ExecutingProject.VBProject` に標準モジュール `moduleName` を作る（無ければ `VBComponents.Add(1)` → 名前を付ける）か、上書きする（`DeleteLines 1, CountOfLines` → `AddFromString source`）。**v1 と違い「名前が C5D_Bridge の先頭のプロジェクト」ではなく、呼び出し元に書く。**
- 戻り値: `"ok"` / `"error:self"`（`moduleName` が `C5D_Bridge_Push`。大文字小文字を無視）/ `error:install.<段階>:<n>:<説明>`。段階: `name`（空の名前）、`apc`、`exec`、`project`（ExecutingProject の Name が `C5D_Bridge` でない）、`components`、`lookup`、`add`、`rename`、`type`（既存の部品が標準モジュールでない）、`codemodule`、`deletelines`、`addfromstring`。
- 書くのはブリッジ自身だけ。利用者のプロジェクトには決して書かない。

---

## 9. `C5D_Bridge_Runtime`

本文: `resources/bridge/C5D_Bridge_Runtime.bas`。版: `c5d-runtime:1`。`src/catiaScripts.ts` が生成していた `operation.catvbs` の処理を移植したもの（探し方・確認・失敗の段階名とメッセージを同じにした）。

共通:

- 状態を持たない（モジュール レベルの変数なし）。呼び出しのたびに `CreateObject("MSAPC.Apc.7.1")`（だめなら `.6.2`、`MSAPC.Apc`）→ `.VBE.VBProjects` を取り直す（今の `ConnectProjects` と同じ。失敗は `connect.apc` / `connect.projects`）。
- 対象プロジェクトの探し方（今の `FindProject` と同じ）: `VBE.VBProjects` を先頭から見て、`Name` が `targetProject` と**大文字小文字を区別して**一致する最初のもの。**ただし Name が `C5D_Bridge` のプロジェクトは一致させない**。見つからなければ `error:project.lookup:5:Target project not found`。`FileName` では探さない。
- モジュールの探し方（今の `FindModule` と同じ）: `VBComponents` を先頭から見て、`Name` が**大文字小文字を無視して**一致する最初のもの。

| 関数 | 戻り値（成功） | 主なエラーの段階 |
|---|---|---|
| `C5D_RT_Version()` | `c5d-runtime:1` | ― |
| `C5D_RT_ListProjects()` | レコード形式。1 レコード = `<VBProject.Name>`（VBE の順、`C5D_Bridge` を除く） | `connect.*`、`list.count`、`list.project`、`list.name` |
| `C5D_RT_ListModules(project)` | レコード形式。1 レコード = `<部品名>␉<種類>`（全種類。100 も含む。絞り込みは C#） | `connect.*`、`project.*`、`export.count`、`export.component`、`export.type`、`export.name` |
| `C5D_RT_ReadModule(project, name)` | `ok:<本文>` | `connect.*`、`project.*`、`export.count` / `export.component` / `export.name`（探す途中）、`export.component:5:Module not found`、`export.codemodule`、`export.linecount`、`export.lines` |
| `C5D_RT_WriteModule(project, name, type, text)` | `ok` | `push.type:5:Invalid module type`（`type` が 1〜3 でない）、`connect.*`、`project.*`、`push.lookup.count` / `.component` / `.name`、`push.add:5:New UserForms require a designer`、`push.add`、`push.name`、`push.type:5:Module type does not match`、`push.codemodule`、`push.linecount`、`push.deletelines`、`push.addfromstring` |
| `C5D_RT_DeleteModule(project, name)` | `ok` | `connect.*`、`project.*`、`push.lookup.*`、`push.delete:5:Module to delete not found`、`push.delete` |

### 9.1 本文の扱い

- **Pull / check**: `CodeModule.Lines(1, CountOfLines)`（`CountOfLines` が 0 なら空文字）。VBE が返すとおりで、行の区切りは CRLF、最後の行の後に CRLF は付かない。**`Attribute` 行は含まれない**（VBE が隠すため。今と同じ）。TS は今と同じく `normalizeModuleText(...) + '\n'` で保存・比較する。
- **Push**: `CountOfLines > 0` なら `DeleteLines 1, CountOfLines`、`Len(text) > 0` なら `AddFromString text`（空の本文なら中身を消すだけ）。本文は変えずに渡す（LF のままでも、今と同じく VBE が行に分ける）。**`Attribute` 行も特別扱いしない**: ローカルのファイルに `Attribute` 行があれば、今と同じくそのまま `AddFromString` に渡る（VBE では構文エラーの行になる）。
- 新しく作ったモジュールに VBE が自動で入れる `Option Explicit` は、`DeleteLines` で消えてから本文が入る（今と同じ）。
- 途中で失敗したときに残るもの（今と同じ・改善しない）: `Add` の後で名前の設定に失敗すると `Module1` などが残る。`DeleteLines` の後で `AddFromString` に失敗すると中身が空のまま残る。
- 大きな本文: 引数・戻り値の BSTR に大きさの上限は設けていない。どこまで通るかは実機で確かめる（プラン 4 の手順 1。目安として 1 MiB の本文の往復を試す）。
- 名前・本文の文字: UTF-16 のまま往復する（日本語・🙂 を含む）。ブリッジの .bas 自体は ASCII だけ。

### 9.2 モジュールの種類（`VBComponent.Type`）

| 値 | 種類 | 拡張子 | Pull | Push |
|---|---|---|---|---|
| 1 | 標準モジュール | `.bas_utf` | する | 作成・上書き |
| 2 | クラス モジュール | `.cls_utf` | する | 作成・上書き |
| 3 | ユーザーフォーム | `.frm_utf` | コード部分だけ | **既存のフォームのコード部分の上書きだけ**。新規作成はしない（デザイン部分を作れないため。今と同じ） |
| 100 | ドキュメント モジュール | ― | しない | しない |

### 9.3 VBA のルール（`tests/bridgeBas.test.js` で静的に検査する）

- ASCII（印字可能文字）だけ、CRLF、末尾も CRLF。`Attribute` 行を書かない（`AddFromString` で入れる本文のため。VBA エディタで取り込むときは `--export-import` で付けたものを使う）。
- 先頭行は `Option Explicit`。モジュール レベルには `Private Const` だけ（変数を持たない）。
- すべての手続きは `Function` で、最初の文が `On Error Resume Next`。
- `MsgBox`・`InputBox`・`Declare`・`GoTo`・`Resume`（`On Error Resume Next` 以外）・`Stop`・`End`・ファイル／シェル操作・`Save`・`Close`・`Run` を使わない。
- `If` / `ElseIf` / `Do While|Until` / `Loop While|Until` / `While` / `Select Case` / `For ... To` の条件でプロパティを読まない（`On Error Resume Next` の下で条件の評価が失敗すると、`If` の中に入ってしまうため）。必ず変数に取ってから比べる。`Err.Number` も同じ。
- `CreateObject` は `MSAPC.Apc.7.1` / `MSAPC.Apc.6.2` / `MSAPC.Apc` だけ。
- 公開関数の名前・引数（すべて `ByVal ... As Variant`）・戻り値（`As Variant`）は本書のとおり。

---

## 10. TS 側で変わること（参考。詳細は TS 担当が決める）

- `src/catiaTypes.ts`: `CatiaRequest` から `encoding` をなくす（または無視）。Push の要求は `{ name, type, text }`（`file` ではなく本文）。`CatiaItem`（結果の `items`）はそのまま。
- `src/catiaExecution.ts`: `operation.catvbs` を書く代わりに `request.json` を書き、`catia-run.exe --bridge <正本> --request <dir>\request.json --result <dir>\result.json` を呼ぶ。`preDispatch` に 2・13・14 を足す。10・13・14 を利用者向けのメッセージに区別できるようにする（例: `CatiaExecutionError` に終了コードを持たせる）。
- `src/catiaSync.ts`: `m<i>.txt` を UTF-8 で読む。Push の本文を `request.modules[].text` に入れる（`iconv` を使わない）。`encoding` を読まない。
- 新しいモジュール（例 `src/bridgeFiles.ts`）: 正本の用意（2.1）と掃除（2.3）。ロックの中で行う。
- `src/i18n.ts`: 10・13・14 のメッセージ（日・英・中）。
- 削除（プラン 3.6）: `buildCatiaScript` と `vbString`（他で使っていなければ）、`scripts/c5d_push.catvbs.bas_utf`、`scripts/update-sync-templates.js` の該当部分、`sync:templates`、`tests/catiaScripts.test.js` の cscript ハーネス（runtime の静的検査は `tests/bridgeBas.test.js` で済んでいる）。
- `.vscodeignore`: `resources/bridge/*.bas` を VSIX から外す（runtime は exe に埋め込むため不要）。`resources/bridge/C5D_Bridge.catvba` は含める（今も含まれている）。
- テスト: `tests/catiaExecution.test.js` / `tests/catiaSync.test.js` の `m0.txt` の文字コード（今は `latin1` / `shift_jis` 前提）と、exe の引数の確認を直す。

---

## 11. ブリッジ `.catvba` の作り直し

### 11.1 自動（`scripts/build-bridge-catvba.js`）

```
node scripts/build-bridge-catvba.js --dry-run       # 計画を JSON で出すだけ（CATIA も exe も使わない、ファイルも作らない）
node scripts/build-bridge-catvba.js                 # 作り直し → 確認 → resources/bridge/C5D_Bridge.catvba を置き換え
node scripts/build-bridge-catvba.js --no-install    # 置き換えない
node scripts/build-bridge-catvba.js --save-method close   # IApcProject.Save が使えないとき、Close(True) で保存する
```

前提（CATIA の状態）:

- CATIA が起動していて、ダイアログが出ていない。
- 元の `.catvba` が v1 のとき（今の同梱品）は、**CATIA を起動し直した直後**（Push もしていない）。v1 の `C5D_Install` は「名前が `C5D_Bridge` の先頭のプロジェクト」に書くので、他の `C5D_Bridge` が読み込まれていないことを、bridge / bridge-build フォルダのファイルのロックで確かめる（ロック中が 1 つでもあれば止まる）。
- VBA エディタは開かなくてよい。利用者のプロジェクトには書き込まない（確認で一覧を読むだけ）。

処理の流れ（各段階は 1 回の ExecuteScript。60 秒で打ち切り。CNEXT.exe の消失・0x800ADF09 などの致命的なエラーで以降を止める）:

1. 元のファイルを `%LOCALAPPDATA%\Cat5Dev\bridge-build\C5D_Bridge-build-<時刻>-w.catvba` に**書き込み可能な**コピーとして置く（正本の bridge フォルダとは別。TS の掃除の対象外）。
2. そのパスで `C5D_Version`（自動読み込み）→ 元の `C5D_Install` で一時モジュール `C5D_Builder` を書く → `C5D_BuildWho` で「ExecutingProject がこのコピー」「`C5D_Bridge` はこれ 1 個」を確かめる。
3. `C5D_Builder.C5D_BuildPut` が `C5D_Bridge_Push` を v2 の本文で置き換え、`C5D_Bridge_Runtime` を書く。`C5D_Version` = `c5d-bridge:2`、`C5D_RT_Version` を確かめる。
4. ExecuteScript が戻った後で、外から（`eval-probe.exe vbacall-obj`）`C5D_Self` の戻り値を guard（Name・DisplayName）してから `VBProject.VBComponents.Remove(VBComponents.Item("C5D_Builder"))`。実行中でないモジュールを外から消すので、自分で自分を消す問題が起きない。v2 に一時的な削除関数を持たせる必要もない。
5. `C5D_Who` で部品がちょうど `C5D_Bridge_Push` と `C5D_Bridge_Runtime` であることを確かめ、`IApcProject.Save(True)`（型情報 `temp/typelib-apc71-20261121-132410.txt`: `Save([opt] ForceSave As BOOL = False)`）→ `Close(False)`。途中で失敗しても、自分を特定できていれば guard 付きで `Close(False)` する（壊れた状態は保存しない）。
6. 保存されたファイルを `temp/bridge-build-<時刻>/C5D_Bridge.catvba` に写し、オフラインで検査する（`scripts/lib/catvba-read.js` で MS-CFB / MS-OVBA を読む）: プロジェクト名、モジュールがちょうど 2 つで本文が `.bas` と同じ、v2、runtime の版、**孤立したストレージ**（どのモジュールにも属さないフォーム等のデータ）。
7. 写したファイルを**新しい**読み取り専用のパス（`...-v.catvba`）で読み込み、`C5D_Version` / `C5D_Who` / `C5D_List` / `C5D_RT_Version` / `C5D_RT_ListProjects` / `C5D_Install`（試験モジュールを書いて呼び、日本語と 🙂 が往復すること）/ `C5D_Install("C5D_Bridge_Push")` = `error:self` を確かめ、guard 付きで `Close(False)`。
8. すべて通れば、`resources/bridge/C5D_Bridge.catvba` を `temp/C5D_Bridge.catvba.bak-<時刻>` に退避して置き換える。孤立したストレージがあれば**置き換えない**（`--allow-orphans` で上書き可。11.4 参照）。

結果の読み方: 標準出力に 1 段階 1 行（`ok` / `FAIL` / `-`（実行しなかった理由））と最後の判定、詳細は `temp/build-bridge-catvba-<日付>-<時刻>.json`（`entries` が各段階、`facts.inspect` がオフライン検査の結果）。

| 終了コード | 意味 | 次にすること |
|---|---|---|
| 0 | 完了（置き換えた、または `--no-install` で確認まで） | `git diff --stat` で `.catvba` の変更を確かめてコミット。sha8 が変わるので、正本の名前も変わる |
| 1 | どこかの段階が失敗 | レポートの最初の `FAIL` を見る。`save` で失敗したら `--save-method close` を試す。だめなら 11.3 |
| 2 | 引数が不正 | ― |
| 3 | 保存したファイルを CATIA が離さず読めない | CATIA を終了してから、表示された `--verify "<...-w.catvba>" --install` を実行する |

`precheck` で止まったら（v1 の元ファイルで、ブリッジのファイルがロック中）: CATIA を起動し直してから再実行する。
`close` が「自分を特定できない」で実行されなかったら: そのコピーは CATIA に残っている。CATIA を（保存せずに）終了する。

### 11.2 確認だけ・検査だけ

```
node scripts/build-bridge-catvba.js --verify <file.catvba> [--install] [--allow-orphans] [--dry-run]
node scripts/build-bridge-catvba.js --inspect <file.catvba>         # CATIA を使わない。0 = v2 で孤立ストレージなし
node scripts/build-bridge-catvba.js --export-import <dir>           # VBA エディタの「ファイルのインポート」用 .bas を書く
```

`--verify` は `<file>` 自体を読み込まない（新しいパスの読み取り専用コピーを読み込む）。

### 11.3 手作業の手順（自動で保存できない場合、または孤立ストレージの無い新しいファイルを作る場合）

1. CATIA を起動し直す。
2. 「ツール」→「マクロ」→「マクロ」→「マクロ ライブラリー」で、ライブラリーの種類を「VBA プロジェクト」にして「新規ライブラリーの作成」。場所は `temp/` の下など、名前は `C5D_Bridge.catvba`（CATIA の版によって表記が違う）。
3. VBA エディタ（Alt+F11）で、そのプロジェクトのプロパティの「プロジェクト名」を `C5D_Bridge` にする。
4. `node scripts/build-bridge-catvba.js --export-import temp/bridge-import` を実行し、VBA エディタの「ファイル」→「ファイルのインポート」で `temp/bridge-import/C5D_Bridge_Push.bas` と `C5D_Bridge_Runtime.bas` を取り込む（モジュール名は `Attribute VB_Name` 行で決まる）。他のモジュール（`Module1` など）があれば削除する。
5. 「デバッグ」→「VBAProject のコンパイル」でエラーが出ないことを確かめ、「上書き保存」。
6. CATIA を終了する（保存したファイルのロックを外すため）。
7. CATIA を起動し、`node scripts/build-bridge-catvba.js --verify <手順 2 のファイル> --install` を実行する。

### 11.4 孤立ストレージ（重要）

今の `resources/bridge/C5D_Bridge.catvba`（v1、5,529,088 バイト）は、モジュールは `C5D_Bridge_Push` 1 つだけだが、ファイルの中に**他のプロジェクトのユーザーフォームのストレージ**（`frmMain`、`frmSelectSeal`、`frmVaridation`、`UserForm1aa` など 25 個前後）と、別のプロジェクトのモジュールの断片が残っている（`node scripts/build-bridge-catvba.js --inspect resources/bridge/C5D_Bridge.catvba` で確認できる）。別のプロジェクトを元に作ったためと思われる。

- このまま VSIX に入れると、そのデータを配布することになる。5.5 MB と大きいのもこのため。
- 自動の作り直し（11.1）は元のファイルを `Save` するので、これらが残る可能性が高い。その場合は置き換えを止める。
- **配布用には 11.3 の手作業（新規ライブラリー）で作ったファイルを使うことを勧める**。

---

## 12. C# 担当と TS 担当の作業範囲

| 担当 | ファイル | 内容 |
|---|---|---|
| C# | `tools/catia-run/CatiaRun.cs`（＋必要なら `CatiaRun.*.cs` に分割） | 3〜7 章のすべて（引数、request.json の検証、リンク、ブリッジの手順、runtime の入れ直し、要求の処理、m<i>.txt と result.json、終了コード、guard 付き Close、`--validate-only`） |
| C# | `scripts/build-bridge.js` | 複数ファイルのコンパイル、runtime の埋め込み（`/resource:`）、必要な参照 |
| C# | `tests/catiaRun.test.js`（新規） | `--validate-only` で引数と request.json の検証（CATIA 非接続）。exe のビルドに csc.exe が要るので、無ければ skip |
| TS | `src/catiaTypes.ts`、`src/catiaExecution.ts`、`src/catiaSync.ts`、新規 `src/bridgeFiles.ts`、`src/extension.ts`、`src/i18n.ts` | 10 章、2.1・2.3、6 章の TS の扱い |
| TS | `src/catiaScripts.ts`、`scripts/c5d_push.catvbs.bas_utf`、`scripts/update-sync-templates.js`、`package.json`（`sync:templates`）、`.vscodeignore` | 3.6 の削除と同梱の調整 |
| TS | `tests/catiaExecution.test.js`、`tests/catiaSync.test.js`、`tests/catiaScripts.test.js`、新規 `tests/bridgeFiles.test.js` | 偽の executor で request.json の中身・終了コードの扱い・UTF-8 の m<i>.txt を確かめる |
| 共有（変更は両者で合意） | `docs/dev/bridge-contract.md`、`resources/bridge/*.bas` | この文書と VBA |

並行して進めるための約束:

- TS は exe を偽の executor で置き換えてテストする（今の `ExecutionOptions.execute` と同じ）。偽の executor は request.json を読み、この文書どおりの result.json と m<i>.txt を書く。
- C# は `--validate-only` と、CATIA に接続した実機スモーク（toQI の `C5D_BridgeTest`）で確かめる。実機には v2 の `.catvba`（11 章）が要る。

---

## 13. 未決事項・リスク

| 項目 | 内容 | 対応 |
|---|---|---|
| 残骸が正本と同じ実ファイル（ハードリンク）のとき | 自分も同じ実ファイル。読み込みと読み出しは実機で問題なかったが、「同じ実ファイルが 2 つ読み込まれた状態で一方を Close する」ことは未検証 | 7 章の手順 7 のとおり、書き込む前に外す。実機で、exe を途中で止めて残骸を作り、次の実行で外れることを確かめる。問題があれば 13 にして CATIA の再起動を求める |
| `IApcProject.Save` の動き | CATIA が読み込んだプロジェクトで `Save(True)` がダイアログなしで保存するかは未検証 | 11.1。だめなら `--save-method close`、それでもだめなら 11.3 |
| 孤立ストレージ | 11.4 | 配布用は 11.3 で作る |
| 大きな本文 | ExecuteScript の引数・戻り値の大きさの上限は未確認 | 実機で 1 MiB を往復させる |
| 呼び出し回数 | Pull はモジュール数＋数回、Push は件数＋数回の ExecuteScript | 1 回数 ms なので当面そのまま |
| 一覧の失敗の細かさ | 今の `ExportModules` は名前の読み出しに失敗した部品だけを失敗にして続けたが、`C5D_RT_ListModules` は一覧全体を失敗にする | 実害は小さいので許容 |
