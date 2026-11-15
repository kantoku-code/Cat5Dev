# tlb-extractor

CATIA V5 の型ライブラリ（`.tlb`）から、VBA Language Server（`tools/vba-ls`）が読む `catia_api.db`（SQLite、スキーマは [schema.sql](schema.sql)）を作るツールです。Windows・CATIA V5 のインストール・[uv](https://docs.astral.sh/uv/) が必要です。

## 同梱 DB の再生成

```sh
cd tools/tlb-extractor
uv run python extract.py --out catia_api.db          # CATIA の標準インストール先から自動検索
uv run python extract.py --verify catia_api.db       # 不変条件の検査
cp catia_api.db ../../bin/catia_api.db               # bin/ にも同じファイルを置く
```

同じ TLB からは同じバイト列の DB ができます（`meta` テーブルに CATIA のリリース・TLB のダイジェストを記録）。同梱 DB には説明文を入れていません。

## 説明文付きの DB を手元で作る（任意）

CATIA に付属するヘルプ `V5Automation.chm` の説明文を DB に入れると、補完・ホバー・引数のヒントに説明が表示されます。Dassault Systèmes の文書を含むため、**生成した DB や CHM の展開物はリポジトリにコミットせず、各自の PC の中だけで使ってください**。

```sh
uv run python extract.py --chm --chm-mode full --chm-workdir D:/ --out D:/catia/catia_api_doc.db
```

- `--chm` … TLB に登録された `V5Automation.chm` を使います（`--chm <パス>` で .chm または展開済みのディレクトリを指定可）。
- `--chm-mode` … `summary`（1 行要約、既定）／`full`（詳細と引数・Enum 値の説明）。
- `--chm-workdir` … CHM を一時的に展開する場所（約 50MB、終了時に削除）。`hh.exe` はパスに空白があると失敗するので、空白の無いディレクトリを指定します。

作った DB は環境変数 `CAT5DEV_TLB_DB` に絶対パスを設定すると、同梱 DB の代わりに読み込まれます（VS Code を再起動して反映）。

```bat
setx CAT5DEV_TLB_DB D:\catia\catia_api_doc.db
```

## テスト

```sh
uv run python -m unittest discover -s .
```

CATIA がインストールされていない環境では、実 TLB を使うテストは skip されます。
