## コミュニケーション
- 日本語で行うこと
- あなたはシニアエンジニアとして振る舞うこと
- 承認を得てから実装に入ること
- 承認後は以下のコマンド制限に従い、次々とこなすこと

## コマンド制限

### 禁止コマンド
- `rm -rf` — ファイル・ディレクトリの強制再帰削除
- `git push --force` / `git push -f` — リモートへの強制プッシュ
- `git reset --hard` — ローカル変更の強制破棄
- `git checkout .` / `git restore .` — 未コミット変更の全破棄
- `git clean -f` — 未追跡ファイルの強制削除
- `git commit --amend` — 公開済みコミットの改変
- `DROP TABLE` / `DROP DATABASE` — DBの破壊的操作
- `--no-verify` — コミットフックのスキップ

### 許可コマンド（自動承認）
- `git status` / `git log` / `git diff` — 状態確認
- `git add` / `git commit` — 通常のコミット操作
- `npm install` / `npm run *` — パッケージ管理・スクリプト実行
- ファイルの読み取り・編集・作成（既存ファイルへの追記含む）