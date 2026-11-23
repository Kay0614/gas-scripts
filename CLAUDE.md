# CLAUDE.md

## プロジェクト概要
Google Apps Script（GAS）のスクリプト群を管理するプロジェクト。

## コミュニケーション
- 返答は日本語で、簡潔に。

## 開発方針
- 言語は Google Apps Script（JavaScript / V8 ランタイム）。
- ローカル開発と GAS への反映には `clasp` を使う想定（導入時は手順をここに追記）。
- 関数・変数名は分かりやすい英語の camelCase。各関数には用途を示す JSDoc コメントを付ける。
- APIキーやIDなどの機密情報はコードに直書きせず、スクリプトプロパティ（`PropertiesService`）で管理する。
- GAS の実行時間制限（6分）やクォータを意識し、シート操作は `getValues` / `setValues` でまとめて行う。

## ディレクトリ構成（予定）
- `src/` … GAS のソースコード（`.js` / `.gs`）
- `src/appsscript.json` … GAS マニフェスト
- `.clasp.json` … clasp 設定（スクリプトID）

## Git運用ルール
- **コードを変更するたびに、コミットして GitHub にプッシュする。**
  - 変更が一段落したら `git add` → `git commit` → `git push` まで行う。
  - コミットメッセージは日本語で、変更内容が分かるように簡潔に書く（例：`売上集計関数を追加`）。
- 1コミット1目的を基本とし、無関係な変更を混ぜない。
- 機密情報を含むファイル（`.clasprc.json`、認証情報、`.env` など）はコミットしない。`.gitignore` に追加する。
- プッシュ前に、構文エラーや動作に問題がないことを確認する。

## 確認方法
- `clasp push` で GAS に反映し、スクリプトエディタで対象関数を実行してログ（`console.log`）を確認する。
