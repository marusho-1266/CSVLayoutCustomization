# CSVレイアウト変更ツール仕様書

## 1. 概要

本ツールはCSVファイルのレイアウト（項目の構成や内容）を柔軟に変更するためのデスクトップアプリケーションである。Windows 11環境で動作し、CSVファイルの項目の並べ替え、結合、文字の除去・追加などの操作を行うことができる。

## 2. 動作環境

- OS: Windows 11
- 実行形式: スタンドアロンのデスクトップアプリケーション
- 必要ライブラリ（開発時）: 
  - Python 3.x
  - pandas
  - tkinter / tkinterdnd2

## 3. 機能仕様

### 3.1 基本機能

| 機能ID | 機能名 | 説明 |
|--------|--------|------|
| F01 | CSVファイル読み込み | UTF-8またはShift-JIS形式のCSVファイルを読み込む |
| F02 | 項目の並べ替え | CSVファイルの列（項目）の順序を変更する |
| F03 | 複数項目の結合 | 複数の列を結合して新しい列を作成する |
| F04 | 文字の除去 | 特定の列から指定された文字を削除する |
| F05 | 文字の追加 | 特定の列の値の前後に文字を追加する |
| F06 | ドラッグ&ドロップ | ファイルをアプリケーション上にドラッグして読み込む |
| F07 | プレビュー表示 | 変更後のCSVファイルのプレビューを表示する |
| F08 | 変換結果の保存 | 変更後のCSVファイルを保存する |
| F09 | プロファイル管理 | 変換設定をプロファイルとして保存・読み込み |
| F10 | 文字コード選択 | 入力/出力の文字コードをUTF-8/Shift-JISから選択 |
| F11 | 都道府県名除去 | 指定した住所項目から都道府県名を削除する |
| F12 | 都道府県コード取得 | 指定した項目に含まれる都道府県名から、対応する都道府県コードを新しい列として取得する |
| F13 | 文字列抽出 | 指定した項目から、指定した開始位置と文字数で文字列を抽出し、新しい列として追加する |
| F14 | ヘッダー行除去 | 出力CSVファイルからヘッダー行（1行目）を除去する |
| F15 | 文字置換 | 特定の列の文字列を別の文字列に置換する |

### 3.2 詳細仕様

#### 3.2.1 CSVファイル読み込み (F01)

- UTF-8とShift-JIS（SJIS）の両方のエンコーディングに対応
- 選択したエンコーディングで読み込めない場合、自動的に代替エンコーディングを試行
- ファイル選択ダイアログまたはドラッグ&ドロップで読み込み可能

#### 3.2.2 項目の並べ替え (F02)

- カンマ区切りで項目名を指定し、指定した順序に列を並べ替え
- 指定されなかった列は出力結果から除外される
- カンマを連続して指定した場合 (`,,`)、その位置にヘッダーおよび値が空の列を挿入する
- 指定されなかった列は元の順序を保持して末尾に配置

#### 3.2.3 複数項目の結合 (F03)

- 指定した複数の列を結合して新しい列を作成
- 結合時の区切り文字を指定可能
- 書式：`新項目名:結合元項目1,結合元項目2...,区切り文字`
  - **注意:** 最後のカンマ (`,`) 以降の文字列全体が区切り文字として扱われます。区切り文字が不要な場合は、最後にカンマのみを記述します (例: `新項目:元1,元2,`)。

#### 3.2.4 文字の除去 (F04)

- 特定の列から指定された文字を削除
- 複数の文字を指定可能
- 書式：`項目名:除去する文字1,除去する文字2...`

#### 3.2.5 文字の追加 (F05)

- 特定の列の値の前または後に文字を追加
- 書式：`項目名:位置:追加文字`
- 位置は「前」または「後」を指定

#### 3.2.6 ドラッグ&ドロップ (F06)

- CSVファイルをアプリケーション上の指定領域にドラッグして読み込み可能
- ドロップ領域は視覚的に区別される

#### 3.2.7 プレビュー表示 (F07)

- 変更後のデータを表形式でプレビュー表示
- 最大10行を表示し、それ以上のデータがある場合はその旨を表示

#### 3.2.8 変換結果の保存 (F08)

- 変換後のCSVファイルを任意の場所に保存
- 保存時のエンコーディングを選択可能（UTF-8またはShift-JIS）
- 出力時、すべてのフィールドはダブルクォートで囲まれる。

#### 3.2.9 プロファイル管理 (F09)

- 変換設定をプロファイルとして保存可能
- 保存したプロファイルを読み込み・編集・削除可能
- プロファイルはJSON形式で保存され、異なる端末間で共有可能

#### 3.2.10 文字コード選択 (F10)

- 入力ファイルの文字コードを選択可能（UTF-8またはShift-JIS）
- 出力ファイルの文字コードを選択可能（UTF-8またはShift-JIS）
- 選択したエンコーディングでファイルを開けない場合、自動的に代替エンコーディングを試行

#### 3.2.11 都道府県名除去 (F11)

- 「文字除去」タブ内の専用チェックボックスで機能の有効/無効を切り替え
- 対象となる住所項目名を指定する入力欄を設ける
- 対象項目名はカンマ区切りで複数指定可能
- 有効化され、対象項目名が指定されている場合、該当する列の値の先頭が日本の47都道府県名（北海道、青森県...沖縄県）と一致する場合に、その都道府県名部分を削除する
- 指定された対象項目名が存在しない場合は、処理をスキップする（または警告を表示する）

#### 3.2.12 都道府県コード取得 (F12)

- 「都道府県コード」タブ内の専用チェックボックスで機能の有効/無効を切り替え
- 都道府県名が含まれる既存の「対象項目名」を指定する入力欄を設ける
- 生成される都道府県コードを格納する「新しい項目名」を指定する入力欄を設ける（デフォルト値: "都道府県コード"）
- 有効化され、対象項目名と新しい項目名が指定されている場合、対象項目の値の先頭が日本の47都道府県名と一致する場合に、対応するJIS X 0401準拠の2桁の都道府県コードを新しい項目名の列に追加する
- 都道府県名が見つからない場合、新しい列には空文字列を設定する
- 指定された対象項目名が存在しない場合、または新しい項目名が既存の項目名と重複する場合は、処理をスキップし警告を表示する

#### 3.2.13 文字列抽出 (F13)

- 「文字列抽出」タブ内のテキストエリアに設定を記述する
- 書式：`新項目名:抽出元項目:開始位置:文字数`
  - 各設定は改行で区切る
- **新項目名**: 抽出結果を格納する新しい列の名前を指定する
- **抽出元項目**: 文字列を抽出する既存の列の名前を指定する
- **開始位置**: 抽出を開始する文字の位置を1以上の整数で指定する (1始まり)
- **文字数**: 抽出する文字数を0以上の整数で指定する
- 指定された抽出元項目が存在しない場合、処理をスキップし警告を表示する
- 新項目名が既に存在する場合、処理を実行し既存の列を上書きする旨の警告を表示する
- 開始位置、文字数が整数でない場合、または開始位置が1未満、文字数が0未満の場合、処理をスキップし警告を表示する
- 抽出元項目の値が欠損値(NaN)の場合、または開始位置が文字列長を超える場合、新しい列には空文字列を設定する
- 抽出元項目の値が数値等の場合、文字列に変換してから抽出処理を行う

#### 3.2.14 ヘッダー行除去 (F14)

- 「プレビューと実行」エリア内の専用チェックボックスで機能の有効/無効を切り替え
- チェックボックスがオンの場合、`変換結果の保存 (F08)` 時に出力されるCSVファイルからヘッダー行（項目名の行）を除去する
- チェックボックスがオフの場合、ヘッダー行は通常通り出力される

#### 3.2.15 文字置換 (F15)

- 「文字置換」タブ内のテキストエリアに設定を記述する
- 書式：`項目名:置換前文字列:置換後文字列`
  - 各設定は改行で区切る
- **項目名**: 置換対象の列の名前を指定する
- **置換前文字列**: 置換したい文字列を指定する
- **置換後文字列**: 置換後の文字列を指定する (空文字列も可)
- 指定された項目名が存在しない場合、処理をスキップし警告を表示する
- 置換前文字列が空の場合、処理をスキップし警告を表示する
- 対象列の値を文字列に変換してから置換処理を行う
- 完全一致での置換を行う (正規表現は使用しない)

## 4. ユーザーインターフェース

### 4.1 画面構成

アプリケーションは以下の主要な領域で構成される：

1. **左側パネル**：プロファイル管理と設定エリア
   - プロファイル選択・操作ボタン
   - 設定タブ（並べ替え、結合、文字列抽出、文字除去、文字追加）
     - 文字除去タブ内に、通常の文字除去設定に加え、都道府県名除去用のチェックボックスと対象項目名入力欄を含む
     - 都道府県コードタブ内に、機能有効化チェックボックス、対象項目名入力欄、新しい項目名入力欄を含む

1. **右側パネル**：ファイル操作とプレビューエリア
   - 文字コード選択
   - ドラッグ&ドロップエリア
   - ファイル選択ボタン
   - プレビュー表示
   - ヘッダー行除去チェックボックス 
   - 変換・保存ボタン

### 4.2 操作フロー

1. プロファイルの選択または新規作成
2. 変換設定の入力（並べ替え、結合、文字除去、文字追加）
3. 文字コードの選択
4. CSVファイルの読み込み（ドラッグ&ドロップまたはファイル選択）
5. プレビューの確認
6. 変換して保存ボタンのクリック
7. 保存先の選択
8. ファイルの保存

## 5. データ構造

### 5.1 プロファイルデータ

プロファイル情報はJSON形式で保存され、以下の構造を持つ：

```json
{
  "プロファイル名1": {
    "reorder": "項目A,,項目B,項目C...",
    "merge": "新項目名:項目1,項目2,区切り文字\n新項目名2:項目A,項目B,\n...",
    "extract": "新項目名1:元項目A:1:5\n新項目名2:元項目B:5:4",
    "remove": "項目名:除去文字1,除去文字2\n...",
    "add": "項目名:位置:追加文字\n...",
    "replace": "項目名1:置換前1:置換後1\n項目名2:置換前2:置換後2",
    "remove_prefecture": {
      "enabled": true,
      "column": "住所1,所在地"
    },
    "get_pref_code": {
      "enabled": true,
      "source_column": "住所1",
      "new_column": "都道府県コード"
    },
    "remove_header": false
  },
  "プロファイル名2": {
    ...
  }
}
```

### 5.2 ファイル形式

- 入力：CSVファイル（UTF-8またはShift-JIS）
- 出力：CSVファイル（UTF-8またはShift-JIS）

## 6. 制約事項と注意点

- 対応OSはWindows 11のみ
- カラム名に特殊文字（コロン、カンマなど）が含まれる場合、操作が正しく機能しない可能性がある
- 大容量CSVファイル（数百MB以上）の処理は遅延する可能性がある
- プロファイル設定ファイル（csv_profiles.json）はアプリケーションと同じディレクトリに保存される

## 7. エラー処理

| エラーID | エラー内容 | 対応 |
|----------|------------|------|
| E01 | ファイル読み込み失敗 | エラーメッセージを表示 |
| E02 | エンコーディングエラー | 代替エンコーディングを試行し、失敗した場合はエラーメッセージを表示 |
| E03 | データ処理エラー | 致命的なエラーの場合はエラーメッセージを表示し、元のデータを維持。設定のパースエラーなど、処理可能なエラーの場合はコンソールに警告を表示し、該当設定行の処理をスキップして続行する。 |
| E04 | プロファイル保存/読み込みエラー | エラーメッセージを表示 |
| E05 | 文字列抽出/結合/除去/追加エラー | 設定形式不正、項目名不正、数値不正等の場合にコンソールに警告を表示し、該当設定行の処理をスキップする。 |
| E06 | 文字置換エラー | 設定形式不正、項目名不正等の場合にコンソールに警告を表示し、該当設定行の処理をスキップする。 |

## 8. 将来の拡張性

以下の機能拡張が検討可能である：

- 正規表現を用いた高度な文字列処理
- CSVフォーマット以外のファイル形式（Excel、TXT等）への対応
- バッチ処理による複数ファイルの一括変換
- 履歴機能による操作の取り消し・やり直し
- 高度なフィルタリング機能

## 9. ビルドと配布

### 9.1 開発環境

- Python 3.x
- 必要ライブラリ：pandas, tkinterdnd2

### 9.2 実行ファイル作成

PyInstallerを使用して単一の実行ファイルを作成：

```
pyinstaller --onefile --windowed --clean CSVLayoutCustomization.py
```

### 9.3 配布方法

- スタンドアロンの実行ファイル（.exe）として配布
- プロファイル設定ファイル（csv_profiles.json）を含めることで設定の共有が可能

## 10. パフォーマンス要件

- 一般的なCSVファイル（〜10MB）を5秒以内に処理
- メモリ使用量：最大500MB程度
- アプリケーション起動時間：3秒以内
