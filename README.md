# 工事書式転記システム

見積決定申請書の内容を出来高・工事完了書式へ自動転記するWebアプリです。

## ファイル構成

```
transfer_app/
├── app.py              # メインアプリ
├── transfer.py         # 転記ロジック
├── requirements.txt    # 必要ライブラリ
└── templates/          # Excelテンプレート
    ├── 出来高.xlsx
    └── 工事完了.xlsx
```

## 公開手順（無料・サーバー不要）

### 1. GitHubにアップロード

1. https://github.com にアクセスしてアカウント作成（無料）
2. 「New repository」でリポジトリを作成（名前は例：`koji-tenki`）
3. このフォルダの全ファイルをアップロード
   - `app.py`
   - `transfer.py`
   - `requirements.txt`
   - `templates/出来高.xlsx`
   - `templates/工事完了.xlsx`

### 2. Streamlit Cloudで公開

1. https://share.streamlit.io にアクセス（GitHubアカウントでログイン）
2. 「New app」をクリック
3. 先ほど作ったリポジトリを選択
4. Main file path: `app.py` を指定
5. 「Deploy!」をクリック → 数分でURLが発行される

### 3. 社内で共有

発行されたURL（例：`https://koji-tenki.streamlit.app`）を
社内のチャットやメールで共有するだけで全員が使えます。

## 使い方

1. 見積決定.xlsx をアップロード
2. 書式を選択（出来高 or 工事完了）
3. 今回が何回目かを入力
4. 工事完了（最終）の場合は過去の出来高ファイルもアップロード
5. 「転記してExcelをダウンロード」ボタンを押す
6. ダウンロードしたExcelを必要に応じて手修正

## 書式追加の方法

新しい書式を追加したい場合は `transfer.py` に関数を追加し、
`app.py` の書式選択ラジオボタンに追加してください。
