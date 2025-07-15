# pdf-keyword-printer

## 概要

**pdf-keyword-printer**は、PDF ファイル内の指定キーワードを含むページだけを抽出し、プレビューや印刷ができる Windows 用 GUI アプリケーションです。  
Python 製で、Windows 標準の印刷機能を利用します。外部 PDF リーダーは不要です。

---

## 特徴

- PDF ファイルからキーワードでページ抽出
- プレビュー（Microsoft Edge で表示）
- 選択したプリンタで直接印刷
- シンプルな日本語 GUI

---

## 動作環境

- Windows 10/11
- Python 3.11 以上

---

## セットアップ手順

### 1. 仮想環境の作成（推奨）

コマンドプロンプトまたは PowerShell でプロジェクトルートにて：

```powershell
python -m venv .venv
.venv\Scripts\activate
```

### 2. 依存パッケージのインストール

```powershell
pip install -r requirements.txt
```

---

## 使い方

### 1. アプリの起動

```powershell
python src/pdf_keyword_printer.py
```

### 2. 操作手順

1. 「参照...」ボタンで PDF ファイルを選択
2. 検索キーワードを入力（例：部品出庫）
3. プリンタを選択（省略時は既定プリンタ）
4. 「プレビュー」ボタンで該当ページを Edge で確認
5. 「印刷」ボタンで該当ページのみ印刷

---

## Windows アプリ（EXE）化手順

### 1. PyInstaller のインストール

```powershell
pip install pyinstaller
```

### 2. EXE ファイルの作成

```powershell
pyinstaller build/pdf_keyword_printer.spec
```

- `dist/`フォルダ内に`pdf_keyword_printer.exe`が生成されます。

---

## 注意事項

- 依存パッケージは`requirements.txt`で管理しています。
- `.venv/`などの仮想環境フォルダは git 管理対象外です。
- 印刷機能は Windows 標準機能を利用します。プリンタドライバが正しくインストールされている必要があります。

---

## ライセンス

MIT License

---

## 作者

- [CrowsNET Co., Ltd.]
