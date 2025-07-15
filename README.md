# PDF キーワード印刷ツール

## 概要

**PDF キーワード印刷ツール**は、ディレクトリ内の PDF ファイルから指定キーワードを含むページのみを抽出し、プレビューや印刷ができる Windows 専用 GUI アプリケーションです。

### 主な特徴

- ディレクトリ内の全 PDF ファイルを自動検索
- キーワードを含むページのみを抽出
- プレビュー機能（システムの既定アプリで表示）
- 印刷機能（Adobe Acrobat 対応）
- 一括印刷機能
- シンプルな日本語 GUI

---

## 動作環境

- **OS**: Windows 10/11
- **Python**: 3.11 以上
- **推奨**: Adobe Acrobat（印刷品質向上のため）

---

## セットアップ手順

### 1. リポジトリのクローン

```powershell
git clone https://github.com/your-repo/pdf-keyword-printer.git
cd pdf-keyword-printer
```

### 2. 仮想環境の作成（推奨）

```powershell
python -m venv .venv
.venv\Scripts\activate
```

### 3. 依存パッケージのインストール

```powershell
pip install -r requirements.txt
```

### 4. アプリケーションの起動

```powershell
python src/pdf_keyword_printer.py
```

---

## 使用方法

### 基本操作手順

1. **ディレクトリ選択**

   - 「参照...」ボタンをクリック
   - PDF ファイルが含まれるフォルダを選択

2. **キーワード入力**

   - 検索したいキーワードを入力（例：部品出庫）
   - 大文字小文字は区別されません

3. **プリンタ選択**（オプション）

   - ドロップダウンから印刷先プリンタを選択
   - 省略時はシステムの既定プリンタを使用

4. **PDF 検索実行**

   - 「検索」ボタンをクリック
   - ディレクトリ内の全 PDF ファイルを検索

5. **結果確認と操作**

   - ヒットした PDF ファイルが一覧表示
   - 各ファイルのヒットページ数が表示
   - 「プレビュー」ボタンで該当ページを確認
   - 「印刷」ボタンで該当ページのみ印刷

6. **一括印刷**
   - 「一括印刷」ボタンで全ヒット PDF を連続印刷

---

## Windows アプリ（EXE）化

### 1. PyInstaller のインストール

```powershell
pip install pyinstaller
```

### 2. EXE ファイルの作成

```powershell
pyinstaller "PDFキーワード印刷ツール.spec"
```

### 3. 配布ファイル

- `dist/PDFキーワード印刷ツール.exe` - 実行ファイル

---

## 使用ライブラリ

- **pypdf**: PDF 読み書き・ページ抽出
- **pdfplumber**: 高精度テキスト抽出
- **pillow**: 画像処理
- **pywin32**: Windows 印刷機能・プリンタ管理
- **tkinter**: GUI フレームワーク

---

## ライセンス

MIT License

---

## 作者

- **開発**: [CrowsNET Co., Ltd.]
- **対応 OS**: Windows 10/11
