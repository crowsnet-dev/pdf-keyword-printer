# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクト概要

Windows 専用の GUI アプリ。指定ディレクトリ内の PDF からキーワードを含むページを抽出し、プレビュー・印刷する。全コードは `src/pdf_keyword_printer.py`（単一ファイル、約 1,270 行）に集約されている。

## コマンド

```powershell
# 依存インストール
pip install -r requirements.txt

# 実行（GUI 起動）
python src/pdf_keyword_printer.py

# EXE 化（成果物: dist/PDFキーワード印刷ツール.exe）
pip install pyinstaller
python -m PyInstaller pdf_keyword_printer.spec
```

- 自動テストは現時点で存在しない。テストを追加する場合は `.cursor/rules/test-strategy.mdc` の観点表・Given/When/Then コメント・境界値網羅ルールに従うこと。
- `.gitignore` は `*.spec` を無視するが、ビルド定義の `pdf_keyword_printer.spec` は意図的にリポジトリへ含まれている。削除・再生成しないこと。

## アーキテクチャ

単一モジュール内で責務ごとにクラス分離している。変更時はこの境界を維持する。

- `Constants` — ウィンドウサイズ、リトライ回数、印刷ジョブ待機秒数などのチューニング定数を集約。印刷信頼性を優先した固定 sleep 値（例: `PRINT_JOB_WAIT = 3.0`）はコメントで根拠が明示されており、安易に削減しないこと。
- `PdfProcessor` — `pdfplumber` でテキスト抽出 → ヒットページ検索、`pypdf` の `PdfReader`/`PdfWriter` で該当ページを新規 PDF としてテンポラリに書き出し（テキストレイヤー保持が要件）。ファイル名は `YYYYMMDDhhmmss_<元名>_<キーワード>抽出版.pdf` で、Windows パス長 260 文字を超えないよう短縮処理あり。
- `PrinterManager` — `win32print` でプリンタ列挙・印刷キュー監視。`print_pdf` は **Adobe Acrobat の CLI (`Acrobat.exe /t`)** を優先的に起動し、見つからない場合のみ `os.startfile(path, "print")` にフォールバック。Acrobat のパスは `print_pdf` 内にハードコードされた候補リストから存在確認して選ぶ。
- `PreviewManager` — プレビューは単に OS 既定アプリで開くだけ（`os.startfile(path)`）。
- `PdfKeywordPrinter(tk.Tk)` — Tk ルート。長時間処理（検索・抽出・印刷）は `threading.Thread(daemon=True)` に逃し、UI 更新は必ず `self.after(0, ...)` 経由でメインスレッドへ戻す。

## 押さえるべき振る舞い

- **キーワード検索**: `keyword_var` はカンマ区切り文字列。`find_keyword_pages` は OR 一致・大文字小文字無視。既定値 `(000101),(000106),外注ケーブル` は特定帳票のコードを想定した初期値で、UI 上で上書き可能。
- **印刷の成否確認**: Adobe 起動直後に固定 3 秒待機 → `check_print_job_sent`（キュー投入を最大 10 秒ポーリング）→ `wait_for_print_job_completion`（最大 60 秒ポーリング）の順で信頼性を担保。タイムアウトしてもエラーにはせず、プロセス終了のみ実施。
- **一時ファイル管理**: 生成した抽出 PDF は `self.temp_pdf_paths` に積み、`atexit` / `signal` (`SIGINT`/`SIGTERM`) / `WM_DELETE_WINDOW` / `finally` の 4 経路で `_cleanup_on_exit` を呼んで削除する。経路を追加・削除する際はこの冗長性を崩さないこと。
- **Windows 依存**: `win32print` / `win32api` / `win32con` の import は `try/except ImportError` で囲んであり、失敗時は `PRINTER_SUPPORT = False` でプリンタ UI を無効化する。pywin32 前提を緩める際はこのフラグ経路を維持する。
- **ログ出力**: 処理経過は `print()` で標準出力に出す一方、ユーザー向けメッセージは `add_message()` で GUI のメッセージエリアへ書く。双方に出したい情報と、内部デバッグ用情報を混同しないこと。

## コミット・PR 規約

`.cursor/rules/` 配下のルールを常に適用する。要点のみ:

- コミットメッセージとPR本文は **日本語**（`language = "ja"`）。`<Prefix>: サマリ` + 箇条書き本文。Prefix は Conventional Commits 準拠（`feat` / `fix` / `refactor` / `docs` / `chore` …）。
- 差分を確認したうえでメッセージを生成する。ブランチ名や Issue タイトルからの推測で書かない。
- テストを伴う変更では観点表（正常系・異常系・境界値 `0/最小/最大/±1/空/NULL`）を先に提示してから実装する。
