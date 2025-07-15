"""
Keyword-based PDF page extractor and printer (Windows GUI)

Dependencies:
    pip install pypdf pdfplumber pillow
    pip install pywin32   # optional, for printer listing

Usage:
    Run this script with Python 3.11+ on Windows.
    The GUI lets users:
        1. Browse and choose a directory containing PDF files.
        2. Enter keyword (case‑insensitive).
        3. Optionally choose printer (defaults to system default).
        4. View list of PDF files with preview and print buttons for each.
        5. Click "一括印刷" to print all PDFs containing the keyword.
        
Note: Printing uses Windows standard functions only. No external PDF readers required.
"""

import os
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional, Tuple, Dict
import time
import hashlib
import shutil
import subprocess
import glob
import datetime

import pdfplumber
from pypdf import PdfReader, PdfWriter

# Printer handling (optional)
try:
    import win32print
    PRINTER_SUPPORT = True
except ImportError:
    PRINTER_SUPPORT = False


class Constants:
    """定数クラス"""
    WINDOW_WIDTH = 800
    WINDOW_HEIGHT = 600
    PADDING = 6
    BUTTON_WIDTH = 12
    ENTRY_WIDTH = 48
    KEYWORD_ENTRY_WIDTH = 30
    PRINTER_MENU_WIDTH = 40
    
    # タイムアウト設定
    PRINT_TIMEOUT = 30
    FILE_ACCESS_WAIT = 5
    
    # UltraThinkプリンタ対応の追加設定
    PRINT_RETRY_COUNT = 3
    PRINT_RETRY_DELAY = 2
    FILE_STABILITY_WAIT = 3
    
    # 出力フォルダ設定
    OUTPUT_FOLDER_NAME = "output"


class PdfProcessor:
    """PDF処理を担当するクラス"""
    
    @staticmethod
    def find_keyword_pages(pdf_path: str, keyword: str) -> List[int]:
        """キーワードを含むページを検索"""
        hit_pages = []
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                if keyword.lower() in text.lower():
                    hit_pages.append(i)
        return hit_pages
    
    @staticmethod
    def extract_pages(pdf_path: str, page_indices: List[int]) -> str:
        """指定されたページを抽出して一時ファイルに保存"""
        if not page_indices:
            raise ValueError("ページが指定されていません")
        
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        
        for idx in page_indices:
            if idx < len(reader.pages):
                writer.add_page(reader.pages[idx])
        
        # 一時ファイル名を生成
        timestamp = str(int(time.time()))
        file_hash = hashlib.md5(pdf_path.encode()).hexdigest()[:8]
        filename = f"extracted_{timestamp}_{file_hash}.pdf"
        
        # 一時ファイルを作成
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, filename)
        
        with open(temp_path, "wb") as f:
            writer.write(f)
            f.flush()
            os.fsync(f.fileno())
        
        # ファイルサイズを確認
        if os.path.getsize(temp_path) == 0:
            raise RuntimeError(f"一時PDFファイルの作成に失敗しました（0バイト）")
        
        return temp_path
    
    @staticmethod
    def find_pdf_files(directory: str) -> List[str]:
        """指定ディレクトリ内のPDFファイルを検索（重複排除）"""
        try:
            pattern = os.path.join(directory, "**", "*.pdf")
            files = set(glob.glob(pattern, recursive=True))
            files.update(glob.glob(pattern.replace("*.pdf", "*.PDF"), recursive=True))
            return sorted(files)
        except Exception:
            return []


class PrinterManager:
    """印刷処理を担当するクラス"""
    
    @staticmethod
    def get_available_printers() -> List[str]:
        """利用可能なプリンタのリストを取得"""
        if not PRINTER_SUPPORT:
            return []
        
        try:
            printers = [p[2] for p in win32print.EnumPrinters(
                win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            )]
            return printers
        except Exception:
            return []
    
    @staticmethod
    def get_default_printer() -> str:
        """既定プリンタを取得"""
        if not PRINTER_SUPPORT:
            return "(既定)"
        
        try:
            return win32print.GetDefaultPrinter()
        except Exception:
            return "(既定)"
    
    @staticmethod
    def print_pdf(pdf_path: str, printer: str) -> None:
        """PDFを印刷（Acrobat優先）"""
        if not os.path.exists(pdf_path):
            raise RuntimeError(f"PDFファイルが見つかりません: {pdf_path}")

        # Acrobat パスを優先順位で設定
        acro_paths = [
            r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            r"C:\Program Files (x86)\Adobe\Reader\AcroRd32.exe",
        ]
        
        # Acrobat で印刷を試行
        for exe in acro_paths:
            if os.path.exists(exe):
                try:
                    subprocess.run(
                        [exe, "/t", pdf_path, printer],
                        check=True,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                    )
                    time.sleep(Constants.FILE_ACCESS_WAIT)
                    return
                except Exception as e:
                    continue   # 次の候補へ

        # Acrobat が見つからない場合のフォールバック
        try:
            # 既定アプリで print verb を試行
            os.startfile(pdf_path, "print")
            time.sleep(Constants.FILE_ACCESS_WAIT)
            return
        except Exception as e:
            # 最後の手段として既定アプリを開く
            try:
                os.startfile(pdf_path)  # open verb
                raise RuntimeError(
                    "Adobe Acrobat が見つからないため、自動印刷できませんでした。\n"
                    "Adobe Acrobat または Acrobat Reader をインストールしてください。"
                )
            except Exception as e2:
                raise RuntimeError(f"印刷に失敗しました: {e2}")


class PreviewManager:
    """プレビュー処理を担当するクラス"""
    
    @staticmethod
    def preview_pdf(pdf_path: str) -> None:
        """システムの既定のPDFアプリケーションでプレビュー"""
        try:
            # システムの既定のアプリケーションで開く
            os.startfile(pdf_path)
        except Exception as e:
            raise RuntimeError(f"PDF表示失敗: {e}")


class PdfKeywordPrinter(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF キーワード印刷ツール")
        self.geometry(f"{Constants.WINDOW_WIDTH}x{Constants.WINDOW_HEIGHT}")
        self.resizable(True, True)

        # State vars
        self.directory_var = tk.StringVar()
        self.keyword_var = tk.StringVar(value="部品出庫")
        self.printer_var = tk.StringVar()
        self.temp_pdf_paths = []  # 一時PDFのパスリスト（削除管理用）
        self.pdf_files = []  # 検索されたPDFファイルのリスト
        self.pdf_hit_info = {}  # {pdf_path: [ヒットページリスト]}

        # Build UI
        self._build_ui()

        # Populate printers
        self._setup_printers()

        # PDF印刷機能の可用性をチェック
        self._check_printing_capabilities()

    def _build_ui(self):
        """UIを構築"""
        main_frame = tk.Frame(self)
        main_frame.pack(fill="both", expand=True, padx=Constants.PADDING, pady=Constants.PADDING)

        # 上部設定エリア
        settings_frame = tk.LabelFrame(main_frame, text="設定", padx=Constants.PADDING, pady=Constants.PADDING)
        settings_frame.pack(fill="x", pady=(0, Constants.PADDING))

        # ディレクトリ選択行
        dir_frame = tk.Frame(settings_frame)
        dir_frame.pack(fill="x", pady=Constants.PADDING)
        tk.Label(dir_frame, text="対象ディレクトリ:").grid(row=0, column=0, sticky="w")
        tk.Entry(dir_frame, textvariable=self.directory_var, width=Constants.ENTRY_WIDTH, 
                state="readonly").grid(row=0, column=1, padx=(0, Constants.PADDING), sticky="ew")
        tk.Button(dir_frame, text="参照...", command=self.browse_directory, width=10).grid(row=0, column=2, sticky="e")
        dir_frame.grid_columnconfigure(1, weight=1)

        # キーワードとプリンタ行
        controls_frame = tk.Frame(settings_frame)
        controls_frame.pack(fill="x", pady=Constants.PADDING)
        tk.Label(controls_frame, text="検索キーワード:").pack(side="left")
        tk.Entry(controls_frame, textvariable=self.keyword_var, width=Constants.KEYWORD_ENTRY_WIDTH).pack(side="left", padx=(0, Constants.PADDING))
        tk.Label(controls_frame, text="プリンタ:").pack(side="left")
        self.printer_menu = tk.OptionMenu(controls_frame, self.printer_var, "")
        self.printer_menu.config(width=Constants.PRINTER_MENU_WIDTH)
        self.printer_menu.pack(side="left")

        # 検索ボタン
        search_frame = tk.Frame(settings_frame)
        search_frame.pack(fill="x", pady=Constants.PADDING)
        tk.Button(search_frame, text="PDFファイル検索", command=self.search_pdf_files, 
                 width=Constants.BUTTON_WIDTH).pack(side="left")

        # メッセージエリア（設定とPDF一覧の間）
        message_frame = tk.LabelFrame(main_frame, text="メッセージ", padx=Constants.PADDING, pady=Constants.PADDING)
        message_frame.pack(fill="x", pady=(0, Constants.PADDING))
        
        # メッセージテキストエリア（スクロール付き）
        message_text_frame = tk.Frame(message_frame)
        message_text_frame.pack(fill="both", expand=True)
        
        self.message_text = tk.Text(message_text_frame, height=6, wrap="word", state="disabled")
        message_scrollbar = tk.Scrollbar(message_text_frame, orient="vertical", command=self.message_text.yview)
        self.message_text.configure(yscrollcommand=message_scrollbar.set)
        
        self.message_text.pack(side="left", fill="both", expand=True)
        message_scrollbar.pack(side="right", fill="y")
        
        # メッセージクリアボタン
        clear_button_frame = tk.Frame(message_frame)
        clear_button_frame.pack(fill="x", pady=(Constants.PADDING, 0))
        tk.Button(clear_button_frame, text="メッセージクリア", command=self.clear_messages, 
                 width=15).pack(side="right")

        # PDFファイル一覧エリア（Canvas+Frame+Scrollbar）
        list_frame = tk.LabelFrame(main_frame, text="PDFファイル一覧", padx=Constants.PADDING, pady=Constants.PADDING)
        list_frame.pack(fill="both", expand=True, pady=(0, Constants.PADDING))
        self.canvas = tk.Canvas(list_frame)
        self.scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview)
        self.inner_frame = tk.Frame(self.canvas)
        self.inner_frame.bind(
            "<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # ヘッダー
        self.header_row = tk.Frame(self.inner_frame)
        self.header_row.pack(fill="x")
        tk.Label(self.header_row, text="ファイル名", width=30, anchor="w", relief="ridge").grid(row=0, column=0, sticky="ew")
        tk.Label(self.header_row, text="ヒット状況", width=12, anchor="center", relief="ridge").grid(row=0, column=1, sticky="ew")
        tk.Label(self.header_row, text="プレビュー", width=10, anchor="center", relief="ridge").grid(row=0, column=2, sticky="ew")
        tk.Label(self.header_row, text="印刷", width=10, anchor="center", relief="ridge").grid(row=0, column=3, sticky="ew")

        # 下部ボタンエリア
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill="x", pady=Constants.PADDING)
        tk.Button(button_frame, text="一括印刷", command=self.batch_print, 
                 width=Constants.BUTTON_WIDTH).pack(side="right", padx=Constants.PADDING)

        # ステータスラベル（下部に常時表示）
        self.status_var = tk.StringVar()
        status_frame = tk.Frame(self)
        status_frame.pack(fill="x", side="bottom")
        tk.Label(status_frame, textvariable=self.status_var, fg="blue", anchor="w", relief="sunken", height=2).pack(fill="x")

    def _setup_printers(self):
        """プリンタ設定"""
        if PRINTER_SUPPORT:
            try:
                self._populate_printers()
            except Exception as e:
                self.printer_menu.configure(state="disabled")
                self.printer_var.set("(既定)")
                self.status_var.set(f"プリンタ情報取得失敗: {e}")
        else:
            self.printer_menu.configure(state="disabled")
            self.printer_var.set("(既定)")

    def _populate_printers(self):
        """プリンタリストを設定"""
        printers = PrinterManager.get_available_printers()
        current_default = PrinterManager.get_default_printer()
        
        menu = self.printer_menu["menu"]
        menu.delete(0, "end")
        
        for printer in printers:
            menu.add_command(label=printer, command=tk._setit(self.printer_var, printer))
        
        # 既定プリンタを設定
        if current_default in printers:
            self.printer_var.set(current_default)
        elif printers:
            self.printer_var.set(printers[0])
        else:
            self.printer_var.set("(既定)")

    def _check_printing_capabilities(self):
        """Windows標準印刷機能の確認"""
        self.status_var.set("印刷: Windows標準機能を使用（デスクトップ/output/日時フォルダに一時保存）")
        self.add_message("アプリケーションが起動しました。印刷: Windows標準機能を使用（デスクトップ/output/日時フォルダに一時保存）", "info")

    def browse_directory(self):
        """ディレクトリを選択"""
        directory = filedialog.askdirectory()
        if directory:
            self.directory_var.set(directory)
            self.status_var.set(f"選択: {directory}")
            self.add_message(f"ディレクトリを選択しました: {directory}", "info")

    def search_pdf_files(self):
        """PDFファイルを検索"""
        directory = self.directory_var.get()
        if not directory or not os.path.isdir(directory):
            self.status_var.set("エラー: ディレクトリを選択してください。")
            self.add_message("エラー: ディレクトリを選択してください。", "error")
            messagebox.showerror("エラー", "ディレクトリを選択してください。")
            return

        self.status_var.set("PDFファイルを検索中...")
        self.add_message("PDFファイルを検索中...", "info")
        
        # バックグラウンドで検索
        threading.Thread(
            target=self._search_pdf_files_background,
            args=(directory,),
            daemon=True
        ).start()

    def _search_pdf_files_background(self, directory: str):
        """バックグラウンドでPDFファイルを検索"""
        try:
            pdf_files = PdfProcessor.find_pdf_files(directory)
            pdf_hit_info = {}
            keyword = self.keyword_var.get().strip()
            for pdf_file in pdf_files:
                try:
                    pages = PdfProcessor.find_keyword_pages(pdf_file, keyword)
                except Exception:
                    pages = []
                pdf_hit_info[pdf_file] = pages
            def update_ui():
                self.pdf_files = pdf_files
                self.pdf_hit_info = pdf_hit_info
                self._update_file_list()
                self.status_var.set(f"検索完了: {len(pdf_files)}個のPDFファイルを発見")
                self.add_message(f"検索完了: {len(pdf_files)}個のPDFファイルを発見", "success")
            self.after(0, update_ui)
        except Exception as e:
            def show_error():
                self.status_var.set(f"検索エラー: {e}")
                self.add_message(f"検索エラー: {e}", "error")
                messagebox.showerror("エラー", f"PDFファイル検索中にエラーが発生しました: {e}")
            self.after(0, show_error)

    def _update_file_list(self):
        # 既存のPDF行・ヘッダーを全て削除
        for widget in self.inner_frame.winfo_children():
            widget.destroy()
        # ヘッダー
        header_font = ("Meiryo", 10, "bold")
        tk.Label(self.inner_frame, text="ファイル名", width=30, anchor="w", relief="ridge", font=header_font).grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        tk.Label(self.inner_frame, text="ヒット状況", width=12, anchor="center", relief="ridge", font=header_font).grid(row=0, column=1, sticky="nsew", padx=1, pady=1)
        tk.Label(self.inner_frame, text="プレビュー", width=10, anchor="center", relief="ridge", font=header_font).grid(row=0, column=2, sticky="nsew", padx=1, pady=1)
        tk.Label(self.inner_frame, text="印刷", width=10, anchor="center", relief="ridge", font=header_font).grid(row=0, column=3, sticky="nsew", padx=1, pady=1)
        # 各PDF行
        for i, pdf_file in enumerate(self.pdf_files):
            filename = os.path.basename(pdf_file)
            pages = self.pdf_hit_info.get(pdf_file, [])
            tk.Label(self.inner_frame, text=filename, width=30, anchor="w").grid(row=i+1, column=0, sticky="nsew", padx=1, pady=1)
            if pages:
                hit_text = f"{len(pages)}ページ" if pages else "-"
                tk.Label(self.inner_frame, text=hit_text, width=12, anchor="center", fg="green").grid(row=i+1, column=1, sticky="nsew", padx=1, pady=1)
                tk.Button(self.inner_frame, text="プレビュー", width=10, command=lambda f=pdf_file: self.preview_single_file(f)).grid(row=i+1, column=2, sticky="nsew", padx=1, pady=1)
                tk.Button(self.inner_frame, text="印刷", width=10, command=lambda f=pdf_file: self.print_single_file(f)).grid(row=i+1, column=3, sticky="nsew", padx=1, pady=1)
            else:
                tk.Label(self.inner_frame, text="該当なし", width=12, anchor="center", fg="gray").grid(row=i+1, column=1, sticky="nsew", padx=1, pady=1)
                tk.Button(self.inner_frame, text="プレビュー", width=10, state="disabled").grid(row=i+1, column=2, sticky="nsew", padx=1, pady=1)
                tk.Button(self.inner_frame, text="印刷", width=10, state="disabled").grid(row=i+1, column=3, sticky="nsew", padx=1, pady=1)
        # 列幅・weight・minsizeを統一
        self.inner_frame.grid_columnconfigure(0, weight=3, minsize=220)
        self.inner_frame.grid_columnconfigure(1, weight=1, minsize=90)
        self.inner_frame.grid_columnconfigure(2, weight=1, minsize=80)
        self.inner_frame.grid_columnconfigure(3, weight=1, minsize=80)

    def preview_single_file(self, pdf_path: str):
        """単一ファイルのプレビュー"""
        if not self._validate_inputs():
            return

        keyword = self.keyword_var.get().strip()
        
        self._set_ui_state("disabled")
        self.status_var.set("処理中...")

        threading.Thread(
            target=self._process_and_preview_single,
            args=(pdf_path, keyword),
            daemon=True
        ).start()

    def print_single_file(self, pdf_path: str):
        """単一ファイルの印刷"""
        if not self._validate_inputs():
            return

        keyword = self.keyword_var.get().strip()
        printer = self.printer_var.get()
        
        self._set_ui_state("disabled")
        self.status_var.set("処理中...")

        threading.Thread(
            target=self._process_and_print_single,
            args=(pdf_path, keyword, printer),
            daemon=True
        ).start()

    def batch_print(self):
        if not self._validate_inputs():
            return

        # ヒットしたPDFのみを対象
        hit_pdfs = [pdf for pdf, pages in self.pdf_hit_info.items() if pages]
        if not hit_pdfs:
            self.status_var.set("エラー: キーワードに該当するPDFがありません。")
            self.add_message("エラー: キーワードに該当するPDFがありません。", "error")
            return

        keyword = self.keyword_var.get().strip()
        printer = self.printer_var.get()
        self._set_ui_state("disabled")
        self.status_var.set(f"一括印刷処理中...（{len(hit_pdfs)}件）")
        self.add_message(f"一括印刷処理中...（{len(hit_pdfs)}件）", "info")
        threading.Thread(
            target=self._process_batch_print,
            args=(hit_pdfs, keyword, printer),
            daemon=True
        ).start()

    def _validate_inputs(self) -> bool:
        """入力値の検証"""
        directory = self.directory_var.get()
        keyword = self.keyword_var.get().strip()

        if not directory or not os.path.isdir(directory):
            self.status_var.set("エラー: ディレクトリを選択してください。")
            self.add_message("エラー: ディレクトリを選択してください。", "error")
            return False
        
        if not keyword:
            self.status_var.set("エラー: 検索キーワードを入力してください。")
            self.add_message("エラー: 検索キーワードを入力してください。", "error")
            return False

        # 依存パッケージチェック
        missing = []
        try:
            import pdfplumber
        except ImportError:
            missing.append('pdfplumber')
        try:
            import pypdf
        except ImportError:
            missing.append('pypdf')
        
        if missing:
            self.status_var.set(f"エラー: 必要なパッケージがありません: {', '.join(missing)} (pip install ...)")
            self.add_message(f"エラー: 必要なパッケージがありません: {', '.join(missing)} (pip install ...)", "error")
            return False

        return True

    def _set_ui_state(self, state: str):
        """UIの有効/無効を切り替え"""
        for child in self.winfo_children():
            try:
                child.configure(state=state)
            except tk.TclError:
                pass  # not all widgets support state

    def _process_and_preview_single(self, pdf_path: str, keyword: str):
        """単一ファイルのプレビュー処理"""
        temp_pdf = None
        try:
            pages = PdfProcessor.find_keyword_pages(pdf_path, keyword)
            if not pages:
                self._done("キーワードを含むページが見つかりませんでした。", error=True)
                return

            temp_pdf = PdfProcessor.extract_pages(pdf_path, pages)
            self.temp_pdf_paths.append(temp_pdf)

            PreviewManager.preview_pdf(temp_pdf)
            self._done(f"プレビュー表示 (ページ: {', '.join(map(str, [p+1 for p in pages]))})", preview=True)
            
        except Exception as e:
            self._cleanup_temp_file(temp_pdf)
            self._done(f"エラー: {e}", error=True)

    def _process_and_print_single(self, pdf_path: str, keyword: str, printer: str):
        """単一ファイルの印刷処理（UltraThink対応）"""
        temp_pdf = None
        output_pdf = None
        try:
            pages = PdfProcessor.find_keyword_pages(pdf_path, keyword)
            if not pages:
                self._done("キーワードを含むページが見つかりませんでした。", error=True)
                return

            temp_pdf = PdfProcessor.extract_pages(pdf_path, pages)
            
            # ファイルの安定化を待つ
            time.sleep(Constants.FILE_STABILITY_WAIT)
            
            # デスクトップ/output/日時フォルダにファイルをコピー
            output_pdf = self._copy_to_desktop(temp_pdf, pdf_path, keyword)
            
            # 出力ファイルで印刷を試行
            try:
                PrinterManager.print_pdf(output_pdf, printer)
                time.sleep(Constants.FILE_ACCESS_WAIT)
            except Exception as e:
                raise RuntimeError(f"印刷に失敗しました: {e}")
            
            self._done(f"印刷完了 (ページ: {', '.join(map(str, [p+1 for p in pages]))})", printed=True)
            
        except Exception as e:
            self._done(f"エラー: {e}", error=True)
        finally:
            # 一時ファイルと出力ファイルをクリーンアップ
            self._cleanup_temp_file(temp_pdf)
            #self._cleanup_temp_file(output_pdf)

    def _process_batch_print(self, pdf_files: List[str], keyword: str, printer: str):
        """一括印刷処理（UltraThink対応）"""
        processed_count = 0
        error_count = 0
        temp_files = []
        output_files = []
        try:
            for i, pdf_path in enumerate(pdf_files):
                temp_pdf = None
                output_pdf = None
                try:
                    self.status_var.set(f"一括印刷中... ({i+1}/{len(pdf_files)})")
                    pages = PdfProcessor.find_keyword_pages(pdf_path, keyword)
                    if not pages:
                        continue
                    
                    temp_pdf = PdfProcessor.extract_pages(pdf_path, pages)
                    temp_files.append(temp_pdf)
                    
                    # ファイルの安定化を待つ
                    time.sleep(Constants.FILE_STABILITY_WAIT)
                    
                    # デスクトップ/output/日時フォルダにファイルをコピー
                    output_pdf = self._copy_to_desktop(temp_pdf, pdf_path, keyword)
                    output_files.append(output_pdf)
                    
                    # 出力ファイルで印刷を試行
                    try:
                        PrinterManager.print_pdf(output_pdf, printer)
                        time.sleep(Constants.FILE_ACCESS_WAIT)
                        processed_count += 1
                    except Exception as e:
                        error_count += 1
                        print(f"印刷エラー ({pdf_path}): {e}")
                            
                except Exception as e:
                    error_count += 1
                    print(f"エラー ({pdf_path}): {e}")
                finally:
                    # 個別ファイルのクリーンアップ
                    self._cleanup_temp_file(temp_pdf)
                    self._cleanup_temp_file(output_pdf)
            
            self._done(f"一括印刷完了: {processed_count}件成功, {error_count}件エラー（ヒットPDFのみ）", printed=True)
        except Exception as e:
            self._done(f"一括印刷エラー: {e}", error=True)
        finally:
            # 残りのファイルをクリーンアップ
            for temp_file in temp_files:
                self._cleanup_temp_file(temp_file)
            for output_file in output_files:
                self._cleanup_temp_file(output_file)

    def _copy_to_desktop(self, temp_pdf: str, original_pdf_path: str = None, keyword: str = None) -> str:
        """一時ファイルをデスクトップ/output/日時フォルダにコピー"""
        # デスクトップフォルダのパスを取得
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        
        # outputフォルダを作成
        output_dir = os.path.join(desktop_path, Constants.OUTPUT_FOLDER_NAME)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # 日時フォルダを作成
        current_datetime = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        datetime_dir = os.path.join(output_dir, current_datetime)
        if not os.path.exists(datetime_dir):
            os.makedirs(datetime_dir)

        # ファイル名を生成
        if original_pdf_path and keyword:
            # 元のPDFファイル名を取得
            original_filename = os.path.splitext(os.path.basename(original_pdf_path))[0]
            # キーワードを安全なファイル名に変換
            safe_keyword = self._make_safe_filename(keyword)
            # 新しいファイル名を生成（長さ制限を考慮）
            base_filename = f"{original_filename}_{safe_keyword}抽出版"
            # ファイル名の長さを制限（拡張子を含めて255文字以内）
            if len(base_filename) > 240:  # .pdf の5文字を考慮
                base_filename = base_filename[:240]
            new_filename = f"{base_filename}.pdf"
        else:
            # フォールバック: 元の一時ファイル名を使用
            filename = os.path.basename(temp_pdf)
            new_filename = filename
        
        desktop_pdf = os.path.join(datetime_dir, new_filename)
        
        try:
            shutil.copy2(temp_pdf, desktop_pdf)
            return desktop_pdf
        except Exception:
            # デスクトップコピーが失敗した場合は元のファイルを使用
            return temp_pdf

    def _make_safe_filename(self, filename: str) -> str:
        """ファイル名を安全な形式に変換"""
        # Windowsで使用できない文字を置換
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        
        # 連続するアンダースコアを単一のアンダースコアに置換
        while '__' in filename:
            filename = filename.replace('__', '_')
        
        # 先頭と末尾のアンダースコアを削除
        filename = filename.strip('_')
        
        # 空文字列の場合はデフォルト値を設定
        if not filename:
            filename = "keyword"
        
        return filename

    def _cleanup_temp_file(self, file_path: Optional[str]):
        """一時ファイルを削除"""
        if file_path and os.path.exists(file_path):
            try:
                os.unlink(file_path)
            except OSError:
                pass

    def _done(self, msg: str, error=False, preview=False, printed=False):
        """処理完了時の処理"""
        def finish():
            self._set_ui_state("normal")
            self.status_var.set(msg)
            
            # メッセージエリアにも追加
            if error:
                self.add_message(msg, "error")
            elif preview:
                self.add_message(msg, "info")
            elif printed:
                self.add_message(msg, "success")
            else:
                self.add_message(msg, "info")
        self.after(0, finish)

    def add_message(self, message: str, message_type: str = "info"):
        """メッセージエリアにメッセージを追加"""
        def update_message():
            self.message_text.configure(state="normal")
            
            # タイムスタンプを追加
            timestamp = time.strftime("%H:%M:%S")
            
            # メッセージタイプに応じて色を設定
            if message_type == "error":
                tag = "error"
                self.message_text.tag_config("error", foreground="red")
            elif message_type == "success":
                tag = "success"
                self.message_text.tag_config("success", foreground="green")
            elif message_type == "warning":
                tag = "warning"
                self.message_text.tag_config("warning", foreground="orange")
            else:
                tag = "info"
                self.message_text.tag_config("info", foreground="black")
            
            # メッセージを追加
            self.message_text.insert("end", f"[{timestamp}] {message}\n", tag)
            
            # 自動スクロール
            self.message_text.see("end")
            self.message_text.configure(state="disabled")
        
        self.after(0, update_message)

    def clear_messages(self):
        """メッセージエリアをクリア"""
        def clear():
            self.message_text.configure(state="normal")
            self.message_text.delete(1.0, "end")
            self.message_text.configure(state="disabled")
        
        self.after(0, clear)


if __name__ == "__main__":
    app = PdfKeywordPrinter()
    app.mainloop()
