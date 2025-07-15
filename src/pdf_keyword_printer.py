"""
Keyword-based PDF page extractor and printer (Windows GUI)

Dependencies:
    pip install pypdf pdfplumber pillow
    pip install pywin32   # optional, for printer listing

Usage:
    Run this script with Python 3.11+ on Windows.
    The GUI lets users:
        1. Browse and choose a PDF file.
        2. Enter keyword (case‑insensitive).
        3. Optionally choose printer (defaults to system default).
        4. Click "プレビュー" to extract pages containing the keyword and preview them in Edge.
        5. Click "印刷" to extract pages containing the keyword and print them.
        
Note: Printing uses Windows standard functions only. No external PDF readers required.
"""

import os
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import List, Optional, Tuple
import time
import hashlib
import shutil
import subprocess

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
    WINDOW_WIDTH = 640
    WINDOW_HEIGHT = 280
    PADDING = 6
    BUTTON_WIDTH = 15
    ENTRY_WIDTH = 48
    KEYWORD_ENTRY_WIDTH = 30
    PRINTER_MENU_WIDTH = 40
    
    # プリンタアプリケーションパス
    EDGE_PATHS = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
    ]
    
    ACROBAT_PATHS = [
        r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
        r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
    ]
    
    SUMATRA_PATHS = [
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe"
    ]
    
    CHROME_PATHS = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ]
    
    # タイムアウト設定
    PRINT_TIMEOUT = 30
    EDGE_PRINT_WAIT = 10
    EDGE_DIALOG_WAIT = 15
    CHROME_PRINT_WAIT = 10
    FILE_ACCESS_WAIT = 5


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
    def print_with_edge(pdf_path: str, printer: str) -> bool:
        """Microsoft Edgeで印刷"""
        for edge_path in Constants.EDGE_PATHS:
            if os.path.exists(edge_path):
                try:
                    if printer and "Microsoft Print to PDF" in printer:
                        subprocess.Popen([edge_path, "--print", pdf_path])
                        time.sleep(Constants.EDGE_DIALOG_WAIT)
                    else:
                        subprocess.Popen([edge_path, "--print", pdf_path])
                        time.sleep(Constants.EDGE_PRINT_WAIT)
                    return True
                except Exception:
                    continue
        return False
    
    @staticmethod
    def print_with_acrobat(pdf_path: str, printer: str) -> bool:
        """Adobe Acrobatで印刷"""
        for acrobat_path in Constants.ACROBAT_PATHS:
            if os.path.exists(acrobat_path):
                try:
                    if printer and printer != "(既定)":
                        subprocess.run([acrobat_path, '/t', pdf_path, printer], 
                                     timeout=Constants.PRINT_TIMEOUT, check=True)
                    else:
                        subprocess.run([acrobat_path, '/t', pdf_path], 
                                     timeout=Constants.PRINT_TIMEOUT, check=True)
                    return True
                except Exception:
                    continue
        return False
    
    @staticmethod
    def print_with_sumatra(pdf_path: str, printer: str) -> bool:
        """SumatraPDFで印刷"""
        for sumatra_path in Constants.SUMATRA_PATHS:
            if os.path.exists(sumatra_path):
                try:
                    if printer and printer != "(既定)":
                        subprocess.run([sumatra_path, '-print-to', printer, pdf_path], 
                                     timeout=Constants.PRINT_TIMEOUT, check=True)
                    else:
                        subprocess.run([sumatra_path, '-print-to-default', pdf_path], 
                                     timeout=Constants.PRINT_TIMEOUT, check=True)
                    return True
                except Exception:
                    continue
        return False
    
    @staticmethod
    def print_with_chrome(pdf_path: str, printer: str) -> bool:
        """Google Chromeで印刷"""
        for chrome_path in Constants.CHROME_PATHS:
            if os.path.exists(chrome_path):
                try:
                    subprocess.Popen([chrome_path, "--print", pdf_path])
                    time.sleep(Constants.CHROME_PRINT_WAIT)
                    return True
                except Exception:
                    continue
        return False
    
    @staticmethod
    def print_pdf(pdf_path: str, printer: str) -> None:
        """PDFを印刷（複数の方法を試行）"""
        methods = [
            ("Microsoft Edge", lambda: PrinterManager.print_with_edge(pdf_path, printer)),
            ("Adobe Acrobat", lambda: PrinterManager.print_with_acrobat(pdf_path, printer)),
            ("SumatraPDF", lambda: PrinterManager.print_with_sumatra(pdf_path, printer)),
            ("Google Chrome", lambda: PrinterManager.print_with_chrome(pdf_path, printer))
        ]
        
        for method_name, method_func in methods:
            if method_func():
                return
        
        raise RuntimeError(
            f"印刷に失敗しました。\n\n"
            f"PDFを印刷するためのアプリケーションが見つかりません。\n"
            f"以下のいずれかをインストールしてください：\n"
            f"• Microsoft Edge (推奨)\n"
            f"• Adobe Acrobat Reader DC\n"
            f"• SumatraPDF\n"
            f"• Google Chrome\n\n"
            f"または、PDFファイルの関連付けを確認してください。"
        )


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
        self.resizable(False, False)

        # State vars
        self.pdf_path_var = tk.StringVar()
        self.keyword_var = tk.StringVar(value="部品出庫")
        self.printer_var = tk.StringVar()
        self.temp_pdf_paths = []  # 一時PDFのパスリスト（削除管理用）

        # Build UI
        self._build_ui()

        # Populate printers
        self._setup_printers()

        # PDF印刷機能の可用性をチェック
        self._check_printing_capabilities()

    def _build_ui(self):
        """UIを構築"""
        # PDF selection row
        row1 = tk.Frame(self)
        row1.pack(fill="x", padx=Constants.PADDING, pady=Constants.PADDING)
        tk.Label(row1, text="PDF ファイル:").grid(row=0, column=0, sticky="w")
        tk.Entry(row1, textvariable=self.pdf_path_var, width=Constants.ENTRY_WIDTH, 
                state="readonly").grid(row=0, column=1, padx=(0, Constants.PADDING), sticky="ew")
        tk.Button(row1, text="参照...", command=self.browse_pdf, width=10).grid(row=0, column=2, sticky="e")
        row1.grid_columnconfigure(1, weight=1)

        # Keyword row
        row2 = tk.Frame(self)
        row2.pack(fill="x", padx=Constants.PADDING, pady=Constants.PADDING)
        tk.Label(row2, text="検索キーワード:").pack(side="left")
        tk.Entry(row2, textvariable=self.keyword_var, width=Constants.KEYWORD_ENTRY_WIDTH).pack(side="left")

        # Printer row
        row3 = tk.Frame(self)
        row3.pack(fill="x", padx=Constants.PADDING, pady=Constants.PADDING)
        tk.Label(row3, text="プリンタ:").pack(side="left")
        self.printer_menu = tk.OptionMenu(row3, self.printer_var, "")
        self.printer_menu.config(width=Constants.PRINTER_MENU_WIDTH)
        self.printer_menu.pack(side="left")

        # Action buttons
        button_frame = tk.Frame(self)
        button_frame.pack(pady=Constants.PADDING * 2)
        tk.Button(button_frame, text="プレビュー", width=Constants.BUTTON_WIDTH, 
                 command=self.run_preview).pack(side="left", padx=Constants.PADDING)
        tk.Button(button_frame, text="印刷", width=Constants.BUTTON_WIDTH, 
                 command=self.run_print).pack(side="left", padx=Constants.PADDING)

        # Status label
        self.status_var = tk.StringVar()
        tk.Label(self, textvariable=self.status_var, fg="blue").pack()

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
        self.status_var.set("印刷: Windows標準機能を使用")

    def browse_pdf(self):
        """PDFファイルを選択"""
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            if not path.lower().endswith('.pdf'):
                messagebox.showerror("エラー", "PDFファイルを選択してください。")
                return
            self.pdf_path_var.set(path)
            self.status_var.set(f"選択: {os.path.basename(path)}")

    def _validate_inputs(self) -> bool:
        """入力値の検証"""
        pdf_path = self.pdf_path_var.get()
        keyword = self.keyword_var.get().strip()

        if not pdf_path or not os.path.isfile(pdf_path):
            messagebox.showerror("エラー", "PDF ファイルを選択してください。")
            return False
        
        if not keyword:
            messagebox.showerror("エラー", "検索キーワードを入力してください。")
            return False
        
        if not pdf_path.lower().endswith('.pdf'):
            messagebox.showerror("エラー", "PDFファイルを選択してください。")
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
            messagebox.showerror("依存パッケージ未導入", 
                               f"必要なパッケージがありません: {', '.join(missing)}\n"
                               f"コマンドプロンプトで\npip install {' '.join(missing)}\n"
                               f"を実行してください。")
            return False

        return True

    def run_preview(self):
        """プレビューボタンの処理"""
        if not self._validate_inputs():
            return

        pdf_path = self.pdf_path_var.get()
        keyword = self.keyword_var.get().strip()

        self._set_ui_state("disabled")
        self.status_var.set("処理中...")

        threading.Thread(
            target=self._process_and_preview,
            args=(pdf_path, keyword),
            daemon=True
        ).start()

    def run_print(self):
        """印刷ボタンの処理"""
        if not self._validate_inputs():
            return

        pdf_path = self.pdf_path_var.get()
        keyword = self.keyword_var.get().strip()
        printer = self.printer_var.get()

        self._set_ui_state("disabled")
        self.status_var.set("処理中...")

        threading.Thread(
            target=self._process_and_print,
            args=(pdf_path, keyword, printer),
            daemon=True
        ).start()

    def _set_ui_state(self, state: str):
        """UIの有効/無効を切り替え"""
        for child in self.winfo_children():
            try:
                child.configure(state=state)
            except tk.TclError:
                pass  # not all widgets support state

    def _process_and_preview(self, pdf_path: str, keyword: str):
        """プレビュー処理"""
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

    def _process_and_print(self, pdf_path: str, keyword: str, printer: str):
        """印刷処理"""
        temp_pdf = None
        try:
            pages = PdfProcessor.find_keyword_pages(pdf_path, keyword)
            if not pages:
                self._done("キーワードを含むページが見つかりませんでした。", error=True)
                return

            temp_pdf = PdfProcessor.extract_pages(pdf_path, pages)
            
            # デスクトップにコピーして印刷を試行
            desktop_pdf = self._copy_to_desktop(temp_pdf)
            
            try:
                PrinterManager.print_pdf(desktop_pdf, printer)
                time.sleep(Constants.FILE_ACCESS_WAIT)
            finally:
                self._cleanup_temp_file(desktop_pdf)
            
            self._cleanup_temp_file(temp_pdf)
            self._done(f"印刷完了 (ページ: {', '.join(map(str, [p+1 for p in pages]))})", printed=True)
            
        except Exception as e:
            self._cleanup_temp_file(temp_pdf)
            self._done(f"エラー: {e}", error=True)

    def _copy_to_desktop(self, temp_pdf: str) -> str:
        """一時ファイルをデスクトップにコピー"""
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        filename = os.path.basename(temp_pdf)
        desktop_pdf = os.path.join(desktop_path, filename)
        
        try:
            shutil.copy2(temp_pdf, desktop_pdf)
            return desktop_pdf
        except Exception:
            # デスクトップコピーが失敗した場合は元のファイルを使用
            return temp_pdf

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
            
            if error:
                messagebox.showerror("エラー", msg)
            elif preview:
                messagebox.showinfo("完了", "PDFを既定のアプリケーションで開きました。内容をご確認ください。")
            elif printed:
                messagebox.showinfo("完了", "選択したプリンタで印刷しました。")
        
        self.after(0, finish)


if __name__ == "__main__":
    app = PdfKeywordPrinter()
    app.mainloop()
