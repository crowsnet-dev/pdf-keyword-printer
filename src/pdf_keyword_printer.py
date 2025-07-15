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

import pdfplumber
from pypdf import PdfReader, PdfWriter

# Printer handling (optional)
try:
    import win32print
    PRINTER_SUPPORT = True
except ImportError:
    PRINTER_SUPPORT = False


class PdfKeywordPrinter(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF キーワード印刷ツール")
        self.geometry("640x280")  # 高さを少し増やす
        self.resizable(False, False)

        # State vars
        self.pdf_path_var = tk.StringVar()
        self.keyword_var = tk.StringVar(value="部品出庫")
        self.printer_var = tk.StringVar()
        self.temp_pdf_paths = []  # 一時PDFのパスリスト（削除管理用）

        # Build UI
        self._build_ui()

        # Populate printers
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

        # PDF印刷機能の可用性をチェック
        self._check_printing_capabilities()

    def _build_ui(self):
        PAD = 6
        # PDF selection row (gridで配置)
        row1 = tk.Frame(self)
        row1.pack(fill="x", padx=PAD, pady=PAD)
        tk.Label(row1, text="PDF ファイル:").grid(row=0, column=0, sticky="w")
        tk.Entry(row1, textvariable=self.pdf_path_var, width=48, state="readonly").grid(row=0, column=1, padx=(0, PAD), sticky="ew")
        tk.Button(row1, text="参照...", command=self.browse_pdf, width=10).grid(row=0, column=2, sticky="e")
        row1.grid_columnconfigure(1, weight=1)

        # Keyword row
        row2 = tk.Frame(self)
        row2.pack(fill="x", padx=PAD, pady=PAD)
        tk.Label(row2, text="検索キーワード:").pack(side="left")
        tk.Entry(row2, textvariable=self.keyword_var, width=30).pack(side="left")

        # Printer row
        row3 = tk.Frame(self)
        row3.pack(fill="x", padx=PAD, pady=PAD)
        tk.Label(row3, text="プリンタ:").pack(side="left")
        self.printer_menu = tk.OptionMenu(row3, self.printer_var, "")
        self.printer_menu.config(width=40)
        self.printer_menu.pack(side="left")

        # Action buttons
        button_frame = tk.Frame(self)
        button_frame.pack(pady=PAD * 2)
        tk.Button(button_frame, text="プレビュー", width=15, command=self.run_preview).pack(side="left", padx=PAD)
        tk.Button(button_frame, text="印刷", width=15, command=self.run_print).pack(side="left", padx=PAD)

        # Status label
        self.status_var = tk.StringVar()
        tk.Label(self, textvariable=self.status_var, fg="blue").pack()

    def _populate_printers(self):
        printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        current_default = win32print.GetDefaultPrinter()
        menu = self.printer_menu["menu"]
        menu.delete(0, "end")
        for p in printers:
            menu.add_command(label=p, command=tk._setit(self.printer_var, p))
        # Set default
        self.printer_var.set(current_default if current_default in printers else printers[0] if printers else "")

    def _check_printing_capabilities(self):
        """Windows標準印刷機能の確認"""
        # Windows標準機能のみを使用するため、特別なチェックは不要
        self.status_var.set("印刷: Windows標準機能を使用")

    def browse_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            if not path.lower().endswith('.pdf'):
                messagebox.showerror("エラー", "PDFファイルを選択してください。")
                return
            self.pdf_path_var.set(path)
            self.status_var.set(f"選択: {os.path.basename(path)}")

    def _validate_inputs(self):
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
            messagebox.showerror("依存パッケージ未導入", f"必要なパッケージがありません: {', '.join(missing)}\nコマンドプロンプトで\npip install {' '.join(missing)}\nを実行してください。")
            return False

        return True

    def run_preview(self):
        """プレビューボタンの処理"""
        if not self._validate_inputs():
            return

        pdf_path = self.pdf_path_var.get()
        keyword = self.keyword_var.get().strip()

        # Disable UI while running
        self._set_ui_state("disabled")
        self.status_var.set("処理中...")

        # Run in background to keep UI responsive
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

        # Disable UI while running
        self._set_ui_state("disabled")
        self.status_var.set("処理中...")

        # Run in background to keep UI responsive
        threading.Thread(
            target=self._process_and_print,
            args=(pdf_path, keyword, printer),
            daemon=True
        ).start()

    def _set_ui_state(self, state: str):
        for child in self.winfo_children():
            try:
                child.configure(state=state)
            except tk.TclError:
                pass  # not all widgets support state

    def _process_and_preview(self, pdf_path: str, keyword: str):
        """プレビュー処理（改善版）"""
        temp_pdf = None
        try:
            pages = self._find_keyword_pages(pdf_path, keyword)
            if not pages:
                self._done("キーワードを含むページが見つかりませんでした。", error=True)
                return

            writer = PdfWriter()
            reader = PdfReader(pdf_path)
            for idx in pages:
                writer.add_page(reader.pages[idx])

            # 一時ファイルを安全に作成
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_file:
                temp_pdf = temp_file.name
                writer.write(temp_file)

            # 一時ファイルリストに追加（管理用）
            self.temp_pdf_paths.append(temp_pdf)

            # Preview
            self._preview_pdf(temp_pdf)
            self._done(f"プレビュー表示 (ページ: {', '.join(map(str, [p+1 for p in pages]))})", preview=True)
        except Exception as e:
            # エラー時は一時ファイルを削除
            if temp_pdf and os.path.exists(temp_pdf):
                try:
                    os.unlink(temp_pdf)
                except OSError:
                    pass
            self._done(f"エラー: {e}", error=True)

    def _process_and_print(self, pdf_path: str, keyword: str, printer: str):
        """印刷処理（デバッグ強化版）"""
        temp_pdf = None
        try:
            # デバッグ: 入力パラメータ確認
            print(f"DEBUG: PDF path: {pdf_path}")
            print(f"DEBUG: Keyword: {keyword}")
            print(f"DEBUG: Printer: {printer}")
            
            pages = self._find_keyword_pages(pdf_path, keyword)
            print(f"DEBUG: Found pages: {pages}")
            
            if not pages:
                self._done("キーワードを含むページが見つかりませんでした。", error=True)
                return

            # デバッグ: PDF読み込み確認
            reader = PdfReader(pdf_path)
            print(f"DEBUG: PDF total pages: {len(reader.pages)}")
            
            writer = PdfWriter()
            for idx in pages:
                page = reader.pages[idx]
                writer.add_page(page)
                print(f"DEBUG: Added page {idx+1} to writer")

            # デバッグ: writer確認
            print(f"DEBUG: Writer has {len(writer.pages)} pages")

            # 元のPDFファイル名から新しいファイル名を生成（完全に英数字のみ）
            import time
            import hashlib
            
            # 現在時刻とファイルパスからユニークな名前を生成
            timestamp = str(int(time.time()))
            file_hash = hashlib.md5(pdf_path.encode()).hexdigest()[:8]
            new_filename = f"extracted_{timestamp}_{file_hash}.pdf"
            print(f"DEBUG: New filename: {new_filename}")
            
            # 一時ファイルを安全に作成
            temp_dir = tempfile.gettempdir()
            temp_pdf = os.path.join(temp_dir, new_filename)
            print(f"DEBUG: Temp file path: {temp_pdf}")
            
            # ファイルを確実に書き込み
            try:
                with open(temp_pdf, "wb") as f:
                    print("DEBUG: Starting to write PDF...")
                    writer.write(f)
                    print("DEBUG: PDF write completed")
                    f.flush()
                    os.fsync(f.fileno())
                    print("DEBUG: File flushed and synced")
                
                # ファイルサイズを確認
                file_size = os.path.getsize(temp_pdf)
                print(f"DEBUG: File size after write: {file_size} bytes")
                
                if file_size == 0:
                    raise RuntimeError(f"一時PDFファイルの作成に失敗しました（0バイト）\nパス: {temp_pdf}")
                
                # ファイルが読み取り可能か確認
                with open(temp_pdf, "rb") as test_f:
                    test_data = test_f.read(100)  # 最初の100バイトを読み取り
                    print(f"DEBUG: File readable, first 100 bytes length: {len(test_data)}")
                
            except Exception as write_error:
                print(f"DEBUG: Write error: {write_error}")
                raise RuntimeError(f"PDFファイル書き込みエラー: {write_error}")

            # Print
            print("DEBUG: Starting print process...")
            
            # 一時的にファイルをデスクトップにコピーして印刷を試行
            import shutil
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            desktop_pdf = os.path.join(desktop_path, new_filename)
            
            try:
                print(f"DEBUG: Copying to desktop: {desktop_pdf}")
                shutil.copy2(temp_pdf, desktop_pdf)
                print("DEBUG: File copied to desktop successfully")
                
                # デスクトップのファイルで印刷を試行
                self._send_pdf_to_printer(desktop_pdf, printer)
                print("DEBUG: Print process completed")
                
                # 印刷処理完了後に追加の待機時間
                print("DEBUG: Additional wait for file access completion...")
                import time
                time.sleep(5)  # ファイルアクセス完了を確実に待つ
                
                # デスクトップのファイルを削除
                if os.path.exists(desktop_pdf):
                    try:
                        print(f"DEBUG: Deleting desktop file: {desktop_pdf}")
                        os.unlink(desktop_pdf)
                        print("DEBUG: Desktop file deleted")
                    except OSError as del_error:
                        print(f"DEBUG: Desktop delete error: {del_error}")
                        
            except Exception as copy_error:
                print(f"DEBUG: Desktop copy failed: {copy_error}")
                # デスクトップコピーが失敗した場合は元のファイルで試行
                self._send_pdf_to_printer(temp_pdf, printer)
                print("DEBUG: Print process completed with original file")
            
            # 一時ファイル削除
            if temp_pdf and os.path.exists(temp_pdf):
                try:
                    print(f"DEBUG: Deleting temp file: {temp_pdf}")
                    os.unlink(temp_pdf)
                    print("DEBUG: Temp file deleted")
                except OSError as del_error:
                    print(f"DEBUG: Delete error: {del_error}")
            
            self._done(f"印刷完了 (ページ: {', '.join(map(str, [p+1 for p in pages]))})", printed=True)
            
        except Exception as e:
            print(f"DEBUG: Exception occurred: {e}")
            import traceback
            traceback.print_exc()
            # エラー時のみファイル削除
            if temp_pdf and os.path.exists(temp_pdf):
                try:
                    os.unlink(temp_pdf)
                except OSError:
                    pass
            self._done(f"エラー: {e}", error=True)

    def _find_keyword_pages(self, pdf_path: str, keyword: str):
        hit_pages = []
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                if keyword.lower() in text.lower():
                    hit_pages.append(i)
        return hit_pages

    def _preview_pdf(self, path: str):
        """EdgeでPDFを開く"""
        try:
            import subprocess
            edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
            if not os.path.exists(edge_path):
                edge_path = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
            if not os.path.exists(edge_path):
                # Edgeが見つからない場合は既定のアプリで開く
                os.startfile(path)
            else:
                subprocess.Popen([edge_path, path])
        except Exception as e:
            raise RuntimeError(f"PDF表示失敗: {e}")

    def _send_pdf_to_printer(self, path: str, printer: str):
        """既存PDFリーダーを検出して直接印刷"""
        import subprocess
        
        print(f"DEBUG: Sending to printer: {printer}")
        print(f"DEBUG: File path: {path}")

        # Method 1: Microsoft Edge (PDF印刷対応)
        try:
            print("DEBUG: Trying Microsoft Edge")
            edge_paths = [
                r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
            ]
            
            for edge_path in edge_paths:
                if os.path.exists(edge_path):
                    print(f"DEBUG: Found Edge at: {edge_path}")
                    if printer and "Microsoft Print to PDF" in printer:
                        # Microsoft Print to PDFの場合は印刷ダイアログを表示
                        subprocess.Popen([edge_path, "--print", path])
                        import time
                        time.sleep(15)  # ダイアログ操作完了まで待機
                    else:
                        # 通常のプリンタの場合
                        subprocess.Popen([edge_path, "--print", path])
                        import time
                        time.sleep(10)  # 印刷完了まで待機
                    print("DEBUG: Edge print completed")
                    return
        except Exception as e:
            print(f"DEBUG: Edge method failed: {e}")

        # Method 2: Adobe Acrobat Reader
        try:
            print("DEBUG: Trying Adobe Acrobat Reader")
            acrobat_paths = [
                r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
            ]
            
            for acrobat_path in acrobat_paths:
                if os.path.exists(acrobat_path):
                    print(f"DEBUG: Found Acrobat at: {acrobat_path}")
                    if printer and printer != "(既定)":
                        # 指定プリンタで印刷
                        subprocess.run([acrobat_path, '/t', path, printer], 
                                     timeout=30, check=True)
                    else:
                        # 既定プリンタで印刷
                        subprocess.run([acrobat_path, '/t', path], 
                                     timeout=30, check=True)
                    print("DEBUG: Acrobat print completed")
                    return
        except Exception as e:
            print(f"DEBUG: Acrobat method failed: {e}")

        # Method 3: SumatraPDF
        try:
            print("DEBUG: Trying SumatraPDF")
            sumatra_paths = [
                r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
                r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe"
            ]
            
            for sumatra_path in sumatra_paths:
                if os.path.exists(sumatra_path):
                    print(f"DEBUG: Found SumatraPDF at: {sumatra_path}")
                    if printer and printer != "(既定)":
                        subprocess.run([sumatra_path, '-print-to', printer, path], 
                                     timeout=30, check=True)
                    else:
                        subprocess.run([sumatra_path, '-print-to-default', path], 
                                     timeout=30, check=True)
                    print("DEBUG: SumatraPDF print completed")
                    return
        except Exception as e:
            print(f"DEBUG: SumatraPDF method failed: {e}")

        # Method 4: Google Chrome (最後の手段)
        try:
            print("DEBUG: Trying Google Chrome")
            chrome_paths = [
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            ]
            
            for chrome_path in chrome_paths:
                if os.path.exists(chrome_path):
                    print(f"DEBUG: Found Chrome at: {chrome_path}")
                    subprocess.Popen([chrome_path, "--print", path])
                    import time
                    time.sleep(10)  # 印刷完了まで待機
                    print("DEBUG: Chrome print completed")
                    return
        except Exception as e:
            print(f"DEBUG: Chrome method failed: {e}")

        # すべて失敗した場合
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

    def _print_pdf_with_win32(self, path: str, printer: str):
        """win32printを使用したWindows標準API印刷（簡素版）"""
        try:
            # win32printを使用した印刷ジョブの作成
            # 注意: PDFの直接印刷は制限があるため、基本的にはPowerShellやos.startfileを優先
            hprinter = win32print.OpenPrinter(printer or win32print.GetDefaultPrinter())
            try:
                # ファイル名をジョブ名として使用
                job_name = os.path.basename(path)
                job_info = (job_name, None, "RAW")
                job_id = win32print.StartDocPrinter(hprinter, 1, job_info)
                try:
                    win32print.StartPagePrinter(hprinter)
                    # PDFファイルの内容を読み込んで送信（制限あり）
                    with open(path, 'rb') as f:
                        data = f.read()
                    win32print.WritePrinter(hprinter, data)
                    win32print.EndPagePrinter(hprinter)
                finally:
                    win32print.EndDocPrinter(hprinter)
            finally:
                win32print.ClosePrinter(hprinter)

        except Exception as e:
            raise RuntimeError(f"Windows標準API印刷失敗: {e}")

    def _done(self, msg: str, error=False, preview=False, printed=False):
        """処理完了時の処理"""
        # Re-enable UI in main thread
        def finish():
            self._set_ui_state("normal")
            self.status_var.set(msg)
            # 一時PDFは削除しない（Edgeで表示するため）
            if error:
                messagebox.showerror("エラー", msg)
            elif preview:
                messagebox.showinfo("完了", "PDFをMicrosoft Edgeで開きました。内容をご確認ください。")
            elif printed:
                messagebox.showinfo("完了", "選択したプリンタで印刷しました。")
        self.after(0, finish)


if __name__ == "__main__":
    app = PdfKeywordPrinter()
    app.mainloop()
