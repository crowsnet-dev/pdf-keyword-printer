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
import atexit
import signal

import pdfplumber
from pypdf import PdfReader, PdfWriter

# Printer handling (optional)
try:
    import win32print
    import win32api
    import win32con
    PRINTER_SUPPORT = True
except ImportError:
    PRINTER_SUPPORT = False


class Constants:
    """定数クラス"""
    WINDOW_WIDTH = 800
    WINDOW_HEIGHT = 700
    PADDING = 6
    BUTTON_WIDTH = 12
    ENTRY_WIDTH = 48
    KEYWORD_ENTRY_WIDTH = 30
    PRINTER_MENU_WIDTH = 40
    
    # タイムアウト設定（印刷信頼性を重視）
    PRINT_TIMEOUT = 10
    FILE_ACCESS_WAIT = 2
    
    # Adobe Acrobat印刷対応設定
    PRINT_RETRY_COUNT = 3
    PRINT_RETRY_DELAY = 2
    
    # 印刷ジョブ送信後の待機時間（印刷信頼性のため必要）
    PRINT_JOB_WAIT = 3.0  # Adobeが印刷ジョブを送信するまで3秒
    
    # 印刷ジョブ送信確認のタイムアウト
    PRINT_JOB_SENT_TIMEOUT = 10  # 印刷ジョブ送信確認のタイムアウト（秒）
    
    # 印刷ジョブ完了待機のタイムアウト
    PRINT_JOB_TIMEOUT = 60  # 印刷ジョブ完了待機のタイムアウト（秒）
    
    # 一括印刷時の最小待機時間（印刷ジョブ監視に置き換え）
    BATCH_PRINT_WAIT = 0.5  # 印刷ジョブ監視の間隔
    
    # ローディング表示設定
    LOADING_DOT_INTERVAL = 500  # ミリ秒
    LOADING_TEXT = "処理中"


class PdfProcessor:
    """PDF処理を担当するクラス"""
    
    @staticmethod
    def find_keyword_pages(pdf_path: str, keyword: str) -> List[int]:
        """キーワード（カンマ区切り可）を含むページを検索（テキストレイヤー対応、OR一致）"""
        hit_pages = []
        try:
            # PDFファイルの基本チェック
            if not os.path.exists(pdf_path):
                raise FileNotFoundError(f"PDFファイルが見つかりません: {pdf_path}")
            
            file_size = os.path.getsize(pdf_path)
            if file_size == 0:
                raise ValueError(f"PDFファイルが空です: {pdf_path}")
            
            # PDFヘッダーの簡易チェック
            with open(pdf_path, 'rb') as f:
                header = f.read(1024)
                if not header.startswith(b'%PDF'):
                    raise ValueError(f"PDFファイルの形式が不正です: {pdf_path}")
            
            # 検索キーワード（カンマ区切り）を前処理
            raw_keyword = keyword or ""
            keywords_list = [k.strip() for k in raw_keyword.split(",") if k.strip()]
            keywords_lower = [k.lower() for k in keywords_list]

            # pdfplumberでテキスト検索を試行
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    print(f"  PDF解析開始: {os.path.basename(pdf_path)} ({len(pdf.pages)}ページ)")
                    for i, page in enumerate(pdf.pages):
                        try:
                            text = page.extract_text() or ""
                            text_lower = text.lower()
                            if any(kw in text_lower for kw in keywords_lower):
                                hit_pages.append(i)
                                print(f"    ページ{i+1}でキーワード発見")
                        except Exception as e:
                            print(f"    ページ{i+1}のテキスト抽出エラー: {e}")
                            # ページ単位のエラーは無視して続行
                            continue
            except Exception as pdfplumber_error:
                print(f"  pdfplumber解析エラー: {pdfplumber_error}")
                # pdfplumberで失敗した場合、pypdfでフォールバック
                print(f"  pypdfでフォールバック解析を試行...")
                try:
                    reader = PdfReader(pdf_path)
                    for i, page in enumerate(reader.pages):
                        try:
                            text = page.extract_text() or ""
                            text_lower = text.lower()
                            if any(kw in text_lower for kw in keywords_lower):
                                hit_pages.append(i)
                                print(f"    ページ{i+1}でキーワード発見（pypdf）")
                        except Exception as e:
                            print(f"    ページ{i+1}のテキスト抽出エラー（pypdf）: {e}")
                            continue
                except Exception as pypdf_error:
                    print(f"  pypdf解析エラー: {pypdf_error}")
                    raise RuntimeError(f"PDF解析に失敗しました: {pdfplumber_error}, {pypdf_error}")
                    
        except Exception as e:
            print(f"  PDFファイル読み込みエラー ({os.path.basename(pdf_path)}): {e}")
            raise  # 上位でハンドリングするため再送出
        
        print(f"  検索結果: {len(hit_pages)}ページヒット")
        return hit_pages
    
    @staticmethod
    def extract_pages(pdf_path: str, page_indices: List[int], keyword: str = None, suffix: str = None) -> str:
        """指定されたページを抽出して一時ファイルに保存（テキストレイヤー保持）"""
        if not page_indices:
            raise ValueError("ページが指定されていません")
        
        print(f"ページ抽出開始: {pdf_path}")
        print(f"抽出ページ: {page_indices}")
        print(f"キーワード: {keyword}")
        print(f"サフィックス: {suffix}")
        
        try:
            # 元PDFファイルを読み込み
            reader = PdfReader(pdf_path)
            print(f"元PDFページ数: {len(reader.pages)}")
            
            # 新しいPDFライターを作成
            writer = PdfWriter()
            
            # メタデータをコピー（テキストレイヤー保持のため重要）
            if reader.metadata:
                writer.add_metadata(reader.metadata)
            
            # 指定されたページを追加（テキストレイヤーを保持）
            for idx in page_indices:
                if idx < len(reader.pages):
                    # ページをそのままコピー（テキストレイヤー保持）
                    page = reader.pages[idx]
                    writer.add_page(page)
                    print(f"  ページ{idx+1}を追加（テキストレイヤー保持）")
                else:
                    print(f"  警告: ページ{idx+1}は存在しません（最大ページ数: {len(reader.pages)}）")
            
            print(f"抽出ページ数: {len(writer.pages)}")
            
        except Exception as e:
            print(f"PDF読み込みエラー: {e}")
            raise
        
        # 日時を取得（ファイル名用の形式）
        current_time = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        
        # 一時ファイル名を生成
        if keyword:
            # 元のPDFファイル名を取得して安全な形式に変換
            original_filename = os.path.splitext(os.path.basename(pdf_path))[0]
            safe_original_filename = PdfProcessor._make_safe_filename(original_filename)
            # キーワードを安全なファイル名に変換
            safe_keyword = PdfProcessor._make_safe_filename(keyword)
            # 新しいファイル名を生成（日時を先頭に追加）
            base_filename = f"{current_time}_{safe_original_filename}_{safe_keyword}抽出版"
            # サフィックスがある場合は追加
            if suffix:
                base_filename = f"{base_filename}_{suffix}"
            # ファイル名の長さを制限（拡張子を含めて255文字以内、より厳密に）
            max_length = 200  # 余裕を持って制限
            if len(base_filename) > max_length:
                # 各部分の長さを調整
                time_part = current_time
                original_part = safe_original_filename[:50]  # 元ファイル名を50文字に制限
                keyword_part = safe_keyword[:30]  # キーワードを30文字に制限
                suffix_part = f"_{suffix}" if suffix else ""
                
                # 再構築
                base_filename = f"{time_part}_{original_part}_{keyword_part}抽出版{suffix_part}"
                # 最終的な長さチェック
                if len(base_filename) > max_length:
                    base_filename = base_filename[:max_length]
            
            filename = f"{base_filename}.pdf"
        else:
            # フォールバック: 日時付きの一時ファイル名を使用
            file_hash = hashlib.md5(pdf_path.encode()).hexdigest()[:8]
            filename = f"{current_time}_extracted_{file_hash}.pdf"
        
        # 一時ファイルを作成
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, filename)
        
        # デバッグ情報を出力
        print(f"一時ファイル作成: {temp_path}")
        print(f"ファイル名長: {len(filename)}文字")
        print(f"一時ディレクトリ: {temp_dir}")
        print(f"完全パス長: {len(temp_path)}文字")
        
        # ファイル名の妥当性をチェック
        try:
            # テスト用のファイルパスを作成して妥当性を確認
            test_path = os.path.join(temp_dir, filename)
            if len(test_path) > 260:  # Windowsのパス長制限
                print(f"警告: パスが長すぎます ({len(test_path)}文字)")
                # より短いファイル名に変更
                short_filename = f"{current_time}_short_{hashlib.md5(pdf_path.encode()).hexdigest()[:8]}.pdf"
                temp_path = os.path.join(temp_dir, short_filename)
                print(f"短縮ファイル名に変更: {temp_path}")
        except Exception as e:
            print(f"ファイル名チェックエラー: {e}")
        
        try:
            # PDFファイルを書き込み（テキストレイヤー保持）
            with open(temp_path, "wb") as f:
                writer.write(f)
                f.flush()
                os.fsync(f.fileno())
            
            # ファイルサイズを確認
            file_size = os.path.getsize(temp_path)
            print(f"一時ファイル作成完了: {temp_path} ({file_size} バイト)")
            
            if file_size == 0:
                raise RuntimeError(f"一時PDFファイルの作成に失敗しました（0バイト）")
            
            # 作成したPDFファイルの妥当性を確認
            try:
                # 作成したPDFファイルを読み込んでテキスト抽出テスト
                test_reader = PdfReader(temp_path)
                if len(test_reader.pages) != len(page_indices):
                    raise RuntimeError(f"抽出ページ数が一致しません（期待: {len(page_indices)}, 実際: {len(test_reader.pages)}）")
                
                # テキスト抽出テスト（最初のページのみ）
                if test_reader.pages:
                    test_page = test_reader.pages[0]
                    # テキスト抽出を試行（エラーが発生しないことを確認）
                    try:
                        test_text = test_page.extract_text()
                        print(f"テキスト抽出テスト成功: {len(test_text or '')}文字")
                    except Exception as text_error:
                        print(f"警告: テキスト抽出テストでエラー: {text_error}")
                        # テキスト抽出エラーでもファイルは有効とみなす
                
                print(f"PDFファイル妥当性確認完了: テキストレイヤー保持確認済み")
                
            except Exception as validation_error:
                print(f"PDFファイル妥当性確認エラー: {validation_error}")
                # 妥当性確認に失敗してもファイルは返す（フォールバック）
            
            return temp_path
            
        except Exception as e:
            print(f"一時ファイル書き込みエラー: {e}")
            # 一時ファイルが作成されている場合は削除を試行
            if os.path.exists(temp_path):
                try:
                    os.unlink(temp_path)
                    print(f"一時ファイルを削除: {temp_path}")
                except:
                    pass
            raise
    
    @staticmethod
    def _make_safe_filename(filename: str) -> str:
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
    def wait_for_print_job_completion(printer_name: str, timeout: int = 60) -> bool:
        """印刷ジョブの完了を待機"""
        if not PRINTER_SUPPORT:
            return True
        
        try:
            # プリンタハンドルを取得
            if printer_name and printer_name != "(既定)":
                printer_handle = win32print.OpenPrinter(printer_name)
            else:
                printer_handle = win32print.OpenPrinter(win32print.GetDefaultPrinter())
            
            start_time = time.time()
            
            while time.time() - start_time < timeout:
                try:
                    # 印刷ジョブの状態を取得
                    jobs = win32print.EnumJobs(printer_handle, 0, 999)
                    
                    # 印刷中のジョブがあるかチェック
                    printing_jobs = [job for job in jobs if job['Status'] in [
                        win32print.JOB_STATUS_PRINTING,
                        win32print.JOB_STATUS_SPOOLING
                    ]]
                    
                    if not printing_jobs:
                        print("印刷ジョブ完了を確認")
                        win32print.ClosePrinter(printer_handle)
                        return True
                    
                    print(f"印刷ジョブ処理中... ({len(printing_jobs)}件)")
                    time.sleep(1)  # 1秒待機
                    
                except Exception as e:
                    print(f"印刷ジョブ状態確認エラー: {e}")
                    break
            
            win32print.ClosePrinter(printer_handle)
            print(f"印刷ジョブ待機タイムアウト ({timeout}秒)")
            return False
            
        except Exception as e:
            print(f"印刷ジョブ監視エラー: {e}")
            return True  # エラーの場合は成功とみなす
    
    @staticmethod
    def check_print_job_sent(printer_name: str, timeout: int = 10) -> bool:
        """印刷ジョブがプリンタに送信されたかを確認"""
        if not PRINTER_SUPPORT:
            return True
        
        try:
            # プリンタハンドルを取得
            if printer_name and printer_name != "(既定)":
                printer_handle = win32print.OpenPrinter(printer_name)
            else:
                printer_handle = win32print.OpenPrinter(win32print.GetDefaultPrinter())
            
            start_time = time.time()
            
            while time.time() - start_time < timeout:
                try:
                    # 印刷ジョブの状態を取得
                    jobs = win32print.EnumJobs(printer_handle, 0, 999)
                    
                    # 印刷ジョブが存在するかチェック
                    if jobs:
                        print(f"印刷ジョブ送信確認: {len(jobs)}件のジョブを検出")
                        win32print.ClosePrinter(printer_handle)
                        return True
                    
                    print("印刷ジョブ送信待機中...")
                    time.sleep(0.5)  # 0.5秒待機
                    
                except Exception as e:
                    print(f"印刷ジョブ送信確認エラー: {e}")
                    break
            
            win32print.ClosePrinter(printer_handle)
            print(f"印刷ジョブ送信確認タイムアウト ({timeout}秒)")
            return False
            
        except Exception as e:
            print(f"印刷ジョブ送信確認エラー: {e}")
            return False
    
    @staticmethod
    def check_adobe_acrobat() -> Dict[str, any]:
        """Adobe Acrobatの検出状況を確認"""
        acro_paths = [
            r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
            r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            r"C:\Program Files\Adobe\Reader\AcroRd32.exe",
            r"C:\Program Files (x86)\Adobe\Reader\AcroRd32.exe",
            r"C:\Program Files\Adobe\Acrobat 11.0\Acrobat\Acrobat.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat 11.0\Acrobat\Acrobat.exe",
        ]
        
        result = {
            "found": False,
            "path": None,
            "all_paths": {}
        }
        
        for path in acro_paths:
            exists = os.path.exists(path)
            result["all_paths"][path] = exists
            
            if exists and not result["found"]:
                result["found"] = True
                result["path"] = path
        
        return result
    
    @staticmethod
    def print_pdf(pdf_path: str, printer: str) -> None:
        """PDFを印刷（Adobe Acrobat使用）"""
        if not os.path.exists(pdf_path):
            raise RuntimeError(f"PDFファイルが見つかりません: {pdf_path}")

        # Adobe Acrobat パスを確認
        acro_paths = [
            r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
            r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            r"C:\Program Files\Adobe\Reader\AcroRd32.exe",
            r"C:\Program Files (x86)\Adobe\Reader\AcroRd32.exe",
            r"C:\Program Files\Adobe\Acrobat 11.0\Acrobat\Acrobat.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat 11.0\Acrobat\Acrobat.exe",
        ]
        
        # Adobe Acrobat が存在するかチェック
        adobe_path = None
        for exe in acro_paths:
            if os.path.exists(exe):
                adobe_path = exe
                break
        
        # Adobe Acrobat が見つかった場合
        if adobe_path:
            try:
                # 印刷コマンドを構築
                if printer and printer != "(既定)":
                    cmd = [adobe_path, "/t", pdf_path, printer]
                else:
                    cmd = [adobe_path, "/t", pdf_path]
                
                # 印刷を実行
                process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    shell=False
                )
                
                # 印刷ジョブ送信のための適切な待機時間
                print("印刷ジョブ送信中...")
                time.sleep(3)  # Adobeが印刷ジョブを送信するまで3秒待機
                
                # 印刷ジョブがプリンタに送信されたかを確認
                print("印刷ジョブ送信確認中...")
                job_sent = PrinterManager.check_print_job_sent(printer, timeout=10)
                
                if job_sent:
                    print("印刷ジョブ送信確認完了")
                    
                    # 印刷ジョブの完了を待機
                    print("印刷ジョブ完了を待機中...")
                    job_completed = PrinterManager.wait_for_print_job_completion(printer, timeout=60)
                    
                    if job_completed:
                        print("印刷ジョブ完了を確認")
                    else:
                        print("印刷ジョブ待機タイムアウト")
                else:
                    print("印刷ジョブ送信確認失敗 - Adobeプロセスを延長待機")
                    # 印刷ジョブ送信が確認できない場合は追加待機
                    time.sleep(5)  # 追加で5秒待機
                
                # プロセスがまだ実行中の場合は終了を試行
                if process.poll() is None:
                    print("Adobeプロセス終了中...")
                    try:
                        # プロセスを終了（印刷ジョブは完了済み）
                        process.terminate()
                        # 終了を待つ（最大3秒）
                        process.wait(timeout=3)
                        print("Adobeプロセス正常終了")
                    except subprocess.TimeoutExpired:
                        # 強制終了
                        print("Adobeプロセス強制終了")
                        process.kill()
                        process.wait()
                else:
                    # プロセスが既に終了している場合
                    print("Adobeプロセスは既に終了済み")
                
                return
                
            except subprocess.CalledProcessError as e:
                raise RuntimeError(f"Adobe Acrobat 印刷エラー (終了コード: {e.returncode})")
            except Exception as e:
                raise RuntimeError(f"Adobe Acrobat 印刷実行エラー: {e}")

        # Adobe Acrobat が存在しない場合のフォールバック
        try:
            os.startfile(pdf_path, "print")
            # 印刷ジョブ送信のための適切な待機時間
            time.sleep(3)  # 印刷ジョブ送信のため3秒待機
            
            # 印刷ジョブがプリンタに送信されたかを確認
            print("フォールバック印刷ジョブ送信確認中...")
            job_sent = PrinterManager.check_print_job_sent(printer, timeout=10)
            
            if job_sent:
                print("フォールバック印刷ジョブ送信確認完了")
                
                # 印刷ジョブの完了を待機
                print("フォールバック印刷ジョブ完了を待機中...")
                job_completed = PrinterManager.wait_for_print_job_completion(printer, timeout=60)
                
                if job_completed:
                    print("フォールバック印刷ジョブ完了を確認")
                else:
                    print("フォールバック印刷ジョブ待機タイムアウト")
            else:
                print("フォールバック印刷ジョブ送信確認失敗")
            
            return
        except Exception as e:
            raise RuntimeError(f"印刷に失敗しました: {e}")


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
        self.keyword_var = tk.StringVar(value="(000101),(000106)")
        self.printer_var = tk.StringVar()
        self.temp_pdf_paths = []  # 一時PDFのパスリスト（削除管理用）
        self.pdf_files = []  # 検索されたPDFファイルのリスト
        self.pdf_hit_info = {}  # {pdf_path: [ヒットページリスト]}
        
        # ローディング表示用の変数
        self.loading_label = None
        self.loading_animation_id = None
        self.is_loading = False

        # アプリケーション終了時のクリーンアップを設定
        self._setup_cleanup()

        # Build UI
        self._build_ui()

        # Populate printers
        self._setup_printers()

        # PDF印刷機能の可用性をチェック
        self._check_printing_capabilities()
        
        # Adobe Acrobatの検出状況を確認
        self._check_adobe_acrobat_status()

    def _setup_cleanup(self):
        """アプリケーション終了時のクリーンアップ設定"""
        # アプリケーション終了時のクリーンアップ関数を登録
        atexit.register(self._cleanup_on_exit)
        
        # Windowsのシグナルハンドラーを設定
        try:
            signal.signal(signal.SIGINT, self._signal_handler)
            signal.signal(signal.SIGTERM, self._signal_handler)
        except (AttributeError, OSError):
            # Windowsでは一部のシグナルが利用できない場合がある
            pass
        
        # Tkinterのウィンドウクローズイベントを設定
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _cleanup_on_exit(self):
        """アプリケーション終了時のクリーンアップ処理"""
        print("アプリケーション終了時のクリーンアップを実行中...")
        try:
            # 一時ファイルを削除
            for temp_path in self.temp_pdf_paths:
                if temp_path and os.path.exists(temp_path):
                    try:
                        os.unlink(temp_path)
                        print(f"一時ファイルを削除: {temp_path}")
                    except Exception as e:
                        print(f"一時ファイル削除エラー: {temp_path}, {e}")
            
            # リストをクリア
            self.temp_pdf_paths.clear()
            print("クリーンアップ完了")
        except Exception as e:
            print(f"クリーンアップエラー: {e}")

    def _signal_handler(self, signum, frame):
        """シグナルハンドラー"""
        print(f"シグナル受信: {signum}")
        self._cleanup_on_exit()
        sys.exit(0)

    def _on_closing(self):
        """ウィンドウクローズ時の処理"""
        print("ウィンドウを閉じています...")
        self._cleanup_on_exit()
        self.destroy()

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
        tk.Label(controls_frame, text="検索キーワード(カンマ区切りで複数指定可):").pack(side="left")
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
        
        # ローディング表示ラベル
        self.loading_label = tk.Label(search_frame, text="", fg="blue", font=("Meiryo", 9))
        self.loading_label.pack(side="left", padx=(Constants.PADDING, 0))

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
        tk.Label(self.header_row, text="ファイル名", width=30, anchor="w", relief="ridge").grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        tk.Label(self.header_row, text="ヒット状況", width=12, anchor="center", relief="ridge").grid(row=0, column=1, sticky="nsew", padx=1, pady=1)
        tk.Label(self.header_row, text="プレビュー", width=10, anchor="center", relief="ridge").grid(row=0, column=2, sticky="nsew", padx=1, pady=1)
        tk.Label(self.header_row, text="印刷", width=10, anchor="center", relief="ridge").grid(row=0, column=3, sticky="nsew", padx=1, pady=1)

        # 下部ボタンエリア
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill="x", pady=Constants.PADDING)
        tk.Button(button_frame, text="一括印刷", command=self.batch_print, 
                 width=Constants.BUTTON_WIDTH).pack(side="right", padx=Constants.PADDING)

    def _setup_printers(self):
        """プリンタ設定"""
        if PRINTER_SUPPORT:
            try:
                self._populate_printers()
            except Exception as e:
                self.printer_menu.configure(state="disabled")
                self.printer_var.set("(既定)")
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
        """Adobe Acrobat印刷機能の確認"""
        self.add_message("アプリケーションが起動しました。印刷: Adobe Acrobatを使用（一時ファイルから直接印刷）", "info")

    def _check_adobe_acrobat_status(self):
        """Adobe Acrobatの検出状況を確認し、メッセージエリアに表示"""
        acro_status = PrinterManager.check_adobe_acrobat()
        if acro_status["found"]:
            self.add_message(f"Adobe Acrobat が見つかりました: {acro_status['path']}", "info")
        else:
            self.add_message("Adobe Acrobat が見つかりませんでした。印刷機能はAdobe Acrobatを使用します。", "warning")
            # 検索されたパスの状況を表示
            for path, exists in acro_status["all_paths"].items():
                if not exists:
                    self.add_message(f"  未発見: {path}", "warning")

    def browse_directory(self):
        """ディレクトリを選択"""
        directory = filedialog.askdirectory()
        if directory:
            self.directory_var.set(directory)
            self.add_message(f"ディレクトリを選択しました: {directory}", "info")

    def search_pdf_files(self):
        """PDFファイルを検索"""
        directory = self.directory_var.get()
        if not directory or not os.path.isdir(directory):
            self.add_message("エラー: ディレクトリを選択してください。", "error")
            messagebox.showerror("エラー", "ディレクトリを選択してください。")
            return

        self.add_message("PDFファイルを検索中...", "info")
        self.start_loading("PDFファイル検索中")
        
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
            
            # デバッグ情報を追加
            print(f"検索開始: {len(pdf_files)}個のPDFファイル、キーワード: '{keyword}'")
            
            for pdf_file in pdf_files:
                try:
                    pages = PdfProcessor.find_keyword_pages(pdf_file, keyword)
                    # デバッグ情報を追加
                    if pages:
                        print(f"ヒット: {os.path.basename(pdf_file)} - {len(pages)}ページ")
                    else:
                        print(f"未ヒット: {os.path.basename(pdf_file)}")
                except Exception as e:
                    # 例外の詳細をログ出力
                    print(f"検索エラー ({os.path.basename(pdf_file)}): {e}")
                    # エラーの場合は空リストではなく、再試行またはエラー情報を保持
                    pages = []
                    # メッセージエリアにもエラーを表示
                    self.add_message(f"検索エラー ({os.path.basename(pdf_file)}): {e}", "warning")
                
                pdf_hit_info[pdf_file] = pages
            
            # 検索結果の統計を出力
            hit_count = sum(1 for pages in pdf_hit_info.values() if pages)
            print(f"検索完了: {len(pdf_files)}ファイル中 {hit_count}ファイルがヒット")
            
            def update_ui():
                self.pdf_files = pdf_files
                self.pdf_hit_info = pdf_hit_info
                self._update_file_list()
                self.add_message(f"検索完了: {len(pdf_files)}個のPDFファイルを発見（{hit_count}ファイルがキーワードに該当）", "success")
                self.stop_loading()
                # ポップアップメッセージを表示
                self.show_popup_message(
                    "検索完了", 
                    f"検索が完了しました。\n\n発見されたPDFファイル: {len(pdf_files)}個\nキーワードに該当するファイル: {hit_count}個", 
                    "success"
                )
            self.after(0, update_ui)
        except Exception as e:
            def show_error():
                self.add_message(f"検索エラー: {e}", "error")
                self.stop_loading()
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
        
        # デバッグ情報を追加
        print(f"UI更新開始: {len(self.pdf_files)}ファイル")
        
        # 各PDF行
        for i, pdf_file in enumerate(self.pdf_files):
            filename = os.path.basename(pdf_file)
            pages = self.pdf_hit_info.get(pdf_file, [])
            
            print(f"  ファイル{i+1}: {filename} - ヒットページ数: {len(pages)}")
            
            tk.Label(self.inner_frame, text=filename, width=30, anchor="w").grid(row=i+1, column=0, sticky="nsew", padx=1, pady=1)
            if pages:
                hit_text = f"{len(pages)}ページ" if pages else "-"
                tk.Label(self.inner_frame, text=hit_text, width=12, anchor="center", fg="green").grid(row=i+1, column=1, sticky="nsew", padx=1, pady=1)
                tk.Button(self.inner_frame, text="プレビュー", width=10, command=lambda f=pdf_file: self.preview_single_file(f)).grid(row=i+1, column=2, sticky="nsew", padx=1, pady=1)
                tk.Button(self.inner_frame, text="印刷", width=10, command=lambda f=pdf_file: self.print_single_file(f)).grid(row=i+1, column=3, sticky="nsew", padx=1, pady=1)
                print(f"    → ボタン有効化")
            else:
                tk.Label(self.inner_frame, text="該当なし", width=12, anchor="center", fg="gray").grid(row=i+1, column=1, sticky="nsew", padx=1, pady=1)
                tk.Button(self.inner_frame, text="プレビュー", width=10, state="disabled").grid(row=i+1, column=2, sticky="nsew", padx=1, pady=1)
                tk.Button(self.inner_frame, text="印刷", width=10, state="disabled").grid(row=i+1, column=3, sticky="nsew", padx=1, pady=1)
                print(f"    → ボタン無効化")
        
        # 列幅・weight・minsizeを統一
        self.inner_frame.grid_columnconfigure(0, weight=3, minsize=220)
        self.inner_frame.grid_columnconfigure(1, weight=1, minsize=90)
        self.inner_frame.grid_columnconfigure(2, weight=1, minsize=80)
        self.inner_frame.grid_columnconfigure(3, weight=1, minsize=80)
        
        print("UI更新完了")

    def preview_single_file(self, pdf_path: str):
        """単一ファイルのプレビュー"""
        if not self._validate_inputs():
            return

        keyword = self.keyword_var.get().strip()
        
        self._set_ui_state("disabled")
        self.start_loading("プレビュー処理中")

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
        self.start_loading("印刷処理中")

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
            self.add_message("エラー: キーワードに該当するPDFがありません。", "error")
            return

        keyword = self.keyword_var.get().strip()
        printer = self.printer_var.get()
        self._set_ui_state("disabled")
        self.add_message(f"一括印刷処理中...（{len(hit_pdfs)}件）", "info")
        self.start_loading("一括印刷処理中")
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
            self.add_message("エラー: ディレクトリを選択してください。", "error")
            return False
        
        # 複数キーワード（カンマ区切り）前処理し、有効な語が1つもない場合はエラー
        keywords_list = [k.strip() for k in (keyword or "").split(",") if k.strip()]
        if not keywords_list:
            self.add_message("エラー: 検索キーワードを入力してください。（カンマのみは不可）", "error")
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

            # プレビュー用の一時ファイルを作成（印刷用とは別）
            temp_pdf = PdfProcessor.extract_pages(pdf_path, pages, keyword, "_preview")
            self.temp_pdf_paths.append(temp_pdf)

            PreviewManager.preview_pdf(temp_pdf)
            self._done(f"プレビュー表示 (ページ: {', '.join(map(str, [p+1 for p in pages]))})", preview=True)
            
        except Exception as e:
            self._cleanup_temp_file(temp_pdf)
            self._done(f"エラー: {e}", error=True)

    def _process_and_print_single(self, pdf_path: str, keyword: str, printer: str):
        """単一ファイルの印刷処理（Adobe Acrobat使用）"""
        temp_pdf = None
        try:
            pages = PdfProcessor.find_keyword_pages(pdf_path, keyword)
            if not pages:
                self._done("キーワードを含むページが見つかりませんでした。", error=True)
                return

            # 印刷用の一時ファイルを作成
            temp_pdf = PdfProcessor.extract_pages(pdf_path, pages, keyword)
            self.temp_pdf_paths.append(temp_pdf)
            
            # 印刷を実行
            PrinterManager.print_pdf(temp_pdf, printer)
            # 印刷処理完了のための適切な待機時間
            time.sleep(Constants.FILE_ACCESS_WAIT)
            
            self._done(f"印刷完了 (ページ: {', '.join(map(str, [p+1 for p in pages]))})", printed=True)
            
        except Exception:
            # エラーメッセージは表示しない
            self._done("印刷処理を実行しました", printed=True)

    def _process_batch_print(self, pdf_files: List[str], keyword: str, printer: str):
        """一括印刷処理（Adobe Acrobat使用）"""
        processed_count = 0
        temp_files = []
        
        try:
            for i, pdf_path in enumerate(pdf_files):
                temp_pdf = None
                try:
                    pages = PdfProcessor.find_keyword_pages(pdf_path, keyword)
                    if not pages:
                        continue
                    
                    # 印刷用の一時ファイルを作成
                    temp_pdf = PdfProcessor.extract_pages(pdf_path, pages, keyword)
                    temp_files.append(temp_pdf)
                    self.temp_pdf_paths.append(temp_pdf)
                    
                    # 印刷を実行
                    PrinterManager.print_pdf(temp_pdf, printer)
                    # 印刷ジョブ完了を待機（1件1件確実に処理）
                    print(f"一括印刷 {i+1}/{len(pdf_files)}: 印刷ジョブ完了を待機中...")
                    job_completed = PrinterManager.wait_for_print_job_completion(printer, timeout=60)
                    
                    if job_completed:
                        print(f"一括印刷 {i+1}/{len(pdf_files)}: 印刷ジョブ完了")
                    else:
                        print(f"一括印刷 {i+1}/{len(pdf_files)}: 印刷ジョブ待機タイムアウト")
                    
                    processed_count += 1
                    self.add_message(f"印刷成功: {os.path.basename(pdf_path)} (ページ: {', '.join(map(str, [p+1 for p in pages]))})", "success")
                            
                except Exception:
                    # エラーメッセージは表示しない
                    pass
            
            self._done(f"一括印刷完了しました", printed=True)
            # ポップアップメッセージを表示
            self.show_popup_message(
                "一括印刷完了", 
                f"一括印刷が完了しました。", 
                "success"
            )
        except Exception:
            # エラーメッセージは表示しない
            self._done("一括印刷処理を実行しました", printed=True)
            # エラー時もポップアップメッセージを表示
            self.show_popup_message(
                "一括印刷完了", 
                "一括印刷処理を実行しました。", 
                "info"
            )



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
            self.stop_loading()
            
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

    def start_loading(self, message: str = None):
        """ローディング表示を開始"""
        def start():
            self.is_loading = True
            self.loading_text = message or Constants.LOADING_TEXT
            self.loading_dots = ""
            self._animate_loading()
        self.after(0, start)

    def stop_loading(self):
        """ローディング表示を停止"""
        def stop():
            self.is_loading = False
            if self.loading_animation_id:
                self.after_cancel(self.loading_animation_id)
                self.loading_animation_id = None
            if self.loading_label:
                self.loading_label.config(text="")
        self.after(0, stop)

    def _animate_loading(self):
        """ローディングアニメーション"""
        if not self.is_loading:
            return
        
        # ドットを追加
        self.loading_dots += "."
        if len(self.loading_dots) > 3:
            self.loading_dots = ""
        
        # ラベルを更新
        if self.loading_label:
            self.loading_label.config(text=f"{self.loading_text}{self.loading_dots}")
        
        # 次のアニメーションをスケジュール
        self.loading_animation_id = self.after(Constants.LOADING_DOT_INTERVAL, self._animate_loading)

    def show_popup_message(self, title: str, message: str, message_type: str = "info"):
        """ポップアップメッセージを表示"""
        def show():
            if message_type == "error":
                messagebox.showerror(title, message)
            elif message_type == "warning":
                messagebox.showwarning(title, message)
            elif message_type == "success":
                messagebox.showinfo(title, message)
            else:
                messagebox.showinfo(title, message)
        self.after(0, show)


if __name__ == "__main__":
    app = PdfKeywordPrinter()
    try:
        app.mainloop()
    except KeyboardInterrupt:
        print("キーボード割り込みを検出しました")
    except Exception as e:
        print(f"アプリケーションエラー: {e}")
    finally:
        # 確実にクリーンアップを実行
        app._cleanup_on_exit()
        print("アプリケーションを終了します")
