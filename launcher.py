"""書店業務自動化システム - GUI ランチャー

デスクトップGUI画面からシステムの起動・停止・管理画面の表示を行う。
tkinterを使用（Python標準ライブラリ、追加インストール不要）。
"""

import os
import sys
import threading
import time
import webbrowser
import subprocess
from datetime import datetime

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext


# ── 定数 ──────────────────────────────────────────
APP_TITLE = "書店業務自動化システム"
APP_VERSION = "1.0.0"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 日本語表示可能なフォント（プラットフォーム別）
# Windows: Yu Gothic UI / Linux: Noto Sans CJK or DejaVu Sans / macOS: Hiragino
if sys.platform == "win32":
    UI_FONT = UI_FONT
elif sys.platform == "darwin":
    UI_FONT = "Hiragino Sans"
else:
    UI_FONT = "Noto Sans CJK JP"


def _open_in_file_manager(path: str) -> None:
    """OSのファイルマネージャでフォルダを開く（クロスプラットフォーム）。"""
    if sys.platform == "win32":
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])

# 色定義
COLOR_BG = "#1a1a2e"
COLOR_BG_LIGHT = "#16213e"
COLOR_ACCENT = "#0f3460"
COLOR_PRIMARY = "#e94560"
COLOR_SUCCESS = "#00b894"
COLOR_WARNING = "#fdcb6e"
COLOR_TEXT = "#eaeaea"
COLOR_TEXT_DIM = "#a0a0a0"
COLOR_CARD = "#1e2a47"
COLOR_CARD_BORDER = "#2d3f5f"


class BookAutoLauncher:
    """メインGUIランチャークラス"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry("680x720")
        self.root.minsize(600, 650)
        self.root.configure(bg=COLOR_BG)
        self.root.resizable(True, True)

        # アイコン設定（Windows のみ .ico に対応; Linux では .png を試す）
        if sys.platform == "win32":
            icon_path = os.path.join(BASE_DIR, "assets", "icon.ico")
            if os.path.exists(icon_path):
                try:
                    self.root.iconbitmap(icon_path)
                except tk.TclError:
                    pass
        else:
            icon_png = os.path.join(BASE_DIR, "assets", "icon.png")
            if os.path.exists(icon_png):
                try:
                    self.root.iconphoto(True, tk.PhotoImage(file=icon_png))
                except tk.TclError:
                    pass

        # 状態管理
        self.pipeline_running = False
        self.pipeline_thread = None
        self.web_thread = None
        self.web_running = False
        self.last_run_time = None
        self.cycle_count = 0
        self.stop_event = threading.Event()

        # UI構築
        self._build_ui()

        # 定期更新タイマー
        self._update_status_display()

        # ウィンドウ閉じる処理
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        """UI全体を構築する"""
        # メインフレーム
        main_frame = tk.Frame(self.root, bg=COLOR_BG, padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ── ヘッダー ──
        self._build_header(main_frame)

        # ── ステータスカード ──
        self._build_status_card(main_frame)

        # ── コントロールボタン ──
        self._build_controls(main_frame)

        # ── 統計情報 ──
        self._build_stats(main_frame)

        # ── ログ表示 ──
        self._build_log_area(main_frame)

        # ── フッター ──
        self._build_footer(main_frame)

    def _build_header(self, parent):
        """ヘッダー部分"""
        header = tk.Frame(parent, bg=COLOR_BG)
        header.pack(fill=tk.X, pady=(0, 15))

        title_label = tk.Label(
            header,
            text=APP_TITLE,
            font=(UI_FONT, 20, "bold"),
            fg=COLOR_TEXT,
            bg=COLOR_BG,
        )
        title_label.pack(side=tk.LEFT)

        version_label = tk.Label(
            header,
            text=f"v{APP_VERSION}",
            font=(UI_FONT, 10),
            fg=COLOR_TEXT_DIM,
            bg=COLOR_BG,
        )
        version_label.pack(side=tk.LEFT, padx=(10, 0), pady=(8, 0))

    def _build_status_card(self, parent):
        """ステータス表示カード"""
        card = tk.Frame(parent, bg=COLOR_CARD, highlightbackground=COLOR_CARD_BORDER,
                        highlightthickness=1, padx=20, pady=15)
        card.pack(fill=tk.X, pady=(0, 15))

        # ステータスインジケーター
        status_row = tk.Frame(card, bg=COLOR_CARD)
        status_row.pack(fill=tk.X)

        self.status_indicator = tk.Label(
            status_row,
            text="\u25cf",
            font=(UI_FONT, 16),
            fg=COLOR_TEXT_DIM,
            bg=COLOR_CARD,
        )
        self.status_indicator.pack(side=tk.LEFT)

        self.status_label = tk.Label(
            status_row,
            text="  \u505c\u6b62\u4e2d",
            font=(UI_FONT, 14, "bold"),
            fg=COLOR_TEXT,
            bg=COLOR_CARD,
        )
        self.status_label.pack(side=tk.LEFT)

        # 詳細情報
        detail_frame = tk.Frame(card, bg=COLOR_CARD)
        detail_frame.pack(fill=tk.X, pady=(10, 0))

        self.detail_labels = {}
        details = [
            ("mode", "\u30e2\u30fc\u30c9:", "\u2500"),
            ("last_run", "\u6700\u7d42\u5b9f\u884c:", "\u2500"),
            ("cycle", "\u5b9f\u884c\u56de\u6570:", "0"),
            ("interval", "\u9593\u9694:", "\u2500"),
        ]

        for key, label_text, default_value in details:
            row = tk.Frame(detail_frame, bg=COLOR_CARD)
            row.pack(fill=tk.X, pady=1)

            tk.Label(row, text=label_text, font=(UI_FONT, 10),
                     fg=COLOR_TEXT_DIM, bg=COLOR_CARD, width=12, anchor="w"
                     ).pack(side=tk.LEFT)

            val_label = tk.Label(row, text=default_value, font=(UI_FONT, 10),
                                fg=COLOR_TEXT, bg=COLOR_CARD, anchor="w")
            val_label.pack(side=tk.LEFT)
            self.detail_labels[key] = val_label

    def _build_controls(self, parent):
        """操作ボタン群"""
        controls = tk.Frame(parent, bg=COLOR_BG)
        controls.pack(fill=tk.X, pady=(0, 15))

        # 上段：メインコントロール
        main_btns = tk.Frame(controls, bg=COLOR_BG)
        main_btns.pack(fill=tk.X, pady=(0, 8))

        self.start_btn = self._make_button(
            main_btns, "\u25b6  \u30b7\u30b9\u30c6\u30e0\u958b\u59cb", self._start_system,
            bg=COLOR_SUCCESS, width=20
        )
        self.start_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.stop_btn = self._make_button(
            main_btns, "\u25a0  \u30b7\u30b9\u30c6\u30e0\u505c\u6b62", self._stop_system,
            bg=COLOR_PRIMARY, width=20, state=tk.DISABLED
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.once_btn = self._make_button(
            main_btns, "\u25b7  1\u56de\u5b9f\u884c", self._run_once,
            bg=COLOR_ACCENT, width=14
        )
        self.once_btn.pack(side=tk.LEFT)

        # 下段：ツールボタン
        tool_btns = tk.Frame(controls, bg=COLOR_BG)
        tool_btns.pack(fill=tk.X)

        self.web_btn = self._make_button(
            tool_btns, "\U0001f310  \u7ba1\u7406\u753b\u9762\u3092\u958b\u304f", self._open_web,
            bg=COLOR_ACCENT, width=20
        )
        self.web_btn.pack(side=tk.LEFT, padx=(0, 8))

        self._make_button(
            tool_btns, "\U0001f4c2  Excel\u51fa\u529b", self._export_excel,
            bg=COLOR_ACCENT, width=14
        ).pack(side=tk.LEFT, padx=(0, 8))

        self._make_button(
            tool_btns, "\u2699  \u8a2d\u5b9a\u30d5\u30a9\u30eb\u30c0", self._open_config,
            bg=COLOR_ACCENT, width=14
        ).pack(side=tk.LEFT)

    def _build_stats(self, parent):
        """統計情報パネル"""
        stats_frame = tk.Frame(parent, bg=COLOR_BG)
        stats_frame.pack(fill=tk.X, pady=(0, 10))

        self.stat_labels = {}
        stat_items = [
            ("total", "\u5168\u4ef6\u6570", "#74b9ff"),
            ("pending", "PENDING", COLOR_TEXT_DIM),
            ("hold", "HOLD", COLOR_WARNING),
            ("ordered", "ORDERED", "#74b9ff"),
            ("error", "ERROR", COLOR_PRIMARY),
        ]

        for key, label_text, color in stat_items:
            card = tk.Frame(stats_frame, bg=COLOR_CARD, highlightbackground=COLOR_CARD_BORDER,
                            highlightthickness=1, padx=12, pady=8)
            card.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

            val = tk.Label(card, text="0", font=(UI_FONT, 18, "bold"),
                           fg=color, bg=COLOR_CARD)
            val.pack()
            tk.Label(card, text=label_text, font=(UI_FONT, 9),
                     fg=COLOR_TEXT_DIM, bg=COLOR_CARD).pack()
            self.stat_labels[key] = val

    def _build_log_area(self, parent):
        """ログ表示エリア"""
        log_frame = tk.Frame(parent, bg=COLOR_BG)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        log_header = tk.Frame(log_frame, bg=COLOR_BG)
        log_header.pack(fill=tk.X)

        tk.Label(log_header, text="\u30b7\u30b9\u30c6\u30e0\u30ed\u30b0",
                 font=(UI_FONT, 11, "bold"), fg=COLOR_TEXT, bg=COLOR_BG
                 ).pack(side=tk.LEFT)

        self._make_button(
            log_header, "\u30af\u30ea\u30a2", self._clear_log,
            bg=COLOR_ACCENT, width=8, font_size=9
        ).pack(side=tk.RIGHT)

        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg="#0d1117",
            fg="#c9d1d9",
            insertbackground=COLOR_TEXT,
            selectbackground=COLOR_ACCENT,
            height=10,
            borderwidth=1,
            relief=tk.SOLID,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        # ログタグ色設定
        self.log_text.tag_config("INFO", foreground="#58a6ff")
        self.log_text.tag_config("WARNING", foreground="#d29922")
        self.log_text.tag_config("ERROR", foreground="#f85149")
        self.log_text.tag_config("SUCCESS", foreground="#3fb950")

    def _build_footer(self, parent):
        """フッター"""
        footer = tk.Frame(parent, bg=COLOR_BG)
        footer.pack(fill=tk.X)

        tk.Label(footer, text=f"{APP_TITLE} v{APP_VERSION}",
                 font=(UI_FONT, 9), fg=COLOR_TEXT_DIM, bg=COLOR_BG
                 ).pack(side=tk.LEFT)

    def _make_button(self, parent, text, command, bg=COLOR_ACCENT,
                     width=12, state=tk.NORMAL, font_size=10):
        """統一スタイルのボタンを作成する"""
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            font=(UI_FONT, font_size),
            fg="white",
            bg=bg,
            activebackground=bg,
            activeforeground="white",
            relief=tk.FLAT,
            cursor="hand2",
            width=width,
            state=state,
            padx=8,
            pady=4,
        )
        return btn

    # ── アクション ──────────────────────────────────

    def _start_system(self):
        """システムを開始する（メインループ）"""
        if self.pipeline_running:
            return

        self.stop_event.clear()
        self.pipeline_running = True
        self._update_button_states()
        self._set_status("running")
        self._log("システムを開始しました", "SUCCESS")

        # 管理画面を自動起動
        if not self.web_running:
            self._start_web_server()

        # パイプラインスレッド開始
        self.pipeline_thread = threading.Thread(target=self._pipeline_loop, daemon=True)
        self.pipeline_thread.start()

    def _stop_system(self):
        """システムを停止する"""
        if not self.pipeline_running:
            return

        self._log("システムを停止しています...", "WARNING")
        self.stop_event.set()
        self.pipeline_running = False
        self._update_button_states()
        self._set_status("stopped")
        self._log("システムを停止しました", "SUCCESS")

    def _run_once(self):
        """パイプラインを1回だけ実行する"""
        if self.pipeline_running:
            self._log("メインループ稼働中のため、1回実行はスキップされました", "WARNING")
            return

        self.once_btn.config(state=tk.DISABLED)
        self._set_status("running_once")
        self._log("パイプラインを1回実行します...", "INFO")

        def run():
            try:
                self._init_system()
                results = self._execute_pipeline()
                self._log_results(results)
            except Exception as e:
                self._log(f"エラー: {e}", "ERROR")
            finally:
                self.root.after(0, lambda: self.once_btn.config(state=tk.NORMAL))
                self.root.after(0, lambda: self._set_status("stopped"))

        threading.Thread(target=run, daemon=True).start()

    def _open_web(self):
        """管理画面をブラウザで開く"""
        if not self.web_running:
            self._start_web_server()
            time.sleep(1.5)

        host = self._get_setting("web", "host", "127.0.0.1")
        port = self._get_setting("web", "port", 5000)
        # ブラウザではlocalhostでアクセス
        url = f"http://127.0.0.1:{port}"
        webbrowser.open(url)
        self._log(f"管理画面を開きました: {url}", "INFO")

    def _export_excel(self):
        """Excel台帳を出力する"""
        self._log("Excel台帳を出力しています...", "INFO")

        def run():
            try:
                self._init_system()
                from src.export.excel_exporter import export_all
                p, s = export_all()
                self._log(f"処理台帳: {p}", "SUCCESS")
                self._log(f"仕入台帳: {s}", "SUCCESS")

                # 出力フォルダを開く
                output_dir = os.path.join(BASE_DIR, "data", "output")
                if os.path.exists(output_dir):
                    _open_in_file_manager(output_dir)

            except Exception as e:
                self._log(f"Excel出力エラー: {e}", "ERROR")

        threading.Thread(target=run, daemon=True).start()

    def _open_config(self):
        """設定フォルダをエクスプローラーで開く"""
        config_dir = os.path.join(BASE_DIR, "config")
        if os.path.exists(config_dir):
            _open_in_file_manager(config_dir)
        else:
            self._log("設定フォルダが見つかりません", "ERROR")

    def _clear_log(self):
        """ログ表示をクリアする"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

    # ── 内部処理 ──────────────────────────────────

    def _init_system(self):
        """システムの初期化（設定読み込み・DB初期化）"""
        # 作業ディレクトリをプロジェクトルートに変更
        os.chdir(BASE_DIR)

        from src.common.logger import setup_logging
        from src.common.config import load_settings, get_setting
        from src.common.database import init_db, sync_suppliers_from_config
        from src.common.config import load_suppliers

        load_settings()
        setup_logging(
            log_dir=get_setting("logging", "dir", default="logs"),
            level=get_setting("logging", "level", default="INFO"),
        )
        init_db()
        suppliers = load_suppliers()
        sync_suppliers_from_config(suppliers)

    def _execute_pipeline(self) -> dict:
        """パイプラインを1回実行する"""
        from main import run_pipeline
        results = run_pipeline()
        self.last_run_time = datetime.now()
        self.cycle_count += 1
        return results

    def _pipeline_loop(self):
        """メインパイプラインループ"""
        try:
            self._init_system()
        except Exception as e:
            self._log(f"初期化エラー: {e}", "ERROR")
            self.pipeline_running = False
            self.root.after(0, lambda: self._update_button_states())
            self.root.after(0, lambda: self._set_status("stopped"))
            return

        interval = int(self._get_setting("outlook", "polling_interval_sec", default=300) or 300)
        self.root.after(0, lambda: self.detail_labels["interval"].config(
            text=f"{interval}秒"))
        self.root.after(0, lambda: self.detail_labels["mode"].config(
            text="自動ループ"))

        while not self.stop_event.is_set():
            try:
                self._log(f"--- パイプライン実行開始 (cycle={self.cycle_count + 1}) ---", "INFO")
                results = self._execute_pipeline()
                self._log_results(results)
                self._refresh_stats()
            except Exception as e:
                self._log(f"パイプラインエラー: {e}", "ERROR")

            # 停止チェック付きの待機
            for _ in range(interval):
                if self.stop_event.is_set():
                    break
                time.sleep(1)

    def _start_web_server(self):
        """管理画面Webサーバーをバックグラウンドで起動する"""
        if self.web_running:
            return

        def run_web():
            try:
                os.chdir(BASE_DIR)
                from src.common.config import load_settings, get_setting
                from src.common.database import init_db
                load_settings()
                init_db()

                from src.web.app import create_app
                app = create_app()
                host = get_setting("web", "host", default="0.0.0.0")
                port = get_setting("web", "port", default=5000)
                self.web_running = True
                self._log(f"管理画面サーバー起動: http://127.0.0.1:{port}", "SUCCESS")
                app.run(host=host, port=port, debug=False, use_reloader=False)
            except Exception as e:
                self._log(f"管理画面起動エラー: {e}", "ERROR")
                self.web_running = False

        self.web_thread = threading.Thread(target=run_web, daemon=True)
        self.web_thread.start()

    def _get_setting(self, *keys, default=None):
        """設定値を安全に取得する"""
        try:
            from src.common.config import get_setting
            result = get_setting(*keys, default=default)
            return result if result is not None else default
        except Exception:
            return default

    def _log_results(self, results: dict):
        """パイプライン結果をログ表示する"""
        errors = results.get("errors", [])
        error_count = len(errors)

        summary = (
            f"メール={results.get('mails_fetched', 0)}, "
            f"解析={results.get('items_parsed', 0)}, "
            f"スクレイピング={results.get('items_scraped', 0)}, "
            f"振分={results.get('items_assigned', 0)}, "
            f"送信={results.get('mails_sent', 0)}"
        )

        if error_count == 0:
            self._log(f"パイプライン完了: {summary}", "SUCCESS")
        else:
            self._log(f"パイプライン完了(警告{error_count}件): {summary}", "WARNING")
            for err in errors:
                self._log(f"  {err}", "WARNING")

    def _refresh_stats(self):
        """DB統計を更新する"""
        try:
            from src.common.database import get_connection
            with get_connection() as conn:
                stats = conn.execute("""
                    SELECT
                        COUNT(*) AS total,
                        SUM(CASE WHEN current_status='PENDING' THEN 1 ELSE 0 END) AS pending,
                        SUM(CASE WHEN current_status='HOLD' THEN 1 ELSE 0 END) AS hold,
                        SUM(CASE WHEN current_status='ORDERED' THEN 1 ELSE 0 END) AS ordered,
                        SUM(CASE WHEN current_status='ERROR' THEN 1 ELSE 0 END) AS error
                    FROM order_items
                """).fetchone()

            for key in ["total", "pending", "hold", "ordered", "error"]:
                val = stats[key] or 0
                self.root.after(0, lambda k=key, v=val: self.stat_labels[k].config(text=str(v)))

        except Exception:
            pass

    # ── UI更新 ──────────────────────────────────

    def _set_status(self, status: str):
        """ステータス表示を更新する"""
        status_map = {
            "running": ("\u25cf", COLOR_SUCCESS, "  \u7a3c\u50cd\u4e2d"),
            "running_once": ("\u25cf", COLOR_WARNING, "  \u5b9f\u884c\u4e2d\uff081\u56de\uff09"),
            "stopped": ("\u25cf", COLOR_TEXT_DIM, "  \u505c\u6b62\u4e2d"),
        }
        indicator, color, text = status_map.get(status, status_map["stopped"])
        self.status_indicator.config(text=indicator, fg=color)
        self.status_label.config(text=text)

    def _update_button_states(self):
        """ボタンの有効/無効を更新する"""
        if self.pipeline_running:
            self.start_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)
            self.once_btn.config(state=tk.DISABLED)
        else:
            self.start_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self.once_btn.config(state=tk.NORMAL)

    def _update_status_display(self):
        """定期的にステータス表示を更新する（1秒ごと）"""
        # 最終実行時間
        if self.last_run_time:
            self.detail_labels["last_run"].config(
                text=self.last_run_time.strftime("%Y-%m-%d %H:%M:%S"))

        # 実行回数
        self.detail_labels["cycle"].config(text=str(self.cycle_count))

        # 次回更新
        self.root.after(1000, self._update_status_display)

    def _log(self, message: str, level: str = "INFO"):
        """ログメッセージを表示する"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted = f"[{timestamp}] [{level}] {message}\n"

        def append():
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, formatted, level)
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)

        self.root.after(0, append)

    def _on_close(self):
        """ウィンドウを閉じる処理"""
        if self.pipeline_running:
            result = messagebox.askokcancel(
                "確認",
                "システムが稼働中です。終了してもよろしいですか？",
                parent=self.root
            )
            if not result:
                return

        self.stop_event.set()
        self.pipeline_running = False
        self.root.destroy()

    def run(self):
        """GUIを起動する"""
        self._log("システム起動準備完了", "SUCCESS")
        self._log("「システム開始」ボタンで自動処理を開始します", "INFO")

        # 初期統計を読み込む
        try:
            os.chdir(BASE_DIR)
            from src.common.config import load_settings
            from src.common.database import init_db
            load_settings()
            init_db()
            self._refresh_stats()
        except Exception:
            pass

        self.root.mainloop()


def main():
    app = BookAutoLauncher()
    app.run()


if __name__ == "__main__":
    main()
