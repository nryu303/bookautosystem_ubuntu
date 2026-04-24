"""ログ機構 - ファイル+コンソール出力、日付ローテーション、DBログ"""

import logging
import os
from datetime import datetime
from logging.handlers import TimedRotatingFileHandler


_initialized = False


class DatabaseLogHandler(logging.Handler):
    """WARNING以上のログをlogsテーブルに書き込むハンドラー。

    DB初期化前に呼ばれても安全にスキップする。
    """

    def emit(self, record: logging.LogRecord) -> None:
        try:
            from src.common.database import get_connection
            with get_connection() as conn:
                conn.execute(
                    "INSERT INTO logs (level, module, message) VALUES (?, ?, ?)",
                    (record.levelname, record.name, self.format(record)),
                )
        except Exception:
            pass  # DB未初期化やインポート循環を安全にスキップ


def setup_logging(log_dir: str = "logs", level: str = "INFO") -> None:
    """ログ設定を初期化する。アプリ起動時に1回だけ呼ぶ。"""
    global _initialized
    if _initialized:
        return

    os.makedirs(log_dir, exist_ok=True)

    log_level = getattr(logging, level.upper(), logging.INFO)
    root_logger = logging.getLogger()
    root_logger.setLevel(log_level)

    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # コンソール出力
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)

    # ファイル出力（日付ローテーション）
    log_file = os.path.join(log_dir, "app.log")
    file_handler = TimedRotatingFileHandler(
        filename=log_file,
        when="midnight",
        interval=1,
        backupCount=30,
        encoding="utf-8",
    )
    file_handler.setLevel(log_level)
    file_handler.setFormatter(formatter)
    file_handler.suffix = "%Y-%m-%d"
    root_logger.addHandler(file_handler)

    # DBログ（WARNING以上をlogsテーブルに保存）
    db_handler = DatabaseLogHandler()
    db_handler.setLevel(logging.WARNING)
    db_handler.setFormatter(logging.Formatter("%(message)s"))
    root_logger.addHandler(db_handler)

    _initialized = True


def get_logger(name: str) -> logging.Logger:
    """モジュール用ロガーを取得する。各モジュールの先頭で呼ぶ。"""
    return logging.getLogger(name)
