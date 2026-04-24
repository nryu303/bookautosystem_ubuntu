"""スクレイピング処理の多重起動防止ロック

Playwrightドライバは同時に複数起動するとChromium接続エラー
（Connection closed while reading from the driver）を起こすため、
全スクレイピングはこのロックの下で1プロセスのみ実行する。
"""

import os
import time
from contextlib import contextmanager

from src.common.config import BASE_DIR
from src.common.logger import get_logger

logger = get_logger(__name__)

LOCK_PATH = os.path.join(BASE_DIR, "data", ".scraping.lock")
MAIN_LOCK_PATH = os.path.join(BASE_DIR, "data", ".main.lock")


class ScrapingLockBusy(Exception):
    """他プロセスがロック保持中のため取得できなかった。"""


class MainInstanceBusy(Exception):
    """main.py が既に別プロセスで起動中。"""


@contextmanager
def scraping_lock(*, blocking: bool = False, poll_sec: float = 5.0, stale_after_sec: int = 3600):
    """ロックファイルを確保してスクレイピングを実行するコンテキストマネージャ。

    Args:
        blocking: True で他プロセス解放待ち、False で取得不可なら即 ScrapingLockBusy
        poll_sec: blocking時のポーリング間隔秒
        stale_after_sec: この秒数以上経過したロックはクラッシュ痕とみなし奪取
    """
    os.makedirs(os.path.dirname(LOCK_PATH), exist_ok=True)

    while True:
        try:
            fd = os.open(LOCK_PATH, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            try:
                os.write(fd, f"{os.getpid()}\n{time.time()}\n".encode())
            finally:
                os.close(fd)
            break
        except FileExistsError:
            try:
                age = time.time() - os.path.getmtime(LOCK_PATH)
            except FileNotFoundError:
                continue
            if age > stale_after_sec:
                logger.warning("古いスクレイピングロックを奪取: age=%.0f秒", age)
                try:
                    os.remove(LOCK_PATH)
                except FileNotFoundError:
                    pass
                continue
            if not blocking:
                raise ScrapingLockBusy(
                    f"他プロセスがスクレイピング実行中です (age={age:.0f}秒, lock={LOCK_PATH})"
                )
            logger.info("スクレイピングロック待機中 (age=%.0f秒)", age)
            time.sleep(poll_sec)

    try:
        yield LOCK_PATH
    finally:
        try:
            os.remove(LOCK_PATH)
        except FileNotFoundError:
            pass


def acquire_main_instance_lock(stale_after_sec: int = 600) -> None:
    """main.py の多重起動を防ぐロックを取得する。

    既存ロックのPIDが実在プロセスなら MainInstanceBusy を送出。
    プロセスが存在しない、または古いロックは奪取する。

    Args:
        stale_after_sec: mtime がこの秒数以上経過していれば、PIDチェックを経ずに奪取
    """
    os.makedirs(os.path.dirname(MAIN_LOCK_PATH), exist_ok=True)

    def _write_lock() -> None:
        with open(MAIN_LOCK_PATH, "w") as f:
            f.write(f"{os.getpid()}\n{time.time()}\n")

    if not os.path.exists(MAIN_LOCK_PATH):
        _write_lock()
        return

    try:
        with open(MAIN_LOCK_PATH) as f:
            lines = f.read().splitlines()
        prev_pid = int(lines[0]) if lines else 0
    except (OSError, ValueError):
        prev_pid = 0

    try:
        age = time.time() - os.path.getmtime(MAIN_LOCK_PATH)
    except FileNotFoundError:
        _write_lock()
        return

    if age > stale_after_sec or not _pid_alive(prev_pid):
        logger.warning("古い main ロックを奪取: pid=%s, age=%.0f秒", prev_pid, age)
        _write_lock()
        return

    raise MainInstanceBusy(f"main.py は既に PID={prev_pid} で起動中です")


def release_main_instance_lock() -> None:
    """main.py 終了時にロックファイルを削除する。"""
    try:
        os.remove(MAIN_LOCK_PATH)
    except FileNotFoundError:
        pass


def _pid_alive(pid: int) -> bool:
    """PIDのプロセスが存在するか確認する（Windows対応）。"""
    if pid <= 0:
        return False
    import sys
    if sys.platform == "win32":
        import subprocess
        try:
            out = subprocess.check_output(
                ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
                stderr=subprocess.DEVNULL,
                text=True,
            )
            return str(pid) in out
        except subprocess.CalledProcessError:
            return False
    else:
        try:
            os.kill(pid, 0)
            return True
        except (OSError, ProcessLookupError):
            return False
