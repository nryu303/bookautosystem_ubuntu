"""Outlook受信メール取得モジュール

win32comを使用してOutlookから新着受注メールを読み取り、
messagesテーブルに保存する。

注意: このモジュールはWindows環境でのみ動作する。
Linux/macOSではIMAP方式（imap_reader.py）またはGraph API方式（graph_reader.py）を使用すること。
"""

import sys
from datetime import datetime

from src.common.database import get_connection
from src.common.exceptions import MailReadError
from src.common.logger import get_logger
from src.common.config import get_setting

logger = get_logger(__name__)


def _get_outlook():
    """Outlook.Applicationオブジェクトを取得する。"""
    if sys.platform != "win32":
        raise MailReadError(
            "Outlook COM方式はWindowsでのみ利用可能です。"
            "settings.yaml の outlook.mode を 'imap' または 'graph' に変更してください。"
        )
    try:
        import win32com.client
        return win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        raise MailReadError(f"Outlook接続に失敗しました。Outlookが起動しているか確認してください: {e}")


def _find_folder(namespace, account_name: str, folder_name: str):
    """指定アカウント・フォルダパスのMAPIFolderを返す。

    folder_name は "/" 区切りで階層を表す。
    例: "受信トレイ/受注"
    """
    # アカウント検索
    target_store = None
    for store in namespace.Stores:
        if account_name.lower() in store.DisplayName.lower():
            target_store = store
            break

    if target_store is None:
        # アカウントが見つからない場合、デフォルトの受信トレイを使う
        logger.warning(
            "アカウント '%s' が見つかりません。デフォルトの受信トレイを使用します。",
            account_name,
        )
        return namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

    # フォルダ階層をたどる
    folder = target_store.GetRootFolder()
    for part in folder_name.split("/"):
        part = part.strip()
        if not part:
            continue
        found = False
        for sub in folder.Folders:
            if sub.Name == part:
                folder = sub
                found = True
                break
        if not found:
            raise MailReadError(
                f"フォルダ '{part}' が見つかりません (パス: {folder_name})"
            )

    return folder


def _get_existing_message_ids() -> set[str]:
    """既にDBに保存済みのoutlook_message_idの集合を返す。"""
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT outlook_message_id FROM messages"
        ).fetchall()
    return {row["outlook_message_id"] for row in rows}


def _save_message(msg) -> bool:
    """1通のOutlookメールをmessagesテーブルに保存する。

    Returns:
        True: 新規保存成功
        False: 既存（スキップ）
    """
    try:
        message_id = msg.EntryID
        if not message_id:
            logger.warning("EntryIDが空のメールをスキップします")
            return False

        # 各フィールド取得（取得失敗に備えて個別にtry）
        sender = ""
        try:
            sender = msg.SenderEmailAddress or msg.SenderName or ""
        except Exception:
            pass

        recipient = ""
        try:
            recipient = msg.To or ""
        except Exception:
            pass

        received_at = ""
        try:
            received_at = msg.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            received_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        subject = ""
        try:
            subject = msg.Subject or ""
        except Exception:
            pass

        raw_html = ""
        try:
            raw_html = msg.HTMLBody or ""
        except Exception:
            pass

        raw_text = ""
        try:
            raw_text = msg.Body or ""
        except Exception:
            pass

        account_name = get_setting("outlook", "account_name", default="")

        with get_connection() as conn:
            conn.execute(
                """
                INSERT INTO messages (
                    outlook_message_id, account_name, sender, recipient,
                    received_at, subject, raw_html, raw_text,
                    parse_status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'PENDING')
                """,
                (
                    message_id, account_name, sender, recipient,
                    received_at, subject, raw_html, raw_text,
                ),
            )
        return True

    except Exception as e:
        logger.error("メール保存エラー (subject=%s): %s", getattr(msg, "Subject", "?"), e)
        return False


def fetch_new_mails(max_items: int = 100) -> int:
    """Outlookから新着メールを取得してDBに保存する。

    Args:
        max_items: 1回の取得で処理する最大件数

    Returns:
        新規保存件数
    """
    account_name = get_setting("outlook", "account_name", default="")
    folder_name = get_setting("outlook", "folder_name", default="受信トレイ")

    logger.info("メール取得開始 (account=%s, folder=%s)", account_name, folder_name)

    try:
        outlook = _get_outlook()
        namespace = outlook.GetNamespace("MAPI")
    except MailReadError:
        raise
    except Exception as e:
        raise MailReadError(f"Outlook MAPI接続エラー: {e}")

    try:
        folder = _find_folder(namespace, account_name, folder_name)
    except MailReadError:
        raise
    except Exception as e:
        raise MailReadError(f"フォルダ取得エラー: {e}")

    # 既存IDを取得して重複チェック用に保持
    existing_ids = _get_existing_message_ids()

    items = folder.Items
    items.Sort("[ReceivedTime]", True)  # 新しい順

    saved_count = 0
    checked_count = 0

    for item in items:
        if checked_count >= max_items:
            break
        checked_count += 1

        try:
            entry_id = item.EntryID
            if entry_id in existing_ids:
                continue

            if _save_message(item):
                saved_count += 1
                existing_ids.add(entry_id)

        except Exception as e:
            logger.warning("メール処理中にエラー（スキップして継続）: %s", e)
            continue

    logger.info(
        "メール取得完了: %d件チェック, %d件新規保存",
        checked_count, saved_count,
    )
    return saved_count
