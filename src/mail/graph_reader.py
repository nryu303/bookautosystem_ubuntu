"""Microsoft Graph API メール取得モジュール

新しいOutlook（UWPアプリ）やOutlookがインストールされていない環境で、
Microsoft Graph API経由でメールを取得する。

認証: MSAL デバイスコードフロー（初回のみブラウザで認証、以降はトークンキャッシュ）
"""

import json
import os
import re
from datetime import datetime

import msal
import requests

from src.common.database import get_connection
from src.common.config import get_setting, BASE_DIR
from src.common.exceptions import MailReadError
from src.common.logger import get_logger

logger = get_logger(__name__)

# Microsoft Graph API endpoints
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.Read"]

# トークンキャッシュファイルパス
TOKEN_CACHE_PATH = os.path.join(BASE_DIR, "data", "graph_token_cache.json")


def _get_msal_app() -> msal.PublicClientApplication:
    """MSALアプリケーションを構築する。トークンキャッシュ付き。"""
    client_id = get_setting("outlook", "graph_client_id", default="")
    tenant_id = get_setting("outlook", "graph_tenant_id", default="consumers")

    if not client_id:
        raise MailReadError(
            "graph_client_id が設定されていません。\n"
            "config/settings.yaml の outlook.graph_client_id に "
            "Azure ADアプリのクライアントIDを設定してください。"
        )

    # トークンキャッシュの読み込み
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_PATH):
        with open(TOKEN_CACHE_PATH, "r", encoding="utf-8") as f:
            cache.deserialize(f.read())

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.PublicClientApplication(
        client_id,
        authority=authority,
        token_cache=cache,
    )
    return app


def _save_token_cache(app: msal.PublicClientApplication) -> None:
    """トークンキャッシュをファイルに保存する。"""
    cache = app.token_cache
    if cache.has_state_changed:
        os.makedirs(os.path.dirname(TOKEN_CACHE_PATH), exist_ok=True)
        with open(TOKEN_CACHE_PATH, "w", encoding="utf-8") as f:
            f.write(cache.serialize())


def _acquire_token(app: msal.PublicClientApplication) -> str:
    """アクセストークンを取得する。

    1. キャッシュからサイレント取得を試みる
    2. 失敗時はデバイスコードフローで対話認証

    Returns:
        アクセストークン文字列
    """
    accounts = app.get_accounts()
    result = None

    # キャッシュからサイレント取得
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if result and "access_token" in result:
        logger.info("トークンをキャッシュから取得しました")
        _save_token_cache(app)
        return result["access_token"]

    # デバイスコードフローで認証
    logger.info("ブラウザでの認証が必要です...")
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise MailReadError(
            f"デバイスコードフローの開始に失敗しました: {flow.get('error_description', '不明なエラー')}"
        )

    # ユーザーに認証URLとコードを表示
    print("\n" + "=" * 60)
    print("【Graph API 認証】")
    print(f"  1. ブラウザで次のURLを開いてください:")
    print(f"     {flow['verification_uri']}")
    print(f"  2. 次のコードを入力してください:")
    print(f"     {flow['user_code']}")
    print("=" * 60 + "\n")

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        error_desc = result.get("error_description", "認証に失敗しました")
        raise MailReadError(f"Graph API認証エラー: {error_desc}")

    logger.info("Graph API認証成功")
    _save_token_cache(app)
    return result["access_token"]


def _graph_get(token: str, url: str, params: dict = None) -> dict:
    """Graph APIにGETリクエストを送る。"""
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers, params=params, timeout=30)

    if resp.status_code == 401:
        raise MailReadError("Graph APIトークンが無効です。再認証が必要です。")
    if resp.status_code != 200:
        raise MailReadError(
            f"Graph APIエラー (HTTP {resp.status_code}): {resp.text[:500]}"
        )
    return resp.json()


def _find_folder_id(token: str, folder_path: str) -> str:
    """フォルダパス（例: '受信トレイ/受注'）からフォルダIDを取得する。

    Graph APIでは日本語フォルダ名を直接使ってフォルダを検索する。
    """
    parts = [p.strip() for p in folder_path.split("/") if p.strip()]

    # 最初のフォルダ（トップレベル）を検索
    url = f"{GRAPH_BASE}/me/mailFolders"
    current_folder_id = None

    for i, part in enumerate(parts):
        if i == 0:
            # トップレベルフォルダを検索
            data = _graph_get(token, url)
            folders = data.get("value", [])
            matched = _match_folder(folders, part)
            if not matched:
                raise MailReadError(
                    f"フォルダ '{part}' が見つかりません。"
                    f"利用可能: {[f['displayName'] for f in folders]}"
                )
            current_folder_id = matched["id"]
        else:
            # 子フォルダを検索
            child_url = f"{GRAPH_BASE}/me/mailFolders/{current_folder_id}/childFolders"
            data = _graph_get(token, child_url)
            folders = data.get("value", [])
            matched = _match_folder(folders, part)
            if not matched:
                raise MailReadError(
                    f"サブフォルダ '{part}' が見つかりません。"
                    f"利用可能: {[f['displayName'] for f in folders]}"
                )
            current_folder_id = matched["id"]

    return current_folder_id


def _match_folder(folders: list[dict], name: str) -> dict | None:
    """フォルダ名で一致するフォルダを探す。

    Graph APIの wellKnownName と displayName の両方をチェック。
    '受信トレイ' は wellKnownName='inbox' にもマッチさせる。
    """
    inbox_aliases = {"受信トレイ", "inbox", "Inbox"}

    for folder in folders:
        display = folder.get("displayName", "")
        well_known = folder.get("wellKnownName", "")

        # 完全一致
        if display == name:
            return folder

        # 受信トレイの特別対応
        if name in inbox_aliases and well_known == "inbox":
            return folder

    return None


def _get_existing_message_ids() -> set[str]:
    """既にDBに保存済みのoutlook_message_idの集合を返す。"""
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT outlook_message_id FROM messages"
        ).fetchall()
    return {row["outlook_message_id"] for row in rows}


def _parse_datetime(iso_str: str) -> str:
    """ISO 8601日時文字列を 'YYYY-MM-DD HH:MM:SS' 形式に変換する。"""
    if not iso_str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        # '2026-03-24T14:30:00Z' or '2026-03-24T14:30:00+09:00'
        cleaned = re.sub(r"[Zz]$", "+00:00", iso_str)
        dt = datetime.fromisoformat(cleaned)
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except (ValueError, TypeError):
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _save_message(msg: dict) -> bool:
    """Graph APIのメッセージをmessagesテーブルに保存する。

    Returns:
        True: 新規保存成功, False: スキップ
    """
    try:
        message_id = msg.get("id", "")
        if not message_id:
            return False

        sender_info = msg.get("from", {}).get("emailAddress", {})
        sender = sender_info.get("address", sender_info.get("name", ""))

        recipients = msg.get("toRecipients", [])
        recipient = ""
        if recipients:
            r = recipients[0].get("emailAddress", {})
            recipient = r.get("address", "")

        received_at = _parse_datetime(msg.get("receivedDateTime", ""))
        subject = msg.get("subject", "")

        # Graph APIはHTMLとテキスト両方を取得可能
        body = msg.get("body", {})
        raw_html = ""
        raw_text = ""
        if body.get("contentType") == "html":
            raw_html = body.get("content", "")
        else:
            raw_text = body.get("content", "")

        # HTMLがある場合、テキスト版も別途取得を試みる
        if raw_html and not raw_text:
            # uniqueBody がある場合はそれをテキストとして使う
            unique_body = msg.get("uniqueBody", {})
            if unique_body:
                raw_text = unique_body.get("content", "")

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
        logger.error("メール保存エラー (subject=%s): %s", msg.get("subject", "?"), e)
        return False


def fetch_new_mails(max_items: int = 100) -> int:
    """Microsoft Graph API経由で新着メールを取得してDBに保存する。

    Args:
        max_items: 1回の取得で処理する最大件数

    Returns:
        新規保存件数
    """
    folder_name = get_setting("outlook", "folder_name", default="受信トレイ")
    logger.info("メール取得開始 [Graph API] (folder=%s)", folder_name)

    try:
        app = _get_msal_app()
        token = _acquire_token(app)
    except MailReadError:
        raise
    except Exception as e:
        raise MailReadError(f"Graph API認証エラー: {e}")

    try:
        folder_id = _find_folder_id(token, folder_name)
    except MailReadError:
        raise
    except Exception as e:
        raise MailReadError(f"フォルダ取得エラー: {e}")

    # 既存IDで重複チェック
    existing_ids = _get_existing_message_ids()

    # メール取得（新しい順）
    url = f"{GRAPH_BASE}/me/mailFolders/{folder_id}/messages"
    params = {
        "$top": max_items,
        "$orderby": "receivedDateTime desc",
        "$select": "id,from,toRecipients,receivedDateTime,subject,body,uniqueBody",
    }

    try:
        data = _graph_get(token, url, params)
    except MailReadError:
        raise
    except Exception as e:
        raise MailReadError(f"メール取得エラー: {e}")

    messages = data.get("value", [])
    saved_count = 0

    for msg in messages:
        msg_id = msg.get("id", "")
        if msg_id in existing_ids:
            continue

        if _save_message(msg):
            saved_count += 1
            existing_ids.add(msg_id)

    logger.info(
        "メール取得完了 [Graph API]: %d件チェック, %d件新規保存",
        len(messages), saved_count,
    )
    return saved_count
