"""IMAP メール取得モジュール

Outlook.comやGmailなど、IMAP対応のメールサービスから
受注メールを取得する。Outlookデスクトップ版が不要。

Outlook.com: OAuth2 (XOAUTH2) 認証 — Basic認証は廃止済み
Gmail:       アプリパスワードによるBasic認証
その他:      Basic認証（パスワード直接入力）

対応サービス:
  - Outlook.com (imap-mail.outlook.com) — OAuth2自動対応
  - Gmail (imap.gmail.com)
  - その他IMAP対応サービス
"""

import base64
import email
import email.header
import imaplib
import json
import os
import re
from datetime import datetime

from src.common.database import get_connection
from src.common.exceptions import MailReadError
from src.common.logger import get_logger
from src.common.config import get_setting, BASE_DIR

logger = get_logger(__name__)

# 主要サービスのIMAPサーバー設定
KNOWN_IMAP_SERVERS = {
    "outlook.com": ("imap-mail.outlook.com", 993),
    "hotmail.com": ("imap-mail.outlook.com", 993),
    "live.com": ("imap-mail.outlook.com", 993),
    "gmail.com": ("imap.gmail.com", 993),
    "yahoo.co.jp": ("imap.mail.yahoo.co.jp", 993),
}

# Microsoft Office Desktop 公開クライアントID（個人/組織アカウント両対応）
OUTLOOK_PUBLIC_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
OUTLOOK_SCOPES = [
    "https://outlook.office.com/IMAP.AccessAsUser.All",
]
OAUTH_TOKEN_CACHE_PATH = os.path.join(BASE_DIR, "data", "imap_oauth_token.json")

# OAuth2が必要なドメイン
OAUTH2_DOMAINS = {"outlook.com", "hotmail.com", "live.com", "outlook.jp"}


def _is_oauth2_required(email_addr: str) -> bool:
    """メールアドレスのドメインがOAuth2必須かどうかを判定する。"""
    domain = email_addr.split("@")[-1].lower()
    return domain in OAUTH2_DOMAINS


def _detect_imap_server(email_addr: str) -> tuple[str, int]:
    """メールアドレスのドメインからIMAPサーバーを自動検出する。"""
    host = get_setting("outlook", "imap_host", default="")
    port = get_setting("outlook", "imap_port", default=993)

    if host:
        return host, int(port)

    domain = email_addr.split("@")[-1].lower()
    if domain in KNOWN_IMAP_SERVERS:
        return KNOWN_IMAP_SERVERS[domain]

    raise MailReadError(
        f"IMAPサーバーを自動検出できません (ドメイン: {domain})。\n"
        f"config/settings.yaml の outlook.imap_host と outlook.imap_port を設定してください。"
    )


def _get_oauth2_token(email_addr: str) -> str:
    """Outlook.com用のOAuth2アクセストークンを取得する。

    MSALデバイスコードフローを使用。初回のみブラウザ認証が必要。
    以降はトークンキャッシュからサイレント取得。
    """
    try:
        import msal
    except ImportError:
        raise MailReadError(
            "msalパッケージが必要です。以下を実行してください:\n"
            "  pip install msal"
        )

    # トークンキャッシュの読み込み
    cache = msal.SerializableTokenCache()
    if os.path.exists(OAUTH_TOKEN_CACHE_PATH):
        with open(OAUTH_TOKEN_CACHE_PATH, "r", encoding="utf-8") as f:
            cache.deserialize(f.read())

    app = msal.PublicClientApplication(
        OUTLOOK_PUBLIC_CLIENT_ID,
        authority="https://login.microsoftonline.com/common",
        token_cache=cache,
    )

    # キャッシュからサイレント取得
    accounts = app.get_accounts()
    result = None
    if accounts:
        # 指定アカウントを優先的に探す
        target_account = None
        for acc in accounts:
            if acc.get("username", "").lower() == email_addr.lower():
                target_account = acc
                break
        if target_account is None:
            target_account = accounts[0]

        result = app.acquire_token_silent(OUTLOOK_SCOPES, account=target_account)

    if result and "access_token" in result:
        logger.info("OAuth2トークンをキャッシュから取得しました")
        _save_oauth_cache(cache)
        return result["access_token"]

    # デバイスコードフローで対話認証
    logger.info("OAuth2認証が必要です。ブラウザでの認証を開始します...")
    flow = app.initiate_device_flow(scopes=OUTLOOK_SCOPES)
    if "user_code" not in flow:
        raise MailReadError(
            f"OAuth2デバイスコードフローの開始に失敗: {flow.get('error_description', '不明なエラー')}"
        )

    print("\n" + "=" * 60)
    print("【IMAP OAuth2 認証】")
    print(f"  1. ブラウザで次のURLを開いてください:")
    print(f"     {flow['verification_uri']}")
    print(f"  2. 次のコードを入力してください:")
    print(f"     {flow['user_code']}")
    print("=" * 60 + "\n")

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        error_desc = result.get("error_description", "認証に失敗しました")
        raise MailReadError(f"OAuth2認証エラー: {error_desc}")

    logger.info("OAuth2認証成功")
    _save_oauth_cache(cache)
    return result["access_token"]


def _save_oauth_cache(cache) -> None:
    """トークンキャッシュをファイルに保存する。"""
    if cache.has_state_changed:
        os.makedirs(os.path.dirname(OAUTH_TOKEN_CACHE_PATH), exist_ok=True)
        with open(OAUTH_TOKEN_CACHE_PATH, "w", encoding="utf-8") as f:
            f.write(cache.serialize())


def _build_xoauth2_string(user: str, token: str) -> str:
    """XOAUTH2認証文字列を構築する。"""
    auth_string = f"user={user}\x01auth=Bearer {token}\x01\x01"
    return auth_string


def _connect(email_addr: str, password: str) -> imaplib.IMAP4_SSL:
    """IMAPサーバーに接続・ログインする。

    Outlook.com: OAuth2 XOAUTH2認証
    その他:      Basic認証（パスワード）
    """
    host, port = _detect_imap_server(email_addr)
    logger.info("IMAP接続: %s:%d (user=%s)", host, port, email_addr)

    try:
        conn = imaplib.IMAP4_SSL(host, port, timeout=30)
    except Exception as e:
        raise MailReadError(f"IMAPサーバーへの接続に失敗しました ({host}:{port}): {e}")

    if _is_oauth2_required(email_addr):
        # OAuth2 XOAUTH2 認証を試行
        logger.info("OAuth2 (XOAUTH2) 認証を使用します")
        try:
            token = _get_oauth2_token(email_addr)
            auth_string = _build_xoauth2_string(email_addr, token)
            conn.authenticate("XOAUTH2", lambda x: auth_string.encode("utf-8"))
        except Exception as e:
            logger.warning("OAuth2認証失敗、パスワード認証にフォールバックします: %s", e)
            # OAuth2失敗時はパスワード認証にフォールバック
            conn = imaplib.IMAP4_SSL(host, port, timeout=30)
            if password:
                try:
                    conn.login(email_addr, password)
                    logger.info("パスワード認証でログイン成功")
                except imaplib.IMAP4.error as e2:
                    raise MailReadError(
                        f"OAuth2認証とパスワード認証の両方に失敗しました: {e2}"
                    )
            else:
                raise MailReadError(
                    f"OAuth2認証に失敗し、パスワードも設定されていません: {e}"
                )
    else:
        # Basic認証（パスワード）
        if not password:
            raise MailReadError(
                "outlook.imap_password が設定されていません。\n"
                "config/settings.yaml の outlook.imap_password にパスワードを設定してください。"
            )
        try:
            conn.login(email_addr, password)
        except imaplib.IMAP4.error as e:
            raise MailReadError(
                f"IMAPログインに失敗しました。メールアドレスとパスワードを確認してください: {e}"
            )

    return conn


def _decode_header_value(raw: str) -> str:
    """MIMEエンコードされたヘッダー値をデコードする。"""
    if not raw:
        return ""
    decoded_parts = email.header.decode_header(raw)
    result = []
    for part, charset in decoded_parts:
        if isinstance(part, bytes):
            result.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            result.append(part)
    return "".join(result)


def _get_body(msg: email.message.Message) -> tuple[str, str]:
    """メールからHTMLとテキスト本文を抽出する。

    Returns:
        (raw_html, raw_text)
    """
    raw_html = ""
    raw_text = ""

    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if part.get("Content-Disposition") == "attachment":
                continue

            try:
                charset = part.get_content_charset() or "utf-8"
                payload = part.get_payload(decode=True)
                if payload is None:
                    continue
                decoded = payload.decode(charset, errors="replace")
            except Exception:
                continue

            if content_type == "text/html" and not raw_html:
                raw_html = decoded
            elif content_type == "text/plain" and not raw_text:
                raw_text = decoded
    else:
        content_type = msg.get_content_type()
        try:
            charset = msg.get_content_charset() or "utf-8"
            payload = msg.get_payload(decode=True)
            if payload:
                decoded = payload.decode(charset, errors="replace")
                if content_type == "text/html":
                    raw_html = decoded
                else:
                    raw_text = decoded
        except Exception:
            pass

    return raw_html, raw_text


def _parse_date(msg: email.message.Message) -> str:
    """メールのDateヘッダーを 'YYYY-MM-DD HH:MM:SS' 形式に変換する。"""
    date_str = msg.get("Date", "")
    if not date_str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        from email.utils import parsedate_to_datetime
        dt = parsedate_to_datetime(date_str)
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _extract_email_address(header_value: str) -> str:
    """ヘッダー値からメールアドレスを抽出する。"""
    if not header_value:
        return ""
    match = re.search(r"[\w.+-]+@[\w.-]+", header_value)
    return match.group(0) if match else header_value


def _get_existing_message_ids() -> set[str]:
    """既にDBに保存済みのoutlook_message_idの集合を返す。"""
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT outlook_message_id FROM messages"
        ).fetchall()
    return {row["outlook_message_id"] for row in rows}


def _resolve_folder_name(conn: imaplib.IMAP4_SSL, folder_path: str) -> str:
    """設定のフォルダパスからIMAPフォルダ名を解決する。

    '受信トレイ/受注' → 'INBOX/受注' のように変換。
    """
    inbox_aliases = {"受信トレイ", "inbox", "Inbox", "INBOX"}

    parts = [p.strip() for p in folder_path.split("/") if p.strip()]
    if not parts:
        return "INBOX"

    # 最初の部分が受信トレイの場合、INBOXに変換
    if parts[0] in inbox_aliases:
        parts[0] = "INBOX"

    # フォルダ一覧を取得して実際の名前を探す
    status, folder_list = conn.list()
    if status != "OK":
        raise MailReadError("IMAPフォルダ一覧の取得に失敗しました")

    target = "/".join(parts)
    available_folders = []

    for folder_data in folder_list:
        if isinstance(folder_data, bytes):
            folder_str = folder_data.decode("utf-8", errors="replace")
        else:
            folder_str = str(folder_data)

        # フォルダ名を抽出（例: '(\\HasNoChildren) "/" "INBOX/受注"' → 'INBOX/受注'）
        match = re.search(r'"([^"]*)"$', folder_str)
        if match:
            fname = match.group(1)
        else:
            # クォートなしの場合
            parts_split = folder_str.rsplit(" ", 1)
            fname = parts_split[-1] if parts_split else folder_str

        available_folders.append(fname)

        if fname == target or fname.upper() == target.upper():
            return fname

    # 区切り文字が "/" でない場合（"." を使うサーバーもある）
    for sep in [".", "/"]:
        alt_target = sep.join(parts)
        for fname in available_folders:
            if fname == alt_target or fname.upper() == alt_target.upper():
                return fname

    raise MailReadError(
        f"IMAPフォルダ '{folder_path}' が見つかりません。\n"
        f"利用可能なフォルダ: {available_folders}"
    )


def _save_message(uid: str, msg: email.message.Message) -> bool:
    """1通のメールをmessagesテーブルに保存する。

    Returns:
        True: 新規保存成功, False: スキップ
    """
    try:
        message_id = f"imap_{uid}"

        subject = _decode_header_value(msg.get("Subject", ""))
        sender = _decode_header_value(msg.get("From", ""))
        sender = _extract_email_address(sender)
        recipient = _decode_header_value(msg.get("To", ""))
        recipient = _extract_email_address(recipient)
        received_at = _parse_date(msg)
        raw_html, raw_text = _get_body(msg)

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
        logger.error("メール保存エラー (uid=%s): %s", uid, e)
        return False


def fetch_new_mails(max_items: int = 100) -> int:
    """IMAP経由で新着メールを取得してDBに保存する。

    Args:
        max_items: 1回の取得で処理する最大件数

    Returns:
        新規保存件数
    """
    account_name = get_setting("outlook", "account_name", default="")
    password = get_setting("outlook", "imap_password", default="")
    folder_name = get_setting("outlook", "folder_name", default="受信トレイ")

    if not account_name:
        raise MailReadError("outlook.account_name が設定されていません。")

    logger.info("メール取得開始 [IMAP] (account=%s, folder=%s)", account_name, folder_name)

    conn = _connect(account_name, password)

    try:
        # フォルダを選択
        imap_folder = _resolve_folder_name(conn, folder_name)
        logger.info("IMAPフォルダ選択: %s", imap_folder)

        status, _ = conn.select(f'"{imap_folder}"', readonly=True)
        if status != "OK":
            raise MailReadError(f"フォルダ '{imap_folder}' の選択に失敗しました")

        # 全メールのUIDを取得（新しい順）
        status, data = conn.uid("search", None, "ALL")
        if status != "OK":
            raise MailReadError("メール検索に失敗しました")

        uid_list = data[0].split() if data[0] else []
        uid_list.reverse()  # 新しい順

        if not uid_list:
            logger.info("メール取得完了 [IMAP]: フォルダにメールがありません")
            conn.close()
            conn.logout()
            return 0

        # 処理上限を適用
        uid_list = uid_list[:max_items]

        # 既存IDで重複チェック
        existing_ids = _get_existing_message_ids()

        saved_count = 0
        checked_count = 0

        for uid_bytes in uid_list:
            uid = uid_bytes.decode("utf-8") if isinstance(uid_bytes, bytes) else str(uid_bytes)
            imap_id = f"imap_{uid}"

            checked_count += 1

            if imap_id in existing_ids:
                continue

            # メール本文を取得
            status, msg_data = conn.uid("fetch", uid_bytes, "(RFC822)")
            if status != "OK" or not msg_data or msg_data[0] is None:
                logger.warning("メール取得失敗 (UID=%s)", uid)
                continue

            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            if _save_message(uid, msg):
                saved_count += 1
                existing_ids.add(imap_id)

        logger.info(
            "メール取得完了 [IMAP]: %d件チェック, %d件新規保存",
            checked_count, saved_count,
        )

        conn.close()
        conn.logout()
        return saved_count

    except MailReadError:
        raise
    except Exception as e:
        raise MailReadError(f"IMAPメール取得エラー: {e}")
    finally:
        try:
            conn.logout()
        except Exception:
            pass
