"""自動メール送信モジュール

実店舗向けの注文メールを自動送信する。
ECサイトへの自動発注は今回スコープ外。
"""

import sys
from datetime import datetime

from src.common.config import get_setting, load_mail_templates
from src.common.database import get_connection
from src.common.enums import MailSendStatus, OrderStatus
from src.common.exceptions import MailSendError
from src.common.logger import get_logger

logger = get_logger(__name__)

# メールテンプレート（mail_templates.yaml が空またはキー不足時のフォールバック）
DEFAULT_SUBJECT_TEMPLATE = "【注文依頼】{product_name} ({product_code})"
DEFAULT_BODY_TEMPLATE = """{supplier_name} 御中

いつもお世話になっております。
下記商品の注文をお願いいたします。

商品名: {product_name}
商品コード: {product_code}
数量: {quantity}冊
金額: {amount}円

よろしくお願いいたします。
"""


def _get_store_order_templates() -> tuple[str, str]:
    """mail_templates.yaml から store_order テンプレートを取得する。

    Returns:
        (subject_template, body_template)
    """
    try:
        templates = load_mail_templates()
        store = templates.get("store_order", {})
        subject = store.get("subject", "").strip()
        body = store.get("body", "").strip()
        return (
            subject if subject else DEFAULT_SUBJECT_TEMPLATE,
            body if body else DEFAULT_BODY_TEMPLATE,
        )
    except Exception:
        return DEFAULT_SUBJECT_TEMPLATE, DEFAULT_BODY_TEMPLATE


def _send_via_outlook(to_address: str, subject: str, body: str) -> None:
    """Outlook経由でメール送信する（Windows専用）。"""
    if sys.platform != "win32":
        raise MailSendError(
            "Outlook COM送信はWindowsでのみ利用可能です。"
            "settings.yaml の outlook.mode を 'imap' に変更してSMTP送信を使用してください。"
        )
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = to_address
        mail.Subject = subject
        mail.Body = body
        mail.Send()
    except Exception as e:
        raise MailSendError(f"Outlookメール送信エラー: {e}")


def _send_via_smtp(to_address: str, subject: str, body: str) -> None:
    """SMTP経由でメール送信する（Outlookがない環境用）。"""
    import smtplib
    from email.mime.text import MIMEText

    account = get_setting("outlook", "account_name", default="")
    password = get_setting("outlook", "imap_password", default="")
    if not account or not password:
        raise MailSendError("SMTP送信: メールアカウントまたはパスワードが設定されていません")

    # SMTPサーバーを自動検出
    domain = account.split("@")[-1].lower()
    smtp_servers = {
        "gmail.com": ("smtp.gmail.com", 587),
        "outlook.com": ("smtp-mail.outlook.com", 587),
        "hotmail.com": ("smtp-mail.outlook.com", 587),
        "yahoo.co.jp": ("smtp.mail.yahoo.co.jp", 465),
    }
    smtp_host, smtp_port = smtp_servers.get(domain, (f"smtp.{domain}", 587))

    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = subject
    msg["From"] = account
    msg["To"] = to_address

    try:
        if smtp_port == 465:
            server = smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=30)
        else:
            server = smtplib.SMTP(smtp_host, smtp_port, timeout=30)
            server.starttls()
        server.login(account, password)
        server.send_message(msg)
        server.quit()
    except Exception as e:
        raise MailSendError(f"SMTPメール送信エラー: {e}")


def send_order_mail(
    order_item_id: int,
    supplier_id: int,
    subject_template: str | None = None,
    body_template: str | None = None,
) -> bool:
    """1商品について注文メールを送信する。

    Args:
        order_item_id: order_items.id
        supplier_id: suppliers.id
        subject_template: 件名テンプレート（省略時はデフォルト）
        body_template: 本文テンプレート（省略時はデフォルト）

    Returns:
        送信成功ならTrue
    """
    with get_connection() as conn:
        item = conn.execute(
            "SELECT * FROM order_items WHERE id = ?",
            (order_item_id,),
        ).fetchone()
        supplier = conn.execute(
            "SELECT * FROM suppliers WHERE id = ?",
            (supplier_id,),
        ).fetchone()

    if item is None or supplier is None:
        logger.error("送信対象が見つかりません (item=%d, supplier=%d)", order_item_id, supplier_id)
        return False

    to_address = supplier["mail_to_address"]
    if not to_address:
        logger.warning(
            "メール宛先未設定: %s (supplier_id=%d)",
            supplier["supplier_name"], supplier_id,
        )
        return False

    # 既に送信済みかチェック
    with get_connection() as conn:
        existing = conn.execute(
            """
            SELECT id FROM outgoing_mails
            WHERE order_item_id = ? AND supplier_id = ? AND send_status = 'SUCCESS'
            """,
            (order_item_id, supplier_id),
        ).fetchone()
    if existing:
        logger.info("送信済みのためスキップ: item=%d, supplier=%d", order_item_id, supplier_id)
        return True

    # テンプレート差込（引数指定 > mail_templates.yaml > ハードコードデフォルト）
    yaml_subject, yaml_body = _get_store_order_templates()
    params = {
        "supplier_name": supplier["supplier_name"],
        "product_name": item["product_name"] or "",
        "product_code": item["product_code_normalized"] or "",
        "amount": item["amount"] or 0,
        "quantity": item["quantity"] or 1,
    }

    subject = (subject_template or yaml_subject).format(**params)
    body = (body_template or yaml_body).format(**params)

    # 確認モード
    confirmation_mode = get_setting("mail", "confirmation_mode", default=False)
    if confirmation_mode:
        logger.info(
            "[確認モード] メール送信をスキップ: to=%s, subject=%s",
            to_address, subject,
        )
        _save_mail_log(order_item_id, supplier_id, to_address, subject, body,
                       MailSendStatus.SUCCESS, "確認モード - 実送信なし")
        return True

    # 送信（Outlookを試し、失敗したらSMTPにフォールバック）
    # Windows以外ではCOMが使えないため、デフォルトはimap（=SMTP送信）
    error_msg = ""
    status = MailSendStatus.SUCCESS
    default_mode = "com" if sys.platform == "win32" else "imap"
    mode = get_setting("outlook", "mode", default=default_mode)
    try:
        if mode == "com":
            _send_via_outlook(to_address, subject, body)
        else:
            _send_via_smtp(to_address, subject, body)
        logger.info("メール送信成功: to=%s, subject=%s", to_address, subject)
    except MailSendError as e:
        # COM失敗時にSMTPフォールバック
        if mode == "com":
            try:
                _send_via_smtp(to_address, subject, body)
                logger.info("メール送信成功(SMTP): to=%s, subject=%s", to_address, subject)
            except MailSendError as e2:
                error_msg = str(e2)
                status = MailSendStatus.FAILED
                logger.error("メール送信失敗: %s", e2)
        else:
            error_msg = str(e)
            status = MailSendStatus.FAILED
            logger.error("メール送信失敗: %s", e)

    # 送信ログ保存
    _save_mail_log(order_item_id, supplier_id, to_address, subject, body, status, error_msg)

    return status == MailSendStatus.SUCCESS


def _save_mail_log(
    order_item_id: int,
    supplier_id: int,
    to_address: str,
    subject: str,
    body: str,
    status: MailSendStatus,
    error_message: str = "",
) -> None:
    """送信メール履歴をDBに保存する。"""
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO outgoing_mails (
                supplier_id, order_item_id, to_address,
                subject, body, send_status, error_message
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (supplier_id, order_item_id, to_address,
             subject, body, status.value, error_message),
        )


def process_auto_mails() -> int:
    """自動メール対象の確定商品に対してメールを送信する。

    対象:
    - current_status = 'ORDERED'
    - planned_supplier の auto_mail_enabled = true
    - 未送信

    Returns:
        送信件数
    """
    max_per_run = get_setting("mail", "max_send_per_run", default=10)

    with get_connection() as conn:
        items = conn.execute(
            """
            SELECT oi.id AS order_item_id, oi.planned_supplier_id
            FROM order_items oi
            JOIN suppliers s ON oi.planned_supplier_id = s.id
            WHERE oi.current_status = ?
              AND s.auto_mail_enabled = 1
              AND oi.id NOT IN (
                  SELECT order_item_id FROM outgoing_mails WHERE send_status = 'SUCCESS'
              )
            ORDER BY oi.id ASC
            LIMIT ?
            """,
            (OrderStatus.ORDERED.value, max_per_run),
        ).fetchall()

    sent_count = 0
    for item in items:
        try:
            if send_order_mail(item["order_item_id"], item["planned_supplier_id"]):
                sent_count += 1
        except Exception as e:
            logger.error("自動メール処理エラー (item=%d): %s", item["order_item_id"], e)

    if sent_count > 0:
        logger.info("自動メール送信完了: %d件", sent_count)
    return sent_count
