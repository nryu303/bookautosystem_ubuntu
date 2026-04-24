"""メール本文解析・商品情報抽出モジュール

messagesテーブルの未解析メールからHTMLを読み、
商品コード・注文番号・商品名・金額・数量を抽出して
order_itemsテーブルに保存する。
"""

import re
from bs4 import BeautifulSoup

from src.common.database import get_connection
from src.common.config import load_mail_patterns, get_setting
from src.common.enums import ParseStatus
from src.common.exceptions import ParseError
from src.common.logger import get_logger
from src.parser.code_normalizer import normalize_code, extract_codes_from_text

logger = get_logger(__name__)


def _html_to_text(html: str) -> str:
    """HTMLをプレーンテキストに変換する。"""
    if not html:
        return ""
    soup = BeautifulSoup(html, "lxml")
    # scriptとstyleを除去
    for tag in soup(["script", "style"]):
        tag.decompose()
    text = soup.get_text(separator="\n")
    # 連続空行を1行にまとめる
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _extract_field(text: str, patterns: list[dict]) -> str | None:
    """パターンリストを順に試して最初にマッチした値を返す。"""
    for pat_def in patterns:
        pattern = pat_def.get("pattern", "")
        group = pat_def.get("group", 1)
        try:
            m = re.search(pattern, text, re.MULTILINE)
            if m:
                return m.group(group).strip()
        except (re.error, IndexError):
            continue
    return None


def _extract_amount(text: str, patterns: list[dict]) -> int:
    """金額を抽出して整数で返す。カンマ・通貨記号を除去。"""
    raw = _extract_field(text, patterns)
    if raw is None:
        return 0
    cleaned = re.sub(r"[¥￥,\s]", "", raw)
    try:
        return int(cleaned)
    except ValueError:
        return 0


def _extract_quantity(text: str, patterns: list[dict]) -> int:
    """数量を抽出して整数で返す。"""
    raw = _extract_field(text, patterns)
    if raw is None:
        return 1  # デフォルト1冊
    try:
        return max(1, int(raw))
    except ValueError:
        return 1


def _split_items(text: str, separator_patterns: list[str]) -> list[str]:
    """テキストを商品ごとのブロックに分割する。

    区切りパターンにマッチしない場合は全体を1ブロックとして返す。
    """
    for pattern in separator_patterns:
        try:
            parts = re.split(pattern, text)
            parts = [p.strip() for p in parts if p and p.strip()]
            if len(parts) > 1:
                return parts
        except re.error:
            continue
    return [text]


def _parse_single_item(
    text: str,
    mail_patterns: dict,
    message_id: int,
) -> dict | None:
    """テキストブロック1つから商品情報を抽出する。"""
    patterns = mail_patterns.get("patterns", {})

    # 注文番号を先に抽出（コード検出時の誤マッチ防止に使う）
    order_number = _extract_field(text, patterns.get("order_number", []))

    # 商品コード抽出
    product_codes = extract_codes_from_text(text)
    if not product_codes:
        # パターン定義からも試す
        # 注文番号を含む行を除外してからコード検出（注文番号の数字列を誤検出しないため）
        code_search_text = text
        if order_number:
            code_search_text = re.sub(
                r"(?:注文番号|オーダー|受注番号)[：:\s]*" + re.escape(order_number) + r"[^\n]*",
                "",
                text,
            )
        code_patterns = patterns.get("product_code", [])
        raw_code = _extract_field(code_search_text, code_patterns)
        if raw_code:
            normalized = normalize_code(raw_code)
            if normalized:
                product_codes = [normalized]

    # SKU（ISBN/ASIN未検出時のフォールバックにも使う）
    sku = _extract_field(text, patterns.get("sku", [])) or ""

    if not product_codes and sku:
        product_codes = [sku]

    if not product_codes:
        return None

    # 商品名
    product_name = _extract_field(text, patterns.get("product_name", []))

    # 金額
    amount = _extract_amount(text, patterns.get("amount", []))

    # 数量
    quantity = _extract_quantity(text, patterns.get("quantity", []))

    # コンディション
    item_condition = _extract_field(text, patterns.get("condition", [])) or ""

    # 宛先
    recipient_name = _extract_field(text, patterns.get("recipient", [])) or ""

    # 版本（商品名から [単行本] 等を抽出）
    edition = _extract_field(text, patterns.get("edition", [])) or ""

    # 出版年度（商品名から [2024] 等を抽出）
    pub_year = _extract_field(text, patterns.get("pub_year", [])) or ""

    # 著者（商品名の [2024] の後ろから抽出）
    author = _extract_field(text, patterns.get("author", [])) or ""

    # 注文日
    order_date = _extract_field(text, patterns.get("order_date", [])) or ""

    # 出荷予定日
    ship_date = _extract_field(text, patterns.get("ship_date", [])) or ""

    # 税金
    tax = _extract_amount(text, patterns.get("tax", []))

    # 配送料
    shipping = _extract_amount(text, patterns.get("shipping", []))

    # 商品名から版本・出版年度・著者を除去してクリーニング
    if product_name:
        # [単行本], [文庫], [新書] 等の版本表記を除去
        if edition:
            product_name = product_name.replace(f"[{edition}]", "")
        # [2024] 等の年度表記を除去
        if pub_year:
            product_name = product_name.replace(f"[{pub_year}]", "")
        # 著者名を除去
        if author:
            product_name = product_name.replace(author, "")
        # 残った余分な空白を整理
        product_name = re.sub(r"\s{2,}", " ", product_name).strip()

    # 最初に見つかった商品コードを使う
    raw_code = product_codes[0]
    normalized = normalize_code(raw_code)

    return {
        "message_id": message_id,
        "order_number": order_number or "",
        "product_code_raw": raw_code,
        "product_code_normalized": normalized or raw_code,
        "product_name": product_name or "",
        "amount": amount,
        "quantity": quantity,
        "item_condition": item_condition,
        "recipient_name": recipient_name,
        "edition": edition,
        "pub_year": pub_year,
        "author": author,
        "sku": sku,
        "order_date": order_date,
        "ship_date": ship_date,
        "tax": tax,
        "shipping": shipping,
    }


def _is_order_mail(sender: str, subject: str) -> bool:
    """メールが受注メールかどうかをフィルタで判定する。

    settings.yaml の mail_filter 設定に基づき、
    allowed_senders または allowed_subject_patterns のいずれかにマッチすれば True。
    フィルタ設定が空の場合は全メールを処理対象とする。
    """
    allowed_senders = get_setting("mail_filter", "allowed_senders", default=[]) or []
    allowed_subjects = get_setting("mail_filter", "allowed_subject_patterns", default=[]) or []

    # フィルタ未設定なら全メールを対象
    if not allowed_senders and not allowed_subjects:
        return True

    sender_lower = (sender or "").lower()
    subject_str = subject or ""

    # 送信者チェック
    for allowed in allowed_senders:
        if allowed and allowed.lower() in sender_lower:
            return True

    # 件名パターンチェック
    for pattern in allowed_subjects:
        if pattern and pattern in subject_str:
            return True

    return False


def parse_pending_messages() -> int:
    """未解析メールを全て解析し、order_itemsに保存する。

    Returns:
        新規作成した order_items の件数
    """
    mail_patterns = load_mail_patterns()
    separator_patterns = []
    sep_config = mail_patterns.get("item_separator", {})
    if sep_config:
        separator_patterns = sep_config.get("patterns", [])

    # 未解析メール取得
    with get_connection() as conn:
        pending = conn.execute(
            "SELECT id, raw_html, raw_text, subject, sender FROM messages WHERE parse_status = 'PENDING'"
        ).fetchall()

    if not pending:
        logger.info("未解析メールはありません")
        return 0

    total_items = 0
    skipped_count = 0

    for msg in pending:
        msg_id = msg["id"]
        subject = msg["subject"] or ""
        sender = msg["sender"] or ""

        # 受注メールフィルタ: 注文と無関係なメールはスキップ
        if not _is_order_mail(sender, subject):
            _update_parse_status(msg_id, ParseStatus.SKIPPED, "注文メール以外")
            skipped_count += 1
            continue

        try:
            # HTML → テキスト変換（HTMLがなければraw_textを使う）
            text = _html_to_text(msg["raw_html"]) if msg["raw_html"] else (msg["raw_text"] or "")
            if not text:
                _update_parse_status(msg_id, ParseStatus.ERROR, "メール本文が空です")
                continue

            # 件名もテキストに含める（注文番号が件名にある場合がある）
            full_text = f"{subject}\n{text}"

            # 商品ブロックに分割
            blocks = _split_items(full_text, separator_patterns)

            items_saved = 0
            items_found = 0
            for block in blocks:
                item = _parse_single_item(block, mail_patterns, msg_id)
                if item is None:
                    continue
                items_found += 1

                # 注文番号が個別ブロックで取れなかった場合、全体テキストから再試行
                if not item["order_number"]:
                    patterns = mail_patterns.get("patterns", {})
                    item["order_number"] = _extract_field(
                        full_text, patterns.get("order_number", [])
                    ) or ""

                if _save_order_item(item):
                    items_saved += 1

            if items_saved == 0 and items_found == 0:
                # 分割せずに全体から1商品として再試行
                item = _parse_single_item(full_text, mail_patterns, msg_id)
                if item:
                    items_found += 1
                    if _save_order_item(item):
                        items_saved = 1

            if items_saved > 0:
                _update_parse_status(msg_id, ParseStatus.DONE)
                total_items += items_saved
                logger.info("メール解析完了 (id=%d, subject=%s): %d商品", msg_id, subject, items_saved)
            elif items_found > 0:
                _update_parse_status(msg_id, ParseStatus.DONE, "商品コード検出済み（重複のためスキップ）")
                logger.info("メール解析完了・重複スキップ (id=%d, subject=%s)", msg_id, subject)
            else:
                _update_parse_status(msg_id, ParseStatus.ERROR, "商品コードを検出できませんでした")
                logger.warning("商品コード未検出 (id=%d, subject=%s)", msg_id, subject)

        except Exception as e:
            logger.error("メール解析エラー (id=%d): %s", msg_id, e)
            _update_parse_status(msg_id, ParseStatus.ERROR, str(e))

    if skipped_count > 0:
        logger.info("注文メール以外をスキップ: %d件", skipped_count)
    logger.info("メール解析処理完了: %d件の商品を登録", total_items)
    return total_items


def _save_order_item(item: dict) -> bool:
    """order_itemsテーブルに1行保存する。

    以下の場合はスキップする:
    - 同一商品コードが既に登録済み（重複防止）
    - 商品名も金額もない不完全なデータ

    Returns:
        True: 新規登録成功, False: スキップ
    """
    code = item.get("product_code_normalized") or item.get("product_code_raw", "")

    # 不完全データの除外（商品名なし AND 金額0）
    if not item.get("product_name") and not item.get("amount"):
        logger.info("不完全データをスキップ: code=%s (商品名・金額なし)", code)
        return False

    order_number = item.get("order_number", "")

    with get_connection() as conn:
        # 同一注文番号＋商品コードの重複チェック
        if code and order_number:
            existing = conn.execute(
                "SELECT id FROM order_items WHERE product_code_normalized = ? AND order_number = ?",
                (code, order_number),
            ).fetchone()
            if existing:
                logger.info(
                    "重複スキップ: code=%s, order=%s (既存 order_item_id=%d)",
                    code, order_number, existing["id"],
                )
                return False
        elif code:
            # 注文番号なしの場合は同一メッセージ内の重複のみチェック
            existing = conn.execute(
                "SELECT id FROM order_items WHERE product_code_normalized = ? AND message_id = ?",
                (code, item["message_id"]),
            ).fetchone()
            if existing:
                logger.info(
                    "重複スキップ: code=%s, message_id=%d (既存 order_item_id=%d)",
                    code, item["message_id"], existing["id"],
                )
                return False

        conn.execute(
            """
            INSERT INTO order_items (
                message_id, order_number,
                product_code_raw, product_code_normalized,
                product_name, amount, quantity,
                item_condition, recipient_name,
                edition, pub_year, author, sku,
                order_date, ship_date, tax, shipping,
                current_status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'PENDING')
            """,
            (
                item["message_id"],
                item["order_number"],
                item["product_code_raw"],
                item["product_code_normalized"],
                item["product_name"],
                item["amount"],
                item["quantity"],
                item.get("item_condition", ""),
                item.get("recipient_name", ""),
                item.get("edition", ""),
                item.get("pub_year", ""),
                item.get("author", ""),
                item.get("sku", ""),
                item.get("order_date", ""),
                item.get("ship_date", ""),
                item.get("tax", 0),
                item.get("shipping", 0),
            ),
        )
        return True


def _update_parse_status(message_id: int, status: ParseStatus, error_msg: str = "") -> None:
    """messagesのparse_statusを更新する。"""
    with get_connection() as conn:
        conn.execute(
            "UPDATE messages SET parse_status = ?, error_message = ? WHERE id = ?",
            (status.value, error_msg, message_id),
        )
