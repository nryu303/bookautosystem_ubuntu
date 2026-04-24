"""仕入先振り分けロジック

クライアント仕様の3段階優先順位に基づいて発注先を決定する。

＜注文保留の考え方（EC/店頭共通）＞
  優先第1位: 保留中金額 < 最低積み上げライン AND 在庫あり → 最小保留合計金額
  優先第2位: 保留中金額 ≥ 最低積み上げライン AND 在庫あり → 最小保留合計金額
  優先第3位: 在庫あり AND 保留中金額 = 0（保留受注残なし） → 仕入先間優先順位順
"""

from datetime import datetime

from src.common.database import get_connection
from src.common.enums import OrderStatus
from src.common.logger import get_logger
from src.judge.stock_judge import get_available_suppliers

logger = get_logger(__name__)


def select_supplier(order_item_id: int) -> int | None:
    """1商品に対して3段階優先順位で最適な仕入先を選択し、order_itemsを更新する。

    Returns:
        選択された supplier_id。選択不可の場合は None。
    """
    available = get_available_suppliers(order_item_id)

    if not available:
        # 在庫なし → 処理台帳に「在庫無し扱い」
        with get_connection() as conn:
            conn.execute(
                "UPDATE order_items SET current_status = ? WHERE id = ?",
                ("NO_STOCK", order_item_id),
            )
        logger.info("在庫あり仕入先なし (order_item_id=%d) → NO_STOCK", order_item_id)
        return None

    # 各仕入先の保留状態を取得
    candidates = []
    with get_connection() as conn:
        for sup in available:
            supplier_id = sup["supplier_id"]
            bucket = conn.execute(
                "SELECT * FROM hold_buckets WHERE supplier_id = ?",
                (supplier_id,),
            ).fetchone()

            total_amount = bucket["total_amount"] if bucket else 0
            item_count = bucket["item_count"] if bucket else 0

            sup_row = conn.execute(
                "SELECT * FROM suppliers WHERE id = ?",
                (supplier_id,),
            ).fetchone()

            hold_limit = sup_row["hold_limit_amount"] if sup_row else 0
            mail_unit_price_limit = sup_row["mail_unit_price_limit"] if sup_row else 0

            candidates.append({
                "supplier_id": supplier_id,
                "supplier_code": sup["supplier_code"],
                "supplier_name": sup["supplier_name"],
                "category": sup["category"],
                "priority": sup["priority"],
                "total_amount": total_amount,
                "item_count": item_count,
                "hold_limit": hold_limit,
                "mail_unit_price_limit": mail_unit_price_limit,
            })

    if not candidates:
        return None

    # 商品金額を取得（制限単価チェック用）
    with get_connection() as conn:
        item_row = conn.execute(
            "SELECT amount FROM order_items WHERE id = ?",
            (order_item_id,),
        ).fetchone()
        item_amount = item_row["amount"] if item_row else 0

    # 制限単価チェック: 高額品を除外
    filtered = []
    for c in candidates:
        limit = c["mail_unit_price_limit"]
        if limit > 0 and item_amount > limit:
            continue
        filtered.append(c)

    if not filtered:
        filtered = candidates  # 全部除外されたらフィルタなしに戻す

    # 3段階優先順位で振り分け
    selected = _apply_three_tier_priority(filtered)

    if selected is None:
        selected = min(filtered, key=lambda c: c["priority"])

    # order_items を更新
    with get_connection() as conn:
        conn.execute(
            """
            UPDATE order_items
            SET planned_supplier_id = ?,
                current_status = ?,
                assigned_at = ?
            WHERE id = ?
            """,
            (
                selected["supplier_id"],
                OrderStatus.HOLD.value,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                order_item_id,
            ),
        )

    logger.info(
        "仕入先決定: order_item_id=%d → %s (tier=%s)",
        order_item_id, selected["supplier_name"], selected.get("_tier", "?"),
    )
    return selected["supplier_id"]


def _apply_three_tier_priority(candidates: list[dict]) -> dict | None:
    """クライアント仕様の3段階優先順位で仕入先を選択する。

    優先第1位: 保留中金額 < 最低積み上げライン(hold_limit) AND 在庫あり
              → その中で最小保留合計金額の仕入先
    優先第2位: 保留中金額 ≥ 最低積み上げライン AND 在庫あり
              → その中で最小保留合計金額の仕入先
    優先第3位: 在庫あり AND 保留中金額 = 0（保留受注残なし）
              → 仕入先間の優先順位順（priority値が小さい方が優先）
    """
    if not candidates:
        return None

    # 第1位: 保留金額 > 0 AND 保留金額 < 最低積み上げライン
    tier1 = [
        c for c in candidates
        if 0 < c["total_amount"] < c["hold_limit"]
    ]
    if tier1:
        selected = min(tier1, key=lambda c: (c["total_amount"], c["priority"]))
        selected["_tier"] = "1"
        return selected

    # 第2位: 保留金額 ≥ 最低積み上げライン
    tier2 = [
        c for c in candidates
        if c["total_amount"] >= c["hold_limit"] and c["hold_limit"] > 0
    ]
    if tier2:
        selected = min(tier2, key=lambda c: (c["total_amount"], c["priority"]))
        selected["_tier"] = "2"
        return selected

    # 第3位: 保留金額 = 0 → 仕入先間の優先順位順
    tier3 = [
        c for c in candidates
        if c["total_amount"] == 0
    ]
    if tier3:
        selected = min(tier3, key=lambda c: c["priority"])
        selected["_tier"] = "3"
        return selected

    return None


def assign_pending_items() -> int:
    """PENDING状態の全order_itemsに対して仕入先振り分けを実行する。

    全有効仕入先のスクレイピングが完了した商品のみ処理する。
    一部の仕入先しかスクレイピングされていない商品は次サイクルまで待機。

    Returns:
        振り分け成功件数
    """
    with get_connection() as conn:
        # スクレイパーが登録されている有効仕入先の数を取得
        # （スクレイパー未登録の仕入先はスクレイピング対象外）
        from src.scraper import SCRAPER_REGISTRY
        registered_codes = set(SCRAPER_REGISTRY.keys())
        enabled_suppliers = conn.execute(
            "SELECT supplier_code FROM suppliers WHERE scrape_enabled = 1"
        ).fetchall()
        scrape_target_count = sum(
            1 for s in enabled_suppliers if s["supplier_code"] in registered_codes
        )

        # スクレイピング結果が全対象仕入先分揃った商品のみ振り分け
        # （全仕入先の在庫状況を把握してから最適な振り分けを行う）
        # PENDING だけでなく NO_STOCK も対象に含める:
        # 一度 NO_STOCK 判定された注文でも、後から仕入先サイトに在庫が復活
        # （再スクレイプで AVAILABLE になる）することがあるため、再評価機会を与える。
        items = conn.execute(
            """
            SELECT oi.id, COUNT(DISTINCT sr.supplier_id) AS scraped_count
            FROM order_items oi
            JOIN scrape_results sr ON sr.order_item_id = oi.id
            WHERE oi.current_status IN (?, ?)
            GROUP BY oi.id
            HAVING scraped_count >= ?
            """,
            (
                OrderStatus.PENDING.value,
                OrderStatus.NO_STOCK.value,
                max(scrape_target_count, 1),
            ),
        ).fetchall()

    assigned_count = 0
    for item in items:
        try:
            result = select_supplier(item["id"])
            if result is not None:
                assigned_count += 1
        except Exception as e:
            logger.error("振り分けエラー (order_item_id=%d): %s", item["id"], e)

    logger.info("仕入先振り分け完了: %d件", assigned_count)
    return assigned_count
