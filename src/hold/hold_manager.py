"""保留バケット管理モジュール

仕入先ごとの保留合計金額・件数・経過日数を管理する。
閾値に達したバケットを検出する。
"""

from datetime import datetime

from src.common.database import get_connection
from src.common.enums import HoldBucketStatus, OrderStatus
from src.common.logger import get_logger

logger = get_logger(__name__)


def add_to_hold(order_item_id: int, supplier_id: int, amount: int) -> None:
    """商品を保留バケットに追加する。"""
    with get_connection() as conn:
        # hold_buckets が存在しなければ作成
        bucket = conn.execute(
            "SELECT id FROM hold_buckets WHERE supplier_id = ?",
            (supplier_id,),
        ).fetchone()

        if bucket is None:
            conn.execute(
                """
                INSERT INTO hold_buckets (supplier_id, total_amount, item_count, oldest_item_date, status)
                VALUES (?, 0, 0, NULL, 'ACTIVE')
                """,
                (supplier_id,),
            )
            bucket = conn.execute(
                "SELECT id FROM hold_buckets WHERE supplier_id = ?",
                (supplier_id,),
            ).fetchone()

        bucket_id = bucket["id"]

        # hold_items に追加
        conn.execute(
            """
            INSERT INTO hold_items (hold_bucket_id, order_item_id, amount, rescrape_required)
            VALUES (?, ?, ?, 0)
            """,
            (bucket_id, order_item_id, amount),
        )

    # バケット集計を再計算
    recalculate_bucket(supplier_id)

    logger.info(
        "保留追加: order_item_id=%d → supplier_id=%d (amount=%d)",
        order_item_id, supplier_id, amount,
    )


def recalculate_bucket(supplier_id: int) -> dict | None:
    """保留バケットの合計金額・件数・最古日付を再計算する。

    Returns:
        更新後のバケット情報
    """
    with get_connection() as conn:
        bucket = conn.execute(
            "SELECT id FROM hold_buckets WHERE supplier_id = ?",
            (supplier_id,),
        ).fetchone()

        if bucket is None:
            return None

        bucket_id = bucket["id"]

        # 集計
        stats = conn.execute(
            """
            SELECT
                COALESCE(SUM(amount), 0) AS total_amount,
                COUNT(*) AS item_count,
                MIN(assigned_at) AS oldest_item_date
            FROM hold_items
            WHERE hold_bucket_id = ?
            """,
            (bucket_id,),
        ).fetchone()

        total = stats["total_amount"]
        count = stats["item_count"]
        oldest = stats["oldest_item_date"]

        # 閾値チェック（金額 + 日数）
        sup = conn.execute(
            "SELECT hold_limit_amount, hold_limit_days FROM suppliers WHERE id = ?",
            (supplier_id,),
        ).fetchone()
        limit_amount = sup["hold_limit_amount"] if sup else 0
        limit_days = sup["hold_limit_days"] if sup else 0

        status = HoldBucketStatus.ACTIVE.value
        if count == 0:
            status = HoldBucketStatus.CLEARED.value
        elif limit_amount > 0 and total >= limit_amount:
            status = HoldBucketStatus.THRESHOLD_REACHED.value
        elif limit_days > 0 and oldest:
            try:
                oldest_dt = datetime.strptime(oldest, "%Y-%m-%d %H:%M:%S")
                if (datetime.now() - oldest_dt).days >= limit_days:
                    status = HoldBucketStatus.THRESHOLD_REACHED.value
            except (ValueError, TypeError):
                pass

        conn.execute(
            """
            UPDATE hold_buckets
            SET total_amount = ?, item_count = ?, oldest_item_date = ?, status = ?
            WHERE id = ?
            """,
            (total, count, oldest, status, bucket_id),
        )

        result = {
            "bucket_id": bucket_id,
            "supplier_id": supplier_id,
            "total_amount": total,
            "item_count": count,
            "oldest_item_date": oldest,
            "status": status,
        }

        if status == HoldBucketStatus.THRESHOLD_REACHED.value:
            logger.info(
                "保留閾値到達: supplier_id=%d, total=%d, limit=%d",
                supplier_id, total, limit_amount,
            )

        return result


def get_threshold_reached_buckets() -> list[dict]:
    """閾値に到達した保留バケットのリストを返す。"""
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT hb.*, s.supplier_code, s.supplier_name, s.hold_limit_amount, s.hold_limit_days
            FROM hold_buckets hb
            JOIN suppliers s ON hb.supplier_id = s.id
            WHERE hb.status = ?
            """,
            (HoldBucketStatus.THRESHOLD_REACHED.value,),
        ).fetchall()
    return [dict(row) for row in rows]


def get_expired_buckets() -> list[dict]:
    """経過日数が上限を超えた保留バケットを返す。"""
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT hb.*, s.supplier_code, s.supplier_name, s.hold_limit_amount, s.hold_limit_days
            FROM hold_buckets hb
            JOIN suppliers s ON hb.supplier_id = s.id
            WHERE hb.status = 'ACTIVE'
              AND hb.oldest_item_date IS NOT NULL
              AND s.hold_limit_days > 0
              AND julianday('now') - julianday(hb.oldest_item_date) >= s.hold_limit_days
            """,
        ).fetchall()
    return [dict(row) for row in rows]


def get_hold_items_for_bucket(bucket_id: int) -> list[dict]:
    """バケット内の保留商品リストを返す。"""
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT hi.*, oi.product_code_normalized, oi.product_name, oi.amount, oi.quantity
            FROM hold_items hi
            JOIN order_items oi ON hi.order_item_id = oi.id
            WHERE hi.hold_bucket_id = ?
            ORDER BY hi.assigned_at ASC
            """,
            (bucket_id,),
        ).fetchall()
    return [dict(row) for row in rows]


def recalculate_all_buckets() -> int:
    """全バケットのステータスを再計算する。

    設定変更（hold_limit_amount引き下げ等）や日数経過による
    ステータス変化を反映する。
    """
    with get_connection() as conn:
        buckets = conn.execute(
            "SELECT supplier_id FROM hold_buckets"
        ).fetchall()

    updated = 0
    for bucket in buckets:
        result = recalculate_bucket(bucket["supplier_id"])
        if result and result["status"] == HoldBucketStatus.THRESHOLD_REACHED.value:
            updated += 1

    if updated > 0:
        logger.info("バケット再計算: %d件が閾値到達", updated)
    return updated


def clear_bucket(bucket_id: int) -> None:
    """バケットをクリアする（発注確定後）。"""
    with get_connection() as conn:
        # hold_items の対応 order_items を ORDERED に更新
        items = conn.execute(
            "SELECT order_item_id FROM hold_items WHERE hold_bucket_id = ?",
            (bucket_id,),
        ).fetchall()

        for item in items:
            conn.execute(
                "UPDATE order_items SET current_status = ? WHERE id = ?",
                (OrderStatus.ORDERED.value, item["order_item_id"]),
            )

        # hold_items 削除
        conn.execute("DELETE FROM hold_items WHERE hold_bucket_id = ?", (bucket_id,))

        # バケット状態クリア
        conn.execute(
            """
            UPDATE hold_buckets
            SET total_amount = 0, item_count = 0, oldest_item_date = NULL, status = 'CLEARED'
            WHERE id = ?
            """,
            (bucket_id,),
        )

    logger.info("保留バケットクリア: bucket_id=%d", bucket_id)


def process_hold_assignments() -> int:
    """仕入先が決まったHOLD商品を保留バケットに振り分ける。

    order_items.current_status='HOLD' かつ hold_items に未登録の商品を処理する。

    Returns:
        処理件数
    """
    with get_connection() as conn:
        items = conn.execute(
            """
            SELECT oi.id, oi.planned_supplier_id, oi.amount
            FROM order_items oi
            WHERE oi.current_status = ?
              AND oi.planned_supplier_id IS NOT NULL
              AND oi.id NOT IN (SELECT order_item_id FROM hold_items)
            """,
            (OrderStatus.HOLD.value,),
        ).fetchall()

    count = 0
    for item in items:
        try:
            add_to_hold(item["id"], item["planned_supplier_id"], item["amount"])
            count += 1
        except Exception as e:
            logger.error("保留登録エラー (order_item_id=%d): %s", item["id"], e)

    if count > 0:
        logger.info("保留バケット振り分け完了: %d件", count)

    # 全バケットのステータスを再計算（設定変更・日数経過に対応）
    recalculate_all_buckets()

    return count
