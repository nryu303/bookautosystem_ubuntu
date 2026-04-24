"""再スクレイピングトリガーモジュール

保留金額が閾値に達した / 経過日数が上限を超えたバケットの商品を
再度スクレイピングして在庫を再確認する。
"""

import asyncio

from src.common.database import get_connection
from src.common.enums import AvailabilityStatus, OrderStatus
from src.common.logger import get_logger
from src.hold.hold_manager import (
    get_threshold_reached_buckets,
    get_expired_buckets,
    get_hold_items_for_bucket,
    recalculate_bucket,
)
from src.scraper import get_scraper
from src.scraper.scrape_lock import ScrapingLockBusy, scraping_lock

logger = get_logger(__name__)


async def _rescrape_item(order_item_id: int, product_code: str, supplier_id: int) -> AvailabilityStatus:
    """1商品を指定仕入先で再スクレイピングする。"""
    with get_connection() as conn:
        sup = conn.execute(
            "SELECT supplier_code FROM suppliers WHERE id = ?",
            (supplier_id,),
        ).fetchone()

    if sup is None:
        return AvailabilityStatus.ERROR

    scraper = get_scraper(sup["supplier_code"])
    if scraper is None:
        logger.warning("スクレイパー未登録: %s", sup["supplier_code"])
        return AvailabilityStatus.ERROR

    return await scraper.scrape_item(order_item_id, product_code)


def process_rescrape() -> dict:
    """閾値到達・期限切れバケットの再スクレイピングを実行する。

    Returns:
        {"checked": int, "still_available": int, "no_longer_available": int, "errors": int}
    """
    stats = {
        "checked": 0,
        "still_available": 0,
        "no_longer_available": 0,
        "errors": 0,
        "lock_busy": False,
    }

    # 閾値到達バケット + 期限切れバケット
    buckets = get_threshold_reached_buckets() + get_expired_buckets()

    # 重複除去
    seen_ids = set()
    unique_buckets = []
    for b in buckets:
        if b["id"] not in seen_ids:
            seen_ids.add(b["id"])
            unique_buckets.append(b)

    if not unique_buckets:
        logger.info("再スクレイピング対象のバケットはありません")
        return stats

    lock_cm = scraping_lock(blocking=False)
    try:
        lock_cm.__enter__()
    except ScrapingLockBusy as e:
        logger.warning("再スクレイピングを多重起動回避のためスキップ: %s", e)
        stats["lock_busy"] = True
        return stats

    try:
        _run_rescrape_loop(unique_buckets, stats)
    finally:
        lock_cm.__exit__(None, None, None)

    logger.info(
        "再スクレイピング完了: checked=%d, available=%d, unavailable=%d, errors=%d",
        stats["checked"], stats["still_available"],
        stats["no_longer_available"], stats["errors"],
    )
    return stats


def _run_rescrape_loop(unique_buckets: list, stats: dict) -> None:
    """ロック取得済み状態で再スクレイピングループを実行する。"""
    for bucket in unique_buckets:
        items = get_hold_items_for_bucket(bucket["id"])

        logger.info(
            "再スクレイピング開始: %s (%d件)",
            bucket["supplier_name"], len(items),
        )

        for item in items:
            stats["checked"] += 1
            order_item_id = item["order_item_id"]
            product_code = item["product_code_normalized"]

            try:
                result = asyncio.run(
                    _rescrape_item(order_item_id, product_code, bucket["supplier_id"])
                )

                if result == AvailabilityStatus.AVAILABLE:
                    stats["still_available"] += 1
                elif result == AvailabilityStatus.ERROR:
                    stats["errors"] += 1
                else:
                    stats["no_longer_available"] += 1
                    # 在庫なし → 保留から外してPENDINGに戻す（再振り分け対象にする）
                    with get_connection() as conn:
                        conn.execute(
                            "DELETE FROM hold_items WHERE id = ?",
                            (item["id"],),
                        )
                        conn.execute(
                            """
                            UPDATE order_items
                            SET current_status = ?, planned_supplier_id = NULL, assigned_at = NULL
                            WHERE id = ?
                            """,
                            (OrderStatus.PENDING.value, order_item_id),
                        )
                    logger.info(
                        "在庫なし → PENDING復帰: order_item_id=%d", order_item_id,
                    )

            except Exception as e:
                logger.error(
                    "再スクレイピングエラー (item_id=%d, code=%s): %s",
                    order_item_id, product_code, e,
                )
                stats["errors"] += 1

        # バケット再計算
        recalculate_bucket(bucket["supplier_id"])

        # 閾値到達バケットは自動クリアせず、THRESHOLD_REACHED のまま待機。
        # 管理画面の「発注確定」ボタンで手動クリアする運用。
