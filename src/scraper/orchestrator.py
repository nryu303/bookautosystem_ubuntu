"""スクレイピング統合モジュール

PENDING状態の全商品を、有効な全仕入先でスクレイピングする。
サイトごとにエラーが発生してもスキップして次のサイト・商品に進む。
"""

from src.common.database import get_connection
from src.common.enums import OrderStatus
from src.common.logger import get_logger
from src.scraper import get_scraper
from src.scraper.base_scraper import run_scraper_sync
from src.scraper.scrape_lock import ScrapingLockBusy, scraping_lock

logger = get_logger(__name__)


def get_scrape_targets() -> list[dict]:
    """スクレイピング対象の商品一覧を取得する。

    CANCELLED を除く全ステータスの商品を対象とする。
    在庫あり（HOLD/ORDERED）でも常に最新の在庫状況を確認する。
    """
    with get_connection() as conn:
        items = conn.execute(
            """
            SELECT id, product_code_normalized, product_name
            FROM order_items
            WHERE current_status != 'CANCELLED'
              AND product_code_normalized IS NOT NULL
              AND product_code_normalized != ''
            ORDER BY id ASC
            """,
        ).fetchall()
    return [dict(i) for i in items]


def get_enabled_suppliers() -> list[dict]:
    """スクレイピングが有効な仕入先一覧をpriority順で返す。"""
    with get_connection() as conn:
        sups = conn.execute(
            """
            SELECT id, supplier_code, supplier_name
            FROM suppliers
            WHERE scrape_enabled = 1
            ORDER BY priority ASC
            """
        ).fetchall()
    return [dict(s) for s in sups]


def has_recent_result(order_item_id: int, supplier_id: int) -> bool:
    """直近2.5時間以内にスクレイピング済みかを確認する。

    2.5時間（150分）以内に同一商品×仕入先の結果があればスキップ。
    """
    with get_connection() as conn:
        row = conn.execute(
            """
            SELECT id FROM scrape_results
            WHERE order_item_id = ? AND supplier_id = ?
              AND scraped_at >= datetime('now', 'localtime', '-150 minutes')
            """,
            (order_item_id, supplier_id),
        ).fetchone()
    return row is not None


def run_all_scraping(*, blocking_lock: bool = False) -> dict:
    """全PENDING商品を全有効サイトでスクレイピングする。

    Playwrightの多重起動を避けるためファイルロックで単一プロセス化する。

    Args:
        blocking_lock: True なら他プロセスの解放を待つ。False（既定）なら
            ロック取得不可時にスキップして即座に返る（main.py向け）。

    Returns:
        {"total_items": int, "total_scraped": int, "errors": int, "skipped": int,
         "lock_busy": bool}
    """
    stats = {
        "total_items": 0,
        "total_scraped": 0,
        "errors": 0,
        "skipped": 0,
        "lock_busy": False,
    }

    try:
        with scraping_lock(blocking=blocking_lock):
            targets = get_scrape_targets()
            suppliers = get_enabled_suppliers()
            stats["total_items"] = len(targets)

            if not targets:
                logger.info("スクレイピング対象の商品がありません")
                return stats

            if not suppliers:
                logger.warning("有効なスクレイピング対象の仕入先がありません")
                return stats

            logger.info(
                "スクレイピング開始: %d商品 × %d仕入先",
                len(targets), len(suppliers),
            )

            # 未登録スクレイパーの警告は1回だけ出す
            warned_suppliers = set()

            for item in targets:
                item_id = item["id"]
                product_code = item["product_code_normalized"]

                for sup in suppliers:
                    # 今日既にこの組み合わせの結果があればスキップ
                    if has_recent_result(item_id, sup["id"]):
                        stats["skipped"] += 1
                        continue

                    scraper = get_scraper(sup["supplier_code"])
                    if scraper is None:
                        if sup["supplier_code"] not in warned_suppliers:
                            logger.warning("スクレイパー未登録: %s (%s)", sup["supplier_code"], sup["supplier_name"])
                            warned_suppliers.add(sup["supplier_code"])
                        stats["skipped"] += 1
                        continue

                    try:
                        run_scraper_sync(scraper, item_id, product_code)
                        stats["total_scraped"] += 1
                    except Exception as e:
                        logger.error(
                            "スクレイピングエラー (item=%d, supplier=%s): %s",
                            item_id, sup["supplier_code"], e,
                        )
                        stats["errors"] += 1

            logger.info(
                "スクレイピング完了: scraped=%d, errors=%d, skipped=%d",
                stats["total_scraped"], stats["errors"], stats["skipped"],
            )
            return stats
    except ScrapingLockBusy as e:
        logger.warning("スクレイピング多重起動を検出したためスキップします: %s", e)
        stats["lock_busy"] = True
        return stats
