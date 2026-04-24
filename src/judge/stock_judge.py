"""在庫判定ロジック

スクレイピング結果をもとに、各商品×各仕入先の在庫有無を正式に判定する。
"""

from src.common.database import get_connection
from src.common.enums import AvailabilityStatus
from src.common.logger import get_logger

logger = get_logger(__name__)


def get_latest_results_for_item(order_item_id: int) -> list[dict]:
    """指定商品の各仕入先の最新スクレイピング結果を取得する。

    各仕入先につき最新1件のみを返す。

    Returns:
        [{"supplier_id": 1, "supplier_code": "...", "availability_status": "...", ...}, ...]
    """
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT sr.*, s.supplier_code, s.supplier_name, s.category, s.priority
            FROM scrape_results sr
            JOIN suppliers s ON sr.supplier_id = s.id
            WHERE sr.order_item_id = ?
              AND sr.id = (
                  SELECT MAX(sr2.id)
                  FROM scrape_results sr2
                  WHERE sr2.order_item_id = sr.order_item_id
                    AND sr2.supplier_id = sr.supplier_id
              )
            ORDER BY s.priority ASC
            """,
            (order_item_id,),
        ).fetchall()
    return [dict(row) for row in rows]


def get_available_suppliers(order_item_id: int) -> list[dict]:
    """在庫ありの仕入先のみをpriority順で返す。"""
    results = get_latest_results_for_item(order_item_id)
    available = [
        r for r in results
        if r["availability_status"] == AvailabilityStatus.AVAILABLE.value
    ]
    return available


def summarize_stock_status(order_item_id: int) -> dict:
    """1商品の在庫状況サマリーを返す。

    Returns:
        {
            "order_item_id": 123,
            "total_checked": 5,
            "available_count": 2,
            "unavailable_count": 2,
            "unknown_count": 1,
            "error_count": 0,
            "available_suppliers": [...],
            "all_results": [...],
        }
    """
    results = get_latest_results_for_item(order_item_id)

    available = []
    unavailable_count = 0
    unknown_count = 0
    error_count = 0

    for r in results:
        status = r["availability_status"]
        if status == AvailabilityStatus.AVAILABLE.value:
            available.append(r)
        elif status == AvailabilityStatus.UNAVAILABLE.value:
            unavailable_count += 1
        elif status == AvailabilityStatus.ERROR.value:
            error_count += 1
        else:
            unknown_count += 1

    return {
        "order_item_id": order_item_id,
        "total_checked": len(results),
        "available_count": len(available),
        "unavailable_count": unavailable_count,
        "unknown_count": unknown_count,
        "error_count": error_count,
        "available_suppliers": available,
        "all_results": results,
    }
