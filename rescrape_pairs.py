"""指定 (order_item_id, supplier_id) ペアだけを再スクレイピングする。"""

from src.common.config import load_settings, get_setting
from src.common.database import init_db, get_connection
from src.common.logger import setup_logging, get_logger
from src.scraper import get_scraper
from src.scraper.base_scraper import run_scraper_sync
from src.scraper.scrape_lock import scraping_lock

load_settings()
setup_logging(log_dir=get_setting("logging", "dir", default="logs"), level="INFO")
init_db()

logger = get_logger(__name__)

PAIRS = [(21, 2), (27, 2)]  # 0V-GUOI-I191 × 楽天ブックス


def resolve(pair_list):
    with get_connection() as conn:
        for item_id, sup_id in pair_list:
            item = conn.execute(
                "SELECT product_code_normalized FROM order_items WHERE id=?",
                (item_id,),
            ).fetchone()
            sup = conn.execute(
                "SELECT supplier_code, supplier_name FROM suppliers WHERE id=?",
                (sup_id,),
            ).fetchone()
            if not item or not sup:
                logger.warning("skip unresolved pair: item=%s sup=%s", item_id, sup_id)
                continue
            yield item_id, item["product_code_normalized"], sup["supplier_code"], sup["supplier_name"]


with scraping_lock(blocking=True):
    for item_id, code, sup_code, sup_name in resolve(PAIRS):
        scraper = get_scraper(sup_code)
        if scraper is None:
            logger.warning("scraper not registered: %s", sup_code)
            continue
        logger.info("rescrape: item=%d code=%s sup=%s(%s)", item_id, code, sup_code, sup_name)
        try:
            status = run_scraper_sync(scraper, item_id, code)
            logger.info("  -> %s", status.value)
        except Exception as e:
            logger.error("  -> exception: %s", e)

print("done")
