"""全商品×全有効仕入先を強制再スクレイピングする一時スクリプト。

screenshot ファイル欠損を修復する目的で、150分スキップをバイパスして実行する。
"""

from src.common.config import load_settings
from src.common.database import init_db
from src.common.logger import setup_logging, get_logger
from src.common.config import get_setting

load_settings()
setup_logging(
    log_dir=get_setting("logging", "dir", default="logs"),
    level=get_setting("logging", "level", default="INFO"),
)
init_db()

logger = get_logger(__name__)

from src.scraper import orchestrator
orchestrator.has_recent_result = lambda *a, **kw: False

logger.info("=== 強制再スクレイピング開始（150分スキップ無効化・ロック待機あり） ===")
stats = orchestrator.run_all_scraping(blocking_lock=True)
logger.info("=== 強制再スクレイピング完了: %s ===", stats)
print(stats)
