"""ヨドバシ.com スクレイパー

HTTP/2エラー対策: ホームページ経由でアクセスし、リクエスト間隔を設ける。
"""

import random
from src.scraper.base_scraper import BaseScraper
from src.common.logger import get_logger

logger = get_logger(__name__)


class YodobashiScraper(BaseScraper):
    supplier_code = "X000001"

    async def navigate(self, url: str) -> bool:
        """ヨドバシ用のナビゲーション。ホームページ経由＋ランダム遅延。"""
        try:
            # ホームページにアクセス
            await self._page.goto(
                "https://www.yodobashi.com/",
                timeout=self._timeout_ms,
                wait_until="domcontentloaded",
            )
            await self._page.wait_for_timeout(random.randint(2000, 4000))

            # 検索ページにアクセス
            response = await self._page.goto(
                url, timeout=self._timeout_ms, wait_until="domcontentloaded"
            )
            if response and response.status >= 400:
                logger.warning("HTTP %d: %s", response.status, url)
                return False
            await self._page.wait_for_timeout(random.randint(2000, 4000))
            return True
        except Exception as e:
            logger.error("ページアクセスエラー (%s): %s", url, e)
            return False
