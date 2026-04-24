"""e-hon スクレイパー

403対策: ホームページを先に読み込んでセッション/Cookieを取得してから検索ページにアクセスする。
"""

from src.scraper.base_scraper import BaseScraper
from src.common.logger import get_logger

logger = get_logger(__name__)


class EhonScraper(BaseScraper):
    supplier_code = "X000004"

    async def navigate(self, url: str) -> bool:
        """e-hon用のナビゲーション。ホームページ経由でCookieを取得する。"""
        try:
            # まずホームページにアクセスしてCookieを取得
            await self._page.goto(
                "https://www.e-hon.ne.jp/bec/EB/Top",
                timeout=self._timeout_ms,
                wait_until="domcontentloaded",
            )
            await self._page.wait_for_timeout(2000)

            # 検索ページにアクセス
            response = await self._page.goto(
                url, timeout=self._timeout_ms, wait_until="domcontentloaded"
            )
            if response and response.status >= 400:
                logger.warning("HTTP %d: %s", response.status, url)
                return False
            await self._page.wait_for_timeout(2000)
            return True
        except Exception as e:
            logger.error("ページアクセスエラー (%s): %s", url, e)
            return False
