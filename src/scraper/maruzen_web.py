"""丸善ジュンク堂ネットストア スクレイパー"""

from src.scraper.base_scraper import BaseScraper


class MaruzenWebScraper(BaseScraper):
    supplier_code = "maruzen_web"
