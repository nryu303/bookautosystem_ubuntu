"""紀伊國屋ウェブストア スクレイパー"""

from src.scraper.base_scraper import BaseScraper


class KinokuniyaEcScraper(BaseScraper):
    supplier_code = "X000003"
