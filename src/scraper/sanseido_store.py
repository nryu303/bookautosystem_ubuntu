"""三省堂実店舗 スクレイパー"""

from src.scraper.base_scraper import BaseScraper


class SanseidoStoreScraper(BaseScraper):
    supplier_code = "sanseido_store"
