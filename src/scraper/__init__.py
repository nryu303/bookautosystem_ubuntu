"""スクレイパーモジュール

supplier_code から対応するスクレイパーインスタンスを取得するレジストリを提供する。
"""

from src.scraper.base_scraper import BaseScraper
from src.scraper.rakuten_books import RakutenBooksScraper
from src.scraper.yodobashi import YodobashiScraper
from src.scraper.kinokuniya_ec import KinokuniyaEcScraper
from src.scraper.ehon import EhonScraper
from src.scraper.maruzen_web import MaruzenWebScraper
from src.scraper.furuhonya import FuruhonyaScraper
from src.scraper.maruzen_store import MaruzenStoreScraper
from src.scraper.kinokuniya_store import KinokuniyaStoreScraper
from src.scraper.sanseido_store import SanseidoStoreScraper

# supplier_code → スクレイパークラスの対応表
SCRAPER_REGISTRY: dict[str, type[BaseScraper]] = {
    "X000001": YodobashiScraper,
    "X000002": RakutenBooksScraper,
    "X000003": KinokuniyaEcScraper,
    "X000004": EhonScraper,
    "X000005": FuruhonyaScraper,
    "X000007": MaruzenWebScraper,
    "X000008": MaruzenStoreScraper,
    "X000009": MaruzenStoreScraper,
    "X000010": MaruzenStoreScraper,
    "X000011": MaruzenStoreScraper,
    "X000012": MaruzenStoreScraper,
    "X000013": MaruzenStoreScraper,
    "X000014": MaruzenStoreScraper,
    "X000015": MaruzenStoreScraper,
    "X000016": MaruzenStoreScraper,
    "X000017": MaruzenStoreScraper,
    "X000018": MaruzenStoreScraper,
    "X000019": MaruzenStoreScraper,
    "X000020": MaruzenStoreScraper,
    "X000021": MaruzenStoreScraper,
    "X000022": MaruzenStoreScraper,
    "X000023": MaruzenStoreScraper,
    "X000024": MaruzenStoreScraper,
    "X000025": MaruzenStoreScraper,
    "X000026": MaruzenStoreScraper,
    "X000027": MaruzenStoreScraper,
    "X000028": MaruzenStoreScraper,
    "X000029": MaruzenStoreScraper,
    "X000030": MaruzenStoreScraper,
    "X000031": MaruzenStoreScraper,
    "X000032": MaruzenStoreScraper,
    "X000033": MaruzenStoreScraper,
    "X000034": MaruzenStoreScraper,
    "X000035": MaruzenStoreScraper,
    "X000036": KinokuniyaStoreScraper,
    "X000037": KinokuniyaStoreScraper,
    "X000038": KinokuniyaStoreScraper,
    "X000039": KinokuniyaStoreScraper,
    "X000040": SanseidoStoreScraper,
    "X000041": SanseidoStoreScraper,
}


def get_scraper(supplier_code: str) -> BaseScraper | None:
    """supplier_code に対応するスクレイパーのインスタンスを返す。

    共有クラス（例: MaruzenStoreScraper）を複数の仕入先で使い回すため、
    インスタンス生成後にsupplier_codeを上書きする。

    ハードコードされたスクレイパーがない場合でも、extraction_patterns が
    登録されていれば汎用 BaseScraper で対応する。
    """
    cls = SCRAPER_REGISTRY.get(supplier_code)
    if cls is not None:
        instance = cls()
        instance.supplier_code = supplier_code
        return instance

    # フォールバック: extraction_patterns が登録されていれば汎用スクレイパーを使う
    from src.common.database import get_connection
    from src.scraper.pattern_engine import has_patterns

    with get_connection() as conn:
        row = conn.execute(
            "SELECT id FROM suppliers WHERE supplier_code = ?", (supplier_code,)
        ).fetchone()
    if row and has_patterns(row["id"]):
        instance = BaseScraper()
        instance.supplier_code = supplier_code
        return instance

    return None
