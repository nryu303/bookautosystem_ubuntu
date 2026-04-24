"""楽天ブックス スクレイパー

検索結果ページ（/search?sitem={code}）から在庫状況を取得する。

ページのサイドバーに「在庫あり」「予約受付中」等のフィルター選択肢テキストが
常に存在するため、全テキストからのパターンマッチでは誤判定が発生する。
商品エリア（class="rbcomp__item-list__item__stock"）のテキストのみを
抽出して判定する。
"""

from src.common.enums import AvailabilityStatus
from src.common.logger import get_logger
from src.scraper.base_scraper import BaseScraper

logger = get_logger(__name__)

# 商品エリアの在庫ステータスを抽出するJavaScript
_JS_EXTRACT_STOCK = r"""() => {
    const stockEls = document.querySelectorAll('.rbcomp__item-list__item__stock');
    return Array.from(stockEls).map(el => el.textContent.trim());
}"""


class RakutenBooksScraper(BaseScraper):
    supplier_code = "rakuten_books"

    async def scrape_item(
        self, order_item_id: int, product_code: str
    ) -> AvailabilityStatus:
        from src.scraper.url_builder import build_url

        supplier_info = self._get_supplier_info()
        if supplier_info is None:
            return AvailabilityStatus.ERROR
        supplier_id = supplier_info["id"]

        url = build_url(supplier_id, product_code)
        if not url:
            self._save_scrape_result(
                order_item_id, supplier_id,
                AvailabilityStatus.ERROR, "", "", "",
                error_flag=True, error_message="URL生成不可",
            )
            return AvailabilityStatus.ERROR

        logger.info("スクレイピング開始: %s → %s", self.supplier_code, url)

        try:
            await self.open_browser()

            if not await self.navigate(url):
                self._save_scrape_result(
                    order_item_id, supplier_id,
                    AvailabilityStatus.ERROR, "", "", "",
                    error_flag=True, error_message="ページアクセス失敗",
                )
                return AvailabilityStatus.ERROR

            html_path = await self.save_html(self.supplier_code, product_code)
            screenshot_path = await self.save_screenshot(self.supplier_code, product_code)

            # 商品エリアのステータステキストのみを抽出
            stock_texts = await self._page.evaluate(_JS_EXTRACT_STOCK)

            positive, negative = self._get_supplier_patterns()

            if not stock_texts:
                # 商品が見つからない場合はフォールバック（全テキスト判定）
                candidates = await self.extract_text_candidates()
                status, matched = self.judge_availability(
                    candidates, positive, negative,
                )
            else:
                # 最初の検索結果（最も関連度の高いヒット）のみで判定する。
                # 全件を結合すると、関連商品の「在庫あり」が本命商品の
                # 「ご注文できない商品」より先に positive マッチしてしまう。
                first_stock = stock_texts[0]
                status, matched = self.judge_availability(
                    [{"text": first_stock}], positive, negative,
                )

            self._save_scrape_result(
                order_item_id, supplier_id,
                status, matched,
                html_path, screenshot_path,
            )

            logger.info(
                "スクレイピング完了: %s / code=%s → %s (match=%s)",
                self.supplier_code, product_code, status.value, matched,
            )
            return status

        except Exception as e:
            logger.error("スクレイピングエラー (%s, %s): %s", self.supplier_code, product_code, e)
            self._save_scrape_result(
                order_item_id, supplier_id,
                AvailabilityStatus.ERROR, "", "", "",
                error_flag=True, error_message=str(e),
            )
            return AvailabilityStatus.ERROR

        finally:
            await self.close_browser()
