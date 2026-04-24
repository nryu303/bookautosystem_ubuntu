"""日本の古本屋 スクレイパー

日本の古本屋は他のECサイトと異なり、各店舗の在庫ごとに個別ページが存在する。
在庫判定:
  - 在庫あり: 商品詳細ページが正常に表示される（HTTP 200 + 商品情報あり）
  - 在庫なし: ページが削除されている（404、またはリダイレクト先で「該当なし」表示）

URL指定型CSV（機能②）で同一ISBNに複数出品がある場合:
  最安値の出品URLにアクセスし、在庫なし（ページ削除）なら次に安い出品へ
  自動的にフォールバックする。
"""

from src.common.database import get_connection
from src.common.enums import AvailabilityStatus
from src.common.logger import get_logger
from src.scraper.base_scraper import BaseScraper
from src.scraper.url_builder import build_url

logger = get_logger(__name__)

# ページ削除・在庫なしを示すテキストパターン
_NOT_FOUND_MARKERS = [
    "ご指定の商品は見つかりませんでした",
    "指定された商品は存在しません",
    "該当する商品がありません",
    "ページが見つかりません",
    "Not Found",
    "404",
]


class FuruhonyaScraper(BaseScraper):
    supplier_code = "X000005"

    async def _check_page_exists(self, url: str) -> tuple[bool, int]:
        """URLにアクセスしてページが存在するかを確認する。

        Returns:
            (page_exists: bool, http_status: int)
        """
        try:
            response = await self._page.goto(
                url, timeout=self._timeout_ms, wait_until="domcontentloaded"
            )
            status_code = response.status if response else 0

            # HTTP 404 等 → ページ削除（在庫なし）
            if status_code >= 400:
                return False, status_code

            # HTTP 200 でもリダイレクト先が検索結果ページの場合がある
            await self._page.wait_for_timeout(2000)
            current_url = self._page.url

            # 商品詳細ページ以外にリダイレクトされた場合
            if "detail.php" not in current_url and "product_id" not in current_url:
                return False, status_code

            # ページ内テキストに「見つからない」系のマーカーがないかチェック
            page_text = await self._page.evaluate("() => document.body.innerText")
            for marker in _NOT_FOUND_MARKERS:
                if marker in page_text:
                    return False, status_code

            return True, status_code

        except Exception as e:
            logger.error("ページアクセスエラー (%s): %s", url, e)
            return False, 0

    def _get_url_csv_candidates(self, product_code: str) -> list[dict]:
        """url_csv_entriesから価格昇順で全候補URLを取得する。

        Returns:
            [{"id": int, "item_url": str, "price": int, "item_text": str}, ...]
        """
        with get_connection() as conn:
            rows = conn.execute(
                """
                SELECT id, item_url, price, item_text
                FROM url_csv_entries
                WHERE product_code = ? AND item_url != ''
                ORDER BY price ASC
                """,
                (product_code,),
            ).fetchall()
        return [dict(r) for r in rows]

    async def scrape_item(self, order_item_id: int, product_code: str) -> AvailabilityStatus:
        """日本の古本屋の商品をスクレイピングする。

        ページの存在=在庫ありとして判定する。
        URL指定型CSVに複数候補がある場合、最安値から順に試行する。
        """
        supplier_info = self._get_supplier_info()
        if supplier_info is None:
            return AvailabilityStatus.ERROR

        supplier_id = supplier_info["id"]

        # URL指定型CSVから価格順の候補リストを取得
        candidates = self._get_url_csv_candidates(product_code)

        if not candidates:
            # CSVに候補がない場合、通常のURL生成を試みる
            url = build_url(supplier_id, product_code)
            if not url:
                logger.warning(
                    "URL生成不可（CSV候補なし）: supplier=%s, code=%s",
                    self.supplier_code, product_code,
                )
                self._save_scrape_result(
                    order_item_id, supplier_id,
                    AvailabilityStatus.UNAVAILABLE, "商品ページ無し",
                    "", "",
                )
                return AvailabilityStatus.UNAVAILABLE
            # 単一URLとして候補リストに追加
            candidates = [{"id": 0, "item_url": url, "price": 0, "item_text": ""}]

        logger.info(
            "日本の古本屋スクレイピング開始: code=%s, 候補%d件",
            product_code, len(candidates),
        )

        try:
            await self.open_browser()

            # 最安値から順にページ存在チェック
            for i, candidate in enumerate(candidates):
                item_url = candidate["item_url"]
                price = candidate["price"]

                logger.debug(
                    "候補 %d/%d: price=%d, url=%s",
                    i + 1, len(candidates), price, item_url,
                )

                page_exists, http_status = await self._check_page_exists(item_url)

                if page_exists:
                    # ページが存在する → 在庫あり
                    html_path = await self.save_html(self.supplier_code, product_code)
                    screenshot_path = await self.save_screenshot(self.supplier_code, product_code)

                    raw_text = f"在庫あり（ページ存在確認済み, {price}円）"
                    if candidate["item_text"]:
                        raw_text += f" / {candidate['item_text']}"

                    self._save_scrape_result(
                        order_item_id, supplier_id,
                        AvailabilityStatus.AVAILABLE, raw_text,
                        html_path, screenshot_path,
                    )

                    # 最安値のis_cheapestフラグを更新（この候補が実際の最安在庫）
                    if candidate["id"] > 0:
                        self._update_cheapest_flag(product_code, candidate["id"])

                    logger.info(
                        "日本の古本屋: code=%s → 在庫あり (候補%d/%d, %d円)",
                        product_code, i + 1, len(candidates), price,
                    )
                    return AvailabilityStatus.AVAILABLE

                else:
                    # ページ削除 → この候補は在庫なし、次の候補へ
                    logger.info(
                        "日本の古本屋: 候補%d/%d 在庫なし (HTTP %d, %d円) → 次の候補へ",
                        i + 1, len(candidates), http_status, price,
                    )
                    continue

            # 全候補が在庫なし
            html_path = await self.save_html(self.supplier_code, product_code)
            screenshot_path = await self.save_screenshot(self.supplier_code, product_code)

            self._save_scrape_result(
                order_item_id, supplier_id,
                AvailabilityStatus.UNAVAILABLE, "全候補ページ削除済み（在庫なし）",
                html_path, screenshot_path,
            )
            logger.info(
                "日本の古本屋: code=%s → 在庫なし（全%d候補チェック済み）",
                product_code, len(candidates),
            )
            return AvailabilityStatus.UNAVAILABLE

        except Exception as e:
            logger.error(
                "日本の古本屋スクレイピングエラー (%s): %s",
                product_code, e,
            )
            self._save_scrape_result(
                order_item_id, supplier_id,
                AvailabilityStatus.ERROR, "", "", "",
                error_flag=True, error_message=str(e),
            )
            return AvailabilityStatus.ERROR

        finally:
            await self.close_browser()

    def _update_cheapest_flag(self, product_code: str, available_entry_id: int) -> None:
        """在庫確認済みの候補にis_cheapestフラグを付け替える。"""
        with get_connection() as conn:
            conn.execute(
                "UPDATE url_csv_entries SET is_cheapest = 0 WHERE product_code = ?",
                (product_code,),
            )
            conn.execute(
                "UPDATE url_csv_entries SET is_cheapest = 1 WHERE id = ?",
                (available_entry_id,),
            )
