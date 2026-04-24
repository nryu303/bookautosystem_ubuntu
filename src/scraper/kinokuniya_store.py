"""紀伊國屋実店舗 スクレイパー

KINOナビの店頭在庫検索を、検索フロー経由で実行する。
KINOナビは暗号化トークンURLを使用しており、直接URLを組み立てることができない。
トップページで店舗選択 → ISBN検索 → 商品ページ（暗号化URL）の流れで在庫を確認する。
"""

from src.common.enums import AvailabilityStatus
from src.common.logger import get_logger
from src.scraper.base_scraper import BaseScraper

logger = get_logger(__name__)

# 店舗コード → KINOナビ store パラメータ の対応
STORE_MAP: dict[str, str] = {
    "X000036": "G2",   # 新宿本店
    "X000037": "J2",   # 梅田本店
    "X000038": "JX",   # グランフロント大阪店
    "X000039": "FA",   # 札幌本店
}

KINONAVI_TOP_URL = "https://www.kinokuniya.co.jp/kinonavi/"


class KinokuniyaStoreScraper(BaseScraper):
    supplier_code = "kinokuniya_store"

    async def _extract_stock_status_text(self) -> str:
        """KINOナビ商品ページから「在庫状況：」の値テキストを抽出する。

        ページ上の表示例: 「在庫��況：× 在庫なし」「在庫状況：◎ 在庫あり��
        """
        try:
            text = await self._page.evaluate("""
                () => {
                    // 方法1: 「在庫状況」を含むテキストノードの周辺を探す
                    const walker = document.createTreeWalker(
                        document.body, NodeFilter.SHOW_TEXT, null, false
                    );
                    let node;
                    while (node = walker.nextNode()) {
                        const t = node.textContent.trim();
                        if (t.includes('在庫状況')) {
                            // 「在庫状況：× 在庫なし」のように同じノードに含まれる場合
                            const match = t.match(/在庫状況[：:]\\s*(.+)/);
                            if (match) return match[1].trim();
                            // 次の兄弟テキストノードに値がある場合
                            const parent = node.parentElement;
                            if (parent) {
                                const fullText = parent.textContent.trim();
                                const m2 = fullText.match(/在庫状況[：:]\\s*(.+)/);
                                if (m2) return m2[1].trim();
                            }
                        }
                    }
                    return '';
                }
            """)
            if text:
                logger.debug("KINOナビ在庫テキスト抽出: %s", text)
            return text or ""
        except Exception as e:
            logger.debug("KINOナビ在庫テキスト抽出失敗: %s", e)
            return ""

    async def scrape_item(self, order_item_id: int, product_code: str) -> AvailabilityStatus:
        """KINOナビの検索フローを経由して店頭在庫を確認する。

        フロー:
          1. KINOナビ トップページにアクセス
          2. JS で店舗セレクトを設定し、検索ボックスに ISBN を入力
          3. フォーム送信 → 暗号化トークン付きの商品ページにリダイレクト
          4. 在庫テキストを抽出・判定
        """
        supplier_info = self._get_supplier_info()
        if supplier_info is None:
            return AvailabilityStatus.ERROR

        supplier_id = supplier_info["id"]
        store_code = STORE_MAP.get(self.supplier_code, "")

        if not store_code:
            logger.warning("KINOナビ店舗コード未定義: %s", self.supplier_code)
            self._save_scrape_result(
                order_item_id, supplier_id,
                AvailabilityStatus.ERROR, "", "", "",
                error_flag=True, error_message="KINOナビ店舗コード未定義",
            )
            return AvailabilityStatus.ERROR

        logger.info(
            "KINOナビ検索開始: %s (store=%s), code=%s",
            self.supplier_code, store_code, product_code,
        )

        try:
            await self.open_browser()

            # Step 1: KINOナビのトップページにアクセス
            if not await self.navigate(KINONAVI_TOP_URL):
                self._save_scrape_result(
                    order_item_id, supplier_id,
                    AvailabilityStatus.ERROR, "", "", "",
                    error_flag=True, error_message="KINOナビトップページアクセス失敗",
                )
                return AvailabilityStatus.ERROR

            # Step 2: JSで店舗を選択（カスタムセレクトウィジェットのため）
            await self._page.evaluate(
                """(storeCode) => {
                    const sel = document.querySelector('select[name="store"]');
                    if (sel) {
                        sel.value = storeCode;
                        sel.dispatchEvent(new Event('change', {bubbles: true}));
                    }
                }""",
                store_code,
            )

            # Step 3: 検索ボックスにISBNを入力
            keyword_input = await self._page.query_selector('input#keyword')
            if not keyword_input:
                keyword_input = await self._page.query_selector('input[name="q"]')
            if not keyword_input:
                logger.error("KINOナビ検索ボックスが見つかりません")
                self._save_scrape_result(
                    order_item_id, supplier_id,
                    AvailabilityStatus.ERROR, "", "", "",
                    error_flag=True, error_message="KINOナビ検索ボックス不在",
                )
                return AvailabilityStatus.ERROR

            await keyword_input.fill(product_code)
            await self._page.wait_for_timeout(500)

            # Step 4: フォーム送信（暗号化トークンURL付きの商品ページにリダイレクトされる）
            await self._page.evaluate(
                """() => {
                    const form = document.getElementById('search_form');
                    if (form) form.submit();
                }"""
            )
            await self._page.wait_for_timeout(4000)

            # Step 5: ページ内容確認 — エラーページかどうかチェック
            page_text = await self._page.evaluate(
                "() => document.body ? document.body.innerText : ''"
            )

            if "エラーが発生しました" in page_text and "再度トップページより操作してください" in page_text:
                logger.warning(
                    "KINOナビ エラーページ検出 (store=%s, code=%s)",
                    store_code, product_code,
                )
                html_path = await self.save_html(self.supplier_code, product_code)
                screenshot_path = await self.save_screenshot(self.supplier_code, product_code)
                self._save_scrape_result(
                    order_item_id, supplier_id,
                    AvailabilityStatus.ERROR, "", html_path, screenshot_path,
                    error_flag=True,
                    error_message="KINOナビ エラーページ",
                )
                return AvailabilityStatus.ERROR

            # Step 6: HTML・スクリーンショット保存
            html_path = await self.save_html(self.supplier_code, product_code)
            screenshot_path = await self.save_screenshot(self.supplier_code, product_code)

            # Step 7: 在��判定
            # KINOナビの在庫状況は「在庫状況：◎ 在庫あり」のような形式で表示される。
            # ページ全体のテキストで判定すると「在庫あり検索」等の UI テキストに
            # 誤マッチするため、在庫状況テキストを直接抽出して判定する。
            stock_text = await self._extract_stock_status_text()
            positive, negative = self._get_supplier_patterns()

            if stock_text:
                status, matched_text = self.judge_availability(
                    [{"text": stock_text}], positive, negative,
                )
            else:
                candidates = await self.extract_text_candidates()
                status, matched_text = self.judge_availability(candidates, positive, negative)

            self._save_scrape_result(
                order_item_id, supplier_id,
                status, matched_text,
                html_path, screenshot_path,
            )

            logger.info(
                "KINOナビ検索完了: %s / code=%s → %s (match=%s)",
                self.supplier_code, product_code, status.value, matched_text,
            )
            return status

        except Exception as e:
            logger.error(
                "KINOナビ スクレイピングエラー (%s, %s): %s",
                self.supplier_code, product_code, e,
            )
            self._save_scrape_result(
                order_item_id, supplier_id,
                AvailabilityStatus.ERROR, "", "", "",
                error_flag=True, error_message=str(e),
            )
            return AvailabilityStatus.ERROR

        finally:
            await self.close_browser()
