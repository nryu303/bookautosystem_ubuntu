"""スクレイピング共通基盤モジュール

全サイト共通のブラウザ操作、HTML保存、スクリーンショット保存、
候補テキスト抽出、在庫パターンマッチを提供する。
各サイト固有のスクレイパーはこのクラスを継承する。
"""

import asyncio
import os
from datetime import datetime

from src.common.config import get_setting, BASE_DIR
from src.common.database import get_connection
from src.common.enums import AvailabilityStatus
from src.common.exceptions import ScrapeError
from src.common.logger import get_logger
from src.scraper.url_builder import build_url

logger = get_logger(__name__)


class BaseScraper:
    """全サイト共通のスクレイパー基底クラス。

    使い方:
        class RakutenBooksScraper(BaseScraper):
            supplier_code = "rakuten_books"
            # 必要に応じてメソッドをオーバーライド
    """

    supplier_code: str = ""  # サブクラスで設定

    def __init__(self):
        self._browser = None
        self._context = None
        self._page = None
        self._timeout_ms = get_setting("scraping", "timeout_sec", default=30) * 1000
        self._headless = get_setting("scraping", "headless", default=True)
        self._screenshots_dir = os.path.join(
            BASE_DIR, get_setting("storage", "screenshots", default="data/screenshots")
        )
        self._html_dir = os.path.join(
            BASE_DIR, get_setting("storage", "html_snapshots", default="data/html_snapshots")
        )
        os.makedirs(self._screenshots_dir, exist_ok=True)
        os.makedirs(self._html_dir, exist_ok=True)

    async def open_browser(self):
        """Playwrightブラウザを起動する。"""
        import sys
        from playwright.async_api import async_playwright

        # PyInstaller exe 環境では同梱ブラウザを使う
        if getattr(sys, "frozen", False):
            browser_dir = os.path.join(os.path.dirname(sys.executable), "playwright-browsers")
            os.environ["PLAYWRIGHT_BROWSERS_PATH"] = browser_dir

        self._playwright = await async_playwright().start()
        self._browser = await self._playwright.chromium.launch(
            headless=self._headless,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
            ],
        )
        self._context = await self._browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/131.0.0.0 Safari/537.36"
            ),
            locale="ja-JP",
            extra_http_headers={
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
                "Accept-Language": "ja,en-US;q=0.7,en;q=0.3",
                "Accept-Encoding": "gzip, deflate, br",
                "Sec-Fetch-Dest": "document",
                "Sec-Fetch-Mode": "navigate",
                "Sec-Fetch-Site": "none",
                "Sec-Fetch-User": "?1",
            },
        )
        self._page = await self._context.new_page()

        # Stealth対策: Bot検出を回避
        try:
            from playwright_stealth import stealth_async
            await stealth_async(self._page)
        except ImportError:
            pass

    async def close_browser(self):
        """ブラウザを閉じる。"""
        if self._page:
            await self._page.close()
        if self._context:
            await self._context.close()
        if self._browser:
            await self._browser.close()
        if hasattr(self, "_playwright") and self._playwright:
            await self._playwright.stop()
        self._page = None
        self._context = None
        self._browser = None

    async def navigate(self, url: str) -> bool:
        """URLにアクセスしてページ読込完了を待つ。

        Returns:
            True: 正常にアクセスできた
            False: アクセス失敗
        """
        try:
            response = await self._page.goto(url, timeout=self._timeout_ms, wait_until="domcontentloaded")
            if response and response.status >= 400:
                logger.warning("HTTP %d: %s", response.status, url)
                return False
            # 追加の描画待ち（JS遅延レンダリング対策）
            await self._page.wait_for_timeout(2000)
            return True
        except Exception as e:
            logger.error("ページアクセスエラー (%s): %s", url, e)
            return False

    async def save_html(self, supplier_code: str, product_code: str) -> str:
        """現在のページHTMLを保存して相対パスを返す。"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{supplier_code}_{product_code}_{timestamp}.html"
        filepath = os.path.join(self._html_dir, filename)
        try:
            html = await self._page.content()
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(html)
            # DBには BASE_DIR からの相対パスを保存
            return os.path.relpath(filepath, BASE_DIR)
        except Exception as e:
            logger.error("HTML保存エラー: %s", e)
            return ""

    async def save_screenshot(self, supplier_code: str, product_code: str) -> str:
        """現在のページのスクリーンショットを保存して相対パスを返す。"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{supplier_code}_{product_code}_{timestamp}.png"
        filepath = os.path.join(self._screenshots_dir, filename)
        try:
            await self._page.screenshot(path=filepath, full_page=True)
            # DBには BASE_DIR からの相対パスを保存
            return os.path.relpath(filepath, BASE_DIR)
        except Exception as e:
            logger.error("スクリーンショット保存エラー: %s", e)
            return ""

    async def extract_text_candidates(self) -> list[dict]:
        """ページ内の全テキストノードを抽出し、属性付きの候補リストを返す。

        Returns:
            [{"text": "在庫あり", "selector": "div.stock", "tag": "span", ...}, ...]
        """
        try:
            candidates = await self._page.evaluate("""
                () => {
                    const results = [];
                    const walker = document.createTreeWalker(
                        document.body,
                        NodeFilter.SHOW_TEXT,
                        null,
                        false
                    );
                    let node;
                    while (node = walker.nextNode()) {
                        const text = node.textContent.trim();
                        if (!text || text.length > 500) continue;
                        const parent = node.parentElement;
                        if (!parent) continue;
                        const tag = parent.tagName.toLowerCase();
                        if (['script', 'style', 'noscript'].includes(tag)) continue;
                        results.push({
                            text: text,
                            tag: tag,
                            id: parent.id || '',
                            className: parent.className || '',
                            selector: parent.tagName + (parent.id ? '#' + parent.id : '') +
                                     (parent.className ? '.' + String(parent.className).split(' ').filter(c=>c).join('.') : ''),
                        });
                    }
                    return results;
                }
            """)
            return candidates or []
        except Exception as e:
            logger.error("テキスト候補抽出エラー: %s", e)
            return []

    def judge_availability(
        self,
        candidates: list[dict],
        positive_patterns: list[str],
        negative_patterns: list[str],
    ) -> tuple[AvailabilityStatus, str]:
        """候補テキストから在庫状態を判定する。

        Args:
            candidates: extract_text_candidatesの戻り値
            positive_patterns: 在庫ありとみなすパターン
            negative_patterns: 在庫なしとみなすパターン

        Returns:
            (AvailabilityStatus, マッチしたテキスト)
        """
        all_text = " ".join(c.get("text", "") for c in candidates)

        # positive チェック（在庫あり）
        for pattern in positive_patterns:
            if pattern and pattern in all_text:
                return AvailabilityStatus.AVAILABLE, pattern

        # negative チェック（在庫なし）
        for pattern in negative_patterns:
            if pattern and pattern in all_text:
                return AvailabilityStatus.UNAVAILABLE, pattern

        return AvailabilityStatus.UNKNOWN, ""

    def _get_supplier_info(self) -> dict | None:
        """DBから自分の仕入先情報を取得する。"""
        with get_connection() as conn:
            row = conn.execute(
                "SELECT * FROM suppliers WHERE supplier_code = ?",
                (self.supplier_code,),
            ).fetchone()
        if row is None:
            logger.error("仕入先未登録: %s", self.supplier_code)
            return None
        return dict(row)

    def _get_supplier_patterns(self) -> tuple[list[str], list[str]]:
        """suppliers.yamlから在庫パターンを取得する。

        Returns:
            (positive_patterns, negative_patterns)
        """
        from src.common.config import load_suppliers
        suppliers = load_suppliers()
        for sup in suppliers:
            if sup.get("code") == self.supplier_code:
                return (
                    sup.get("positive_patterns", []),
                    sup.get("negative_patterns", []),
                )
        return [], []

    def _save_scrape_result(
        self,
        order_item_id: int,
        supplier_id: int,
        status: AvailabilityStatus,
        raw_text: str,
        html_path: str,
        screenshot_path: str,
        error_flag: bool = False,
        error_message: str = "",
    ) -> None:
        """scrape_resultsテーブルに結果を保存する。"""
        # 現在のページURLを取得（HTMLプレビューでの相対パス解決に使用）
        page_url = ""
        if self._page:
            try:
                page_url = self._page.url or ""
            except Exception:
                pass

        with get_connection() as conn:
            conn.execute(
                """
                INSERT INTO scrape_results (
                    order_item_id, supplier_id,
                    availability_status, raw_stock_text,
                    html_snapshot_path, screenshot_path,
                    error_flag, error_message, page_url
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    order_item_id, supplier_id,
                    status.value, raw_text,
                    html_path, screenshot_path,
                    int(error_flag), error_message,
                    page_url,
                ),
            )

    async def scrape_item(self, order_item_id: int, product_code: str) -> AvailabilityStatus:
        """1商品を1サイトでスクレイピングする。

        Args:
            order_item_id: order_items.id
            product_code: 正規化済み商品コード

        Returns:
            判定結果
        """
        supplier_info = self._get_supplier_info()
        if supplier_info is None:
            return AvailabilityStatus.ERROR

        supplier_id = supplier_info["id"]

        # URL生成
        url = build_url(supplier_id, product_code)
        if not url:
            logger.warning(
                "URL生成不可: supplier=%s, code=%s",
                self.supplier_code, product_code,
            )
            self._save_scrape_result(
                order_item_id, supplier_id,
                AvailabilityStatus.ERROR, "", "", "",
                error_flag=True, error_message="URL生成不可",
            )
            return AvailabilityStatus.ERROR

        logger.info("スクレイピング開始: %s → %s", self.supplier_code, url)

        try:
            await self.open_browser()

            # ページアクセス
            if not await self.navigate(url):
                self._save_scrape_result(
                    order_item_id, supplier_id,
                    AvailabilityStatus.ERROR, "", "", "",
                    error_flag=True, error_message="ページアクセス失敗",
                )
                return AvailabilityStatus.ERROR

            # HTML・スクリーンショット保存
            html_path = await self.save_html(self.supplier_code, product_code)
            screenshot_path = await self.save_screenshot(self.supplier_code, product_code)

            # テキスト候補抽出
            candidates = await self.extract_text_candidates()

            # 在庫判定: extraction_patterns があればそちらを優先
            from src.scraper.pattern_engine import get_extraction_patterns, extract_by_patterns
            ep = get_extraction_patterns(supplier_id)
            positive, negative = self._get_supplier_patterns()

            if ep:
                extracted = await extract_by_patterns(self._page, ep)
                stock_text = extracted.get("stock_text", "")
                if stock_text:
                    status, matched_text = self.judge_availability(
                        [{"text": stock_text}], positive, negative,
                    )
                else:
                    # stock_text が抽出できなかった場合は全候補テキストでフォールバック
                    status, matched_text = self.judge_availability(candidates, positive, negative)
            else:
                status, matched_text = self.judge_availability(candidates, positive, negative)

            # 結果保存
            self._save_scrape_result(
                order_item_id, supplier_id,
                status, matched_text,
                html_path, screenshot_path,
            )

            logger.info(
                "スクレイピング完了: %s / code=%s → %s (match=%s)",
                self.supplier_code, product_code, status.value, matched_text,
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


def run_scraper_sync(scraper: BaseScraper, order_item_id: int, product_code: str) -> AvailabilityStatus:
    """同期的にスクレイパーを実行するヘルパー。"""
    return asyncio.run(scraper.scrape_item(order_item_id, product_code))
