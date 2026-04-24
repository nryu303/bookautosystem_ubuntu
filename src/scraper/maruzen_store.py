"""丸善ジュンク堂 実店舗スクレイパー

店舗在庫ページ (shoplist?product={code}) は全店舗の在庫を
6ページにわたって表示する。「次へ」ボタン（changePage(N)）で
全ページをめくり、各店舗の shop_name と _status を取得する。

全28店舗が同一URLなので、最初の1回で全ページを取得し
クラスレベルのキャッシュに保持する。同一商品の残り27店舗は
キャッシュから結果を返す（ブラウザ不要）。

スクレイピングのたびに全店舗の在庫状況を xlsx ファイルとして
data/csv/ に出力する。
"""

import os
import time
from datetime import datetime

from src.common.config import BASE_DIR
from src.common.enums import AvailabilityStatus
from src.common.logger import get_logger
from src.scraper.base_scraper import BaseScraper
from src.scraper.url_builder import build_url

logger = get_logger(__name__)

# キャッシュの有効期限（秒）
_CACHE_TTL = 600  # 10分

# クラスレベルキャッシュ:
#   {product_code: (timestamp, [{name, status, address, screenshot_path}])}
_store_cache: dict[str, tuple[float, list[dict[str, str]]]] = {}

# ページ内のショップ名・ステータスを抽出するJavaScript
_JS_EXTRACT_SHOPS = r"""() => {
    const entries = document.querySelectorAll('li.shop_details');
    const shops = [];
    entries.forEach(li => {
        if (li.style.display === 'none') return;
        const nameEl = li.querySelector('.shop_name');
        const statusEl = li.querySelector('._status');
        const addrEl = li.querySelector('.shop_address');
        if (nameEl) {
            shops.push({
                name: nameEl.textContent.trim(),
                status: statusEl ? statusEl.textContent.trim() : '',
                address: addrEl ? addrEl.textContent.trim() : ''
            });
        }
    });
    return shops;
}"""

_JS_GET_MAX_PAGE = r"""() => {
    const el = document.querySelector('.last-page-num a');
    return el ? parseInt(el.textContent) : 1;
}"""


def _match_store_name(
    stores: list[dict[str, str]], supplier_name: str
) -> dict[str, str] | None:
    """suppliers.yaml の店舗名とページ上の店舗名をマッチングする。

    suppliers.yaml: "丸善ジュンク堂書店-札幌店"
    ページ上:       "MARUZEN＆ジュンク堂書店 札幌店"

    ハイフン以降の店舗所在名（"札幌店"等）で照合する。
    マッチした店舗の dict（name, status, screenshot_path 等）を返す。
    """
    if "-" in supplier_name:
        location = supplier_name.split("-", 1)[1]
    else:
        location = supplier_name

    for entry in stores:
        if location in entry["name"]:
            return entry
    return None


# def _export_xlsx(product_code: str, stores: list[dict[str, str]]) -> str:
#     """全店舗の在庫状況を xlsx ファイルに出力する。

#     出力先: C:/Users/Administrator/OneDrive/BookAutoSystem/maruzen_store/maruzen_stores_{product_code}_{timestamp}.xlsx
#     Returns:
#         出力ファイルパス（BASE_DIR 相対）
#     """
#     import openpyxl

#     out_dir = "C:/Users/Administrator/OneDrive/BookAutoSystem/maruzen_store"
#     os.makedirs(out_dir, exist_ok=True)
#     timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
#     filename = f"maruzen_stores_{product_code}_{timestamp}.xlsx"
#     filepath = os.path.join(out_dir, filename)

#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = "店舗在庫"

#     # ヘッダー
#     ws.append(["No", "店舗名", "住所", "在庫状況", "商品コード", "取得日時"])

#     now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
#     for i, entry in enumerate(stores, 1):
#         ws.append([
#             i,
#             entry["name"],
#             entry.get("address", ""),
#             entry["status"],
#             product_code,
#             now_str,
#         ])

#     # 列幅調整
#     ws.column_dimensions["A"].width = 5
#     ws.column_dimensions["B"].width = 40
#     ws.column_dimensions["C"].width = 60
#     ws.column_dimensions["D"].width = 12
#     ws.column_dimensions["E"].width = 18
#     ws.column_dimensions["F"].width = 22

#     wb.save(filepath)
#     rel_path = os.path.relpath(filepath, BASE_DIR)
#     logger.info("丸善店舗在庫xlsx出力: %s (%d店舗)", rel_path, len(stores))
#     return rel_path


class MaruzenStoreScraper(BaseScraper):
    supplier_code = "maruzen_store"

    async def _scrape_all_pages(self, product_code: str) -> list[dict[str, str]]:
        """全ページをめくり、各ページのスクリーンショットを撮影する。

        Returns:
            [{name, status, address, screenshot_path}, ...]
            screenshot_path は該当店舗が表示されていたページのスクリーンショット。
        """
        all_stores: list[dict[str, str]] = []

        # li.shop_details が実際にDOMに現れるまで明示的に待つ。
        # 固定 wait_for_timeout だけでは JS 描画が間に合わないケースがあり、
        # その場合この後の evaluate が空配列を返してしまう（全店舗マッチ失敗）。
        try:
            await self._page.wait_for_selector(
                "li.shop_details", state="attached", timeout=10000,
            )
        except Exception as e:
            logger.warning("丸善店舗一覧 li.shop_details 待機タイムアウト: %s", e)

        max_page = await self._page.evaluate(_JS_GET_MAX_PAGE)
        max_page = min(max_page, 10)  # 安全上限

        for page_num in range(1, max_page + 1):
            if page_num > 1:
                await self._page.evaluate(f"changePage({page_num})")
                # ページ切り替え後も shop_details の描画を待つ
                try:
                    await self._page.wait_for_selector(
                        "li.shop_details", state="attached", timeout=10000,
                    )
                except Exception:
                    pass
                await self._page.wait_for_timeout(800)

            # このページのスクリーンショットを保存
            ss_path = await self.save_screenshot(
                f"maruzen_store_p{page_num}", product_code,
            )

            page_shops = await self._page.evaluate(_JS_EXTRACT_SHOPS)
            for shop in page_shops:
                shop["screenshot_path"] = ss_path
            all_stores.extend(page_shops)

            logger.debug(
                "丸善店舗一覧 page %d/%d: %d店舗取得, ss=%s",
                page_num, max_page, len(page_shops), ss_path,
            )

        return all_stores

    async def scrape_item(
        self, order_item_id: int, product_code: str
    ) -> AvailabilityStatus:
        supplier_info = self._get_supplier_info()
        if supplier_info is None:
            return AvailabilityStatus.ERROR
        supplier_id = supplier_info["id"]
        supplier_name = supplier_info["supplier_name"]

        # キャッシュ確認
        cached = _store_cache.get(product_code)
        if cached and (time.time() - cached[0]) < _CACHE_TTL:
            stores = cached[1]
            logger.info(
                "丸善店舗キャッシュヒット: code=%s (%d店舗)",
                product_code, len(stores),
            )
        else:
            # ブラウザで全ページをスクレイピング
            url = build_url(supplier_id, product_code)
            if not url:
                self._save_scrape_result(
                    order_item_id, supplier_id,
                    AvailabilityStatus.ERROR, "", "", "",
                    error_flag=True, error_message="URL生成不可",
                )
                return AvailabilityStatus.ERROR

            logger.info("丸善店舗一覧スクレイピング開始: %s", url)

            try:
                await self.open_browser()
                if not await self.navigate(url):
                    self._save_scrape_result(
                        order_item_id, supplier_id,
                        AvailabilityStatus.ERROR, "", "", "",
                        error_flag=True, error_message="ページアクセス失敗",
                    )
                    return AvailabilityStatus.ERROR

                # 追加の描画待ち（店舗一覧の初期描画）
                await self._page.wait_for_timeout(3000)

                stores = await self._scrape_all_pages(product_code)

                # 空結果（初回描画ミス等）はキャッシュしない。
                # キャッシュすると 10分間、他の丸善店舗全てが「店舗名マッチ失敗」になる。
                if stores:
                    _store_cache[product_code] = (time.time(), stores)
                else:
                    logger.warning(
                        "丸善店舗一覧が空のためキャッシュを更新しません (code=%s)",
                        product_code,
                    )

                logger.info(
                    "丸善店舗一覧取得完了: %d店舗, code=%s",
                    len(stores), product_code,
                )

            except Exception as e:
                logger.error("丸善店舗一覧スクレイピングエラー: %s", e)
                self._save_scrape_result(
                    order_item_id, supplier_id,
                    AvailabilityStatus.ERROR, "", "", "",
                    error_flag=True, error_message=str(e),
                )
                return AvailabilityStatus.ERROR
            finally:
                await self.close_browser()

        # 該当店舗をマッチング
        matched = _match_store_name(stores, supplier_name)
        if matched is None:
            logger.warning(
                "丸善店舗マッチング失敗: supplier=%s, 取得店舗=%d件",
                supplier_name, len(stores),
            )
            self._save_scrape_result(
                order_item_id, supplier_id,
                AvailabilityStatus.UNKNOWN, "", "", "",
                error_message=f"店舗名マッチ失敗: {supplier_name}",
            )
            return AvailabilityStatus.UNKNOWN

        matched_name = matched["name"]
        status_text = matched["status"]
        screenshot_path = matched.get("screenshot_path", "")
        positive, negative = self._get_supplier_patterns()

        # ステータス文字列をパターンマッチ
        status, matched_pattern = self.judge_availability(
            [{"text": status_text}], positive, negative,
        )

        self._save_scrape_result(
            order_item_id, supplier_id,
            status, status_text, "", screenshot_path,
        )

        logger.info(
            "丸善店舗在庫判定: %s → %s [%s] (page=%s)",
            supplier_name, status.value, status_text, matched_name,
        )
        return status
