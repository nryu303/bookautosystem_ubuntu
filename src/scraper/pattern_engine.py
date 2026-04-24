"""パターンベース抽出エンジン

extraction_patterns テーブルに登録されたCSSセレクタ/テキストヒントを使い、
ハードコードなしで任意のサイトから在庫情報を抽出する。
"""

from src.common.database import get_connection
from src.common.logger import get_logger

logger = get_logger(__name__)


def get_extraction_patterns(supplier_id: int) -> list[dict]:
    """指定仕入先の有効な抽出パターンを取得する。

    Returns:
        [{"field_name": "stock_text", "selector": "div.stock", ...}, ...]
    """
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT field_name, selector, xpath, class_hint, text_hint, pattern_name
            FROM extraction_patterns
            WHERE supplier_id = ? AND active_flag = 1
            ORDER BY id ASC
            """,
            (supplier_id,),
        ).fetchall()
    return [dict(r) for r in rows]


def has_patterns(supplier_id: int) -> bool:
    """指定仕入先にパターンが登録されているか。"""
    with get_connection() as conn:
        row = conn.execute(
            "SELECT COUNT(*) AS cnt FROM extraction_patterns WHERE supplier_id = ? AND active_flag = 1",
            (supplier_id,),
        ).fetchone()
    return row["cnt"] > 0


async def extract_by_patterns(page, patterns: list[dict]) -> dict[str, str]:
    """登録パターンを使ってページからフィールド値を抽出する。

    Args:
        page: Playwright page object
        patterns: get_extraction_patterns() の戻り値

    Returns:
        {"stock_text": "在庫あり", "price": "1,870円", ...}
    """
    results = {}

    for pat in patterns:
        field = pat["field_name"]
        selector = pat.get("selector", "").strip()
        text_hint = pat.get("text_hint", "").strip()
        class_hint = pat.get("class_hint", "").strip()

        extracted = None

        # Strategy 1: CSS selector
        if selector:
            try:
                el = await page.query_selector(selector)
                if el:
                    extracted = (await el.text_content() or "").strip()
            except Exception as e:
                logger.debug("Selector failed (%s): %s", selector, e)

        # Strategy 2: class hint — find element by class name
        if not extracted and class_hint:
            try:
                el = await page.query_selector(f".{class_hint}")
                if el:
                    extracted = (await el.text_content() or "").strip()
            except Exception as e:
                logger.debug("Class hint failed (%s): %s", class_hint, e)

        # Strategy 3: text hint — search all text for a line containing the hint
        if not extracted and text_hint:
            try:
                all_text = await page.evaluate("""
                    (hint) => {
                        const walker = document.createTreeWalker(
                            document.body, NodeFilter.SHOW_TEXT, null, false
                        );
                        let node;
                        while (node = walker.nextNode()) {
                            const t = node.textContent.trim();
                            if (t && t.includes(hint)) return t;
                        }
                        return '';
                    }
                """, text_hint)
                if all_text:
                    extracted = all_text.strip()
            except Exception as e:
                logger.debug("Text hint failed (%s): %s", text_hint, e)

        if extracted:
            results[field] = extracted

    return results
