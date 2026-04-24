"""URL生成モジュール

仕入先ごとの設定に基づき、商品コードからスクレイピング対象URLを生成する。
3つの方式に対応:
  ① template: JAN/ISBNをURLテンプレートに組み込む方式
  ② csv_lookup: 従来の簡易CSV参照方式（コード→URL 2列）
  ③ url_csv_db: URL指定型CSV（機能②）からDB経由で最安値URLを取得する方式
"""

import csv
import os

from src.common.config import BASE_DIR
from src.common.database import get_connection
from src.common.logger import get_logger

logger = get_logger(__name__)

# CSV参照方式のキャッシュ {csv_path: {product_code: url}}
_csv_cache: dict[str, dict[str, str]] = {}


def _load_csv_lookup(csv_path: str) -> dict[str, str]:
    """商品コード→URL対応CSVを読み込んでキャッシュする。"""
    if csv_path in _csv_cache:
        return _csv_cache[csv_path]

    full_path = os.path.join(BASE_DIR, csv_path)
    mapping = {}
    if not os.path.exists(full_path):
        logger.warning("URL参照CSV が見つかりません: %s", full_path)
        return mapping

    try:
        with open(full_path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            next(reader, None)  # ヘッダー読み飛ばし
            for row in reader:
                if len(row) >= 2 and row[0].strip() and row[1].strip():
                    mapping[row[0].strip()] = row[1].strip()
    except Exception as e:
        logger.error("URL参照CSV読込エラー (%s): %s", csv_path, e)

    _csv_cache[csv_path] = mapping
    return mapping


def _lookup_url_csv_db(product_code: str) -> str | None:
    """url_csv_entriesテーブルから最安値のURLを取得する（機能②）。

    9列目（item_url）を返す。同一コードで最安値の行（is_cheapest=1）を優先。
    """
    with get_connection() as conn:
        row = conn.execute(
            """
            SELECT item_url FROM url_csv_entries
            WHERE product_code = ? AND is_cheapest = 1
            LIMIT 1
            """,
            (product_code,),
        ).fetchone()

    if row and row["item_url"]:
        return row["item_url"]
    return None


def build_url(supplier_id: int, product_code: str) -> str | None:
    """仕入先とコードからスクレイピング用URLを生成する。

    Args:
        supplier_id: suppliersテーブルのID
        product_code: 正規化済みの商品コード

    Returns:
        URL文字列。生成できない場合はNone。
    """
    with get_connection() as conn:
        rule = conn.execute(
            "SELECT mode, url_template, lookup_csv_path FROM url_rules WHERE supplier_id = ?",
            (supplier_id,),
        ).fetchone()

    if rule is None:
        logger.warning("URL生成ルール未定義 (supplier_id=%d)", supplier_id)
        return None

    mode = rule["mode"]
    template = rule["url_template"] or ""
    csv_path = rule["lookup_csv_path"] or ""

    # 機能② URL指定型CSV（DBテーブル参照、最安値選択）
    if mode == "url_csv_db":
        url = _lookup_url_csv_db(product_code)
        if url:
            return url
        logger.debug("URL指定型CSVでURL未検出 (code=%s)", product_code)
        if template:
            return template.replace("{code}", product_code)
        return None

    # 従来のCSV参照方式
    if mode == "csv_lookup" and csv_path:
        mapping = _load_csv_lookup(csv_path)
        url = mapping.get(product_code)
        if url:
            return url
        logger.debug("CSV参照でURL未検出 (code=%s, csv=%s)", product_code, csv_path)
        # フォールバック: テンプレートがあればそちらを使う
        if template:
            return template.replace("{code}", product_code)
        return None

    # テンプレート方式（機能①）
    if template:
        return template.replace("{code}", product_code)

    logger.warning(
        "URL生成不可 (supplier_id=%d, mode=%s): テンプレートもCSVも未設定",
        supplier_id, mode,
    )
    return None


def clear_csv_cache() -> None:
    """CSV参照キャッシュをクリアする。"""
    global _csv_cache
    _csv_cache = {}
