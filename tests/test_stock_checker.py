"""self_stock_checker のユニットテスト"""

import os
import tempfile
import pytest

from src.common.logger import setup_logging

setup_logging(level="WARNING")


def _setup_test_env():
    """テスト用の一時DBとCSVを作成する"""
    tmp_db = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
    tmp_db.close()
    db_path = tmp_db.name

    tmp_csv = tempfile.NamedTemporaryFile(suffix=".csv", delete=False, mode="w", encoding="utf-8")
    tmp_csv.write("商品コード,在庫数\n")
    tmp_csv.write("9784101010014,5\n")
    tmp_csv.write("978-4-06-293842-6,3\n")
    tmp_csv.write("B0CXYZ1234,10\n")
    tmp_csv.close()
    csv_path = tmp_csv.name

    import src.common.config as cfg
    cfg._config_cache = {
        "database": {"path": db_path},
        "self_stock": {"csv_path": csv_path},
        "storage": {"screenshots": "data/screenshots", "html_snapshots": "data/html_snapshots",
                     "output_dir": "data/output"},
        "logging": {"level": "WARNING", "dir": "logs"},
    }

    import src.common.database as db_mod
    db_mod._get_db_path = lambda: db_path

    from src.common.database import init_db
    init_db()

    return db_path, csv_path


class TestSelfStockChecker:
    def setup_method(self):
        self._db_path, self._csv_path = _setup_test_env()

    def teardown_method(self):
        for p in (self._db_path, self._csv_path):
            try:
                os.unlink(p)
            except OSError:
                pass

    def test_import_csv(self):
        from src.common.database import get_connection
        from src.stock.self_stock_checker import import_self_stock_csv

        count = import_self_stock_csv(self._csv_path)
        assert count == 3

        with get_connection() as conn:
            rows = conn.execute("SELECT * FROM self_stock").fetchall()
            assert len(rows) == 3
            codes = [r["product_code"] for r in rows]
            assert "9784101010014" in codes
            assert "B0CXYZ1234" in codes

    def test_check_self_stock_hit(self):
        from src.common.database import get_connection
        from src.stock.self_stock_checker import import_self_stock_csv, check_self_stock

        import_self_stock_csv(self._csv_path)

        # PENDING商品を追加（自家在庫にある商品コード）
        with get_connection() as conn:
            conn.execute(
                """
                INSERT INTO messages (outlook_message_id, sender, subject, parse_status)
                VALUES ('SC_MSG001', 'a@b.com', 'test', 'DONE')
                """
            )
            msg = conn.execute("SELECT id FROM messages ORDER BY id DESC LIMIT 1").fetchone()
            conn.execute(
                """
                INSERT INTO order_items (message_id, product_code_normalized, product_name, amount, current_status)
                VALUES (?, '9784101010014', '在庫あり書籍', 1000, 'PENDING')
                """,
                (msg["id"],),
            )

        hits = check_self_stock()
        assert hits == 1

        with get_connection() as conn:
            item = conn.execute(
                "SELECT current_status FROM order_items WHERE product_code_normalized = '9784101010014'"
            ).fetchone()
            assert item["current_status"] == "SELF_STOCK"

    def test_check_self_stock_no_hit(self):
        from src.common.database import get_connection
        from src.stock.self_stock_checker import import_self_stock_csv, check_self_stock

        import_self_stock_csv(self._csv_path)

        with get_connection() as conn:
            conn.execute(
                """
                INSERT INTO messages (outlook_message_id, sender, subject, parse_status)
                VALUES ('SC_MSG002', 'a@b.com', 'test', 'DONE')
                """
            )
            msg = conn.execute("SELECT id FROM messages ORDER BY id DESC LIMIT 1").fetchone()
            conn.execute(
                """
                INSERT INTO order_items (message_id, product_code_normalized, product_name, amount, current_status)
                VALUES (?, '9789999999999', '在庫なし書籍', 2000, 'PENDING')
                """,
                (msg["id"],),
            )

        hits = check_self_stock()
        assert hits == 0

    def test_import_missing_csv(self):
        from src.stock.self_stock_checker import import_self_stock_csv

        count = import_self_stock_csv("/nonexistent/path.csv")
        assert count == 0
