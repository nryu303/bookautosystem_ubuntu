"""hold_manager のユニットテスト"""

import os
import tempfile
import pytest

from src.common.logger import setup_logging

setup_logging(level="WARNING")


def _setup_test_db():
    """テスト用の一時DBを初期化し、仕入先とorder_itemsを入れる"""
    tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
    tmp.close()
    db_path = tmp.name

    import src.common.config as cfg
    cfg._config_cache = {
        "database": {"path": db_path},
        "storage": {"screenshots": "data/screenshots", "html_snapshots": "data/html_snapshots",
                     "output_dir": "data/output"},
        "hold": {"strategy": "fill_smallest"},
        "logging": {"level": "WARNING", "dir": "logs"},
    }

    import src.common.database as db_mod
    db_mod._get_db_path = lambda: db_path

    from src.common.database import init_db, get_connection
    init_db()

    with get_connection() as conn:
        # 仕入先追加
        conn.execute(
            """
            INSERT INTO suppliers (supplier_code, supplier_name, category, priority, hold_limit_amount, hold_limit_days)
            VALUES ('test_sup', 'テスト仕入先', 'ec', 1, 5000, 14)
            """
        )
        # メッセージ追加
        conn.execute(
            """
            INSERT INTO messages (outlook_message_id, sender, subject, parse_status)
            VALUES ('HM_MSG001', 'a@b.com', 'test', 'DONE')
            """
        )
        msg = conn.execute("SELECT id FROM messages ORDER BY id DESC LIMIT 1").fetchone()
        # order_items 追加
        for i in range(3):
            conn.execute(
                """
                INSERT INTO order_items (message_id, product_code_normalized, product_name, amount, quantity, current_status)
                VALUES (?, ?, ?, ?, 1, 'HOLD')
                """,
                (msg["id"], f"978410101{i:04d}", f"テスト書籍{i+1}", 2000),
            )

    return db_path


class TestHoldManager:
    def setup_method(self):
        self._db_path = _setup_test_db()

    def teardown_method(self):
        try:
            os.unlink(self._db_path)
        except OSError:
            pass

    def test_add_to_hold(self):
        from src.common.database import get_connection
        from src.hold.hold_manager import add_to_hold

        with get_connection() as conn:
            sup = conn.execute("SELECT id FROM suppliers WHERE supplier_code = 'test_sup'").fetchone()
            items = conn.execute("SELECT id, amount FROM order_items ORDER BY id").fetchall()

        add_to_hold(items[0]["id"], sup["id"], items[0]["amount"])

        with get_connection() as conn:
            bucket = conn.execute(
                "SELECT * FROM hold_buckets WHERE supplier_id = ?", (sup["id"],)
            ).fetchone()
            assert bucket["total_amount"] == 2000
            assert bucket["item_count"] == 1

    def test_threshold_reached(self):
        from src.common.database import get_connection
        from src.hold.hold_manager import add_to_hold

        with get_connection() as conn:
            sup = conn.execute("SELECT id FROM suppliers WHERE supplier_code = 'test_sup'").fetchone()
            items = conn.execute("SELECT id, amount FROM order_items ORDER BY id").fetchall()

        # 3件×2000円=6000円 → hold_limit_amount=5000を超える
        for item in items:
            add_to_hold(item["id"], sup["id"], item["amount"])

        with get_connection() as conn:
            bucket = conn.execute(
                "SELECT * FROM hold_buckets WHERE supplier_id = ?", (sup["id"],)
            ).fetchone()
            assert bucket["total_amount"] == 6000
            assert bucket["item_count"] == 3
            assert bucket["status"] == "THRESHOLD_REACHED"

    def test_clear_bucket(self):
        from src.common.database import get_connection
        from src.hold.hold_manager import add_to_hold, clear_bucket

        with get_connection() as conn:
            sup = conn.execute("SELECT id FROM suppliers WHERE supplier_code = 'test_sup'").fetchone()
            items = conn.execute("SELECT id, amount FROM order_items ORDER BY id LIMIT 1").fetchall()

        add_to_hold(items[0]["id"], sup["id"], items[0]["amount"])

        with get_connection() as conn:
            bucket = conn.execute(
                "SELECT id FROM hold_buckets WHERE supplier_id = ?", (sup["id"],)
            ).fetchone()

        clear_bucket(bucket["id"])

        with get_connection() as conn:
            bucket = conn.execute(
                "SELECT * FROM hold_buckets WHERE supplier_id = ?", (sup["id"],)
            ).fetchone()
            assert bucket["total_amount"] == 0
            assert bucket["item_count"] == 0
            assert bucket["status"] == "CLEARED"

            item = conn.execute(
                "SELECT current_status FROM order_items WHERE id = ?", (items[0]["id"],)
            ).fetchone()
            assert item["current_status"] == "ORDERED"

    def test_get_threshold_reached_buckets(self):
        from src.common.database import get_connection
        from src.hold.hold_manager import add_to_hold, get_threshold_reached_buckets

        with get_connection() as conn:
            sup = conn.execute("SELECT id FROM suppliers WHERE supplier_code = 'test_sup'").fetchone()
            items = conn.execute("SELECT id, amount FROM order_items ORDER BY id").fetchall()

        for item in items:
            add_to_hold(item["id"], sup["id"], item["amount"])

        reached = get_threshold_reached_buckets()
        assert len(reached) >= 1
        assert reached[0]["supplier_code"] == "test_sup"
