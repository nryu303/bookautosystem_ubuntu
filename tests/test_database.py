"""database モジュールのユニットテスト"""

import os
import tempfile
import pytest

from src.common.logger import setup_logging

setup_logging(level="WARNING")


class TestDatabase:
    """DB初期化・テーブル作成・仕入先同期のテスト"""

    def setup_method(self):
        """テストごとに一時DBを使う"""
        self._tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
        self._tmp.close()
        self._db_path = self._tmp.name

        # config のDB pathを一時的に上書き
        import src.common.config as cfg
        cfg._config_cache = {
            "database": {"path": self._db_path},
            "storage": {
                "screenshots": "data/screenshots",
                "html_snapshots": "data/html_snapshots",
                "output_dir": "data/output",
            },
            "logging": {"level": "WARNING", "dir": "logs"},
        }

        # database モジュールの DB パス解決を一時ファイルに差し替え
        import src.common.database as db_mod
        self._original_get_db_path = db_mod._get_db_path
        db_mod._get_db_path = lambda: self._db_path

    def teardown_method(self):
        import src.common.database as db_mod
        db_mod._get_db_path = self._original_get_db_path
        try:
            os.unlink(self._db_path)
        except OSError:
            pass

    def test_init_db_creates_tables(self):
        from src.common.database import init_db, get_connection

        init_db()

        with get_connection() as conn:
            tables = conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
            ).fetchall()
            table_names = [t["name"] for t in tables]

        expected = [
            "extraction_patterns", "hold_buckets", "hold_items",
            "logs", "messages", "order_items", "outgoing_mails",
            "scrape_results", "self_stock", "suppliers", "url_rules",
        ]
        for name in expected:
            assert name in table_names, f"テーブル '{name}' が作成されていません"

    def test_sync_suppliers(self):
        from src.common.database import init_db, sync_suppliers_from_config, get_connection

        init_db()

        test_suppliers = [
            {
                "code": "test_site",
                "name": "テストサイト",
                "category": "ec",
                "scrape_enabled": True,
                "auto_mail_enabled": False,
                "priority": 1,
                "hold_limit_amount": 10000,
                "hold_limit_days": 7,
                "mail_limit_amount": 0,
                "mail_limit_days": 0,
                "mail_to_address": "",
                "login_required": False,
                "url_mode": "template",
                "url_template": "https://example.com/?q={code}",
                "lookup_csv_path": "",
            },
        ]

        sync_suppliers_from_config(test_suppliers)

        with get_connection() as conn:
            sup = conn.execute(
                "SELECT * FROM suppliers WHERE supplier_code = 'test_site'"
            ).fetchone()
            assert sup is not None
            assert sup["supplier_name"] == "テストサイト"
            assert sup["priority"] == 1
            assert sup["hold_limit_amount"] == 10000

            # url_rules も作成されている
            rule = conn.execute(
                "SELECT * FROM url_rules WHERE supplier_id = ?", (sup["id"],)
            ).fetchone()
            assert rule is not None
            assert rule["url_template"] == "https://example.com/?q={code}"

            # hold_buckets も初期化されている
            bucket = conn.execute(
                "SELECT * FROM hold_buckets WHERE supplier_id = ?", (sup["id"],)
            ).fetchone()
            assert bucket is not None
            assert bucket["total_amount"] == 0

    def test_upsert_suppliers(self):
        """同じコードで2回syncしても重複しない"""
        from src.common.database import init_db, sync_suppliers_from_config, get_connection

        init_db()

        suppliers = [{"code": "dup_test", "name": "初回", "priority": 1,
                       "url_mode": "template", "url_template": ""}]
        sync_suppliers_from_config(suppliers)

        suppliers[0]["name"] = "更新後"
        suppliers[0]["priority"] = 5
        sync_suppliers_from_config(suppliers)

        with get_connection() as conn:
            rows = conn.execute(
                "SELECT * FROM suppliers WHERE supplier_code = 'dup_test'"
            ).fetchall()
            assert len(rows) == 1
            assert rows[0]["supplier_name"] == "更新後"
            assert rows[0]["priority"] == 5

    def test_message_insert_and_query(self):
        from src.common.database import init_db, get_connection

        init_db()

        with get_connection() as conn:
            conn.execute(
                """
                INSERT INTO messages (outlook_message_id, sender, subject, parse_status)
                VALUES ('MSG001', 'test@example.com', 'テスト注文', 'PENDING')
                """
            )

        with get_connection() as conn:
            msg = conn.execute(
                "SELECT * FROM messages WHERE outlook_message_id = 'MSG001'"
            ).fetchone()
            assert msg is not None
            assert msg["sender"] == "test@example.com"
            assert msg["parse_status"] == "PENDING"

    def test_order_item_foreign_key(self):
        from src.common.database import init_db, get_connection

        init_db()

        with get_connection() as conn:
            conn.execute(
                """
                INSERT INTO messages (outlook_message_id, sender, subject, parse_status)
                VALUES ('MSG002', 'a@b.com', 'order', 'DONE')
                """
            )
            msg = conn.execute("SELECT id FROM messages WHERE outlook_message_id = 'MSG002'").fetchone()

            conn.execute(
                """
                INSERT INTO order_items (message_id, product_code_normalized, product_name, amount, quantity, current_status)
                VALUES (?, '9784101010014', 'テスト書籍', 1500, 1, 'PENDING')
                """,
                (msg["id"],),
            )

        with get_connection() as conn:
            item = conn.execute(
                "SELECT * FROM order_items WHERE product_code_normalized = '9784101010014'"
            ).fetchone()
            assert item is not None
            assert item["amount"] == 1500
            assert item["message_id"] == msg["id"]
