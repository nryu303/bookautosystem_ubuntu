"""mail_parser のユニットテスト"""

import os
import tempfile
import pytest

from src.common.logger import setup_logging

setup_logging(level="WARNING")


def _setup_test_db():
    """テスト用の一時DBを初期化する"""
    tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
    tmp.close()
    db_path = tmp.name

    import src.common.config as cfg
    cfg._config_cache = {
        "database": {"path": db_path},
        "storage": {"screenshots": "data/screenshots", "html_snapshots": "data/html_snapshots",
                     "output_dir": "data/output"},
        "logging": {"level": "WARNING", "dir": "logs"},
    }
    cfg._mail_patterns_cache = None  # 実際のmail_patterns.yamlを読ませる

    import src.common.database as db_mod
    db_mod._get_db_path = lambda: db_path

    from src.common.database import init_db
    init_db()

    return db_path


class TestMailParser:
    def setup_method(self):
        self._db_path = _setup_test_db()

    def teardown_method(self):
        try:
            os.unlink(self._db_path)
        except OSError:
            pass

    def test_html_to_text(self):
        from src.parser.mail_parser import _html_to_text

        html = "<html><body><p>テスト<br>メール</p><script>alert(1)</script></body></html>"
        text = _html_to_text(html)
        assert "テスト" in text
        assert "メール" in text
        assert "alert" not in text

    def test_extract_field_order_number(self):
        from src.parser.mail_parser import _extract_field

        text = "ご注文ありがとうございます。\n注文番号：ORD-12345\n商品をお届けします。"
        patterns = [{"pattern": r"注文番号[：:\s]*([A-Za-z0-9\-]+)", "group": 1}]
        result = _extract_field(text, patterns)
        assert result == "ORD-12345"

    def test_extract_amount(self):
        from src.parser.mail_parser import _extract_amount

        text = "価格：¥1,500"
        patterns = [{"pattern": r"価格[：:\s]*[¥￥]?([\d,]+)", "group": 1}]
        result = _extract_amount(text, patterns)
        assert result == 1500

    def test_extract_amount_no_match(self):
        from src.parser.mail_parser import _extract_amount

        text = "本日はありがとうございました"
        patterns = [{"pattern": r"価格[：:\s]*[¥￥]?([\d,]+)", "group": 1}]
        result = _extract_amount(text, patterns)
        assert result == 0

    def test_parse_pending_messages_with_isbn(self):
        from src.common.database import get_connection
        from src.parser.mail_parser import parse_pending_messages

        # テストメールを登録
        with get_connection() as conn:
            conn.execute(
                """
                INSERT INTO messages (outlook_message_id, sender, subject, raw_html, raw_text, parse_status)
                VALUES ('TEST001', 'shop@example.com', '注文確認 注文番号：A-001',
                        '', '注文番号：A-001\n商品名：テスト書籍\n価格：1500\n数量：2\nISBN：978-4-10-101001-4', 'PENDING')
                """
            )

        count = parse_pending_messages()
        assert count >= 1

        with get_connection() as conn:
            items = conn.execute("SELECT * FROM order_items").fetchall()
            assert len(items) >= 1

            item = items[0]
            assert item["product_code_normalized"] == "9784101010014"
            assert item["current_status"] == "PENDING"

            # メッセージがDONEになっている
            msg = conn.execute(
                "SELECT parse_status FROM messages WHERE outlook_message_id = 'TEST001'"
            ).fetchone()
            assert msg["parse_status"] == "DONE"

    def test_parse_empty_message(self):
        from src.common.database import get_connection
        from src.parser.mail_parser import parse_pending_messages

        with get_connection() as conn:
            conn.execute(
                """
                INSERT INTO messages (outlook_message_id, sender, subject, raw_html, raw_text, parse_status)
                VALUES ('EMPTY001', 'x@y.com', 'empty', '', '', 'PENDING')
                """
            )

        parse_pending_messages()

        with get_connection() as conn:
            msg = conn.execute(
                "SELECT parse_status FROM messages WHERE outlook_message_id = 'EMPTY001'"
            ).fetchone()
            assert msg["parse_status"] == "ERROR"
