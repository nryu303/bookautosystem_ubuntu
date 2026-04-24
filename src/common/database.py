"""SQLiteデータベース接続・テーブル作成・共通操作"""

import os
import sqlite3
from contextlib import contextmanager

from src.common.config import get_setting, BASE_DIR
from src.common.exceptions import DatabaseError
from src.common.logger import get_logger

logger = get_logger(__name__)


def _get_db_path() -> str:
    """設定からDBファイルパスを解決する。"""
    rel_path = get_setting("database", "path", default="data/db/main.db")
    db_path = os.path.join(BASE_DIR, rel_path)
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    return db_path


@contextmanager
def get_connection():
    """DBコネクションをコンテキストマネージャで返す。

    使い方:
        with get_connection() as conn:
            conn.execute("SELECT ...")
    """
    db_path = _get_db_path()
    conn = sqlite3.connect(db_path, timeout=60)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=60000")
    conn.execute("PRAGMA foreign_keys=ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


# ── テーブル定義 ──────────────────────────────────────────

TABLES_SQL = [
    # 1. messages（受信メール）
    """
    CREATE TABLE IF NOT EXISTS messages (
        id                  INTEGER PRIMARY KEY AUTOINCREMENT,
        outlook_message_id  TEXT    UNIQUE NOT NULL,
        account_name        TEXT,
        sender              TEXT,
        recipient           TEXT,
        received_at         TEXT,
        subject             TEXT,
        raw_html            TEXT,
        raw_text            TEXT,
        parse_status        TEXT    DEFAULT 'PENDING',
        error_message       TEXT,
        created_at          TEXT    DEFAULT (datetime('now', 'localtime'))
    )
    """,

    # 2. order_items（商品単位の処理行）
    """
    CREATE TABLE IF NOT EXISTS order_items (
        id                      INTEGER PRIMARY KEY AUTOINCREMENT,
        message_id              INTEGER REFERENCES messages(id),
        order_number            TEXT,
        product_code_raw        TEXT,
        product_code_normalized TEXT,
        product_name            TEXT,
        amount                  INTEGER DEFAULT 0,
        quantity                INTEGER DEFAULT 1,
        item_condition          TEXT    DEFAULT '',
        recipient_name          TEXT    DEFAULT '',
        edition                 TEXT    DEFAULT '',
        pub_year                TEXT    DEFAULT '',
        author                  TEXT    DEFAULT '',
        sku                     TEXT    DEFAULT '',
        order_date              TEXT    DEFAULT '',
        ship_date               TEXT    DEFAULT '',
        tax                     INTEGER DEFAULT 0,
        shipping                INTEGER DEFAULT 0,
        current_status          TEXT    DEFAULT 'PENDING',
        planned_supplier_id     INTEGER REFERENCES suppliers(id),
        final_supplier_id       INTEGER REFERENCES suppliers(id),
        assigned_at             TEXT,
        created_at              TEXT    DEFAULT (datetime('now', 'localtime'))
    )
    """,

    # 3. self_stock（自家在庫）
    """
    CREATE TABLE IF NOT EXISTS self_stock (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        product_code    TEXT    NOT NULL,
        stock_qty       INTEGER DEFAULT 0,
        imported_at     TEXT    DEFAULT (datetime('now', 'localtime'))
    )
    """,

    # 4. suppliers（仕入先マスタ）
    """
    CREATE TABLE IF NOT EXISTS suppliers (
        id                  INTEGER PRIMARY KEY AUTOINCREMENT,
        supplier_code       TEXT    UNIQUE NOT NULL,
        supplier_name       TEXT    NOT NULL,
        category            TEXT    DEFAULT 'ec',
        scrape_enabled      INTEGER DEFAULT 1,
        auto_mail_enabled   INTEGER DEFAULT 0,
        priority            INTEGER DEFAULT 99,
        hold_limit_amount   INTEGER DEFAULT 0,
        hold_limit_days     INTEGER DEFAULT 14,
        mail_limit_amount       INTEGER DEFAULT 0,
        mail_limit_days         INTEGER DEFAULT 0,
        mail_unit_price_limit   INTEGER DEFAULT 0,
        mail_quantity_limit     INTEGER DEFAULT 0,
        mail_to_address         TEXT    DEFAULT '',
        login_required          INTEGER DEFAULT 0
    )
    """,

    # 5. url_rules（URL生成ルール）
    """
    CREATE TABLE IF NOT EXISTS url_rules (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        supplier_id     INTEGER REFERENCES suppliers(id) UNIQUE,
        mode            TEXT    DEFAULT 'template',
        url_template    TEXT    DEFAULT '',
        lookup_csv_path TEXT    DEFAULT ''
    )
    """,

    # 6. extraction_patterns（HTML抽出ルール）
    """
    CREATE TABLE IF NOT EXISTS extraction_patterns (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        supplier_id     INTEGER REFERENCES suppliers(id),
        pattern_name    TEXT,
        field_name      TEXT,
        selector        TEXT    DEFAULT '',
        xpath           TEXT    DEFAULT '',
        class_hint      TEXT    DEFAULT '',
        text_hint       TEXT    DEFAULT '',
        active_flag     INTEGER DEFAULT 1
    )
    """,

    # 7. scrape_results（スクレイピング結果）
    """
    CREATE TABLE IF NOT EXISTS scrape_results (
        id                  INTEGER PRIMARY KEY AUTOINCREMENT,
        order_item_id       INTEGER REFERENCES order_items(id),
        supplier_id         INTEGER REFERENCES suppliers(id),
        scraped_at          TEXT    DEFAULT (datetime('now', 'localtime')),
        availability_status TEXT    DEFAULT 'UNKNOWN',
        raw_stock_text      TEXT    DEFAULT '',
        html_snapshot_path  TEXT    DEFAULT '',
        screenshot_path     TEXT    DEFAULT '',
        error_flag          INTEGER DEFAULT 0,
        error_message       TEXT    DEFAULT ''
    )
    """,

    # 8. hold_buckets（保留バケット）
    """
    CREATE TABLE IF NOT EXISTS hold_buckets (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        supplier_id     INTEGER REFERENCES suppliers(id) UNIQUE,
        total_amount    INTEGER DEFAULT 0,
        item_count      INTEGER DEFAULT 0,
        oldest_item_date TEXT,
        status          TEXT    DEFAULT 'ACTIVE'
    )
    """,

    # 9. hold_items（保留中の個別商品）
    """
    CREATE TABLE IF NOT EXISTS hold_items (
        id                  INTEGER PRIMARY KEY AUTOINCREMENT,
        hold_bucket_id      INTEGER REFERENCES hold_buckets(id),
        order_item_id       INTEGER REFERENCES order_items(id),
        amount              INTEGER DEFAULT 0,
        assigned_at         TEXT    DEFAULT (datetime('now', 'localtime')),
        rescrape_required   INTEGER DEFAULT 0
    )
    """,

    # 10. outgoing_mails（送信メール履歴）
    """
    CREATE TABLE IF NOT EXISTS outgoing_mails (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        supplier_id     INTEGER REFERENCES suppliers(id),
        order_item_id   INTEGER REFERENCES order_items(id),
        to_address      TEXT,
        subject         TEXT,
        body            TEXT,
        sent_at         TEXT    DEFAULT (datetime('now', 'localtime')),
        send_status     TEXT    DEFAULT 'SUCCESS',
        error_message   TEXT    DEFAULT ''
    )
    """,

    # 11. logs（システムログ）
    """
    CREATE TABLE IF NOT EXISTS logs (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        level       TEXT,
        module      TEXT,
        message     TEXT,
        created_at  TEXT    DEFAULT (datetime('now', 'localtime'))
    )
    """,

    # 12. url_csv_entries（URL指定型CSV — 機能②）
    """
    CREATE TABLE IF NOT EXISTS url_csv_entries (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        product_code    TEXT    NOT NULL,
        book_title      TEXT    DEFAULT '',
        publisher       TEXT    DEFAULT '',
        author          TEXT    DEFAULT '',
        pub_year        TEXT    DEFAULT '',
        list_price      INTEGER DEFAULT 0,
        page_url        TEXT    DEFAULT '',
        item_text       TEXT    DEFAULT '',
        item_url        TEXT    DEFAULT '',
        price           INTEGER DEFAULT 0,
        field2          TEXT    DEFAULT '',
        field3          TEXT    DEFAULT '',
        scraped_at      TEXT    DEFAULT '',
        is_cheapest     INTEGER DEFAULT 0,
        imported_at     TEXT    DEFAULT (datetime('now', 'localtime'))
    )
    """,
]

# インデックス定義
INDEXES_SQL = [
    "CREATE INDEX IF NOT EXISTS idx_order_items_message_id ON order_items(message_id)",
    "CREATE INDEX IF NOT EXISTS idx_order_items_status ON order_items(current_status)",
    "CREATE INDEX IF NOT EXISTS idx_order_items_product_code ON order_items(product_code_normalized)",
    "CREATE INDEX IF NOT EXISTS idx_scrape_results_order_item ON scrape_results(order_item_id)",
    "CREATE INDEX IF NOT EXISTS idx_scrape_results_supplier ON scrape_results(supplier_id)",
    "CREATE INDEX IF NOT EXISTS idx_hold_items_bucket ON hold_items(hold_bucket_id)",
    "CREATE INDEX IF NOT EXISTS idx_hold_items_order_item ON hold_items(order_item_id)",
    "CREATE INDEX IF NOT EXISTS idx_self_stock_code ON self_stock(product_code)",
    "CREATE INDEX IF NOT EXISTS idx_messages_parse_status ON messages(parse_status)",
    "CREATE INDEX IF NOT EXISTS idx_outgoing_mails_order_item ON outgoing_mails(order_item_id)",
    "CREATE INDEX IF NOT EXISTS idx_url_csv_entries_code ON url_csv_entries(product_code)",
    "CREATE INDEX IF NOT EXISTS idx_url_csv_entries_cheapest ON url_csv_entries(product_code, is_cheapest)",
]


_MIGRATIONS = [
    "ALTER TABLE order_items ADD COLUMN item_condition TEXT DEFAULT ''",
    "ALTER TABLE order_items ADD COLUMN condition_note TEXT DEFAULT ''",
    "ALTER TABLE order_items ADD COLUMN recipient_name TEXT DEFAULT ''",
    "ALTER TABLE order_items ADD COLUMN edition TEXT DEFAULT ''",
    "ALTER TABLE order_items ADD COLUMN pub_year TEXT DEFAULT ''",
    "ALTER TABLE order_items ADD COLUMN author TEXT DEFAULT ''",
    "ALTER TABLE order_items ADD COLUMN sku TEXT DEFAULT ''",
    "ALTER TABLE order_items ADD COLUMN order_date TEXT DEFAULT ''",
    "ALTER TABLE order_items ADD COLUMN ship_date TEXT DEFAULT ''",
    "ALTER TABLE order_items ADD COLUMN tax INTEGER DEFAULT 0",
    "ALTER TABLE order_items ADD COLUMN shipping INTEGER DEFAULT 0",
]


def init_db() -> None:
    """全テーブルとインデックスを作成する。アプリ起動時に呼ぶ。"""
    try:
        with get_connection() as conn:
            for sql in TABLES_SQL:
                conn.execute(sql)
            for sql in INDEXES_SQL:
                conn.execute(sql)
            # 既存DBへのカラム追加マイグレーション
            for sql in _MIGRATIONS:
                try:
                    conn.execute(sql)
                except Exception:
                    pass  # カラムが既に存在する場合は無視
        logger.info("データベース初期化完了")
    except Exception as e:
        raise DatabaseError(f"データベース初期化に失敗しました: {e}")


def sync_suppliers_from_config(suppliers_config: list[dict]) -> None:
    """suppliers.yamlの内容をsuppliersテーブルとurl_rulesテーブルに同期する。"""
    with get_connection() as conn:
        for sup in suppliers_config:
            # suppliers テーブル UPSERT
            conn.execute(
                """
                INSERT INTO suppliers (
                    supplier_code, supplier_name, category,
                    scrape_enabled, auto_mail_enabled, priority,
                    hold_limit_amount, hold_limit_days,
                    mail_limit_amount, mail_limit_days,
                    mail_unit_price_limit, mail_quantity_limit,
                    mail_to_address, login_required
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(supplier_code) DO UPDATE SET
                    supplier_name = excluded.supplier_name,
                    category = excluded.category,
                    scrape_enabled = excluded.scrape_enabled,
                    auto_mail_enabled = excluded.auto_mail_enabled,
                    priority = excluded.priority,
                    hold_limit_amount = excluded.hold_limit_amount,
                    hold_limit_days = excluded.hold_limit_days,
                    mail_limit_amount = excluded.mail_limit_amount,
                    mail_limit_days = excluded.mail_limit_days,
                    mail_unit_price_limit = excluded.mail_unit_price_limit,
                    mail_quantity_limit = excluded.mail_quantity_limit,
                    mail_to_address = excluded.mail_to_address,
                    login_required = excluded.login_required
                """,
                (
                    sup["code"], sup["name"], sup.get("category", "ec"),
                    int(sup.get("scrape_enabled", True)),
                    int(sup.get("auto_mail_enabled", False)),
                    sup.get("priority", 99),
                    sup.get("hold_limit_amount", 0),
                    sup.get("hold_limit_days", 14),
                    sup.get("mail_limit_amount", 0),
                    sup.get("mail_limit_days", 0),
                    sup.get("mail_unit_price_limit", 0),
                    sup.get("mail_quantity_limit", 0),
                    sup.get("mail_to_address", ""),
                    int(sup.get("login_required", False)),
                ),
            )

            # supplier_id を取得
            row = conn.execute(
                "SELECT id FROM suppliers WHERE supplier_code = ?",
                (sup["code"],),
            ).fetchone()
            supplier_id = row["id"]

            # url_rules テーブル UPSERT
            conn.execute(
                """
                INSERT INTO url_rules (supplier_id, mode, url_template, lookup_csv_path)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(supplier_id) DO UPDATE SET
                    mode = excluded.mode,
                    url_template = excluded.url_template,
                    lookup_csv_path = excluded.lookup_csv_path
                """,
                (
                    supplier_id,
                    sup.get("url_mode", "template"),
                    sup.get("url_template", ""),
                    sup.get("lookup_csv_path", ""),
                ),
            )

            # hold_buckets 初期化（なければ作成）
            conn.execute(
                """
                INSERT OR IGNORE INTO hold_buckets (supplier_id, total_amount, item_count, status)
                VALUES (?, 0, 0, 'CLEARED')
                """,
                (supplier_id,),
            )

    logger.info("仕入先マスタ同期完了: %d件", len(suppliers_config))
