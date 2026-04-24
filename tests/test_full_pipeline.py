"""全工程テストスクリプト

ダミーの受注データ（ISBN付き）をDBに投入し、
スクレイピング結果の擬似投入 → 在庫判定 → 振り分け → 保留管理 → Excel出力
までの全工程を自動で通すテスト。

使い方:
  python -m pytest tests/test_full_pipeline.py -v
  python tests/test_full_pipeline.py              # 直接実行も可
"""

import os
import sys
import sqlite3

# プロジェクトルートをパスに追加
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.common.config import load_settings, load_suppliers, BASE_DIR
from src.common.database import init_db, sync_suppliers_from_config, get_connection
from src.common.logger import setup_logging


# ── テスト用ダミーデータ ──────────────────────────────────────

DUMMY_MESSAGES = [
    {
        "outlook_message_id": "TEST-MSG-001",
        "account_name": "test@example.com",
        "sender": "amazon@marketplace.co.jp",
        "recipient": "test@example.com",
        "received_at": "2026-03-20 10:00:00",
        "subject": "新規注文 #9784336065704SHINPIN07",
        "raw_html": "<p>注文番号: 9784336065704SHINPIN07</p><p>商品名: A書籍</p><p>金額: 4186</p>",
        "raw_text": "注文番号: 9784336065704SHINPIN07\n商品名: A書籍\n金額: 4186",
        "parse_status": "DONE",
    },
    {
        "outlook_message_id": "TEST-MSG-002",
        "account_name": "test@example.com",
        "sender": "amazon@marketplace.co.jp",
        "recipient": "test@example.com",
        "received_at": "2026-03-20 10:05:00",
        "subject": "新規注文 #9784336065705SHINPIN07",
        "raw_html": "<p>注文番号: 9784336065705SHINPIN07</p><p>商品名: B書籍</p><p>金額: 4800</p>",
        "raw_text": "注文番号: 9784336065705SHINPIN07\n商品名: B書籍\n金額: 4800",
        "parse_status": "DONE",
    },
    {
        "outlook_message_id": "TEST-MSG-003",
        "account_name": "test@example.com",
        "sender": "amazon@marketplace.co.jp",
        "recipient": "test@example.com",
        "received_at": "2026-03-21 09:00:00",
        "subject": "新規注文 #9784336065706SHINPIN07",
        "raw_html": "<p>注文番号: 9784336065706SHINPIN07</p><p>商品名: C書籍</p><p>金額: 5600</p>",
        "raw_text": "注文番号: 9784336065706SHINPIN07\n商品名: C書籍\n金額: 5600",
        "parse_status": "DONE",
    },
    {
        "outlook_message_id": "TEST-MSG-004",
        "account_name": "test@example.com",
        "sender": "amazon@marketplace.co.jp",
        "recipient": "test@example.com",
        "received_at": "2026-03-21 11:00:00",
        "subject": "新規注文 #9784336065707SHINPIN07",
        "raw_html": "<p>注文番号: 9784336065707SHINPIN07</p><p>商品名: D書籍（在庫なし）</p><p>金額: 4100</p>",
        "raw_text": "注文番号: 9784336065707SHINPIN07\n商品名: D書籍（在庫なし）\n金額: 4100",
        "parse_status": "DONE",
    },
    {
        "outlook_message_id": "TEST-MSG-005",
        "account_name": "test@example.com",
        "sender": "amazon@marketplace.co.jp",
        "recipient": "test@example.com",
        "received_at": "2026-03-22 08:00:00",
        "subject": "新規注文 #9784893586247SHINPIN07",
        "raw_html": "<p>注文番号: 9784893586247SHINPIN07</p><p>商品名: E書籍（自家在庫あり）</p><p>金額: 2200</p>",
        "raw_text": "注文番号: 9784893586247SHINPIN07\n商品名: E書籍（自家在庫あり）\n金額: 2200",
        "parse_status": "DONE",
    },
]

DUMMY_ORDER_ITEMS = [
    # A書籍: EC複数サイトに在庫あり → 保留ロジックで振り分け
    {
        "message_id": 1,
        "order_number": "9784336065704SHINPIN07",
        "product_code_raw": "9784336065704SHINPIN07",
        "product_code_normalized": "9784336065704",
        "product_name": "A書籍",
        "amount": 4186,
        "quantity": 1,
        "current_status": "PENDING",
    },
    # B書籍: 店頭のみ在庫あり → 店頭に振り分け
    {
        "message_id": 2,
        "order_number": "9784336065705SHINPIN07",
        "product_code_raw": "9784336065705SHINPIN07",
        "product_code_normalized": "9784336065705",
        "product_name": "B書籍",
        "amount": 4800,
        "quantity": 1,
        "current_status": "PENDING",
    },
    # C書籍: EC1箇所だけ在庫あり → そのECに振り分け
    {
        "message_id": 3,
        "order_number": "9784336065706SHINPIN07",
        "product_code_raw": "9784336065706SHINPIN07",
        "product_code_normalized": "9784336065706",
        "product_name": "C書籍",
        "amount": 5600,
        "quantity": 1,
        "current_status": "PENDING",
    },
    # D書籍: どこにも在庫なし → NO_STOCK
    {
        "message_id": 4,
        "order_number": "9784336065707SHINPIN07",
        "product_code_raw": "9784336065707SHINPIN07",
        "product_code_normalized": "9784336065707",
        "product_name": "D書籍（在庫なし）",
        "amount": 4100,
        "quantity": 1,
        "current_status": "PENDING",
    },
    # E書籍: 自家在庫ヒット → SELF_STOCK
    {
        "message_id": 5,
        "order_number": "9784893586247SHINPIN07",
        "product_code_raw": "9784893586247SHINPIN07",
        "product_code_normalized": "9784893586247",
        "product_name": "E書籍（自家在庫あり）",
        "amount": 2200,
        "quantity": 1,
        "current_status": "PENDING",
    },
]

# スクレイピング結果を擬似的に投入（サイト別の在庫状態）
# supplier_code → availability_status のマッピング
DUMMY_SCRAPE_MAP = {
    # A書籍 (item_id=1): ヨドバシと楽天に在庫あり
    1: {
        "X000001": ("AVAILABLE", "在庫あり"),
        "X000002": ("AVAILABLE", "在庫あり"),
        "X000003": ("UNAVAILABLE", "ご注文いただけません"),
        "X000004": ("UNAVAILABLE", "絶版商品です"),
    },
    # B書籍 (item_id=2): EC在庫なし、丸善池袋本店のみ在庫あり
    2: {
        "X000001": ("UNAVAILABLE", "販売休止中です"),
        "X000002": ("UNAVAILABLE", "ご注文できない商品"),
        "X000003": ("UNAVAILABLE", "ご注文いただけません"),
        "X000018": ("AVAILABLE", "在庫〇 取り置き可"),
        "X000036": ("UNAVAILABLE", "× 在庫なし"),
    },
    # C書籍 (item_id=3): 紀伊國屋ECのみ在庫あり
    3: {
        "X000001": ("UNAVAILABLE", "販売休止中です"),
        "X000002": ("UNAVAILABLE", "ご注文できない商品"),
        "X000003": ("AVAILABLE", "ウェブストア専用在庫あり"),
        "X000004": ("UNAVAILABLE", "現在ご注文できません"),
    },
    # D書籍 (item_id=4): 全サイト在庫なし
    4: {
        "X000001": ("UNAVAILABLE", "販売休止中です"),
        "X000002": ("UNAVAILABLE", "ご注文できない商品"),
        "X000003": ("UNAVAILABLE", "ご注文いただけません"),
        "X000004": ("UNAVAILABLE", "絶版商品です"),
    },
    # E書籍 (item_id=5): 自家在庫ヒットなのでスクレイピング対象外
}


def _reset_db():
    """テスト用にDBを初期化する。"""
    db_path = os.path.join(BASE_DIR, "data", "db", "main.db")
    if os.path.exists(db_path):
        os.remove(db_path)
    load_settings(force_reload=True)
    init_db()
    suppliers = load_suppliers(force_reload=True)
    sync_suppliers_from_config(suppliers)


def _insert_dummy_data():
    """ダミーメール・注文データを投入する。"""
    with get_connection() as conn:
        for msg in DUMMY_MESSAGES:
            conn.execute(
                """
                INSERT INTO messages (
                    outlook_message_id, account_name, sender, recipient,
                    received_at, subject, raw_html, raw_text, parse_status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    msg["outlook_message_id"], msg["account_name"],
                    msg["sender"], msg["recipient"], msg["received_at"],
                    msg["subject"], msg["raw_html"], msg["raw_text"],
                    msg["parse_status"],
                ),
            )

        for item in DUMMY_ORDER_ITEMS:
            conn.execute(
                """
                INSERT INTO order_items (
                    message_id, order_number, product_code_raw,
                    product_code_normalized, product_name, amount,
                    quantity, current_status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    item["message_id"], item["order_number"],
                    item["product_code_raw"], item["product_code_normalized"],
                    item["product_name"], item["amount"],
                    item["quantity"], item["current_status"],
                ),
            )


def _insert_dummy_scrape_results():
    """擬似スクレイピング結果をDBに投入する。

    DUMMY_SCRAPE_MAP に明示的に指定されていない仕入先は
    UNAVAILABLE として自動補完し、全対象仕入先の結果を揃える。
    """
    from src.scraper import SCRAPER_REGISTRY

    with get_connection() as conn:
        # スクレイパー登録済みかつ有効な仕入先を全取得
        all_suppliers = conn.execute(
            "SELECT id, supplier_code FROM suppliers WHERE scrape_enabled = 1"
        ).fetchall()
        registered_suppliers = [
            dict(s) for s in all_suppliers
            if s["supplier_code"] in SCRAPER_REGISTRY
        ]

        for item_id, explicit_results in DUMMY_SCRAPE_MAP.items():
            # 明示指定分を投入
            for supplier_code, (status, text) in explicit_results.items():
                row = conn.execute(
                    "SELECT id FROM suppliers WHERE supplier_code = ?",
                    (supplier_code,),
                ).fetchone()
                if not row:
                    continue
                conn.execute(
                    """
                    INSERT INTO scrape_results (
                        order_item_id, supplier_id, availability_status,
                        raw_stock_text, error_flag
                    ) VALUES (?, ?, ?, ?, 0)
                    """,
                    (item_id, row["id"], status, text),
                )

            # 未指定の登録済み仕入先には UNAVAILABLE を自動補完
            for sup in registered_suppliers:
                if sup["supplier_code"] not in explicit_results:
                    conn.execute(
                        """
                        INSERT INTO scrape_results (
                            order_item_id, supplier_id, availability_status,
                            raw_stock_text, error_flag
                        ) VALUES (?, ?, 'UNAVAILABLE', '', 0)
                        """,
                        (item_id, sup["id"]),
                    )


def _mark_self_stock():
    """E書籍を自家在庫ヒットにする。"""
    with get_connection() as conn:
        conn.execute(
            "INSERT INTO self_stock (product_code, stock_qty) VALUES (?, ?)",
            ("9784893586247", 1),
        )
        conn.execute(
            "UPDATE order_items SET current_status = 'SELF_STOCK' WHERE id = 5",
        )


def test_full_pipeline():
    """全工程テスト"""
    print("=" * 60)
    print("BookAutoSystem 全工程テスト")
    print("=" * 60)

    # Step 1: DB初期化
    print("\n[Step 1] DB初期化...")
    _reset_db()
    with get_connection() as conn:
        count = conn.execute("SELECT COUNT(*) as c FROM suppliers").fetchone()["c"]
    print(f"  仕入先マスタ: {count}件")
    assert count == 40, f"Expected 40 suppliers, got {count}"

    # Step 2: ダミーデータ投入
    print("\n[Step 2] ダミーデータ投入...")
    _insert_dummy_data()
    with get_connection() as conn:
        msg_count = conn.execute("SELECT COUNT(*) as c FROM messages").fetchone()["c"]
        item_count = conn.execute("SELECT COUNT(*) as c FROM order_items").fetchone()["c"]
    print(f"  メール: {msg_count}件, 注文商品: {item_count}件")
    assert msg_count == 5
    assert item_count == 5

    # Step 3: 自家在庫照合
    print("\n[Step 3] 自家在庫照合...")
    _mark_self_stock()
    with get_connection() as conn:
        self_stock_items = conn.execute(
            "SELECT * FROM order_items WHERE current_status = 'SELF_STOCK'"
        ).fetchall()
    print(f"  自家在庫ヒット: {len(self_stock_items)}件")
    assert len(self_stock_items) == 1
    assert self_stock_items[0]["product_name"] == "E書籍（自家在庫あり）"

    # Step 4: スクレイピング結果投入（擬似）
    print("\n[Step 4] スクレイピング結果投入（擬似）...")
    _insert_dummy_scrape_results()
    with get_connection() as conn:
        scrape_count = conn.execute("SELECT COUNT(*) as c FROM scrape_results").fetchone()["c"]
    print(f"  スクレイピング結果: {scrape_count}件")
    assert scrape_count > 0

    # Step 5: 仕入先振り分け
    print("\n[Step 5] 仕入先振り分け（3段階優先順位）...")
    from src.judge.supplier_selector import assign_pending_items
    assigned = assign_pending_items()
    print(f"  振り分け成功: {assigned}件")

    with get_connection() as conn:
        items = conn.execute(
            "SELECT oi.*, s.supplier_name, s.supplier_code FROM order_items oi LEFT JOIN suppliers s ON oi.planned_supplier_id = s.id ORDER BY oi.id"
        ).fetchall()

    print("\n  --- 振り分け結果 ---")
    for item in items:
        sup_name = item["supplier_name"] or "-"
        print(f"  [{item['current_status']:10s}] {item['product_name']:20s} → {sup_name}")

    # A書籍: 在庫ありサイト複数 → HOLD
    a_item = items[0]
    assert a_item["current_status"] == "HOLD", f"A書籍 expected HOLD, got {a_item['current_status']}"

    # B書籍: 店頭のみ → HOLD
    b_item = items[1]
    assert b_item["current_status"] == "HOLD", f"B書籍 expected HOLD, got {b_item['current_status']}"

    # C書籍: EC1箇所 → HOLD
    c_item = items[2]
    assert c_item["current_status"] == "HOLD", f"C書籍 expected HOLD, got {c_item['current_status']}"

    # D書籍: 在庫なし → NO_STOCK
    d_item = items[3]
    assert d_item["current_status"] == "NO_STOCK", f"D書籍 expected NO_STOCK, got {d_item['current_status']}"

    # E書籍: 自家在庫 → SELF_STOCK
    e_item = items[4]
    assert e_item["current_status"] == "SELF_STOCK", f"E書籍 expected SELF_STOCK, got {e_item['current_status']}"

    # Step 6: 保留管理
    print("\n[Step 6] 保留バケット管理...")
    from src.hold.hold_manager import process_hold_assignments
    hold_count = process_hold_assignments()
    print(f"  保留処理: {hold_count}件")

    with get_connection() as conn:
        buckets = conn.execute(
            """
            SELECT hb.*, s.supplier_name
            FROM hold_buckets hb
            JOIN suppliers s ON hb.supplier_id = s.id
            WHERE hb.item_count > 0
            ORDER BY hb.total_amount DESC
            """
        ).fetchall()

    print("\n  --- 保留バケット状態 ---")
    for b in buckets:
        print(f"  {b['supplier_name']:20s}  合計={b['total_amount']:>8,}円  件数={b['item_count']}  状態={b['status']}")

    # Step 7: Excel台帳出力
    print("\n[Step 7] Excel台帳出力...")
    from src.export.excel_exporter import export_all
    proc_path, sup_path = export_all()
    print(f"  処理台帳: {proc_path}")
    print(f"  仕入台帳: {sup_path}")
    assert os.path.exists(proc_path), f"処理台帳が存在しない: {proc_path}"
    assert os.path.exists(sup_path), f"仕入台帳が存在しない: {sup_path}"

    # Step 8: 再スクレイピングトリガー確認
    print("\n[Step 8] 再スクレイピングトリガー確認...")
    from src.hold.rescrape_trigger import process_rescrape
    rescrape_result = process_rescrape()
    print(f"  チェック: {rescrape_result.get('checked', 0)}件")

    # Step 9: 最終状態サマリー
    print("\n" + "=" * 60)
    print("テスト結果サマリー")
    print("=" * 60)
    with get_connection() as conn:
        summary = conn.execute(
            """
            SELECT current_status, COUNT(*) as cnt
            FROM order_items
            GROUP BY current_status
            ORDER BY current_status
            """
        ).fetchall()
    for row in summary:
        print(f"  {row['current_status']:12s}: {row['cnt']}件")

    print("\n  全テスト合格!")
    return True


def test_three_tier_priority():
    """3段階優先順位ロジックの単体テスト"""
    print("\n" + "=" * 60)
    print("3段階優先順位ロジック テスト")
    print("=" * 60)

    from src.judge.supplier_selector import _apply_three_tier_priority

    # テストケース1: 第1位 - ライン未満で在庫あり → 最小保留合計金額
    candidates_1 = [
        {"supplier_id": 1, "supplier_code": "A", "supplier_name": "EC1", "category": "ec",
         "priority": 1, "total_amount": 5000, "item_count": 2, "hold_limit": 10000, "mail_unit_price_limit": 0},
        {"supplier_id": 2, "supplier_code": "B", "supplier_name": "EC2", "category": "ec",
         "priority": 2, "total_amount": 3000, "item_count": 1, "hold_limit": 10000, "mail_unit_price_limit": 0},
        {"supplier_id": 3, "supplier_code": "C", "supplier_name": "EC3", "category": "ec",
         "priority": 3, "total_amount": 8000, "item_count": 3, "hold_limit": 10000, "mail_unit_price_limit": 0},
    ]
    result = _apply_three_tier_priority(candidates_1)
    print(f"\n  Case1 (ライン未満): {result['supplier_name']} (tier={result['_tier']})")
    assert result["supplier_name"] == "EC2", "最小保留合計金額のEC2が選ばれるべき"
    assert result["_tier"] == "1"

    # テストケース2: 第2位 - 全てライン以上 → 最小保留合計金額
    candidates_2 = [
        {"supplier_id": 1, "supplier_code": "A", "supplier_name": "EC1", "category": "ec",
         "priority": 1, "total_amount": 15000, "item_count": 5, "hold_limit": 10000, "mail_unit_price_limit": 0},
        {"supplier_id": 2, "supplier_code": "B", "supplier_name": "EC2", "category": "ec",
         "priority": 2, "total_amount": 12000, "item_count": 4, "hold_limit": 10000, "mail_unit_price_limit": 0},
    ]
    result = _apply_three_tier_priority(candidates_2)
    print(f"  Case2 (ライン以上): {result['supplier_name']} (tier={result['_tier']})")
    assert result["supplier_name"] == "EC2", "最小保留合計金額のEC2が選ばれるべき"
    assert result["_tier"] == "2"

    # テストケース3: 第3位 - 全て保留0 → 優先順位順
    candidates_3 = [
        {"supplier_id": 1, "supplier_code": "A", "supplier_name": "EC1", "category": "ec",
         "priority": 3, "total_amount": 0, "item_count": 0, "hold_limit": 10000, "mail_unit_price_limit": 0},
        {"supplier_id": 2, "supplier_code": "B", "supplier_name": "EC2", "category": "ec",
         "priority": 1, "total_amount": 0, "item_count": 0, "hold_limit": 10000, "mail_unit_price_limit": 0},
        {"supplier_id": 3, "supplier_code": "C", "supplier_name": "EC3", "category": "ec",
         "priority": 2, "total_amount": 0, "item_count": 0, "hold_limit": 10000, "mail_unit_price_limit": 0},
    ]
    result = _apply_three_tier_priority(candidates_3)
    print(f"  Case3 (保留0): {result['supplier_name']} (tier={result['_tier']})")
    assert result["supplier_name"] == "EC2", "priority=1のEC2が選ばれるべき"
    assert result["_tier"] == "3"

    # テストケース4: 混在 - ライン未満1件 + ライン以上1件 + 保留01件 → 第1位優先
    candidates_4 = [
        {"supplier_id": 1, "supplier_code": "A", "supplier_name": "EC1", "category": "ec",
         "priority": 1, "total_amount": 15000, "item_count": 5, "hold_limit": 10000, "mail_unit_price_limit": 0},
        {"supplier_id": 2, "supplier_code": "B", "supplier_name": "EC2", "category": "ec",
         "priority": 2, "total_amount": 5000, "item_count": 2, "hold_limit": 10000, "mail_unit_price_limit": 0},
        {"supplier_id": 3, "supplier_code": "C", "supplier_name": "EC3", "category": "ec",
         "priority": 3, "total_amount": 0, "item_count": 0, "hold_limit": 10000, "mail_unit_price_limit": 0},
    ]
    result = _apply_three_tier_priority(candidates_4)
    print(f"  Case4 (混在): {result['supplier_name']} (tier={result['_tier']})")
    assert result["supplier_name"] == "EC2", "ライン未満のEC2が第1位で選ばれるべき"
    assert result["_tier"] == "1"

    print("\n  3段階優先順位ロジック 全テスト合格!")


def test_code_normalizer():
    """商品コード正規化の単体テスト"""
    print("\n" + "=" * 60)
    print("商品コード正規化 テスト")
    print("=" * 60)

    from src.parser.code_normalizer import normalize_code

    tests = [
        ("9784336065704", "9784336065704", "ISBN-13そのまま"),
        ("978-4-336-06570-4", "9784336065704", "ハイフン付きISBN-13"),
        ("9784336065704SHINPIN07", "9784336065704", "管理番号から13桁切出し"),
        ("B00BC06HV8", "B00BC06HV8", "ASINそのまま"),
    ]

    for input_code, expected, desc in tests:
        result = normalize_code(input_code)
        status = "OK" if result == expected else "NG"
        print(f"  [{status}] {desc}: {input_code} → {result} (expected: {expected})")
        assert result == expected, f"Failed: {desc}"

    print("\n  商品コード正規化 全テスト合格!")


if __name__ == "__main__":
    setup_logging(log_dir="logs", level="INFO")

    test_code_normalizer()
    test_three_tier_priority()
    test_full_pipeline()

    print("\n" + "=" * 60)
    print("全テスト完了!")
    print("=" * 60)
