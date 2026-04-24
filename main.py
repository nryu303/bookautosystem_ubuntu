"""書店業務自動化システム - エントリーポイント

メインループ:
  Outlookメール監視 → 商品コード抽出 → 自家在庫照合 →
  スクレイピング → 在庫判定 → 仕入先振り分け → 保留管理 →
  再スクレイピング → 自動メール → Excel台帳出力

使い方:
  python main.py              # メインループ開始
  python main.py --web        # 管理画面のみ起動
  python main.py --once       # 1回だけ実行して終了
  python main.py --export     # Excel台帳出力のみ
"""

import argparse
import atexit
import sys
import time
import threading
import webbrowser

from src.common.logger import setup_logging, get_logger
from src.common.config import load_settings, load_suppliers, get_setting
from src.common.database import init_db, sync_suppliers_from_config
from src.scraper.scrape_lock import (
    MainInstanceBusy,
    acquire_main_instance_lock,
    release_main_instance_lock,
)


logger = get_logger(__name__)


def run_pipeline() -> dict:
    """メイン処理パイプラインを1回実行する。

    Returns:
        各ステップの処理結果サマリー
    """
    results = {
        "mails_fetched": 0,
        "items_parsed": 0,
        "self_stock_hits": 0,
        "url_csv_loaded": 0,
        "items_scraped": 0,
        "items_assigned": 0,
        "hold_processed": 0,
        "rescrape_checked": 0,
        "mails_sent": 0,
        "errors": [],
    }

    # 1. メール取得（モードに応じてCOM or Graph APIを使用）
    # Windows以外ではCOMが使えないため、デフォルトはimap
    default_mode = "com" if sys.platform == "win32" else "imap"
    try:
        mail_mode = get_setting("outlook", "mode", default=default_mode)
        if mail_mode == "imap":
            from src.mail.imap_reader import fetch_new_mails
        elif mail_mode == "graph":
            from src.mail.graph_reader import fetch_new_mails
        else:
            from src.mail.outlook_reader import fetch_new_mails
        results["mails_fetched"] = fetch_new_mails()
    except Exception as e:
        if "Invalid class string" in str(e) or "Outlook" in str(e):
            logger.warning("Outlook未検出 - メール取得をスキップします（Outlookが起動していない可能性があります）")
        else:
            logger.error("メール取得エラー: %s", e)
        results["errors"].append(f"メール取得: {e}")

    # 2. メール解析・商品コード抽出
    try:
        from src.parser.mail_parser import parse_pending_messages
        results["items_parsed"] = parse_pending_messages()
    except Exception as e:
        logger.error("メール解析エラー: %s", e)
        results["errors"].append(f"メール解析: {e}")

    # 3. 自家在庫CSV照合
    try:
        from src.stock.self_stock_checker import import_self_stock_csv, check_self_stock
        import_self_stock_csv()
        results["self_stock_hits"] = check_self_stock()
    except Exception as e:
        logger.error("自家在庫照合エラー: %s", e)
        results["errors"].append(f"自家在庫: {e}")

    # 3.5. URL指定型CSV読込（機能②）
    try:
        from src.stock.url_csv_loader import import_url_csv
        url_csv_count = import_url_csv()
        results["url_csv_loaded"] = url_csv_count
    except Exception as e:
        logger.error("URL指定型CSV読込エラー: %s", e)
        results["errors"].append(f"URL指定型CSV: {e}")

    # 4. スクレイピング
    try:
        from src.scraper.orchestrator import run_all_scraping
        scrape_stats = run_all_scraping()
        results["items_scraped"] = scrape_stats["total_scraped"]
    except Exception as e:
        logger.error("スクレイピングエラー: %s", e)
        results["errors"].append(f"スクレイピング: {e}")

    # 5. 仕入先振り分け
    try:
        from src.judge.supplier_selector import assign_pending_items
        results["items_assigned"] = assign_pending_items()
    except Exception as e:
        logger.error("振り分けエラー: %s", e)
        results["errors"].append(f"振り分け: {e}")

    # 6. 保留バケット処理
    try:
        from src.hold.hold_manager import process_hold_assignments
        results["hold_processed"] = process_hold_assignments()
    except Exception as e:
        logger.error("保留処理エラー: %s", e)
        results["errors"].append(f"保留処理: {e}")

    # 7. 再スクレイピング（閾値到達分）
    try:
        from src.hold.rescrape_trigger import process_rescrape
        rescrape_result = process_rescrape()
        results["rescrape_checked"] = rescrape_result.get("checked", 0)
    except Exception as e:
        logger.error("再スクレイピングエラー: %s", e)
        results["errors"].append(f"再スクレイピング: {e}")

    # 8. 自動メール送信
    try:
        from src.mailer.auto_mailer import process_auto_mails
        results["mails_sent"] = process_auto_mails()
    except Exception as e:
        logger.error("自動メールエラー: %s", e)
        results["errors"].append(f"自動メール: {e}")

    # 9. Excel台帳出力
    try:
        from src.export.excel_exporter import export_all
        export_all()
    except Exception as e:
        logger.error("Excel出力エラー: %s", e)
        results["errors"].append(f"Excel出力: {e}")

    return results


def main_loop():
    """メインポーリングループ。"""
    from src.sync.onedrive_sync import sync_before_pipeline, sync_after_pipeline

    interval = get_setting("outlook", "polling_interval_sec", default=300)
    logger.info("メインループ開始 (interval=%d秒)", interval)

    cycle = 0
    while True:
        cycle += 1
        try:
            # OneDrive同期: 制御フラグ・設定変更の読み取り
            sync_info = sync_before_pipeline()

            if sync_info["action"] == "stop":
                logger.info("リモート停止指示を受信しました。パイプラインを停止します。")
                break
            if sync_info["action"] == "pause":
                logger.info("リモート一時停止中。%d秒後に再チェックします。", interval)
                time.sleep(interval)
                continue

            if sync_info["suppliers_changed"] > 0:
                # 設定変更があった場合はDBに再同期
                suppliers = load_suppliers(force_reload=True)
                sync_suppliers_from_config(suppliers)

            logger.info("--- パイプライン実行開始 (cycle=%d) ---", cycle)
            results = run_pipeline()

            # リモートからの即時Excel出力指示
            if sync_info["export_now"]:
                try:
                    from src.export.excel_exporter import export_all
                    export_all()
                    logger.info("リモート指示によるExcel出力完了")
                except Exception as e:
                    logger.error("リモート指示Excel出力エラー: %s", e)

            error_count = len(results["errors"])
            if error_count == 0:
                logger.info(
                    "--- パイプライン完了 (cycle=%d) --- "
                    "メール=%d, 解析=%d, スクレイピング=%d, 振分=%d, 送信=%d [正常]",
                    cycle,
                    results["mails_fetched"],
                    results["items_parsed"],
                    results["items_scraped"],
                    results["items_assigned"],
                    results["mails_sent"],
                )
            else:
                logger.warning(
                    "--- パイプライン完了 (cycle=%d) --- "
                    "メール=%d, 解析=%d, スクレイピング=%d, 振分=%d, 送信=%d [警告: %d件のエラー]",
                    cycle,
                    results["mails_fetched"],
                    results["items_parsed"],
                    results["items_scraped"],
                    results["items_assigned"],
                    results["mails_sent"],
                    error_count,
                )

            # OneDrive同期: 稼働状況・設定の書き出し
            sync_after_pipeline(cycle, results)

        except Exception as e:
            logger.error("パイプライン全体エラー (cycle=%d): %s", cycle, e)

        logger.info("次回実行まで %d秒 待機 (Ctrl+C で停止)", interval)
        time.sleep(interval)


def main():
    """エントリーポイント。"""
    parser = argparse.ArgumentParser(description="書店業務自動化システム")
    parser.add_argument("--web", action="store_true", help="管理画面のみ起動")
    parser.add_argument("--once", action="store_true", help="1回だけ実行して終了")
    parser.add_argument("--export", action="store_true", help="Excel台帳出力のみ")
    args = parser.parse_args()

    # 初期化
    settings = load_settings()
    setup_logging(
        log_dir=get_setting("logging", "dir", default="logs"),
        level=get_setting("logging", "level", default="INFO"),
    )
    logger.info("=== 書店業務自動化システム 起動 ===")

    # 多重起動防止: --export はDB読み取りのみなのでスキップ
    if not args.export:
        try:
            acquire_main_instance_lock()
            atexit.register(release_main_instance_lock)
        except MainInstanceBusy as e:
            logger.error("多重起動検出: %s", e)
            print(f"エラー: {e}", file=sys.stderr)
            sys.exit(1)

    init_db()
    suppliers = load_suppliers()
    sync_suppliers_from_config(suppliers)

    # モード別実行
    if args.web:
        from src.web.app import run_web
        run_web()
        return

    if args.export:
        from src.export.excel_exporter import export_all
        p, s = export_all()
        print(f"Processing ledger: {p}")
        print(f"Supplier ledger: {s}")
        return

    if args.once:
        results = run_pipeline()
        logger.info("パイプライン結果: %s", results)
        return

    # デフォルト: メインループ + Web管理画面を並行起動
    host = get_setting("web", "host", default="127.0.0.1")
    port = get_setting("web", "port", default=5000)
    url = f"http://{host}:{port}"

    web_thread = threading.Thread(target=_start_web, daemon=True)
    web_thread.start()
    logger.info("管理画面をバックグラウンドで起動しました (%s)", url)

    # ブラウザを自動で開く（Webサーバー起動を少し待ってから）
    def _open_browser():
        time.sleep(1.5)
        webbrowser.open(url)
    threading.Thread(target=_open_browser, daemon=True).start()

    main_loop()


def _start_web():
    """管理画面をバックグラウンドスレッドで起動する。"""
    try:
        from src.web.app import create_app
        app = create_app()
        host = get_setting("web", "host", default="127.0.0.1")
        port = get_setting("web", "port", default=5000)
        app.run(host=host, port=port, debug=False, use_reloader=False)
    except Exception as e:
        logger.error("管理画面起動エラー: %s", e)


if __name__ == "__main__":
    main()
