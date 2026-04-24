"""Flask管理画面

処理台帳・仕入台帳の閲覧、コントロールパネル、
スクリーンショット表示、手動操作を提供する。
"""

import os
from datetime import datetime

from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify

from src.common.config import (
    get_setting, BASE_DIR, load_suppliers, save_suppliers,
    load_mail_templates, save_mail_templates,
)
from src.common.database import get_connection
from src.common.logger import get_logger

logger = get_logger(__name__)

# 在庫ステータスの日本語ラベル（テンプレート・Excel共通）
STOCK_STATUS_LABELS = {
    "AVAILABLE": "在庫あり",
    "UNAVAILABLE": "在庫なし",
    "BACKORDER": "取り寄せ可",
    "UNKNOWN": "不明",
    "ERROR": "エラー",
}

ORDER_STATUS_LABELS = {
    "PENDING": "未処理",
    "SELF_STOCK": "自家在庫",
    "HOLD": "保留中",
    "ORDERED": "発注確定",
    "CANCELLED": "キャンセル",
    "NO_STOCK": "在庫なし",
    "ERROR": "エラー",
}

BUCKET_STATUS_LABELS = {
    "ACTIVE": "積上中",
    "THRESHOLD_REACHED": "閾値到達",
    "CLEARED": "クリア",
}


def _get_template_folder() -> str:
    """PyInstaller exe / 通常実行の両方でテンプレートフォルダを解決する。"""
    import sys
    if getattr(sys, "frozen", False):
        return os.path.join(sys._MEIPASS, "src", "web", "templates")
    return os.path.join(os.path.dirname(__file__), "templates")


def create_app() -> Flask:
    app = Flask(
        __name__,
        template_folder=_get_template_folder(),
    )
    app.secret_key = "book-auto-system-secret"

    # テンプレートで使える日本語ラベル変換関数を登録
    app.jinja_env.globals["stock_label"] = lambda s: STOCK_STATUS_LABELS.get(s, s)
    app.jinja_env.globals["order_label"] = lambda s: ORDER_STATUS_LABELS.get(s, s)
    app.jinja_env.globals["bucket_label"] = lambda s: BUCKET_STATUS_LABELS.get(s, s)

    @app.route("/")
    def index():
        """ダッシュボード（処理台帳）"""
        status_filter = request.args.get("status", "")

        with get_connection() as conn:
            if status_filter:
                items = conn.execute(
                    """
                    SELECT oi.*, m.received_at, m.subject,
                           ps.supplier_name AS planned_supplier_name
                    FROM order_items oi
                    LEFT JOIN messages m ON oi.message_id = m.id
                    LEFT JOIN suppliers ps ON oi.planned_supplier_id = ps.id
                    WHERE oi.current_status = ?
                    ORDER BY oi.id DESC
                    """,
                    (status_filter,),
                ).fetchall()
            else:
                items = conn.execute(
                    """
                    SELECT oi.*, m.received_at, m.subject,
                           ps.supplier_name AS planned_supplier_name
                    FROM order_items oi
                    LEFT JOIN messages m ON oi.message_id = m.id
                    LEFT JOIN suppliers ps ON oi.planned_supplier_id = ps.id
                    ORDER BY oi.id DESC
                    """
                ).fetchall()

            # 統計
            stats = conn.execute(
                """
                SELECT
                    COUNT(*) AS total,
                    SUM(CASE WHEN current_status='PENDING' THEN 1 ELSE 0 END) AS pending,
                    SUM(CASE WHEN current_status='SELF_STOCK' THEN 1 ELSE 0 END) AS self_stock,
                    SUM(CASE WHEN current_status='HOLD' THEN 1 ELSE 0 END) AS hold,
                    SUM(CASE WHEN current_status='NO_STOCK' THEN 1 ELSE 0 END) AS no_stock,
                    SUM(CASE WHEN current_status='ORDERED' THEN 1 ELSE 0 END) AS ordered,
                    SUM(CASE WHEN current_status='ERROR' THEN 1 ELSE 0 END) AS error
                FROM order_items
                """
            ).fetchone()

            # メール取込統計
            mail_stats = conn.execute(
                """
                SELECT
                    COUNT(*) AS total,
                    SUM(CASE WHEN parse_status='DONE' THEN 1 ELSE 0 END) AS done,
                    SUM(CASE WHEN parse_status='PENDING' THEN 1 ELSE 0 END) AS pending,
                    SUM(CASE WHEN parse_status='SKIPPED' THEN 1 ELSE 0 END) AS skipped,
                    SUM(CASE WHEN parse_status='ERROR' THEN 1 ELSE 0 END) AS error
                FROM messages
                """
            ).fetchone()

        return render_template("index.html",
                               items=[dict(i) for i in items],
                               stats=dict(stats),
                               mail_stats=dict(mail_stats),
                               status_filter=status_filter)

    @app.route("/suppliers")
    def suppliers():
        """仕入台帳"""
        with get_connection() as conn:
            buckets = conn.execute(
                """
                SELECT hb.*, s.supplier_code, s.supplier_name, s.category,
                       s.hold_limit_amount, s.priority, s.scrape_enabled, s.auto_mail_enabled,
                       CASE WHEN hb.oldest_item_date IS NOT NULL
                            THEN CAST(julianday('now', 'localtime') - julianday(hb.oldest_item_date) AS INTEGER)
                            ELSE NULL END AS elapsed_days
                FROM hold_buckets hb
                JOIN suppliers s ON hb.supplier_id = s.id
                ORDER BY s.priority
                """
            ).fetchall()
        return render_template("suppliers.html", buckets=[dict(b) for b in buckets])

    @app.route("/item/<int:item_id>")
    def item_detail(item_id):
        """商品詳細（スクレイピング結果・スクリーンショット）"""
        with get_connection() as conn:
            item = conn.execute(
                """
                SELECT oi.*, m.received_at, m.subject
                FROM order_items oi
                LEFT JOIN messages m ON oi.message_id = m.id
                WHERE oi.id = ?
                """,
                (item_id,),
            ).fetchone()

            # 全仕入先の最新結果サマリー（各仕入先1行）
            # 最新結果にスクリーンショットがない場合、直近のスクリーンショットを補完する
            summary_raw = conn.execute(
                """
                SELECT sr.*, s.supplier_name, s.supplier_code, s.priority
                FROM scrape_results sr
                JOIN suppliers s ON sr.supplier_id = s.id
                WHERE sr.id IN (
                    SELECT MAX(id) FROM scrape_results
                    WHERE order_item_id = ?
                    GROUP BY supplier_id
                )
                ORDER BY s.priority
                """,
                (item_id,),
            ).fetchall()

            # 仕入先ごとの最新スクリーンショットを取得
            latest_screenshots = conn.execute(
                """
                SELECT supplier_id, screenshot_path, id
                FROM scrape_results
                WHERE id IN (
                    SELECT MAX(id) FROM scrape_results
                    WHERE order_item_id = ?
                      AND screenshot_path IS NOT NULL
                      AND screenshot_path != ''
                    GROUP BY supplier_id
                )
                """,
                (item_id,),
            ).fetchall()
            ss_map = {row["supplier_id"]: row["screenshot_path"] for row in latest_screenshots}

            summary = []
            for row in summary_raw:
                d = dict(row)
                # 最新結果にスクリーンショットがなければ、直近のものを補完
                if not d.get("screenshot_path") and d["supplier_id"] in ss_map:
                    d["screenshot_path"] = ss_map[d["supplier_id"]]
                summary.append(d)

            # 仕入先別の履歴（約1日前まで）
            results = conn.execute(
                """
                SELECT sr.*, s.supplier_name, s.supplier_code, s.priority
                FROM scrape_results sr
                JOIN suppliers s ON sr.supplier_id = s.id
                WHERE sr.order_item_id = ?
                  AND sr.scraped_at >= datetime('now', 'localtime', '-1 day')
                ORDER BY s.priority, sr.scraped_at DESC
                """,
                (item_id,),
            ).fetchall()

        if item is None:
            return "Not found", 404

        # 仕入先別にグループ化
        from collections import OrderedDict
        grouped = OrderedDict()
        for r in results:
            r = dict(r)
            name = r["supplier_name"]
            if name not in grouped:
                grouped[name] = []
            grouped[name].append(r)

        return render_template("item_detail.html",
                               item=dict(item),
                               summary=[dict(r) for r in summary],
                               grouped_results=grouped)

    @app.route("/screenshot/<path:filepath>")
    def screenshot(filepath):
        """スクリーンショット画像を返す"""
        # DB内のパスが絶対パスの場合はそのまま使う
        if os.path.isabs(filepath):
            full_path = filepath
        else:
            full_path = os.path.join(BASE_DIR, filepath)
        if os.path.exists(full_path):
            return send_file(full_path, mimetype="image/png")
        return "Not found", 404

    @app.route("/item/<int:item_id>/delete", methods=["POST"])
    def delete_item(item_id):
        """注文商品を削除する（関連データも含む）。"""
        with get_connection() as conn:
            # 関連テーブルを先に削除（外部キー制約対応）
            conn.execute("DELETE FROM hold_items WHERE order_item_id = ?", (item_id,))
            conn.execute("DELETE FROM outgoing_mails WHERE order_item_id = ?", (item_id,))
            conn.execute("DELETE FROM scrape_results WHERE order_item_id = ?", (item_id,))
            conn.execute("DELETE FROM order_items WHERE id = ?", (item_id,))
        logger.info("注文商品を削除: order_item_id=%d", item_id)
        return redirect(url_for("index"))

    @app.route("/settings", methods=["GET"])
    def settings_page():
        """コントロールパネル"""
        templates = load_mail_templates()
        return render_template("settings.html",
                               suppliers=_get_suppliers_for_settings(),
                               templates=templates,
                               active_tab=request.args.get("tab", "suppliers"))

    @app.route("/settings/update", methods=["POST"])
    def settings_update():
        """仕入先設定を更新する（DB + suppliers.yaml 同期）"""
        supplier_id = request.form.get("supplier_id")
        if not supplier_id:
            return redirect(url_for("settings_page"))

        new_values = {
            "scrape_enabled": int(request.form.get("scrape_enabled", 0)),
            "auto_mail_enabled": int(request.form.get("auto_mail_enabled", 0)),
            "mail_to_address": request.form.get("mail_to_address", "").strip(),
            "priority": int(request.form.get("priority", 99)),
            "hold_limit_amount": int(request.form.get("hold_limit_amount", 0)),
            "hold_limit_days": int(request.form.get("hold_limit_days", 14)),
            "mail_limit_amount": int(request.form.get("mail_limit_amount", 0)),
            "mail_limit_days": int(request.form.get("mail_limit_days", 0)),
            "mail_unit_price_limit": int(request.form.get("mail_unit_price_limit", 0)),
            "mail_quantity_limit": int(request.form.get("mail_quantity_limit", 0)),
        }

        # パターン（YAML専用フィールド）: テキストエリアから行分割してリスト化
        pos_text = request.form.get("positive_patterns", "")
        neg_text = request.form.get("negative_patterns", "")
        new_values["positive_patterns"] = [
            line.strip() for line in pos_text.splitlines() if line.strip()
        ]
        new_values["negative_patterns"] = [
            line.strip() for line in neg_text.splitlines() if line.strip()
        ]

        # DB更新
        with get_connection() as conn:
            conn.execute(
                """
                UPDATE suppliers SET
                    scrape_enabled = ?,
                    auto_mail_enabled = ?,
                    mail_to_address = ?,
                    priority = ?,
                    hold_limit_amount = ?,
                    hold_limit_days = ?,
                    mail_limit_amount = ?,
                    mail_limit_days = ?,
                    mail_unit_price_limit = ?,
                    mail_quantity_limit = ?
                WHERE id = ?
                """,
                (
                    new_values["scrape_enabled"],
                    new_values["auto_mail_enabled"],
                    new_values["mail_to_address"],
                    new_values["priority"],
                    new_values["hold_limit_amount"],
                    new_values["hold_limit_days"],
                    new_values["mail_limit_amount"],
                    new_values["mail_limit_days"],
                    new_values["mail_unit_price_limit"],
                    new_values["mail_quantity_limit"],
                    supplier_id,
                ),
            )

            # supplier_code を取得
            row = conn.execute(
                "SELECT supplier_code FROM suppliers WHERE id = ?",
                (supplier_id,),
            ).fetchone()

        # suppliers.yaml に同期
        if row:
            _sync_supplier_to_yaml(row["supplier_code"], new_values)

        return redirect(url_for("settings_page"))

    @app.route("/logs")
    def logs_page():
        """システムログ"""
        with get_connection() as conn:
            logs = conn.execute(
                "SELECT * FROM logs ORDER BY id DESC LIMIT 200"
            ).fetchall()
        return render_template("logs.html", logs=[dict(l) for l in logs])

    @app.route("/api/stats")
    def api_stats():
        """ダッシュボード用の統計API"""
        with get_connection() as conn:
            stats = conn.execute(
                """
                SELECT
                    COUNT(*) AS total,
                    SUM(CASE WHEN current_status='PENDING' THEN 1 ELSE 0 END) AS pending,
                    SUM(CASE WHEN current_status='HOLD' THEN 1 ELSE 0 END) AS hold,
                    SUM(CASE WHEN current_status='ORDERED' THEN 1 ELSE 0 END) AS ordered
                FROM order_items
                """
            ).fetchone()
        return jsonify(dict(stats))

    @app.route("/supplier/<int:supplier_id>/confirm", methods=["POST"])
    def confirm_order(supplier_id):
        """閾値到達バケットを手動で発注確定する。

        電話発注完了後にこのボタンを押すことで:
        1. バケット内の全商品を HOLD → ORDERED に変更
        2. バケットをクリア
        3. auto_mail_enabled の場合は自動メール送信対象になる
        """
        from src.hold.hold_manager import clear_bucket

        with get_connection() as conn:
            bucket = conn.execute(
                "SELECT * FROM hold_buckets WHERE supplier_id = ?",
                (supplier_id,),
            ).fetchone()

        if bucket is None:
            return redirect(url_for("suppliers"))

        if bucket["item_count"] == 0:
            return redirect(url_for("supplier_hold_items", supplier_id=supplier_id))

        clear_bucket(bucket["id"])
        logger.info("手動発注確定: supplier_id=%d, bucket_id=%d", supplier_id, bucket["id"])

        return redirect(url_for("supplier_hold_items", supplier_id=supplier_id))

    @app.route("/supplier/<int:supplier_id>/items")
    def supplier_hold_items(supplier_id):
        """仕入先の保留商品一覧をJSON/HTMLで返す（モーダル表示用）"""
        with get_connection() as conn:
            supplier = conn.execute(
                "SELECT * FROM suppliers WHERE id = ?", (supplier_id,),
            ).fetchone()

            bucket = conn.execute(
                "SELECT * FROM hold_buckets WHERE supplier_id = ?", (supplier_id,),
            ).fetchone()

            items = []
            if bucket:
                items = conn.execute(
                    """
                    SELECT hi.assigned_at, hi.amount,
                           oi.id AS order_item_id, oi.order_number,
                           oi.product_code_normalized, oi.product_name,
                           oi.quantity, oi.current_status,
                           m.received_at
                    FROM hold_items hi
                    JOIN order_items oi ON hi.order_item_id = oi.id
                    LEFT JOIN messages m ON oi.message_id = m.id
                    WHERE hi.hold_bucket_id = ?
                    ORDER BY hi.assigned_at ASC
                    """,
                    (bucket["id"],),
                ).fetchall()

        if request.args.get("format") == "json":
            return jsonify({
                "supplier": dict(supplier) if supplier else {},
                "items": [dict(i) for i in items],
            })

        return render_template("supplier_items.html",
                               supplier=dict(supplier) if supplier else {},
                               bucket=dict(bucket) if bucket else {},
                               items=[dict(i) for i in items])

    @app.route("/settings/templates", methods=["GET"])
    def templates_page():
        """メールテンプレート編集ページ"""
        templates = load_mail_templates()
        return render_template("settings.html",
                               suppliers=_get_suppliers_for_settings(),
                               templates=templates,
                               active_tab="templates")

    @app.route("/settings/templates/update", methods=["POST"])
    def templates_update():
        """メールテンプレートを更新する（mail_templates.yaml に保存）"""
        templates = load_mail_templates()

        # store_order テンプレート
        store_subject = request.form.get("store_order_subject", "").strip()
        store_body = request.form.get("store_order_body", "").strip()
        if store_subject or store_body:
            if "store_order" not in templates:
                templates["store_order"] = {}
            templates["store_order"]["subject"] = store_subject
            templates["store_order"]["body"] = store_body + "\n"

        # store_bulk_order テンプレート
        bulk_subject = request.form.get("store_bulk_order_subject", "").strip()
        bulk_body = request.form.get("store_bulk_order_body", "").strip()
        if bulk_subject or bulk_body:
            if "store_bulk_order" not in templates:
                templates["store_bulk_order"] = {}
            templates["store_bulk_order"]["subject"] = bulk_subject
            templates["store_bulk_order"]["body"] = bulk_body + "\n"

        save_mail_templates(templates)
        return redirect(url_for("settings_page"))

    # ====== パターンエディタ API ======

    @app.route("/api/patterns/<int:supplier_id>")
    def api_get_patterns(supplier_id):
        """仕入先の抽出パターン一覧を返す。"""
        with get_connection() as conn:
            rows = conn.execute(
                """
                SELECT id, field_name, selector, xpath, class_hint, text_hint,
                       pattern_name, active_flag
                FROM extraction_patterns
                WHERE supplier_id = ?
                ORDER BY id ASC
                """,
                (supplier_id,),
            ).fetchall()
        return jsonify([dict(r) for r in rows])

    @app.route("/api/patterns/<int:supplier_id>", methods=["POST"])
    def api_save_pattern(supplier_id):
        """抽出パターンを1件追加または更新する。"""
        data = request.get_json()
        if not data or not data.get("field_name"):
            return jsonify({"error": "field_name is required"}), 400

        pattern_id = data.get("id")
        with get_connection() as conn:
            if pattern_id:
                conn.execute(
                    """
                    UPDATE extraction_patterns
                    SET field_name = ?, selector = ?, xpath = ?,
                        class_hint = ?, text_hint = ?, pattern_name = ?, active_flag = ?
                    WHERE id = ? AND supplier_id = ?
                    """,
                    (
                        data["field_name"],
                        data.get("selector", ""),
                        data.get("xpath", ""),
                        data.get("class_hint", ""),
                        data.get("text_hint", ""),
                        data.get("pattern_name", ""),
                        int(data.get("active_flag", 1)),
                        pattern_id,
                        supplier_id,
                    ),
                )
            else:
                conn.execute(
                    """
                    INSERT INTO extraction_patterns
                        (supplier_id, field_name, selector, xpath, class_hint, text_hint, pattern_name, active_flag)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        supplier_id,
                        data["field_name"],
                        data.get("selector", ""),
                        data.get("xpath", ""),
                        data.get("class_hint", ""),
                        data.get("text_hint", ""),
                        data.get("pattern_name", ""),
                        int(data.get("active_flag", 1)),
                    ),
                )
        return jsonify({"ok": True})

    @app.route("/api/patterns/<int:supplier_id>/<int:pattern_id>", methods=["DELETE"])
    def api_delete_pattern(supplier_id, pattern_id):
        """抽出パターンを1件削除する。"""
        with get_connection() as conn:
            conn.execute(
                "DELETE FROM extraction_patterns WHERE id = ? AND supplier_id = ?",
                (pattern_id, supplier_id),
            )
        return jsonify({"ok": True})

    @app.route("/api/url_template/<int:supplier_id>")
    def api_get_url_template(supplier_id):
        """仕入先のURLテンプレートを取得する。"""
        with get_connection() as conn:
            row = conn.execute(
                "SELECT url_template FROM url_rules WHERE supplier_id = ?",
                (supplier_id,),
            ).fetchone()
        return jsonify({"url_template": row["url_template"] if row else ""})

    @app.route("/api/url_template/<int:supplier_id>", methods=["POST"])
    def api_save_url_template(supplier_id):
        """仕入先のURLテンプレートを更新する（DB + suppliers.yaml 同期）。"""
        data = request.get_json()
        new_template = data.get("url_template", "").strip()

        with get_connection() as conn:
            # url_rules に既にある場合は更新、なければ挿入
            existing = conn.execute(
                "SELECT id FROM url_rules WHERE supplier_id = ?",
                (supplier_id,),
            ).fetchone()
            if existing:
                conn.execute(
                    "UPDATE url_rules SET url_template = ? WHERE supplier_id = ?",
                    (new_template, supplier_id),
                )
            else:
                conn.execute(
                    "INSERT INTO url_rules (supplier_id, mode, url_template) VALUES (?, 'template', ?)",
                    (supplier_id, new_template),
                )
            # supplier_code を取得して YAML に同期
            row = conn.execute(
                "SELECT supplier_code FROM suppliers WHERE id = ?",
                (supplier_id,),
            ).fetchone()

        if row:
            try:
                suppliers_list = load_suppliers(force_reload=True)
                for sup in suppliers_list:
                    if sup.get("code") == row["supplier_code"]:
                        sup["url_template"] = new_template
                        break
                save_suppliers(suppliers_list)
            except Exception as e:
                logger.error("suppliers.yaml URL同期エラー: %s", e)

        return jsonify({"ok": True})

    @app.route("/api/html_preview/<int:supplier_id>")
    def api_html_preview(supplier_id):
        """仕入先の最新HTMLスナップショットからテキストノード一覧を返す。"""
        with get_connection() as conn:
            row = conn.execute(
                """
                SELECT html_snapshot_path FROM scrape_results
                WHERE supplier_id = ?
                  AND html_snapshot_path IS NOT NULL AND html_snapshot_path != ''
                ORDER BY id DESC LIMIT 1
                """,
                (supplier_id,),
            ).fetchone()

        if not row or not row["html_snapshot_path"]:
            return jsonify({"error": "HTMLスナップショットがありません", "nodes": []}), 404

        html_path = row["html_snapshot_path"]
        if not os.path.isabs(html_path):
            html_path = os.path.join(BASE_DIR, html_path)
        if not os.path.exists(html_path):
            return jsonify({"error": "HTMLファイルが見つかりません", "nodes": []}), 404

        from bs4 import BeautifulSoup
        with open(html_path, "r", encoding="utf-8") as f:
            soup = BeautifulSoup(f.read(), "lxml")

        # script/style 除去
        for tag in soup(["script", "style", "noscript"]):
            tag.decompose()

        nodes = []
        seen_texts = set()
        for el in soup.body.find_all(True) if soup.body else []:
            text = el.get_text(strip=True)
            if not text or len(text) > 500 or text in seen_texts:
                continue
            seen_texts.add(text)

            # CSSセレクタを構築
            tag_name = el.name
            el_id = el.get("id", "")
            el_classes = " ".join(el.get("class", []))
            selector = tag_name
            if el_id:
                selector += f"#{el_id}"
            elif el_classes:
                selector += "." + ".".join(el.get("class", []))

            nodes.append({
                "text": text[:200],
                "tag": tag_name,
                "selector": selector,
                "className": el_classes,
            })

        return jsonify({"nodes": nodes, "path": row["html_snapshot_path"]})

    @app.route("/api/html_snapshot/<int:supplier_id>")
    def api_html_snapshot(supplier_id):
        """仕入先の最新HTMLスナップショットをiframe用に返す。

        元HTMLに操作用JSを注入し、クリックでセレクタを親ウィンドウに送信する。
        CSSや画像の相対パスが正しく解決されるよう <base> タグも注入する。
        """
        with get_connection() as conn:
            row = conn.execute(
                """
                SELECT html_snapshot_path, page_url FROM scrape_results
                WHERE supplier_id = ?
                  AND html_snapshot_path IS NOT NULL AND html_snapshot_path != ''
                ORDER BY id DESC LIMIT 1
                """,
                (supplier_id,),
            ).fetchone()

            # フォールバック用: URL テンプレートを取得
            url_rule = conn.execute(
                "SELECT url_template FROM url_rules WHERE supplier_id = ?",
                (supplier_id,),
            ).fetchone()

        if not row or not row["html_snapshot_path"]:
            return "<html><body><p style='padding:2rem;color:#999;'>HTMLスナップショットがありません</p></body></html>"

        html_path = row["html_snapshot_path"]
        if not os.path.isabs(html_path):
            html_path = os.path.join(BASE_DIR, html_path)
        if not os.path.exists(html_path):
            return "<html><body><p style='padding:2rem;color:#999;'>HTMLファイルが見つかりません</p></body></html>"

        with open(html_path, "r", encoding="utf-8") as f:
            html_content = f.read()

        # 相対パス（CSS/JS/画像）を絶対URLに書き換える
        try:
            page_url = (row["page_url"] or "") if row else ""
        except (IndexError, KeyError):
            page_url = ""
        base_url = _detect_base_url(page_url, url_rule, html_content)
        if base_url:
            html_content = _rewrite_relative_urls(html_content, base_url)

        # 操作用JSを注入（</body>の直前に挿入）
        interaction_js = """
<style>
  ._pe-hover { outline: 3px solid #f59e0b !important; outline-offset: 2px; cursor: crosshair !important; }
  ._pe-selected { outline: 3px solid #2563eb !important; outline-offset: 2px; background-color: rgba(37,99,235,.08) !important; }
  ._pe-assigned-stock_text { outline: 3px solid #16a34a !important; outline-offset: 2px; background-color: rgba(22,163,74,.08) !important; }
  ._pe-assigned-price { outline: 3px solid #9333ea !important; outline-offset: 2px; background-color: rgba(147,51,234,.08) !important; }
  ._pe-assigned-product_title { outline: 3px solid #0891b2 !important; outline-offset: 2px; background-color: rgba(8,145,178,.08) !important; }
  ._pe-assigned-positive_signal { outline: 3px solid #16a34a !important; outline-offset: 2px; background-color: rgba(22,163,74,.15) !important; }
  ._pe-assigned-negative_signal { outline: 3px solid #dc2626 !important; outline-offset: 2px; background-color: rgba(220,38,38,.08) !important; }
  ._pe-popup {
    position: fixed; z-index: 999999; background: #fff; border: 2px solid #2563eb;
    border-radius: 8px; box-shadow: 0 8px 32px rgba(0,0,0,.25); padding: 12px 16px;
    font-family: 'Yu Gothic','Meiryo',sans-serif; font-size: 13px; min-width: 260px;
  }
  ._pe-popup select { width: 100%; padding: 6px 8px; margin: 8px 0; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 13px; }
  ._pe-popup .pe-btn-row { display: flex; gap: 6px; justify-content: flex-end; }
  ._pe-popup button { padding: 5px 14px; border: none; border-radius: 4px; cursor: pointer; font-size: 12px; font-weight: 600; }
  ._pe-popup .pe-ok { background: #2563eb; color: #fff; }
  ._pe-popup .pe-cancel { background: #e2e8f0; color: #334155; }
  * { pointer-events: auto !important; }
</style>
<script>
(function(){
  // Disable all links and forms
  document.querySelectorAll('a').forEach(a => { a.removeAttribute('href'); a.onclick = e => e.preventDefault(); });
  document.querySelectorAll('form').forEach(f => { f.onsubmit = e => e.preventDefault(); });

  let hovered = null;
  let popup = null;
  let selectedEl = null;

  function buildSelector(el) {
    if (el.id) return el.tagName.toLowerCase() + '#' + el.id;
    let sel = el.tagName.toLowerCase();
    if (el.className && typeof el.className === 'string') {
      const classes = el.className.split(/\\s+/).filter(c => c && !c.startsWith('_pe-'));
      if (classes.length) sel += '.' + classes.join('.');
    }
    // Add nth-child for uniqueness if no id/class
    if (!el.id && (!el.className || typeof el.className !== 'string' || !el.className.trim())) {
      const parent = el.parentElement;
      if (parent) {
        const siblings = Array.from(parent.children).filter(c => c.tagName === el.tagName);
        if (siblings.length > 1) {
          const idx = siblings.indexOf(el) + 1;
          sel += ':nth-child(' + idx + ')';
        }
      }
    }
    // Walk up to build a more specific path (max 3 levels)
    let parts = [sel];
    let cur = el.parentElement;
    let depth = 0;
    while (cur && cur !== document.body && depth < 3) {
      let pSel = cur.tagName.toLowerCase();
      if (cur.id) { pSel += '#' + cur.id; parts.unshift(pSel); break; }
      if (cur.className && typeof cur.className === 'string') {
        const pc = cur.className.split(/\\s+/).filter(c => c && !c.startsWith('_pe-'));
        if (pc.length) pSel += '.' + pc.join('.');
      }
      parts.unshift(pSel);
      cur = cur.parentElement;
      depth++;
    }
    return parts.join(' > ');
  }

  document.addEventListener('mouseover', function(e) {
    if (popup) return;
    const el = e.target;
    if (el === document.body || el === document.documentElement) return;
    if (el.closest('._pe-popup')) return;
    if (hovered && hovered !== el) hovered.classList.remove('_pe-hover');
    el.classList.add('_pe-hover');
    hovered = el;
  }, true);

  document.addEventListener('mouseout', function(e) {
    if (hovered) hovered.classList.remove('_pe-hover');
  }, true);

  document.addEventListener('click', function(e) {
    const el = e.target;
    if (el.closest('._pe-popup')) return;
    e.preventDefault();
    e.stopPropagation();

    // Close existing popup
    if (popup) { popup.remove(); popup = null; }
    if (selectedEl) selectedEl.classList.remove('_pe-selected');

    el.classList.remove('_pe-hover');
    el.classList.add('_pe-selected');
    selectedEl = el;

    const text = (el.textContent || '').trim().substring(0, 120);
    const selector = buildSelector(el);

    // Create popup near click
    popup = document.createElement('div');
    popup.className = '_pe-popup';
    const rx = Math.min(e.clientX + 10, window.innerWidth - 300);
    const ry = Math.min(e.clientY + 10, window.innerHeight - 200);
    popup.style.left = rx + 'px';
    popup.style.top = ry + 'px';
    popup.innerHTML = `
      <div style="margin-bottom:6px;"><strong>選択したテキスト:</strong></div>
      <div style="background:#f8fafc;padding:6px 8px;border-radius:4px;margin-bottom:8px;font-size:12px;max-height:60px;overflow:auto;word-break:break-all;">${text || '(空)'}</div>
      <div style="font-size:11px;color:#64748b;margin-bottom:4px;">セレクタ: <code>${selector}</code></div>
      <label style="font-weight:600;">データ種別を選択:</label>
      <select id="_pe-field-sel">
        <option value="">-- 選んでください --</option>
        <option value="stock_text">在庫テキスト</option>
        <option value="price">価格</option>
        <option value="product_title">商品名</option>
        <option value="positive_signal">在庫あり判定</option>
        <option value="negative_signal">在庫なし判定</option>
      </select>
      <div class="pe-btn-row">
        <button class="pe-cancel" id="_pe-btn-cancel">キャンセル</button>
        <button class="pe-ok" id="_pe-btn-ok">登録</button>
      </div>`;
    document.body.appendChild(popup);

    document.getElementById('_pe-btn-cancel').onclick = function() {
      popup.remove(); popup = null;
      if (selectedEl) { selectedEl.classList.remove('_pe-selected'); selectedEl = null; }
    };
    document.getElementById('_pe-btn-ok').onclick = function() {
      const field = document.getElementById('_pe-field-sel').value;
      if (!field) { alert('データ種別を選んでください'); return; }
      // Send to parent window
      window.parent.postMessage({
        type: 'pe-assign',
        field_name: field,
        selector: selector,
        text: text,
      }, '*');
      // Mark as assigned
      if (selectedEl) {
        selectedEl.classList.remove('_pe-selected');
        selectedEl.classList.add('_pe-assigned-' + field);
      }
      popup.remove(); popup = null;
      selectedEl = null;
    };
  }, true);

  // Listen for highlight commands from parent
  window.addEventListener('message', function(e) {
    if (e.data && e.data.type === 'pe-highlight') {
      // Highlight existing patterns
      (e.data.patterns || []).forEach(function(p) {
        try {
          const el = document.querySelector(p.selector);
          if (el) el.classList.add('_pe-assigned-' + p.field_name);
        } catch(ex) {}
      });
    }
  });
})();
</script>
"""
        # Inject before </body> or at end
        body_lower = html_content.lower()
        idx = body_lower.rfind("</body>")
        if idx != -1:
            html_content = html_content[:idx] + interaction_js + html_content[idx:]
        else:
            html_content += interaction_js

        return html_content

    @app.route("/pattern-editor")
    def pattern_editor():
        """ビジュアルパターンエディタ（専用ページ）"""
        return render_template("pattern_editor.html",
                               suppliers=_get_suppliers_for_settings())

    return app


def _rewrite_relative_urls(html_content: str, base_url: str) -> str:
    """HTML内の相対パス（href/src）を絶対URLに書き換える。

    <base> タグ方式より確実。iframe内でも外部CSSが正しく読み込まれる。
    """
    import re
    from urllib.parse import urljoin

    def _replace(match):
        attr = match.group(1)     # href or src
        quote = match.group(2)    # " or '
        url = match.group(3)      # the URL value
        # 既に絶対URL、データURI、フラグメント、JS の場合はそのまま
        if url.startswith(("http://", "https://", "data:", "javascript:", "#", "//")):
            return match.group(0)
        # 空の場合もそのまま
        if not url.strip():
            return match.group(0)
        # 相対パスを絶対URLに変換
        absolute = urljoin(base_url, url)
        return f'{attr}={quote}{absolute}{quote}'

    # href="..." と src="..." を書き換え
    html_content = re.sub(
        r'(href|src)=(["\'])(.*?)\2',
        _replace,
        html_content,
        flags=re.IGNORECASE,
    )
    return html_content


def _detect_base_url(page_url: str, url_rule, html_content: str = "") -> str:
    """スナップショットの <base> タグ用ベースURLを決定する。

    優先順位:
      1. スクレイピング時に保存された実際のページURL
      2. HTML内の絶対URLと相対パスから推定
      3. url_rules の url_template からの推定
    """
    import re
    from urllib.parse import urlparse

    # 方法1: 実際のページURL（最も正確）
    if page_url and page_url.startswith("http"):
        parsed = urlparse(page_url)
        path_dir = parsed.path.rsplit("/", 1)[0] + "/"
        return f"{parsed.scheme}://{parsed.netloc}{path_dir}"

    # 方法2: HTML内の絶対URL + 相対パスから推定
    if html_content:
        abs_urls = re.findall(
            r'(?:href|src)=["\']?(https?://[^"\'>\s]+)',
            html_content[:5000],
            re.IGNORECASE,
        )
        rel_refs = re.findall(
            r'(?:href|src)=["\'](\.\./[^"\'>\s]+)',
            html_content[:5000],
            re.IGNORECASE,
        )
        if abs_urls and rel_refs:
            # 相対パスの ../X/ 部分と絶対URLの /A/X/ 部分をマッチさせて
            # ページのディレクトリ深度を推定する
            # 例: rel=../assets/css/x.css, abs=https://ex.com/kino/assets/img/y.png
            #   → assets/ が共通 → ../ は /kino/ に解決 → ページは /kino/???/
            rel_first_dir = rel_refs[0].lstrip("./").split("/")[0]  # "assets"
            for abs_url in abs_urls:
                p = urlparse(abs_url)
                parts = p.path.strip("/").split("/")
                if rel_first_dir in parts:
                    idx = parts.index(rel_first_dir)
                    # ../ が解決する先 = 絶対URLのパスで rel_first_dir の前まで
                    parent_path = "/" + "/".join(parts[:idx]) + "/"
                    # ページはその1階層下にある
                    base = f"{p.scheme}://{p.netloc}{parent_path}_page/"
                    return base

        # 絶対URLだけでもドメインを取得
        if abs_urls:
            # CDN系（googletagmanager等）を除外
            for abs_url in abs_urls:
                p = urlparse(abs_url)
                if "google" not in p.netloc and "cdn" not in p.netloc:
                    return f"{p.scheme}://{p.netloc}/"

    # 方法3: url_template からフォールバック
    if url_rule and url_rule["url_template"]:
        tmpl = url_rule["url_template"]
        if tmpl.startswith("http"):
            parsed = urlparse(tmpl)
            path_dir = parsed.path.rsplit("/", 1)[0] + "/"
            return f"{parsed.scheme}://{parsed.netloc}{path_dir}"

    return ""


def _get_suppliers_for_settings() -> list[dict]:
    """設定ページ用の仕入先一覧を取得する。

    DB の列に加えて、suppliers.yaml からパターン情報を読み込み
    テキストエリア表示用の文字列に変換して付与する。
    """
    with get_connection() as conn:
        sups = conn.execute(
            "SELECT * FROM suppliers ORDER BY priority"
        ).fetchall()

    result = [dict(s) for s in sups]

    # suppliers.yaml からパターンを取得してマージ
    yaml_suppliers = load_suppliers(force_reload=True)
    yaml_map = {s.get("code"): s for s in yaml_suppliers}

    for s in result:
        yaml_entry = yaml_map.get(s["supplier_code"], {})
        pos = yaml_entry.get("positive_patterns", [])
        neg = yaml_entry.get("negative_patterns", [])
        s["positive_patterns_text"] = "\n".join(pos) if pos else ""
        s["negative_patterns_text"] = "\n".join(neg) if neg else ""

    return result


def _sync_supplier_to_yaml(supplier_code: str, new_values: dict) -> None:
    """Web UIでの変更を suppliers.yaml に書き戻す。"""
    try:
        suppliers_list = load_suppliers(force_reload=True)

        # YAML内のフィールド名はDB列名と一部異なる
        yaml_field_map = {
            "scrape_enabled": "scrape_enabled",
            "auto_mail_enabled": "auto_mail_enabled",
            "mail_to_address": "mail_to_address",
            "priority": "priority",
            "hold_limit_amount": "hold_limit_amount",
            "hold_limit_days": "hold_limit_days",
            "mail_limit_amount": "mail_limit_amount",
            "mail_limit_days": "mail_limit_days",
            "mail_unit_price_limit": "mail_unit_price_limit",
            "mail_quantity_limit": "mail_quantity_limit",
        }

        for sup in suppliers_list:
            if sup.get("code") == supplier_code:
                for db_field, yaml_field in yaml_field_map.items():
                    value = new_values.get(db_field)
                    if value is not None:
                        # YAML側は bool で保存（scrape_enabled, auto_mail_enabled）
                        if yaml_field in ("scrape_enabled", "auto_mail_enabled"):
                            sup[yaml_field] = bool(value)
                        else:
                            sup[yaml_field] = value
                # パターン（YAML専用フィールド）
                if "positive_patterns" in new_values:
                    sup["positive_patterns"] = new_values["positive_patterns"]
                if "negative_patterns" in new_values:
                    sup["negative_patterns"] = new_values["negative_patterns"]
                break

        save_suppliers(suppliers_list)
    except Exception as e:
        logger.error("suppliers.yaml 同期エラー: %s", e)


def run_web():
    """管理画面Webサーバーを起動する（ブラウザ自動起動付き）。"""
    import threading
    import time
    import webbrowser

    host = get_setting("web", "host", default="127.0.0.1")
    port = get_setting("web", "port", default=5000)
    url = f"http://{host}:{port}"

    def _open_browser():
        time.sleep(1.5)
        webbrowser.open(url)
    threading.Thread(target=_open_browser, daemon=True).start()

    app = create_app()
    app.run(host=host, port=port, debug=False)
