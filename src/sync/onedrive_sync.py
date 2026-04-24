"""OneDrive同期モジュール

OneDrive同期フォルダを介したリモート閲覧・制御機能を提供する。

同期対象:
  view/   → 本体からOneDriveへ書き出し（閲覧用）
  control/ → OneDriveから本体が読み取り（操作フラグ）
  edit/   → OneDriveから本体へ設定反映（編集用）
"""

import json
import os
import shutil
from datetime import datetime

import yaml

from src.common.config import (
    BASE_DIR, CONFIG_DIR, get_setting, load_settings, load_suppliers,
    save_suppliers,
)
from src.common.database import get_connection
from src.common.logger import get_logger

logger = get_logger(__name__)

ONEDRIVE_BASE = os.path.join(
    os.path.expanduser("~"), "OneDrive", "BookAutoSystem"
)

CONTROL_FILE = os.path.join(ONEDRIVE_BASE, "control", "control.json")
STATUS_FILE = os.path.join(ONEDRIVE_BASE, "view", "status.json")
SETTINGS_VIEW = os.path.join(ONEDRIVE_BASE, "view", "settings_readonly.yaml")
SUPPLIERS_VIEW = os.path.join(ONEDRIVE_BASE, "view", "suppliers_readonly.yaml")
SUPPLIERS_EDIT = os.path.join(ONEDRIVE_BASE, "edit", "suppliers_edit.yaml")

# 設定ファイル内のマスク対象キー
_MASK_KEYS = {"imap_password", "graph_client_id", "graph_tenant_id"}


def _ensure_dirs() -> bool:
    """OneDriveフォルダが存在するか確認し、なければスキップ。"""
    if not os.path.isdir(ONEDRIVE_BASE):
        logger.debug("OneDrive同期フォルダが見つかりません: %s", ONEDRIVE_BASE)
        return False
    for sub in ("control", "view", "edit"):
        os.makedirs(os.path.join(ONEDRIVE_BASE, sub), exist_ok=True)
    return True


# ─── Step 2: リモート操作フラグの読み取り ───

def read_control() -> dict:
    """control.json を読み取り、操作指示を返す。

    Returns:
        {"action": "run"|"pause"|"stop",
         "export_now": bool,
         "rescrape_item_ids": list[int]}
    """
    default = {"action": "run", "export_now": False, "rescrape_item_ids": []}
    if not os.path.isfile(CONTROL_FILE):
        return default
    try:
        with open(CONTROL_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return {
            "action": data.get("action", "run"),
            "export_now": bool(data.get("export_now", False)),
            "rescrape_item_ids": list(data.get("rescrape_item_ids", [])),
        }
    except Exception as e:
        logger.warning("control.json 読み取りエラー: %s", e)
        return default


def reset_control_flags() -> None:
    """一度きりのフラグ（export_now, rescrape_item_ids）をリセットする。"""
    if not os.path.isfile(CONTROL_FILE):
        return
    try:
        with open(CONTROL_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        data["export_now"] = False
        data["rescrape_item_ids"] = []
        with open(CONTROL_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception as e:
        logger.warning("control.json リセットエラー: %s", e)


# ─── Step 3a: 閲覧用ファイルの書き出し ───

def export_status(cycle: int, results: dict) -> None:
    """パイプラインの稼働状況を status.json に書き出す。"""
    if not _ensure_dirs():
        return
    try:
        interval = get_setting("outlook", "polling_interval_sec", default=300)
        now = datetime.now()
        next_run = datetime.fromtimestamp(now.timestamp() + interval)

        # DBから商品ステータス集計
        stats = {}
        try:
            with get_connection() as conn:
                row = conn.execute("""
                    SELECT
                        COUNT(*) AS total_items,
                        SUM(CASE WHEN current_status='PENDING' THEN 1 ELSE 0 END) AS pending,
                        SUM(CASE WHEN current_status='SELF_STOCK' THEN 1 ELSE 0 END) AS self_stock,
                        SUM(CASE WHEN current_status='HOLD' THEN 1 ELSE 0 END) AS hold,
                        SUM(CASE WHEN current_status='ORDERED' THEN 1 ELSE 0 END) AS ordered,
                        SUM(CASE WHEN current_status='NO_STOCK' THEN 1 ELSE 0 END) AS no_stock,
                        SUM(CASE WHEN current_status='ERROR' THEN 1 ELSE 0 END) AS error
                    FROM order_items
                """).fetchone()
                stats = dict(row) if row else {}
        except Exception:
            pass

        status_data = {
            "system": "BookAutoSystem",
            "status": "running",
            "last_cycle": cycle,
            "last_run_at": now.strftime("%Y-%m-%d %H:%M:%S"),
            "next_run_at": next_run.strftime("%Y-%m-%d %H:%M:%S"),
            "polling_interval_sec": interval,
            "last_results": {
                "mails_fetched": results.get("mails_fetched", 0),
                "items_parsed": results.get("items_parsed", 0),
                "items_scraped": results.get("items_scraped", 0),
                "items_assigned": results.get("items_assigned", 0),
                "hold_processed": results.get("hold_processed", 0),
                "mails_sent": results.get("mails_sent", 0),
                "error_count": len(results.get("errors", [])),
            },
            "stats": stats,
        }

        with open(STATUS_FILE, "w", encoding="utf-8") as f:
            json.dump(status_data, f, ensure_ascii=False, indent=4)

    except Exception as e:
        logger.warning("status.json 書き出しエラー: %s", e)


def export_settings_view() -> None:
    """settings.yaml のマスク済みコピーを閲覧用に書き出す。"""
    if not _ensure_dirs():
        return
    try:
        settings = load_settings()
        masked = _mask_sensitive(settings)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with open(SETTINGS_VIEW, "w", encoding="utf-8") as f:
            f.write(f"# === 閲覧専用 (このファイルを編集しても本体には反映されません) ===\n")
            f.write(f"# 最終更新: {now}\n\n")
            yaml.dump(masked, f, default_flow_style=False,
                      allow_unicode=True, sort_keys=False)

    except Exception as e:
        logger.warning("settings_readonly.yaml 書き出しエラー: %s", e)


def export_suppliers_view() -> None:
    """suppliers.yaml のコピーを閲覧用に書き出す。"""
    if not _ensure_dirs():
        return
    try:
        src = os.path.join(CONFIG_DIR, "suppliers.yaml")
        shutil.copy2(src, SUPPLIERS_VIEW)
    except Exception as e:
        logger.warning("suppliers_readonly.yaml 書き出しエラー: %s", e)


def _mask_sensitive(data: dict, depth: int = 0) -> dict:
    """辞書内のパスワード等の値をマスクする。"""
    if depth > 5:
        return data
    result = {}
    for k, v in data.items():
        if k in _MASK_KEYS and isinstance(v, str) and v:
            result[k] = "********"
        elif isinstance(v, dict):
            result[k] = _mask_sensitive(v, depth + 1)
        else:
            result[k] = v
    return result


# ─── Step 3b: 編集ファイルの取り込み ───

# 編集可能フィールド
_EDITABLE_FIELDS = {
    "priority", "hold_limit_amount", "hold_limit_days",
    "scrape_enabled", "auto_mail_enabled",
    "mail_limit_amount", "mail_limit_days",
    "mail_unit_price_limit", "mail_quantity_limit",
}


def import_suppliers_edit() -> int:
    """edit/suppliers_edit.yaml の変更をローカル suppliers.yaml に反映する。

    Returns:
        変更された仕入先の数
    """
    if not os.path.isfile(SUPPLIERS_EDIT):
        return 0

    try:
        with open(SUPPLIERS_EDIT, "r", encoding="utf-8") as f:
            edit_data = yaml.safe_load(f)
        edit_list = edit_data.get("suppliers", [])
        if not edit_list:
            return 0

        # code をキーにした辞書
        edit_map = {s["code"]: s for s in edit_list if "code" in s}

        # 現在の suppliers.yaml を読み込み
        current_list = load_suppliers(force_reload=True)
        changed_count = 0

        for supplier in current_list:
            code = supplier.get("code")
            if code not in edit_map:
                continue

            edit_sup = edit_map[code]
            for field in _EDITABLE_FIELDS:
                if field in edit_sup and edit_sup[field] != supplier.get(field):
                    old_val = supplier.get(field)
                    supplier[field] = edit_sup[field]
                    logger.info(
                        "仕入先設定変更 (%s): %s: %s → %s",
                        code, field, old_val, edit_sup[field],
                    )
                    changed_count += 1

        if changed_count > 0:
            save_suppliers(current_list)
            # 閲覧用コピーも更新
            export_suppliers_view()
            logger.info("OneDrive編集ファイルから %d件 の設定変更を反映しました", changed_count)

        return changed_count

    except Exception as e:
        logger.warning("suppliers_edit.yaml 取り込みエラー: %s", e)
        return 0


# ─── 統合: パイプラインから呼ぶ関数 ───

def sync_before_pipeline() -> dict:
    """パイプライン実行前の同期処理。

    Returns:
        {"action": str, "export_now": bool, "rescrape_item_ids": list,
         "suppliers_changed": int}
    """
    if not _ensure_dirs():
        return {"action": "run", "export_now": False,
                "rescrape_item_ids": [], "suppliers_changed": 0}

    control = read_control()
    suppliers_changed = import_suppliers_edit()

    return {
        "action": control["action"],
        "export_now": control["export_now"],
        "rescrape_item_ids": control["rescrape_item_ids"],
        "suppliers_changed": suppliers_changed,
    }


def sync_after_pipeline(cycle: int, results: dict) -> None:
    """パイプライン実行後の同期処理。"""
    if not _ensure_dirs():
        return

    export_status(cycle, results)
    export_settings_view()
    export_suppliers_view()
    reset_control_flags()
