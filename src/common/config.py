"""設定ファイル(YAML)読み込みモジュール"""

import os
import yaml
from src.common.exceptions import ConfigError
from src.common.logger import get_logger

logger = get_logger(__name__)

_config_cache: dict | None = None
_suppliers_cache: dict | None = None
_mail_patterns_cache: dict | None = None

def _resolve_base_dir() -> str:
    """PyInstaller exe / 通常実行の両方で正しいベースディレクトリを返す。"""
    import sys
    if getattr(sys, "frozen", False):
        # PyInstaller exe: exe の隣に config/ data/ がある想定
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


BASE_DIR = _resolve_base_dir()
CONFIG_DIR = os.path.join(BASE_DIR, "config")


def _load_yaml(file_path: str) -> dict:
    """YAMLファイルを読み込んで辞書で返す。"""
    if not os.path.exists(file_path):
        raise ConfigError(f"設定ファイルが見つかりません: {file_path}")
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
        if data is None:
            raise ConfigError(f"設定ファイルが空です: {file_path}")
        return data
    except yaml.YAMLError as e:
        raise ConfigError(f"YAML解析エラー ({file_path}): {e}")


def load_settings(force_reload: bool = False) -> dict:
    """settings.yaml を読み込む。キャッシュあり。"""
    global _config_cache
    if _config_cache is not None and not force_reload:
        return _config_cache
    path = os.path.join(CONFIG_DIR, "settings.yaml")
    _config_cache = _load_yaml(path)
    logger.info("設定ファイル読込完了: %s", path)
    return _config_cache


def load_suppliers(force_reload: bool = False) -> list[dict]:
    """suppliers.yaml を読み込み、仕入先リストを返す。"""
    global _suppliers_cache
    if _suppliers_cache is not None and not force_reload:
        return _suppliers_cache
    path = os.path.join(CONFIG_DIR, "suppliers.yaml")
    data = _load_yaml(path)
    _suppliers_cache = data.get("suppliers", [])
    logger.info("仕入先設定読込完了: %d件", len(_suppliers_cache))
    return _suppliers_cache


def load_mail_patterns(force_reload: bool = False) -> dict:
    """mail_patterns.yaml を読み込む。"""
    global _mail_patterns_cache
    if _mail_patterns_cache is not None and not force_reload:
        return _mail_patterns_cache
    path = os.path.join(CONFIG_DIR, "mail_patterns.yaml")
    _mail_patterns_cache = _load_yaml(path)
    logger.info("メールパターン設定読込完了: %s", path)
    return _mail_patterns_cache


def save_suppliers(suppliers_list: list[dict]) -> None:
    """仕入先リストを suppliers.yaml に書き戻す。キャッシュも更新する。"""
    global _suppliers_cache
    path = os.path.join(CONFIG_DIR, "suppliers.yaml")
    data = {"suppliers": suppliers_list}
    try:
        with open(path, "w", encoding="utf-8") as f:
            yaml.dump(
                data, f,
                default_flow_style=False,
                allow_unicode=True,
                sort_keys=False,
            )
        _suppliers_cache = suppliers_list
        logger.info("仕入先設定を保存しました: %s (%d件)", path, len(suppliers_list))
    except Exception as e:
        raise ConfigError(f"仕入先設定の保存に失敗しました: {e}")


def load_mail_templates(force_reload: bool = False) -> dict:
    """mail_templates.yaml を読み込む。"""
    path = os.path.join(CONFIG_DIR, "mail_templates.yaml")
    if not os.path.exists(path):
        return {}
    return _load_yaml(path)


def save_mail_templates(templates: dict) -> None:
    """メールテンプレートを mail_templates.yaml に書き戻す。"""
    path = os.path.join(CONFIG_DIR, "mail_templates.yaml")
    try:
        with open(path, "w", encoding="utf-8") as f:
            yaml.dump(
                templates, f,
                default_flow_style=False,
                allow_unicode=True,
                sort_keys=False,
            )
        logger.info("メールテンプレートを保存しました: %s", path)
    except Exception as e:
        raise ConfigError(f"メールテンプレートの保存に失敗しました: {e}")


def get_setting(*keys: str, default=None):
    """ネストされた設定値をドット的にたどって取得する。

    例: get_setting("outlook", "account_name")
    """
    settings = load_settings()
    value = settings
    for key in keys:
        if isinstance(value, dict):
            value = value.get(key)
        else:
            return default
        if value is None:
            return default
    return value
