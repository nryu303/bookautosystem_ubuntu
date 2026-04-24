"""在庫状態・注文状態などの列挙型定義"""

from enum import Enum


class AvailabilityStatus(str, Enum):
    """スクレイピング結果の在庫判定ステータス"""
    AVAILABLE = "AVAILABLE"         # 在庫あり
    UNAVAILABLE = "UNAVAILABLE"     # 在庫なし
    BACKORDER = "BACKORDER"         # 取り寄せ可
    UNKNOWN = "UNKNOWN"             # 判定不能
    ERROR = "ERROR"                 # スクレイピングエラー


class OrderStatus(str, Enum):
    """商品(order_item)の処理ステータス"""
    PENDING = "PENDING"             # 未処理（スクレイピング待ち）
    SELF_STOCK = "SELF_STOCK"       # 自家在庫あり
    HOLD = "HOLD"                   # 保留中
    ORDERED = "ORDERED"             # 発注確定
    CANCELLED = "CANCELLED"         # キャンセル
    ERROR = "ERROR"                 # エラー


class ParseStatus(str, Enum):
    """メール解析ステータス"""
    PENDING = "PENDING"             # 未解析
    DONE = "DONE"                   # 解析完了
    SKIPPED = "SKIPPED"             # 注文メール以外（スキップ）
    ERROR = "ERROR"                 # 解析エラー


class HoldBucketStatus(str, Enum):
    """保留バケットのステータス"""
    ACTIVE = "ACTIVE"               # 積み上げ中
    THRESHOLD_REACHED = "THRESHOLD_REACHED"  # 閾値到達
    CLEARED = "CLEARED"             # 発注済みクリア


class SupplierCategory(str, Enum):
    """仕入先カテゴリ"""
    EC = "ec"                       # ECサイト（通販）
    STORE = "store"                 # 実店舗


class UrlMode(str, Enum):
    """URL生成モード"""
    TEMPLATE = "template"           # テンプレート方式（機能①）
    CSV_LOOKUP = "csv_lookup"       # CSV参照方式（従来の簡易方式）
    URL_CSV_DB = "url_csv_db"       # URL指定型CSV方式（機能②、DB経由・最安値選択）


class MailSendStatus(str, Enum):
    """メール送信ステータス"""
    SUCCESS = "SUCCESS"
    FAILED = "FAILED"
