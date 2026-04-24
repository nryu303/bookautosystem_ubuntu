"""カスタム例外クラス定義"""


class BookAutoError(Exception):
    """システム共通の基底例外"""
    pass


class MailReadError(BookAutoError):
    """Outlookメール読取時のエラー"""
    pass


class ParseError(BookAutoError):
    """メール解析・商品コード抽出時のエラー"""
    pass


class ScrapeError(BookAutoError):
    """スクレイピング実行時のエラー"""
    pass


class StockJudgeError(BookAutoError):
    """在庫判定時のエラー"""
    pass


class HoldError(BookAutoError):
    """保留管理処理時のエラー"""
    pass


class MailSendError(BookAutoError):
    """メール送信時のエラー"""
    pass


class ConfigError(BookAutoError):
    """設定ファイル読込時のエラー"""
    pass


class DatabaseError(BookAutoError):
    """データベース操作時のエラー"""
    pass
