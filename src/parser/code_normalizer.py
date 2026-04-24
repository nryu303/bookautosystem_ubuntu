"""商品コード正規化モジュール

ISBN-13, ISBN-10, ASIN, その他管理番号を正規化する。
"""

import re
from src.common.logger import get_logger

logger = get_logger(__name__)


def _remove_hyphens(code: str) -> str:
    """ハイフン・スペースを除去する。"""
    return re.sub(r"[\s\-]", "", code)


def _isbn10_check_digit(digits: str) -> str:
    """ISBN-10のチェックディジットを計算する。"""
    total = sum((10 - i) * int(d) for i, d in enumerate(digits[:9]))
    remainder = (11 - (total % 11)) % 11
    return "X" if remainder == 10 else str(remainder)


def _isbn13_check_digit(digits: str) -> str:
    """ISBN-13のチェックディジットを計算する。"""
    total = sum(
        int(d) * (1 if i % 2 == 0 else 3)
        for i, d in enumerate(digits[:12])
    )
    return str((10 - (total % 10)) % 10)


def _isbn10_to_isbn13(isbn10: str) -> str:
    """ISBN-10をISBN-13に変換する。"""
    base = "978" + isbn10[:9]
    check = _isbn13_check_digit(base)
    return base + check


def validate_isbn13(code: str) -> bool:
    """ISBN-13のチェックディジットを検証する。"""
    if len(code) != 13 or not code.isdigit():
        return False
    return _isbn13_check_digit(code) == code[-1]


def validate_isbn10(code: str) -> bool:
    """ISBN-10のチェックディジットを検証する。"""
    if len(code) != 10:
        return False
    if not code[:9].isdigit():
        return False
    return _isbn10_check_digit(code) == code[-1].upper()


def normalize_code(raw_code: str) -> str | None:
    """商品コードを正規化する。

    Args:
        raw_code: 元の商品コード文字列

    Returns:
        正規化されたコード。正規化できない場合はNone。
    """
    if not raw_code or not raw_code.strip():
        return None

    code = _remove_hyphens(raw_code.strip())

    # ISBN-13 (978 or 979 で始まる13桁)
    if len(code) == 13 and code.isdigit() and code[:3] in ("978", "979"):
        if validate_isbn13(code):
            return code
        logger.warning("ISBN-13チェックディジット不一致: %s", raw_code)
        return code  # チェック不一致でもそのまま使う

    # ISBN-10 (10桁、末尾がXの場合あり)
    if len(code) == 10 and code[:9].isdigit() and (code[-1].isdigit() or code[-1].upper() == "X"):
        if validate_isbn10(code):
            return _isbn10_to_isbn13(code)
        logger.warning("ISBN-10チェックディジット不一致: %s → ISBN-13変換を試行", raw_code)
        return _isbn10_to_isbn13(code)

    # ASIN (Bで始まる10桁英数字)
    if len(code) == 10 and code[0].upper() == "B" and code.isalnum():
        return code.upper()

    # 13桁以上の管理番号 → 先頭13桁を切り出してISBN-13として試す
    if len(code) >= 13 and code[:13].isdigit() and code[:3] in ("978", "979"):
        isbn13 = code[:13]
        if validate_isbn13(isbn13):
            return isbn13
        return isbn13

    # その他の数字列（JANコードなど）
    if code.isdigit() and len(code) >= 8:
        return code

    # 正規化不能
    logger.warning("正規化できない商品コード: %s", raw_code)
    return raw_code.strip()


def extract_codes_from_text(text: str) -> list[str]:
    """テキストから商品コードらしき文字列を全て抽出する。

    ISBN-13, ISBN-10, ASIN パターンを検出する。
    """
    codes = []

    # ISBN-13 パターン (978/979 + 10桁、ハイフンあり/なし)
    for m in re.finditer(r"(?:978|979)[\d\-]{10,}", text):
        cleaned = _remove_hyphens(m.group())
        if len(cleaned) == 13:
            codes.append(cleaned)

    # ISBN-10 パターン (数字9桁 + 数字orX、ISBNラベル付き)
    for m in re.finditer(r"ISBN[：:\s]*(\d[\d\-]{8,}[\dXx])", text):
        cleaned = _remove_hyphens(m.group(1))
        if len(cleaned) == 10:
            normalized = normalize_code(cleaned)
            if normalized and normalized not in codes:
                codes.append(normalized)

    # ASIN パターン
    for m in re.finditer(r"\b(B[A-Z0-9]{9})\b", text):
        code = m.group(1)
        if code not in codes:
            codes.append(code)

    # 重複除去（順序保持）
    seen = set()
    unique = []
    for c in codes:
        if c not in seen:
            seen.add(c)
            unique.append(c)

    return unique
